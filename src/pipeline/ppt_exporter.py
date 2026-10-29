"""Stitch a lecture's deduplicated iCourse PPT screenshots into a .pptx deck.

While a lecture plays, the iCourse recorder screenshots the desktop every
20-30 seconds (``/pptnote/v1/schedule/search-ppt``).  ``PPTPipeline`` then
keeps only the screenshots that survive the dHash dedup and the garbage
filter — i.e. exactly the slides the lecture actually used — and stores
them in ``ppt_pages`` with ``ocr_status='done'``.

This module turns those surviving pages back into a real presentation:
one slide per kept page, plus a title slide, with the page's OCR text
attached as speaker notes so the deck stays searchable.  The deck is
written under ``config.PPT_EXPORT_DIR`` and attached to the daily email
by ``Emailer``; files are removed once the email has been sent.

Kept deliberately dependency-light: only ``python-pptx`` + ``Pillow``
(``Pillow`` is already required by the OCR pipeline).
"""

from __future__ import annotations

import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from typing import TYPE_CHECKING, Optional

from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt

from src.api import icourse
from src.pipeline import ppt_slide_cache
from src.runtime import config

if TYPE_CHECKING:
    from src.api.icourse import ICourseClient
    from src.data.database import Database
    from src.runtime.reporter import Reporter

# Presentation geometry: 16:9 with a small margin around the picture so
# screenshots whose aspect ratio is not exactly 16:9 don't touch the edges.
SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5
PICTURE_MARGIN_IN = 0.18

# (max_width_px, jpeg_quality) ladder used only when the first pass exceeds
# PPT_EXPORT_MAX_MB.  Keeps typical decks at full 1600 px / q82 while
# guaranteeing a size bound for slide-heavy lectures.
_SIZE_LADDER = ((1.0, 1.0), (0.85, 0.9), (0.70, 0.8), (0.55, 0.7),
                (0.45, 0.6))


def _safe_filename(title: str) -> str:
    """Sanitise a title for use inside an attachment filename."""
    return "".join(
        c if c.isalnum() or c in " _-" else "_" for c in (title or "")
    )[:80]


def _fmt_offset(seconds: int) -> str:
    m, s = divmod(max(int(seconds or 0), 0), 60)
    return f"{m:02d}:{s:02d}"


def _to_jpeg(data: bytes, width: int, quality: int) -> Optional[bytes]:
    """Decode one screenshot and re-encode it as a size-bounded JPEG."""
    try:
        with Image.open(BytesIO(data)) as im:
            im = im.convert("RGB")
            if im.width > width:
                new_h = max(round(im.height * width / im.width), 1)
                im = im.resize((width, new_h), Image.LANCZOS)
            buf = BytesIO()
            im.save(buf, "JPEG", quality=quality, optimize=True,
                    progressive=True)
            return buf.getvalue()
    except Exception as e:                              # noqa: BLE001
        print(f"    [PPT] could not decode a slide screenshot: "
              f"{type(e).__name__}: {e}", flush=True)
        return None


def _fit_budget(raw_images: list[bytes]) -> list[bytes]:
    """Re-encode screenshots until the deck fits ``PPT_EXPORT_MAX_MB``.

    Returns the encoded JPEGs; the last ladder step is returned even when it
    still exceeds the budget — ``Emailer`` then decides whether to drop the
    deck, so we never lose slides by silently truncating the lecture.
    """
    budget = int(config.PPT_EXPORT_MAX_MB * 1024 * 1024)
    base_w = config.PPT_EXPORT_IMAGE_WIDTH
    base_q = config.PPT_EXPORT_JPEG_QUALITY
    encoded: list[bytes] = []
    for w_scale, q_scale in _SIZE_LADDER:
        width = max(int(base_w * w_scale), 640)
        quality = max(min(int(base_q * q_scale), 95), 40)
        encoded = []
        for raw in raw_images:
            jpeg = _to_jpeg(raw, width, quality)
            if jpeg is not None:
                encoded.append(jpeg)
        total = sum(len(b) for b in encoded)
        if total <= budget or (w_scale, q_scale) == _SIZE_LADDER[-1]:
            print(f"    [PPT] encoded {len(encoded)} slides at "
                  f"{width}px/q{quality}: {total / (1024 * 1024):.1f} MB "
                  f"(budget {config.PPT_EXPORT_MAX_MB:.0f} MB)", flush=True)
            return encoded
    return encoded


def _add_title_slide(prs: Presentation, course_title: str, sub_title: str,
                     date: str, slide_count: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])   # blank
    box = slide.shapes.add_textbox(Inches(0.8), Inches(2.0),
                                   Inches(SLIDE_WIDTH_IN - 1.6), Inches(1.6))
    frame = box.text_frame
    frame.word_wrap = True
    title_run = frame.paragraphs[0].add_run()
    title_run.text = course_title
    title_run.font.size = Pt(34)
    title_run.font.bold = True
    p = frame.add_paragraph()
    p.text = sub_title
    p.font.size = Pt(22)

    box2 = slide.shapes.add_textbox(Inches(0.8), Inches(4.1),
                                    Inches(SLIDE_WIDTH_IN - 1.6), Inches(1.2))
    frame2 = box2.text_frame
    frame2.word_wrap = True
    info = frame2.paragraphs[0]
    info.text = f"{date} · 共 {slide_count} 页幻灯片（iCourse 自动截图）"
    info.font.size = Pt(14)


def _add_picture_slide(prs: Presentation, jpeg: bytes, index: int,
                       total: int, note: str, offset_sec: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])   # blank
    pic = slide.shapes.add_picture(BytesIO(jpeg), 0, 0)
    # Fit inside the slide box, preserving aspect ratio, then centre it.
    max_w = Inches(SLIDE_WIDTH_IN - 2 * PICTURE_MARGIN_IN)
    max_h = Inches(SLIDE_HEIGHT_IN - 2 * PICTURE_MARGIN_IN)
    ratio = min(max_w / pic.width, max_h / pic.height)
    pic.width = int(pic.width * ratio)
    pic.height = int(pic.height * ratio)
    pic.left = int((Inches(SLIDE_WIDTH_IN) - pic.width) / 2)
    pic.top = int((Inches(SLIDE_HEIGHT_IN) - pic.height) / 2)

    # Small timestamp/counter caption in the bottom-right corner.
    cap = slide.shapes.add_textbox(
        Inches(SLIDE_WIDTH_IN - 2.2), Inches(SLIDE_HEIGHT_IN - 0.42),
        Inches(2.0), Inches(0.32),
    )
    cap.text_frame.word_wrap = False
    cap.text_frame.paragraphs[0].text = (
        f"{index}/{total} · {_fmt_offset(offset_sec)}"
    )

    if note and note.strip():
        slide.notes_slide.notes_text_frame.text = note.strip()[:4000]


def build_lecture_pptx(client: "ICourseClient", db: "Database",
                       sub_id: str, course_title: str, sub_title: str,
                       date: str,
                       reporter: "Reporter | None" = None) -> Optional[str]:
    """Build ``PPT_EXPORT_DIR/<sub_id>.pptx`` for one lecture.

    Returns the file path, or ``None`` when the lecture has no surviving
    slides / no screenshot could be downloaded.  Never raises for a single
    broken image; the caller treats export as best-effort.
    """
    sub_id = str(sub_id)
    try:
        rows = db.get_ppt_export_pages(sub_id)
    except Exception as e:                              # noqa: BLE001
        print(f"    [PPT] cannot read slide rows: "
              f"{type(e).__name__}: {e}", flush=True)
        return None
    if not rows:
        print("    [PPT] No deduplicated slides for this lecture — "
              "skipping deck.", flush=True)
        return None

    if len(rows) > config.PPT_EXPORT_MAX_SLIDES:
        print(f"    [PPT] {len(rows)} slides, keeping the first "
              f"{config.PPT_EXPORT_MAX_SLIDES}", flush=True)
        rows = rows[:config.PPT_EXPORT_MAX_SLIDES]

    slide_bytes = _resolve_slide_bytes(client, sub_id, rows)

    raw_images: list[bytes] = []
    notes: list[str] = []
    offsets: list[int] = []
    for row in rows:
        data = slide_bytes.get(int(row.get("page_num") or 0))
        if data is None:
            continue
        raw_images.append(data)
        notes.append(row.get("text") or "")
        offsets.append(int(row.get("created_sec") or 0))

    failed = len(rows) - len(raw_images)
    if not raw_images:
        print(f"    [PPT] all {len(rows)} slide screenshots failed to "
              f"download — skipping deck.", flush=True)
        return None
    if failed:
        print(f"    [PPT] {failed}/{len(rows)} slide screenshots could not "
              f"be downloaded", flush=True)

    encoded = _fit_budget(raw_images)
    if not encoded:
        return None

    os.makedirs(config.PPT_EXPORT_DIR, exist_ok=True)
    out_path = os.path.join(
        config.PPT_EXPORT_DIR, f"{_safe_filename(sub_id)}.pptx"
    )

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH_IN)
    prs.slide_height = Inches(SLIDE_HEIGHT_IN)
    _add_title_slide(prs, course_title, sub_title, date, len(encoded))
    for i, (jpeg, note, offset) in enumerate(
            zip(encoded, notes, offsets), start=1):
        _add_picture_slide(prs, jpeg, i, len(encoded), note, offset)
    prs.save(out_path)

    size_mb = os.path.getsize(out_path) / (1024 * 1024)
    print(f"    [PPT] Deck saved: {len(encoded)} slides, {size_mb:.1f} MB "
          f"→ {out_path}", flush=True)
    if reporter is not None:
        reporter.info(f"    [PPT] {sub_title}: {len(encoded)} slides, "
                      f"{size_mb:.1f} MB")
    # The deck has the pixels now; drop the per-lecture screenshot cache.
    ppt_slide_cache.clear(sub_id)
    return out_path


def _resolve_slide_bytes(client: "ICourseClient", sub_id: str,
                         rows: list[dict]) -> dict[int, bytes]:
    """Return {page_num: image bytes} for the deck's rows.

    Fresh lectures already had their deduplicated screenshots cached to disk
    by ``PPTPipeline`` (see ``ppt_slide_cache``); only pages that were
    deduplicated in an earlier run need an HTTP round-trip.  Those are
    fetched concurrently — the same pattern as the OCR image pool — with the
    retrying ``download_ppt_image`` helper.
    """
    found: dict[int, bytes] = {}
    missing: list[dict] = []
    for row in rows:
        page_num = int(row.get("page_num") or 0)
        data = ppt_slide_cache.load_page(sub_id, page_num)
        if data:
            found[page_num] = data
        else:
            missing.append(row)

    cached_n = len(found)
    downloaded_n = 0
    if missing:
        workers = max(1, min(config.PPT_EXPORT_DOWNLOAD_WORKERS,
                             len(missing)))
        with ThreadPoolExecutor(max_workers=workers) as pool:
            futures = {
                pool.submit(
                    icourse.download_ppt_image, client,
                    {"pptimgurl": row.get("pptimgurl"),
                     "page_num": row.get("page_num")},
                ): row
                for row in missing
            }
            for fut in as_completed(futures):
                row = futures[fut]
                try:
                    data = fut.result()
                except Exception:                       # noqa: BLE001
                    data = None
                if data:
                    found[int(row.get("page_num") or 0)] = data
                    downloaded_n += 1

    print(f"    [PPT] slides ready: {len(found)}/{len(rows)} "
          f"({cached_n} cached from this run, {downloaded_n} downloaded, "
          f"{len(missing) - downloaded_n} failed)", flush=True)
    return found


def cleanup_exports() -> None:
    """Remove every generated deck (called at run start / after the email)."""
    import shutil
    shutil.rmtree(config.PPT_EXPORT_DIR, ignore_errors=True)
