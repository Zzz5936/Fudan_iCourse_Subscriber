"""On-disk cache for the deduplicated PPT screenshots of the current run.

The PPT pipeline already downloads every page screenshot once (image pool →
dHash dedup → OCR).  When deck export is enabled, the survivors are written
to ``config.PPT_EXPORT_DIR/cache/<sub_id>/`` right after dedup — the same
pattern the audio path uses (ffmpeg writes the decoded lecture to disk once
and everything downstream reads that file instead of re-downloading).

The exporter then builds the deck from those cached bytes, and only falls
back to a (concurrent, retrying) HTTP download for lectures whose pages were
deduplicated in an earlier run, i.e. before this cache existed.

This module deliberately has no heavy imports: ``ppt_pipeline`` imports it
on the hot path and must stay importable even when python-pptx is absent.
"""

from __future__ import annotations

import os
import shutil

from src.runtime import config

CACHE_DIRNAME = "cache"


def _cache_root() -> str:
    return os.path.join(config.PPT_EXPORT_DIR, CACHE_DIRNAME)


def lecture_dir(sub_id: str) -> str:
    return os.path.join(_cache_root(), str(sub_id))


def _page_path(sub_id: str, page_num: int) -> str:
    return os.path.join(lecture_dir(sub_id), f"{int(page_num):05d}.bin")


def save_pages(sub_id: str, images: dict[int, bytes]) -> int:
    """Persist raw screenshot bytes for the pages that survived dedup.

    Best-effort: a write failure only means the exporter re-downloads that
    page later.  Returns the number of files written.
    """
    if not images or not config.PPT_EXPORT_ENABLED:
        return 0
    directory = lecture_dir(sub_id)
    written = 0
    try:
        os.makedirs(directory, exist_ok=True)
    except OSError:
        return 0
    for page_num, data in images.items():
        if not data:
            continue
        try:
            with open(_page_path(sub_id, page_num), "wb") as f:
                f.write(data)
            written += 1
        except OSError:
            continue
    return written


def load_page(sub_id: str, page_num: int) -> bytes | None:
    """Return the cached screenshot for one page, or ``None``."""
    path = _page_path(sub_id, page_num)
    try:
        with open(path, "rb") as f:
            return f.read()
    except OSError:
        return None


def clear(sub_id: str) -> None:
    """Drop one lecture's cached screenshots (after the deck was built)."""
    shutil.rmtree(lecture_dir(sub_id), ignore_errors=True)
