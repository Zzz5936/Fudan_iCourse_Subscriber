"""Build one course's archive: a single merged Markdown + all slide decks.

Two entry points share the same data:

* ``--print-meta`` resolves the course and prints what is already in the
  database, so a workflow can name the branch and fetch its current content
  before doing any expensive work.
* the default mode walks the course's lectures and fills in only what is
  missing (see "Reuse" below), writing the Markdown and the new decks to
  ``--out-dir``.

Reuse — the point of this script is to avoid redoing work:

* Markdown is rebuilt from the stored summaries; nothing is re-summarised.
* A deck that already exists on the course branch (``--existing-dir``) is
  skipped entirely — no PPT fetch, no OCR, no download.
* Lectures whose ``ppt_pages`` are already OCR'd in the database skip the
  PPT pipeline and go straight to deck assembly (only the slide images are
  re-fetched, which is unavoidable for building a .pptx).

    python scripts/backfill_pptx.py --course 数字逻辑 --print-meta
    python scripts/backfill_pptx.py --course 数字逻辑 \
        --existing-dir existing --out-dir course_out
"""

import argparse
import os
import shutil
import sys
import time
import traceback

# Allow importing from the project root when run as
# `python scripts/backfill_pptx.py`
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.data.database import Database  # noqa: E402
from src.pipeline.archive import (  # noqa: E402
    deck_filename, write_course_markdown,
)
from src.runtime.reporter import Reporter  # noqa: E402


def _login(max_attempts: int = 5):
    """Log in to WebVPN + iCourse CAS, the same way ``main.login_with_retry``
    does (kept local so the script does not import the ASR stack)."""
    from src.api.webvpn import WebVPNSession

    for attempt in range(max_attempts):
        try:
            vpn = WebVPNSession()
            print(f"[Login] WebVPN (attempt {attempt + 1}/{max_attempts})...",
                  flush=True)
            vpn.login()
            print("[Login] iCourse CAS...", flush=True)
            vpn.authenticate_icourse()
            return vpn
        except Exception as e:                          # noqa: BLE001
            print(f"  Failed: {type(e).__name__}: {e}", flush=True)
            if attempt == max_attempts - 1:
                raise
            time.sleep(5)
    raise RuntimeError("login failed")


def _resolve_course(db: Database, needle: str) -> dict:
    """Accept either a ``course_id`` or a (unique) fragment of the title."""
    row = db.conn.execute(
        "SELECT course_id, title FROM courses WHERE course_id = ?", (needle,)
    ).fetchone()
    if row:
        return dict(row)

    # An exact title ("数字逻辑") beats a fragment that also matches a sibling
    # course ("数字逻辑实验").
    row = db.conn.execute(
        "SELECT course_id, title FROM courses WHERE title = ?", (needle,)
    ).fetchone()
    if row:
        return dict(row)

    rows = db.conn.execute(
        "SELECT course_id, title FROM courses WHERE title LIKE ?",
        (f"%{needle}%",),
    ).fetchall()
    if len(rows) == 1:
        return dict(rows[0])
    if not rows:
        raise SystemExit(f"找不到课程：{needle}（用 --list 查看库里的课程）")
    listing = "\n".join(f"  {r['course_id']}  {r['title']}" for r in rows)
    raise SystemExit(f"“{needle}”匹配到多门课，请改用 course_id：\n{listing}")


def _coverage(db: Database, course_id: str) -> dict:
    def one(sql: str) -> int:
        return db.conn.execute(sql, (course_id,)).fetchone()[0]

    return {
        "lectures": one("SELECT COUNT(*) FROM lectures WHERE course_id = ?"),
        "summarized": one(
            "SELECT COUNT(*) FROM lectures"
            " WHERE course_id = ? AND summary IS NOT NULL"),
        "with_pages": one(
            "SELECT COUNT(*) FROM lectures l WHERE l.course_id = ?"
            " AND EXISTS (SELECT 1 FROM ppt_pages p WHERE p.sub_id = l.sub_id"
            "             AND p.ocr_status = 'done')"),
    }


def _list_courses(db: Database) -> None:
    """Print every stored course with its coverage, so the caller can pick."""
    rows = db.conn.execute(
        """SELECT c.course_id, c.title,
                  (SELECT COUNT(*) FROM lectures l
                    WHERE l.course_id = c.course_id) AS total,
                  (SELECT COUNT(*) FROM lectures l
                    WHERE l.course_id = c.course_id
                      AND l.summary IS NOT NULL) AS summarized,
                  (SELECT COUNT(*) FROM lectures l
                    WHERE l.course_id = c.course_id
                      AND EXISTS (SELECT 1 FROM ppt_pages p
                                   WHERE p.sub_id = l.sub_id
                                     AND p.ocr_status = 'done')) AS with_slides
             FROM courses c
            ORDER BY c.title"""
    ).fetchall()
    print(f"{'course_id':<12}{'讲次':>5}{'已总结':>7}{'有课件':>7}  课程", flush=True)
    for r in rows:
        print(f"{r['course_id']:<12}{r['total']:>5}{r['summarized']:>7}"
              f"{r['with_slides']:>7}  {r['title']}", flush=True)


def _lectures_of(db: Database, course_id: str,
                 sub_ids: set[str] | None) -> list[dict]:
    rows = db.conn.execute(
        """SELECT sub_id, sub_title, date, summary FROM lectures
            WHERE course_id = ?
            ORDER BY CAST(sub_id AS INTEGER) ASC""",
        (course_id,),
    ).fetchall()
    lectures = [dict(r) for r in rows]
    if sub_ids:
        lectures = [l for l in lectures if str(l["sub_id"]) in sub_ids]
    return lectures


def _existing_decks(directory: str) -> set[str]:
    """Deck file names already present on the course branch."""
    if not directory:
        return set()
    ppt_dir = os.path.join(directory, "ppt")
    if not os.path.isdir(ppt_dir):
        return set()
    return {n for n in os.listdir(ppt_dir) if n.lower().endswith(".pptx")}


def main() -> None:
    parser = argparse.ArgumentParser(description="Build a one-course archive.")
    parser.add_argument("--course", default="",
                        help="course_id or a unique fragment of the title")
    parser.add_argument("--list", action="store_true",
                        help="list stored courses and exit")
    parser.add_argument("--print-meta", action="store_true",
                        help="print course id/title/branch/coverage and exit")
    parser.add_argument("--sub-ids", default="",
                        help="comma-separated sub_ids to restrict the run")
    parser.add_argument("--limit", type=int, default=0,
                        help="stop after N lectures (0 = all)")
    parser.add_argument("--existing-dir", default="",
                        help="tree of the course branch (decks already built)")
    parser.add_argument("--out-dir", default="",
                        help="where to write the Markdown + new decks")
    args = parser.parse_args()

    db = Database()

    if args.list or not args.course:
        _list_courses(db)
        return

    course = _resolve_course(db, args.course)
    course_id, course_title = course["course_id"], course["title"]
    coverage = _coverage(db, course_id)

    if args.print_meta:
        print(f"COURSE_ID={course_id}")
        print(f"COURSE_TITLE={course_title}")
        print(f"BRANCH=course-{course_id}")
        print(f"LECTURES={coverage['lectures']}")
        print(f"SUMMARIZED={coverage['summarized']}")
        print(f"WITH_PAGES={coverage['with_pages']}")
        return

    sub_ids = {s.strip() for s in args.sub_ids.split(",") if s.strip()}
    lectures = _lectures_of(db, course_id, sub_ids or None)
    if not lectures:
        raise SystemExit(f"{course_title} 没有匹配的讲次")
    if args.limit > 0:
        lectures = lectures[:args.limit]

    if coverage["summarized"] < coverage["lectures"]:
        print(f"[Warn] {course_title} 有 "
              f"{coverage['lectures'] - coverage['summarized']} 讲还没总结；"
              f"先跑 Single Run 补齐，否则 md 会缺这些讲次。", flush=True)

    existing = _existing_decks(args.existing_dir)
    if existing:
        print(f"[Reuse] 分支上已有 {len(existing)} 份课件，将跳过重建",
              flush=True)

    reporter = Reporter()
    reporter.run_header()

    # Imported here so `--list` / `--print-meta` work without the network and
    # OCR stack installed.
    from src.api.icourse import ICourseClient
    from src.pipeline.ppt_exporter import build_lecture_pptx
    from src.pipeline.ppt_pipeline import PPTPipeline
    from src.runtime.scheduler import Scheduler

    vpn = _login()
    client = ICourseClient(vpn)
    scheduler = Scheduler(reporter=reporter)
    pipeline = PPTPipeline(db, scheduler, reporter)

    out_ppt_dir = os.path.join(args.out_dir, "ppt") if args.out_dir else ""
    if out_ppt_dir:
        os.makedirs(out_ppt_dir, exist_ok=True)

    built = empty = failed = skipped = reused = 0
    t0 = time.time()
    try:
        for i, lec in enumerate(lectures, 1):
            sub_id = str(lec["sub_id"])
            dest_name = deck_filename(course_title, lec.get("date") or "",
                                      sub_id)

            if dest_name in existing:
                skipped += 1
                print(f"\n[{i}/{len(lectures)}] {sub_id} 已有课件，跳过",
                      flush=True)
                continue

            print(f"\n[{i}/{len(lectures)}] {course_title} · "
                  f"{lec['sub_title']} ({lec['date']})", flush=True)
            try:
                if not client.check_alive():
                    print("  Session expired, re-logging in...", flush=True)
                    client.vpn = _login()
                    client._userinfo = None

                # Pages already OCR'd in the database: skip fetch/dedup/OCR.
                if db.get_ppt_export_pages(sub_id):
                    reused += 1
                    print("  复用数据库里已处理的幻灯片页面", flush=True)
                else:
                    handle = pipeline.submit(client, course_id, sub_id)
                    handle.drain()

                deck = build_lecture_pptx(
                    client, db, sub_id, course_title,
                    lec.get("sub_title") or sub_id, lec.get("date") or "",
                    reporter,
                )
                if deck:
                    built += 1
                    if out_ppt_dir:
                        shutil.copyfile(deck, os.path.join(out_ppt_dir,
                                                          dest_name))
                    print(f"  → {dest_name}", flush=True)
                else:
                    empty += 1
                    print("  → 没有可导出的课件", flush=True)
            except Exception:                           # noqa: BLE001
                failed += 1
                print("  → 失败：", flush=True)
                traceback.print_exc()
            finally:
                scheduler.image_cache.discard(sub_id)
    finally:
        scheduler.shutdown()

    md = {"filename": "-", "lecture_count": 0}
    if args.out_dir:
        md = write_course_markdown(db, course_id, args.out_dir)

    elapsed = time.time() - t0
    print(f"\n[Done] {course_title}: 新建 {built}，跳过 {skipped}（分支已有），"
          f"复用 OCR {reused}，无课件 {empty}，失败 {failed}，"
          f"耗时 {elapsed / 60:.1f} 分钟", flush=True)
    if args.out_dir:
        print(f"[Done] {md['filename']}（{md['lecture_count']} 讲）→ "
              f"{args.out_dir}", flush=True)
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
