"""Build slide decks for a course's past lectures.

The normal run only builds a deck for the lectures it summarises *in that run*
(``main._drive_lectures``), so lectures summarised before the PPT pipeline
existed never get one.  This script walks every lecture of one course, re-runs
the PPT stages (fetch → dHash dedup → OCR) and writes
``data/ppt_export/<sub_id>.pptx``; the workflow then hands that directory to
``archive_run`` / the ``archive`` branch.

Summaries are *not* produced here — it is a deck backfill.  Lectures that have
no summary yet are reported and skipped; run the ``Single Run`` workflow for
those first.

    python scripts/backfill_pptx.py --list
    python scripts/backfill_pptx.py --course 数字逻辑
    python scripts/backfill_pptx.py --course 30004 --limit 2
"""

import argparse
import os
import sys
import time
import traceback

# Allow importing from the project root when run as
# `python scripts/backfill_pptx.py`
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.data.database import Database  # noqa: E402
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


def main() -> None:
    parser = argparse.ArgumentParser(description="Backfill PPT decks.")
    parser.add_argument("--course", default="",
                        help="course_id or a unique fragment of the title")
    parser.add_argument("--list", action="store_true",
                        help="list stored courses and exit")
    parser.add_argument("--sub-ids", default="",
                        help="comma-separated sub_ids to restrict the run")
    parser.add_argument("--limit", type=int, default=0,
                        help="stop after N lectures (0 = all)")
    args = parser.parse_args()

    db = Database()

    if args.list or not args.course:
        _list_courses(db)
        return

    course = _resolve_course(db, args.course)
    course_id, course_title = course["course_id"], course["title"]
    sub_ids = {s.strip() for s in args.sub_ids.split(",") if s.strip()}

    lectures = _lectures_of(db, course_id, sub_ids or None)
    if not lectures:
        raise SystemExit(f"{course_title} 没有匹配的讲次")
    if args.limit > 0:
        lectures = lectures[:args.limit]

    pending = [l for l in lectures if not l["summary"]]
    if pending:
        print(f"[Warn] {len(pending)}/{len(lectures)} 讲还没有总结"
              f"（如 {pending[0]['sub_id']}）——先跑 Single Run 补齐，"
              f"本次只生成已有总结讲次的课件。", flush=True)

    reporter = Reporter()
    reporter.run_header()

    # Imported here so `--list` works without the network / OCR stack.
    from src.api.icourse import ICourseClient
    from src.pipeline.ppt_pipeline import PPTPipeline
    from src.pipeline.ppt_exporter import build_lecture_pptx
    from src.runtime.scheduler import Scheduler

    vpn = _login()
    client = ICourseClient(vpn)
    scheduler = Scheduler(reporter=reporter)
    pipeline = PPTPipeline(db, scheduler, reporter)

    built = empty = failed = 0
    t0 = time.time()
    try:
        for i, lec in enumerate(lectures, 1):
            sub_id = str(lec["sub_id"])
            print(f"\n[{i}/{len(lectures)}] {course_title} · "
                  f"{lec['sub_title']} ({lec['date']})", flush=True)
            try:
                if not client.check_alive():
                    print("  Session expired, re-logging in...", flush=True)
                    client.vpn = _login()
                    client._userinfo = None

                handle = pipeline.submit(client, course_id, sub_id)
                stats = handle.drain()
                deck = build_lecture_pptx(
                    client, db, sub_id, course_title,
                    lec.get("sub_title") or sub_id, lec.get("date") or "",
                    reporter,
                )
                if deck:
                    built += 1
                    print(f"  → {deck}", flush=True)
                else:
                    empty += 1
                    print(f"  → 没有可导出的课件（保留页面 {stats.done}）",
                          flush=True)
            except Exception:                           # noqa: BLE001
                failed += 1
                print(f"  → 失败：", flush=True)
                traceback.print_exc()
            finally:
                scheduler.image_cache.discard(sub_id)
    finally:
        scheduler.shutdown()

    elapsed = time.time() - t0
    print(f"\n[Done] {course_title}: 生成 {built}，无课件 {empty}，失败 {failed}，"
          f"耗时 {elapsed / 60:.1f} 分钟", flush=True)
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
