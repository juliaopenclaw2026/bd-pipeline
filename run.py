"""
run.py — SYTE Corp BD Pipeline 一键入口

用法:
  python run.py               完整流程（拉数据 + 生成报告）
  python run.py --data-only   只拉取 SAM.gov 数据
  python run.py --report-only 只生成报告（使用最新已有 CSV）
"""
import sys
import os
import re
import glob
import shutil
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))

def archive_old_files():
    """把早于当前月的带日期输出文件移到 Archive/YYYY-MM/。"""
    now = datetime.now()
    cur_ym = now.year * 100 + now.month  # 如 202606
    date_re = re.compile(r"(\d{4})(\d{2})\d{2}")
    exts = (".xlsx", ".html", ".md", ".csv")
    moved = 0
    for fname in os.listdir(BASE):
        path = os.path.join(BASE, fname)
        if not os.path.isfile(path) or not fname.endswith(exts):
            continue
        m = date_re.search(fname)
        if not m:
            continue  # 无日期的文件（如 canonical csv）跳过
        ym = int(m.group(1)) * 100 + int(m.group(2))
        if ym >= cur_ym:
            continue  # 当前月或未来，保留
        dest_dir = os.path.join(BASE, "Archive", f"{m.group(1)}-{m.group(2)}")
        os.makedirs(dest_dir, exist_ok=True)
        shutil.move(path, os.path.join(dest_dir, fname))
        moved += 1
    if moved:
        print(f"  [Archive] 归档了 {moved} 个往月文件到 Archive/")

def find_latest_csv():
    pattern = os.path.join(BASE, "syte_opportunities_????????.csv")
    files = sorted(glob.glob(pattern))
    if not files:
        fallback = os.path.join(BASE, "syte_opportunities.csv")
        return fallback if os.path.exists(fallback) else None
    return files[-1]

def run_data():
    print("=" * 50)
    print("Step 1: Fetching SAM.gov data (incremental)")
    print("=" * 50)
    import sam_opportunities
    raw = sam_opportunities.get_opportunities(limit=100)
    processed = sam_opportunities.filter_and_sort_opportunities(raw)
    new_keys, updated_keys = sam_opportunities.save_to_csv(processed)
    sam_opportunities.write_last_run_date()
    print(f"Last run date updated to {datetime.now().strftime('%Y-%m-%d')}.\n")
    return new_keys, updated_keys

def run_report():
    print("=" * 50)
    print("Step 2: Generating Excel report")
    print("=" * 50)
    csv_path = find_latest_csv()
    if not csv_path:
        print("ERROR: No syte_opportunities CSV found. Run with --data-only first.")
        sys.exit(1)
    today = datetime.now()
    print(f"Using data: {os.path.basename(csv_path)}")
    import generate_report
    generate_report.main(csv_path, today)

def update_claude_md(today):
    """Update the 上次数据更新 line in CLAUDE.md."""
    claude_md = os.path.join(BASE, "CLAUDE.md")
    if not os.path.exists(claude_md):
        return
    with open(claude_md, encoding="utf-8") as f:
        content = f.read()
    import re
    updated = re.sub(
        r"> \*\*上次数据更新：.*?\*\*",
        f"> **上次数据更新：{today.strftime('%Y-%m-%d')}**",
        content
    )
    with open(claude_md, "w", encoding="utf-8") as f:
        f.write(updated)

def run_score():
    from score_opportunities import refresh_profile, score_all, evolve_keywords

    print("=" * 50)
    print("Step 0: Refreshing SYTE profile from PDFs")
    print("=" * 50)
    refresh_profile(BASE)

    csv_path = find_latest_csv()
    if not csv_path:
        return

    print("=" * 50)
    print("Step 1.5: LLM semantic scoring (Claude Haiku)")
    print("=" * 50)
    score_all(csv_path)

    print("=" * 50)
    print("Step 1.6: Keyword evolution suggestions")
    print("=" * 50)
    evolve_keywords(csv_path)


def main():
    args = sys.argv[1:]
    data_only   = "--data-only"   in args
    report_only = "--report-only" in args

    if data_only and report_only:
        print("ERROR: Cannot use --data-only and --report-only together.")
        sys.exit(1)

    print(f"\n{'='*50}")
    print(f"SYTE Corp BD Pipeline — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"{'='*50}\n")

    archive_old_files()

    new_keys, updated_keys = set(), set()
    if not report_only:
        new_keys, updated_keys = run_data()

    if not data_only:
        run_score()
        run_report()

    print("=" * 50)
    print("Step 3: Generating briefing")
    print("=" * 50)
    csv_path = find_latest_csv()
    if csv_path:
        import generate_briefing
        generate_briefing.main(csv_path, new_keys=new_keys, updated_keys=updated_keys,
                               today=datetime.now())

    print("=" * 50)
    print("Step 4: Generating CEO Decision Brief")
    print("=" * 50)
    csv_path = find_latest_csv()
    if csv_path:
        import generate_ceo_brief
        generate_ceo_brief.main(csv_path, today=datetime.now())

    update_claude_md(datetime.now())
    print("\nDone.")

if __name__ == "__main__":
    main()
