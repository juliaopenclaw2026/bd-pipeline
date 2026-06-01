"""
generate_briefing.py
每次 run.py 运行后，生成一份中文 markdown 简报：本次新增了什么、什么紧急、SDVOSB 重点。
输出：briefing_YYYYMMDD.md
"""
import csv
import os
from datetime import datetime, timedelta

TODAY = datetime.now()
NAICS_LABELS = {
    "237110": "水/污水管线",
    "237120": "油气管道",
    "237310": "公路/桥梁",
    "237990": "其他重型土建",
}


def _parse_date(s):
    if not s or str(s).strip() in ("", "nan"):
        return None
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d")
    except ValueError:
        return None


def _days_left(dt):
    if dt is None:
        return None
    return (dt.replace(hour=0, minute=0, second=0, microsecond=0) -
            TODAY.replace(hour=0, minute=0, second=0, microsecond=0)).days


def _fmt_rom(s):
    if not s or str(s).strip() in ("", "nan"):
        return "—"
    return str(s).split("|")[0].strip()


def _is_sdvosb(row):
    desc = str(row.get("typeOfSetAsideDescription", "")).lower()
    return "veteran" in desc or "sdvosb" in desc


def _urgency(days):
    if days is None:
        return "—"
    if days <= 7:
        return "🔴 极紧急"
    if days <= 14:
        return "🟠 紧急"
    if days <= 30:
        return "🟡 关注"
    return "⬜ 正常"


def _shorten(text, n=60):
    return text[:n] + "…" if len(text) > n else text


def _naics_label(code):
    return NAICS_LABELS.get(str(code), str(code))


def load_rows(csv_path):
    rows = []
    with open(csv_path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            dl_dt = _parse_date(row.get("responseDeadLine", ""))
            dl = _days_left(dl_dt)
            if dl_dt is not None and dl < 0:
                continue  # 已截止，跳过
            row["_dl_dt"] = dl_dt
            row["_dl"] = dl
            row["_sdvosb"] = _is_sdvosb(row)
            rows.append(row)
    return rows


def _md_table_row(cols):
    return "| " + " | ".join(str(c) for c in cols) + " |"


def _md_table(headers, rows_data):
    lines = []
    lines.append(_md_table_row(headers))
    lines.append(_md_table_row(["---"] * len(headers)))
    for r in rows_data:
        lines.append(_md_table_row(r))
    return "\n".join(lines)


def main(csv_path, new_keys=None, updated_keys=None, today=None):
    global TODAY
    if today is not None:
        TODAY = today

    new_keys = new_keys or set()
    updated_keys = updated_keys or set()

    rows = load_rows(csv_path)
    sol_rows   = [r for r in rows if str(r.get("type","")).lower() == "solicitation"]
    presol_rows= [r for r in rows if str(r.get("type","")).lower() in ("presolicitation","special notice")]

    # 新增条目（从本次拉取中第一次出现的）
    new_rows = [r for r in rows if (r.get("solicitationNumber") or r.get("title","")) in new_keys]

    # 紧急：solicitation 且 ≤14 天
    urgent_rows = sorted(
        [r for r in sol_rows if r["_dl"] is not None and r["_dl"] <= 14],
        key=lambda r: r["_dl"]
    )

    # SDVOSB solicitation
    sdvosb_sol = [r for r in sol_rows if r["_sdvosb"]]

    # 即将截止 15-30 天
    soon_rows = sorted(
        [r for r in sol_rows if r["_dl"] is not None and 15 <= r["_dl"] <= 30],
        key=lambda r: r["_dl"]
    )

    lines = []
    date_str = TODAY.strftime("%Y-%m-%d")
    lines.append(f"# SYTE Corp BD 简报 — {date_str}")
    lines.append("")
    lines.append(f"> 数据来源：SAM.gov ｜ 更新时间：{TODAY.strftime('%Y-%m-%d %H:%M')}")
    lines.append("")

    # ── 摘要 ──────────────────────────────────────────────────────────────────
    lines.append("## 📌 本次更新摘要")
    lines.append("")
    lines.append(f"| 指标 | 数值 |")
    lines.append(f"|---|---|")
    lines.append(f"| 本次**新增**机会 | **{len(new_rows)}** 个 |")
    lines.append(f"| 本次**更新**记录 | {len(updated_keys)} 个 |")
    lines.append(f"| 当前有效机会总数 | {len(rows)} 个 |")
    lines.append(f"| 其中 Solicitation | {len(sol_rows)} 个 |")
    lines.append(f"| 其中 Presolicitation | {len(presol_rows)} 个 |")
    lines.append(f"| 🔴 极紧急（≤7天） | {sum(1 for r in urgent_rows if r['_dl'] <= 7)} 个 |")
    lines.append(f"| 🟠 紧急（8–14天） | {sum(1 for r in urgent_rows if 8 <= r['_dl'] <= 14)} 个 |")
    lines.append(f"| ⭐ SDVOSB Solicitation | {len(sdvosb_sol)} 个 |")
    lines.append("")

    # ── 新增机会 ──────────────────────────────────────────────────────────────
    lines.append("## 🆕 本次新增机会")
    lines.append("")
    if not new_rows:
        lines.append("_本次运行无新增机会（均为历史记录更新）。_")
    else:
        headers = ["类型", "项目名称", "地点", "NAICS", "截止日期", "剩余天数", "ROM估算", "优先级"]
        table_data = []
        for r in sorted(new_rows, key=lambda x: (x["_dl"] is None, x["_dl"] or 9999)):
            dl_str = r["_dl_dt"].strftime("%m/%d/%Y") if r["_dl_dt"] else "—"
            dl_days = f"{r['_dl']}d" if r["_dl"] is not None else "—"
            priority = ("⭐ SDVOSB" if r["_sdvosb"] else "") + _urgency(r["_dl"])
            table_data.append([
                r.get("type","—"),
                _shorten(r.get("title","—"), 55),
                r.get("location","—"),
                _naics_label(r.get("naicsCode","")),
                dl_str,
                dl_days,
                _fmt_rom(r.get("ROM","")),
                priority,
            ])
        lines.append(_md_table(headers, table_data))
    lines.append("")

    # ── 立即行动 ──────────────────────────────────────────────────────────────
    lines.append("## 🔴 立即行动（截止日期 ≤14 天）")
    lines.append("")
    if not urgent_rows:
        lines.append("_目前没有 14 天内截止的 Solicitation。_")
    else:
        headers = ["Sol #", "项目名称", "地点", "截止日期", "剩余天数", "ROM估算", "标志"]
        table_data = []
        for r in urgent_rows:
            dl_str = r["_dl_dt"].strftime("%m/%d/%Y") if r["_dl_dt"] else "—"
            flag = "⭐ SDVOSB" if r["_sdvosb"] else ""
            table_data.append([
                r.get("solicitationNumber","—"),
                _shorten(r.get("title","—"), 50),
                r.get("location","—"),
                dl_str,
                f"**{r['_dl']}d**",
                _fmt_rom(r.get("ROM","")),
                flag,
            ])
        lines.append(_md_table(headers, table_data))
    lines.append("")

    # ── SDVOSB 重点 ───────────────────────────────────────────────────────────
    lines.append("## ⭐ SDVOSB 重点机会（最高优先级）")
    lines.append("")
    if not sdvosb_sol:
        lines.append("_本期无 SDVOSB Solicitation。_")
    else:
        headers = ["Sol #", "项目名称", "地点", "截止日期", "剩余天数", "ROM估算"]
        table_data = []
        for r in sorted(sdvosb_sol, key=lambda x: (x["_dl"] is None, x["_dl"] or 9999)):
            dl_str = r["_dl_dt"].strftime("%m/%d/%Y") if r["_dl_dt"] else "—"
            dl_days = f"{r['_dl']}d" if r["_dl"] is not None else "—"
            table_data.append([
                r.get("solicitationNumber","—"),
                _shorten(r.get("title","—"), 50),
                r.get("location","—"),
                dl_str,
                dl_days,
                _fmt_rom(r.get("ROM","")),
            ])
        lines.append(_md_table(headers, table_data))
    lines.append("")

    # ── 下周关注 ─────────────────────────────────────────────────────────────
    lines.append("## 🟡 下周关注（截止日期 15–30 天）")
    lines.append("")
    if not soon_rows:
        lines.append("_目前没有 15–30 天内截止的 Solicitation。_")
    else:
        headers = ["Sol #", "项目名称", "地点", "截止日期", "剩余天数", "ROM估算"]
        table_data = []
        for r in soon_rows[:10]:
            dl_str = r["_dl_dt"].strftime("%m/%d/%Y") if r["_dl_dt"] else "—"
            table_data.append([
                r.get("solicitationNumber","—"),
                _shorten(r.get("title","—"), 50),
                r.get("location","—"),
                dl_str,
                f"{r['_dl']}d",
                _fmt_rom(r.get("ROM","")),
            ])
        lines.append(_md_table(headers, table_data))
    lines.append("")

    # ── Presolicitation 预警 ──────────────────────────────────────────────────
    top_presols = sorted(presol_rows, key=lambda r: (not r["_sdvosb"], 0))[:8]
    if top_presols:
        lines.append("## 👀 Presolicitation 预警（SDVOSB 优先）")
        lines.append("")
        headers = ["Sol #", "项目名称", "地点", "NAICS", "标志"]
        table_data = []
        for r in top_presols:
            flag = "⭐ SDVOSB" if r["_sdvosb"] else "SBA"
            table_data.append([
                r.get("solicitationNumber","—"),
                _shorten(r.get("title","—"), 50),
                r.get("location","—"),
                _naics_label(r.get("naicsCode","")),
                flag,
            ])
        lines.append(_md_table(headers, table_data))
        lines.append("")

    lines.append("---")
    lines.append(f"_自动生成 by SYTE BD Pipeline ｜ {date_str}_")

    # ── 写文件 ────────────────────────────────────────────────────────────────
    base = os.path.dirname(os.path.abspath(__file__))
    slot = "AM" if TODAY.hour < 12 else "PM"
    out_path = os.path.join(base, f"briefing_{TODAY.strftime('%Y%m%d')}_{slot}.md")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"\n✅ 简报已生成: {out_path}")
    return out_path


if __name__ == "__main__":
    import glob
    base = os.path.dirname(os.path.abspath(__file__))
    pattern = os.path.join(base, "syte_opportunities_????????.csv")
    candidates = sorted(glob.glob(pattern))
    csv_path = candidates[-1] if candidates else os.path.join(base, "syte_opportunities.csv")
    main(csv_path)
