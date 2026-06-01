"""
generate_ceo_brief.py — SYTE Corp BD Pipeline
Generates a mobile-friendly HTML CEO Decision Brief.

Shows top 5-8 opportunities with:
  - AI score + reason (when ANTHROPIC_API_KEY is set)
  - Action recommendation (BID NOW / BID / REVIEW / RESPOND TO RFI / WATCH)
  - Best-matching SYTE past performance project
  - Direct SAM.gov link

Output: CEO_Brief_YYYYMMDD.html
"""

import csv
import os
import glob
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ── SYTE Past Performance (from capability documents) ─────────────────────────
PAST_PERFORMANCE = [
    {"title": "Gary Airport Sanitary Sewer Improvements",
     "client": "USACE Chicago", "value": "$1.77M",
     "keywords": ["sewer", "sanitary", "gravity", "pump station", "manhole", "directional drill", "utility"]},
    {"title": "Morgan Shoals Erosion Protection",
     "client": "USACE Chicago", "value": "$2.05M",
     "keywords": ["erosion", "riprap", "shoreline", "lake", "concrete", "protection", "scour"]},
    {"title": "West Basin Scour Repairs",
     "client": "USACE Memphis", "value": "$4.65M",
     "keywords": ["scour", "levee", "riprap", "excavation", "erosion berm", "turfing", "bank"]},
    {"title": "Ste. Genevieve Levee Repairs",
     "client": "USACE St. Louis", "value": "$2.71M",
     "keywords": ["levee", "repair", "stone revetment", "embankment", "compacted fill", "flood control"]},
    {"title": "Fort Chartres Flood Recovery & Levee Repairs",
     "client": "USACE St. Louis", "value": "$3.11M",
     "keywords": ["flood", "levee", "earthwork", "pipe repair", "slip-lining", "drainage", "recovery"]},
    {"title": "Oldtown Seepage Remediation",
     "client": "USACE Memphis", "value": "$4.59M",
     "keywords": ["seepage", "berm", "compacted fill", "gravel", "aggregate", "drainage", "remediation"]},
    {"title": "Commerce MS / Norfolk Clack Relief Wells",
     "client": "USACE Memphis", "value": "$2.3–2.8M",
     "keywords": ["seepage", "relief well", "drilling", "well installation", "gravel pack", "pump test"]},
    {"title": "Streambank Stabilization — Forest County MS",
     "client": "USDA", "value": "$1.02M",
     "keywords": ["streambank", "bank stabilization", "riprap", "channel", "ditch", "concrete lining"]},
    {"title": "AMLD Emissions Survey",
     "client": "Southern Company Gas", "value": "Multi-year",
     "keywords": ["methane", "gas", "leak detection", "emissions", "pipeline survey", "amld", "picarro"]},
    {"title": "Design-Build Site Prep — BEP Intaglio Presses",
     "client": "Bureau of Engraving & Printing", "value": "$3.81M",
     "keywords": ["design-build", "site preparation", "industrial", "equipment installation", "electrical", "ductwork"]},
    {"title": "HVAC & Electrical Systems Repair",
     "client": "USACE Rock Island", "value": "$2.64M",
     "keywords": ["hvac", "electrical", "mechanical", "facility repair", "building", "upgrade"]},
    {"title": "Bois Brule Road Raise & Drainage",
     "client": "USACE St. Louis", "value": "$1.79M",
     "keywords": ["road", "drainage", "culvert", "fill", "ditch", "pavement", "highway", "bridge"]},
]


# ── Helpers ───────────────────────────────────────────────────────────────────

def _parse_date(s):
    if not s or str(s).strip() in ("", "nan"):
        return None
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d")
    except ValueError:
        return None

def _days_left(dt, today):
    if dt is None:
        return None
    return (dt.replace(hour=0,minute=0,second=0,microsecond=0) -
            today.replace(hour=0,minute=0,second=0,microsecond=0)).days

def _is_sdvosb(row):
    desc = str(row.get("typeOfSetAsideDescription", "")).lower()
    return "veteran" in desc or "sdvosb" in desc

def _fmt_rom(s):
    if not s or str(s).strip() in ("", "nan"):
        return "—"
    return str(s).split("|")[0].strip()[:30]

def _shorten(text, n):
    return str(text)[:n] + "…" if len(str(text)) > n else str(text)

def _match_past_perf(title: str, dept: str):
    text = f"{title} {dept}".lower()
    best, best_score = None, 0
    for pp in PAST_PERFORMANCE:
        score = sum(1 for kw in pp["keywords"] if kw.lower() in text)
        if score > best_score:
            best_score, best = score, pp
    return best if best_score > 0 else None

def _action(row) -> tuple:
    """Returns (label, color_hex) for the action badge."""
    is_sol   = "solicitation" in str(row.get("type","")).lower() and \
               "presolicitation" not in str(row.get("type","")).lower()
    is_presol = "presolicitation" in str(row.get("type","")).lower()
    sdvosb   = row["_sdvosb"]
    dl       = row["_dl"]
    score    = row.get("_llm_score")

    if is_sol:
        if dl is not None and dl <= 7:
            return "🎯  BID NOW", "#E94F37"
        if sdvosb and (score is None or score >= 6):
            return "🎯  BID — SDVOSB", "#2E86AB"
        if score is not None and score >= 8:
            return "🎯  BID — Strong Fit", "#44BBA4"
        if score is not None and score >= 6:
            return "📋  REVIEW", "#F18F01"
        return "📋  REVIEW", "#718096"
    if is_presol:
        if sdvosb:
            return "👀  RESPOND TO RFI", "#F18F01"
        return "📌  WATCH", "#718096"
    return "📋  REVIEW", "#718096"

def _score_color(s):
    if s is None: return "#718096"
    if s >= 8: return "#44BBA4"
    if s >= 6: return "#F18F01"
    return "#E94F37"

def _select_opportunities(rows):
    """Pick top 5-8 opportunities for the CEO brief."""
    has_scores = any(r.get("_llm_score") is not None for r in rows)

    solicitations = [r for r in rows
                     if "solicitation" in str(r.get("type","")).lower()
                     and "presolicitation" not in str(r.get("type","")).lower()]
    presols = [r for r in rows
               if "presolicitation" in str(r.get("type","")).lower()]

    if has_scores:
        # With AI scores: filter by score then urgency
        sol_scored = sorted(
            [r for r in solicitations if r.get("_llm_score") is not None],
            key=lambda r: (not r["_sdvosb"], -(r["_llm_score"] or 0),
                           r["_dl"] is None, r["_dl"] or 9999)
        )
        sol_unscored = sorted(
            [r for r in solicitations if r.get("_llm_score") is None and r["_sdvosb"]],
            key=lambda r: (r["_dl"] is None, r["_dl"] or 9999)
        )
        top_sols = [r for r in sol_scored if (r["_llm_score"] or 0) >= 6][:5]
        if len(top_sols) < 3:
            top_sols = sol_scored[:5]
        top_sols = (top_sols + sol_unscored)[:5]

        top_presols = sorted(
            [r for r in presols if r["_sdvosb"] or (r.get("_llm_score") or 0) >= 7],
            key=lambda r: (not r["_sdvosb"], -(r.get("_llm_score") or 0))
        )[:3]
    else:
        # Without AI scores: urgency + SDVOSB + relevance
        top_sols = sorted(
            solicitations,
            key=lambda r: (not r["_sdvosb"], r["_dl"] is None, r["_dl"] or 9999)
        )[:5]
        top_presols = sorted(
            [r for r in presols if r["_sdvosb"]],
            key=lambda r: (r["_dl"] is None, r["_dl"] or 9999)
        )[:3]

    combined = top_sols + top_presols
    return combined[:8]


# ── HTML generation ───────────────────────────────────────────────────────────

def _card_html(row, idx):
    title    = _shorten(row.get("title", "—"), 90)
    sol_num  = row.get("solicitationNumber", "—")
    dept     = row.get("department", row.get("fullParentPathName", "—"))
    dept     = _shorten(str(dept).split(".")[0].strip().title(), 50)
    location = row.get("location", "—")
    naics    = row.get("naicsCode", "—")
    rom      = _fmt_rom(row.get("ROM", ""))
    # Best available link: direct SAM workspace URL > source_link > SAM search fallback
    link = row.get("uiLink", "") or row.get("source_link", "")
    if not link or str(link) in ("", "nan"):
        from urllib.parse import quote as _q
        _sol = str(row.get("solicitationNumber", "")).strip()
        if _sol:
            link = f"https://sam.gov/search/?keywords={_q(_sol)}&index=opp"
        else:
            link = ""
    sdvosb   = row["_sdvosb"]
    dl       = row["_dl"]
    dl_dt    = row["_dl_dt"]
    ai_score = row.get("_llm_score")
    ai_reason = row.get("LLM_Reason", "") or ""
    notice_type = row.get("type", "—")

    action_label, action_color = _action(row)
    pp = _match_past_perf(row.get("title",""), row.get("department",""))

    # Deadline display
    if dl_dt:
        dl_str = dl_dt.strftime("%b %d, %Y")
        if dl is not None and dl <= 7:
            dl_color, dl_bg = "#FFFFFF", "#E94F37"
        elif dl is not None and dl <= 14:
            dl_color, dl_bg = "#FFFFFF", "#E07B3A"
        elif dl is not None and dl <= 30:
            dl_color, dl_bg = "#1B2A4A", "#FEF3C7"
        else:
            dl_color, dl_bg = "#374151", "#F3F4F6"
        dl_badge = (f'<span class="badge" style="background:{dl_bg};color:{dl_color}">'
                    f'📅 {dl_str}'
                    + (f' · {dl}d' if dl is not None else '')
                    + '</span>')
    else:
        dl_badge = '<span class="badge" style="background:#F3F4F6;color:#718096">📅 TBD</span>'

    sdvosb_badge = ('<span class="badge" style="background:#FEF3C7;color:#92400E;'
                    'font-weight:700">⭐ SDVOSB</span>') if sdvosb else ""

    ai_block = ""
    if ai_score is not None:
        sc = _score_color(ai_score)
        ai_block = f"""
        <div class="ai-row">
          <span class="ai-score" style="color:{sc}">AI {ai_score}/10</span>
          {"<span class='ai-reason'>&ldquo;" + ai_reason + "&rdquo;</span>" if ai_reason else ""}
        </div>"""
    elif not any(r.get("_llm_score") is not None for r in [row]):
        ai_block = '<div class="ai-row"><span class="ai-score" style="color:#CBD5E0">AI scoring not enabled</span></div>'

    pp_block = ""

    platform = row.get("source_platform", "SAM.gov")
    link_label = f"View on {platform} →"
    link_html = (f'<a href="{link}" target="_blank" class="sam-link">{link_label}</a>'
                 if link and link not in ("", "nan") else "")

    border_color = action_color

    return f"""
  <div class="card" style="border-left: 5px solid {border_color}">
    <div class="card-top">
      <span class="action-badge" style="background:{action_color}">{action_label}</span>
      <span class="notice-type">{notice_type}</span>
      {sdvosb_badge}
    </div>
    <div class="card-title">{title}</div>
    <div class="card-meta">
      <span class="meta-item">🏛 {dept}</span>
      <span class="meta-item">📍 {location}</span>
      <span class="meta-item">🏗 NAICS {naics}</span>
      {"<span class='meta-item'>💰 " + rom + "</span>" if rom != "—" else ""}
    </div>
    <div class="card-badges">
      {dl_badge}
    </div>
    {ai_block}
    {pp_block}
    <div class="card-footer">
      <span class="sol-num">#{sol_num}</span>
      {link_html}
    </div>
  </div>"""


def _build_html(rows, today, total_rows, stats):
    date_str  = today.strftime("%B %d, %Y")
    date_file = today.strftime("%Y-%m-%d")
    n_opps    = len(rows)
    n_sol     = stats["n_sol"]
    n_urgent  = stats["n_urgent"]
    n_sdvosb  = stats["n_sdvosb"]
    pipeline  = stats["pipeline"]

    def fmt_pipeline(v):
        if v >= 1e9: return f"${v/1e9:.1f}B"
        if v >= 1e6: return f"${v/1e6:.1f}M"
        if v >= 1e3: return f"${v/1e3:.0f}K"
        return f"${v:,.0f}"

    cards_html = "\n".join(_card_html(r, i) for i, r in enumerate(rows))

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SYTE Corp — CEO Brief {date_file}</title>
  <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: #F5F7FA; color: #1B2A4A; font-size: 15px; }}

    /* ── Header ── */
    .header {{ background: #1B2A4A; color: white; padding: 20px 24px 16px; }}
    .header-top {{ display: flex; justify-content: space-between; align-items: flex-start;
                   flex-wrap: wrap; gap: 8px; }}
    .brand {{ font-size: 11px; font-weight: 700; letter-spacing: 1.5px;
              color: #93C5FD; text-transform: uppercase; margin-bottom: 4px; }}
    .header-title {{ font-size: 22px; font-weight: 700; line-height: 1.2; }}
    .header-date {{ font-size: 13px; color: #93C5FD; margin-top: 2px; }}
    .header-note {{ font-size: 11px; color: #64748B; margin-top: 8px; font-style: italic; }}

    /* ── KPI bar ── */
    .kpi-bar {{ background: #2E86AB; display: flex; flex-wrap: wrap; }}
    .kpi {{ flex: 1; min-width: 80px; padding: 12px 16px; text-align: center;
            border-right: 1px solid rgba(255,255,255,0.15); }}
    .kpi:last-child {{ border-right: none; }}
    .kpi-value {{ font-size: 22px; font-weight: 700; color: white; line-height: 1; }}
    .kpi-label {{ font-size: 10px; color: rgba(255,255,255,0.75); text-transform: uppercase;
                  letter-spacing: 0.5px; margin-top: 3px; }}

    /* ── Section ── */
    .section-header {{ background: white; margin: 16px 16px 0;
                       border-radius: 10px 10px 0 0; padding: 14px 18px 10px;
                       border-bottom: 2px solid #F5F7FA; }}
    .section-title {{ font-size: 13px; font-weight: 700; color: #E94F37;
                      text-transform: uppercase; letter-spacing: 0.8px; }}
    .section-sub {{ font-size: 12px; color: #718096; margin-top: 2px; }}

    /* ── Cards ── */
    .cards-wrap {{ margin: 0 16px 16px; display: flex; flex-direction: column; gap: 0; }}
    .card {{ background: white; padding: 16px 18px; border-bottom: 1px solid #F0F4F8;
             transition: background 0.1s; }}
    .card:last-child {{ border-radius: 0 0 10px 10px; border-bottom: none; }}

    .card-top {{ display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-bottom: 8px; }}
    .action-badge {{ font-size: 11px; font-weight: 700; color: white;
                     padding: 3px 10px; border-radius: 20px; letter-spacing: 0.3px; }}
    .notice-type {{ font-size: 11px; color: #718096; background: #F3F4F6;
                    padding: 2px 8px; border-radius: 10px; }}
    .badge {{ font-size: 11px; padding: 3px 8px; border-radius: 10px; font-weight: 500; }}

    .card-title {{ font-size: 15px; font-weight: 600; color: #1B2A4A;
                   line-height: 1.4; margin-bottom: 8px; }}

    .card-meta {{ display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 8px; }}
    .meta-item {{ font-size: 12px; color: #4A5568; }}

    .card-badges {{ display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 10px; }}

    .ai-row {{ display: flex; align-items: baseline; gap: 10px; margin-bottom: 7px;
               flex-wrap: wrap; }}
    .ai-score {{ font-size: 13px; font-weight: 700; white-space: nowrap; }}
    .ai-reason {{ font-size: 12px; color: #4A5568; font-style: italic; }}

.card-footer {{ display: flex; justify-content: space-between; align-items: center;
                    margin-top: 4px; }}
    .sol-num {{ font-size: 11px; color: #A0AEC0; font-family: monospace; }}
    .sam-link {{ font-size: 12px; color: #2E86AB; text-decoration: none; font-weight: 600; }}
    .sam-link:hover {{ text-decoration: underline; }}

    /* ── Footer ── */
    .footer {{ text-align: center; padding: 20px 16px; font-size: 11px; color: #A0AEC0; }}

    /* ── Desktop ── */
    @media (min-width: 700px) {{
      body {{ max-width: 680px; margin: 0 auto; box-shadow: 0 0 40px rgba(0,0,0,0.08); }}
      .header {{ padding: 24px 28px 20px; }}
      .header-title {{ font-size: 26px; }}
      .section-header {{ margin: 20px 20px 0; }}
      .cards-wrap {{ margin: 0 20px 20px; }}
    }}

    @media print {{
      body {{ background: white; }}
      .sam-link {{ color: #2E86AB; }}
    }}
  </style>
</head>
<body>

<div class="header">
  <div class="header-top">
    <div>
      <div class="brand">SYTE Corp · Business Development</div>
      <div class="header-title">CEO Decision Brief</div>
      <div class="header-date">{date_str}</div>
    </div>
  </div>
  <div class="header-note">
    Showing {n_opps} of {total_rows} active opportunities · Full report: SAM_Opportunities_Report_{today.strftime('%Y%m%d')}_{"AM" if today.hour < 12 else "PM"}.xlsx
  </div>
</div>

<div class="kpi-bar">
  <div class="kpi">
    <div class="kpi-value">{total_rows}</div>
    <div class="kpi-label">Total Opps</div>
  </div>
  <div class="kpi">
    <div class="kpi-value">{n_sol}</div>
    <div class="kpi-label">Solicitations</div>
  </div>
  <div class="kpi">
    <div class="kpi-value" style="color:#FCA5A5">{n_urgent}</div>
    <div class="kpi-label">Urgent ≤14d</div>
  </div>
  <div class="kpi">
    <div class="kpi-value" style="color:#FDE68A">{n_sdvosb}</div>
    <div class="kpi-label">SDVOSB</div>
  </div>
  <div class="kpi">
    <div class="kpi-value">{fmt_pipeline(pipeline) if pipeline else "—"}</div>
    <div class="kpi-label">Pipeline ROM</div>
  </div>
</div>

<div class="section-header">
  <div class="section-title">⚡ {n_opps} Opportunities — Your Decision Needed</div>
  <div class="section-sub">Sorted by urgency · SDVOSB highlighted · AI score when available</div>
</div>
<div class="cards-wrap">
{cards_html}
</div>

<div class="footer">
  Auto-generated by SYTE BD Pipeline · {today.strftime('%Y-%m-%d %H:%M')} ·
  Full pipeline: {total_rows} active opportunities
</div>

</body>
</html>"""


# ── Main ──────────────────────────────────────────────────────────────────────

def main(csv_path=None, today=None):
    if today is None:
        today = datetime.now()

    base = os.path.dirname(os.path.abspath(__file__))

    if csv_path is None:
        pattern = os.path.join(base, "syte_opportunities_????????.csv")
        candidates = sorted(glob.glob(pattern))
        csv_path = candidates[-1] if candidates else os.path.join(base, "syte_opportunities.csv")

    # Load all rows
    all_rows = []
    with open(csv_path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            dl_dt = _parse_date(row.get("responseDeadLine", ""))
            dl    = _days_left(dl_dt, today)
            if dl_dt is not None and dl is not None and dl < 0:
                continue
            row["_dl_dt"]    = dl_dt
            row["_dl"]       = dl
            row["_sdvosb"]   = _is_sdvosb(row)
            raw_score = row.get("LLM_Score", "")
            row["_llm_score"] = int(raw_score) if str(raw_score).strip().isdigit() else None

            # Parse ROM for pipeline total
            try:
                import re
                suffix_map = {"thousand":1e3,"k":1e3,"million":1e6,"m":1e6,"billion":1e9,"b":1e9}
                matches = re.findall(r'\$([\d,]+(?:\.\d+)?)\s*(million|billion|thousand|M|B|K)?',
                                     str(row.get("ROM","")), re.IGNORECASE)
                amounts = []
                for num, suf in matches:
                    v = float(num.replace(",",""))
                    if suf: v *= suffix_map.get(suf.lower(), 1)
                    amounts.append(v)
                row["_rom_val"] = sum(amounts)/len(amounts) if amounts else 0
            except Exception:
                row["_rom_val"] = 0

            all_rows.append(row)

    solicitations = [r for r in all_rows
                     if "solicitation" in str(r.get("type","")).lower()
                     and "presolicitation" not in str(r.get("type","")).lower()]

    stats = {
        "n_sol":    len(solicitations),
        "n_urgent": sum(1 for r in solicitations if r["_dl"] is not None and r["_dl"] <= 14),
        "n_sdvosb": sum(1 for r in all_rows if r["_sdvosb"]),
        "pipeline": sum(r["_rom_val"] for r in all_rows if r["_rom_val"]),
    }

    top_opps = _select_opportunities(all_rows)

    html = _build_html(top_opps, today, len(all_rows), stats)

    slot = "AM" if today.hour < 12 else "PM"
    out_path = os.path.join(base, f"CEO_Brief_{today.strftime('%Y%m%d')}_{slot}.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ CEO Brief generated: {out_path}  ({len(top_opps)} opportunities)")
    return out_path


if __name__ == "__main__":
    main()
