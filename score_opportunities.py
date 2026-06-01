"""
score_opportunities.py — SYTE Corp BD Pipeline
LLM-powered semantic scoring (Claude Haiku) + keyword evolution suggestions.

Steps run by run.py:
  1. refresh_profile(BASE_DIR)   — re-extract syte_profile.md if any PDF changed
  2. score_all(csv_path)         — score every opportunity; write LLM_Score + LLM_Reason
  3. evolve_keywords(csv_path)   — print keyword update suggestions to terminal
"""

import os
import json
import csv
import re
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
PROFILE_PATH = os.path.join(BASE_DIR, "syte_profile.md")
CACHE_PATH   = os.path.join(BASE_DIR, "llm_scores_cache.json")


# ── Anthropic client ──────────────────────────────────────────────────────────

def _get_client():
    from dotenv import load_dotenv
    load_dotenv()
    api_key = os.getenv("ANTHROPIC_API_KEY", "").strip()
    if not api_key:
        return None
    import anthropic
    return anthropic.Anthropic(api_key=api_key)


# ── PDF → profile ─────────────────────────────────────────────────────────────

def _extract_pdf_text(pdf_path: str, max_chars: int = 12000) -> str:
    try:
        import pypdf
        reader = pypdf.PdfReader(pdf_path)
        parts = []
        for page in reader.pages:
            parts.append(page.extract_text() or "")
        return "\n".join(parts)[:max_chars]
    except Exception as e:
        print(f"  [Profile] Could not read {os.path.basename(pdf_path)}: {e}")
        return ""


def refresh_profile(pdf_dir: str):
    """Re-extract syte_profile.md if any PDF in pdf_dir is newer than the profile."""
    pdf_files = sorted(Path(pdf_dir).glob("*.pdf"))
    if not pdf_files:
        return

    profile_mtime = os.path.getmtime(PROFILE_PATH) if os.path.exists(PROFILE_PATH) else 0
    if not any(p.stat().st_mtime > profile_mtime for p in pdf_files):
        return  # all PDFs older than current profile — nothing to do

    client = _get_client()
    if client is None:
        print("  [Profile] ANTHROPIC_API_KEY not set — skipping profile refresh.")
        return

    print(f"  [Profile] New/updated PDF detected. Refreshing syte_profile.md from "
          f"{len(pdf_files)} file(s)...")

    sections = []
    for pdf in pdf_files:
        text = _extract_pdf_text(str(pdf))
        if text:
            sections.append(f"=== {pdf.name} ===\n{text}")

    if not sections:
        print("  [Profile] No text extracted. Skipping.")
        return

    combined = "\n\n".join(sections)

    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1800,
        messages=[{
            "role": "user",
            "content": (
                "Based on the following SYTE Corp documents, write a structured capability "
                "profile in Markdown. Include:\n"
                "1. Company overview (type, certifications, HQ)\n"
                "2. Core capabilities (specific bullet list)\n"
                "3. Top past performance projects (client, value, scope — 10 max)\n"
                "4. Primary agency/client relationships\n"
                "5. Typical contract size range\n"
                "6. Geographic focus areas\n"
                "7. What SYTE does NOT do (helps filter out irrelevant opportunities)\n\n"
                "Be concise and specific. This profile will be used to score federal "
                "contracting opportunities for fit.\n\n"
                f"DOCUMENTS:\n{combined}"
            )
        }]
    )

    content = response.content[0].text.strip()
    with open(PROFILE_PATH, "w", encoding="utf-8") as f:
        f.write(f"# SYTE Corp Capability Profile\n")
        f.write(f"*Auto-generated {datetime.now().strftime('%Y-%m-%d')} from PDFs*\n\n")
        f.write(content)

    print(f"  [Profile] syte_profile.md updated ({len(content)} chars).")


# ── Scoring ───────────────────────────────────────────────────────────────────

def _load_profile() -> str:
    if not os.path.exists(PROFILE_PATH):
        return ""
    with open(PROFILE_PATH, encoding="utf-8") as f:
        return f.read()


def _load_cache() -> dict:
    if not os.path.exists(CACHE_PATH):
        return {}
    with open(CACHE_PATH, encoding="utf-8") as f:
        try:
            return json.load(f)
        except (json.JSONDecodeError, ValueError):
            return {}


def _save_cache(cache: dict):
    with open(CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2, ensure_ascii=False)


def _score_one(client, profile: str, opp: dict) -> dict:
    prompt = (
        "You are a BD analyst for SYTE Corp. Score this federal contracting opportunity "
        "for fit with SYTE Corp (0 = no fit, 10 = perfect fit).\n\n"
        f"SYTE CORP PROFILE:\n{profile}\n\n"
        "OPPORTUNITY:\n"
        f"Title: {opp.get('title', 'N/A')}\n"
        f"Type: {opp.get('type', 'N/A')} | Set-Aside: {opp.get('typeOfSetAsideDescription', 'N/A')}\n"
        f"Agency: {opp.get('department', 'N/A')}\n"
        f"NAICS: {opp.get('naicsCode', 'N/A')}\n"
        f"Location: {opp.get('location', 'N/A')}\n"
        f"ROM: {opp.get('ROM', 'N/A')}\n\n"
        'Return JSON only, no other text: {"score": <integer 0-10>, "reason": "<15 words max, English>"}'
    )
    try:
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=80,
            messages=[{"role": "user", "content": prompt}]
        )
        text = resp.content[0].text.strip()
        m = re.search(r'\{[^}]+\}', text)
        result = json.loads(m.group() if m else text)
        return {"score": int(result.get("score", 0)), "reason": str(result.get("reason", ""))}
    except Exception:
        return {"score": 0, "reason": "scoring unavailable"}


def score_all(csv_path: str):
    """Score all rows in csv_path; write LLM_Score and LLM_Reason columns back."""
    client = _get_client()
    if client is None:
        print("  [Score] ANTHROPIC_API_KEY not set — skipping LLM scoring.")
        return

    profile = _load_profile()
    if not profile:
        print("  [Score] syte_profile.md not found — skipping LLM scoring.")
        return

    with open(csv_path, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    if not rows:
        return

    cache = _load_cache()

    to_score = [
        (row.get("solicitationNumber") or row.get("title", ""), row)
        for row in rows
        if (row.get("solicitationNumber") or row.get("title", "")) not in cache
    ]

    if to_score:
        print(f"  [Score] Scoring {len(to_score)} new opportunities via Claude Haiku "
              f"({len(rows) - len(to_score)} served from cache)...")

        def _task(item):
            key, opp = item
            return key, _score_one(client, profile, opp)

        with ThreadPoolExecutor(max_workers=10) as ex:
            for future in as_completed({ex.submit(_task, item): item for item in to_score}):
                key, result = future.result()
                cache[key] = {**result, "scored_at": datetime.now().strftime("%Y-%m-%d")}

        _save_cache(cache)
    else:
        print(f"  [Score] All {len(rows)} opportunities already cached — no API calls needed.")

    # Write scores back to CSV
    fieldnames = list(rows[0].keys())
    for col in ("LLM_Score", "LLM_Reason"):
        if col not in fieldnames:
            fieldnames.append(col)

    for row in rows:
        key = row.get("solicitationNumber") or row.get("title", "")
        entry = cache.get(key, {})
        row["LLM_Score"]  = entry.get("score", "")
        row["LLM_Reason"] = entry.get("reason", "")

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)

    scored = [r for r in rows if str(r.get("LLM_Score", "")).isdigit()]
    if scored:
        avg = sum(int(r["LLM_Score"]) for r in scored) / len(scored)
        print(f"  [Score] Done. Avg score: {avg:.1f}/10 across {len(scored)} opportunities.")


# ── Keyword evolution ─────────────────────────────────────────────────────────

def _current_keywords() -> list:
    kw_path = os.path.join(BASE_DIR, "sam_opportunities.py")
    if not os.path.exists(kw_path):
        return []
    with open(kw_path, encoding="utf-8") as f:
        content = f.read()
    m = re.search(r'SYTE_KEYWORDS\s*=\s*\[([^\]]+)\]', content, re.DOTALL)
    if not m:
        return []
    return re.findall(r'"([^"]+)"', m.group(1))


def evolve_keywords(csv_path: str):
    """Analyze scored results and print keyword evolution suggestions."""
    client = _get_client()
    if client is None:
        return

    with open(csv_path, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    scored = [r for r in rows if str(r.get("LLM_Score", "")).isdigit()]
    if not scored:
        return

    avg_score = sum(int(r["LLM_Score"]) for r in scored) / len(scored)
    high = [r for r in scored if int(r["LLM_Score"]) >= 7]
    low  = [r for r in scored if int(r["LLM_Score"]) <= 3]

    summary_lines = sorted(
        [(int(r["LLM_Score"]), r.get("naicsCode", ""), r.get("type", ""), r.get("title", ""))
         for r in scored],
        key=lambda x: -x[0]
    )
    summary = "\n".join(
        f"Score {s} | NAICS {n} | {t} | {title[:80]}"
        for s, n, t, title in summary_lines
    )

    current_kw = _current_keywords()

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=400,
        messages=[{
            "role": "user",
            "content": (
                "Analyze scored federal contracting opportunities for SYTE Corp "
                "(SDVOSB, heavy civil/utilities/levee/pipeline construction) and "
                "recommend keyword list updates.\n\n"
                f"Current SYTE_KEYWORDS: {current_kw}\n\n"
                "Scored opportunities this run (score desc):\n"
                f"{summary}\n\n"
                "Respond in this exact format:\n"
                "ADD: word1, word2\n"
                "REMOVE: word3, word4\n"
                "REASON_EN: <one sentence>\n"
                "REASON_ZH: <一句话>"
            )
        }]
    )

    result = response.content[0].text.strip()

    print()
    print("═" * 62)
    print("  Keyword Evolution Suggestions / 关键词进化建议")
    print("═" * 62)
    print(f"  Run stats: {len(scored)} scored | avg {avg_score:.1f}/10 | "
          f"{len(high)} high (≥7) | {len(low)} low (≤3)")
    print()
    for line in result.splitlines():
        print(f"  {line}")
    print()
    print("  To apply: update SYTE_KEYWORDS in sam_opportunities.py (lines 30-33)")
    print("═" * 62)
    print()
