"""
Build the public GitHub Pages site.

Only HTML report pages are copied into site/. Data files, workbooks, source
code, and CSV exports are intentionally excluded from the published artifact.
"""
from __future__ import annotations

import glob
import html
import os
import re
import shutil
from dataclasses import dataclass
from datetime import datetime


BASE = os.path.dirname(os.path.abspath(__file__))
SITE_DIR = os.path.join(BASE, "site")
DATE_RE = re.compile(r"_(\d{8})(?:_(AM|PM))?\.html$")


@dataclass(frozen=True)
class Page:
    filename: str
    date_key: str
    slot: str

    @property
    def label(self) -> str:
        try:
            dt = datetime.strptime(self.date_key, "%Y%m%d")
            date_text = dt.strftime("%b %-d, %Y")
        except ValueError:
            date_text = self.date_key
        return f"CEO Brief - {date_text} {self.slot}".strip()


def collect_pages() -> list[Page]:
    pages: list[Page] = []
    for path in glob.glob(os.path.join(BASE, "CEO_Brief_*.html")):
        filename = os.path.basename(path)
        match = DATE_RE.search(filename)
        date_key = match.group(1) if match else "00000000"
        slot = match.group(2) or ""
        pages.append(Page(filename=filename, date_key=date_key, slot=slot))
    return sorted(pages, key=lambda page: (page.date_key, page.slot, page.filename), reverse=True)


def copy_html_pages(pages: list[Page]) -> None:
    if os.path.exists(SITE_DIR):
        shutil.rmtree(SITE_DIR)
    os.makedirs(SITE_DIR, exist_ok=True)

    for page in pages:
        with open(os.path.join(BASE, page.filename), encoding="utf-8") as src:
            content = src.read()
        content = (
            content
            .replace("SYTE Corp · Business Development", "Business Development")
            .replace("SYTE Corp — CEO Brief", "Executive Brief")
            .replace("SYTE BD Pipeline", "Federal Opportunity Pipeline")
            .replace("SYTE Corp", "Company")
            .replace("SYTE", "Company")
        )
        with open(os.path.join(SITE_DIR, page.filename), "w", encoding="utf-8") as dest:
            dest.write(content)

    with open(os.path.join(SITE_DIR, ".nojekyll"), "w", encoding="utf-8") as f:
        f.write("")


def build_index(pages: list[Page]) -> None:
    latest = pages[0].filename if pages else ""
    latest_button = (
        f'<a class="primary" href="{html.escape(latest)}">View latest CEO brief</a>'
        if latest
        else '<span class="primary disabled">No CEO brief has been generated yet</span>'
    )
    links = "\n".join(
        f'<a class="report-link" href="{html.escape(page.filename)}"><span>{html.escape(page.label)}</span><em>HTML</em></a>'
        for page in pages
    ) or '<p class="empty">No CEO brief pages found.</p>'

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    latest_label = html.escape(pages[0].label if pages else "Waiting for first CEO brief")
    page_count = len(pages)
    content = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Federal Opportunity Pipeline</title>
  <style>
    * {{ box-sizing: border-box; }}
    html {{ scroll-behavior: smooth; }}
    body {{
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: #f4f7fb;
      color: #152033;
      line-height: 1.5;
    }}
    a {{ color: inherit; }}
    .hero {{
      min-height: 76vh;
      display: grid;
      align-items: center;
      background:
        linear-gradient(110deg, rgba(18, 32, 54, 0.96), rgba(18, 32, 54, 0.88)),
        repeating-linear-gradient(90deg, rgba(255,255,255,0.08) 0 1px, transparent 1px 80px),
        repeating-linear-gradient(0deg, rgba(255,255,255,0.06) 0 1px, transparent 1px 80px);
      color: white;
      padding: 22px 0 28px;
    }}
    .wrap {{
      width: min(1120px, calc(100% - 32px));
      margin: 0 auto;
    }}
    nav {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 18px;
      margin-bottom: 34px;
      font-size: 13px;
      font-weight: 700;
    }}
    .brand {{ letter-spacing: 0.8px; text-transform: uppercase; }}
    .nav-links {{
      display: flex;
      gap: 18px;
      color: #bfd4ef;
    }}
    .nav-links a {{ text-decoration: none; }}
    .hero-grid {{
      display: grid;
      grid-template-columns: minmax(0, 1fr) minmax(340px, 450px);
      gap: 34px;
      align-items: center;
    }}
    .eyebrow {{
      color: #91c4ff;
      font-size: 12px;
      font-weight: 700;
      letter-spacing: 1.2px;
      text-transform: uppercase;
    }}
    h1 {{
      margin: 8px 0 14px;
      font-size: clamp(40px, 7vw, 76px);
      line-height: 1.1;
      max-width: 780px;
    }}
    .summary {{
      max-width: 700px;
      color: #d9e4f2;
      margin: 0 0 24px;
      font-size: 18px;
    }}
    .hero-actions {{ display: flex; flex-wrap: wrap; gap: 12px; }}
    .primary {{
      display: inline-flex;
      align-items: center;
      min-height: 44px;
      padding: 0 16px;
      border-radius: 8px;
      background: #2e86ab;
      color: white;
      text-decoration: none;
      font-weight: 700;
    }}
    .secondary {{
      display: inline-flex;
      align-items: center;
      min-height: 44px;
      padding: 0 16px;
      border: 1px solid rgba(255,255,255,0.28);
      border-radius: 8px;
      color: white;
      text-decoration: none;
      font-weight: 700;
    }}
    .primary.disabled {{
      background: #64748b;
    }}
    .brief-preview {{
      background: #ffffff;
      color: #152033;
      border-radius: 8px;
      overflow: hidden;
      box-shadow: 0 22px 70px rgba(0,0,0,0.28);
    }}
    .preview-head {{
      background: #2e86ab;
      color: white;
      padding: 14px 16px;
    }}
    .preview-head strong {{ display: block; font-size: 18px; }}
    .preview-head span {{ display: block; color: #d9eef7; font-size: 13px; margin-top: 2px; }}
    .metric-grid {{
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      border-bottom: 1px solid #e5ebf3;
    }}
    .metric {{
      padding: 15px 12px;
      border-right: 1px solid #e5ebf3;
    }}
    .metric:last-child {{ border-right: 0; }}
    .metric b {{ display: block; font-size: 24px; color: #e94f37; line-height: 1; }}
    .metric span {{ display: block; color: #637083; font-size: 11px; margin-top: 5px; text-transform: uppercase; }}
    .pipeline-visual {{ padding: 16px; }}
    .stage {{
      display: grid;
      grid-template-columns: 92px 1fr;
      gap: 12px;
      align-items: center;
      padding: 10px 0;
      border-bottom: 1px solid #edf1f7;
    }}
    .stage:last-child {{ border-bottom: 0; }}
    .stage code {{
      display: inline-flex;
      justify-content: center;
      padding: 5px 8px;
      border-radius: 999px;
      background: #eff7fb;
      color: #2e86ab;
      font-weight: 800;
      font-family: inherit;
      font-size: 12px;
    }}
    .stage span {{ color: #4d5b6f; font-size: 14px; }}
    main {{
      padding: 36px 0 52px;
    }}
    .section {{
      margin-top: 36px;
    }}
    .split {{
      display: grid;
      grid-template-columns: minmax(0, 0.95fr) minmax(0, 1.05fr);
      gap: 28px;
      align-items: start;
    }}
    h2 {{
      font-size: 24px;
      margin: 0 0 12px;
      color: #152033;
      letter-spacing: 0;
    }}
    .section-label {{
      font-size: 12px;
      color: #e94f37;
      text-transform: uppercase;
      letter-spacing: 0.8px;
      font-weight: 800;
      margin-bottom: 6px;
    }}
    .body-copy {{
      color: #4d5b6f;
      margin: 0;
      max-width: 680px;
    }}
    .feature-grid {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      margin-top: 16px;
    }}
    .feature, .output {{
      background: white;
      border: 1px solid #dbe3ef;
      border-radius: 8px;
      padding: 16px;
    }}
    .feature b, .output b {{ display: block; margin-bottom: 6px; }}
    .feature span, .output span {{ color: #637083; font-size: 14px; }}
    .output-list {{
      display: grid;
      gap: 10px;
    }}
    .output {{
      display: grid;
      grid-template-columns: 82px 1fr;
      gap: 12px;
      align-items: start;
    }}
    .tag {{
      display: inline-flex;
      justify-content: center;
      padding: 5px 8px;
      border-radius: 999px;
      background: #fff1ed;
      color: #e94f37;
      font-weight: 800;
      font-size: 12px;
    }}
    .report-list {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
      gap: 10px;
    }}
    .report-link {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 14px;
      min-height: 56px;
      padding: 13px 14px;
      border: 1px solid #dbe3ef;
      border-radius: 8px;
      background: white;
      color: #17233d;
      text-decoration: none;
      font-weight: 650;
    }}
    .report-link:hover {{
      border-color: #2e86ab;
    }}
    .report-link em {{
      color: #637083;
      font-size: 12px;
      font-style: normal;
      white-space: nowrap;
    }}
    .empty, footer {{
      color: #64748b;
      font-size: 13px;
    }}
    footer {{
      margin-top: 28px;
    }}
    @media (max-width: 820px) {{
      .hero-grid, .split, .feature-grid {{
        grid-template-columns: 1fr;
      }}
      .hero {{ min-height: auto; }}
      nav {{ align-items: flex-start; }}
      .nav-links {{ flex-wrap: wrap; justify-content: flex-end; }}
    }}
    @media (max-width: 520px) {{
      .metric-grid {{ grid-template-columns: 1fr; }}
      .metric {{ border-right: 0; border-bottom: 1px solid #e5ebf3; }}
      .metric:last-child {{ border-bottom: 0; }}
      .output {{ grid-template-columns: 1fr; }}
      h1 {{ font-size: 38px; }}
    }}
  </style>
</head>
<body>
  <header class="hero">
    <div class="wrap">
      <nav>
        <div class="brand">BD Automation</div>
        <div class="nav-links">
          <a href="#workflow">Workflow</a>
          <a href="#outputs">Outputs</a>
          <a href="#reports">Reports</a>
        </div>
      </nav>
      <div class="hero-grid">
        <div>
          <div class="eyebrow">SAM.gov business development automation</div>
          <h1>Federal Opportunity Pipeline</h1>
          <p class="summary">A cloud workflow that turns federal opportunity noise into a focused executive briefing, making it easier to see what is new, urgent, and worth action.</p>
          <div class="hero-actions">
            {latest_button}
            <a class="secondary" href="#workflow">See how it works</a>
          </div>
        </div>
        <div class="brief-preview" aria-label="Pipeline preview">
          <div class="preview-head">
            <strong>{latest_label}</strong>
            <span>Web-only public preview</span>
          </div>
          <div class="metric-grid">
            <div class="metric"><b>{page_count}</b><span>Briefs</span></div>
            <div class="metric"><b>2x</b><span>Daily runs</span></div>
            <div class="metric"><b>0</b><span>CSV files published</span></div>
          </div>
          <div class="pipeline-visual">
            <div class="stage"><code>FETCH</code><span>Pull incremental SAM.gov opportunities.</span></div>
            <div class="stage"><code>SCORE</code><span>Rank opportunities for fit, urgency, and set-aside value.</span></div>
            <div class="stage"><code>BRIEF</code><span>Generate a concise web report for review.</span></div>
            <div class="stage"><code>PUBLISH</code><span>Deploy only HTML pages to GitHub Pages.</span></div>
          </div>
        </div>
      </div>
    </div>
  </header>
  <main class="wrap">
    <section class="section split" id="workflow">
      <div>
        <div class="section-label">Purpose</div>
        <h2>Built to reduce BD scanning time.</h2>
        <p class="body-copy">The pipeline watches SAM.gov for relevant opportunities, keeps a dated history, and converts raw procurement data into decision-ready briefings. It is designed for a practical weekly rhythm: find the right items faster, review urgency clearly, and keep the team aligned.</p>
      </div>
      <div class="feature-grid">
        <div class="feature"><b>Opportunity intake</b><span>Incremental SAM.gov pulls keep the active pipeline current.</span></div>
        <div class="feature"><b>Priority signal</b><span>Deadline, notice type, set-aside, and fit are surfaced quickly.</span></div>
        <div class="feature"><b>Executive view</b><span>The CEO brief highlights the work that deserves attention first.</span></div>
      </div>
    </section>

    <section class="section split" id="outputs">
      <div>
        <div class="section-label">Publishing model</div>
        <h2>Public site, controlled contents.</h2>
        <p class="body-copy">The repository can generate CSV, Markdown, Excel, and HTML internally. This page publishes only web report files. Excel workbooks, CSV exports, source code, and secrets are excluded from the GitHub Pages artifact.</p>
      </div>
      <div class="output-list">
        <div class="output"><span class="tag">HTML</span><div><b>Published</b><span>CEO decision briefs are available as web pages.</span></div></div>
        <div class="output"><span class="tag">DATA</span><div><b>Not published</b><span>CSV snapshots and Excel reports stay out of the public Pages site.</span></div></div>
        <div class="output"><span class="tag">AUTO</span><div><b>Updated by Actions</b><span>The site refreshes after the scheduled pipeline or manual workflow run.</span></div></div>
      </div>
    </section>

    <section class="section" id="reports">
      <div class="section-label">Live web reports</div>
      <h2>CEO Briefs</h2>
      <div class="report-list">
        {links}
      </div>
    </section>
    <footer>Generated {html.escape(generated_at)}. Excel and CSV files are not published on this site.</footer>
  </main>
</body>
</html>
"""
    with open(os.path.join(SITE_DIR, "index.html"), "w", encoding="utf-8") as f:
        f.write(content)


def main() -> None:
    pages = collect_pages()
    copy_html_pages(pages)
    build_index(pages)


if __name__ == "__main__":
    main()
