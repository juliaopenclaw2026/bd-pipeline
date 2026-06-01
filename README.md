# SYTE BD Pipeline

Cloud-ready SAM.gov business development pipeline for SYTE Corp.

Phase 1 runs the core pipeline in GitHub Actions:

- Fetch incremental SAM.gov opportunities.
- Merge them into `syte_opportunities.csv`.
- Write a dated CSV snapshot.
- Generate Markdown, HTML, and Excel reports.
- Commit outputs back to this private repository.

Pipedrive and the local Flask app are intentionally excluded from this phase.

## Schedule

GitHub Actions cron runs in UTC:

- `30 13 * * 1-5` = 08:30 America/Chicago during daylight saving time.
- `0 20 * * 1-5` = 15:00 America/Chicago during daylight saving time.

When Chicago returns to standard time in November, update the cron entries to:

- `30 14 * * 1-5`
- `0 21 * * 1-5`

## Required Secrets

Set these in GitHub repository settings, or with GitHub CLI:

```bash
gh secret set SAM_API_KEY
gh secret set ANTHROPIC_API_KEY
```

`SAM_API_KEY` is required. `ANTHROPIC_API_KEY` is optional; if it is empty, LLM scoring is skipped automatically.

Do not set `PIPEDRIVE_API_TOKEN` for phase 1.

## Manual Run

From the repository page, open the Actions tab, choose `SYTE BD Pipeline`, then click `Run workflow`.

With GitHub CLI:

```bash
gh workflow run "SYTE BD Pipeline"
```

## Outputs

After each successful run, the workflow commits updated outputs:

- `syte_opportunities.csv`
- `syte_opportunities_YYYYMMDD.csv`
- `briefing_YYYYMMDD*.md`
- `CEO_Brief_YYYYMMDD*.html`
- `SAM_Opportunities_Report_YYYYMMDD*.xlsx`
- `last_run.txt`

The Markdown briefing is the easiest first-phase report to read in GitHub web or the GitHub mobile app.

## Local Development

```bash
python -m venv venv_new
source venv_new/bin/activate
pip install -r requirements.txt
python run.py
```

For a report-only rerun using existing CSV data:

```bash
python run.py --report-only
```

## Migration Notes

This repository is seeded from the June 2026 canonical CSV and `last_run.txt`, so GitHub Actions continues incrementally from the current state.

After the cloud workflow is verified, remove the local cron job to avoid duplicate runs:

```bash
crontab -r
```
