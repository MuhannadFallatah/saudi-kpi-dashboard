# Saudi KPI Data Monitor — GitHub Pages

This repository template provides a static dashboard for all 16 Saudi KPIs,
plus a scheduled GitHub Actions job that extracts available official GASTAT
workbooks and rebuilds the Excel and JSON downloads.

## Hosting and privacy

The repository is intended to be private. The dashboard and files in `docs/`
are published to GitHub Pages and are publicly accessible, as authorized for
this deployment. The Excel and JSON downloads contain the KPI observations.
Do not put credentials, private API keys, or non-public source data in this
repository.

## Deployment

1. Create a **private** GitHub repository named `saudi-kpi-dashboard`.
2. Upload the contents of this folder to the repository's `main` branch.
3. In repository **Settings → Pages**, set the build/deployment source to
   **GitHub Actions**.
4. Run **Actions → Refresh and publish Saudi KPI dashboard → Run workflow**.
5. The workflow refreshes daily at 21:00 UTC (00:00 Riyadh time). It can also be run manually.

The scheduled workflow publishes the report files and dashboard as a static
site. It cannot discover newly released GASTAT report URLs; update the source
register URLs in `Saudi_16_KPI_Data_Mart.xlsx` when GASTAT publishes a new
edition. A Pages deployment is not a continuously running Python service:
JSON endpoints are static files, and there is no POST refresh API.

## Included extractors and status meaning

Automated source workbooks currently cover KPIs 05, 07, 08, 11, 12 and 13
(GASTAT, fixed URLs) and KPI 16 (Ministry of Finance quarterly budget
performance reports, "Non-financial Assets (CAPEX)", quarterly from 2022; the
latest report per fiscal year is discovered from the MOF year pages).
Other KPIs appear in the status table as unavailable until their source routes
and access are verified. KPI 10 (MISA) remains excluded by instruction.
Green indicates a configured extraction succeeded; red indicates a failed or
unavailable connector, no data, or intentional exclusion. Check observation
dates; success does not imply that the source has published a newer period.

The original fixed URLs in the mart are used for GASTAT KPIs. The exporter produces:

- `docs/data/status.json`
- `docs/data/facts.json`
- `docs/data/Saudi_KPI_Extract.xlsx`

## Rebuild locally

```powershell
python -m pip install -r .\requirements.txt
python .\publish_static.py
```

## Access notice

This is static public hosting authorized by the repository owner. The source
repository itself remains private. GitHub account/organization policy may
require enabling Pages for a private repository.
