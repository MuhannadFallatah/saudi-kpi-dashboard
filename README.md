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
5. The workflow refreshes daily at 07:07 UTC (10:07 Riyadh time). It can also be run manually from the Actions tab.

The scheduled workflow publishes the report files and dashboard as a static
site. File-based sources are refreshed from their listing pages so each run
can select the newest release: SAMA's latest POS PDFs, GASTAT's latest monthly
ITR workbook for KPIs 08 and 11, and GASTAT's latest monthly STS workbook for
KPI 13. A Pages deployment is not a continuously running Python service:
JSON endpoints are static files, and there is no POST refresh API.
Before refreshing, each successful workflow run archives the latest
pre-refresh full dataset once per Riyadh calendar date under `docs/archive/`.
The compressed snapshot contains the prior KPI status and all facts, is
retained permanently in the repository, and is available from the dashboard's
Archive page. A manual rerun on the same Riyadh date does not replace that
day's first snapshot.
Each workflow refresh also appends its per-KPI outcomes to
`docs/refresh-history/`, including manual and push-triggered runs. These
timestamped records are retained permanently and are separate from the
once-daily fact archive. `docs/data/refresh-logs.json` remains the latest-run
view.

## Included extractors and status meaning

Automated sources currently cover KPI 01 (DataSaudi's public monthly PMI data
endpoint, attributed by DataSaudi to Riyad Bank and linked to the S&P Global
PMI release archive), KPI 02 (SAMA weekly POS report PDFs; transaction count
and transaction value are stored as separate fact rows), KPI 03 (Yamama
Cement's monthly company sales table, summed across companies; Argaam's Saudi
market report is used as a cross-check), KPI 04 (DataSaudi's monthly Grand
Total of banks' claims on the private sector, sourced to SAMA), KPI 09
(DataSaudi's monthly Saudi interbank offered rate, sourced to SAMA;
tenor is not specified by the source), and KPI 14 (Investing.com daily Saudi
5-year USD CDS close; its initial historical baseline was loaded once from a
user-provided CSV covering 2025-01-01 through 2026-11-04. The CSV is not read
on recurring runs; existing workbook history is retained while successful
live observations may be added. The page exposes only about one month and may
block runners), and KPI 15 (DataSaudi's monthly Tadawul All Share Index level;
monthly, not daily).
KPIs 05, 07 and 12 use GASTAT monthly trade database reports. KPIs 08 and 11
use the newest monthly ITR workbook linked on GASTAT's [monthly goods-trade
statistics page](https://www.stats.gov.sa/ar/statistics-tabs?tab=436312&category=123481).
KPI 08 uses table 1.2's national non-oil exports plus re-exports, and KPI 11
uses the same table's oil exports value in SAR million; both series start in
January 2017. ITR observations marked with an asterisk are labeled preliminary.
KPI 13 uses the newest monthly GASTAT STS workbook linked on its
[statistics page](https://www.stats.gov.sa/ar/statistics-tabs?tab=436312&category=123342);
KPI 16 uses DataSaudi's
[Public Finances](https://datasaudi.sa/ar/#public-finances) quarterly
"Capital Expenditures" series, sourced to the Ministry of Finance. The
DataSaudi API reports SAR billions; values are converted to SAR million.
KPI 6 uses the newest SAMA
[Monthly Statistical Bulletin](https://www.sama.gov.sa/ar-sa/statistics/pages/monthlystatistics.aspx)
workbook discovered at each refresh: the monthly Grand Total of Table 20
(letters of credit settled and bills received) and Table 21 (new letters of
credit opened), in SAR million, one `SourceTable` label each. Table 22 has no
monthly data and is not loaded.
Other KPIs appear in the status table as unavailable until their source routes
and access are verified. KPI 10 (MISA) remains excluded by instruction.
Green indicates a configured extraction succeeded; red indicates a failed or
unavailable connector, no data, or intentional exclusion. Check observation
dates; success does not imply that the source has published a newer period.
DataSaudi's PMI feed provides monthly values from August 2009. Confirm
redistribution terms with the source owners before sharing the series with
external systems. DataSaudi's private-sector bank-claims series is monthly,
reported in SAR million, and includes the Grand Total and component rows; the
extractor uses only the Grand Total. Yamama's report currently provides monthly sales history
from January 2010; the September 2026 total of 4,439 thousand tonnes matches
the Argaam Saudi-wide report. SAMA's POS page currently lists the latest 10 weekly PDFs;
each report covers four weeks, so the POS extraction provides a rolling,
overlap-deduplicated recent history rather than a complete historical series.
SAMA reports both count (thousand transactions) and value (SAR thousand).

The GASTAT database-report routes for KPIs 05, 07 and 12 are fixed. Periodic
workbook and PDF sources are discovered from their listing pages during each
refresh. The exporter produces:

- `docs/data/status.json`
- `docs/data/refresh-logs.json` (latest result per KPI; see `docs/logs.html`)
- `docs/refresh-history/index.json` and one JSON log snapshot per workflow run
- `docs/data/facts.json`
- `docs/data/kpi/KPI_01.json` through `KPI_16.json` and `docs/data/kpi/index.json`
- `docs/data/Saudi_KPI_Extract.xlsx`
- `docs/archive/index.json` and one `docs/archive/YYYY-MM-DD.json.gz` full snapshot per day

## Rebuild locally

```powershell
python -m pip install -r .\requirements.txt
python .\publish_static.py
```

## Access notice

This is static public hosting authorized by the repository owner. The source
repository itself remains private. GitHub account/organization policy may
require enabling Pages for a private repository.
