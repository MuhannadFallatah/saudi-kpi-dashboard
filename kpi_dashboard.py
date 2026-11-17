# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import threading
import time
import calendar
import re
import ssl
from concurrent.futures import ThreadPoolExecutor, as_completed
from copy import copy
from html import unescape
from io import BytesIO
from datetime import date, datetime, timedelta, timezone
from urllib.parse import quote, urlencode, urljoin

from pypdf import PdfReader
from pypdf.errors import PdfReadError

RIYADH_TZ = timezone(timedelta(hours=3), "Asia/Riyadh")
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse
from urllib.request import Request, urlopen

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


APP_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = APP_DIR / "Saudi_16_KPI_Data_Mart.xlsx"
WORKBOOK_PATH = Path(os.environ.get("KPI_DATA_WORKBOOK", DEFAULT_WORKBOOK))
REFRESH_MINUTES = max(1, int(os.environ.get("KPI_REFRESH_MINUTES", "1440")))
HOST = "127.0.0.1"
PORT = int(os.environ.get("KPI_PORT", "8765"))
ALLOWED_HOSTS = {
    "www.stats.gov.sa",
    "api.datasaudi.sa",
    "www.sama.gov.sa",
    "www.yamamacement.com",
    "www.investing.com",
    "muhannadfallatah.github.io",
    "database.stats.gov.sa",
}
GASTAT_DB_SELECT_URL = "https://database.stats.gov.sa/Build_Reports/flexmonster/select"
GASTAT_IMPORTS_REPORT_URL = (
    "https://database.stats.gov.sa/Build_Reports/main/home/build-report?reportData=9516309&fetchSavedSlice=true"
)
GASTAT_IMPORTS_INDEX = "DPV_IMP_CP21_EFFT020101"
GASTAT_END_USE_REPORT_URL = (
    "https://database.stats.gov.sa/Build_Reports/main/home/build-report?reportData=9516308&fetchSavedSlice=true"
)
GASTAT_END_USE_INDEX = "DPV_IMP_ENDU_EFFT020101"
GFCF_END_USE = "Gross Fixed Capital Formation"
GASTAT_KPI08_SOURCE_PAGE = (
    "https://www.stats.gov.sa/ar/statistics-tabs?tab=436312&category=123481"
)
GASTAT_KPI13_SOURCE_PAGE = (
    "https://www.stats.gov.sa/ar/statistics-tabs?tab=436312&category=123342"
)
GASTAT_TOTAL_EXPORTS_REPORT_URL = (
    "https://database.stats.gov.sa/Build_Reports/main/home/build-report?reportData=9516298&fetchSavedSlice=true"
)
GASTAT_TOTAL_EXPORTS_INDEX = "DPV_EXP_ENDU_EFFT010101"
ITR_WORKBOOK_URL = (
    "https://www.stats.gov.sa/documents/20117/2435267/ITR+Jul+2026.xlsx/"
    "4cd729e2-a62f-9b94-b06a-23f5f66c54ec?t=1790669050760"
)
# Fallback only; the live refresh uses the newest monthly ITR file linked on GASTAT_KPI08_SOURCE_PAGE.
ITR_FILE_LINK = re.compile(r"/documents/20117/2435267/(ITR[^/\"'\s<>]*)/([^\"'\s<>]+)")
MONTHLY_FILE_PERIOD = re.compile(r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s_+\-]*(20\d{2})", re.IGNORECASE)
STS_FILE_LINK = re.compile(r"/documents/20117/2435267/(STS[^/\"'\s<>]*)/([^\"'\s<>]+)", re.IGNORECASE)
ITR_MONTH_NAMES = ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
KPI08_ITR_NOTE = (
    "Extracted from the newest monthly GASTAT ITR workbook linked on the GASTAT goods-trade statistics page "
    "(table 1.2: national non-oil exports and re-exports as separate source-table facts), SAR million."
)
KPI11_ITR_NOTE = (
    "Extracted from the newest monthly GASTAT ITR workbook linked on the GASTAT goods-trade statistics page "
    "(table 1.2: oil exports value), SAR million; starred observations are preliminary."
)
DATABASE_REPORT_URLS = {
    "KPI_05": GASTAT_IMPORTS_REPORT_URL,
    "KPI_07": GASTAT_END_USE_REPORT_URL,
    "KPI_12": GASTAT_TOTAL_EXPORTS_REPORT_URL,
}
DATABASE_REPORT_NOTES = {
    "KPI_12": "GASTAT database report 'Exports Value by End Use': monthly sum of all end-use categories "
              "(total goods exports), SAR million.",
    "KPI_05": "GASTAT database report 'Imports Value by Central Product Classification (CPC2.1)': "
              "monthly sum of all sections, SAR million.",
    "KPI_07": "GASTAT database report 'Imports Value by End Use': Gross Fixed Capital Formation, SAR million.",
}

DATASAUDI_TASI_URL = (
    "https://api.datasaudi.sa/tesseract/data.jsonrecords"
    "?cube=tadawul_indicators&drilldowns=Indicator%2CMonth&measures=Value&locale=en"
)
DATASAUDI_TASI_PAGE = "https://datasaudi.sa/en/#tadawul-indicators"
SAUDI_EXCHANGE_TASI_URL = (
    "https://www.saudiexchange.sa/wps/portal/saudiexchange/newsandreports/"
    "reports-publications/historical-reports?locale=en"
)
TASI_BASELINE_NOTE = (
    "Daily TASI closes were initially loaded from the Saudi Exchange Historical Reports page "
    "for 2007-01-06 through 2026-11-15 (4,958 observations). Automated requests to the page "
    "return HTTP 403, so this historical baseline does not receive scheduled updates."
)
DATASAUDI_CAPEX_URL = (
    "https://api.datasaudi.sa/tesseract/data.jsonrecords"
    "?cube=mof_government_revenues_expenditures_quarter"
    "&drilldowns=Type%2CQuarter&measures=SAR+Billions&locale=en"
)
DATASAUDI_CAPEX_PAGE = "https://datasaudi.sa/ar/#public-finances"
DATASAUDI_CAPEX_NOTE = (
    "Quarterly capital expenditures from DataSaudi's Public Finances series, sourced to the Ministry of Finance. "
    "Values are converted from SAR billions to SAR million."
)
INVESTING_CDS_URL = "https://www.investing.com/rates-bonds/saudi-arabia-cds-5-year-usd-historical-data"
INVESTING_CDS_PAGE = "https://www.investing.com/rates-bonds/saudi-arabia-cds-5-year-usd"
PUBLISHED_KPI_URL = "https://muhannadfallatah.github.io/saudi-kpi-dashboard/data/kpi/{kpi_id}.json"
DATASAUDI_PMI_URL = (
    "https://api.datasaudi.sa/tesseract/data.jsonrecords"
    "?cube=pmi&drilldowns=Month&measures=Purchasing+Manager+Index"
)
DATASAUDI_PMI_PAGE = "https://datasaudi.sa/en/#purchasing-manager-index"
DATASAUDI_PRIVATE_LENDING_URL = (
    "https://api.datasaudi.sa/tesseract/data.jsonrecords"
    "?cube=lending_private_sector&drilldowns=Lending+Type%2CMonth&measures=Million+SAR&locale=en"
)
DATASAUDI_PRIVATE_LENDING_PAGE = "https://datasaudi.sa/ar/#lending-to-the-private-sector"
DATASAUDI_SAIBOR_URL = (
    "https://api.datasaudi.sa/tesseract/data.jsonrecords"
    "?cube=sama_interest_rates_month&drilldowns=Month"
    "&measures=Saudi+Interbank+Offered+Rates&locale=en"
)
DATASAUDI_SAIBOR_PAGE = "https://datasaudi.sa/ar/#lending-rates"
SAMA_MONTHLY_STATS_PAGE = "https://www.sama.gov.sa/en-US/Statistics/Pages/MonthlyStatistics.aspx"
SAMA_KPI06_SOURCE_PAGE = "https://www.sama.gov.sa/ar-sa/statistics/pages/monthlystatistics.aspx"
KPI06_SHEETS = {
    "20": "Private Sector Imports Financed Through Commercial Banks  (Letters of Credit Settled and Bills Received)",
    "21": "Private Sector Imports Financed Through Commercial Banks  (New Letters of Credit Opened)",
}
KPI06_NOTE = (
    "Monthly Grand Total from SAMA Monthly Statistical Bulletin Tables 20 and 21 (newest bulletin workbook "
    "discovered on the SAMA page at refresh). Table 22 (by exporting countries) has no monthly data and is not loaded."
)
SAMA_POS_PAGE = "https://www.sama.gov.sa/en-US/Statistics/Indices/Pages/POS.aspx"
YAMAMA_CEMENT_URL = "https://www.yamamacement.com/ar/Reports/FullReport"
ARGAAM_CEMENT_REFERENCE_URL = "https://www.argaam.com/ar/financial-reports/cement-monthly-report/2026/66"

FACT_HEADERS = (
    "CountryCode",
    "KPI_ID",
    "PeriodStart",
    "PeriodEnd",
    "Year",
    "Month",
    "MonthName",
    "Value",
    "Unit",
    "SourceTable",
    "ObservationStatus",
    "SourceURL",
    "SourceNote",
    "Open",
    "High",
    "Low",
    "Close",
    "VolumeTraded",
    "TurnoverSAR",
    "TradeCount",
)
BEC_CODES = {"112", "212", "312", "412", "512", "612", "712", "812"}
KPI_NAMES_EN = {
    "KPI_01": "Purchasing Managers' Index (PMI)",
    "KPI_02": "Point-of-sale transactions (POS)",
    "KPI_03": "Cement sales",
    "KPI_04": "Private-sector bank credit (12a)",
    "KPI_05": "Total goods imports",
    "KPI_06": "Private-sector imports",
    "KPI_07": "Capital imports",
    "KPI_08": "Non-oil exports including re-exports",
    "KPI_09": "Monthly average SAIBOR",
    "KPI_10": "Investment licences",
    "KPI_11": "Oil exports",
    "KPI_12": "Goods exports",
    "KPI_13": "Issued building permits",
    "KPI_14": "Credit default swap (CDS) spreads",
    "KPI_15": "Tadawul All Share Index (TASI)",
    "KPI_16": "Government capital expenditure",
}
KPI_SHEET_SLUGS = {
    "KPI_01": "PMI",
    "KPI_02": "POS_Transactions",
    "KPI_03": "Cement_Sales",
    "KPI_04": "Private_Bank_Credit",
    "KPI_05": "Total_Goods_Imports",
    "KPI_06": "Private_Sector_Imports",
    "KPI_07": "Capital_Imports",
    "KPI_08": "NonOil_Exports",
    "KPI_09": "SAIBOR",
    "KPI_10": "Investment_Licences",
    "KPI_11": "Oil_Exports",
    "KPI_12": "Goods_Exports",
    "KPI_13": "Building_Permits",
    "KPI_14": "CDS_Spreads",
    "KPI_15": "TASI",
    "KPI_16": "Govt_Capex",
}
MONTHS = {
    name.lower(): number
    for number, name in enumerate(
        ("", "January", "February", "March", "April", "May", "June", "July",
         "August", "September", "October", "November", "December")
    )
    if name
}


def serialize(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def workbook_data(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise FileNotFoundError(f"Data workbook does not exist: {path}")

    workbook = load_workbook(path, read_only=True, data_only=True)
    required = {"KPI_Catalog", "Source_Register", "Fact_All"}
    missing = required.difference(workbook.sheetnames)
    if missing:
        workbook.close()
        raise ValueError(f"Workbook is missing required sheets: {', '.join(sorted(missing))}")

    catalog_sheet = workbook["KPI_Catalog"]
    catalog_headers = [cell.value for cell in catalog_sheet[1]]
    catalog = [
        dict(zip(catalog_headers, row))
        for row in catalog_sheet.iter_rows(min_row=2, values_only=True)
        if row and row[0]
    ]
    source_sheet = workbook["Source_Register"]
    source_headers = [cell.value for cell in source_sheet[1]]
    sources = {
        row[0]: dict(zip(source_headers, row))
        for row in source_sheet.iter_rows(min_row=2, values_only=True)
        if row and row[0]
    }

    facts = []
    for row in workbook["Fact_All"].iter_rows(min_row=2, values_only=True):
        if not row or not row[1]:
            continue
        facts.append(dict(zip(FACT_HEADERS, (serialize(value) for value in row[: len(FACT_HEADERS)]))))
    components = []
    if "Fact_KPI07_BEC" in workbook.sheetnames:
        component_sheet = workbook["Fact_KPI07_BEC"]
        component_headers = [cell.value for cell in component_sheet[1]]
        components = [
            dict(zip(component_headers, (serialize(value) for value in row)))
            for row in component_sheet.iter_rows(min_row=2, values_only=True)
            if row and row[1] == "KPI_07"
        ]
    export_components = []
    if "08_KPI_Components" in workbook.sheetnames:
        component_sheet = workbook["08_KPI_Components"]
        component_headers = [cell.value for cell in component_sheet[1]]
        export_components = [
            dict(zip(component_headers, (serialize(value) for value in row)))
            for row in component_sheet.iter_rows(min_row=2, values_only=True)
            if row and row[1] == "KPI_08"
        ]
    workbook.close()
    apply_configured_source_metadata(catalog, sources)
    return {
        "catalog": catalog,
        "sources": sources,
        "facts": facts,
        "components": components,
        "kpi08_components": export_components,
    }


def tls_context() -> ssl.SSLContext:
    context = ssl.create_default_context()
    # www.mof.gov.sa omits its DigiCert intermediate certificate from the TLS handshake.
    for pem in sorted((APP_DIR / "certs").glob("*.pem")):
        context.load_verify_locations(cafile=str(pem))
    return context


def download_bytes(url: str) -> bytes:
    parsed = urlparse(url)
    if parsed.scheme != "https" or parsed.hostname not in ALLOWED_HOSTS:
        raise ValueError("Live extractors only allow HTTPS URLs on approved source hosts.")
    request = Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
            "Accept": "text/html,application/pdf,application/json,*/*",
        },
    )
    with urlopen(request, timeout=45, context=tls_context()) as response:
        if response.status != 200:
            raise OSError(f"Source returned HTTP {response.status}: {url}")
        return response.read()


def apply_configured_source_metadata(catalog: list[dict[str, Any]], sources: dict[str, dict[str, Any]]) -> None:
    # KPIs 05, 07 and 12 use GASTAT reports; KPIs 08 and 11 use ITR; KPI 16 uses DataSaudi.
    for item in catalog:
        kpi_id = item.get("KPI_ID")
        if kpi_id in {"KPI_08", "KPI_11"}:
            url = ITR_WORKBOOK_URL
            source_page_url = GASTAT_KPI08_SOURCE_PAGE
            route = f"GASTAT monthly ITR workbook: {url}"
            note = KPI08_ITR_NOTE if kpi_id == "KPI_08" else KPI11_ITR_NOTE
        elif kpi_id == "KPI_06":
            url = (sources.get(kpi_id, {}).get("ExtractedDataURL") or SAMA_KPI06_SOURCE_PAGE)
            source_page_url = SAMA_KPI06_SOURCE_PAGE
            route = f"Newest SAMA monthly statistical bulletin workbook discovered from: {source_page_url}"
            note = KPI06_NOTE
        elif kpi_id == "KPI_16":
            url = DATASAUDI_CAPEX_URL
            source_page_url = DATASAUDI_CAPEX_PAGE
            route =  f"DataSaudi Public Finances API: {url}"
            note = DATASAUDI_CAPEX_NOTE
        elif kpi_id == "KPI_13":
            source = sources.get(kpi_id, {})
            url = source.get("ExtractedDataURL") or item.get("VerifiedData_or_API_ReferenceURL")
            source_page_url = GASTAT_KPI13_SOURCE_PAGE
            route = f"Newest monthly GASTAT STS workbook discovered from: {source_page_url}"
            note = (
                "The newest monthly workbook linked on the GASTAT STS statistics page is selected at refresh; "
                "Table 7 contains issued building permits, and starred 2026 values are preliminary."
            )
        elif kpi_id == "KPI_15":
            url = SAUDI_EXCHANGE_TASI_URL
            source_page_url = SAUDI_EXCHANGE_TASI_URL
            route = "One-time daily historical baseline from the Saudi Exchange Historical Reports page"
            note = TASI_BASELINE_NOTE
        elif kpi_id in DATABASE_REPORT_URLS:
            url = DATABASE_REPORT_URLS[kpi_id]
            source_page_url = url
            route = f"GASTAT database report (Build Reports API): {url}"
            note = DATABASE_REPORT_NOTES[kpi_id]
        else:
            continue
        item.update({
            "SourceURL": source_page_url,
            "API_or_DataRoute": route,
            "VerifiedData_or_API_ReferenceURL": source_page_url,
            "Access_Caveat": None,
            "ExtractionNote": note,
        })
        if kpi_id == "KPI_16":
            item.update({
                "Source": "DataSaudi (Ministry of Finance)",
                "Frequency_or_TimeFrame": "Quarterly",
                "APIMatchStatus": "Verified API",
                "ExtractionStatus": "Live API extraction configured",
            })
        if kpi_id == "KPI_15":
            item.update({
                "Source": "Saudi Exchange",
                "Frequency_or_TimeFrame": "Daily",
                "APIMatchStatus": "Historical report; no automated API verified",
                "Access_Caveat": (
                    "Automated requests to the historical reports page return HTTP 403; "
                    "scheduled updates are not configured."
                ),
                "ExtractionStatus": "Historical baseline loaded; live refresh unavailable",
            })
        source = sources.setdefault(kpi_id, {"KPI_ID": kpi_id})
        source.update({
            "PrimarySourceURL": source_page_url,
            "ExtractedDataURL": url,
            "ExtractionStatus": (
                "Live API extraction configured" if kpi_id == "KPI_16"
                else "Historical baseline loaded; live refresh unavailable" if kpi_id == "KPI_15"
                else "Extracted"
            ),
            "ExtractionNote": note,
        })


def post_json(url: str, body: dict[str, Any]) -> dict[str, Any]:
    parsed = urlparse(url)
    if parsed.scheme != "https" or parsed.hostname not in ALLOWED_HOSTS:
        raise ValueError("Live extractors only allow HTTPS URLs on approved GASTAT and MOF hosts.")
    request = Request(
        url,
        data=json.dumps(body).encode("utf-8"),
        headers={"User-Agent": "Saudi-KPI-Monitor/1.0", "Content-Type": "application/json"},
        method="POST",
    )
    with urlopen(request, timeout=60, context=tls_context()) as response:
        if response.status != 200:
            raise OSError(f"Source returned HTTP {response.status}: {url}")
        return json.loads(response.read().decode("utf-8"))


def fetch_gastat_imports() -> dict[str, Any]:
    # Same aggregate the GASTAT "Imports Value by CPC2.1" report requests, without the product rows.
    return post_json(GASTAT_DB_SELECT_URL, {
        "type": "select",
        "index": GASTAT_IMPORTS_INDEX,
        "query": {"aggs": {
            "by": {"cols": [{"uniqueName": "YEAR_MONTH"}]},
            "values": [{"func": "sum", "field": {"uniqueName": "OBS_VALUE"}}],
        }},
        "page": 0,
    })


def fetch_gastat_total_exports() -> dict[str, Any]:
    return post_json(GASTAT_DB_SELECT_URL, {
        "type": "select",
        "index": GASTAT_TOTAL_EXPORTS_INDEX,
        "query": {"aggs": {
            "by": {"rows": [{"uniqueName": "ENDU_EN"}], "cols": [{"uniqueName": "YEAR_MONTH"}]},
            "values": [{"func": "sum", "field": {"uniqueName": "OBS_VALUE"}}],
        }},
        "page": 0,
    })


def extract_total_exports(payload: dict[str, Any], source_url: str) -> list[dict[str, Any]]:
    if payload.get("pageTotal", 1) != 1:
        raise ValueError("GASTAT exports response is paged; expected a single page.")
    totals: dict[tuple[int, int], float] = {}
    for item in payload.get("aggs", []):
        keys = item.get("keys") or {}
        if not keys.get("ENDU_EN") or not keys.get("YEAR_MONTH"):
            continue  # subtotal and grand-total rows
        match = re.fullmatch(r"\s*(\d{4})\s*/\s*(\d{1,2})\s*", str(keys["YEAR_MONTH"]))
        value = ((item.get("values") or {}).get("OBS_VALUE") or {}).get("sum")
        if not match or value is None:
            raise ValueError(f"Unexpected GASTAT exports observation: {item}")
        key = (int(match.group(1)), int(match.group(2)))
        totals[key] = totals.get(key, 0.0) + value
    if not totals:
        raise ValueError("GASTAT exports response contained no monthly observations.")
    return [
        make_fact(
            "KPI_12", year, month, total / 1_000_000, "SAR million",
            "GASTAT Exports Value by End Use", source_url,
            "Monthly total goods exports (sum of end-use categories), converted from SAR to SAR million.",
        )
        for (year, month), total in sorted(totals.items())
    ]


def fetch_gastat_end_use_imports() -> dict[str, Any]:
    return post_json(GASTAT_DB_SELECT_URL, {
        "type": "select",
        "index": GASTAT_END_USE_INDEX,
        "query": {"aggs": {
            "by": {"rows": [{"uniqueName": "ENDU_EN"}], "cols": [{"uniqueName": "YEAR_MONTH"}]},
            "values": [{"func": "sum", "field": {"uniqueName": "OBS_VALUE"}}],
        }},
        "page": 0,
    })


def latest_itr_workbook_url(page: bytes) -> str:
    # The non-oil ITR files are a separate series and are skipped.
    html = unescape(page.decode("utf-8", "replace"))
    latest: tuple[tuple[int, int], str] | None = None
    for file_name, token in dict.fromkeys(ITR_FILE_LINK.findall(html)):
        name = unquote(file_name)
        if ".xlsx" not in name.lower() or "non" in name.lower():
            continue
        match = MONTHLY_FILE_PERIOD.search(name)
        if not match:
            continue
        period = (int(match.group(2)), ITR_MONTH_NAMES.index(match.group(1).lower()) + 1)
        if latest is None or period > latest[0]:
            latest = (period, f"https://www.stats.gov.sa/documents/20117/2435267/{file_name}/{token}")
    if latest is None:
        raise ValueError("No monthly ITR workbook was found on the GASTAT statistics page.")
    return latest[1]


def latest_sts_workbook_url(page: bytes) -> str:
    html = unescape(page.decode("utf-8", "replace"))
    latest: tuple[tuple[int, int], str] | None = None
    for file_name, token in dict.fromkeys(STS_FILE_LINK.findall(html)):
        name = unquote(file_name)
        if ".xlsx" not in name.lower():
            continue
        match = MONTHLY_FILE_PERIOD.search(name)
        if not match:
            continue
        period = (int(match.group(2)), ITR_MONTH_NAMES.index(match.group(1).lower()) + 1)
        if latest is None or period > latest[0]:
            latest = (period, f"https://www.stats.gov.sa/documents/20117/2435267/{file_name}/{token}")
    if latest is None:
        raise ValueError("No monthly STS workbook was found on the GASTAT statistics page.")
    return latest[1]


def latest_sama_bulletin_url(page: bytes) -> str:
    html = unescape(page.decode("utf-8", "replace")).replace(r"\u002f", "/")
    latest: tuple[tuple[int, int], str] | None = None
    for path in dict.fromkeys(re.findall(r'"FileRef"\s*:\s*"([^"]+\.xlsx)"', html, flags=re.IGNORECASE)):
        if "monthly_statistical_bulletin" not in path.lower():
            continue
        match = MONTHLY_FILE_PERIOD.search(unquote(path).rsplit("/", 1)[-1])
        if not match:
            continue
        period = (int(match.group(2)), ITR_MONTH_NAMES.index(match.group(1).lower()) + 1)
        if latest is None or period > latest[0]:
            latest = (period, urljoin(SAMA_KPI06_SOURCE_PAGE, path))
    if latest is None:
        raise ValueError("No monthly statistical bulletin workbook was found on the SAMA page.")
    return latest[1]


def extract_sama_private_imports(content: bytes, source_url: str) -> list[dict[str, Any]]:
    try:
        workbook = load_workbook(BytesIO(content), data_only=True, read_only=True)
    except Exception as error:
        raise ValueError(f"Unable to read SAMA monthly bulletin workbook: {error}") from error

    facts: list[dict[str, Any]] = []
    for sheet_name, source_table in KPI06_SHEETS.items():
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"SAMA bulletin workbook has no sheet {sheet_name}.")
        rows = list(workbook[sheet_name].iter_rows(values_only=True))
        total_column = None
        for row in rows[:14]:
            for index, cell in enumerate(row):
                if isinstance(cell, str) and cell.strip().lower() == "grand":
                    total_column = index
        if total_column is None:
            raise ValueError(f"SAMA bulletin sheet {sheet_name} has no Grand Total column.")
        # Annual, quarterly and monthly blocks are stacked; monthly dates follow the quarterly rows.
        monthly = False
        for row in rows:
            label = row[0] if row else None
            if isinstance(label, str) and re.fullmatch(r"Q[1-4]", label.strip()):
                monthly = True
                continue
            if not monthly or not isinstance(label, datetime):
                continue
            value = row[total_column] if total_column < len(row) else None
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                facts.append(make_fact(
                    "KPI_06", label.year, label.month, value, "SAR million", source_table,
                    source_url, KPI06_NOTE,
                ))
    if not facts:
        raise ValueError("SAMA bulletin sheets 20 and 21 contain no monthly observations.")
    return facts


def extract_gfcf_imports(payload: dict[str, Any], source_url: str) -> list[dict[str, Any]]:
    if payload.get("pageTotal", 1) != 1:
        raise ValueError("GASTAT end-use response is paged; expected a single page.")
    facts = []
    for item in payload.get("aggs", []):
        keys = item.get("keys") or {}
        if keys.get("ENDU_EN") != GFCF_END_USE or not keys.get("YEAR_MONTH"):
            continue
        match = re.fullmatch(r"\s*(\d{4})\s*/\s*(\d{1,2})\s*", str(keys["YEAR_MONTH"]))
        total = ((item.get("values") or {}).get("OBS_VALUE") or {}).get("sum")
        if not match or total is None:
            raise ValueError(f"Unexpected GASTAT end-use observation: {item}")
        facts.append(make_fact(
            "KPI_07", int(match.group(1)), int(match.group(2)), total / 1_000_000, "SAR million",
            "GASTAT Imports Value by End Use - Gross Fixed Capital Formation", source_url,
            "Monthly imports classified as Gross Fixed Capital Formation, converted from SAR to SAR million.",
        ))
    if not facts:
        raise ValueError(f"GASTAT end-use response has no '{GFCF_END_USE}' observations.")
    return sorted(facts, key=lambda fact: fact["PeriodStart"])


def extract_imports_cpc(payload: dict[str, Any], source_url: str) -> list[dict[str, Any]]:
    if payload.get("pageTotal", 1) != 1:
        raise ValueError("GASTAT imports response is paged; expected a single page of monthly totals.")
    facts = []
    for item in payload.get("aggs", []):
        period = (item.get("keys") or {}).get("YEAR_MONTH")
        if not period:
            continue  # grand-total row
        match = re.fullmatch(r"\s*(\d{4})\s*/\s*(\d{1,2})\s*", str(period))
        total = ((item.get("values") or {}).get("OBS_VALUE") or {}).get("sum")
        if not match or total is None:
            raise ValueError(f"Unexpected GASTAT imports observation: {item}")
        year, month = int(match.group(1)), int(match.group(2))
        facts.append(make_fact(
            "KPI_05", year, month, total / 1_000_000, "SAR million",
            "GASTAT Imports Value by Central Product Classification (CPC2.1)", source_url,
            "Monthly sum of OBS_VALUE across all CPC2.1 sections, converted from SAR to SAR million.",
        ))
    if not facts:
        raise ValueError("GASTAT imports response contained no monthly observations.")
    return sorted(facts, key=lambda fact: fact["PeriodStart"])


def make_quarter_fact(
    kpi_id: str, year: int, quarter: int, value: float, unit: str,
    source_table: str, source_url: str, source_note: str,
) -> dict[str, Any]:
    first_month = 3 * quarter - 2
    fact = make_fact(kpi_id, year, first_month, value, unit, source_table, source_url, source_note)
    last_month = first_month + 2
    fact["PeriodEnd"] = date(year, last_month, calendar.monthrange(year, last_month)[1]).isoformat()
    fact["Month"] = None
    fact["MonthName"] = f"Q{quarter}"
    return fact


def make_fact(
    kpi_id: str,
    year: int,
    month: int,
    value: float,
    unit: str,
    source_table: str,
    source_url: str,
    source_note: str,
    status: str = "Reported",
) -> dict[str, Any]:
    start = date(year, month, 1)
    end = date(year, month, calendar.monthrange(year, month)[1])
    return {
        "CountryCode": "SAU",
        "KPI_ID": kpi_id,
        "PeriodStart": start.isoformat(),
        "PeriodEnd": end.isoformat(),
        "Year": year,
        "Month": month,
        "MonthName": start.strftime("%B"),
        "Value": float(value),
        "Unit": unit,
        "SourceTable": source_table,
        "ObservationStatus": status,
        "SourceURL": source_url,
        "SourceNote": source_note,
        "Open": None,
        "High": None,
        "Low": None,
        "Close": None,
        "VolumeTraded": None,
        "TurnoverSAR": None,
        "TradeCount": None,
    }


def extract_datasaudi_pmi(content: bytes) -> list[dict[str, Any]]:
    try:
        payload = json.loads(content)
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValueError("DataSaudi PMI response is not valid JSON.") from error

    if not isinstance(payload, dict):
        raise ValueError("DataSaudi PMI response must be a JSON object.")

    expected_measure = "Purchasing Manager Index"
    columns = payload.get("columns")
    data = payload.get("data")
    page = payload.get("page")
    total = page.get("total") if isinstance(page, dict) else None
    if not isinstance(columns, list) or expected_measure not in columns:
        raise ValueError("DataSaudi PMI response is missing the expected measure column.")
    if not isinstance(data, list) or not data:
        raise ValueError("DataSaudi PMI response contains no monthly observations.")
    if not isinstance(total, int) or total != len(data):
        raise ValueError("DataSaudi PMI response is incomplete or has an invalid total.")

    note = (
        "Retrieved from DataSaudi's public PMI monthly data endpoint. DataSaudi attributes "
        "the series to Riyad Bank and links the S&P Global PMI release archive."
    )
    rows = []
    seen_periods: set[str] = set()
    for item in data:
        if not isinstance(item, dict):
            raise ValueError("DataSaudi PMI response contains an invalid observation.")
        period = item.get("Month")
        value = item.get(expected_measure)
        if not isinstance(period, str) or not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", period):
            raise ValueError(f"DataSaudi PMI response has an invalid month: {period!r}.")
        if period in seen_periods:
            raise ValueError(f"DataSaudi PMI response repeats month {period}.")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not 0 <= value <= 100:
            raise ValueError(f"DataSaudi PMI response has an invalid value for {period}.")
        seen_periods.add(period)
        year, month = (int(part) for part in period.split("-"))
        rows.append(
            make_fact(
                "KPI_01",
                year,
                month,
                value,
                "index points",
                "Purchasing Managers' Index - Monthly (DataSaudi)",
                DATASAUDI_PMI_PAGE,
                note,
            )
        )
    return sorted(rows, key=lambda row: row["PeriodStart"])


def extract_datasaudi_private_lending(content: bytes) -> list[dict[str, Any]]:
    try:
        payload = json.loads(content)
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValueError("DataSaudi private lending response is not valid JSON.") from error
    if not isinstance(payload, dict):
        raise ValueError("DataSaudi private lending response must be a JSON object.")

    required_columns = {"Month", "Lending Type", "Million SAR"}
    columns = payload.get("columns")
    data = payload.get("data")
    page = payload.get("page")
    total = page.get("total") if isinstance(page, dict) else None
    if not isinstance(columns, list) or not required_columns.issubset(columns):
        raise ValueError("DataSaudi private lending response is missing expected data columns.")
    if not isinstance(data, list) or not data:
        raise ValueError("DataSaudi private lending response contains no monthly observations.")
    if not isinstance(total, int) or total != len(data):
        raise ValueError("DataSaudi private lending response is incomplete or has an invalid total.")

    rows_by_period: dict[str, dict[str, Any]] = {}
    for item in data:
        if not isinstance(item, dict):
            raise ValueError("DataSaudi private lending response contains an invalid observation.")
        if item.get("Lending Type") != "Grand Total":
            continue
        period = item.get("Month")
        value = item.get("Million SAR")
        if not isinstance(period, str) or not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", period):
            raise ValueError(f"DataSaudi private lending response has an invalid month: {period!r}.")
        if period in rows_by_period:
            raise ValueError(f"DataSaudi private lending response repeats total for {period}.")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or value <= 0:
            raise ValueError(f"DataSaudi private lending response has an invalid total for {period}.")
        year, month = (int(part) for part in period.split("-"))
        rows_by_period[period] = make_fact(
            "KPI_04",
            year,
            month,
            value,
            "SAR million",
            "Banks' Claims on Private Sector - Grand Total (DataSaudi/SAMA)",
            DATASAUDI_PRIVATE_LENDING_PAGE,
            "Monthly grand total of banks' claims on the private sector, from DataSaudi's series sourced "
            "to Saudi Central Bank (SAMA). Grand Total is used directly; lending components are not added again.",
        )
    if not rows_by_period:
        raise ValueError("DataSaudi private lending response contains no Grand Total observations.")
    return [rows_by_period[period] for period in sorted(rows_by_period)]


def extract_datasaudi_saibor(content: bytes) -> list[dict[str, Any]]:
    try:
        payload = json.loads(content)
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValueError("DataSaudi SAIBOR response is not valid JSON.") from error
    if not isinstance(payload, dict):
        raise ValueError("DataSaudi SAIBOR response must be a JSON object.")
    data = payload.get("data")
    page = payload.get("page")
    total = page.get("total") if isinstance(page, dict) else None
    if not isinstance(data, list) or not data:
        raise ValueError("DataSaudi SAIBOR response contains no monthly observations.")
    if not isinstance(total, int) or total != len(data):
        raise ValueError("DataSaudi SAIBOR response is incomplete or has an invalid total.")

    rows_by_period: dict[str, dict[str, Any]] = {}
    for item in data:
        if not isinstance(item, dict):
            raise ValueError("DataSaudi SAIBOR response contains an invalid observation.")
        period = item.get("Month")
        value = item.get("Saudi Interbank Offered Rates")
        if not isinstance(period, str) or not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", period):
            raise ValueError(f"DataSaudi SAIBOR response has an invalid month: {period!r}.")
        if period in rows_by_period:
            raise ValueError(f"DataSaudi SAIBOR response repeats {period}.")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or value < 0:
            raise ValueError(f"DataSaudi SAIBOR response has an invalid rate for {period}.")
        year, month = (int(part) for part in period.split("-"))
        rows_by_period[period] = make_fact(
            "KPI_09",
            year,
            month,
            value,
            "percent",
            "Saudi Interbank Offered Rates - monthly (DataSaudi/SAMA)",
            DATASAUDI_SAIBOR_PAGE,
            "Monthly Saudi interbank offered rate (SAIBOR) from DataSaudi's 'Interest Rate - Monthly' "
            f"table sourced to SAMA ({SAMA_MONTHLY_STATS_PAGE}). Tenor is not specified by the source.",
        )
    return [rows_by_period[period] for period in sorted(rows_by_period)]


def extract_investing_cds(content: bytes) -> list[dict[str, Any]]:
    html = content.decode("utf-8", "replace")
    match = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', html, re.S)
    if not match:
        raise ValueError("Investing.com page has no embedded historical data (possibly blocked).")
    try:
        state = json.loads(match.group(1))["props"]["pageProps"]["state"]
        records = state["historicalDataStore"]["historicalData"]["data"]
    except (json.JSONDecodeError, KeyError, TypeError) as error:
        raise ValueError("Investing.com historical data structure was not recognised.") from error
    if not isinstance(records, list) or not records:
        raise ValueError("Investing.com page contains no CDS observations.")

    rows: dict[str, dict[str, Any]] = {}
    for record in records:
        try:
            day = datetime.fromisoformat(record["rowDateTimestamp"].replace("Z", "+00:00")).date()
            close = float(record["last_closeRaw"])
            values = {
                key: float(record[f"last_{name}Raw"])
                for key, name in (("Open", "open"), ("High", "max"), ("Low", "min"))
            }
        except (KeyError, TypeError, ValueError) as error:
            raise ValueError("Investing.com CDS record is malformed.") from error
        if close <= 0:
            raise ValueError(f"Investing.com CDS close is invalid for {day}.")
        fact = make_fact(
            "KPI_14",
            day.year,
            day.month,
            close,
            "basis points",
            "Saudi Arabia CDS 5 Year USD - daily close (Investing.com)",
            INVESTING_CDS_PAGE,
            "Daily close of the Saudi Arabia 5-year USD CDS spread from Investing.com's historical data "
            "page. The page only exposes about one month, so older days come from earlier refreshes.",
        )
        fact["PeriodStart"] = fact["PeriodEnd"] = day.isoformat()
        fact["MonthName"] = day.strftime("%b %d")
        fact.update(values)
        fact["Close"] = close
        rows[fact["PeriodStart"]] = fact
    return [rows[key] for key in sorted(rows)]


def merge_cds_history(*groups: list[dict[str, Any]]) -> list[dict[str, Any]]:
    history: dict[str, dict[str, Any]] = {}
    for group in groups:
        for fact in group:
            if fact.get("KPI_ID") == "KPI_14" and fact.get("PeriodStart"):
                period = str(fact["PeriodStart"]).split("T", 1)[0]
                normalized = dict(fact)
                normalized["PeriodStart"] = period
                normalized["PeriodEnd"] = period
                history[period] = normalized
    return [history[key] for key in sorted(history)]


def extract_datasaudi_tasi(content: bytes) -> list[dict[str, Any]]:
    try:
        payload = json.loads(content)
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValueError("DataSaudi Tadawul response is not valid JSON.") from error
    if not isinstance(payload, dict):
        raise ValueError("DataSaudi Tadawul response must be a JSON object.")
    data = payload.get("data")
    page = payload.get("page")
    total = page.get("total") if isinstance(page, dict) else None
    if not isinstance(data, list) or not data:
        raise ValueError("DataSaudi Tadawul response contains no observations.")
    if not isinstance(total, int) or total != len(data):
        raise ValueError("DataSaudi Tadawul response is incomplete or has an invalid total.")

    rows: dict[str, dict[str, Any]] = {}
    for item in data:
        if not isinstance(item, dict):
            raise ValueError("DataSaudi Tadawul response contains an invalid observation.")
        if item.get("Indicator") != "Index":
            continue
        period = item.get("Month")
        value = item.get("Value")
        if not isinstance(period, str) or not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", period):
            raise ValueError(f"DataSaudi Tadawul response has an invalid month: {period!r}.")
        if period in rows:
            raise ValueError(f"DataSaudi Tadawul response repeats the index for {period}.")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or value <= 0:
            raise ValueError(f"DataSaudi Tadawul response has an invalid index for {period}.")
        year, month = (int(part) for part in period.split("-"))
        rows[period] = make_fact(
            "KPI_15",
            year,
            month,
            value,
            "index points",
            "Tadawul All Share Index (TASI) - monthly (DataSaudi/Tadawul)",
            DATASAUDI_TASI_PAGE,
            "Monthly Tadawul All Share Index level from DataSaudi's 'Tadawul Indicators - Monthly' table "
            "('Index' indicator); the series is monthly, not daily.",
        )
    if not rows:
        raise ValueError("DataSaudi Tadawul response contains no 'Index' observations.")
    return [rows[period] for period in sorted(rows)]


def extract_datasaudi_capex(content: bytes) -> list[dict[str, Any]]:
    try:
        payload = json.loads(content)
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValueError("DataSaudi public-finances response is not valid JSON.") from error
    if not isinstance(payload, dict):
        raise ValueError("DataSaudi public-finances response must be a JSON object.")

    required_columns = {"Quarter", "Type ID", "Type", "SAR Billions"}
    columns = payload.get("columns")
    data = payload.get("data")
    page = payload.get("page")
    total = page.get("total") if isinstance(page, dict) else None
    if not isinstance(columns, list) or not required_columns.issubset(columns):
        raise ValueError("DataSaudi public-finances response is missing expected data columns.")
    if not isinstance(data, list) or not data:
        raise ValueError("DataSaudi public-finances response contains no quarterly observations.")
    if not isinstance(total, int) or total != len(data):
        raise ValueError("DataSaudi public-finances response is incomplete or has an invalid total.")

    rows: dict[str, dict[str, Any]] = {}
    for item in data:
        if not isinstance(item, dict):
            raise ValueError("DataSaudi public-finances response contains an invalid observation.")
        if item.get("Type ID") != 4 or item.get("Type") != "Capital Expenditures":
            continue
        period = item.get("Quarter")
        match = re.fullmatch(r"(\d{4})-Q([1-4])", period) if isinstance(period, str) else None
        value = item.get("SAR Billions")
        if not match:
            raise ValueError(f"DataSaudi public-finances response has an invalid quarter: {period!r}.")
        if period in rows:
            raise ValueError(f"DataSaudi public-finances response repeats capital expenditures for {period}.")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or value <= 0:
            raise ValueError(f"DataSaudi public-finances response has an invalid value for {period}.")
        year, quarter = int(match.group(1)), int(match.group(2))
        rows[period] = make_quarter_fact(
            "KPI_16",
            year,
            quarter,
            value * 1000,
            "SAR million",
            "Government Revenues and Expenditures - Quarterly (DataSaudi/MOF)",
            DATASAUDI_CAPEX_PAGE,
            DATASAUDI_CAPEX_NOTE,
        )
    if not rows:
        raise ValueError("DataSaudi public-finances response contains no capital-expenditure observations.")
    return [rows[period] for period in sorted(rows)]


def sama_pos_report_urls(fetch, full_history: bool = False) -> list[str]:
    page_url = SAMA_POS_PAGE
    visited: set[str] = set()
    urls: list[str] = []
    while page_url:
        if page_url in visited:
            raise ValueError("SAMA POS report list returned a repeated pagination link.")
        visited.add(page_url)
        html = fetch(page_url).decode("utf-8", "replace")
        html = html.replace(r"\u002f", "/")
        paths = re.findall(r'"FileRef"\s*:\s*"([^"]+\.pdf)"', html, flags=re.IGNORECASE)
        paths = [path for path in paths if "/statistics/indices/pos_en/" in path.lower()]
        page_urls = list(dict.fromkeys(
            urljoin(SAMA_POS_PAGE, quote(path, safe="/%")) for path in paths
        ))
        if not page_urls:
            if not urls:
                title = re.search(r"<title[^>]*>(.*?)</title>", html, flags=re.IGNORECASE | re.DOTALL)
                page_title = re.sub(r"\s+", " ", title.group(1)).strip() if title else "unavailable"
                raise ValueError(
                    f"SAMA POS page did not contain weekly report PDF links "
                    f"(title={page_title!r}, response length={len(html)})."
                )
            raise ValueError("A SAMA POS report-list page contained no weekly report PDF links.")
        urls.extend(url for url in page_urls if url not in urls)
        if not full_history:
            break
        next_page = re.search(r'"NextHref"\s*:\s*"([^"]+)"', html, flags=re.IGNORECASE)
        if not next_page:
            break
        next_href = json.loads(f'"{next_page.group(1)}"')
        page_url = urljoin(SAMA_POS_PAGE, next_href)
    return urls if full_history else urls[:10]


def extract_sama_pos_pdf(content: bytes, source_url: str) -> list[dict[str, Any]]:
    try:
        reader = PdfReader(BytesIO(content))
        if not reader.pages:
            raise ValueError("SAMA POS report PDF has no pages.")
        text = reader.pages[0].extract_text() or ""
    except (PdfReadError, OSError, EOFError) as error:
        raise ValueError(f"Unable to read SAMA POS report PDF: {error}") from error

    return extract_sama_pos_text(text, source_url)


def extract_sama_pos_text(text: str, source_url: str) -> list[dict[str, Any]]:
    text = text.replace("\xa0", " ")
    text = re.sub(r"[\u2010-\u2015\u2212]", "-", text)
    date_token = r"\d{1,2}(?:\s+[A-Za-z]{3})?(?:,\s*\d{2}|\s+\d{4})?"
    period_pattern = (
        rf"({date_token})\s*-\s*({date_token})"
    )
    report_date_match = re.search(
        r"(\d{1,2})(?:st|nd|rd|th)?[\s_-]*([A-Za-z]{3,9})[,\s_-]*(20\d{2})",
        unquote(source_url),
        flags=re.IGNORECASE,
    )
    years_to_date = re.findall(r"\bYear to Date\s+(\d{4})\b", text, flags=re.IGNORECASE)
    reference_year = int(years_to_date[-1]) if years_to_date else None
    report_date = None
    if report_date_match:
        month_format = "%b" if len(report_date_match.group(2)) == 3 else "%B"
        report_date = datetime.strptime(
            f"{report_date_match.group(1)} {report_date_match.group(2)} "
            f"{report_date_match.group(3)}",
            f"%d {month_format} %Y",
        ).date()
    else:
        report_note = re.search(
            r"\bNote:\s*Until\s+([A-Za-z]{3,9})\s+(\d{1,2})(?:st|nd|rd|th)?"
            r"(?:,\s*(20\d{2}))?",
            text,
            flags=re.IGNORECASE,
        )
        note_year = int(report_note.group(3)) if report_note and report_note.group(3) else reference_year
        if report_note and note_year:
            month_format = "%b" if len(report_note.group(1)) == 3 else "%B"
            report_date = datetime.strptime(
                f"{report_note.group(2)} {report_note.group(1)} {note_year}",
                f"%d {month_format} %Y",
            ).date()
    reference_year = report_date.year if report_date else int(years_to_date[-1]) if years_to_date else None

    def period_range(period: tuple[str, str]) -> tuple[date, date]:
        def components(value: str) -> tuple[int, int | None, int | None]:
            match = re.fullmatch(
                r"(\d{1,2})(?:\s+([A-Za-z]{3}))?(?:(?:,\s*(\d{2}))|(?:\s+(\d{4})))?",
                value,
            )
            if not match:
                raise ValueError("SAMA POS report contains an invalid weekly date.")
            day = int(match.group(1))
            month = datetime.strptime(match.group(2), "%b").month if match.group(2) else None
            year_text = match.group(3) or match.group(4)
            year = (
                datetime.strptime(year_text, "%y").year
                if year_text and len(year_text) == 2
                else int(year_text) if year_text else None
            )
            return day, month, year

        start_day, start_month, start_year = components(period[0])
        end_day, end_month, end_year = components(period[1])
        if start_month is None and end_month is None:
            if not report_date:
                raise ValueError("SAMA POS report dates omit months and provide no report date.")
            if end_day < start_day:
                end_month = report_date.month
                start_month = (end_month - 2) % 12 + 1
            else:
                start_month = end_month = report_date.month
        elif start_month is None:
            start_month = end_month - int(start_day > end_day)
            if start_month == 0:
                start_month = 12
        elif end_month is None:
            end_month = start_month + int(end_day < start_day)
            if end_month == 13:
                end_month = 1
        if end_year is None:
            if report_date:
                end_year = report_date.year - int(
                    (end_month, end_day) > (report_date.month, report_date.day)
                )
            elif reference_year:
                end_year = reference_year
            else:
                raise ValueError("SAMA POS report dates omit a year and provide no year reference.")
        if start_year is None:
            start_year = end_year - int((start_month, start_day) > (end_month, end_day))
        return date(start_year, start_month, start_day), date(end_year, end_month, end_day)

    period_sequences = []
    period_matches = []
    for line in text.splitlines():
        line_periods = [
            tuple(re.sub(r"\s+", " ", value).strip() for value in period)
            for period in re.findall(period_pattern, line)
            if any(re.search(r"[A-Za-z]{3}", value) for value in period)
        ]
        period_matches.extend(line_periods)
        if len(line_periods) >= 2:
            period_sequences.append(line_periods)
    descending_sequence = next(
        (
            sequence for sequence in reversed(period_sequences)
            if all(
                period_range(sequence[index])[0] > period_range(sequence[index + 1])[0]
                for index in range(len(sequence) - 1)
            )
        ),
        None,
    )
    if descending_sequence:
        periods = sorted(
            set(period_matches),
            key=lambda period: period_range(period)[0],
            reverse=True,
        )[:4]
    elif period_sequences and len(period_sequences[-1]) >= 4:
        periods = period_sequences[-1][-4:]
    else:
        periods = list(dict.fromkeys(period_matches))
    total = re.search(
        r"(?m)^\s*Total[ \t]*(?:\r?\n[ \t]*)?([^\r\n]+)",
        text,
    )
    if len(periods) != 4 or not total:
        raise ValueError("SAMA POS report is missing its four weekly periods or total activity row.")

    total_row = re.sub(r"(?<=\d),\s+(?=\d)", ",", total.group(1))
    values = re.findall(r"(?<![\w.])-?\d[\d,]*(?:\.\d+)?", total_row)
    if len(values) < 8:
        raise ValueError("SAMA POS total row does not contain four count/value pairs.")
    try:
        date_ranges = [period_range(period) for period in periods]
        pairs = [(float(values[index].replace(",", "")), float(values[index + 1].replace(",", "")))
                 for index in range(0, 8, 2)]
    except ValueError as error:
        raise ValueError("SAMA POS report contains an invalid date or total value.") from error

    rows = []
    for (period_start, period_end), (count, value) in zip(date_ranges, pairs):
        if period_end < period_start or count < 0 or value < 0:
            raise ValueError("SAMA POS report contains an invalid weekly period or total.")
        for measure, amount, unit in (
            ("transaction count", count, "thousand transactions"),
            ("transaction value", value, "SAR thousand"),
        ):
            fact = make_fact(
                "KPI_02",
                period_start.year,
                period_start.month,
                amount,
                unit,
                f"SAMA POS total {measure}",
                source_url,
                "Weekly total across all activities from the SAMA weekly POS transactions report; "
                "the source reports both measures in thousands.",
            )
            fact["PeriodStart"] = period_start.isoformat()
            fact["PeriodEnd"] = period_end.isoformat()
            fact["MonthName"] = f"{period_start.strftime('%b %d')}-{period_end.strftime('%b %d')}"
            rows.append(fact)
    return rows


ARABIC_MONTHS = {
    "يناير": 1,
    "فبراير": 2,
    "مارس": 3,
    "ابريل": 4,
    "أبريل": 4,
    "مايو": 5,
    "يونيو": 6,
    "يونيه": 6,
    "يوليو": 7,
    "اغسطس": 8,
    "أغسطس": 8,
    "سبتمبر": 9,
    "اكتوبر": 10,
    "أكتوبر": 10,
    "نوفمبر": 11,
    "ديسمبر": 12,
}


def _html_table_cells(table_html: str) -> list[list[str]]:
    rows = re.findall(r"<tr\b[^>]*>(.*?)</tr>", table_html, flags=re.IGNORECASE | re.DOTALL)
    parsed = []
    for row in rows:
        cells = re.findall(
            r"<(?:th|td)\b[^>]*>(.*?)</(?:th|td)>",
            row,
            flags=re.IGNORECASE | re.DOTALL,
        )
        parsed.append(
            [
                re.sub(r"\s+", " ", unescape(re.sub(r"<[^>]+>", " ", cell))).strip()
                for cell in cells
            ]
        )
    return parsed


def extract_yamama_cement(content: bytes, source_url: str) -> list[dict[str, Any]]:
    page = content.decode("utf-8", "replace")
    table_match = re.search(
        r'<table\b[^>]*id=["\']monthlySalesOfCementTable["\'][^>]*>.*?</table>',
        page,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if not table_match:
        raise ValueError("Yamama Cement page is missing its monthly company sales table.")

    table = _html_table_cells(table_match.group())
    if len(table) < 11 or len(table[0]) < 2:
        raise ValueError("Yamama Cement monthly sales table is incomplete.")

    periods: list[tuple[int, int]] = []
    for label in table[0][1:]:
        match = re.fullmatch(r"(\d{4})-(.+)", label)
        month = ARABIC_MONTHS.get(match.group(2).strip()) if match else None
        if not match or not month:
            raise ValueError(f"Yamama Cement table contains an unknown month label: {label!r}.")
        periods.append((int(match.group(1)), month))
    if len(periods) != len(set(periods)):
        raise ValueError("Yamama Cement monthly sales table repeats a month.")

    totals = [0.0] * len(periods)
    companies = 0
    for row in table[1:]:
        if len(row) != len(periods) + 1 or not row[0]:
            raise ValueError("Yamama Cement monthly sales table has a malformed company row.")
        values = row[1:]
        if not any(value.strip() for value in values):
            continue
        companies += 1
        for index, value in enumerate(values):
            if not value.strip() or value.strip() == "-":
                continue
            try:
                amount = float(value.replace(",", ""))
            except ValueError as error:
                raise ValueError(
                    f"Yamama Cement table contains an invalid sales value for {periods[index]}."
                ) from error
            if amount < 0:
                raise ValueError(f"Yamama Cement table contains negative sales for {periods[index]}.")
            totals[index] += amount
    if companies < 10:
        raise ValueError(f"Yamama Cement table contained only {companies} company rows.")

    rows = []
    for (year, month), value in zip(periods, totals):
        if value == 0:
            continue
        rows.append(
            make_fact(
                "KPI_03",
                year,
                month,
                value,
                "thousand tonnes",
                f"Yamama Cement monthly sales total ({companies} Saudi cement companies)",
                source_url,
                "Sum of monthly sales across the Saudi cement companies listed in Yamama Cement's report. "
                f"Argaam's September 2026 report independently confirms 4,439 thousand tonnes "
                f"({ARGAAM_CEMENT_REFERENCE_URL}).",
            )
        )
    if not rows:
        raise ValueError("Yamama Cement monthly sales table contains no reported totals.")
    return sorted(rows, key=lambda row: row["PeriodStart"])


def yamama_cement_report_url() -> str:
    latest = datetime.now(RIYADH_TZ)
    query = urlencode(
        {
            "IsRequested": "true",
            "FirstYear": 2000,
            "FirstMonth": 1,
            "SecondYear": latest.year,
            "SecondMonth": latest.month,
            "ReportType": 4,
        }
    )
    return f"{YAMAMA_CEMENT_URL}?{query}"


def monthly_rows(sheet, year_index: int, month_index: int):
    for row in sheet.iter_rows(values_only=True):
        if len(row) <= max(year_index, month_index):
            continue
        month = MONTHS.get(str(row[month_index] or "").strip().lower())
        if not month:
            continue
        try:
            year = int(str(row[year_index]).replace("*", "").strip())
        except (TypeError, ValueError):
            continue
        yield year, month, row


def extract_trade(content: bytes, source_url: str) -> dict[str, list[dict[str, Any]]]:
    workbook = load_workbook(BytesIO(content), read_only=True, data_only=True)
    result: dict[str, list[dict[str, Any]]] = {
        key: [] for key in ("KPI_05", "KPI_07", "KPI_08", "KPI_11", "KPI_12")
    }
    for year, month, row in monthly_rows(workbook["1.1"], 0, 2):
        if row[6] is not None:
            result["KPI_05"].append(make_fact(
                "KPI_05", year, month, row[6], "SAR million",
            (
                row[3],
                "GASTAT ITR 1.2 - national non-oil exports",
                "National non-oil exports, excluding re-exports.",
            ),
            (
                row[11],
                "GASTAT ITR 1.2 - re-exports",
                "Re-exports, reported separately from national non-oil exports.",
            ),
        ):
            if value is None:
                continue
            fact = make_fact(
                "KPI_08", year, month, value, "SAR million",
                source_table, source_url, source_note, observation_status,
            )
            result["KPI_08"].append(fact)
            result.setdefault("KPI_08_Components", []).append(dict(fact))
