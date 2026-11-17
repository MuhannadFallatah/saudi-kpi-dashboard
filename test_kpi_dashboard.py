import json
import tempfile
import threading
import unittest
from http.server import ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
from unittest.mock import patch
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from openpyxl import Workbook, load_workbook

from publish_static import (
    archive_daily_snapshot,
    archive_refresh_run,
    build_refresh_log,
    last_successful_archive_facts,
    load_pre_refresh_snapshot,
)
from kpi_dashboard import (
    DashboardHandler,
    MartService,
    DATASAUDI_CAPEX_PAGE,
    DATASAUDI_CAPEX_URL,
    WORKBOOK_PATH,
    GASTAT_KPI08_SOURCE_PAGE,
    GASTAT_KPI13_SOURCE_PAGE,
    ITR_WORKBOOK_URL,
    apply_configured_source_metadata,
    extract_datasaudi_capex,
    extract_datasaudi_pmi,
    extract_gfcf_imports,
    extract_imports_cpc,
    latest_itr_workbook_url,
    latest_sts_workbook_url,
    extract_trade,
    extract_total_exports,
    extract_datasaudi_private_lending,
    extract_datasaudi_saibor,
    extract_investing_cds,
    merge_cds_history,
    extract_datasaudi_tasi,
    workbook_data,
    extract_sama_pos_text,
    extract_yamama_cement,
    sama_pos_report_urls,
    extract_sama_private_imports,
    latest_sama_bulletin_url,
    KPI06_SHEETS,
)


class TasiInitialBaselineTests(unittest.TestCase):
    def test_daily_initial_load_covers_full_source_history(self) -> None:
        rows = [
            fact for fact in workbook_data(WORKBOOK_PATH)["facts"]
            if fact["KPI_ID"] == "KPI_15"
        ]
        periods = [fact["PeriodStart"][:10] for fact in rows]
        self.assertEqual(len(rows), 4958)
        self.assertEqual(len(set(periods)), 4958)
        self.assertEqual(min(periods), "2007-01-06")
        self.assertEqual(max(periods), "2026-11-15")
        self.assertTrue(all(fact["Value"] == fact["Close"] for fact in rows))
        self.assertTrue(all(fact["Unit"] == "index points" for fact in rows))

    def test_failed_live_refresh_preserves_daily_baseline(self) -> None:
        service = MartService(WORKBOOK_PATH, 1440)
        with (
            patch("kpi_dashboard.download_bytes", side_effect=OSError("source unavailable")),
            patch("kpi_dashboard.post_json", side_effect=OSError("source unavailable")),
        ):
            result = service.refresh(live=True)

        self.assertTrue(result["ok"])
        rows = service.facts_for("KPI_15")
        self.assertEqual(len(rows), 4958)
        self.assertEqual(rows[0]["PeriodStart"][:10], "2007-01-06")
        self.assertEqual(rows[-1]["PeriodStart"][:10], "2026-11-15")
        kpi = next(item for item in service.snapshot()["kpis"] if item["id"] == "KPI_15")
        self.assertEqual(kpi["refreshLabel"], "Historical baseline loaded")
        self.assertFalse(kpi["liveExtractionConfigured"])


class SamaPrivateImportsTests(unittest.TestCase):
    def _workbook(self) -> bytes:
        from datetime import datetime
        workbook = Workbook()
        workbook.remove(workbook.active)
        for name, offset in (("20", 0), ("21", 1000)):
            sheet = workbook.create_sheet(name)
            sheet.append(["Table"])
            sheet.append(["Period", None, "Grand", "Total"][:2] + ["Other", "Grand"])
            sheet.append([1993, None, 1, offset + 111])
            sheet.append([datetime(2024, 12, 31), None, 1, offset + 222])
            sheet.append(["Q1", 1993, 1, offset + 333])
            sheet.append([datetime(2026, 8, 31), None, 1, offset + 5])
            sheet.append([datetime(2026, 9, 30), None, 1, None])
        buffer = BytesIO()
        workbook.save(buffer)
        return buffer.getvalue()

    def test_extracts_monthly_grand_total_only(self) -> None:
        facts = extract_sama_private_imports(self._workbook(), "https://example/file.xlsx")
        self.assertEqual(len(facts), 2)
        by_table = {fact["SourceTable"]: fact for fact in facts}
        self.assertEqual(set(by_table), set(KPI06_SHEETS.values()))
        self.assertEqual(by_table[KPI06_SHEETS["20"]]["Value"], 5.0)
        self.assertEqual(by_table[KPI06_SHEETS["21"]]["Value"], 1005.0)
        self.assertEqual(by_table[KPI06_SHEETS["20"]]["PeriodStart"], "2026-08-01")

    def test_picks_newest_bulletin(self) -> None:
        page = (
            b'{"FileRef":"\\u002far-sa\\u002fStatistics\\u002fMonthlyStatistics\\u002fMonthly_Statistical_Bulletin_For_Aug_2026.xlsx"}'
            b'{"FileRef":"/ar-sa/Statistics/MonthlyStatistics/Monthly_Statistical_Bulletin_For_Sep_2026.xlsx"}'
        )
        self.assertTrue(latest_sama_bulletin_url(page).endswith("For_Sep_2026.xlsx"))


class GastatTotalExportsTests(unittest.TestCase):
    def test_end_use_categories_are_summed_per_month(self) -> None:
        def item(category, period, value):
            return {"values": {"OBS_VALUE": {"sum": value}}, "keys": {"ENDU_EN": category, "YEAR_MONTH": period}}
        payload = {"aggs": [item("Other", "2026 / 07", 1_000_000_000), item("Final Consumption", "2026 / 07", 2_000_000_000),
                            {"values": {"OBS_VALUE": {"sum": 9}}, "keys": {"YEAR_MONTH": "2026 / 07"}}], "pageTotal": 1}
        rows = extract_total_exports(payload, "u")
        self.assertEqual([(r["PeriodStart"], r["Value"], r["KPI_ID"]) for r in rows], [("2026-07-01", 3000.0, "KPI_12")])

        kpi08_facts = result["KPI_08"]
        self.assertEqual(
            [
                (fact["PeriodStart"], fact["Value"], fact["SourceTable"])
                for fact in kpi08_facts
            ],
            [
                ("2026-01-01", 120.0, "GASTAT ITR 1.2 - national non-oil exports"),
                ("2026-01-01", 20.0, "GASTAT ITR 1.2 - re-exports"),
            ],
        )
        self.assertEqual(
            [fact["SourceTable"] for fact in result["KPI_08_Components"]],
            [fact["SourceTable"] for fact in kpi08_facts],
        )
        self.assertTrue(all(fact["ObservationStatus"] == "Preliminary" for fact in kpi08_facts))
