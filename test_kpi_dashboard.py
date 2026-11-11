import json
import tempfile
import threading
import unittest
from http.server import ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from openpyxl import Workbook, load_workbook

from publish_static import archive_daily_snapshot, build_refresh_log, load_pre_refresh_snapshot
from kpi_dashboard import (
    DashboardHandler,
    MartService,
    DATASAUDI_CAPEX_PAGE,
    DATASAUDI_CAPEX_URL,
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
    extract_sama_pos_text,
    extract_yamama_cement,
    sama_pos_report_urls,
    extract_sama_private_imports,
    latest_sama_bulletin_url,
    KPI06_SHEETS,
)


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


class GastatItrSourceTests(unittest.TestCase):
    def test_source_page_is_reported_separately_from_extraction_route(self) -> None:
        catalog = [{"KPI_ID": "KPI_08"}, {"KPI_ID": "KPI_11"}]
        sources = {}
        apply_configured_source_metadata(catalog, sources)
        self.assertEqual(catalog[0]["SourceURL"], GASTAT_KPI08_SOURCE_PAGE)
        self.assertEqual(sources["KPI_08"]["PrimarySourceURL"], GASTAT_KPI08_SOURCE_PAGE)
        self.assertEqual(sources["KPI_08"]["ExtractedDataURL"], ITR_WORKBOOK_URL)
        self.assertEqual(catalog[1]["SourceURL"], GASTAT_KPI08_SOURCE_PAGE)
        self.assertEqual(sources["KPI_11"]["PrimarySourceURL"], GASTAT_KPI08_SOURCE_PAGE)
        self.assertEqual(sources["KPI_11"]["ExtractedDataURL"], ITR_WORKBOOK_URL)

    def test_it_extracts_oil_export_value_from_monthly_itr(self) -> None:
        workbook = Workbook()
        workbook.active.title = "1.1"
        exports = workbook.create_sheet("1.2")
        row = ["2026*", None, "January", 120.0, None, None, None, 540.0, None, None, None, 20.0]
        exports.append(row)
        end_use = workbook.create_sheet("7.2")
        end_use.cell(row=4, column=1, value="Imports")
        end_use.cell(row=5, column=1, value="Year")
        end_use.cell(row=6, column=1, value="Month")
        content = BytesIO()
        workbook.save(content)

        result = extract_trade(content.getvalue(), "itr-url")

        self.assertEqual(len(result["KPI_11"]), 1)
        fact = result["KPI_11"][0]
        self.assertEqual((fact["PeriodStart"], fact["Value"], fact["Unit"]),
                         ("2026-01-01", 540.0, "SAR million"))
        self.assertEqual(fact["SourceURL"], "itr-url")
        self.assertEqual(fact["ObservationStatus"], "Preliminary")

    def test_newest_monthly_itr_workbook_is_selected(self) -> None:
        page = (
            '<a href="/documents/20117/2435267/ITR+Jun+2026.xlsx/aaa?t=1">x</a>'
            '<a href="/documents/20117/2435267/ITR_non_oil_Sep2026.xlsx/bbb?t=2">x</a>'
            '<a href="/documents/20117/2435267/ITR_Aug+2026.xlsx/ccc?t=3">x</a>'
            '<a href="/documents/20117/2435267/ITR+Aug+2026+AR.pdf/ddd?t=4">x</a>'
            '<a href="/documents/20117/2435267/ITR_Oct_-2025_AR-EN_%283%29.xlsx_fixed_1/eee?t=5">x</a>'
        ).encode()
        self.assertEqual(
            latest_itr_workbook_url(page),
            "https://www.stats.gov.sa/documents/20117/2435267/ITR_Aug+2026.xlsx/ccc?t=3",
        )

    def test_page_without_monthly_itr_workbook_is_rejected(self) -> None:
        with self.assertRaises(ValueError):
            latest_itr_workbook_url(b'<a href="/documents/20117/2435267/ITR_non_oil_Sep2026.xlsx/x">x</a>')


class GastatStsFileTests(unittest.TestCase):
    def test_source_metadata_uses_listing_page_and_retains_latest_file(self) -> None:
        catalog = [{"KPI_ID": "KPI_13"}]
        sources = {"KPI_13": {"ExtractedDataURL": "https://www.stats.gov.sa/old.xlsx"}}
        apply_configured_source_metadata(catalog, sources)
        self.assertEqual(catalog[0]["SourceURL"], GASTAT_KPI13_SOURCE_PAGE)
        self.assertIn(GASTAT_KPI13_SOURCE_PAGE, catalog[0]["API_or_DataRoute"])
        self.assertEqual(sources["KPI_13"]["PrimarySourceURL"], GASTAT_KPI13_SOURCE_PAGE)
        self.assertEqual(sources["KPI_13"]["ExtractedDataURL"], "https://www.stats.gov.sa/old.xlsx")

    def test_newest_monthly_sts_workbook_is_selected_and_pdfs_skipped(self) -> None:
        page = (
            '<a href="/documents/20117/2435267/STS+July+2026+(1).xlsx/old?t=1">July workbook</a>'
            '<a href="/documents/20117/2435267/STS+August+2026+-+AR.pdf/pdf?t=2">August report</a>'
            '<a href="/documents/20117/2435267/STS+Aug+2026+(1).xlsx/new?t=3">August workbook</a>'
            '<a href="/documents/20117/2435267/STS+Sep+2025.xlsx/older?t=4">Old workbook</a>'
        ).encode()
        self.assertEqual(
            latest_sts_workbook_url(page),
            "https://www.stats.gov.sa/documents/20117/2435267/STS+Aug+2026+(1).xlsx/new?t=3",
        )

    def test_page_without_monthly_sts_workbook_is_rejected(self) -> None:
        page = b'<a href="/documents/20117/2435267/STS+August+2026+AR.pdf/pdf">PDF</a>'
        with self.assertRaises(ValueError):
            latest_sts_workbook_url(page)


class GastatGfcfImportsTests(unittest.TestCase):
    def test_only_gfcf_end_use_is_extracted(self) -> None:
        def item(category, period, value):
            return {"values": {"OBS_VALUE": {"sum": value}}, "keys": {"ENDU_EN": category, "YEAR_MONTH": period}}
        payload = {"aggs": [
            item("Gross Fixed Capital Formation", "2026 / 02", 10_000_000_000),
            item("Final Consumption", "2026 / 02", 99_000_000_000),
            item("Gross Fixed Capital Formation", "2026 / 01", 12_500_000_000),
            {"values": {"OBS_VALUE": {"sum": 1}}},
        ], "pageTotal": 1}
        rows = extract_gfcf_imports(payload, "u")
        self.assertEqual([(r["PeriodStart"], r["Value"], r["KPI_ID"]) for r in rows],
                         [("2026-01-01", 12500.0, "KPI_07"), ("2026-02-01", 10000.0, "KPI_07")])

    def test_missing_category_is_rejected(self) -> None:
        with self.assertRaises(ValueError):
            extract_gfcf_imports({"aggs": [], "pageTotal": 1}, "u")


class GastatImportsTests(unittest.TestCase):
    def test_monthly_totals_are_converted_to_sar_million_and_sorted(self) -> None:
        payload = {"aggs": [
            {"values": {"OBS_VALUE": {"sum": 2_500_000_000.0}}, "keys": {"YEAR_MONTH": "2026 / 02"}},
            {"values": {"OBS_VALUE": {"sum": 1_000_000_000}}, "keys": {"YEAR_MONTH": "2025 / 12"}},
            {"values": {"OBS_VALUE": {"sum": 3_500_000_000.0}}},
        ], "page": 0, "pageTotal": 1}
        rows = extract_imports_cpc(payload, "https://database.stats.gov.sa/x")
        self.assertEqual([(r["PeriodStart"], r["Value"], r["KPI_ID"]) for r in rows],
                         [("2025-12-01", 1000.0, "KPI_05"), ("2026-02-01", 2500.0, "KPI_05")])

    def test_empty_or_paged_response_is_rejected(self) -> None:
        with self.assertRaises(ValueError):
            extract_imports_cpc({"aggs": [], "pageTotal": 1}, "u")
        with self.assertRaises(ValueError):
            extract_imports_cpc({"aggs": [], "pageTotal": 2}, "u")


class RefreshLogTests(unittest.TestCase):
    def test_builds_summary_and_entries_from_latest_snapshot(self) -> None:
        log = build_refresh_log(
            {
                "lastRefresh": "2026-11-04T08:00:00+03:00",
                "lastError": None,
                "kpis": [
                    {
                        "id": "KPI_01",
                        "name": "PMI",
                        "refreshStatus": "green",
                        "refreshLabel": "Live extraction succeeded",
                        "observationCount": 12,
                        "latestObservation": "2026-10-31",
                        "refreshNote": "Extracted successfully",
                        "lastChecked": "2026-11-04T08:00:00+03:00",
                    },
                    {
                        "id": "KPI_02",
                        "name": "POS",
                        "refreshStatus": "red",
                        "refreshLabel": "Live extraction failed",
                        "observationCount": 0,
                        "refreshNote": "Source timed out",
                    },
                ],
            },
            "2026-11-04T08:01:00+03:00",
        )
        self.assertEqual(log["summary"], {"kpiCount": 2, "successCount": 1, "failedCount": 1})
        self.assertEqual(log["entries"][0]["timestamp"], "2026-11-04T08:00:00+03:00")
        self.assertEqual(log["entries"][1]["detail"], "Source timed out")
        self.assertEqual(log["entries"][1]["timestamp"], "2026-11-04T08:00:00+03:00")


class DailyArchiveTests(unittest.TestCase):
    def test_creates_one_immutable_snapshot_per_riyadh_date(self) -> None:
        import gzip

        with tempfile.TemporaryDirectory() as directory:
            archive_dir = Path(directory)
            status = {"kpis": [{"id": "KPI_01"}]}
            facts = [{"KPI_ID": "KPI_01", "Value": 10}]
            first = archive_daily_snapshot(
                archive_dir, status, facts, "2026-11-11T10:07:00+03:00"
            )
            path = archive_dir / "2026-11-11.json.gz"
            with gzip.open(path, "rt", encoding="utf-8") as file:
                archived = json.load(file)
            second = archive_daily_snapshot(
                archive_dir, status, [{"KPI_ID": "KPI_01", "Value": 99}],
                "2026-11-11T15:00:00+03:00",
            )

            self.assertTrue(first["created"])
            self.assertEqual(archived["facts"], facts)
            self.assertFalse(second["created"])
            with gzip.open(path, "rt", encoding="utf-8") as file:
                self.assertEqual(json.load(file)["facts"], facts)

            index = json.loads((archive_dir / "index.json").read_text(encoding="utf-8"))
            self.assertEqual(len(index["snapshots"]), 1)
            self.assertEqual(index["snapshots"][0]["factCount"], 1)

    def test_indexes_snapshots_in_newest_first_order(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            archive_dir = Path(directory)
            for timestamp in (
                "2026-11-10T10:07:00+03:00",
                "2026-11-11T10:07:00+03:00",
            ):
                archive_daily_snapshot(archive_dir, {"kpis": []}, [], timestamp)

            index = json.loads((archive_dir / "index.json").read_text(encoding="utf-8"))
            self.assertEqual(
                [item["date"] for item in index["snapshots"]],
                ["2026-11-11", "2026-11-10"],
            )

    def test_uses_current_published_dataset_for_pre_refresh_archive(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory) / "data"
            published_status = {"generatedAt": "2026-11-11T10:07:00+03:00", "kpis": []}
            published_facts = [{"KPI_ID": "KPI_01", "Value": 10}]
            responses = {
                "https://muhannadfallatah.github.io/saudi-kpi-dashboard/data/status.json":
                    json.dumps(published_status).encode("utf-8"),
                "https://muhannadfallatah.github.io/saudi-kpi-dashboard/data/facts.json":
                    json.dumps({"count": len(published_facts), "data": published_facts}).encode("utf-8"),
            }
            status, facts = load_pre_refresh_snapshot(
                data_dir,
                {"generatedAt": "fallback"},
                [{"KPI_ID": "KPI_01", "Value": 20}],
                fetch=lambda url: responses[url],
            )

            self.assertEqual(status, published_status)
            self.assertEqual(facts, published_facts)

    def test_falls_back_to_local_data_when_site_unavailable(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            data_dir = root / "data"
            data_dir.mkdir()
            status = {"generatedAt": "2026-11-10T10:07:00+03:00", "kpis": []}
            facts = [{"KPI_ID": "KPI_01", "Value": 10}]
            (data_dir / "status.json").write_text(json.dumps(status), encoding="utf-8")
            (data_dir / "facts.json").write_text(
                json.dumps({"count": len(facts), "data": facts}),
                encoding="utf-8",
            )

            def unavailable(_url: str) -> bytes:
                raise OSError("site unavailable")

            with self.assertWarns(RuntimeWarning):
                archived_status, archived_facts = load_pre_refresh_snapshot(
                    data_dir,
                    {"generatedAt": "fallback"},
                    [],
                    fetch=unavailable,
                )

            self.assertEqual(archived_status, status)
            self.assertEqual(archived_facts, facts)


class DataSaudiCapexTests(unittest.TestCase):
    def test_metadata_uses_public_finances_page_and_api(self) -> None:
        catalog = [{"KPI_ID": "KPI_16"}]
        sources = {}
        apply_configured_source_metadata(catalog, sources)
        self.assertEqual(catalog[0]["Source"], "DataSaudi (Ministry of Finance)")
        self.assertEqual(catalog[0]["SourceURL"], DATASAUDI_CAPEX_PAGE)
        self.assertEqual(sources["KPI_16"]["PrimarySourceURL"], DATASAUDI_CAPEX_PAGE)
        self.assertEqual(sources["KPI_16"]["ExtractedDataURL"], DATASAUDI_CAPEX_URL)
        self.assertEqual(catalog[0]["Frequency_or_TimeFrame"], "Quarterly")

    def test_extracts_only_quarterly_capital_expenditure_and_converts_units(self) -> None:
        payload = {
            "columns": ["Quarter ID", "Quarter", "Type ID", "Type", "SAR Billions"],
            "page": {"total": 5},
            "data": [
                {"Quarter ID": 20261, "Quarter": "2026-Q1", "Type ID": 1, "Type": "Oil Revenues", "SAR Billions": 300.0},
                {"Quarter ID": 20261, "Quarter": "2026-Q1", "Type ID": 3, "Type": "Current Expenditures", "SAR Billions": 250.0},
                {"Quarter ID": 20261, "Quarter": "2026-Q1", "Type ID": 4, "Type": "Capital Expenditures", "SAR Billions": 43.433},
                {"Quarter ID": 20262, "Quarter": "2026-Q2", "Type ID": 4, "Type": "Capital Expenditures", "SAR Billions": 46.226},
                {"Quarter ID": 20263, "Quarter": "2026-Q3", "Type ID": 5, "Type": "Deficit Surplus (Actual)", "SAR Billions": 20.0},
            ],
        }
        rows = extract_datasaudi_capex(json.dumps(payload).encode())
        self.assertEqual(
            [(row["PeriodStart"], row["PeriodEnd"], row["Value"], row["Unit"]) for row in rows],
            [
                ("2026-01-01", "2026-03-31", 43433.0, "SAR million"),
                ("2026-04-01", "2026-06-30", 46226.0, "SAR million"),
            ],
        )
        self.assertEqual(rows[0]["SourceURL"], DATASAUDI_CAPEX_PAGE)

    def test_rejects_incomplete_api_response(self) -> None:
        payload = {
            "columns": ["Quarter", "Type ID", "Type", "SAR Billions"],
            "page": {"total": 1},
            "data": [],
        }
        with self.assertRaises(ValueError):
            extract_datasaudi_capex(json.dumps(payload).encode())


class DataSaudiPmiTests(unittest.TestCase):
    def test_extracts_monthly_pmi_facts(self) -> None:
        payload = {
            "page": {"total": 2},
            "columns": ["Month ID", "Month", "Purchasing Manager Index"],
            "data": [
                {"Month ID": 202609, "Month": "2026-09", "Purchasing Manager Index": 55.26107857},
                {"Month ID": 202608, "Month": "2026-08", "Purchasing Manager Index": 53.75571254},
            ],
        }
        rows = extract_datasaudi_pmi(json.dumps(payload).encode("utf-8"))
        self.assertEqual([row["PeriodStart"] for row in rows], ["2026-08-01", "2026-09-01"])
        self.assertEqual(rows[-1]["Value"], 55.26107857)
        self.assertEqual(rows[-1]["Unit"], "index points")
        self.assertEqual(rows[-1]["SourceURL"], "https://datasaudi.sa/en/#purchasing-manager-index")

    def test_rejects_incomplete_api_response(self) -> None:
        payload = {
            "page": {"total": 2},
            "columns": ["Month", "Purchasing Manager Index"],
            "data": [{"Month": "2026-09", "Purchasing Manager Index": 55.3}],
        }
        with self.assertRaisesRegex(ValueError, "incomplete"):
            extract_datasaudi_pmi(json.dumps(payload).encode("utf-8"))


class InvestingCdsTests(unittest.TestCase):
    def test_tasi_extracts_only_index_indicator(self) -> None:
        payload = {
            "page": {"total": 3},
            "data": [
                {"Month": "2026-09", "Indicator": "Index", "Value": 10440.64},
                {"Month": "2026-09", "Indicator": "Banks", "Value": 999.0},
                {"Month": "2026-08", "Indicator": "Index", "Value": 11127.08},
            ],
        }
        rows = extract_datasaudi_tasi(json.dumps(payload).encode("utf-8"))
        self.assertEqual([row["PeriodStart"] for row in rows], ["2026-08-01", "2026-09-01"])
        self.assertEqual(rows[1]["KPI_ID"], "KPI_15")
        self.assertEqual(rows[1]["Value"], 10440.64)
    def test_extracts_daily_close(self) -> None:
        record = {
            "rowDateTimestamp": "2026-10-29T00:00:00Z", "last_closeRaw": "68.1", "last_openRaw": "69.2",
            "last_maxRaw": "69.3", "last_minRaw": "68.0",
        }
        state = {"props": {"pageProps": {"state": {"historicalDataStore": {"historicalData": {"data": [record]}}}}}}
        html = f'<script id="__NEXT_DATA__" type="application/json">{json.dumps(state)}</script>'
        rows = extract_investing_cds(html.encode("utf-8"))
        self.assertEqual(rows[0]["KPI_ID"], "KPI_14")
        self.assertEqual(rows[0]["PeriodStart"], "2026-10-29")
        self.assertEqual(rows[0]["Close"], 68.1)
        self.assertEqual(rows[0]["High"], 69.3)

    def test_rejects_blocked_page(self) -> None:
        with self.assertRaisesRegex(ValueError, "blocked"):
            extract_investing_cds(b"<html>Just a moment</html>")

    def test_history_merge_retains_one_time_initial_load(self) -> None:
        initial = {"KPI_ID": "KPI_14", "PeriodStart": "2025-01-01T00:00:00", "Value": 64.69}
        old_recent = {"KPI_ID": "KPI_14", "PeriodStart": "2026-11-04T00:00:00", "Value": 66.0}
        live_recent = {"KPI_ID": "KPI_14", "PeriodStart": "2026-11-04", "Value": 67.07}
        merged = merge_cds_history([initial, old_recent], [], [live_recent])
        self.assertEqual([fact["PeriodStart"] for fact in merged], ["2025-01-01", "2026-11-04"])
        self.assertEqual(merged[-1]["PeriodEnd"], "2026-11-04")
        self.assertEqual(merged[-1]["Value"], 67.07)


class DataSaudiSaiborTests(unittest.TestCase):
    def test_extracts_monthly_rate(self) -> None:
        payload = {
            "page": {"total": 2},
            "data": [
                {"Month": "2026-09", "Saudi Interbank Offered Rates": 5.0127},
                {"Month": "2026-08", "Saudi Interbank Offered Rates": 4.83},
            ],
        }
        rows = extract_datasaudi_saibor(json.dumps(payload).encode("utf-8"))
        self.assertEqual([row["PeriodStart"] for row in rows], ["2026-08-01", "2026-09-01"])
        self.assertEqual(rows[1]["KPI_ID"], "KPI_09")
        self.assertEqual(rows[1]["Value"], 5.0127)

    def test_rejects_incomplete_response(self) -> None:
        payload = {"page": {"total": 3}, "data": [{"Month": "2026-09", "Saudi Interbank Offered Rates": 5.0}]}
        with self.assertRaisesRegex(ValueError, "incomplete"):
            extract_datasaudi_saibor(json.dumps(payload).encode("utf-8"))


class DataSaudiPrivateLendingTests(unittest.TestCase):
    def test_extracts_only_monthly_grand_total(self) -> None:
        payload = {
            "page": {"total": 4},
            "columns": ["Month", "Lending Type", "Million SAR"],
            "data": [
                {"Month": "2026-08", "Lending Type": "Loans, Advances & Overdrafts", "Million SAR": 3165439.25},
                {"Month": "2026-08", "Lending Type": "Bills Discounted", "Million SAR": 22006.04},
                {"Month": "2026-08", "Lending Type": "Investments in Private Securities", "Million SAR": 109217.17},
                {"Month": "2026-08", "Lending Type": "Grand Total", "Million SAR": 3296662.5},
            ],
        }
        rows = extract_datasaudi_private_lending(json.dumps(payload).encode("utf-8"))
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["KPI_ID"], "KPI_04")
        self.assertEqual(rows[0]["PeriodStart"], "2026-08-01")
        self.assertEqual(rows[0]["Value"], 3296662.5)
        self.assertEqual(rows[0]["Unit"], "SAR million")

    def test_rejects_incomplete_response(self) -> None:
        payload = {
            "page": {"total": 2},
            "columns": ["Month", "Lending Type", "Million SAR"],
            "data": [{"Month": "2026-08", "Lending Type": "Grand Total", "Million SAR": 10}],
        }
        with self.assertRaisesRegex(ValueError, "incomplete"):
            extract_datasaudi_private_lending(json.dumps(payload).encode("utf-8"))


class SamaPosTests(unittest.TestCase):
    def test_finds_report_pdf_links_in_sharepoint_list_data(self) -> None:
        html = r'''{"FileRef":"\u002fen-US\u002fStatistics\u002fIndices\u002fPOS_EN\u002fWeekly_Points_of_Sale_Transactions_Report.pdf"}'''
        urls = sama_pos_report_urls(lambda _: html.encode("utf-8"))
        self.assertEqual(len(urls), 1)
        self.assertTrue(urls[0].endswith("/POS_EN/Weekly_Points_of_Sale_Transactions_Report.pdf"))

    def test_extracts_weekly_count_and_value_as_separate_rows(self) -> None:
        text = """
        27 Sep,26 - 03 Oct,26 04 Oct,26 - 10 Oct,26 11 Oct,26 - 17 Oct,26 18 Oct,26 - 24 Oct,26
        Number of Transactions Value of Transactions
        Total 273,924 16,507,357 253,359 13,525,658 246,421 12,602,805 229,116 11,779,845 -7.0 -6.5
        """
        rows = extract_sama_pos_text(text, "https://www.sama.gov.sa/report.pdf")
        self.assertEqual(len(rows), 8)
        latest = [row for row in rows if row["PeriodStart"] == "2026-10-18"]
        self.assertEqual({row["SourceTable"] for row in latest}, {
            "SAMA POS total transaction count",
            "SAMA POS total transaction value",
        })
        count = next(row for row in latest if row["Unit"] == "thousand transactions")
        value = next(row for row in latest if row["Unit"] == "SAR thousand")
        self.assertEqual((count["Value"], count["PeriodEnd"]), (229116.0, "2026-10-24"))
        self.assertEqual(value["Value"], 11779845.0)

    def test_rejects_report_without_total_or_periods(self) -> None:
        with self.assertRaisesRegex(ValueError, "four weekly periods or total"):
            extract_sama_pos_text("Weekly Points of Sale Transactions", "u")


class YamamaCementTests(unittest.TestCase):
    def test_sums_company_sales_by_month(self) -> None:
        rows = ['<table id="monthlySalesOfCementTable"><tr><th>Company</th><td>2026-سبتمبر</td><td>2026-أغسطس</td></tr>']
        rows.extend(
            f"<tr><th>Company {index}</th><td>100</td><td>90</td></tr>"
            for index in range(10)
        )
        rows.append("</table>")
        facts = extract_yamama_cement("".join(rows).encode("utf-8"), "https://yamama.example/report")
        self.assertEqual(len(facts), 2)
        self.assertEqual(facts[0]["PeriodStart"], "2026-08-01")
        self.assertEqual(facts[0]["Value"], 900.0)
        self.assertEqual(facts[1]["PeriodStart"], "2026-09-01")
        self.assertEqual(facts[1]["Value"], 1000.0)
        self.assertEqual(facts[1]["Unit"], "thousand tonnes")

    def test_rejects_missing_company_sales_table(self) -> None:
        with self.assertRaisesRegex(ValueError, "missing its monthly company sales table"):
            extract_yamama_cement(b"<html></html>", "u")


def make_workbook(path: Path) -> None:
    workbook = Workbook()
    catalog = workbook.active
    catalog.title = "KPI_Catalog"
    catalog.append(
        [
            "KPI_ID",
            "KPI_Name_Arabic",
            "KPI_Name_Original",
            "Source",
            "SourceURL",
            "Frequency_or_TimeFrame",
            "ExtractionStatus",
            "ExtractionNote",
        ]
    )
    catalog.append(["KPI_05", "الواردات", "Imports", "GASTAT", "", "Monthly", "Extracted", "Fixture"])
    catalog.append(["KPI_10", "التراخيص", "Licences", "MISA", "", "Monthly", "Excluded", "Excluded by request"])

    sources = workbook.create_sheet("Source_Register")
    sources.append(["KPI_ID", "ObservedFrom", "ObservedTo", "ExtractionStatus", "ExtractionNote"])
    sources.append(["KPI_05", "2025-01", "2025-01", "Extracted", "Fixture"])
    sources.append(["KPI_10", None, None, "Excluded", "Excluded by request"])

    facts = workbook.create_sheet("Fact_All")
    facts.append(
        [
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
        ]
    )
    facts.append(
        [
            "SAU",
            "KPI_05",
            "2025-01-01",
            "2025-01-31",
            2025,
            1,
            "January",
            12.5,
            "SAR million",
            "Test table",
            "Reported",
            "",
            "",
            None,
            None,
            None,
            None,
            None,
            None,
            None,
        ]
    )
    workbook.save(path)


class MartServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.workbook_path = Path(self.temp_dir.name) / "test.xlsx"
        make_workbook(self.workbook_path)
        self.service = MartService(self.workbook_path, refresh_minutes=60)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_status_includes_all_kpis_and_distinguishes_missing_data(self) -> None:
        result = self.service.snapshot()
        self.assertEqual(len(result["kpis"]), 2)
        self.assertEqual(result["factCount"], 1)
        self.assertEqual(result["kpis"][0]["refreshStatus"], "green")
        self.assertEqual(result["kpis"][0]["name"], "Total goods imports")
        self.assertEqual(result["kpis"][0]["nameArabic"], "الواردات")
        self.assertEqual(result["kpis"][0]["frequency"], "Monthly")
        self.assertEqual(result["kpis"][1]["refreshLabel"], "Excluded by request")
        self.assertEqual(result["kpis"][1]["refreshStatus"], "red")

    def test_facts_can_be_filtered(self) -> None:
        self.assertEqual(len(self.service.facts_for("KPI_05")), 1)
        self.assertEqual(self.service.facts_for("KPI_10"), [])

    def test_refresh_reloads_updated_workbook(self) -> None:
        result = self.service.refresh()
        self.assertTrue(result["ok"])
        self.assertEqual(result["factCount"], 1)
        self.assertIsNotNone(self.service.snapshot()["lastRefresh"])

    def test_excel_export_has_one_named_table_per_kpi(self) -> None:
        workbook = load_workbook(BytesIO(self.service.export_bytes()), data_only=True)
        kpi_sheets = [name for name in workbook.sheetnames if name.startswith("KPI_")]
        self.assertEqual(len(kpi_sheets), 16)
        self.assertNotIn("Refresh_Status", workbook.sheetnames)
        self.assertNotIn("Fact_All", workbook.sheetnames)
        sheet = workbook["KPI_05_Total_Goods_Imports"]
        self.assertEqual(list(sheet.tables), ["tbl_KPI_05_Total_Goods_Imports"])
        self.assertEqual(sheet.max_row, 2)


class DashboardApiTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        path = Path(self.temp_dir.name) / "test.xlsx"
        make_workbook(path)
        DashboardHandler.service = MartService(path, refresh_minutes=60)
        self.server = ThreadingHTTPServer(("127.0.0.1", 0), DashboardHandler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.base_url = f"http://127.0.0.1:{self.server.server_port}"

    def tearDown(self) -> None:
        self.server.shutdown()
        self.server.server_close()
        self.thread.join(timeout=2)
        self.temp_dir.cleanup()

    def test_kpi_api_returns_json(self) -> None:
        with urlopen(f"{self.base_url}/api/kpis") as response:
            self.assertIn("charset=utf-8", response.headers["Content-Type"])
            raw = response.read()
            self.assertIn("الواردات".encode("utf-8"), raw)
            body = json.loads(raw.decode("utf-8"))
        self.assertEqual(len(body["kpis"]), 2)

    def test_filtered_fact_api_returns_rows(self) -> None:
        with urlopen(f"{self.base_url}/api/facts?kpi_id=KPI_05") as response:
            body = json.loads(response.read())
        self.assertEqual(body["count"], 1)
        self.assertEqual(body["data"][0]["Value"], 12.5)

    def test_unknown_kpi_returns_404(self) -> None:
        with self.assertRaises(HTTPError) as error:
            urlopen(f"{self.base_url}/api/kpis/KPI_99")
        self.assertEqual(error.exception.code, 404)

    def test_refresh_endpoint_returns_status(self) -> None:
        request = Request(f"{self.base_url}/api/refresh", method="POST")
        with urlopen(request) as response:
            body = json.loads(response.read())
        self.assertTrue(body["ok"])
        self.assertEqual(body["kpiCount"], 2)

    def test_excel_download_returns_workbook(self) -> None:
        with urlopen(f"{self.base_url}/export.xlsx") as response:
            self.assertEqual(response.status, 200)
            self.assertIn("spreadsheetml", response.headers["Content-Type"])
            workbook = load_workbook(BytesIO(response.read()), read_only=True, data_only=True)
        self.assertIn("KPI_05_Total_Goods_Imports", workbook.sheetnames)
        self.assertEqual(workbook["KPI_05_Total_Goods_Imports"].max_row, 2)


if __name__ == "__main__":
    unittest.main()
