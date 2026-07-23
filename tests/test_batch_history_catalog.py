import contextlib
import io
import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import get_tsys_batch as report


class BatchHistoryCatalogTests(unittest.TestCase):
    def setUp(self):
        self.records = [
            {
                "accountNumber": "A",
                "termID": "T1",
                "created": "2026-01-01T10:00:00Z",
            },
            {
                "accountNumber": "A",
                "termID": "T1",
                "created": "2026-01-03T10:00:00Z",
            },
            {
                "accountNumber": "A",
                "termID": "T2",
                "batchDate": "2026-01-02T10:00:00Z",
            },
            {
                "accountNumber": "B",
                "termID": "",
                "created": "2026-01-04T10:00:00Z",
            },
            {
                "accountNumber": "C",
                "termID": "T3",
                "batchDate": "2026-01-05T10:00:00Z",
            },
            {
                "accountNumber": "",
                "termID": "T4",
                "created": "2026-01-06T10:00:00Z",
            },
        ]

    def test_catalog_builds_all_history_views_from_same_records(self):
        catalog = report.BatchHistoryCatalog(self.records)

        self.assertEqual(
            catalog.last_batch_by_account,
            {
                "A": "01/03/2026 10:00:00 AM",
                "B": "01/04/2026 10:00:00 AM",
                "C": "01/05/2026 10:00:00 AM",
            },
        )
        self.assertEqual(
            catalog.term_history_by_account,
            {
                "A": {
                    "termIDs": ["T1", "T2"],
                    "lastBatchDate": "01/03/2026 10:00:00 AM",
                },
                "C": {
                    "termIDs": ["T3"],
                    "lastBatchDate": "01/05/2026 10:00:00 AM",
                },
            },
        )
        self.assertEqual(
            catalog.term_history_rows,
            [
                {
                    "accountNumber": "A",
                    "termID": "T1",
                    "lastBatchDate": "01/03/2026 10:00:00 AM",
                },
                {
                    "accountNumber": "A",
                    "termID": "T2",
                    "lastBatchDate": "01/02/2026 10:00:00 AM",
                },
                {
                    "accountNumber": "C",
                    "termID": "T3",
                    "lastBatchDate": "01/05/2026 10:00:00 AM",
                },
            ],
        )

    def test_compatibility_helpers_match_catalog_views(self):
        catalog = report.BatchHistoryCatalog(self.records)

        self.assertEqual(report.batch_index(self.records), (catalog.by_pair, catalog.term_accounts))
        self.assertEqual(report.latest_batch_dates_by_account(self.records), catalog.last_batch_by_account)
        self.assertEqual(report.term_history_by_account(self.records), catalog.term_history_by_account)
        self.assertEqual(report.build_term_history_rows(self.records), catalog.term_history_rows)

    def test_report_emits_null_term_without_term_history(self):
        catalog = report.BatchHistoryCatalog(self.records)
        roster = [
            {
                "account": {
                    "number": "A",
                    "product": {"id": "3"},
                    "active": True,
                    "status": "Open",
                },
                "location": {"name": "Store A"},
            },
            {
                "account": {
                    "number": "B",
                    "product": {"id": "3"},
                    "active": True,
                    "status": "Open",
                },
                "location": {"name": "Store B"},
            },
        ]
        activity_summary = {
            "A": {
                "activityCount": 1,
                "approvedAmount": 10.0,
                "approvedCount": 1,
                "terminals": {"AUTH-TERM-A"},
            },
            "B": {
                "activityCount": 1,
                "approvedAmount": 20.0,
                "approvedCount": 1,
                "terminals": {"AUTH-TERM-B"},
            },
        }

        email_rows, review_rows = report.create_store_report(
            roster,
            [],
            activity_summary,
            [self.records[1]],
            catalog.last_batch_by_account,
            catalog.term_history_by_account,
            {"B": 20.0},
            "2026-01-01",
        )

        self.assertEqual(len(email_rows), 1)
        self.assertEqual(email_rows[0]["accountNumber"], "B")
        self.assertEqual(email_rows[0]["termID"], "NULL")
        self.assertEqual(email_rows[0]["authorizedUnbatchedAmount"], "20.00")
        self.assertEqual(review_rows, [])

    def test_stale_rules_use_any_activity_and_raw_batch_status(self):
        activity = report.build_account_activity_summary(
            [
                {
                    "accountNumber": "A",
                    "authorizedAmount": "12.50",
                    "authorizationResponseStatus": "Declined",
                },
                {
                    "accountNumber": "A",
                    "authorizedAmount": "7.50",
                    "authorizationResponseStatus": "approved",
                },
            ]
        )

        self.assertEqual(activity["A"]["activityCount"], 2)
        self.assertEqual(activity["A"]["approvedCount"], 1)
        self.assertEqual(activity["A"]["approvedAmount"], 7.5)
        self.assertEqual(report.batch_status({"rejected": "Yes"}), "rejected")
        self.assertEqual(report.batch_status({"rejected": "No"}), "accepted")
        self.assertEqual(report.batch_status({}), "unknown")

        roster = [
            {
                "account": {
                    "number": "A",
                    "product": {"id": "3"},
                    "active": True,
                    "status": "Open",
                },
                "location": {"name": "Store A"},
            },
        ]
        email_rows, review_rows = report.create_store_report(
            roster,
            [],
            activity,
            [],
            {},
            {},
            {"A": 7.5},
            "2026-01-01",
        )
        self.assertEqual(len(email_rows), 1)
        self.assertEqual(email_rows[0]["accountNumber"], "A")
        self.assertEqual(email_rows[0]["termID"], "NULL")
        self.assertEqual(review_rows, [])

    def test_absolute_authorization_url_contains_explicit_window(self):
        url = report.authorization_absolute_url(
            "https://fixture.invalid",
            "2026-01-01",
            "2026-01-03",
            {"B", "A"},
        )

        self.assertIn("dr_type=abs", url)
        self.assertIn("2026-01-01T00%3A00%3A00.000Z", url)
        self.assertIn("2026-01-03T23%3A59%3A59.999Z", url)

    def test_any_batch_record_suppresses_stale_candidate(self):
        roster = [
            {
                "account": {
                    "number": "A",
                    "product": {"id": "3"},
                    "active": True,
                    "status": "Open",
                },
                "location": {"name": "Store A"},
            },
        ]
        activity = report.build_account_activity_summary(
            [{"accountNumber": "A", "authorizationResponseStatus": "Declined"}]
        )

        email_rows, review_rows = report.create_store_report(
            roster,
            [],
            activity,
            [{"accountNumber": "A", "rejected": "Yes"}],
            {},
            {},
            {"A": 10.0},
            "2026-01-01",
        )

        self.assertEqual(email_rows, [])
        self.assertEqual(review_rows, [])


class AccountReportCatalogTests(unittest.TestCase):
    def setUp(self):
        self.history_views = {
            "lastBatchByAccount": {"A": "01/03/2026 10:00:00 AM"},
            "termHistoryByAccount": {
                "A": {
                    "termIDs": ["T1"],
                    "lastBatchDate": "01/03/2026 10:00:00 AM",
                },
            },
            "termHistoryRows": [
                {
                    "accountNumber": "A",
                    "termID": "T1",
                    "lastBatchDate": "01/03/2026 10:00:00 AM",
                },
            ],
        }
        self.fingerprint = "history-fingerprint"

    def test_constant_value_promotes_after_three_reports_and_resets_on_change(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            catalog = report.AccountReportCatalog(Path(temp_dir) / "catalog.json")
            observations = {"A": {"storeName": "Store A"}}

            for report_number in range(1, 3):
                catalog.record_report(
                    f"run-{report_number}", observations, self.fingerprint, self.history_views
                )
                self.assertEqual(catalog.stable_account_values("storeName"), {})

            catalog.record_report("run-3", observations, self.fingerprint, self.history_views)
            self.assertEqual(catalog.stable_account_values("storeName"), {"A": "Store A"})

            catalog.record_report(
                "run-4",
                {"A": {"storeName": "Renamed Store A"}},
                "changed-history",
                self.history_views,
            )
            self.assertEqual(catalog.stable_account_values("storeName"), {})

    def test_matching_history_view_is_reused_only_after_three_reports(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            catalog = report.AccountReportCatalog(Path(temp_dir) / "catalog.json")
            for report_number in range(1, 4):
                catalog.record_report(
                    f"run-{report_number}",
                    {},
                    self.fingerprint,
                    self.history_views,
                )
                if report_number < 3:
                    self.assertIsNone(catalog.cached_history_views(self.fingerprint))

            self.assertEqual(catalog.cached_history_views(self.fingerprint), self.history_views)

    def test_catalog_integrity_failure_disables_reuse(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / report.REPORT_CATALOG_FILENAME
            catalog = report.AccountReportCatalog(path)
            for report_number in range(1, 4):
                catalog.record_report(
                    f"run-{report_number}",
                    {},
                    self.fingerprint,
                    self.history_views,
                )
            catalog.save()

            reloaded = report.AccountReportCatalog(path)
            self.assertTrue(reloaded.valid)
            self.assertEqual(reloaded.cached_history_views(self.fingerprint), self.history_views)

            path.write_text(path.read_text(encoding="utf-8").replace("history-fingerprint", "tampered"), encoding="utf-8")
            reloaded = report.AccountReportCatalog(path)

            self.assertFalse(reloaded.valid)
            self.assertIsNone(reloaded.cached_history_views(self.fingerprint))


class EndToEndFixtureTests(unittest.TestCase):
    def _run_fixture_report(self, config_path, history_records):
        current_records = [
            {
                "accountNumber": "B",
                "termID": "TB",
                "rejected": "No",
                "created": "2026-01-04T10:00:00Z",
            },
        ]
        uar_records = [
            {
                "account": {
                    "number": "A",
                    "product": {"id": "3"},
                    "active": True,
                    "status": "Open",
                },
                "location": {"name": "Store A"},
            },
            {
                "account": {
                    "number": "B",
                    "product": {"id": "3"},
                    "active": True,
                    "status": "Open",
                },
                "location": {"name": "Store B"},
            },
        ]
        authorization_records = [
            {
                "accountNumber": "A",
                "terminalNumber": "AUTH-A",
                "authorizedAmount": "10.00",
                "authorizationResponseStatus": "Approved",
            },
        ]

        class FixtureClient:
            base_url = "https://fixture.invalid"

            def __init__(self):
                self.batch_request_count = 0

            def authenticate(self):
                return None

            def request(self, method, url, payload=None):
                if method == "GET":
                    return uar_records
                if report.BATCH_PATH in url:
                    self.batch_request_count += 1
                    records = current_records if self.batch_request_count == 1 else history_records
                    return {"records": records, "totalRecords": len(records)}
                if report.AUTHORIZATION_PATH in url:
                    return {"records": authorization_records, "totalRecords": len(authorization_records)}
                raise AssertionError(f"Unexpected fixture request: {method} {url}")

        args = SimpleNamespace(
            ui=False,
            run_report=True,
            config=str(config_path),
            batch_days=3,
            historical_days=4,
            auth_check_days=None,
            auth_window=None,
            log_file=None,
        )
        output = io.StringIO()
        with patch.object(report, "parse_args", return_value=args), \
                patch.object(report, "make_client", return_value=FixtureClient()), \
                contextlib.redirect_stdout(output):
            self.assertEqual(report.main(), 0)
        return output.getvalue()

    def test_main_fixture_regression_promotes_reuses_and_invalidates_catalog(self):
        history_records = [
            {
                "accountNumber": "A",
                "termID": "TA",
                "rejected": "No",
                "created": "2026-01-03T10:00:00Z",
            },
        ]
        changed_history_records = history_records + [
            {
                "accountNumber": "A",
                "termID": "TA2",
                "rejected": "No",
                "created": "2026-01-02T10:00:00Z",
            },
        ]

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            output_root = root / "reports"
            config_path = root / "config.json"
            config_path.write_text(
                json.dumps(
                    {
                        "apiBaseUrl": "https://fixture.invalid",
                        "batchLookbackDays": 3,
                        "historicalLookbackDays": 4,
                        "authorizationWindow": "last_24_h",
                        "outputDirectory": str(output_root),
                        "devices": [],
                    }
                ),
                encoding="utf-8",
            )

            run_outputs = [
                self._run_fixture_report(config_path, history_records)
                for _ in range(3)
            ]
            reuse_output = self._run_fixture_report(config_path, history_records)
            changed_output = self._run_fixture_report(config_path, changed_history_records)

            self.assertTrue(all("Stores flagged: 1" in output for output in run_outputs))
            self.assertIn("Reusing the validated three-report history catalog.", reuse_output)
            self.assertNotIn("Reusing the validated three-report history catalog.", changed_output)

            run_directories = sorted(path for path in output_root.iterdir() if path.is_dir())
            self.assertEqual(len(run_directories), 5)
            for run_directory in run_directories:
                html_path = run_directory / "BottlePOS PAX Batch Report.html"
                self.assertTrue(html_path.exists())
                html = html_path.read_text(encoding="utf-8")
                self.assertIn("PINPAD_BATCH_NOT_CLOSED", html)
                self.assertIn("TERMID_ACCOUNT_HISTORY", html)

            catalog_path = output_root / report.REPORT_CATALOG_FILENAME
            catalog = report.AccountReportCatalog(catalog_path)
            self.assertTrue(catalog.valid)
            self.assertEqual(catalog._payload["historyView"]["reportCount"], 1)
            self.assertEqual(catalog.stable_account_values(report.TERM_HISTORY_CONSTANT_FIELD), {})


class InteractiveSummaryHtmlTests(unittest.TestCase):
    def test_dataset_tabs_include_filtered_csv_export_support(self):
        data_specs = [
            ("Primary exception report", "PINPAD_BATCH_NOT_CLOSED.csv", ["accountNumber"], []),
            ("Store review", "NEEDS_MAPPING_OR_REVIEW.csv", ["accountNumber"], []),
            ("Term/account history", "TERMID_ACCOUNT_HISTORY.csv", ["accountNumber", "termID"], []),
            ("Batch history", "BATCH_HISTORY.csv", ["accountNumber", "termID"], []),
        ]

        with tempfile.TemporaryDirectory() as temp_dir:
            html_path = Path(temp_dir) / "summary.html"
            report.write_interactive_summary_html(
                html_path,
                3,
                "last_24_h",
                datetime(2026, 1, 1),
                datetime(2026, 1, 3),
                data_specs,
            )
            html = html_path.read_text(encoding="utf-8")

        self.assertIn('class="export-button"', html)
        self.assertIn('>Export CSV</button>', html)
        self.assertIn("function getVisibleRows(datasetId)", html)
        self.assertIn("function exportDataset(datasetId)", html)
        self.assertIn("link.download = exportFilename(dataset)", html)
        self.assertIn("const csv = [", html)


if __name__ == "__main__":
    unittest.main()
