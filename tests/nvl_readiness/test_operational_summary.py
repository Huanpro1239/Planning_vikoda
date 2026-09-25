"""Contracts for NVL operational run summaries."""

from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest

from nvl.ops_summary import build_operational_summary, render_markdown


class NVLOperationalSummaryTests(unittest.TestCase):
    def _write(self, root: Path, name: str, payload: dict):
        (root / name).write_text(
            json.dumps(payload),
            encoding="utf-8",
        )

    def test_success_summary_surfaces_release_and_changed_counts(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            self._write(
                root,
                "nvl_production_readiness_report.json",
                {
                    "status": "passed",
                    "gate_version": "nvl_production_readiness_v1",
                },
            )
            self._write(
                root,
                "nvl_stock_report.json",
                {
                    "mode": "publish",
                    "status": "published",
                    "metrics": {"changed_count": 5},
                    "target": {"revision": "etag-before"},
                },
            )
            self._write(
                root,
                "nvl_open_po_report.json",
                {
                    "published": True,
                    "target": {"changed_cells": 3},
                    "publish_evidence": {
                        "target_revision_after": "etag-after",
                    },
                },
            )
            self._write(
                root,
                "nvl_release_manifest.json",
                {
                    "release_id": "nvl-release-1",
                    "release_version": "v1",
                    "published_workbook": {"sha256": "workbook-sha"},
                },
            )
            self._write(
                root,
                "nvl_release_index.json",
                {
                    "status": "passed",
                    "summary": {
                        "release_count": 4,
                        "failed_count": 0,
                        "warning_count": 0,
                    },
                },
            )

            result = build_operational_summary(
                root,
                job_status="success",
                event_name="schedule",
                run_id="123",
                commit_sha="abc",
            )

        self.assertEqual(result["outcome"], "PASS")
        self.assertEqual(result["stock"]["changed_D"], 5)
        self.assertEqual(result["open_po"]["changed_E"], 3)
        self.assertEqual(
            result["release"]["release_id"],
            "nvl-release-1",
        )
        self.assertEqual(
            result["release"]["workbook_sha256"],
            "workbook-sha",
        )
        self.assertEqual(result["ledger"]["status"], "passed")

    def test_failure_summary_reports_stock_phase(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            self._write(
                root,
                "nvl_production_readiness_report.json",
                {"status": "passed"},
            )
            self._write(
                root,
                "nvl_stock_report.json",
                {
                    "status": "failed",
                    "mode": "publish",
                    "phase": "post_upload_verify",
                    "message": "server mismatch",
                },
            )
            result = build_operational_summary(
                root,
                job_status="failure",
            )

        self.assertEqual(result["outcome"], "FAIL")
        self.assertEqual(
            result["failed_phase"],
            "post_upload_verify",
        )
        self.assertEqual(result["failure_detail"], "server mismatch")
        markdown = render_markdown(result)
        self.assertIn("**Outcome:** FAIL", markdown)
        self.assertIn("post_upload_verify", markdown)

    def test_readiness_failure_has_priority(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            self._write(
                root,
                "nvl_production_readiness_report.json",
                {
                    "status": "failed",
                    "failed_phase": "etag_concurrency",
                },
            )
            result = build_operational_summary(
                root,
                job_status="failure",
            )

        self.assertEqual(
            result["failed_phase"],
            "etag_concurrency",
        )

    def test_missing_optional_artifacts_still_produces_summary(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            result = build_operational_summary(
                tmpdir,
                job_status="cancelled",
            )

        self.assertEqual(result["outcome"], "FAIL")
        self.assertEqual(result["readiness"]["status"], "missing")
        self.assertEqual(result["ledger"]["status"], "missing")


if __name__ == "__main__":
    unittest.main()
