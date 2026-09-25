"""Release manifest contracts for successful NVL production publishes."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import tempfile
import unittest

from nvl.release import (
    NVL_RELEASE_SCHEMA,
    build_release_manifest,
    stable_workbook_sha256,
    write_release_manifest,
)
from tests.test_sync_nvl_stock import make_mock_target_bytes


class NVLReleaseManifestTests(unittest.TestCase):
    def _reports(self):
        stock = {
            "mode": "publish",
            "status": "published",
            "source": {
                "revision": '"stock-source-etag"',
            },
            "target": {
                "revision": '"target-before-stock"',
            },
            "metrics": {
                "changed_count": 2,
            },
            "post_upload_verified": True,
            "upload_result": {
                "name": "Kế hoạch mua hàng.xlsx",
                "eTag": '"target-after-stock"',
            },
        }
        open_po = {
            "publish_requested": True,
            "published": True,
            "source": {
                "revision": '"open-po-source-etag"',
            },
            "target": {
                "revision_before": '"target-before-open-po"',
                "revision_after": '"target-after-open-po"',
                "changed_cells": 3,
            },
            "publish_evidence": {
                "expected_etag": '"target-before-open-po"',
                "upload_skipped": False,
                "upload_etag": '"target-after-open-po"',
                "target_revision_after": '"target-after-open-po"',
                "post_upload_verified": True,
                "server_sha256": "server-sha",
            },
        }
        return stock, open_po

    def _readiness(self, root: Path, *, sha="commit-sha", status="passed"):
        path = root / "nvl_production_readiness_report.json"
        path.write_text(
            json.dumps(
                {
                    "status": status,
                    "gate_version": "nvl_production_readiness_v1",
                    "git_sha": sha,
                }
            ),
            encoding="utf-8",
        )
        return path

    def test_manifest_binds_commit_gate_revisions_hashes_and_publish_evidence(self):
        stock_report, open_po_report = self._reports()
        stock_bytes = make_mock_target_bytes([("VT001", 100)])
        final_bytes = make_mock_target_bytes([("VT001", 100), ("VT002", 200)])

        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            readiness = self._readiness(root)
            manifest = build_release_manifest(
                stock_report,
                open_po_report,
                stock_bytes,
                final_bytes,
                published_at="2026-09-25T08:00:00+00:00",
                environ={
                    "NVL_REQUIRE_READINESS": "1",
                    "GITHUB_SHA": "commit-sha",
                    "GITHUB_REF": "refs/heads/main",
                    "GITHUB_REPOSITORY": "Huanpro1239/Planning_vikoda",
                    "GITHUB_RUN_ID": "12345",
                    "GITHUB_RUN_ATTEMPT": "2",
                    "GITHUB_WORKFLOW": "Sync SharePoint NVL Stock",
                    "GITHUB_ACTOR": "tester",
                },
                readiness_path=readiness,
            )

        self.assertEqual(manifest["schema"], NVL_RELEASE_SCHEMA)
        self.assertEqual(manifest["commit"]["sha"], "commit-sha")
        self.assertEqual(
            manifest["readiness"]["gate_version"],
            "nvl_production_readiness_v1",
        )
        self.assertEqual(
            manifest["input_revision"]["stock_source_etag"],
            '"stock-source-etag"',
        )
        self.assertEqual(
            manifest["input_revision"]["open_po_source_etag"],
            '"open-po-source-etag"',
        )
        self.assertEqual(
            manifest["input_revision"]["open_po_target_etag_after"],
            '"target-after-open-po"',
        )
        self.assertEqual(
            manifest["artifacts"]["stock_proposal"]["sha256"],
            hashlib.sha256(stock_bytes).hexdigest(),
        )
        self.assertEqual(
            manifest["published_workbook"]["sha256"],
            hashlib.sha256(final_bytes).hexdigest(),
        )
        self.assertEqual(
            manifest["published_workbook"]["stable_sha256"],
            stable_workbook_sha256(final_bytes),
        )
        self.assertTrue(
            manifest["publish_evidence"]["stock"]["post_upload_verified"]
        )
        self.assertTrue(
            manifest["publish_evidence"]["open_po"]["evidence"][
                "post_upload_verified"
            ]
        )
        self.assertIn("run-12345.2", manifest["release_version"])

    def test_strict_manifest_rejects_readiness_commit_mismatch(self):
        stock_report, open_po_report = self._reports()
        workbook = make_mock_target_bytes([("VT001", 100)])

        with tempfile.TemporaryDirectory() as tmpdir:
            readiness = self._readiness(Path(tmpdir), sha="other-sha")
            with self.assertRaisesRegex(RuntimeError, "không khớp"):
                build_release_manifest(
                    stock_report,
                    open_po_report,
                    workbook,
                    workbook,
                    environ={
                        "NVL_REQUIRE_READINESS": "1",
                        "GITHUB_SHA": "commit-sha",
                    },
                    readiness_path=readiness,
                )

    def test_manifest_rejects_incomplete_publish(self):
        stock_report, open_po_report = self._reports()
        open_po_report["published"] = False
        workbook = make_mock_target_bytes([("VT001", 100)])

        with tempfile.TemporaryDirectory() as tmpdir:
            readiness = self._readiness(Path(tmpdir))
            with self.assertRaisesRegex(RuntimeError, "Open-PO publish"):
                build_release_manifest(
                    stock_report,
                    open_po_report,
                    workbook,
                    workbook,
                    environ={
                        "NVL_REQUIRE_READINESS": "1",
                        "GITHUB_SHA": "commit-sha",
                    },
                    readiness_path=readiness,
                )

    def test_manifest_serialization_is_deterministic(self):
        stock_report, open_po_report = self._reports()
        workbook = make_mock_target_bytes([("VT001", 100)])

        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            readiness = self._readiness(root)
            manifest = build_release_manifest(
                stock_report,
                open_po_report,
                workbook,
                workbook,
                published_at="2026-09-25T08:00:00+00:00",
                environ={
                    "NVL_REQUIRE_READINESS": "1",
                    "GITHUB_SHA": "commit-sha",
                    "GITHUB_RUN_ID": "12345",
                    "GITHUB_RUN_ATTEMPT": "1",
                },
                readiness_path=readiness,
            )
            first = root / "first.json"
            second = root / "second.json"
            write_release_manifest(manifest, path=first)
            write_release_manifest(manifest, path=second)
            self.assertEqual(first.read_bytes(), second.read_bytes())


if __name__ == "__main__":
    unittest.main()
