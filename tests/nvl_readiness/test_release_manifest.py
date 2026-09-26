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
                    "NVL_RELEASE_COMMIT_SHA": "commit-sha",
                    "GITHUB_SHA": "workflow-run-default-sha",
                    "GITHUB_REF": "refs/heads/main",
                    "GITHUB_REPOSITORY": "Huanpro1239/Planning_vikoda",
                    "GITHUB_RUN_ID": "12345",
                    "GITHUB_RUN_ATTEMPT": "2",
                    "GITHUB_WORKFLOW": "Sync SharePoint NVL Stock",
                    "GITHUB_ACTOR": "tester",
                    "NVL_UPSTREAM_PLANNING_WORKFLOW": "Sync SharePoint Stock",
                    "NVL_UPSTREAM_PLANNING_RUN_ID": "998877",
                    "NVL_UPSTREAM_PLANNING_HEAD_SHA": "commit-sha",
                    "NVL_UPSTREAM_PLANNING_EVENT": "schedule",
                },
                readiness_path=readiness,
            )

        self.assertEqual(manifest["schema"], NVL_RELEASE_SCHEMA)
        self.assertEqual(manifest["commit"]["sha"], "commit-sha")
        self.assertEqual(
            manifest["upstream_planning"],
            {
                "workflow": "Sync SharePoint Stock",
                "run_id": "998877",
                "head_sha": "commit-sha",
                "event": "schedule",
            },
        )
        self.assertEqual(
            manifest["readiness"]["gate_version"],
            "nvl_production_readiness_v1",
        )
        self.assertEqual(
            manifest["input_revision"]["stock_source_etag"],
            '"stock-source-etag"',
        )
        self.assertEqual(
            manifest["input_revision"]["stock_target_etag_after"],
            '"target-after-stock"',
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
            "server-sha",
        )
        self.assertEqual(
            manifest["published_workbook"]["proposal_sha256"],
            hashlib.sha256(final_bytes).hexdigest(),
        )
        self.assertEqual(
            manifest["published_workbook"]["stable_sha256"],
            stable_workbook_sha256(final_bytes),
        )
        self.assertTrue(
            manifest["published_workbook"]["post_upload_verified"]
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

    def test_manual_release_has_no_upstream_planning_pair(self):
        stock_report, open_po_report = self._reports()
        workbook = make_mock_target_bytes([("VT001", 100)])

        with tempfile.TemporaryDirectory() as tmpdir:
            readiness = self._readiness(Path(tmpdir))
            manifest = build_release_manifest(
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

        self.assertEqual(
            manifest["upstream_planning"],
            {
                "workflow": None,
                "run_id": None,
                "head_sha": None,
                "event": None,
            },
        )

    def test_partial_upstream_planning_provenance_is_rejected(self):
        stock_report, open_po_report = self._reports()
        workbook = make_mock_target_bytes([("VT001", 100)])

        with tempfile.TemporaryDirectory() as tmpdir:
            readiness = self._readiness(Path(tmpdir))
            with self.assertRaisesRegex(
                RuntimeError,
                "run_id/head_sha/event",
            ):
                build_release_manifest(
                    stock_report,
                    open_po_report,
                    workbook,
                    workbook,
                    environ={
                        "NVL_REQUIRE_READINESS": "1",
                        "NVL_RELEASE_COMMIT_SHA": "commit-sha",
                        "NVL_UPSTREAM_PLANNING_RUN_ID": "123",
                    },
                    readiness_path=readiness,
                )

    def test_upstream_head_sha_must_match_release_commit(self):
        stock_report, open_po_report = self._reports()
        workbook = make_mock_target_bytes([("VT001", 100)])

        with tempfile.TemporaryDirectory() as tmpdir:
            readiness = self._readiness(Path(tmpdir))
            with self.assertRaisesRegex(
                RuntimeError,
                "head_sha",
            ):
                build_release_manifest(
                    stock_report,
                    open_po_report,
                    workbook,
                    workbook,
                    environ={
                        "NVL_REQUIRE_READINESS": "1",
                        "NVL_RELEASE_COMMIT_SHA": "commit-sha",
                        "NVL_UPSTREAM_PLANNING_RUN_ID": "123",
                        "NVL_UPSTREAM_PLANNING_HEAD_SHA": "other-sha",
                        "NVL_UPSTREAM_PLANNING_EVENT": "schedule",
                    },
                    readiness_path=readiness,
                )

    def test_first_release_is_chain_anchor(self):
        stock_report, open_po_report = self._reports()
        workbook = make_mock_target_bytes([("VT001", 100)])

        with tempfile.TemporaryDirectory() as tmpdir:
            readiness = self._readiness(Path(tmpdir))
            manifest = build_release_manifest(
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

        self.assertEqual(
            manifest["chain"],
            {
                "previous_release_id": None,
                "previous_manifest_sha256": None,
            },
        )

    def test_release_links_previous_manifest_by_id_and_sha256(self):
        stock_report, open_po_report = self._reports()
        workbook = make_mock_target_bytes([("VT001", 100)])

        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            readiness = self._readiness(root)
            previous = build_release_manifest(
                stock_report,
                open_po_report,
                workbook,
                workbook,
                published_at="2026-09-25T07:00:00+00:00",
                environ={
                    "NVL_REQUIRE_READINESS": "1",
                    "GITHUB_SHA": "commit-sha",
                    "GITHUB_RUN_ID": "100",
                },
                readiness_path=readiness,
            )
            previous_path = root / "previous.json"
            write_release_manifest(previous, path=previous_path)
            previous_sha = hashlib.sha256(
                previous_path.read_bytes()
            ).hexdigest()

            current = build_release_manifest(
                stock_report,
                open_po_report,
                workbook,
                workbook,
                published_at="2026-09-25T08:00:00+00:00",
                environ={
                    "NVL_REQUIRE_READINESS": "1",
                    "GITHUB_SHA": "commit-sha",
                    "GITHUB_RUN_ID": "101",
                },
                readiness_path=readiness,
                previous_manifest=previous,
                previous_manifest_sha256=previous_sha,
            )

        self.assertEqual(
            current["chain"]["previous_release_id"],
            previous["release_id"],
        )
        self.assertEqual(
            current["chain"]["previous_manifest_sha256"],
            previous_sha,
        )

    def test_manifest_accepts_published_with_warnings(self):
        stock_report, open_po_report = self._reports()
        stock_report["status"] = "published_with_warnings"
        workbook = make_mock_target_bytes([("VT001", 100)])

        with tempfile.TemporaryDirectory() as tmpdir:
            readiness = self._readiness(Path(tmpdir))
            manifest = build_release_manifest(
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

        self.assertEqual(
            manifest["publish_evidence"]["stock"]["status"],
            "published_with_warnings",
        )

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
