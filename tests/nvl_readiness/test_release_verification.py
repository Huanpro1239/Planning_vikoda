"""Contracts for offline NVL release verification."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from nvl.release import (
    build_release_manifest,
    write_release_manifest,
)
from nvl.verify_release import (
    NVLReleaseVerificationError,
    verify_nvl_release,
)


ROOT = Path(__file__).resolve().parents[2]


class NVLReleaseVerificationTests(unittest.TestCase):
    def _bundle(self, root: Path):
        stock_proposal = root / "nvl_stock_proposal.xlsx"
        final_workbook = root / "nvl_open_po_proposal.xlsx"
        readiness = root / "nvl_production_readiness_report.json"
        stock_report_path = root / "nvl_stock_report.json"
        open_po_report_path = root / "nvl_open_po_report.json"
        manifest_path = root / "nvl_release_manifest.json"

        stock_bytes = b"deterministic-stock-proposal"
        final_bytes = b"deterministic-final-workbook"
        stock_proposal.write_bytes(stock_bytes)
        final_workbook.write_bytes(final_bytes)

        readiness_payload = {
            "status": "passed",
            "gate_version": "nvl_production_readiness_v1",
            "git_sha": "deadbeef",
            "phases": [
                {
                    "name": "full_regression",
                    "status": "passed",
                }
            ],
        }
        readiness.write_text(
            json.dumps(
                readiness_payload,
                sort_keys=True,
            ),
            encoding="utf-8",
        )

        stock_report = {
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
        stock_report_path.write_text(
            json.dumps(stock_report, sort_keys=True),
            encoding="utf-8",
        )

        final_sha = hashlib.sha256(final_bytes).hexdigest()
        open_evidence = {
            "expected_etag": '"target-before-open-po"',
            "upload_skipped": False,
            "upload_etag": '"target-after-open-po"',
            "target_revision_after": '"target-after-open-po"',
            "post_upload_verified": True,
            "server_sha256": final_sha,
        }
        open_po_report = {
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
            "publish_evidence": open_evidence,
        }
        open_po_report_path.write_text(
            json.dumps(open_po_report, sort_keys=True),
            encoding="utf-8",
        )

        manifest = build_release_manifest(
            stock_report,
            open_po_report,
            stock_bytes,
            final_bytes,
            published_at="2026-09-25T09:00:00+00:00",
            environ={
                "NVL_REQUIRE_READINESS": "1",
                "GITHUB_SHA": "deadbeef",
                "GITHUB_REF": "refs/heads/main",
                "GITHUB_REPOSITORY": "Huanpro1239/Planning_vikoda",
                "GITHUB_RUN_ID": "12345",
                "GITHUB_RUN_ATTEMPT": "1",
                "GITHUB_WORKFLOW": "Sync SharePoint NVL Stock",
                "GITHUB_ACTOR": "tester",
                "NVL_RELEASE_COMMIT_SHA": "deadbeef",
                "NVL_UPSTREAM_PLANNING_WORKFLOW": "Sync SharePoint Stock",
                "NVL_UPSTREAM_PLANNING_RUN_ID": "777",
                "NVL_UPSTREAM_PLANNING_HEAD_SHA": "deadbeef",
                "NVL_UPSTREAM_PLANNING_EVENT": "schedule",
            },
            readiness_path=readiness,
        )
        write_release_manifest(
            manifest,
            path=manifest_path,
        )

        return {
            "manifest": manifest_path,
            "stock_proposal": stock_proposal,
            "final_workbook": final_workbook,
            "readiness": readiness,
            "stock_report": stock_report_path,
            "open_po_report": open_po_report_path,
        }

    def test_strict_verification_passes_complete_bundle(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            paths = self._bundle(Path(tmpdir))
            result = verify_nvl_release(
                paths["manifest"],
                strict=True,
                verify_commit=False,
            )

        self.assertTrue(result["verified"])
        self.assertTrue(result["strict"])
        self.assertEqual(result["warnings"], [])
        self.assertEqual(result["commit_sha"], "deadbeef")
        self.assertEqual(
            result["upstream_planning"]["run_id"],
            "777",
        )
        self.assertEqual(
            result["upstream_planning"]["head_sha"],
            "deadbeef",
        )
        self.assertTrue(
            result["release_id"].startswith("nvl-")
        )

    def test_tampered_upstream_planning_head_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            paths = self._bundle(Path(tmpdir))
            manifest = json.loads(
                paths["manifest"].read_text(encoding="utf-8")
            )
            manifest["upstream_planning"]["head_sha"] = "tampered-sha"
            paths["manifest"].write_text(
                json.dumps(manifest, sort_keys=True),
                encoding="utf-8",
            )

            with self.assertRaisesRegex(
                NVLReleaseVerificationError,
                "upstream_planning_head_matches_commit",
            ):
                verify_nvl_release(
                    paths["manifest"],
                    strict=True,
                    verify_commit=False,
                )

    def test_tampered_stock_proposal_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            paths = self._bundle(Path(tmpdir))
            paths["stock_proposal"].write_bytes(b"tampered-stock")

            with self.assertRaisesRegex(
                NVLReleaseVerificationError,
                "stock_proposal.raw_sha256",
            ):
                verify_nvl_release(
                    paths["manifest"],
                    strict=True,
                    verify_commit=False,
                )

    def test_tampered_final_workbook_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            paths = self._bundle(Path(tmpdir))
            paths["final_workbook"].write_bytes(b"tampered-final")

            with self.assertRaisesRegex(
                NVLReleaseVerificationError,
                "final_workbook.proposal_sha256",
            ):
                verify_nvl_release(
                    paths["manifest"],
                    strict=True,
                    verify_commit=False,
                )

    def test_tampered_readiness_report_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            paths = self._bundle(Path(tmpdir))
            readiness = json.loads(
                paths["readiness"].read_text(encoding="utf-8")
            )
            readiness["gate_version"] = "tampered-gate"
            paths["readiness"].write_text(
                json.dumps(readiness, sort_keys=True),
                encoding="utf-8",
            )

            with self.assertRaisesRegex(
                NVLReleaseVerificationError,
                "readiness.report_sha256",
            ):
                verify_nvl_release(
                    paths["manifest"],
                    strict=True,
                    verify_commit=False,
                )

    def test_tampered_stock_revision_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            paths = self._bundle(Path(tmpdir))
            report = json.loads(
                paths["stock_report"].read_text(encoding="utf-8")
            )
            report["source"]["revision"] = '"tampered-etag"'
            paths["stock_report"].write_text(
                json.dumps(report, sort_keys=True),
                encoding="utf-8",
            )

            with self.assertRaisesRegex(
                NVLReleaseVerificationError,
                "audit.stock.source_revision",
            ):
                verify_nvl_release(
                    paths["manifest"],
                    strict=True,
                    verify_commit=False,
                )

    def test_tampered_open_po_evidence_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            paths = self._bundle(Path(tmpdir))
            report = json.loads(
                paths["open_po_report"].read_text(encoding="utf-8")
            )
            report["publish_evidence"]["server_sha256"] = "tampered"
            paths["open_po_report"].write_text(
                json.dumps(report, sort_keys=True),
                encoding="utf-8",
            )

            with self.assertRaisesRegex(
                NVLReleaseVerificationError,
                "audit.open_po.publish_evidence",
            ):
                verify_nvl_release(
                    paths["manifest"],
                    strict=True,
                    verify_commit=False,
                )

    def test_tampered_manifest_release_identity_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            paths = self._bundle(Path(tmpdir))
            manifest = json.loads(
                paths["manifest"].read_text(encoding="utf-8")
            )
            manifest["release_id"] = "nvl-wrong-identity"
            paths["manifest"].write_text(
                json.dumps(manifest, sort_keys=True),
                encoding="utf-8",
            )

            with self.assertRaisesRegex(
                NVLReleaseVerificationError,
                "manifest.release_id_identity",
            ):
                verify_nvl_release(
                    paths["manifest"],
                    strict=True,
                    verify_commit=False,
                )

    def test_missing_evidence_warns_non_strict_and_fails_strict(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            paths = self._bundle(Path(tmpdir))
            paths["stock_proposal"].unlink()

            relaxed = verify_nvl_release(
                paths["manifest"],
                strict=False,
                verify_commit=False,
            )
            warning_names = {
                item["name"]
                for item in relaxed["warnings"]
            }
            self.assertIn(
                "evidence.stock_proposal",
                warning_names,
            )

            with self.assertRaisesRegex(
                NVLReleaseVerificationError,
                "evidence.stock_proposal",
            ):
                verify_nvl_release(
                    paths["manifest"],
                    strict=True,
                    verify_commit=False,
                )

    def test_cli_json_output_passes_complete_bundle(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            paths = self._bundle(Path(tmpdir))
            result = subprocess.run(
                [
                    sys.executable,
                    "-X",
                    "utf8",
                    str(ROOT / "scripts" / "verify_nvl_release.py"),
                    str(paths["manifest"]),
                    "--strict",
                    "--no-git",
                    "--json",
                ],
                cwd=ROOT,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=False,
            )

        self.assertEqual(
            result.returncode,
            0,
            msg=result.stderr or result.stdout,
        )
        payload = json.loads(result.stdout)
        self.assertTrue(payload["verified"])
        self.assertEqual(payload["commit_sha"], "deadbeef")


if __name__ == "__main__":
    unittest.main()
