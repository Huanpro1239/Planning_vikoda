"""Contracts for offline release verification."""

from __future__ import annotations

import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from planning.publish.release import build_release_manifest, write_release_manifest
from planning.publish.verify_release import (
    ReleaseVerificationError,
    verify_release,
)


ROOT = Path(__file__).resolve().parents[2]


class ReleaseVerificationTests(unittest.TestCase):
    def _bundle(self, root: Path):
        proposal = root / "planning_proposal.xlsx"
        readiness = root / "production_readiness_report.json"
        report_path = root / "planning_schedule_report.json"
        decision_path = root / "planning_publish_decision.json"
        manifest_path = root / "planning_release_manifest.json"

        proposal_bytes = b"deterministic-release-workbook"
        proposal.write_bytes(proposal_bytes)

        readiness_payload = {
            "status": "passed",
            "gate_version": "production_readiness_v1",
            "git_sha": "deadbeef",
            "phases": [{"name": "full_regression", "status": "passed"}],
        }
        readiness.write_text(
            json.dumps(readiness_payload, sort_keys=True),
            encoding="utf-8",
        )

        report = {
            "proposal_id": "placeholder",
            "plan_month": "2026-09",
            "algorithm": "algo-v1",
            "input_revision": {
                "target": {"etag": "target-etag"},
                "sources": {
                    "actual_stock": {"etag": "actual-etag"},
                },
            },
            "pipeline": {
                "engine_version": "engine-v1",
                "conversion_hash": "conversion",
                "fc_hash": "fc",
                "no_kho_hash": "debt",
                "planning_inputs_hash": "planning",
            },
        }
        decision = {
            "state": "published",
            "basis": "publish_status",
            "proposal_id": "",
            "publisher": "planner",
        }

        # First build derives the stable proposal identity contract from the
        # same fields used by production.
        from planning.publish.proposal import proposal_id

        identity, _, _ = proposal_id(report, proposal_bytes)
        report["proposal_id"] = identity
        decision["proposal_id"] = identity

        manifest = build_release_manifest(
            report,
            decision,
            proposal_bytes,
            published_at="2026-09-23T10:00:00+00:00",
            environ={
                "PLANNING_REQUIRE_READINESS": "1",
                "GITHUB_SHA": "deadbeef",
                "GITHUB_RUN_ID": "98765",
                "GITHUB_RUN_ATTEMPT": "1",
                "GITHUB_REPOSITORY": "Huanpro1239/Planning_vikoda",
                "GITHUB_REF": "refs/heads/main",
                "GITHUB_WORKFLOW": "Sync SharePoint Stock",
                "GITHUB_ACTOR": "planner",
            },
            readiness_path=readiness,
        )
        write_release_manifest(manifest, path=manifest_path)

        report_path.write_text(
            json.dumps(
                {
                    **report,
                    "proposal_output_sha256": manifest["proposal"]["stable_output_sha256"],
                    "output_sha256": manifest["proposal"]["output_sha256"],
                },
                sort_keys=True,
            ),
            encoding="utf-8",
        )
        decision_path.write_text(
            json.dumps(decision, sort_keys=True),
            encoding="utf-8",
        )
        return {
            "manifest": manifest_path,
            "proposal": proposal,
            "readiness": readiness,
            "report": report_path,
            "decision": decision_path,
            "release_id": manifest["release_id"],
            "proposal_id": identity,
        }

    def test_strict_verification_passes_with_complete_evidence(self):
        with tempfile.TemporaryDirectory() as tmp:
            bundle = self._bundle(Path(tmp))
            result = verify_release(
                bundle["manifest"],
                strict=True,
                verify_commit=False,
            )

        self.assertTrue(result["verified"])
        self.assertEqual(result["release_id"], bundle["release_id"])
        self.assertEqual(result["proposal_id"], bundle["proposal_id"])
        self.assertFalse(result["warnings"])

    def test_manifest_only_verification_warns_about_missing_external_evidence(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            bundle = self._bundle(root)
            isolated = root / "history"
            isolated.mkdir()
            historical = isolated / "release.json"
            historical.write_bytes(bundle["manifest"].read_bytes())

            result = verify_release(
                historical,
                strict=False,
                verify_commit=False,
            )

        self.assertTrue(result["verified"])
        warning_names = {
            item["name"]
            for item in result["warnings"]
        }
        self.assertEqual(
            warning_names,
            {
                "evidence.proposal",
                "evidence.readiness",
                "evidence.report",
                "evidence.decision",
            },
        )

    def test_tampered_proposal_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            bundle = self._bundle(Path(tmp))
            bundle["proposal"].write_bytes(b"tampered-workbook")
            with self.assertRaisesRegex(
                ReleaseVerificationError,
                "proposal.raw_sha256",
            ):
                verify_release(
                    bundle["manifest"],
                    strict=True,
                    verify_commit=False,
                )

    def test_tampered_readiness_evidence_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            bundle = self._bundle(Path(tmp))
            payload = json.loads(
                bundle["readiness"].read_text(encoding="utf-8")
            )
            payload["gate_version"] = "tampered-gate"
            bundle["readiness"].write_text(
                json.dumps(payload, sort_keys=True),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(
                ReleaseVerificationError,
                "readiness.report_sha256",
            ):
                verify_release(
                    bundle["manifest"],
                    strict=True,
                    verify_commit=False,
                )

    def test_tampered_audit_decision_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            bundle = self._bundle(Path(tmp))
            decision = json.loads(
                bundle["decision"].read_text(encoding="utf-8")
            )
            decision["publisher"] = "other-user"
            bundle["decision"].write_text(
                json.dumps(decision, sort_keys=True),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(
                ReleaseVerificationError,
                "audit.publish_decision",
            ):
                verify_release(
                    bundle["manifest"],
                    strict=True,
                    verify_commit=False,
                )

    def test_cli_strict_json_returns_zero_for_valid_release(self):
        with tempfile.TemporaryDirectory() as tmp:
            bundle = self._bundle(Path(tmp))
            result = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "verify_release.py"),
                    str(bundle["manifest"]),
                    "--strict",
                    "--no-git",
                    "--json",
                ],
                cwd=ROOT,
                text=True,
                capture_output=True,
                check=False,
            )

        self.assertEqual(result.returncode, 0, result.stderr)
        payload = json.loads(result.stdout)
        self.assertTrue(payload["verified"])
        self.assertEqual(payload["release_id"], bundle["release_id"])


if __name__ == "__main__":
    unittest.main()
