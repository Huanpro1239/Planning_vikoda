"""Release verification command contracts."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path
import unittest

from planning.publish.proposal import with_proposal_identity
from planning.publish.release import build_release_manifest
from planning.publish.verify_release import (
    ReleaseVerificationError,
    verify_release,
)


ROOT = Path(__file__).resolve().parents[2]


class ReleaseVerifierTests(unittest.TestCase):
    def _fixture(self, root: Path, *, commit_sha="deadbeef"):
        proposal_bytes = b"verified-release-workbook-bytes"
        base_report = {
            "plan_month": "2026-09",
            "algorithm": "contract-algorithm-v1",
            "input_revision": {
                "target": {
                    "etag": "target-etag-1",
                    "sha256": "target-sha",
                },
                "sources": {
                    "actual_stock": {
                        "etag": "actual-etag-1",
                        "sha256": "actual-sha",
                    },
                },
            },
            "pipeline": {
                "engine_version": "engine-v1",
                "conversion_hash": "conversion-hash",
                "fc_hash": "fc-hash",
                "no_kho_hash": "debt-hash",
                "planning_inputs_hash": "planning-hash",
            },
        }
        report = with_proposal_identity(
            base_report,
            proposal_bytes,
        )
        decision = {
            "state": "published",
            "basis": "publish_status",
            "publish_status": "ready_for_publish",
            "proposal_id": report["proposal_id"],
            "publisher": "release-test",
        }

        readiness_path = root / "production_readiness_report.json"
        readiness_path.write_text(
            json.dumps({
                "schema_version": 1,
                "gate_version": "production_readiness_v1",
                "status": "passed",
                "git_sha": commit_sha,
                "phases": [
                    {"name": "compile", "status": "passed"},
                    {"name": "full_regression", "status": "passed"},
                ],
            }, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

        manifest = build_release_manifest(
            report,
            decision,
            proposal_bytes,
            published_at="2026-09-23T16:00:00+00:00",
            environ={
                "PLANNING_REQUIRE_READINESS": "1",
                "GITHUB_SHA": commit_sha,
                "GITHUB_RUN_ID": "9001",
                "GITHUB_RUN_ATTEMPT": "1",
                "GITHUB_REPOSITORY": "Huanpro1239/Planning_vikoda",
                "GITHUB_REF": "refs/heads/main",
                "GITHUB_WORKFLOW": "Sync SharePoint Stock",
                "GITHUB_ACTOR": "release-test",
            },
            readiness_path=readiness_path,
        )

        manifest_path = root / "planning_release_manifest.json"
        manifest_path.write_text(
            json.dumps(
                manifest,
                ensure_ascii=False,
                indent=2,
                sort_keys=True,
            ) + "\n",
            encoding="utf-8",
        )
        proposal_path = root / "planning_proposal.xlsx"
        proposal_path.write_bytes(proposal_bytes)

        report_path = root / "planning_schedule_report.json"
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        decision_path = root / "planning_publish_decision.json"
        decision_path.write_text(
            json.dumps(decision, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

        return {
            "manifest": manifest_path,
            "proposal": proposal_path,
            "readiness": readiness_path,
            "report": report_path,
            "decision": decision_path,
            "manifest_data": manifest,
            "report_data": report,
            "decision_data": decision,
        }

    def test_strict_verification_passes_complete_consistent_evidence(self):
        with tempfile.TemporaryDirectory() as tmp:
            fixture = self._fixture(Path(tmp))
            result = verify_release(
                fixture["manifest"],
                strict=True,
                verify_commit=False,
            )

        self.assertTrue(result["verified"])
        self.assertTrue(result["strict"])
        self.assertEqual(result["warnings"], [])
        self.assertEqual(
            result["proposal_id"],
            fixture["report_data"]["proposal_id"],
        )
        self.assertTrue(
            all(item["status"] == "passed" for item in result["checks"])
        )

    def test_manifest_only_mode_passes_with_explicit_missing_evidence_warnings(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            fixture = self._fixture(root)

            fixture["proposal"].unlink()
            fixture["readiness"].unlink()
            fixture["report"].unlink()
            fixture["decision"].unlink()

            result = verify_release(
                fixture["manifest"],
                strict=False,
                verify_commit=False,
            )

        warning_names = {
            item["name"]
            for item in result["warnings"]
        }
        self.assertTrue(result["verified"])
        self.assertEqual(
            warning_names,
            {
                "evidence.proposal",
                "evidence.readiness",
                "evidence.report",
                "evidence.decision",
            },
        )

    def test_strict_verification_rejects_tampered_workbook(self):
        with tempfile.TemporaryDirectory() as tmp:
            fixture = self._fixture(Path(tmp))
            fixture["proposal"].write_bytes(b"tampered-workbook")

            with self.assertRaisesRegex(
                ReleaseVerificationError,
                "proposal.raw_sha256",
            ):
                verify_release(
                    fixture["manifest"],
                    proposal_path=fixture["proposal"],
                    readiness_path=fixture["readiness"],
                    report_path=fixture["report"],
                    decision_path=fixture["decision"],
                    strict=True,
                    verify_commit=False,
                )

    def test_strict_verification_rejects_tampered_readiness_evidence(self):
        with tempfile.TemporaryDirectory() as tmp:
            fixture = self._fixture(Path(tmp))
            readiness = json.loads(
                fixture["readiness"].read_text(encoding="utf-8")
            )
            readiness["gate_version"] = "tampered-gate"
            fixture["readiness"].write_text(
                json.dumps(readiness, indent=2) + "\n",
                encoding="utf-8",
            )

            with self.assertRaisesRegex(
                ReleaseVerificationError,
                "readiness.report_sha256",
            ):
                verify_release(
                    fixture["manifest"],
                    proposal_path=fixture["proposal"],
                    readiness_path=fixture["readiness"],
                    report_path=fixture["report"],
                    decision_path=fixture["decision"],
                    strict=True,
                    verify_commit=False,
                )

    def test_strict_verification_rejects_audit_report_mismatch(self):
        with tempfile.TemporaryDirectory() as tmp:
            fixture = self._fixture(Path(tmp))
            report = json.loads(
                fixture["report"].read_text(encoding="utf-8")
            )
            report["plan_month"] = "2026-10"
            fixture["report"].write_text(
                json.dumps(report, indent=2) + "\n",
                encoding="utf-8",
            )

            with self.assertRaisesRegex(
                ReleaseVerificationError,
                "audit.plan_month",
            ):
                verify_release(
                    fixture["manifest"],
                    proposal_path=fixture["proposal"],
                    readiness_path=fixture["readiness"],
                    report_path=fixture["report"],
                    decision_path=fixture["decision"],
                    strict=True,
                    verify_commit=False,
                )

    def test_strict_verification_rejects_publish_decision_mismatch(self):
        with tempfile.TemporaryDirectory() as tmp:
            fixture = self._fixture(Path(tmp))
            decision = json.loads(
                fixture["decision"].read_text(encoding="utf-8")
            )
            decision["publisher"] = "tampered-actor"
            fixture["decision"].write_text(
                json.dumps(decision, indent=2) + "\n",
                encoding="utf-8",
            )

            with self.assertRaisesRegex(
                ReleaseVerificationError,
                "audit.publish_decision",
            ):
                verify_release(
                    fixture["manifest"],
                    proposal_path=fixture["proposal"],
                    readiness_path=fixture["readiness"],
                    report_path=fixture["report"],
                    decision_path=fixture["decision"],
                    strict=True,
                    verify_commit=False,
                )


    def test_strict_verification_confirms_release_commit_exists_in_git_history(self):
        head = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            cwd=ROOT,
            capture_output=True,
            text=True,
            check=True,
        ).stdout.strip()

        with tempfile.TemporaryDirectory() as tmp:
            fixture = self._fixture(
                Path(tmp),
                commit_sha=head,
            )
            result = verify_release(
                fixture["manifest"],
                proposal_path=fixture["proposal"],
                readiness_path=fixture["readiness"],
                report_path=fixture["report"],
                decision_path=fixture["decision"],
                strict=True,
                verify_commit=True,
                repo_root=ROOT,
            )

        git_checks = [
            item
            for item in result["checks"]
            if item["name"] == "git.commit_exists"
        ]
        self.assertEqual(len(git_checks), 1)
        self.assertEqual(git_checks[0]["status"], "passed")

    def test_cli_returns_zero_and_json_for_valid_release(self):
        with tempfile.TemporaryDirectory() as tmp:
            fixture = self._fixture(Path(tmp))
            result = subprocess.run(
                [
                    sys.executable,
                    "-X",
                    "utf8",
                    str(ROOT / "scripts" / "verify_release.py"),
                    str(fixture["manifest"]),
                    "--strict",
                    "--no-git",
                    "--proposal",
                    str(fixture["proposal"]),
                    "--readiness",
                    str(fixture["readiness"]),
                    "--report",
                    str(fixture["report"]),
                    "--decision",
                    str(fixture["decision"]),
                    "--json",
                ],
                cwd=ROOT,
                capture_output=True,
                text=True,
                check=False,
            )

        self.assertEqual(result.returncode, 0, result.stderr)
        payload = json.loads(result.stdout)
        self.assertTrue(payload["verified"])

    def test_cli_returns_nonzero_for_tampered_release(self):
        with tempfile.TemporaryDirectory() as tmp:
            fixture = self._fixture(Path(tmp))
            manifest = json.loads(
                fixture["manifest"].read_text(encoding="utf-8")
            )
            manifest["schema"] = "unknown-schema"
            fixture["manifest"].write_text(
                json.dumps(manifest, indent=2) + "\n",
                encoding="utf-8",
            )

            result = subprocess.run(
                [
                    sys.executable,
                    "-X",
                    "utf8",
                    str(ROOT / "scripts" / "verify_release.py"),
                    str(fixture["manifest"]),
                    "--no-git",
                    "--json",
                ],
                cwd=ROOT,
                capture_output=True,
                text=True,
                check=False,
            )

        self.assertEqual(result.returncode, 1)
        payload = json.loads(result.stdout)
        self.assertFalse(payload["verified"])
        self.assertIn("manifest.schema", payload["error"])


if __name__ == "__main__":
    unittest.main()
