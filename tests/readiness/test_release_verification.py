"""Historical release verification contracts."""

from __future__ import annotations

import copy
import json
import subprocess
import tempfile
from pathlib import Path
import unittest

from planning.publish.proposal import with_proposal_identity
from planning.publish.release import (
    RELEASE_SCHEMA,
    RELEASE_SCHEMA_VERSION,
    build_release_manifest,
    write_release_manifest,
)
from planning.publish.release_verify import verify_release


ROOT = Path(__file__).resolve().parents[2]


class ReleaseVerificationTests(unittest.TestCase):
    def _fixture(self, root: Path):
        proposal_bytes = b"release-proposal-workbook"
        base_report = {
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
                "conversion_hash": "conv",
                "fc_hash": "fc",
                "no_kho_hash": "debt",
                "planning_inputs_hash": "planning",
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
            "publisher": "planner",
            "target_etag": "target-etag",
        }
        report["publish_decision"] = copy.deepcopy(decision)

        head = subprocess.check_output(
            ["git", "-C", str(ROOT), "rev-parse", "HEAD"],
            text=True,
        ).strip()

        readiness = {
            "status": "passed",
            "gate_version": "production_readiness_v1",
            "git_sha": head,
            "phases": [{"name": "full_regression", "status": "passed"}],
        }

        evidence = root / "release_evidence" / "release-test"
        evidence.mkdir(parents=True)
        readiness_path = evidence / "production_readiness_report.json"
        report_path = evidence / "planning_schedule_report.json"
        revision_path = evidence / "planning_input_revision.json"
        decision_path = evidence / "planning_publish_decision.json"
        proposal_path = evidence / "planning_proposal.xlsx"

        readiness_path.write_text(
            json.dumps(readiness, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )

        manifest = build_release_manifest(
            report,
            decision,
            proposal_bytes,
            published_at="2026-09-23T10:00:00+00:00",
            environ={
                "PLANNING_REQUIRE_READINESS": "1",
                "GITHUB_SHA": head,
                "GITHUB_RUN_ID": "12345",
                "GITHUB_RUN_ATTEMPT": "1",
                "GITHUB_REPOSITORY": "Huanpro1239/Planning_vikoda",
                "GITHUB_REF": "refs/heads/main",
                "GITHUB_WORKFLOW": "Sync SharePoint Stock",
                "GITHUB_ACTOR": "planner",
            },
            readiness_path=readiness_path,
        )

        release_id = manifest["release_id"]
        manifest_path = root / "releases" / f"{release_id}.json"
        manifest_path.parent.mkdir(parents=True)
        write_release_manifest(manifest, path=manifest_path)

        actual_evidence = root / "release_evidence" / release_id
        evidence.rename(actual_evidence)
        report_path = actual_evidence / "planning_schedule_report.json"
        revision_path = actual_evidence / "planning_input_revision.json"
        decision_path = actual_evidence / "planning_publish_decision.json"
        proposal_path = actual_evidence / "planning_proposal.xlsx"

        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )
        revision_path.write_text(
            json.dumps(report["input_revision"], ensure_ascii=False, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )
        decision_path.write_text(
            json.dumps(decision, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )
        proposal_path.write_bytes(proposal_bytes)

        return {
            "manifest": manifest,
            "manifest_path": manifest_path,
            "proposal": proposal_path,
            "readiness": actual_evidence / "production_readiness_report.json",
            "report": report_path,
            "revision": revision_path,
            "decision": decision_path,
        }

    def test_v2_release_verifies_all_evidence_in_strict_mode(self):
        with tempfile.TemporaryDirectory() as tmp:
            fixture = self._fixture(Path(tmp))
            result = verify_release(
                fixture["manifest_path"],
                proposal_path=fixture["proposal"],
                repo_root=ROOT,
                strict=True,
            )

        self.assertEqual(result["verification_status"], "verified")
        self.assertEqual(result["summary"]["failed"], 0)
        self.assertEqual(result["summary"]["skipped"], 0)

    def test_tampered_proposal_is_detected(self):
        with tempfile.TemporaryDirectory() as tmp:
            fixture = self._fixture(Path(tmp))
            fixture["proposal"].write_bytes(b"tampered-workbook")
            result = verify_release(
                fixture["manifest_path"],
                proposal_path=fixture["proposal"],
                repo_root=ROOT,
                strict=True,
            )

        self.assertEqual(result["verification_status"], "failed")
        failed = {
            item["name"]
            for item in result["checks"]
            if item["status"] == "failed"
        }
        self.assertIn("proposal_raw_hash", failed)
        self.assertIn("proposal_id_recomputed", failed)

    def test_tampered_audit_decision_is_detected(self):
        with tempfile.TemporaryDirectory() as tmp:
            fixture = self._fixture(Path(tmp))
            decision = json.loads(
                fixture["decision"].read_text(encoding="utf-8")
            )
            decision["publisher"] = "tampered"
            fixture["decision"].write_text(
                json.dumps(decision, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
                encoding="utf-8",
            )
            result = verify_release(
                fixture["manifest_path"],
                proposal_path=fixture["proposal"],
                repo_root=ROOT,
                strict=True,
            )

        self.assertEqual(result["verification_status"], "failed")
        failed = {
            item["name"]
            for item in result["checks"]
            if item["status"] == "failed"
        }
        self.assertIn("publish_decision_consistency", failed)
        self.assertIn("publish_decision_hash", failed)

    def test_manifest_only_is_partial_not_false_pass(self):
        with tempfile.TemporaryDirectory() as tmp:
            fixture = self._fixture(Path(tmp))
            # Copy only the manifest away from its auto-discoverable evidence.
            isolated = Path(tmp) / "isolated.json"
            isolated.write_bytes(fixture["manifest_path"].read_bytes())
            result = verify_release(
                isolated,
                repo_root=ROOT,
                strict=False,
            )

        self.assertEqual(result["verification_status"], "partial")
        self.assertGreater(result["summary"]["skipped"], 0)
        self.assertEqual(result["summary"]["failed"], 0)

    def test_v1_manifest_remains_supported(self):
        with tempfile.TemporaryDirectory() as tmp:
            fixture = self._fixture(Path(tmp))
            manifest = copy.deepcopy(fixture["manifest"])
            proposal_id = manifest["proposal"]["proposal_id"]
            manifest["schema"] = "planning_release_manifest_v1"
            manifest["schema_version"] = 1
            manifest["release_version"] = (
                f"planning_release_manifest_v1.proposal-{proposal_id[:12]}.run-12345.1"
            )
            manifest.pop("audit_evidence", None)
            write_release_manifest(
                manifest,
                path=fixture["manifest_path"],
            )

            result = verify_release(
                fixture["manifest_path"],
                proposal_path=fixture["proposal"],
                repo_root=ROOT,
                strict=True,
            )

        self.assertEqual(result["verification_status"], "verified")
        self.assertEqual(result["release_schema"], "planning_release_manifest_v1")


if __name__ == "__main__":
    unittest.main()
