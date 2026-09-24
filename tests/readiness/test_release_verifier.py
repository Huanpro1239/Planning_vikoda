"""Release verification command contracts."""

from __future__ import annotations

import hashlib
import json
import tempfile
from pathlib import Path
import unittest
from unittest.mock import patch

from planning.publish.proposal import proposal_output_sha256
from planning.publish.release import RELEASE_SCHEMA, RELEASE_SCHEMA_VERSION
from planning.publish.verify_release import (
    ReleaseVerificationError,
    verify_release,
)


class ReleaseVerificationTests(unittest.TestCase):
    def _write_fixture(self, root: Path):
        proposal_bytes = b"not-an-xlsx-but-stable-fallback"
        stable = proposal_output_sha256(proposal_bytes)
        input_revision = {
            "target": {"etag": "target-1"},
            "sources": {"actual_stock": {"etag": "source-1"}},
        }
        payload = {
            "algorithm": "algo-v1",
            "plan_month": "2026-09",
            "input_revision": input_revision,
            "output_sha256": stable,
        }
        canonical = json.dumps(
            payload,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
        proposal_id = hashlib.sha256(canonical).hexdigest()

        readiness = {
            "status": "passed",
            "gate_version": "production_readiness_v1",
            "git_sha": "deadbeef",
        }
        readiness_path = root / "production_readiness_report.json"
        readiness_path.write_text(
            json.dumps(readiness, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

        decision = {
            "state": "published",
            "basis": "publish_status",
            "proposal_id": proposal_id,
            "publisher": "planner",
        }
        report = {
            "proposal_id": proposal_id,
            "plan_month": "2026-09",
            "algorithm": "algo-v1",
            "input_revision": input_revision,
        }
        manifest = {
            "schema": RELEASE_SCHEMA,
            "schema_version": RELEASE_SCHEMA_VERSION,
            "release_id": f"2026-09-{proposal_id[:12]}-gh-123-1",
            "release_version": (
                f"{RELEASE_SCHEMA}.proposal-{proposal_id[:12]}.run-123.1"
            ),
            "published_at": "2026-09-23T10:00:00+00:00",
            "plan_month": "2026-09",
            "algorithm": "algo-v1",
            "engine_version": "engine-v1",
            "commit": {
                "sha": "deadbeef",
                "ref": "refs/heads/main",
                "repository": "Huanpro1239/Planning_vikoda",
                "run_id": "123",
                "run_attempt": "1",
                "workflow": "Sync SharePoint Stock",
                "actor": "planner",
            },
            "readiness": {
                "status": "passed",
                "gate_version": "production_readiness_v1",
                "git_sha": "deadbeef",
                "report_sha256": hashlib.sha256(
                    readiness_path.read_bytes()
                ).hexdigest(),
            },
            "proposal": {
                "proposal_id": proposal_id,
                "output_sha256": hashlib.sha256(proposal_bytes).hexdigest(),
                "stable_output_sha256": stable,
            },
            "input_revision": input_revision,
            "publish_decision": decision,
            "pipeline_fingerprints": {},
        }

        (root / "planning_proposal.xlsx").write_bytes(proposal_bytes)
        (root / "planning_schedule_report.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        (root / "planning_publish_decision.json").write_text(
            json.dumps(decision, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        manifest_path = root / "planning_release_manifest.json"
        manifest_path.write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        return manifest_path, proposal_id

    def test_strict_full_evidence_verifies_release(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            manifest_path, proposal_id = self._write_fixture(root)
            with patch(
                "planning.publish.verify_release._git_commit_exists",
                return_value=(True, "commit exists"),
            ):
                result = verify_release(
                    manifest_path,
                    strict=True,
                    repo_root=root,
                )

        self.assertTrue(result["verified"])
        self.assertEqual(result["proposal_id"], proposal_id)
        self.assertEqual(result["warnings"], [])
        self.assertTrue(
            all(item["status"] == "passed" for item in result["checks"])
        )

    def test_manifest_only_non_strict_returns_warnings_for_missing_evidence(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            manifest_path, _ = self._write_fixture(root)
            for name in (
                "planning_proposal.xlsx",
                "production_readiness_report.json",
                "planning_schedule_report.json",
                "planning_publish_decision.json",
            ):
                (root / name).unlink()

            result = verify_release(
                manifest_path,
                strict=False,
                verify_commit=False,
            )

        self.assertTrue(result["verified"])
        warning_names = {item["name"] for item in result["warnings"]}
        self.assertIn("evidence.proposal", warning_names)
        self.assertIn("evidence.readiness", warning_names)
        self.assertIn("evidence.report", warning_names)
        self.assertIn("evidence.decision", warning_names)

    def test_strict_rejects_tampered_proposal(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            manifest_path, _ = self._write_fixture(root)
            (root / "planning_proposal.xlsx").write_bytes(b"tampered")

            with patch(
                "planning.publish.verify_release._git_commit_exists",
                return_value=(True, "commit exists"),
            ):
                with self.assertRaisesRegex(
                    ReleaseVerificationError,
                    "proposal.raw_sha256",
                ):
                    verify_release(
                        manifest_path,
                        strict=True,
                        repo_root=root,
                    )

    def test_rejects_manifest_schema_mismatch(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            manifest_path, _ = self._write_fixture(root)
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["schema"] = "unknown"
            manifest_path.write_text(
                json.dumps(manifest),
                encoding="utf-8",
            )

            with self.assertRaisesRegex(
                ReleaseVerificationError,
                "manifest.schema",
            ):
                verify_release(
                    manifest_path,
                    strict=False,
                    verify_commit=False,
                )


if __name__ == "__main__":
    unittest.main()
