"""Release manifest/versioning contracts."""

from __future__ import annotations

import json
import tempfile
from pathlib import Path
import unittest

from planning.publish.release import (
    RELEASE_SCHEMA,
    RELEASE_SCHEMA_VERSION,
    build_release_manifest,
)


class ReleaseManifestTests(unittest.TestCase):
    def _report(self):
        return {
            "proposal_id": "abcdef1234567890",
            "plan_month": "2026-09",
            "algorithm": "algo-v1",
            "input_revision": {
                "target": {"etag": "target-etag"},
                "sources": {"actual_stock": {"etag": "source-etag"}},
            },
            "pipeline": {
                "engine_version": "engine-v1",
                "conversion_hash": "conv",
                "fc_hash": "fc",
                "no_kho_hash": "debt",
                "planning_inputs_hash": "planning",
            },
        }

    def _decision(self):
        return {
            "state": "published",
            "basis": "publish_status",
            "proposal_id": "abcdef1234567890",
            "publisher": "planner",
        }

    def test_manifest_contains_required_traceability_fields(self):
        with tempfile.TemporaryDirectory() as tmp:
            readiness = Path(tmp) / "readiness.json"
            readiness.write_text(
                json.dumps({
                    "status": "passed",
                    "gate_version": "production_readiness_v1",
                    "git_sha": "deadbeef",
                }),
                encoding="utf-8",
            )
            manifest = build_release_manifest(
                self._report(),
                self._decision(),
                b"workbook-bytes",
                published_at="2026-09-23T10:00:00+00:00",
                environ={
                    "PLANNING_REQUIRE_READINESS": "1",
                    "GITHUB_SHA": "deadbeef",
                    "GITHUB_RUN_ID": "12345",
                    "GITHUB_RUN_ATTEMPT": "2",
                    "GITHUB_REPOSITORY": "Huanpro1239/Planning_vikoda",
                    "GITHUB_REF": "refs/heads/main",
                    "GITHUB_WORKFLOW": "Sync SharePoint Stock",
                    "GITHUB_ACTOR": "planner",
                },
                readiness_path=readiness,
            )

        self.assertEqual(manifest["schema"], RELEASE_SCHEMA)
        self.assertEqual(manifest["schema_version"], RELEASE_SCHEMA_VERSION)
        self.assertEqual(manifest["commit"]["sha"], "deadbeef")
        self.assertEqual(manifest["readiness"]["status"], "passed")
        self.assertEqual(
            manifest["readiness"]["gate_version"],
            "production_readiness_v1",
        )
        self.assertEqual(
            manifest["proposal"]["proposal_id"],
            "abcdef1234567890",
        )
        self.assertTrue(manifest["proposal"]["output_sha256"])
        self.assertTrue(manifest["proposal"]["stable_output_sha256"])
        self.assertEqual(
            manifest["input_revision"]["target"]["etag"],
            "target-etag",
        )
        self.assertIn("12345", manifest["release_id"])
        self.assertIn("proposal-abcdef123456", manifest["release_version"])

    def test_strict_production_manifest_requires_passed_readiness(self):
        with tempfile.TemporaryDirectory() as tmp:
            readiness = Path(tmp) / "readiness.json"
            readiness.write_text(
                json.dumps({
                    "status": "failed",
                    "gate_version": "production_readiness_v1",
                    "git_sha": "deadbeef",
                }),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(RuntimeError, "chưa PASS"):
                build_release_manifest(
                    self._report(),
                    self._decision(),
                    b"workbook",
                    environ={
                        "PLANNING_REQUIRE_READINESS": "1",
                        "GITHUB_SHA": "deadbeef",
                    },
                    readiness_path=readiness,
                )

    def test_manifest_rejects_non_publish_decision(self):
        decision = dict(self._decision(), state="blocked")
        with self.assertRaisesRegex(RuntimeError, "publish thành công"):
            build_release_manifest(
                self._report(),
                decision,
                b"workbook",
                environ={},
            )


if __name__ == "__main__":
    unittest.main()
