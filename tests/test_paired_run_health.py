"""Contracts for Planning -> NVL paired-run health monitoring."""

from __future__ import annotations

from datetime import datetime, timezone
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from ops.paired_run_health import (
    evaluate_paired_run_health,
    parse_paired_nvl_run,
)


ROOT = Path(__file__).resolve().parents[1]
NOW = datetime(2026, 9, 26, 4, 0, tzinfo=timezone.utc)
CUTOFF = "2026-09-26T03:30:00Z"


def planning_run(
    run_id: int = 100,
    *,
    sha: str = "a" * 40,
    event: str = "schedule",
    conclusion: str = "success",
    created_at: str = "2026-09-26T03:31:00Z",
    updated_at: str = "2026-09-26T03:32:00Z",
):
    return {
        "id": run_id,
        "name": "Sync SharePoint Stock",
        "event": event,
        "status": "completed",
        "conclusion": conclusion,
        "head_sha": sha,
        "created_at": created_at,
        "updated_at": updated_at,
        "html_url": f"https://github.example/planning/{run_id}",
    }


def nvl_run(
    nvl_id: int = 200,
    *,
    planning_id: int = 100,
    sha: str = "a" * 40,
    planning_event: str = "schedule",
    status: str = "completed",
    conclusion: str | None = "success",
):
    return {
        "id": nvl_id,
        "name": "Sync SharePoint NVL Stock",
        "event": "workflow_run",
        "status": status,
        "conclusion": conclusion,
        "display_title": (
            f"NVL paired planning={planning_id} "
            f"sha={sha} event={planning_event}"
        ),
        "created_at": "2026-09-26T03:33:00Z",
        "updated_at": "2026-09-26T03:36:00Z",
        "html_url": f"https://github.example/nvl/{nvl_id}",
    }


def release_pair(
    *,
    planning_id: int = 100,
    sha: str = "a" * 40,
    event: str = "schedule",
    release_id: str = "nvl-release-1",
):
    return {
        "release_id": release_id,
        "planning_run_id": str(planning_id),
        "planning_head_sha": sha,
        "planning_event": event,
        "planning_workflow": "Sync SharePoint Stock",
        "release_commit_sha": sha,
        "published_at": "2026-09-26T03:35:00Z",
    }


class PairedRunHealthTests(unittest.TestCase):
    def _evaluate(self, planning=None, nvl=None, releases=None, **kwargs):
        return evaluate_paired_run_health(
            planning if planning is not None else [planning_run()],
            nvl if nvl is not None else [nvl_run()],
            releases if releases is not None else [release_pair()],
            enforce_after=CUTOFF,
            grace_minutes=20,
            now=NOW,
            **kwargs,
        )

    def test_parse_paired_run_name(self):
        parsed = parse_paired_nvl_run(nvl_run())
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed["planning_run_id"], "100")
        self.assertEqual(parsed["planning_head_sha"], "a" * 40)
        self.assertEqual(parsed["planning_event"], "schedule")

    def test_healthy_one_to_one_pair_passes(self):
        health = self._evaluate()
        self.assertEqual(health["status"], "passed")
        self.assertEqual(health["summary"]["planning_runs_checked"], 1)
        self.assertEqual(health["summary"]["passed_pairs"], 1)
        self.assertEqual(health["summary"]["failed_pairs"], 0)

    def test_missing_nvl_is_detected(self):
        health = self._evaluate(nvl=[], releases=[])
        self.assertEqual(health["status"], "failed")
        self.assertEqual(health["summary"]["missing_nvl"], 1)
        codes = {
            issue["code"]
            for issue in health["pairs"][0]["issues"]
        }
        self.assertIn("missing_nvl", codes)

    def test_duplicate_nvl_is_detected(self):
        health = self._evaluate(
            nvl=[
                nvl_run(200),
                nvl_run(201),
            ]
        )
        self.assertEqual(health["status"], "failed")
        self.assertEqual(health["summary"]["duplicate_nvl"], 1)
        self.assertEqual(health["pairs"][0]["nvl_run_count"], 2)

    def test_head_sha_mismatch_is_detected(self):
        health = self._evaluate(
            nvl=[nvl_run(sha="b" * 40)],
        )
        self.assertEqual(health["status"], "failed")
        self.assertGreaterEqual(
            health["summary"]["head_sha_mismatch"],
            1,
        )
        codes = {
            issue["code"]
            for issue in health["pairs"][0]["issues"]
        }
        self.assertIn("head_sha_mismatch", codes)

    def test_release_head_sha_mismatch_is_detected(self):
        health = self._evaluate(
            releases=[release_pair(sha="b" * 40)]
        )
        self.assertEqual(health["status"], "failed")
        self.assertGreaterEqual(
            health["summary"]["head_sha_mismatch"],
            1,
        )
        codes = {
            issue["code"]
            for issue in health["pairs"][0]["issues"]
        }
        self.assertIn("release_head_sha_mismatch", codes)
        self.assertIn("release_commit_sha_mismatch", codes)

    def test_downstream_failure_is_detected(self):
        health = self._evaluate(
            nvl=[nvl_run(conclusion="failure")],
            releases=[],
        )
        self.assertEqual(health["status"], "failed")
        self.assertEqual(
            health["summary"]["nvl_downstream_failed"],
            1,
        )
        codes = {
            issue["code"]
            for issue in health["pairs"][0]["issues"]
        }
        self.assertIn("nvl_downstream_failed", codes)
        self.assertNotIn("missing_release_provenance", codes)

    def test_successful_downstream_without_release_is_detected(self):
        health = self._evaluate(releases=[])
        self.assertEqual(health["status"], "failed")
        self.assertEqual(
            health["summary"]["missing_release_provenance"],
            1,
        )

    def test_duplicate_release_provenance_is_detected(self):
        health = self._evaluate(
            releases=[
                release_pair(release_id="nvl-release-1"),
                release_pair(release_id="nvl-release-2"),
            ]
        )
        self.assertEqual(health["status"], "failed")
        codes = {
            issue["code"]
            for issue in health["pairs"][0]["issues"]
        }
        self.assertIn("duplicate_release_provenance", codes)

    def test_manual_or_direct_nvl_runs_are_ignored(self):
        direct = {
            "id": 300,
            "display_title": (
                "NVL direct event=workflow_dispatch sha=" + "a" * 40
            ),
            "status": "completed",
            "conclusion": "success",
        }
        health = self._evaluate(
            nvl=[direct, nvl_run()],
        )
        self.assertEqual(health["status"], "passed")
        self.assertEqual(health["pairs"][0]["nvl_run_count"], 1)

    def test_recent_planning_run_inside_grace_is_not_failed(self):
        recent = planning_run(
            created_at="2026-09-26T03:50:00Z",
            updated_at="2026-09-26T03:55:00Z",
        )
        health = self._evaluate(
            planning=[recent],
            nvl=[],
            releases=[],
        )
        self.assertEqual(health["status"], "empty")
        self.assertEqual(health["summary"]["planning_runs_checked"], 0)

    def test_pre_cutover_planning_run_is_not_retroactively_failed(self):
        old = planning_run(
            created_at="2026-09-26T03:20:00Z",
            updated_at="2026-09-26T03:22:00Z",
        )
        health = self._evaluate(
            planning=[old],
            nvl=[],
            releases=[],
        )
        self.assertEqual(health["status"], "empty")

    def test_failed_planning_run_is_not_expected_to_have_nvl(self):
        failed = planning_run(conclusion="failure")
        health = self._evaluate(
            planning=[failed],
            nvl=[],
            releases=[],
        )
        self.assertEqual(health["status"], "empty")

    def test_cli_offline_fixtures_write_health_artifacts(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            planning_path = root / "planning.json"
            nvl_path = root / "nvl.json"
            runtime = root / "runtime-state"
            release_dir = runtime / "nvl" / "releases"
            release_dir.mkdir(parents=True)

            planning_fixture = planning_run()
            planning_fixture["created_at"] = "2026-09-26T01:31:00Z"
            planning_fixture["updated_at"] = "2026-09-26T01:32:00Z"
            nvl_fixture = nvl_run()
            nvl_fixture["created_at"] = "2026-09-26T01:33:00Z"
            nvl_fixture["updated_at"] = "2026-09-26T01:36:00Z"

            planning_path.write_text(
                json.dumps([planning_fixture]),
                encoding="utf-8",
            )
            nvl_path.write_text(
                json.dumps([nvl_fixture]),
                encoding="utf-8",
            )
            (release_dir / "nvl-release-1.json").write_text(
                json.dumps(
                    {
                        "release_id": "nvl-release-1",
                        "published_at": "2026-09-26T03:35:00Z",
                        "upstream_planning": {
                            "workflow": "Sync SharePoint Stock",
                            "run_id": "100",
                            "head_sha": "a" * 40,
                            "event": "schedule",
                        },
                        "commit": {"sha": "a" * 40},
                    }
                ),
                encoding="utf-8",
            )

            out_json = root / "health.json"
            out_csv = root / "health.csv"
            out_md = root / "health.md"

            result = subprocess.run(
                [
                    sys.executable,
                    "-X",
                    "utf8",
                    str(ROOT / "scripts" / "check_paired_runs.py"),
                    "--runtime-state",
                    str(runtime),
                    "--planning-runs-file",
                    str(planning_path),
                    "--nvl-runs-file",
                    str(nvl_path),
                    "--enforce-after",
                    "2026-09-26T00:00:00Z",
                    "--grace-minutes",
                    "0",
                    "--out-json",
                    str(out_json),
                    "--out-csv",
                    str(out_csv),
                    "--out-md",
                    str(out_md),
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
            self.assertTrue(out_json.is_file())
            self.assertTrue(out_csv.is_file())
            self.assertTrue(out_md.is_file())
            payload = json.loads(out_json.read_text(encoding="utf-8"))
            self.assertEqual(payload["status"], "passed")


class PairedRunHealthWorkflowTests(unittest.TestCase):
    def test_nvl_run_name_encodes_upstream_pair_identity(self):
        text = (
            ROOT / ".github" / "workflows" / "sync-nvl-stock.yml"
        ).read_text(encoding="utf-8")
        self.assertIn("run-name:", text)
        self.assertIn("NVL paired planning={0} sha={1} event={2}", text)
        self.assertIn("github.event.workflow_run.id", text)
        self.assertIn("github.event.workflow_run.head_sha", text)

    def test_health_workflow_has_event_driven_and_daily_safety_net(self):
        text = (
            ROOT / ".github" / "workflows" / "paired-run-health.yml"
        ).read_text(encoding="utf-8")
        self.assertIn("- Sync SharePoint NVL Stock", text)
        self.assertIn('cron: "45 23 * * *"', text)
        self.assertIn("actions: read", text)
        self.assertIn("contents: read", text)
        self.assertIn("ref: runtime-state", text)
        self.assertIn("scripts/check_paired_runs.py", text)
        self.assertIn("--grace-minutes 20", text)
        self.assertIn(
            '--enforce-after "2026-09-26T03:30:00Z"',
            text,
        )

    def test_health_artifacts_are_uploaded_before_final_failure(self):
        text = (
            ROOT / ".github" / "workflows" / "paired-run-health.yml"
        ).read_text(encoding="utf-8")
        check_index = text.find(
            "- name: Check Planning to NVL paired-run health"
        )
        upload_index = text.find(
            "- name: Upload paired-run health evidence"
        )
        fail_index = text.find(
            "- name: Fail when paired-run health is unhealthy"
        )
        self.assertGreater(check_index, 0)
        self.assertGreater(upload_index, check_index)
        self.assertGreater(fail_index, upload_index)
        self.assertIn(
            "continue-on-error: true",
            text[check_index:upload_index],
        )
        self.assertIn("if: always()", text[upload_index:fail_index])


if __name__ == "__main__":
    unittest.main()
