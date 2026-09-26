"""Regression guards for the NVL production-readiness gate wiring."""

from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]
GATE_COMMAND = "python -X utf8 scripts/nvl_production_readiness.py"
REPORT_NAME = "nvl_production_readiness_report.json"


class NVLReadinessGateWiringTests(unittest.TestCase):
    def test_workflow_uses_single_nvl_readiness_command(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-nvl-stock.yml"
        ).read_text(encoding="utf-8")
        self.assertIn(GATE_COMMAND, workflow)
        self.assertNotIn(
            "python -X utf8 -m unittest discover -s tests -v",
            workflow,
        )

    def test_gate_runs_before_runtime_state_and_any_nvl_sharepoint_operation(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-nvl-stock.yml"
        ).read_text(encoding="utf-8")
        gate_index = workflow.find("- name: Run NVL production readiness gate")
        runtime_index = workflow.find("- name: Checkout runtime state")
        staging_index = workflow.find("- name: Ensure staging copy on SharePoint")
        publish_index = workflow.find("- name: Generate NVL proposal or controlled publish")
        self.assertGreaterEqual(gate_index, 0)
        self.assertGreater(runtime_index, gate_index)
        self.assertGreater(staging_index, gate_index)
        self.assertGreater(publish_index, gate_index)

    def test_nvl_automatic_trigger_follows_successful_planning_run(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-nvl-stock.yml"
        ).read_text(encoding="utf-8")
        self.assertIn("workflow_run:", workflow)
        self.assertIn("- Sync SharePoint Stock", workflow)
        self.assertIn("- completed", workflow)
        self.assertNotIn('cron: "15 23 * * *"', workflow)
        self.assertIn(
            "github.event.workflow_run.conclusion == 'success'",
            workflow,
        )
        self.assertIn(
            "github.event.workflow_run.event == 'schedule'",
            workflow,
        )
        self.assertIn(
            "github.event.workflow_run.event == 'repository_dispatch'",
            workflow,
        )
        self.assertIn(
            "github.event.workflow_run.head_sha || github.sha",
            workflow,
        )

    def test_manual_planning_run_does_not_auto_publish_nvl(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-nvl-stock.yml"
        ).read_text(encoding="utf-8")
        guard = workflow[
            workflow.find("jobs:"):
            workflow.find("runs-on: ubuntu-latest")
        ]
        self.assertNotIn(
            "github.event.workflow_run.event == 'workflow_dispatch'",
            guard,
        )
        self.assertIn(
            "github.event_name == 'workflow_dispatch' && inputs.publish",
            workflow,
        )

    def test_workflow_run_is_treated_as_authorized_production_trigger(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-nvl-stock.yml"
        ).read_text(encoding="utf-8")
        self.assertGreaterEqual(
            workflow.count("github.event_name == 'workflow_run'"),
            5,
        )
        self.assertIn("UPSTREAM_RUN_ID:", workflow)
        self.assertIn("UPSTREAM_EVENT:", workflow)

    def test_workflow_passes_paired_planning_provenance_to_release(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-nvl-stock.yml"
        ).read_text(encoding="utf-8")
        self.assertIn(
            "NVL_RELEASE_COMMIT_SHA: ${{ github.event_name == 'workflow_run' && github.event.workflow_run.head_sha || github.sha }}",
            workflow,
        )
        self.assertIn(
            "NVL_UPSTREAM_PLANNING_RUN_ID: ${{ github.event_name == 'workflow_run' && github.event.workflow_run.id || '' }}",
            workflow,
        )
        self.assertIn(
            "NVL_UPSTREAM_PLANNING_HEAD_SHA: ${{ github.event_name == 'workflow_run' && github.event.workflow_run.head_sha || '' }}",
            workflow,
        )
        self.assertIn(
            "NVL_UPSTREAM_PLANNING_EVENT: ${{ github.event_name == 'workflow_run' && github.event.workflow_run.event || '' }}",
            workflow,
        )
        self.assertIn(
            "NVL_UPSTREAM_PLANNING_WORKFLOW: ${{ github.event_name == 'workflow_run' && github.event.workflow_run.name || '' }}",
            workflow,
        )

    def test_release_manifest_is_recorded_after_publish_and_persisted_state_only(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-nvl-stock.yml"
        ).read_text(encoding="utf-8")
        publish_index = workflow.find("- name: Generate NVL proposal or controlled publish")
        refresh_index = workflow.find("- name: Refresh NVL release history")
        manifest_index = workflow.find("- name: Record NVL release manifest")
        history_index = workflow.find("- name: Save NVL release history to runtime-state branch")
        self.assertGreater(refresh_index, publish_index)
        self.assertGreater(manifest_index, refresh_index)
        self.assertGreater(history_index, manifest_index)
        self.assertIn("python -X utf8 -m nvl.release", workflow)
        self.assertIn(
            "--previous-manifest .runtime-state/nvl/latest_release.json",
            workflow,
        )
        self.assertIn('NVL_REQUIRE_READINESS: "1"', workflow)
        self.assertIn(".runtime-state/nvl/latest_release.json", workflow)
        self.assertIn(".runtime-state/nvl/releases/", workflow)
        self.assertIn("nvl_release_manifest.json", workflow)
        self.assertIn("contents: write", workflow)

    def test_release_history_is_audited_and_indexed_before_push(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-nvl-stock.yml"
        ).read_text(encoding="utf-8")
        self.assertIn("scripts/audit_nvl_releases.py", workflow)
        self.assertIn(".runtime-state/nvl/release_index.json", workflow)
        self.assertIn(".runtime-state/nvl/release_index.csv", workflow)
        audit_index = workflow.find("scripts/audit_nvl_releases.py")
        commit_index = workflow.find('git commit -m "chore: record NVL release')
        push_index = workflow.find("git push origin HEAD:runtime-state")
        self.assertGreater(audit_index, 0)
        self.assertGreater(commit_index, audit_index)
        self.assertGreater(push_index, commit_index)

    def test_workflow_always_emits_operational_summary_before_artifact_upload(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-nvl-stock.yml"
        ).read_text(encoding="utf-8")
        summary_index = workflow.find("- name: Write NVL operational summary")
        upload_index = workflow.find("- name: Upload NVL proposal and report audit")
        self.assertGreaterEqual(summary_index, 0)
        self.assertGreater(upload_index, summary_index)
        self.assertIn("if: always()", workflow[summary_index:upload_index])
        self.assertIn("python -X utf8 scripts/summarize_nvl_run.py", workflow)
        self.assertIn("JOB_STATUS: ${{ job.status }}", workflow)
        self.assertIn("nvl_operational_summary.json", workflow)

    def test_readiness_report_is_uploaded(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-nvl-stock.yml"
        ).read_text(encoding="utf-8")
        self.assertIn(REPORT_NAME, workflow)

    def test_readiness_uses_checked_out_release_sha_for_provenance(self):
        gate = (
            ROOT / "scripts" / "nvl_production_readiness.py"
        ).read_text(encoding="utf-8")
        self.assertIn('os.getenv("NVL_RELEASE_COMMIT_SHA")', gate)
        self.assertIn('or os.getenv("GITHUB_SHA")', gate)

    def test_gate_declares_all_required_phases_in_order(self):
        gate = (
            ROOT / "scripts" / "nvl_production_readiness.py"
        ).read_text(encoding="utf-8")
        phases = [
            '"architecture"',
            '"business_contracts"',
            '"workbook_roundtrip"',
            '"deterministic_proposal"',
            '"etag_concurrency"',
            '"publish_dry_run"',
            '"full_regression"',
        ]
        positions = [gate.find(phase) for phase in phases]
        self.assertTrue(all(position >= 0 for position in positions))
        self.assertEqual(positions, sorted(positions))


if __name__ == "__main__":
    unittest.main()
