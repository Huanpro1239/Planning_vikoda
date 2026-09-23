"""The same readiness command must gate merge CI and production publish."""

from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[2]
GATE_COMMAND = "python -X utf8 scripts/production_readiness.py"


class ReleaseGateWiringTests(unittest.TestCase):
    def test_pull_request_ci_uses_single_readiness_gate(self):
        workflow = (
            ROOT / ".github" / "workflows" / "test.yml"
        ).read_text(encoding="utf-8")
        self.assertIn(GATE_COMMAND, workflow)
        self.assertNotIn(
            "python -m unittest discover -s tests -v",
            workflow,
        )

    def test_production_publish_runs_same_gate_before_pipeline(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-stock.yml"
        ).read_text(encoding="utf-8")

        gate_index = workflow.find(GATE_COMMAND)
        publish_index = workflow.find(
            "python -X utf8 sync_planning_pipeline.py"
        )
        self.assertGreaterEqual(gate_index, 0)
        self.assertGreater(publish_index, gate_index)

    def test_production_gate_cannot_be_conditionally_skipped(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-stock.yml"
        ).read_text(encoding="utf-8")
        gate_step = workflow.split(
            "- name: Run production readiness gate",
            1,
        )[1].split("- name:", 1)[0]
        self.assertNotIn("\n        if:", gate_step)
        self.assertIn(GATE_COMMAND, gate_step)


    def test_production_audit_persists_release_manifest_history(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-stock.yml"
        ).read_text(encoding="utf-8")
        self.assertIn("planning_release_manifest.json", workflow)
        self.assertIn(".runtime-state/latest_release.json", workflow)
        self.assertIn(".runtime-state/releases/", workflow)
        self.assertIn('PLANNING_REQUIRE_READINESS: "1"', workflow)


    def test_release_history_step_has_no_unindented_python_heredoc(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-stock.yml"
        ).read_text(encoding="utf-8")
        self.assertNotIn("<<'PY'", workflow)
        self.assertIn(
            'RELEASE_ID="$(python -c ',
            workflow,
        )

    def test_release_history_detects_untracked_manifest_files(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-stock.yml"
        ).read_text(encoding="utf-8")
        self.assertIn(
            "git status --porcelain -- state.json planning_runtime.json "
            "latest_release.json releases",
            workflow,
        )
        self.assertNotIn(
            "git diff --quiet -- state.json planning_runtime.json "
            "latest_release.json releases",
            workflow,
        )


if __name__ == "__main__":
    unittest.main()
