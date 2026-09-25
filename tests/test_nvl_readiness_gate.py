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

    def test_gate_runs_before_any_nvl_sharepoint_operation(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-nvl-stock.yml"
        ).read_text(encoding="utf-8")
        gate_index = workflow.find("- name: Run NVL production readiness gate")
        staging_index = workflow.find("- name: Ensure staging copy on SharePoint")
        publish_index = workflow.find("- name: Generate NVL proposal or controlled publish")
        self.assertGreaterEqual(gate_index, 0)
        self.assertGreater(staging_index, gate_index)
        self.assertGreater(publish_index, gate_index)

    def test_readiness_report_is_uploaded(self):
        workflow = (
            ROOT / ".github" / "workflows" / "sync-nvl-stock.yml"
        ).read_text(encoding="utf-8")
        self.assertIn(REPORT_NAME, workflow)

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
