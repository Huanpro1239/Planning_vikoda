import unittest
from pathlib import Path


class NVLWorkflowContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text = Path(".github/workflows/sync-nvl-stock.yml").read_text(encoding="utf-8")

    def test_automatic_nvl_follows_planning_completion(self):
        self.assertIn("workflow_run:", self.text)
        self.assertIn("- Sync SharePoint Stock", self.text)
        self.assertIn("- completed", self.text)
        self.assertNotIn('cron: "15 23 * * *"', self.text)

    def test_sharepoint_dispatch_events_are_enabled(self):
        for event_name in (
            "sharepoint_nvl_updated",
            "nvl_stock_updated",
            "nvl_open_po_updated",
        ):
            self.assertIn(event_name, self.text)

    def test_automatic_events_publish_but_manual_default_remains_controlled(self):
        self.assertIn("github.event_name == 'workflow_run'", self.text)
        self.assertIn(
            "github.event.workflow_run.event == 'schedule'",
            self.text,
        )
        self.assertIn(
            "github.event.workflow_run.event == 'repository_dispatch'",
            self.text,
        )
        self.assertIn("github.event_name == 'repository_dispatch'", self.text)
        self.assertIn(
            "github.event_name == 'workflow_dispatch' && inputs.publish",
            self.text,
        )

    def test_obsolete_feature_branch_push_trigger_removed(self):
        self.assertNotIn("feat/sync-nvl-stock", self.text)


if __name__ == "__main__":
    unittest.main()
