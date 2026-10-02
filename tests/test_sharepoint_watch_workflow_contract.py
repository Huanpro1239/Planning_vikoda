import unittest
from pathlib import Path


class SharePointWatchWorkflowContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.watch = Path(
            ".github/workflows/sharepoint-watch.yml"
        ).read_text(encoding="utf-8")
        cls.planning = Path(
            ".github/workflows/sync-stock.yml"
        ).read_text(encoding="utf-8")

    def test_watcher_polls_and_dispatches_expected_event(self):
        self.assertIn('cron: "*/10 * * * *"', self.watch)
        self.assertIn(
            '"event_type": "sharepoint_stock_updated"',
            self.watch,
        )
        self.assertIn(
            '"source": "github-sharepoint-watch"',
            self.watch,
        )

    def test_watcher_serializes_with_planning(self):
        self.assertIn("group: sharepoint-stock-sync", self.watch)
        self.assertIn("cancel-in-progress: false", self.watch)
        self.assertIn("group: sharepoint-stock-sync", self.planning)

    def test_watcher_is_oidc_and_read_only_until_dispatch(self):
        self.assertIn("id-token: write", self.watch)
        self.assertIn(
            "python -X utf8 scripts/watch_sharepoint_changes.py",
            self.watch,
        )
        self.assertNotIn("planning.publish --publish", self.watch)

    def test_planning_listener_keeps_skip_if_unchanged(self):
        self.assertIn("sharepoint_stock_updated", self.planning)
        self.assertIn('EXTRA_FLAGS="--skip-if-unchanged"', self.planning)
        self.assertIn("sharepoint-event", self.planning)


if __name__ == "__main__":
    unittest.main()
