import unittest
from pathlib import Path


class OIDCWorkflowContractTests(unittest.TestCase):
    def _read(self, name):
        return Path(".github/workflows", name).read_text(encoding="utf-8")

    def test_planning_workflow_is_oidc_only(self):
        text = self._read("sync-stock.yml")
        self.assertIn("id-token: write", text)
        self.assertIn("MS_AUTH_MODE: ${{ vars.MS_AUTH_MODE || 'oidc' }}", text)
        self.assertIn("python -X utf8 sync_planning_pipeline.py", text)
        self.assertNotIn("auth_runner.py", text)
        self.assertNotIn("MS_CLIENT_SECRET", text)

    def test_nvl_workflow_is_oidc_only(self):
        text = self._read("sync-nvl-stock.yml")
        self.assertIn("id-token: write", text)
        self.assertIn("MS_AUTH_MODE: ${{ vars.MS_AUTH_MODE || 'oidc' }}", text)
        self.assertIn("python -X utf8 sync_nvl_stock.py", text)
        self.assertIn("python -X utf8 sync_nvl_open_po_safe.py", text)
        self.assertIn("python -X utf8 scripts/survey_nvl_sharepoint.py", text)
        self.assertIn("python -X utf8 scripts/ensure_staging_copy.py", text)
        self.assertNotIn("auth_runner.py", text)
        self.assertNotIn("MS_CLIENT_SECRET", text)

    def test_workflows_use_node24_actions(self):
        planning = self._read("sync-stock.yml")
        nvl = self._read("sync-nvl-stock.yml")
        tests = self._read("test.yml")

        for text in (planning, nvl, tests):
            self.assertIn("actions/checkout@v5", text)
            self.assertIn("actions/setup-python@v6", text)
            self.assertNotIn("actions/checkout@v4", text)
            self.assertNotIn("actions/setup-python@v5", text)

        self.assertIn("actions/upload-artifact@v6", planning)
        self.assertIn("actions/upload-artifact@v6", nvl)
        self.assertNotIn("actions/upload-artifact@v4", planning)
        self.assertNotIn("actions/upload-artifact@v4", nvl)


if __name__ == "__main__":
    unittest.main()
