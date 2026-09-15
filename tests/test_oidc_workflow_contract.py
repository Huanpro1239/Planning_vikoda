import unittest
from pathlib import Path


class OIDCWorkflowContractTests(unittest.TestCase):
    def _read(self, name):
        return Path(".github/workflows", name).read_text(encoding="utf-8")

    def test_planning_workflow_has_oidc_permission_and_safe_default(self):
        text = self._read("sync-stock.yml")
        self.assertIn("id-token: write", text)
        self.assertIn("MS_AUTH_MODE: ${{ vars.MS_AUTH_MODE || 'secret' }}", text)
        self.assertIn("scripts/auth_runner.py sync_planning_pipeline", text)
        self.assertIn("MS_CLIENT_SECRET", text)  # transitional until Entra cutover is verified

    def test_nvl_workflow_has_oidc_permission_and_runner(self):
        text = self._read("sync-nvl-stock.yml")
        self.assertIn("id-token: write", text)
        self.assertIn("MS_AUTH_MODE: ${{ vars.MS_AUTH_MODE || 'secret' }}", text)
        self.assertIn("scripts/auth_runner.py sync_nvl_stock", text)
        self.assertIn("scripts/auth_runner.py sync_nvl_open_po_safe", text)
        self.assertIn("scripts/auth_runner.py scripts.survey_nvl_sharepoint", text)
        self.assertIn("scripts/auth_runner.py scripts.ensure_staging_copy", text)


if __name__ == "__main__":
    unittest.main()
