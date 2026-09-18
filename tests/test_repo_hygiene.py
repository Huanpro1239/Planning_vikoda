"""Repository hygiene regression checks."""

import ast
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class RepoHygieneTests(unittest.TestCase):
    def test_root_has_no_temporary_review_or_antigravity_docs(self):
        forbidden_patterns = ("REVIEW_*.md", "RECHECK_*.md", "PLAN_*ANTIGRAVITY*.md")
        leftovers = []
        for pattern in forbidden_patterns:
            leftovers.extend(path.name for path in ROOT.glob(pattern))
        self.assertEqual(leftovers, [], f"Root còn tài liệu tạm: {leftovers}")

    def test_dev_only_trigger_tool_is_not_in_production_scripts_root(self):
        self.assertFalse((ROOT / "scripts" / "trigger_and_download_dryrun.py").exists())
        self.assertTrue(
            (ROOT / "scripts" / "dev" / "trigger_and_download_dryrun.py").is_file()
        )

    def test_package_facades_do_not_use_wildcard_imports(self):
        paths = (
            ROOT / "planning" / "pipeline.py",
            ROOT / "nvl" / "stock.py",
            ROOT / "nvl" / "open_po.py",
        )
        for path in paths:
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            wildcard_imports = [
                node
                for node in ast.walk(tree)
                if isinstance(node, ast.ImportFrom)
                and any(alias.name == "*" for alias in node.names)
            ]
            self.assertEqual(
                wildcard_imports,
                [],
                f"{path.relative_to(ROOT)} không được dùng import *",
            )


if __name__ == "__main__":
    unittest.main()
