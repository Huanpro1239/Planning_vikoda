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


    def test_nvl_runtime_has_no_direct_sync_stock_import(self):
        paths = [
            ROOT / "sync_nvl_stock.py",
            ROOT / "sync_nvl_open_po.py",
            *sorted((ROOT / "nvl").glob("*.py")),
        ]
        offenders = []
        for path in paths:
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            for node in ast.walk(tree):
                if isinstance(node, ast.Import):
                    if any(alias.name == "sync_stock" for alias in node.names):
                        offenders.append(str(path.relative_to(ROOT)))
                elif isinstance(node, ast.ImportFrom) and node.module == "sync_stock":
                    offenders.append(str(path.relative_to(ROOT)))
        self.assertEqual(offenders, [], f"NVL còn phụ thuộc trực tiếp sync_stock: {offenders}")

    def test_legacy_nvl_cli_is_thin(self):
        line_count = len((ROOT / "sync_nvl_stock.py").read_text(encoding="utf-8").splitlines())
        self.assertLess(line_count, 180, f"sync_nvl_stock.py vẫn quá lớn: {line_count} dòng")


if __name__ == "__main__":
    unittest.main()
