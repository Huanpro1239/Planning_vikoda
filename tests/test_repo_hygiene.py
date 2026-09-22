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


    def test_graph_api_is_not_imported_from_sync_stock(self):
        forbidden = {
            "GRAPH",
            "HOSTNAME",
            "SITE_PATH",
            "GraphClient",
            "GraphRequestError",
            "get_access_token",
            "is_retryable_graph_error",
        }
        offenders = []
        for path in sorted(ROOT.rglob("*.py")):
            if path.name == "sync_stock.py":
                continue
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            for node in ast.walk(tree):
                if isinstance(node, ast.ImportFrom) and node.module == "sync_stock":
                    imported = {alias.name for alias in node.names}
                    bad = sorted(imported & forbidden)
                    if bad:
                        offenders.append(
                            f"{path.relative_to(ROOT)}: {', '.join(bad)}"
                        )
        self.assertEqual(
            offenders,
            [],
            "Graph/auth API phải import từ sharepoint.client, không từ sync_stock: "
            + "; ".join(offenders),
        )

    def test_sharepoint_client_does_not_depend_on_sync_stock(self):
        path = ROOT / "sharepoint" / "client.py"
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        offenders = []
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                offenders.extend(
                    alias.name for alias in node.names if alias.name == "sync_stock"
                )
            elif isinstance(node, ast.ImportFrom) and node.module == "sync_stock":
                offenders.append("from sync_stock")
        self.assertEqual(offenders, [], "sharepoint.client không được phụ thuộc legacy sync_stock")

    def test_auth_runner_is_removed(self):
        self.assertFalse((ROOT / "scripts" / "auth_runner.py").exists())


    def test_planning_runtime_uses_canonical_modules(self):
        forbidden_modules = {
            "sync_planning_metrics",
            "sync_planning_metrics_compat",
            "sync_planning_metrics_direct",
            "sync_planning_metrics_all_months",
            "sync_planning_khsx_ki",
            "planning_pipeline",
        }
        shim_paths = set()
        offenders = []
        for path in sorted(ROOT.rglob("*.py")):
            if path in shim_paths:
                continue
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            for node in ast.walk(tree):
                if isinstance(node, ast.Import):
                    for alias in node.names:
                        if alias.name in forbidden_modules:
                            offenders.append(
                                f"{path.relative_to(ROOT)}: import {alias.name}"
                            )
                elif isinstance(node, ast.ImportFrom) and node.module in forbidden_modules:
                    offenders.append(
                        f"{path.relative_to(ROOT)}: from {node.module}"
                    )
        self.assertEqual(
            offenders,
            [],
            "Planning runtime/test phải dùng planning.* canonical: " + "; ".join(offenders),
        )

    def test_legacy_planning_root_modules_are_removed(self):
        legacy = [
            ROOT / "sync_planning_metrics.py",
            ROOT / "sync_planning_metrics_compat.py",
            ROOT / "sync_planning_metrics_direct.py",
            ROOT / "sync_planning_metrics_all_months.py",
            ROOT / "sync_planning_khsx_ki.py",
            ROOT / "planning_pipeline.py",
        ]
        leftovers = [str(path.relative_to(ROOT)) for path in legacy if path.exists()]
        self.assertEqual(leftovers, [], "Planning shim cũ vẫn còn: " + ", ".join(leftovers))


    def test_planning_metrics_is_package_and_monolith_is_removed(self):
        self.assertFalse((ROOT / "planning" / "metrics.py").exists())
        required = {
            "constants.py",
            "state.py",
            "readers.py",
            "calculation.py",
            "workbook.py",
            "service.py",
            "__init__.py",
        }
        actual = {
            path.name
            for path in (ROOT / "planning" / "metrics").glob("*.py")
        }
        self.assertTrue(required <= actual, f"Thiếu metrics modules: {required - actual}")

    def test_planning_metrics_has_no_runtime_monkey_patching(self):
        offenders = []
        for path in sorted(ROOT.rglob("*.py")):
            if path == ROOT / "tests" / "test_repo_hygiene.py":
                continue
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            for node in ast.walk(tree):
                targets = []
                if isinstance(node, ast.Assign):
                    targets = node.targets
                elif isinstance(node, (ast.AnnAssign, ast.AugAssign)):
                    targets = [node.target]
                for target in targets:
                    if (
                        isinstance(target, ast.Attribute)
                        and isinstance(target.value, ast.Name)
                        and target.value.id == "metrics"
                    ):
                        offenders.append(
                            f"{path.relative_to(ROOT)}:{getattr(node, 'lineno', '?')} "
                            f"metrics.{target.attr}"
                        )
        self.assertEqual(
            offenders,
            [],
            "Không được rebind planning metrics lúc runtime: " + "; ".join(offenders),
        )


if __name__ == "__main__":
    unittest.main()
