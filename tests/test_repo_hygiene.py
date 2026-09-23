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
            ROOT / "planning" / "khsx_ki" / "__init__.py",
            ROOT / "planning" / "publish" / "__init__.py",
            ROOT / "planning" / "weekly_model" / "__init__.py",
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
            "sync_planning_fc",
            "sync_planning_fc_compat",
            "sync_planning_calendar",
            "sync_planning_calendar_all_months",
            "sync_planning_layout",
            "sync_planning_stock_inputs",
            "sync_planning_pipeline",
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
            ROOT / "sync_planning_fc_compat.py",
            ROOT / "sync_planning_calendar_all_months.py",
            ROOT / "sync_planning_layout.py",
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


    def test_khsx_ki_is_package_and_monolith_is_removed(self):
        self.assertFalse((ROOT / "planning" / "khsx_ki.py").exists())
        required = {
            "__init__.py",
            "calendar.py",
            "layout.py",
            "workbook.py",
            "verification.py",
        }
        actual = {
            path.name
            for path in (ROOT / "planning" / "khsx_ki").glob("*.py")
        }
        self.assertTrue(
            required <= actual,
            f"Thiếu KHSX_ki modules: {required - actual}",
        )

    def test_khsx_ki_dependency_direction(self):
        allowed = {
            "calendar.py": set(),
            "layout.py": {"calendar"},
            "workbook.py": {"calendar", "layout"},
            "verification.py": {"calendar", "layout", "workbook"},
        }
        offenders = []
        root = ROOT / "planning" / "khsx_ki"
        for filename, allowed_local in allowed.items():
            path = root / filename
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            for node in ast.walk(tree):
                if not isinstance(node, ast.ImportFrom):
                    continue
                module = node.module or ""
                if node.level != 1:
                    continue
                local = module.split(".")[0]
                if local and local not in allowed_local:
                    offenders.append(
                        f"{filename}: .{module} not in {sorted(allowed_local)}"
                    )
        self.assertEqual(
            offenders,
            [],
            "KHSX_ki dependency boundary bị đảo: " + "; ".join(offenders),
        )

    def test_khsx_ki_modules_stay_bounded(self):
        limits = {
            "calendar.py": 150,
            "layout.py": 350,
            "workbook.py": 500,
            "verification.py": 300,
            "__init__.py": 80,
        }
        oversized = []
        root = ROOT / "planning" / "khsx_ki"
        for name, limit in limits.items():
            count = len((root / name).read_text(encoding="utf-8").splitlines())
            if count >= limit:
                oversized.append(f"{name}={count} dòng")
        self.assertEqual(
            oversized,
            [],
            "KHSX_ki module bị phình lại: " + "; ".join(oversized),
        )


    def test_finished_goods_runtime_has_no_direct_sync_stock_import(self):
        offenders = []
        for path in sorted(ROOT.rglob("*.py")):
            if path == ROOT / "sync_stock.py":
                continue
            if "tests" in path.parts:
                continue
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            for node in ast.walk(tree):
                if isinstance(node, ast.Import):
                    if any(alias.name == "sync_stock" for alias in node.names):
                        offenders.append(str(path.relative_to(ROOT)))
                elif isinstance(node, ast.ImportFrom) and node.module == "sync_stock":
                    offenders.append(str(path.relative_to(ROOT)))
        self.assertEqual(
            offenders,
            [],
            "Runtime còn phụ thuộc root sync_stock: " + ", ".join(offenders),
        )

    def test_legacy_sync_stock_cli_is_thin(self):
        line_count = len(
            (ROOT / "sync_stock.py").read_text(encoding="utf-8").splitlines()
        )
        self.assertLess(
            line_count,
            150,
            f"sync_stock.py vẫn quá lớn: {line_count} dòng",
        )

    def test_sync_stock_compat_is_removed(self):
        self.assertFalse((ROOT / "sync_stock_compat.py").exists())

    def test_stock_package_does_not_depend_on_planning_or_sync_stock(self):
        offenders = []
        for path in sorted((ROOT / "stock").glob("*.py")):
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            for node in ast.walk(tree):
                modules = []
                if isinstance(node, ast.Import):
                    modules = [alias.name for alias in node.names]
                elif isinstance(node, ast.ImportFrom) and node.module:
                    modules = [node.module]
                for module in modules:
                    if (
                        module == "sync_stock"
                        or module == "planning"
                        or module.startswith("planning.")
                    ):
                        offenders.append(
                            f"{path.relative_to(ROOT)}: {module}"
                        )
        self.assertEqual(
            offenders,
            [],
            "stock package dependency bị đảo: " + "; ".join(offenders),
        )


    def test_planning_root_entrypoints_are_thin(self):
        limits = {
            "sync_planning_fc.py": 30,
            "sync_planning_calendar.py": 30,
            "sync_planning_stock_inputs.py": 30,
            "sync_planning_pipeline.py": 30,
        }
        oversized = []
        for name, limit in limits.items():
            count = len((ROOT / name).read_text(encoding="utf-8").splitlines())
            if count >= limit:
                oversized.append(f"{name}={count} dòng")
        self.assertEqual(
            oversized,
            [],
            "Planning root entrypoint bị phình lại: " + "; ".join(oversized),
        )

    def test_canonical_planning_does_not_import_root_sync_modules(self):
        offenders = []
        for path in sorted((ROOT / "planning").rglob("*.py")):
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            for node in ast.walk(tree):
                modules = []
                if isinstance(node, ast.Import):
                    modules = [alias.name for alias in node.names]
                elif isinstance(node, ast.ImportFrom) and node.module:
                    modules = [node.module]
                for module in modules:
                    if module.startswith("sync_planning_"):
                        offenders.append(
                            f"{path.relative_to(ROOT)}: {module}"
                        )
        self.assertEqual(
            offenders,
            [],
            "Canonical planning còn import root sync module: " + "; ".join(offenders),
        )

    def test_planning_has_no_import_time_compat_monkey_patches(self):
        forbidden_assignments = {
            ("fc", "read_planning_fc_targets"),
            ("calendar_sync", "prepare_calendar_update"),
        }
        offenders = []
        for path in sorted(ROOT.rglob("*.py")):
            if "tests" in path.parts:
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
                        and (target.value.id, target.attr) in forbidden_assignments
                    ):
                        offenders.append(
                            f"{path.relative_to(ROOT)}:{getattr(node, 'lineno', '?')}"
                        )
        self.assertEqual(
            offenders,
            [],
            "Planning compatibility monkey-patch quay trở lại: " + "; ".join(offenders),
        )


    def test_publish_is_package_and_monolith_is_removed(self):
        self.assertFalse((ROOT / "planning" / "publish.py").exists())
        required = {
            "__init__.py",
            "constants.py",
            "snapshot.py",
            "proposal.py",
            "release.py",
            "policy.py",
            "state.py",
            "service.py",
            "runner.py",
        }
        actual = {
            path.name
            for path in (ROOT / "planning" / "publish").glob("*.py")
        }
        self.assertTrue(
            required <= actual,
            f"Thiếu Planning publish modules: {required - actual}",
        )

    def test_publish_dependency_direction(self):
        allowed = {
            "constants.py": set(),
            "snapshot.py": {"constants"},
            "proposal.py": {"constants"},
            "release.py": {"proposal"},
            "policy.py": {"constants"},
            "state.py": {"constants"},
            "service.py": {"policy", "proposal", "release", "snapshot", "state"},
            "runner.py": {"service"},
        }
        offenders = []
        root = ROOT / "planning" / "publish"
        for filename, allowed_local in allowed.items():
            path = root / filename
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            for node in ast.walk(tree):
                if not isinstance(node, ast.ImportFrom):
                    continue
                module = node.module or ""
                if node.level != 1:
                    continue
                local = module.split(".")[0]
                if local and local not in allowed_local:
                    offenders.append(
                        f"{filename}: .{module} not in {sorted(allowed_local)}"
                    )
        self.assertEqual(
            offenders,
            [],
            "Planning publish dependency boundary bị đảo: " + "; ".join(offenders),
        )

    def test_publish_modules_stay_bounded(self):
        limits = {
            "constants.py": 80,
            "snapshot.py": 100,
            "proposal.py": 140,
            "release.py": 260,
            "policy.py": 160,
            "state.py": 200,
            "service.py": 320,
            "runner.py": 130,
            "__init__.py": 100,
        }
        oversized = []
        root = ROOT / "planning" / "publish"
        for name, limit in limits.items():
            count = len((root / name).read_text(encoding="utf-8").splitlines())
            if count >= limit:
                oversized.append(f"{name}={count} dòng")
        self.assertEqual(
            oversized,
            [],
            "Planning publish module bị phình lại: " + "; ".join(oversized),
        )


    def test_weekly_model_is_package_and_monolith_is_removed(self):
        self.assertFalse((ROOT / "planning" / "weekly_model.py").exists())
        required = {
            "__init__.py",
            "policy.py",
            "inputs.py",
            "schedule.py",
            "workbook.py",
            "report.py",
            "verification.py",
            "service.py",
        }
        actual = {
            path.name
            for path in (ROOT / "planning" / "weekly_model").glob("*.py")
        }
        self.assertTrue(
            required <= actual,
            f"Thiếu weekly_model modules: {required - actual}",
        )

    def test_weekly_model_dependency_direction(self):
        allowed = {
            "policy.py": set(),
            "inputs.py": {"policy"},
            "schedule.py": {"inputs"},
            "workbook.py": {"inputs", "schedule"},
            "report.py": {"inputs", "schedule"},
            "verification.py": {"report", "schedule"},
            "service.py": {"report", "schedule", "workbook"},
        }
        offenders = []
        root = ROOT / "planning" / "weekly_model"
        for filename, allowed_local in allowed.items():
            path = root / filename
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            for node in ast.walk(tree):
                if not isinstance(node, ast.ImportFrom):
                    continue
                module = node.module or ""
                if node.level != 1:
                    continue
                local = module.split(".")[0]
                if local and local not in allowed_local:
                    offenders.append(
                        f"{filename}: .{module} not in {sorted(allowed_local)}"
                    )
        self.assertEqual(
            offenders,
            [],
            "weekly_model dependency boundary bị đảo: " + "; ".join(offenders),
        )

    def test_weekly_model_modules_stay_bounded(self):
        limits = {
            "policy.py": 150,
            "inputs.py": 320,
            "schedule.py": 300,
            "workbook.py": 160,
            "report.py": 400,
            "verification.py": 240,
            "service.py": 80,
            "__init__.py": 100,
        }
        oversized = []
        root = ROOT / "planning" / "weekly_model"
        for name, limit in limits.items():
            count = len((root / name).read_text(encoding="utf-8").splitlines())
            if count >= limit:
                oversized.append(f"{name}={count} dòng")
        self.assertEqual(
            oversized,
            [],
            "weekly_model module bị phình lại: " + "; ".join(oversized),
        )

    def test_weekly_engine_does_not_depend_on_weekly_model(self):
        path = ROOT / "planning" / "weekly_engine.py"
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        offenders = []
        for node in ast.walk(tree):
            modules = []
            if isinstance(node, ast.Import):
                modules = [alias.name for alias in node.names]
            elif isinstance(node, ast.ImportFrom) and node.module:
                modules = [node.module]
            for module in modules:
                if module == "planning.weekly_model" or module.startswith(
                    "planning.weekly_model."
                ):
                    offenders.append(module)
        self.assertEqual(
            offenders,
            [],
            "weekly_engine phải giữ thuần, không phụ thuộc weekly_model: "
            + "; ".join(offenders),
        )


if __name__ == "__main__":
    unittest.main()
