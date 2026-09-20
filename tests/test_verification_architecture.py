import ast
import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
VERIFICATION_DIR = REPO_ROOT / "planning" / "verification"

ALLOWED_INTERNAL_DEPENDENCIES = {
    "__init__": {"workbook", "row_mass_balance", "resources"},
    "workbook": set(),
    "row_mass_balance": {"workbook", "shared_machine"},
    "resources": {"shared_machine"},
    "shared_machine": set(),
}


def _module_path(module_name):
    filename = "__init__.py" if module_name == "__init__" else f"{module_name}.py"
    return VERIFICATION_DIR / filename


def _parse_module(module_name):
    return ast.parse(
        _module_path(module_name).read_text(encoding="utf-8"),
        filename=str(_module_path(module_name)),
    )


def _internal_dependencies(module_name):
    dependencies = set()
    tree = _parse_module(module_name)

    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom):
            if any(alias.name == "*" for alias in node.names):
                raise AssertionError(
                    f"{module_name} uses wildcard import at line {node.lineno}."
                )

            if node.level == 1:
                if node.module:
                    dependencies.add(node.module.split(".", 1)[0])
                continue

            imported = node.module or ""
            prefix = "planning.verification"
            if imported == prefix:
                dependencies.add("__init__")
            elif imported.startswith(prefix + "."):
                dependencies.add(imported[len(prefix) + 1 :].split(".", 1)[0])

        elif isinstance(node, ast.Import):
            for alias in node.names:
                imported = alias.name
                prefix = "planning.verification"
                if imported == prefix:
                    dependencies.add("__init__")
                elif imported.startswith(prefix + "."):
                    dependencies.add(imported[len(prefix) + 1 :].split(".", 1)[0])

    dependencies.discard(module_name)
    return dependencies


class VerificationArchitectureTests(unittest.TestCase):
    def test_verification_modules_are_explicitly_registered(self):
        discovered = {
            path.stem
            for path in VERIFICATION_DIR.glob("*.py")
            if path.name != "__pycache__"
        }
        discovered.discard("__init__")
        discovered.add("__init__")

        self.assertEqual(
            discovered,
            set(ALLOWED_INTERNAL_DEPENDENCIES),
            "New verification modules must be added to the dependency policy explicitly.",
        )

    def test_internal_dependencies_follow_layer_policy(self):
        for module_name, allowed in ALLOWED_INTERNAL_DEPENDENCIES.items():
            with self.subTest(module=module_name):
                actual = _internal_dependencies(module_name)
                self.assertEqual(
                    actual,
                    allowed,
                    f"{module_name} changed its verification-layer dependencies. "
                    "Update architecture intentionally instead of adding a reverse import.",
                )

    def test_verification_dependency_graph_is_acyclic(self):
        graph = {
            module_name: _internal_dependencies(module_name)
            for module_name in ALLOWED_INTERNAL_DEPENDENCIES
        }
        visiting = set()
        visited = set()

        def visit(module_name, path):
            if module_name in visiting:
                cycle = " -> ".join(path + [module_name])
                self.fail(f"Verification dependency cycle detected: {cycle}")
            if module_name in visited:
                return

            visiting.add(module_name)
            for dependency in graph[module_name]:
                visit(dependency, path + [module_name])
            visiting.remove(module_name)
            visited.add(module_name)

        for module_name in graph:
            visit(module_name, [])

    def test_package_init_remains_orchestration_only(self):
        tree = _parse_module("__init__")
        top_level_functions = [
            node.name for node in tree.body if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        ]
        top_level_classes = [
            node.name for node in tree.body if isinstance(node, ast.ClassDef)
        ]

        self.assertEqual(top_level_functions, ["verify_workbook", "main"])
        self.assertEqual(top_level_classes, [])

        verify_workbook = next(
            node
            for node in tree.body
            if isinstance(node, ast.FunctionDef) and node.name == "verify_workbook"
        )
        called_names = [
            node.func.id
            for node in ast.walk(verify_workbook)
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
        ]

        for expected_call in (
            "validate_workbook_context",
            "validate_planning_rows",
            "validate_schedule_resources",
        ):
            self.assertEqual(
                called_names.count(expected_call),
                1,
                f"verify_workbook must orchestrate {expected_call} exactly once.",
            )

        forbidden_detail_calls = {
            name
            for name in called_names
            if name.startswith("_validate_") or name.startswith("_find_shared_")
        }
        self.assertEqual(
            forbidden_detail_calls,
            set(),
            "Package orchestration must not call shared-machine implementation details directly.",
        )


if __name__ == "__main__":
    unittest.main()
