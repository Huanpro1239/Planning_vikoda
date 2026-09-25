"""Architecture guardrails for the NVL package."""

import ast
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[2]
NVL_ROOT = ROOT / "nvl"

ALLOWED_LOCAL_IMPORTS = {
    "models": set(),
    "values": set(),
    "config": {"models"},
    "reconcile": {"models", "values"},
    "workbook": {"models", "values"},
    "reporting": {"models"},
    "service": {"models", "reconcile", "reporting", "workbook"},
    "stock": {"config", "models", "reconcile", "reporting", "service", "values", "workbook"},
    "open_po": {"values"},
    "open_po_safe": {"open_po"},
}


def _module_name(path: Path) -> str:
    return path.stem


def _imported_modules(path: Path) -> list[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    modules: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            modules.extend(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            modules.append(node.module)
    return modules


class NVLArchitectureTests(unittest.TestCase):
    def test_nvl_dependency_direction_is_one_way(self):
        offenders = []
        graph = {name: set() for name in ALLOWED_LOCAL_IMPORTS}

        for path in sorted(NVL_ROOT.glob("*.py")):
            name = _module_name(path)
            if name == "__init__":
                continue
            self.assertIn(name, ALLOWED_LOCAL_IMPORTS, f"Thiếu architecture rule cho nvl/{path.name}")

            allowed = ALLOWED_LOCAL_IMPORTS[name]
            for module in _imported_modules(path):
                if not module.startswith("nvl."):
                    continue
                local = module.split(".", 1)[1].split(".", 1)[0]
                graph[name].add(local)
                if local not in allowed:
                    offenders.append(
                        f"{path.name}: {module} not in allowed={sorted(allowed)}"
                    )

        self.assertEqual(
            offenders,
            [],
            "NVL dependency boundary bị đảo: " + "; ".join(offenders),
        )

        visiting = set()
        visited = set()

        def visit(node: str, trail: list[str]) -> None:
            if node in visiting:
                cycle = " -> ".join(trail + [node])
                self.fail(f"NVL dependency cycle: {cycle}")
            if node in visited:
                return
            visiting.add(node)
            for dep in sorted(graph.get(node, ())):
                if dep in graph:
                    visit(dep, trail + [node])
            visiting.remove(node)
            visited.add(node)

        for node in sorted(graph):
            visit(node, [])

    def test_nvl_does_not_depend_on_planning_stock_or_root_sync_modules(self):
        forbidden = []
        for path in sorted(NVL_ROOT.glob("*.py")):
            for module in _imported_modules(path):
                if (
                    module == "planning"
                    or module.startswith("planning.")
                    or module == "stock"
                    or module.startswith("stock.")
                    or module.startswith("sync_")
                ):
                    forbidden.append(f"{path.name}: {module}")

        self.assertEqual(
            forbidden,
            [],
            "NVL canonical package có dependency ngoài boundary: " + "; ".join(forbidden),
        )


if __name__ == "__main__":
    unittest.main()
