"""Package-level dependency architecture for the Planning domain."""

import ast
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]
PLANNING = ROOT / "planning"

DOMAIN_PREFIXES = (
    ("planning.weekly_engine", "weekly_engine"),
    ("planning.metrics", "metrics"),
    ("planning.weekly_model", "weekly_model"),
    ("planning.khsx_ki", "khsx_ki"),
    ("planning.pipeline", "pipeline"),
    ("planning.publish", "publish"),
    ("planning.verification", "verification"),
)

SUPPORT_MODULES = {
    "planning.fc": "fc",
    "planning.layout": "layout",
    "planning.calendar": "calendar",
    "planning.stock_inputs": "stock_inputs",
}

# Package-level architecture:
#
#   support/workbook primitives
#          |
#   weekly_engine
#      ^      ^
#      |      |
# metrics  weekly_model   khsx_ki
#       \      |        /
#              pipeline
#             /        \
#        publish      verification
#
# publish is the online proposal/publish shell. verification is an independent
# read-only outer verifier. Neither is allowed to leak back into lower layers.
ALLOWED_DOMAIN_EDGES = {
    "weekly_engine": set(),
    "metrics": {"fc"},
    "weekly_model": {"weekly_engine"},
    "khsx_ki": set(),
    "pipeline": {
        "fc",
        "calendar",
        "stock_inputs",
        "metrics",
        "weekly_model",
        "khsx_ki",
    },
    "publish": {"pipeline", "metrics"},
    "verification": {
        "fc",
        "calendar",
        "weekly_engine",
    },
    "fc": set(),
    "layout": {"fc"},
    "calendar": {"fc", "layout"},
    "stock_inputs": {"fc"},
}

MIDDLE_DOMAINS = {"metrics", "weekly_model", "khsx_ki"}


def module_name(path: Path) -> str:
    relative = path.relative_to(ROOT).with_suffix("")
    parts = list(relative.parts)
    if parts[-1] == "__init__":
        parts.pop()
    return ".".join(parts)


def domain_for_module(module: str) -> str | None:
    if module in SUPPORT_MODULES:
        return SUPPORT_MODULES[module]
    for prefix, domain in DOMAIN_PREFIXES:
        if module == prefix or module.startswith(prefix + "."):
            return domain
    return None


def resolve_from_import(current_module: str, node: ast.ImportFrom):
    """Yield concrete imported Planning module names for one ImportFrom."""
    module = node.module or ""

    if node.level:
        current_parts = current_module.split(".")
        # A package __init__ has already been normalized to the package name.
        if (PLANNING / Path(*current_parts[1:])).is_dir():
            package_parts = current_parts
        else:
            package_parts = current_parts[:-1]

        up = node.level - 1
        if up:
            package_parts = package_parts[:-up]
        base = ".".join(package_parts + ([module] if module else []))
    else:
        base = module

    if base == "planning":
        for alias in node.names:
            if alias.name != "*":
                yield f"planning.{alias.name}"
        return

    if base.startswith("planning"):
        yield base


def planning_imports(path: Path):
    current = module_name(path)
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name == "planning" or alias.name.startswith("planning."):
                    yield alias.name
        elif isinstance(node, ast.ImportFrom):
            yield from resolve_from_import(current, node)


class PlanningPackageArchitectureTests(unittest.TestCase):
    def test_package_dependency_graph(self):
        offenders = []

        for path in sorted(PLANNING.rglob("*.py")):
            source_module = module_name(path)
            source_domain = domain_for_module(source_module)
            if source_domain is None:
                continue

            allowed = ALLOWED_DOMAIN_EDGES[source_domain]
            for imported_module in planning_imports(path):
                target_domain = domain_for_module(imported_module)
                if target_domain is None or target_domain == source_domain:
                    continue
                if target_domain not in allowed:
                    offenders.append(
                        f"{path.relative_to(ROOT)}: "
                        f"{source_domain} -> {target_domain} "
                        f"via {imported_module}"
                    )

        self.assertEqual(
            offenders,
            [],
            "Planning dependency boundary bị vi phạm:\n"
            + "\n".join(offenders),
        )

    def test_middle_domains_do_not_import_each_other(self):
        offenders = []

        for path in sorted(PLANNING.rglob("*.py")):
            source_domain = domain_for_module(module_name(path))
            if source_domain not in MIDDLE_DOMAINS:
                continue

            for imported_module in planning_imports(path):
                target_domain = domain_for_module(imported_module)
                if (
                    target_domain in MIDDLE_DOMAINS
                    and target_domain != source_domain
                ):
                    offenders.append(
                        f"{path.relative_to(ROOT)}: "
                        f"{source_domain} -> {target_domain}"
                    )

        self.assertEqual(
            offenders,
            [],
            "Các domain metrics/weekly_model/khsx_ki "
            "không được import lẫn nhau:\n"
            + "\n".join(offenders),
        )

    def test_weekly_engine_is_bottom_layer(self):
        path = PLANNING / "weekly_engine.py"
        imports = list(planning_imports(path))
        self.assertEqual(
            imports,
            [],
            "weekly_engine phải là engine thuần, không import Planning layer khác: "
            + ", ".join(imports),
        )

    def test_pipeline_cannot_depend_on_outer_shells(self):
        path = PLANNING / "pipeline.py"
        forbidden = []

        for imported_module in planning_imports(path):
            target = domain_for_module(imported_module)
            if target in {"publish", "verification"}:
                forbidden.append(imported_module)

        self.assertEqual(
            forbidden,
            [],
            "pipeline không được phụ thuộc ngược outer shell: "
            + ", ".join(forbidden),
        )

    def test_lower_layers_cannot_import_pipeline_or_outer_shells(self):
        lower = {
            "weekly_engine",
            "metrics",
            "weekly_model",
            "khsx_ki",
            "fc",
            "layout",
            "calendar",
            "stock_inputs",
        }
        forbidden_targets = {"pipeline", "publish", "verification"}
        offenders = []

        for path in sorted(PLANNING.rglob("*.py")):
            source = domain_for_module(module_name(path))
            if source not in lower:
                continue
            for imported_module in planning_imports(path):
                target = domain_for_module(imported_module)
                if target in forbidden_targets:
                    offenders.append(
                        f"{path.relative_to(ROOT)}: {source} -> {target}"
                    )

        self.assertEqual(
            offenders,
            [],
            "Lower Planning layers không được phụ thuộc orchestration phía trên:\n"
            + "\n".join(offenders),
        )


if __name__ == "__main__":
    unittest.main()
