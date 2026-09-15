"""Run a legacy entrypoint through the canonical SharePoint auth provider.

This is a migration shim: many existing modules do ``from sync_stock import
get_access_token``. The runner replaces that function *before* importing the
target module, so legacy imports receive ``sharepoint.auth.get_access_token``
without a risky whole-file rewrite of large production modules.

Usage:
    python -X utf8 scripts/auth_runner.py sync_planning_pipeline --publish ...
    python -X utf8 scripts/auth_runner.py sync_nvl_stock --config ...
    python -X utf8 scripts/auth_runner.py scripts.survey_nvl_sharepoint --config ...
"""

from __future__ import annotations

import importlib
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def run_module(module_name: str, argv: list[str] | None = None):
    import sync_stock
    from sharepoint.auth import get_access_token

    # Patch before importing the target so ``from sync_stock import get_access_token``
    # inside the target captures the canonical provider.
    original_provider = sync_stock.get_access_token
    sync_stock.get_access_token = get_access_token

    old_argv = sys.argv
    try:
        module = importlib.import_module(module_name)
        main = getattr(module, "main", None)
        if not callable(main):
            raise RuntimeError(f"Module {module_name!r} không có hàm main() callable.")

        sys.argv = [module_name, *(argv or [])]
        return main()
    finally:
        sys.argv = old_argv
        # Quan trọng cho test/in-process tooling. Các target đã import provider mới
        # vào namespace riêng; production CLI cũng kết thúc ngay sau main().
        sync_stock.get_access_token = original_provider


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python scripts/auth_runner.py <module> [args...]", file=sys.stderr)
        return 2

    module_name = sys.argv[1]
    args = sys.argv[2:]
    result = run_module(module_name, args)
    return int(result) if isinstance(result, int) else 0


if __name__ == "__main__":
    raise SystemExit(main())
