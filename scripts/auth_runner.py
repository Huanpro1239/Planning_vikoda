"""Run a repository entrypoint without mutating authentication state.

Authentication is resolved by canonical callers through sharepoint.auth /
sharepoint.client. This runner only preserves the existing workflow invocation
shape while modules migrate away from root entrypoints.
"""

from __future__ import annotations

import importlib
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def run_module(module_name: str, argv: list[str] | None = None):
    module = importlib.import_module(module_name)
    main = getattr(module, "main", None)
    if not callable(main):
        raise RuntimeError(f"Module {module_name!r} không có hàm main() callable.")

    old_argv = sys.argv
    try:
        sys.argv = [module_name, *(argv or [])]
        return main()
    finally:
        sys.argv = old_argv


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python scripts/auth_runner.py <module> [args...]", file=sys.stderr)
        return 2

    result = run_module(sys.argv[1], sys.argv[2:])
    return int(result) if isinstance(result, int) else 0


if __name__ == "__main__":
    raise SystemExit(main())
