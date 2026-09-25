"""Build a unified Planning + NVL release overview from runtime-state."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ops.release_overview import (  # noqa: E402
    build_release_overview,
    write_overview_csv,
    write_overview_json,
    write_overview_markdown,
)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "runtime_state",
        nargs="?",
        default=".runtime-state",
        help="Checkout/path của runtime-state branch",
    )
    parser.add_argument("--out-json", default="release_overview.json")
    parser.add_argument("--out-csv", default="release_overview.csv")
    parser.add_argument("--out-md", default="release_overview.md")
    parser.add_argument("--repo-root", default=str(ROOT))
    parser.add_argument(
        "--no-git",
        action="store_true",
        help="Bỏ qua local Git commit-history verification",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        dest="json_output",
        help="In full JSON ra stdout",
    )
    args = parser.parse_args(argv)

    overview = build_release_overview(
        args.runtime_state,
        verify_commit=not args.no_git,
        repo_root=args.repo_root,
    )
    write_overview_json(overview, args.out_json)
    write_overview_csv(overview, args.out_csv)
    write_overview_markdown(overview, args.out_md)

    if args.json_output:
        print(json.dumps(overview, ensure_ascii=False, indent=2))
    else:
        summary = overview["summary"]
        print(
            f"[RELEASE-OVERVIEW] status={overview['status']} "
            f"planning={summary['planning_release_count']} "
            f"nvl={summary['nvl_release_count']} "
            f"total={summary['total_release_count']}"
        )
        print(
            f"[RELEASE-OVERVIEW] JSON={args.out_json} "
            f"CSV={args.out_csv} MD={args.out_md}"
        )

    return 1 if overview["status"] == "failed" else 0


if __name__ == "__main__":
    raise SystemExit(main())
