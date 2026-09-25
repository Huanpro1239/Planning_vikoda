"""Audit and index the NVL release ledger.

Examples:
    python -X utf8 scripts/audit_nvl_releases.py runtime-state/nvl/releases
    python -X utf8 scripts/audit_nvl_releases.py .runtime-state/nvl/releases --no-git
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from nvl.release_audit import (  # noqa: E402
    audit_nvl_release_ledger,
    write_index_csv,
    write_index_json,
)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "releases_dir",
        nargs="?",
        default="runtime-state/nvl/releases",
        help="Directory chứa immutable NVL release manifests",
    )
    parser.add_argument(
        "--latest",
        help=(
            "Path tới latest_release.json; mặc định dùng sibling "
            "../latest_release.json"
        ),
    )
    parser.add_argument(
        "--out-json",
        default="nvl_release_index.json",
        help="Output JSON index",
    )
    parser.add_argument(
        "--out-csv",
        default="nvl_release_index.csv",
        help="Output CSV history table",
    )
    parser.add_argument(
        "--repo-root",
        default=str(ROOT),
        help="Git repository root để kiểm commit existence",
    )
    parser.add_argument(
        "--no-git",
        action="store_true",
        help="Bỏ qua local Git commit-history verification",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        dest="json_output",
        help="In full index JSON ra stdout thay vì bảng ngắn",
    )
    return parser.parse_args(argv)


def _print_table(index):
    print(
        "SEQ | STATUS  | CHAIN           | PUBLISHED_AT              "
        "| RELEASE_ID | COMMIT"
    )
    print("-" * 112)
    for row in index.get("releases", []):
        release_id = str(row.get("release_id") or "")
        commit = str(row.get("commit_sha") or "")
        print(
            f"{row.get('sequence', ''):>3} | "
            f"{str(row.get('status') or ''):<7} | "
            f"{str(row.get('chain_status') or ''):<15} | "
            f"{str(row.get('published_at') or ''):<25} | "
            f"{release_id[:40]:<40} | "
            f"{commit[:12]}"
        )
    summary = index.get("summary") or {}
    latest = index.get("latest") or {}
    print(
        f"\n[NVL-LEDGER] status={index.get('status')} "
        f"releases={summary.get('release_count', 0)} "
        f"failed={summary.get('failed_count', 0)} "
        f"warnings={summary.get('warning_count', 0)} "
        f"latest={latest.get('status')}"
    )


def main(argv=None) -> int:
    args = parse_args(argv)
    index = audit_nvl_release_ledger(
        args.releases_dir,
        latest_path=args.latest,
        verify_commit=not args.no_git,
        repo_root=args.repo_root,
    )
    write_index_json(index, args.out_json)
    write_index_csv(index, args.out_csv)

    if args.json_output:
        print(json.dumps(index, ensure_ascii=False, indent=2))
    else:
        _print_table(index)
        print(
            f"[NVL-LEDGER] JSON={args.out_json} CSV={args.out_csv}"
        )

    return 1 if index.get("status") == "failed" else 0


if __name__ == "__main__":
    raise SystemExit(main())
