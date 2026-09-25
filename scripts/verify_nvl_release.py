"""Verify an NVL release manifest and its available audit evidence.

Examples:
    python -X utf8 scripts/verify_nvl_release.py nvl_release_manifest.json
    python -X utf8 scripts/verify_nvl_release.py runtime-state/nvl/releases/<id>.json
    python -X utf8 scripts/verify_nvl_release.py nvl_release_manifest.json --strict \
        --stock-proposal nvl_stock_proposal.xlsx \
        --final-workbook nvl_open_po_proposal.xlsx \
        --readiness nvl_production_readiness_report.json \
        --stock-report nvl_stock_report.json \
        --open-po-report nvl_open_po_report.json
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from nvl.verify_release import (  # noqa: E402
    NVLReleaseVerificationError,
    verify_nvl_release,
)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "manifest",
        help="Path to NVL release manifest JSON",
    )
    parser.add_argument(
        "--stock-proposal",
        help="Path to nvl_stock_proposal.xlsx",
    )
    parser.add_argument(
        "--final-workbook",
        help="Path to nvl_open_po_proposal.xlsx",
    )
    parser.add_argument(
        "--readiness",
        help="Path to nvl_production_readiness_report.json",
    )
    parser.add_argument(
        "--stock-report",
        help="Path to nvl_stock_report.json",
    )
    parser.add_argument(
        "--open-po-report",
        help="Path to nvl_open_po_report.json",
    )
    parser.add_argument(
        "--repo-root",
        default=str(ROOT),
        help="Git repository root used to verify commit existence",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Require all evidence files and commit existence",
    )
    parser.add_argument(
        "--no-git",
        action="store_true",
        help="Skip local Git commit-history verification",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        dest="json_output",
        help="Print machine-readable JSON result",
    )
    return parser.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    try:
        result = verify_nvl_release(
            args.manifest,
            stock_proposal_path=args.stock_proposal,
            final_workbook_path=args.final_workbook,
            readiness_path=args.readiness,
            stock_report_path=args.stock_report,
            open_po_report_path=args.open_po_report,
            strict=args.strict,
            verify_commit=not args.no_git,
            repo_root=args.repo_root,
        )
    except NVLReleaseVerificationError as exc:
        if args.json_output:
            print(
                json.dumps(
                    {
                        "verified": False,
                        "error": str(exc),
                    },
                    ensure_ascii=False,
                    indent=2,
                )
            )
        else:
            print(
                f"[NVL_RELEASE_VERIFY] FAIL: {exc}",
                file=sys.stderr,
            )
        return 1

    if args.json_output:
        print(
            json.dumps(
                result,
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0

    print(
        f"[NVL_RELEASE_VERIFY] PASS "
        f"release={result['release_id']} "
        f"commit={result['commit_sha']} "
        f"workbook={result['published_workbook_sha256']}"
    )
    for check in result["checks"]:
        print(
            f"  [{check['status'].upper()}] {check['name']}"
            + (
                f": {check['detail']}"
                if check.get("detail")
                else ""
            )
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
