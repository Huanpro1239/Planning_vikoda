"""Verify a Planning release manifest and its available audit evidence.

Examples:
    python -X utf8 scripts/verify_release.py planning_release_manifest.json
    python -X utf8 scripts/verify_release.py runtime-state/releases/<id>.json
    python -X utf8 scripts/verify_release.py manifest.json --strict \
        --proposal planning_proposal.xlsx \
        --readiness production_readiness_report.json \
        --report planning_schedule_report.json \
        --decision planning_publish_decision.json
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from planning.publish.verify_release import (  # noqa: E402
    ReleaseVerificationError,
    verify_release,
)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", help="Path to planning release manifest JSON")
    parser.add_argument("--proposal", help="Path to planning_proposal.xlsx")
    parser.add_argument("--readiness", help="Path to production_readiness_report.json")
    parser.add_argument("--report", help="Path to planning_schedule_report.json")
    parser.add_argument("--decision", help="Path to planning_publish_decision.json")
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
        result = verify_release(
            args.manifest,
            proposal_path=args.proposal,
            readiness_path=args.readiness,
            report_path=args.report,
            decision_path=args.decision,
            strict=args.strict,
            verify_commit=not args.no_git,
            repo_root=args.repo_root,
        )
    except ReleaseVerificationError as exc:
        if args.json_output:
            print(json.dumps({
                "verified": False,
                "error": str(exc),
            }, ensure_ascii=False, indent=2))
        else:
            print(f"[RELEASE_VERIFY] FAIL: {exc}", file=sys.stderr)
        return 1

    if args.json_output:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    print(
        f"[RELEASE_VERIFY] PASS release={result['release_id']} "
        f"proposal={result['proposal_id']} commit={result['commit_sha']}"
    )
    for check in result["checks"]:
        print(
            f"  [{check['status'].upper()}] {check['name']}"
            + (f": {check['detail']}" if check.get("detail") else "")
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
