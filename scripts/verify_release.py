"""Verify a historical Planning release manifest and its evidence."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from planning.publish.release_verify import verify_release


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", help="Path to release manifest JSON")
    parser.add_argument("--proposal", help="Proposal workbook used by the release")
    parser.add_argument("--readiness", help="production_readiness_report.json")
    parser.add_argument("--report", help="planning_schedule_report.json")
    parser.add_argument("--revision", help="planning_input_revision.json")
    parser.add_argument("--decision", help="planning_publish_decision.json")
    parser.add_argument("--repo-root", help="Git repo root used to verify commit ancestry")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Fail when any evidence is missing, including proposal workbook/commit objects.",
    )
    parser.add_argument(
        "--out",
        help="Optional JSON verification report path.",
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    result = verify_release(
        args.manifest,
        proposal_path=args.proposal,
        readiness_path=args.readiness,
        report_path=args.report,
        revision_path=args.revision,
        decision_path=args.decision,
        repo_root=args.repo_root,
        strict=args.strict,
    )

    for item in result["checks"]:
        print(
            f"[VERIFY_RELEASE][{item['status'].upper()}] "
            f"{item['name']}: {item['detail']}"
        )

    summary = result["summary"]
    print(
        "[VERIFY_RELEASE] "
        f"status={result['verification_status']}; "
        f"passed={summary['passed']}; "
        f"failed={summary['failed']}; "
        f"skipped={summary['skipped']}"
    )

    if args.out:
        output = Path(args.out)
        output.write_text(
            json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )
        print(f"[VERIFY_RELEASE] report={output}")

    return 1 if result["verification_status"] == "failed" else 0


if __name__ == "__main__":
    raise SystemExit(main())
