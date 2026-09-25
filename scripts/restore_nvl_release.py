"""Create or publish a controlled NVL recovery from a historical release.

Examples:
    python -X utf8 scripts/restore_nvl_release.py <release_id> \
        --historical-workbook nvl_open_po_proposal.xlsx

    python -X utf8 scripts/restore_nvl_release.py <release_id> \
        --historical-workbook nvl_open_po_proposal.xlsx \
        --publish --approve-release-id <release_id>
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from nvl.recovery import run_recovery  # noqa: E402


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "release",
        help="Release ID hoặc path tới NVL release manifest",
    )
    parser.add_argument(
        "--historical-workbook",
        required=True,
        help="Historical nvl_open_po_proposal.xlsx của release cần phục hồi",
    )
    parser.add_argument(
        "--config",
        default="nvl_stock_config.json",
        help="NVL target config",
    )
    parser.add_argument(
        "--releases-dir",
        default="runtime-state/nvl/releases",
    )
    parser.add_argument(
        "--latest",
        help="Path tới runtime-state/nvl/latest_release.json",
    )
    parser.add_argument(
        "--out-dir",
        default=".",
    )
    parser.add_argument(
        "--publish",
        action="store_true",
        help="Cho phép upload recovery proposal sau các guard",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Explicit dry-run; đây cũng là mode mặc định",
    )
    parser.add_argument(
        "--approve-release-id",
        help="Bắt buộc và phải khớp release ID khi dùng --publish",
    )
    parser.add_argument(
        "--verify-git",
        action="store_true",
        help="Kiểm commit history local khi audit release ledger",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        dest="json_output",
    )
    return parser.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    if args.publish and args.dry_run:
        print(
            "[NVL-RECOVERY] FAIL: không dùng đồng thời --publish và --dry-run",
            file=sys.stderr,
        )
        return 2

    try:
        report = run_recovery(
            args.release,
            historical_workbook=args.historical_workbook,
            config_path=args.config,
            releases_dir=args.releases_dir,
            latest_path=args.latest,
            out_dir=args.out_dir,
            publish=args.publish,
            approve_release_id=args.approve_release_id,
            verify_git=args.verify_git,
        )
    except Exception as exc:
        if args.json_output:
            print(
                json.dumps(
                    {"ok": False, "error": str(exc)},
                    ensure_ascii=False,
                    indent=2,
                )
            )
        else:
            print(f"[NVL-RECOVERY] FAIL: {exc}", file=sys.stderr)
        return 1

    if args.json_output:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        proposal = report.get("proposal") or {}
        print(
            f"[NVL-RECOVERY] {report.get('status')} "
            f"release={report.get('release_id')} "
            f"changed={proposal.get('changed_cells')} "
            f"D={proposal.get('changed_D')} E={proposal.get('changed_E')}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
