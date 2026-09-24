"""CLI/auth boundary for Planning publish."""

import argparse
import os

from sharepoint.retry import install_retry_after_support
from sharepoint.client import GraphClient, get_access_token

from .service import run_pipeline_with_retry


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description=(
            "Tạo proposal planning read-only hoặc publish có kiểm soát."
        )
    )
    parser.add_argument(
        "--publish",
        action="store_true",
        help=(
            "Cho phép publish theo policy. "
            "Mặc định chỉ tạo proposal read-only."
        ),
    )
    parser.add_argument(
        "--skip-if-unchanged",
        action="store_true",
        help=(
            "Bỏ qua publish nếu không có thay đổi ở nguồn tồn kho, "
            "sheet Danh_muc hoặc sheet FC."
        ),
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Bắt buộc publish kể cả khi không có thay đổi đầu vào.",
    )
    parser.add_argument(
        "--approval-proposal-id",
        default="",
    )
    parser.add_argument(
        "--approval-reason",
        default="",
    )
    parser.add_argument(
        "--approved-by",
        default=os.getenv("GITHUB_ACTOR", ""),
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    install_retry_after_support()

    token = get_access_token()
    graph = GraphClient(token)
    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    review_approval = None
    if (
        args.approval_proposal_id
        or args.approval_reason
    ):
        review_approval = {
            "proposal_id": args.approval_proposal_id,
            "reason": args.approval_reason,
            "approved_by": args.approved_by,
        }

    result = run_pipeline_with_retry(
        graph,
        drive_id,
        publish_mode=(
            "publish"
            if args.publish
            else "proposal"
        ),
        review_approval=review_approval,
        publisher=args.approved_by,
        skip_if_unchanged=args.skip_if_unchanged,
        force=args.force,
    )
    if result.get("publish_blocked"):
        raise SystemExit(2)
    return result
