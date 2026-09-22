"""Publish authorization policy.

Only validated stockout risk can be explicitly approved. Resource failures and
carryover remain non-waivable.
"""

from .constants import READY_PUBLISH_STATUSES, REVIEW_REQUIRED_STATUS


def blocked_decision(report, publisher, basis, reason):
    return {
        "state": "blocked",
        "basis": basis,
        "publish_status": report.get("publish_status"),
        "proposal_id": report.get("proposal_id"),
        "publisher": publisher or None,
        "reason": reason,
    }


def publish_decision(report, review_approval, publisher):
    status = str(report.get("publish_status") or "").strip()
    proposal_id = report.get("proposal_id")

    if status in READY_PUBLISH_STATUSES:
        return True, {
            "state": "authorized",
            "basis": "publish_status",
            "publish_status": status,
            "proposal_id": proposal_id,
            "publisher": publisher or None,
            "review_approval": None,
        }

    if status != REVIEW_REQUIRED_STATUS:
        return False, blocked_decision(
            report,
            publisher,
            "unknown_publish_status",
            "Publish status không thuộc chính sách cho phép.",
        )

    sections = report.get("status") or {}
    monthly = sections.get("monthly_quantity") or {}
    resource = sections.get("resource_validation") or {}
    service = sections.get("service") or {}

    if not bool(resource.get("ok")):
        return False, blocked_decision(
            report,
            publisher,
            "non_waivable_validation",
            "review_required do resource validation không được phép override.",
        )

    if (
        not bool(monthly.get("ok"))
        and monthly.get("state") == "carryover"
    ):
        return False, blocked_decision(
            report,
            publisher,
            "non_waivable_validation",
            "review_required do carryover/resource validation không được phép override.",
        )

    if bool(service.get("ok")) or service.get("state") != "stockout_risk":
        return False, blocked_decision(
            report,
            publisher,
            "unsupported_review_reason",
            "Chỉ stockout_risk đã validate mới có thể dùng review approval.",
        )

    approval = dict(review_approval or {})
    approved_proposal_id = str(
        approval.get("proposal_id") or ""
    ).strip()
    reason = str(approval.get("reason") or "").strip()
    approved_by = str(
        approval.get("approved_by") or publisher or ""
    ).strip()

    if not approved_proposal_id or approved_proposal_id != proposal_id:
        return False, blocked_decision(
            report,
            publisher,
            "review_approval",
            "review_required cần approval đúng proposal_id của snapshot hiện tại.",
        )

    if not reason:
        return False, blocked_decision(
            report,
            publisher,
            "review_approval",
            "review_required cần lý do chấp nhận rủi ro/thiếu hàng.",
        )

    return True, {
        "state": "authorized",
        "basis": "review_approval",
        "publish_status": status,
        "proposal_id": proposal_id,
        "publisher": publisher or approved_by or None,
        "review_approval": {
            "proposal_id": approved_proposal_id,
            "reason": reason,
            "approved_by": approved_by or None,
        },
    }
