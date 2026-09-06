import argparse
import hashlib
import json
import os
import time
from pathlib import Path

from graph_retry import install_retry_after_support, retry_delay_seconds as retry_wait_seconds
from planning_pipeline import prepare_pipeline_output
from planning_schedule_report import print_operational_report, save_schedule_report
import sync_planning_metrics as metrics
import sync_stock
from sync_stock import GraphClient, get_access_token, is_retryable_graph_error


AUDIT_REVISION_FILE = Path("planning_input_revision.json")
PROPOSAL_WORKBOOK_FILE = Path("planning_proposal.xlsx")
PUBLISH_DECISION_FILE = Path("planning_publish_decision.json")
READY_PUBLISH_STATUSES = {"ready_for_publish", "feasible"}  # feasible = legacy tests/reports
REVIEW_REQUIRED_STATUS = "review_required"
SOURCES = {
    "actual_stock": sync_stock.SOURCE_ACTUAL_PATH,
    "factory_vikoda": sync_stock.SOURCE_FACTORY_VIKODA_PATH,
    "factory_vkd": sync_stock.SOURCE_FACTORY_VKD_PATH,
    "accounting_vikoda": sync_stock.SOURCE_ACCOUNTING_VIKODA_PATH,
    "accounting_vkd": sync_stock.SOURCE_ACCOUNTING_VKD_PATH,
}


def _read_snapshot(graph, drive_id):
    target_item = graph.get_item_by_path(drive_id, sync_stock.DEST_PATH)
    target_bytes = graph.download_file(drive_id, target_item["id"])

    source_bytes = {}
    source_items = {}
    for key, path in SOURCES.items():
        item = graph.get_item_by_path(drive_id, path)
        data = graph.download_file(drive_id, item["id"])
        source_items[key] = item
        source_bytes[key] = data

    revision = {
        "target": {
            "path": sync_stock.DEST_PATH,
            "item_id": target_item.get("id"),
            "etag": target_item.get("eTag"),
            "last_modified": target_item.get("lastModifiedDateTime"),
            "sha256": hashlib.sha256(target_bytes).hexdigest(),
        },
        "sources": {
            key: {
                "path": SOURCES[key],
                "item_id": source_items[key].get("id"),
                "etag": source_items[key].get("eTag"),
                "last_modified": source_items[key].get("lastModifiedDateTime"),
                "sha256": hashlib.sha256(source_bytes[key]).hexdigest(),
            }
            for key in SOURCES
        },
    }
    return target_item, target_bytes, source_items, source_bytes, revision


def _proposal_id(report, final_bytes):
    output_sha256 = hashlib.sha256(final_bytes).hexdigest()
    payload = {
        "algorithm": report.get("algorithm"),
        "plan_month": report.get("plan_month"),
        "input_revision": report.get("input_revision") or {},
        "output_sha256": output_sha256,
    }
    canonical = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(canonical).hexdigest(), output_sha256


def _with_proposal_identity(report, final_bytes):
    result = dict(report)
    proposal_id, output_sha256 = _proposal_id(result, final_bytes)
    result["output_sha256"] = output_sha256
    result["proposal_id"] = proposal_id
    return result


def _save_proposal_artifacts(final_bytes, report):
    PROPOSAL_WORKBOOK_FILE.write_bytes(final_bytes)
    save_schedule_report(report)
    AUDIT_REVISION_FILE.write_text(
        json.dumps(report.get("input_revision") or {}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def _save_publish_decision(decision):
    PUBLISH_DECISION_FILE.write_text(
        json.dumps(decision, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def _save_states_after_success(source_items, report, proposed_runtime_state):
    source_etags = {
        key: source_items[key].get("eTag")
        for key in SOURCES
    }
    conversion_hash = report.get("pipeline", {}).get("conversion_hash")
    if not conversion_hash:
        raise RuntimeError("Pipeline report thiếu conversion_hash; không ghi state.")

    sync_stock.save_state(source_etags, conversion_hash)
    metrics.save_runtime_state(proposed_runtime_state)
    save_schedule_report(report)
    AUDIT_REVISION_FILE.write_text(
        json.dumps(report.get("input_revision") or {}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def _blocked_decision(report, publisher, basis, reason):
    return {
        "state": "blocked",
        "basis": basis,
        "publish_status": report.get("publish_status"),
        "proposal_id": report.get("proposal_id"),
        "publisher": publisher or None,
        "reason": reason,
    }


def _publish_decision(report, review_approval, publisher):
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
        return False, _blocked_decision(
            report,
            publisher,
            "unknown_publish_status",
            "Publish status không thuộc chính sách cho phép.",
        )

    sections = report.get("status") or {}
    monthly = sections.get("monthly_quantity") or {}
    resource = sections.get("resource_validation") or {}
    service = sections.get("service") or {}

    # Review approval is intentionally narrow: it may accept a validated service
    # shortage, but it never waives monthly mass-balance/carryover or physical
    # resource validation. Hard workbook validation has already passed earlier.
    if not bool(monthly.get("ok")) or not bool(resource.get("ok")):
        return False, _blocked_decision(
            report,
            publisher,
            "non_waivable_validation",
            "review_required do carryover/resource validation không được phép override.",
        )
    if bool(service.get("ok")) or service.get("state") != "stockout_risk":
        return False, _blocked_decision(
            report,
            publisher,
            "unsupported_review_reason",
            "Chỉ stockout_risk đã validate mới có thể dùng review approval.",
        )

    approval = dict(review_approval or {})
    approved_proposal_id = str(approval.get("proposal_id") or "").strip()
    reason = str(approval.get("reason") or "").strip()
    approved_by = str(approval.get("approved_by") or publisher or "").strip()

    if not approved_proposal_id or approved_proposal_id != proposal_id:
        return False, _blocked_decision(
            report,
            publisher,
            "review_approval",
            "review_required cần approval đúng proposal_id của snapshot hiện tại.",
        )
    if not reason:
        return False, _blocked_decision(
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


def run_pipeline_with_retry(
    graph,
    drive_id,
    *,
    publish_mode=None,
    review_approval=None,
    publisher=None,
    sleep_func=time.sleep,
    max_attempts=6,
    retry_delay_seconds=10,
):
    """
    Read all inputs -> compute entire workbook -> validate -> optional ETag upload.

    For current report schema, omitted publish_mode defaults to proposal/read-only.
    The obsolete `feasible` status keeps implicit publish behavior only for
    backward-compatible legacy callers/tests; production reports now emit
    `ready_for_publish` or `review_required`.

    publish_mode="proposal" is read-only for SharePoint and runtime state: the
    computed workbook/report are saved only as local review artifacts.

    publish_mode="publish" writes only when publish policy authorizes the exact
    computed proposal. `review_required` may override only validated
    `stockout_risk`, and needs an approval bound to proposal_id plus a non-empty
    reason. Any 412/transient Graph error restarts from fresh target + source
    snapshots; an approval for an older proposal_id is therefore not reused.
    """
    if publish_mode not in {None, "proposal", "publish"}:
        raise ValueError("publish_mode phải là 'proposal' hoặc 'publish'.")

    for attempt in range(1, max_attempts + 1):
        try:
            (
                target_item,
                target_bytes,
                source_items,
                source_bytes,
                revision,
            ) = _read_snapshot(graph, drive_id)

            final_bytes, report, proposed_runtime_state = prepare_pipeline_output(
                target_bytes,
                source_bytes,
                runtime_state=metrics.load_runtime_state(),
                input_revision=revision,
            )
            report = _with_proposal_identity(report, final_bytes)
            # prepare_pipeline_output already validates the exact final bytes.
            print_operational_report(report)

            changed = hashlib.sha256(final_bytes).digest() != hashlib.sha256(target_bytes).digest()
            effective_mode = publish_mode
            if effective_mode is None:
                effective_mode = (
                    "publish"
                    if report.get("publish_status") == "feasible"
                    else "proposal"
                )

            if effective_mode == "proposal":
                report["publish_decision"] = {
                    "state": "proposal_only",
                    "basis": "explicit_proposal_mode" if publish_mode == "proposal" else "safe_default",
                    "publish_status": report.get("publish_status"),
                    "proposal_id": report.get("proposal_id"),
                    "publisher": publisher or None,
                }
                _save_proposal_artifacts(final_bytes, report)
                print(
                    f"[PIPELINE] Proposal {report['proposal_id']} đã tạo; "
                    "không upload SharePoint và không lưu runtime state."
                )
                return {
                    "uploaded": False,
                    "would_change": changed,
                    "publish_blocked": False,
                    "publish_mode": "proposal",
                    "report": report,
                    "input_revision": revision,
                }

            authorized, decision = _publish_decision(
                report,
                review_approval,
                publisher,
            )
            report["publish_decision"] = decision
            _save_proposal_artifacts(final_bytes, report)

            if not authorized:
                print(
                    f"[PIPELINE] PUBLISH BLOCKED proposal={report['proposal_id']}: "
                    f"{decision.get('reason')}"
                )
                return {
                    "uploaded": False,
                    "would_change": changed,
                    "publish_blocked": True,
                    "publish_mode": "publish",
                    "report": report,
                    "input_revision": revision,
                }

            if changed:
                graph.upload_file(
                    drive_id,
                    target_item["id"],
                    final_bytes,
                    expected_etag=target_item["eTag"],
                )
                print(
                    "[PIPELINE] Authorized snapshot uploaded exactly once with If-Match ETag."
                )
            else:
                print("[PIPELINE] Authorized workbook already matches target; no upload needed.")

            decision = dict(decision)
            decision["state"] = "published" if changed else "published_no_change"
            decision["target_etag"] = target_item.get("eTag")
            report["publish_decision"] = decision
            _save_proposal_artifacts(final_bytes, report)
            _save_states_after_success(
                source_items,
                report,
                proposed_runtime_state,
            )
            _save_publish_decision(decision)
            return {
                "uploaded": changed,
                "would_change": changed,
                "publish_blocked": False,
                "publish_mode": "publish",
                "report": report,
                "input_revision": revision,
            }
        except Exception as exc:
            if not is_retryable_graph_error(exc) or attempt == max_attempts:
                raise
            delay = retry_wait_seconds(exc, retry_delay_seconds)
            print(
                f"[PIPELINE] Graph tạm lỗi/file thay đổi; bỏ toàn bộ attempt và "
                f"đọc lại tất cả snapshot ({attempt}/{max_attempts}) sau {delay:g}s."
            )
            sleep_func(delay)

    raise RuntimeError("Không thể hoàn tất validated planning pipeline.")


def _parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Tạo proposal planning read-only hoặc publish có kiểm soát."
    )
    parser.add_argument(
        "--publish",
        action="store_true",
        help="Cho phép publish theo policy. Mặc định chỉ tạo proposal read-only.",
    )
    parser.add_argument("--approval-proposal-id", default="")
    parser.add_argument("--approval-reason", default="")
    parser.add_argument("--approved-by", default=os.getenv("GITHUB_ACTOR", ""))
    return parser.parse_args(argv)


def main(argv=None):
    args = _parse_args(argv)
    install_retry_after_support()
    token = get_access_token()
    graph = GraphClient(token)
    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    review_approval = None
    if args.approval_proposal_id or args.approval_reason:
        review_approval = {
            "proposal_id": args.approval_proposal_id,
            "reason": args.approval_reason,
            "approved_by": args.approved_by,
        }

    result = run_pipeline_with_retry(
        graph,
        drive_id,
        publish_mode="publish" if args.publish else "proposal",
        review_approval=review_approval,
        publisher=args.approved_by,
    )
    if result.get("publish_blocked"):
        raise SystemExit(2)
    return result


if __name__ == "__main__":
    main()
