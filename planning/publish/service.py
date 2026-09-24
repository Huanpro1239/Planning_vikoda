"""Proposal/publish orchestration with retry and exact-snapshot semantics."""

import hashlib
import time

import stock
from graph_retry import retry_delay_seconds as retry_wait_seconds
from planning import metrics
from planning.pipeline import prepare_pipeline_output
from planning.schedule_report import print_operational_report
from sharepoint.client import is_retryable_graph_error

from .policy import publish_decision
from .proposal import (
    save_proposal_artifacts,
    with_proposal_identity,
)
from .release import (
    build_release_manifest,
    write_release_manifest,
)
from .snapshot import read_snapshot
from .state import (
    detect_input_changes,
    save_publish_decision,
    save_states_after_success,
)


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
    skip_if_unchanged=False,
    force=False,
):
    """Read snapshot, compute proposal, validate policy and optionally publish."""
    if publish_mode not in {None, "proposal", "publish"}:
        raise ValueError(
            "publish_mode phải là 'proposal' hoặc 'publish'."
        )

    for attempt in range(1, max_attempts + 1):
        try:
            (
                target_item,
                target_bytes,
                source_items,
                source_bytes,
                revision,
            ) = read_snapshot(graph, drive_id)

            (
                final_bytes,
                report,
                proposed_runtime_state,
            ) = prepare_pipeline_output(
                target_bytes,
                source_bytes,
                runtime_state=metrics.load_runtime_state(),
                input_revision=revision,
            )
            report = with_proposal_identity(
                report,
                final_bytes,
            )
            print_operational_report(report)

            changed = (
                hashlib.sha256(final_bytes).digest()
                != hashlib.sha256(target_bytes).digest()
            )

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
                    "basis": (
                        "explicit_proposal_mode"
                        if publish_mode == "proposal"
                        else "safe_default"
                    ),
                    "publish_status": report.get(
                        "publish_status"
                    ),
                    "proposal_id": report.get("proposal_id"),
                    "publisher": publisher or None,
                }
                save_proposal_artifacts(
                    final_bytes,
                    report,
                )
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

            old_state = stock.load_state()
            pipeline_info = report.get("pipeline", {})
            changed_inputs = detect_input_changes(
                old_state,
                source_items,
                pipeline_info,
            )

            if (
                skip_if_unchanged
                and not force
                and not changed_inputs
            ):
                print(
                    "[PIPELINE] Không có thay đổi đầu vào "
                    "(5 nguồn tồn kho, sheet Danh_muc, sheet FC đều không đổi)."
                )
                print(
                    "[PIPELINE] Bỏ qua publish để tránh lặp vòng "
                    "(Loop Prevention)."
                )
                decision = {
                    "state": "skipped_unchanged",
                    "basis": "no_input_changes",
                    "publish_status": report.get(
                        "publish_status"
                    ),
                    "proposal_id": report.get("proposal_id"),
                    "publisher": publisher or None,
                    "reason": (
                        "Tất cả nguồn tồn kho và sheet FC/Danh_muc "
                        "không đổi so với state.json."
                    ),
                }
                report["publish_decision"] = decision
                save_proposal_artifacts(
                    final_bytes,
                    report,
                )
                save_publish_decision(decision)
                return {
                    "uploaded": False,
                    "would_change": False,
                    "publish_blocked": False,
                    "publish_mode": "publish",
                    "report": report,
                    "input_revision": revision,
                    "skipped_unchanged": True,
                }

            if changed_inputs:
                print(
                    "[PIPELINE] Phát hiện thay đổi đầu vào: "
                    + ", ".join(changed_inputs)
                    + "."
                )

            authorized, decision = publish_decision(
                report,
                review_approval,
                publisher,
            )
            report["publish_decision"] = decision
            save_proposal_artifacts(final_bytes, report)

            if not authorized:
                print(
                    f"[PIPELINE] PUBLISH BLOCKED "
                    f"proposal={report['proposal_id']}: "
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
                    "[PIPELINE] Authorized snapshot uploaded exactly once "
                    "with If-Match ETag."
                )
            else:
                print(
                    "[PIPELINE] Authorized workbook already matches target; "
                    "no upload needed."
                )

            decision = dict(decision)
            decision["state"] = (
                "published"
                if changed
                else "published_no_change"
            )
            decision["target_etag"] = target_item.get("eTag")
            report["publish_decision"] = decision

            # Validate traceability before any production upload is attempted.
            # The manifest is persisted only after upload/state/decision succeed.
            release_manifest = build_release_manifest(
                report,
                decision,
                final_bytes,
            )

            save_proposal_artifacts(final_bytes, report)
            save_states_after_success(
                source_items,
                report,
                proposed_runtime_state,
            )
            save_publish_decision(decision)
            write_release_manifest(release_manifest)
            return {
                "uploaded": changed,
                "would_change": changed,
                "publish_blocked": False,
                "publish_mode": "publish",
                "report": report,
                "input_revision": revision,
            }

        except Exception as exc:
            if (
                not is_retryable_graph_error(exc)
                or attempt == max_attempts
            ):
                raise
            delay = retry_wait_seconds(
                exc,
                retry_delay_seconds,
            )
            print(
                "[PIPELINE] Graph tạm lỗi/file thay đổi; "
                "bỏ toàn bộ attempt và đọc lại tất cả snapshot "
                f"({attempt}/{max_attempts}) sau {delay:g}s."
            )
            sleep_func(delay)

    raise RuntimeError(
        "Không thể hoàn tất validated planning pipeline."
    )
