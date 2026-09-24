"""Planning publish state comparison and persistence."""

import json

import stock
from planning import metrics
from planning.schedule_report import save_schedule_report

from .constants import (
    AUDIT_REVISION_FILE,
    PUBLISH_DECISION_FILE,
    SOURCES,
)


def save_publish_decision(decision):
    PUBLISH_DECISION_FILE.write_text(
        json.dumps(
            decision,
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )
        + "\n",
        encoding="utf-8",
    )


def save_states_after_success(
    source_items,
    report,
    proposed_runtime_state,
):
    source_etags = {
        key: source_items[key].get("eTag")
        for key in SOURCES
    }
    pipeline_info = report.get("pipeline") or {}
    conversion_hash = (
        pipeline_info.get("conversion_hash")
        or pipeline_info.get("danh_muc_hash")
    )
    if not conversion_hash:
        raise RuntimeError(
            "Pipeline report thiếu conversion_hash; không ghi state."
        )

    stock.save_state(
        source_etags,
        conversion_hash,
        fc_hash=pipeline_info.get("fc_hash"),
        no_kho_hash=pipeline_info.get("no_kho_hash"),
        planning_inputs_hash=pipeline_info.get(
            "planning_inputs_hash"
        ),
        engine_version=(
            pipeline_info.get("engine_version")
            or pipeline_info.get("engine")
        ),
    )
    metrics.save_runtime_state(proposed_runtime_state)
    save_schedule_report(report)
    AUDIT_REVISION_FILE.write_text(
        json.dumps(
            report.get("input_revision") or {},
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )


def detect_input_changes(old_state, source_items, pipeline_info):
    """Detect all source/master/planning changes relative to state.json."""
    old_sources = (
        old_state.get("sources", {})
        if isinstance(old_state, dict)
        else {}
    )
    changed_inputs = []

    for key in SOURCES:
        old_etag = old_sources.get(key)
        new_etag = source_items.get(key, {}).get("eTag")
        if old_etag != new_etag:
            changed_inputs.append(f"source:{key}")

    old_conv = (
        old_state.get("conversion_hash")
        if isinstance(old_state, dict)
        else None
    )
    new_conv = (
        pipeline_info.get("conversion_hash")
        or pipeline_info.get("danh_muc_hash")
    )
    if old_conv != new_conv:
        changed_inputs.append("sheet:Danh_muc")

    old_fc = (
        old_state.get("fc_hash")
        if isinstance(old_state, dict)
        else None
    )
    if old_fc != pipeline_info.get("fc_hash"):
        changed_inputs.append("sheet:FC")

    old_no_kho = (
        old_state.get("no_kho_hash")
        if isinstance(old_state, dict)
        else None
    )
    new_no_kho = pipeline_info.get("no_kho_hash")
    if (
        new_no_kho is not None
        and old_no_kho != new_no_kho
    ):
        changed_inputs.append("sheet:No_kho")

    old_planning = (
        old_state.get("planning_inputs_hash")
        if isinstance(old_state, dict)
        else None
    )
    new_planning = pipeline_info.get("planning_inputs_hash")
    if (
        new_planning is not None
        and old_planning != new_planning
    ):
        changed_inputs.append("sheet:Ke_hoach_SX")

    old_engine = (
        old_state.get("engine_version")
        if isinstance(old_state, dict)
        else None
    )
    new_engine = (
        pipeline_info.get("engine_version")
        or pipeline_info.get("engine")
    )
    if (
        new_engine is not None
        and old_engine != new_engine
    ):
        changed_inputs.append("engine_version")

    return changed_inputs
