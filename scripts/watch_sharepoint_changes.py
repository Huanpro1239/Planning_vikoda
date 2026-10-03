"""Detect material Planning input changes in SharePoint.

This watcher is read-only. It compares the same production input
fingerprints used by Planning publish against runtime-state/state.json.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import stock  # noqa: E402
from planning import metrics  # noqa: E402
from planning.fc import compute_fc_hash  # noqa: E402
from planning.publish.constants import SOURCES  # noqa: E402
from planning.publish.state import detect_input_changes  # noqa: E402
from planning.weekly_model import (  # noqa: E402
    ENGINE_VERSION,
    compute_planning_inputs_hash,
)
from sharepoint.client import GraphClient, get_access_token  # noqa: E402


def collect_current_inputs(graph, drive_id):
    """Read only the metadata/workbook bytes needed for change detection."""
    source_items = {
        key: graph.get_item_by_path(drive_id, path)
        for key, path in SOURCES.items()
    }

    target_item = graph.get_item_by_path(drive_id, stock.DEST_PATH)
    target_bytes = graph.download_file(drive_id, target_item["id"])

    _, conversion_hash = stock.read_conversion_factors(target_bytes)
    _, no_kho_hash = metrics.hash_debt_sheet(target_bytes)

    pipeline_info = {
        "conversion_hash": conversion_hash,
        "danh_muc_hash": conversion_hash,
        "fc_hash": compute_fc_hash(target_bytes),
        "no_kho_hash": no_kho_hash,
        "planning_inputs_hash": compute_planning_inputs_hash(target_bytes),
        "engine_version": ENGINE_VERSION,
        "engine": ENGINE_VERSION,
    }
    return source_items, pipeline_info


def detect_sharepoint_changes(graph, drive_id, *, state_path="state.json"):
    """Return the exact production inputs changed since the last release."""
    old_state = stock.load_state(path=state_path)
    source_items, pipeline_info = collect_current_inputs(graph, drive_id)
    return detect_input_changes(old_state, source_items, pipeline_info)


def emit_outputs(changed_inputs):
    changed = bool(changed_inputs)
    reasons_json = json.dumps(
        changed_inputs,
        ensure_ascii=False,
        separators=(",", ":"),
    )

    print(
        "[WATCH] "
        + (
            "Phát hiện thay đổi đầu vào: " + ", ".join(changed_inputs)
            if changed
            else "Không có thay đổi đầu vào production."
        )
    )

    output_path = os.getenv("GITHUB_OUTPUT")
    if output_path:
        with open(output_path, "a", encoding="utf-8") as handle:
            handle.write(f"changed={'true' if changed else 'false'}\n")
            handle.write(f"reasons_json={reasons_json}\n")

    return changed


def main():
    token = get_access_token()
    graph = GraphClient(token)
    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    changed_inputs = detect_sharepoint_changes(graph, drive_id)
    emit_outputs(changed_inputs)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
