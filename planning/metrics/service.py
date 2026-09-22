"""Standalone SharePoint service for Planning metrics."""

import time

from sharepoint.client import (
    GraphClient,
    get_access_token,
    is_retryable_graph_error,
)
from sync_stock import (
    DEST_PATH,
    SOURCE_ACTUAL_PATH,
    SOURCE_FACTORY_VIKODA_PATH,
)

from .calculation import calculate_metrics
from .constants import FC_SELECTOR_CELL, FC_SHEET, PLANNING_SHEET
from .readers import (
    read_actual_inputs,
    read_conversion_factors_and_leadtime,
    read_planning_rows,
    read_system_receipts,
)
from .state import load_runtime_state, save_runtime_state
from .workbook import count_changes, patch_workbook


def main():
    token = get_access_token()
    graph = GraphClient(token)

    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    dest_item = graph.get_item_by_path(drive_id, DEST_PATH)
    actual_item = graph.get_item_by_path(drive_id, SOURCE_ACTUAL_PATH)
    system_item = graph.get_item_by_path(
        drive_id,
        SOURCE_FACTORY_VIKODA_PATH,
    )

    dest_bytes = graph.download_file(drive_id, dest_item["id"])
    actual_bytes = graph.download_file(drive_id, actual_item["id"])
    system_bytes = graph.download_file(drive_id, system_item["id"])

    selector, plan_month, planning_rows = read_planning_rows(dest_bytes)
    conversion_factors, _, leadtimes = (
        read_conversion_factors_and_leadtime(dest_bytes)
    )
    report_date, actual_receipts, consignments = read_actual_inputs(
        actual_bytes
    )
    system_receipts = read_system_receipts(
        system_bytes,
        conversion_factors,
        set(planning_rows),
    )

    state = load_runtime_state()
    metric_values, state_changed, source_key, plan_key = calculate_metrics(
        report_date=report_date,
        plan_month=plan_month,
        planning_rows=planning_rows,
        actual_receipts=actual_receipts,
        system_receipts=system_receipts,
        current_consignments=consignments,
        state=state,
        leadtimes=leadtimes,
    )

    changed_count = count_changes(planning_rows, metric_values)
    if changed_count:
        updated_bytes, patched = patch_workbook(
            dest_bytes,
            metric_values,
        )
        result = graph.upload_file(
            drive_id,
            dest_item["id"],
            updated_bytes,
            expected_etag=dest_item["eTag"],
        )
        print(
            f"[{PLANNING_SHEET}] Đã tính trực tiếp M:R cho "
            f"{patched} mã; {changed_count} dòng thay đổi."
        )
        print("Upload thành công:", result.get("name", DEST_PATH))
    else:
        print(
            f"[{PLANNING_SHEET}] M:R đã đúng, không cần upload lại."
        )

    if state_changed:
        save_runtime_state(state)

    print(
        f"Nguồn tháng {source_key}; kế hoạch {plan_key}; "
        f"{FC_SHEET}!{FC_SELECTOR_CELL}={selector!r}."
    )


def main_with_retry(
    *,
    sleep_func=time.sleep,
    max_attempts=6,
    retry_delay_seconds=10,
):
    for attempt in range(1, max_attempts + 1):
        try:
            return main()
        except Exception as exc:
            if (
                not is_retryable_graph_error(exc)
                or attempt == max_attempts
            ):
                raise
            print(
                f"[{PLANNING_SHEET}] File đang khóa/thay đổi; "
                f"tải lại và thử lại ({attempt}/{max_attempts})."
            )
            sleep_func(retry_delay_seconds)
