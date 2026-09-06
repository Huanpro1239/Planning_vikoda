import hashlib
import json
import time
from pathlib import Path

from graph_retry import install_retry_after_support, retry_delay_seconds as retry_wait_seconds
from planning_pipeline import prepare_pipeline_output
from planning_schedule_report import print_operational_report, save_schedule_report
import sync_planning_metrics as metrics
import sync_stock
from sync_stock import GraphClient, get_access_token, is_retryable_graph_error


AUDIT_REVISION_FILE = Path("planning_input_revision.json")
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


def run_pipeline_with_retry(
    graph,
    drive_id,
    *,
    sleep_func=time.sleep,
    max_attempts=6,
    retry_delay_seconds=10,
):
    """
    Read all inputs -> compute entire workbook -> validate -> one ETag upload.

    Any 412/transient read/write error restarts from fresh target + source
    snapshots. State files are saved only after a validated no-op or successful
    upload.
    """
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
            # prepare_pipeline_output already validates the exact final bytes.
            print_operational_report(report)

            changed = hashlib.sha256(final_bytes).digest() != hashlib.sha256(target_bytes).digest()
            if changed:
                graph.upload_file(
                    drive_id,
                    target_item["id"],
                    final_bytes,
                    expected_etag=target_item["eTag"],
                )
                print(
                    "[PIPELINE] Validated snapshot uploaded exactly once with If-Match ETag."
                )
            else:
                print("[PIPELINE] Final workbook already matches validated output; no upload.")

            _save_states_after_success(
                source_items,
                report,
                proposed_runtime_state,
            )
            return {
                "uploaded": changed,
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


def main():
    install_retry_after_support()
    token = get_access_token()
    graph = GraphClient(token)
    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)
    run_pipeline_with_retry(graph, drive_id)


if __name__ == "__main__":
    main()
