"""Safe entrypoint cho đồng bộ Ton_NVL!E.

Khóa đúng file BCTheodoiDMHANG.xlsm bằng SharePoint driveItem ID đã cấu hình,
thay vì phụ thuộc Microsoft Graph root/search. Các Graph GET idempotent dùng để
resolve/đọc metadata được retry khi gặp lỗi tạm thời 429/5xx/network.
"""

from __future__ import annotations

import json
from pathlib import Path
import sys
import time

import sync_nvl_open_po as base
from sync_stock import is_retryable_graph_error


_ORIGINAL_ITEM_METADATA = base._item_metadata
_ORIGINAL_RESOLVE_SOURCE_ITEM = base.resolve_source_item
READ_ATTEMPTS = 3


def _source_config_path_from_argv() -> Path:
    """Đọc đúng --source-config mà base.main() sẽ sử dụng."""
    args = sys.argv[1:]
    for index, arg in enumerate(args):
        if arg == "--source-config" and index + 1 < len(args):
            return Path(args[index + 1])
        if arg.startswith("--source-config="):
            return Path(arg.split("=", 1)[1])
    return Path(base.DEFAULT_SOURCE_CONFIG)


def _configured_source_item_id() -> str:
    path = _source_config_path_from_argv()
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    return str(data.get("source", {}).get("item_id", "")).strip()


def _retry_read(label: str, operation, attempts: int = READ_ATTEMPTS):
    for attempt in range(1, attempts + 1):
        try:
            return operation()
        except Exception as exc:
            if attempt >= attempts or not is_retryable_graph_error(exc):
                raise
            delay = min(2 ** (attempt - 1), 8)
            print(
                f"[Ton_NVL!E] Graph tạm lỗi tại {label} "
                f"(lượt {attempt}/{attempts}): {exc}. Thử lại sau {delay}s..."
            )
            time.sleep(delay)
    raise RuntimeError(f"Không thể hoàn tất Graph read: {label}")


def _item_metadata_with_retry(graph, drive_id: str, item_id: str) -> dict:
    return _retry_read(
        f"metadata item {item_id}",
        lambda: _ORIGINAL_ITEM_METADATA(graph, drive_id, item_id),
    )


def resolve_source_item_pinned(
    graph,
    drive_id: str,
    config,
    *,
    item_id: str | None = None,
) -> dict:
    """Resolve đúng source bằng item_id + sourcedoc; fallback chỉ khi config cũ."""
    pinned_item_id = str(item_id or _configured_source_item_id()).strip()
    if not pinned_item_id:
        return _ORIGINAL_RESOLVE_SOURCE_ITEM(graph, drive_id, config)

    item = _item_metadata_with_retry(graph, drive_id, pinned_item_id)

    actual_name = str(item.get("name", ""))
    if actual_name.casefold() != config.source_name.casefold():
        raise RuntimeError(
            f"Sai file nguồn cho item_id {pinned_item_id}: "
            f"nhận {actual_name!r}, mong đợi {config.source_name!r}."
        )

    expected_guid = base._normalize_guid(config.source_sourcedoc)
    actual_guid = base._normalize_guid(
        item.get("sharepointIds", {}).get("listItemUniqueId")
    )
    if expected_guid and actual_guid != expected_guid:
        raise RuntimeError(
            f"Sai sourcedoc file nguồn cho item_id {pinned_item_id}: "
            f"nhận {actual_guid}, mong đợi {expected_guid}."
        )

    print(
        f"[Ton_NVL!E] Đã khóa chính xác nguồn {config.source_name} "
        f"bằng item_id={pinned_item_id}; không dùng Graph Search."
    )
    return item


def install_safe_resolver() -> None:
    # Mọi metadata GET sau resolve (sau download / trước upload) cũng được retry.
    base._item_metadata = _item_metadata_with_retry
    base.resolve_source_item = resolve_source_item_pinned


def main() -> None:
    install_safe_resolver()
    base.main()


if __name__ == "__main__":
    main()
