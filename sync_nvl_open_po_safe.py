"""Safe entrypoint cho đồng bộ Ton_NVL!E.

Resolve đúng BCTheodoiDMHANG.xlsm từ URL SharePoint exact do người dùng cung cấp,
thay vì Graph root/search hoặc hard-code driveItem ID. URL được đổi sang Graph share-id,
resolve về driveItem, sau đó xác minh lại tên file + sourcedoc trước khi đọc.
Các Graph GET idempotent được retry khi gặp lỗi tạm thời 429/5xx/network.
"""

from __future__ import annotations

import base64
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


def _load_source_config() -> dict:
    path = _source_config_path_from_argv()
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("source", {})


def _configured_source_url() -> str:
    return str(_load_source_config().get("source_url", "")).strip()


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


def _graph_share_id(url: str) -> str:
    """Encode SharePoint sharing/browser URL theo chuẩn Microsoft Graph /shares/u!..."""
    token = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")
    return f"u!{token}"


def _validate_source_identity(item: dict, config, *, label: str) -> dict:
    remote = item.get("remoteItem")
    if isinstance(remote, dict):
        merged = dict(remote)
        # Giữ parentReference ngoài nếu remoteItem không có.
        if "parentReference" not in merged and "parentReference" in item:
            merged["parentReference"] = item["parentReference"]
        item = merged

    actual_name = str(item.get("name", ""))
    if actual_name.casefold() != config.source_name.casefold():
        raise RuntimeError(
            f"Sai file nguồn từ {label}: nhận {actual_name!r}, "
            f"mong đợi {config.source_name!r}."
        )

    expected_guid = base._normalize_guid(config.source_sourcedoc)
    actual_guid = base._normalize_guid(
        item.get("sharepointIds", {}).get("listItemUniqueId")
    )
    if expected_guid and actual_guid != expected_guid:
        raise RuntimeError(
            f"Sai sourcedoc file nguồn từ {label}: nhận {actual_guid or '<trống>'}, "
            f"mong đợi {expected_guid}."
        )

    if not item.get("id"):
        raise RuntimeError(f"Graph không trả về driveItem id cho nguồn {label}.")
    return item


def resolve_source_item_exact_url(graph, drive_id: str, config) -> dict:
    """Resolve source bằng exact SharePoint URL; không Search theo tên, không pin item ID."""
    source_url = _configured_source_url()
    if not source_url:
        # Tương thích config cũ, nhưng production config hiện phải có source_url.
        return _ORIGINAL_RESOLVE_SOURCE_ITEM(graph, drive_id, config)

    share_id = _graph_share_id(source_url)
    endpoint = (
        f"{base.GRAPH}/shares/{share_id}/driveItem"
        "?$select=id,name,eTag,cTag,size,webUrl,lastModifiedDateTime,sharepointIds,parentReference,remoteItem"
    )
    item = _retry_read(
        "resolve exact SharePoint URL",
        lambda: graph.get_json(endpoint),
    )
    item = _validate_source_identity(item, config, label="source_url")

    actual_drive_id = str(item.get("parentReference", {}).get("driveId", "")).strip()
    if actual_drive_id and actual_drive_id != drive_id:
        raise RuntimeError(
            "Nguồn PO resolve sang document library khác với drive mặc định của workflow. "
            f"drive nguồn={actual_drive_id}, drive workflow={drive_id}. "
            "Cần mở rộng workflow để dùng source drive riêng trước khi publish."
        )

    print(
        f"[Ton_NVL!E] Đã resolve đúng nguồn {config.source_name} từ exact SharePoint URL; "
        f"item_id={item['id']}; sourcedoc đã xác minh."
    )
    return item


def install_safe_resolver() -> None:
    # Mọi metadata GET sau resolve (sau download / trước upload) cũng được retry.
    base._item_metadata = _item_metadata_with_retry
    base.resolve_source_item = resolve_source_item_exact_url


def main() -> None:
    install_safe_resolver()
    base.main()


if __name__ == "__main__":
    main()
