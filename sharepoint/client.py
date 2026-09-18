"""Canonical SharePoint / Microsoft Graph integration helpers.

`sync_stock.py` hiện vẫn giữ implementation thấp tầng để không phá backward compatibility,
nhưng toàn bộ code mới nên đi qua module này. `SharePointClient` và `GraphClient`
trỏ tới cùng một class object; các helper ở đây gom resolve URL, identity validation,
retry idempotent reads và ETag-aware metadata handling về một tầng dùng chung.
"""

from __future__ import annotations

import base64
import re
import time
from typing import Callable, TypeVar

from sharepoint.auth import get_access_token
from sync_stock import (
    GRAPH,
    HOSTNAME,
    SITE_PATH,
    GraphClient,
    GraphRequestError,
    is_retryable_graph_error,
)


SharePointClient = GraphClient
T = TypeVar("T")
_GUID_RE = re.compile(
    r"^[{(]?([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-"
    r"[0-9a-fA-F]{4}-[0-9a-fA-F]{12})[})]?$"
)


def normalize_guid(value) -> str:
    """Chuẩn hóa GUID để so identity SharePoint ổn định, không phụ thuộc dấu ngoặc/case."""
    text = str(value or "").strip()
    match = _GUID_RE.fullmatch(text)
    return match.group(1).upper() if match else ""


def graph_share_id(url: str) -> str:
    """Encode browser/sharing URL thành Graph share-id chuẩn `u!…`."""
    if not str(url or "").strip():
        raise ValueError("SharePoint URL không được để trống.")
    token = base64.urlsafe_b64encode(str(url).encode("utf-8")).decode("ascii").rstrip("=")
    return f"u!{token}"


def retry_graph_read(
    label: str,
    operation: Callable[[], T],
    *,
    attempts: int = 3,
    max_delay_seconds: int = 8,
    log_prefix: str = "[SharePoint]",
) -> T:
    """Retry chỉ các thao tác đọc idempotent khi Graph/network lỗi tạm thời."""
    if attempts < 1:
        raise ValueError("attempts phải >= 1")

    for attempt in range(1, attempts + 1):
        try:
            return operation()
        except Exception as exc:
            if attempt >= attempts or not is_retryable_graph_error(exc):
                raise
            delay = min(2 ** (attempt - 1), max_delay_seconds)
            print(
                f"{log_prefix} Graph tạm lỗi tại {label} "
                f"(lượt {attempt}/{attempts}): {exc}. Thử lại sau {delay}s..."
            )
            time.sleep(delay)

    raise RuntimeError(f"Không thể hoàn tất Graph read: {label}")


def flatten_remote_item(item: dict) -> dict:
    """Chuẩn hóa kết quả `/shares/.../driveItem` khi Graph bọc dữ liệu trong remoteItem."""
    remote = item.get("remoteItem") if isinstance(item, dict) else None
    if not isinstance(remote, dict):
        return item

    merged = dict(remote)
    for key in ("parentReference", "sharepointIds", "webUrl", "eTag", "cTag"):
        if key not in merged and key in item:
            merged[key] = item[key]
    return merged


def validate_item_identity(
    item: dict,
    *,
    expected_name: str | None = None,
    expected_sourcedoc: str | None = None,
    expected_drive_id: str | None = None,
    label: str = "SharePoint item",
) -> dict:
    """Fail-closed khi Graph resolve nhầm file, sourcedoc hoặc document library."""
    item = flatten_remote_item(item)

    item_id = str(item.get("id", "")).strip()
    if not item_id:
        raise RuntimeError(f"Graph không trả về driveItem id cho {label}.")

    if expected_name:
        actual_name = str(item.get("name", ""))
        if actual_name.casefold() != str(expected_name).casefold():
            raise RuntimeError(
                f"Sai file cho {label}: nhận {actual_name!r}, mong đợi {expected_name!r}."
            )

    expected_guid = normalize_guid(expected_sourcedoc)
    if expected_guid:
        actual_guid = normalize_guid(
            item.get("sharepointIds", {}).get("listItemUniqueId")
        )
        if actual_guid != expected_guid:
            raise RuntimeError(
                f"Sai sourcedoc cho {label}: nhận {actual_guid or '<trống>'}, "
                f"mong đợi {expected_guid}."
            )

    if expected_drive_id:
        actual_drive_id = str(item.get("parentReference", {}).get("driveId", "")).strip()
        if actual_drive_id and actual_drive_id != str(expected_drive_id).strip():
            raise RuntimeError(
                f"{label} nằm ở document library khác: "
                f"drive thực tế={actual_drive_id}, drive mong đợi={expected_drive_id}."
            )

    return item


def resolve_share_url(
    graph: GraphClient,
    url: str,
    *,
    expected_name: str | None = None,
    expected_sourcedoc: str | None = None,
    expected_drive_id: str | None = None,
    attempts: int = 3,
    label: str = "SharePoint URL",
) -> dict:
    """Resolve exact browser/sharing URL; tuyệt đối không search theo tên file."""
    share_id = graph_share_id(url)
    endpoint = (
        f"{GRAPH}/shares/{share_id}/driveItem"
        "?$select=id,name,eTag,cTag,size,webUrl,lastModifiedDateTime,"
        "sharepointIds,parentReference,remoteItem"
    )
    item = retry_graph_read(
        f"resolve {label}",
        lambda: graph.get_json(endpoint),
        attempts=attempts,
    )
    return validate_item_identity(
        item,
        expected_name=expected_name,
        expected_sourcedoc=expected_sourcedoc,
        expected_drive_id=expected_drive_id,
        label=label,
    )


def get_item_metadata(
    graph: GraphClient,
    drive_id: str,
    item_id: str,
    *,
    attempts: int = 3,
    select: str = "id,name,eTag,cTag,size,lastModifiedDateTime,sharepointIds,parentReference",
) -> dict:
    """Đọc metadata item theo ID với retry đọc thống nhất."""
    endpoint = f"{GRAPH}/drives/{drive_id}/items/{item_id}"
    return retry_graph_read(
        f"metadata item {item_id}",
        lambda: graph.get_json(endpoint, {"$select": select}),
        attempts=attempts,
    )


def download_file_with_retry(
    graph: GraphClient,
    drive_id: str,
    item_id: str,
    *,
    attempts: int = 3,
) -> bytes:
    """Download snapshot với retry có giới hạn cho lỗi GET/network tạm thời."""
    return retry_graph_read(
        f"download item {item_id}",
        lambda: graph.download_file(drive_id, item_id),
        attempts=attempts,
    )


def assert_same_revision(before: dict, after: dict, *, label: str = "SharePoint item") -> None:
    """Chặn publish nếu nguồn/đích đổi revision giữa hai thời điểm đọc metadata."""
    before_etag = str(before.get("eTag", ""))
    after_etag = str(after.get("eTag", ""))
    if before_etag and after_etag and before_etag != after_etag:
        raise GraphRequestError(
            f"{label} đã thay đổi revision: {before_etag} -> {after_etag}.",
            status_code=412,
            error_code="preconditionFailed",
            detail={"before_etag": before_etag, "after_etag": after_etag},
        )


__all__ = [
    "GRAPH",
    "HOSTNAME",
    "SITE_PATH",
    "GraphClient",
    "SharePointClient",
    "GraphRequestError",
    "get_access_token",
    "is_retryable_graph_error",
    "normalize_guid",
    "graph_share_id",
    "retry_graph_read",
    "flatten_remote_item",
    "validate_item_identity",
    "resolve_share_url",
    "get_item_metadata",
    "download_file_with_retry",
    "assert_same_revision",
]
