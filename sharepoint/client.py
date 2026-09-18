"""Canonical Microsoft Graph / SharePoint client.

This module owns Graph transport, structured errors, retry classification and
SharePoint identity helpers. Legacy modules may re-export these names, but the
dependency direction is always legacy -> sharepoint.client.
"""

from __future__ import annotations

import base64
import re
import time
from typing import Callable, TypeVar
from urllib.parse import quote

import requests

from sharepoint.auth import get_access_token


GRAPH = "https://graph.microsoft.com/v1.0"
HOSTNAME = "vikodacomvn.sharepoint.com"
SITE_PATH = "/sites/Planning"


class GraphRequestError(RuntimeError):
    """Structured Microsoft Graph failure used by retry policies."""

    def __init__(self, message, *, status_code=None, error_code=None, detail=None):
        super().__init__(message)
        self.status_code = status_code
        self.error_code = error_code
        self.detail = detail


def is_retryable_graph_error(exc):
    """Return True only for transient/concurrency/network failures."""
    if isinstance(exc, GraphRequestError):
        if exc.status_code in {412, 423, 429, 500, 502, 503, 504}:
            return True
        if str(exc.error_code or "").casefold() in {
            "resourcelocked",
            "preconditionfailed",
            "toomanyrequests",
            "timeout",
            "serviceunavailable",
        }:
            return True
        return False

    if isinstance(exc, (requests.Timeout, requests.ConnectionError)):
        return True

    # Backward compatibility for existing fakes/tests while callers migrate.
    if isinstance(exc, RuntimeError):
        message = str(exc)
        return any(
            marker in message
            for marker in ("412", "423", "429", "resourceLocked", "500", "502", "503", "504")
        )
    return False


class GraphClient:
    def __init__(self, token):
        self.session = requests.Session()
        self.session.headers.update(
            {"Authorization": f"Bearer {token}"}
        )

    @staticmethod
    def _raise(response):
        if response.ok:
            return

        try:
            detail = response.json()
        except Exception:
            detail = getattr(response, "text", "")

        error_code = None
        if isinstance(detail, dict):
            error = detail.get("error")
            if isinstance(error, dict):
                error_code = error.get("code")

        raise GraphRequestError(
            f"Microsoft Graph lỗi {response.status_code}: {detail}",
            status_code=response.status_code,
            error_code=error_code,
            detail=detail,
        )

    def get_json(self, url, params=None):
        response = self.session.get(
            url,
            params=params,
            timeout=60,
        )
        self._raise(response)
        return response.json()

    def get_site_id(self):
        url = f"{GRAPH}/sites/{HOSTNAME}:{SITE_PATH}"
        return self.get_json(url, {"$select": "id"})["id"]

    def get_default_drive_id(self, site_id):
        url = f"{GRAPH}/sites/{site_id}/drive"
        return self.get_json(url, {"$select": "id"})["id"]

    def get_item_by_path(self, drive_id, file_path):
        encoded = quote(file_path, safe="/")
        url = f"{GRAPH}/drives/{drive_id}/root:/{encoded}"
        return self.get_json(
            url,
            {"$select": "id,name,eTag,size,lastModifiedDateTime"},
        )

    def list_folder_children(self, drive_id, folder_path=""):
        """Liệt kê các file/folder con trong một thư mục SharePoint."""
        if not folder_path or folder_path.strip() in ("", "/"):
            url = f"{GRAPH}/drives/{drive_id}/root/children"
        else:
            encoded = quote(folder_path.strip().strip("/"), safe="/")
            url = f"{GRAPH}/drives/{drive_id}/root:/{encoded}:/children"
        res = self.get_json(
            url,
            {"$select": "id,name,eTag,size,lastModifiedDateTime,folder,file"},
        )
        return res.get("value", [])

    def download_file(self, drive_id, item_id):
        url = f"{GRAPH}/drives/{drive_id}/items/{item_id}/content"
        response = self.session.get(
            url,
            timeout=120,
            allow_redirects=True,
        )
        self._raise(response)
        return response.content

    def upload_file(self, drive_id, item_id, content, expected_etag):
        url = f"{GRAPH}/drives/{drive_id}/items/{item_id}/content"
        headers = {
            "Content-Type": (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            "If-Match": expected_etag,
        }

        response = self.session.put(
            url,
            headers=headers,
            data=content,
            timeout=180,
        )

        if response.status_code == 412:
            raise GraphRequestError(
                "Microsoft Graph HTTP 412 preconditionFailed: file đích vừa thay đổi; "
                "phải tải lại workbook, tính lại patch và dùng ETag mới.",
                status_code=412,
                error_code="preconditionFailed",
            )

        self._raise(response)
        return response.json()

    def create_file_by_path(
        self,
        drive_id,
        file_path,
        content,
        conflict_behavior="fail",
    ):
        encoded = quote(file_path.strip().strip("/"), safe="/")
        url = f"{GRAPH}/drives/{drive_id}/root:/{encoded}:/content"
        params = {}
        if conflict_behavior:
            params["@microsoft.graph.conflictBehavior"] = conflict_behavior
        headers = {
            "Content-Type": (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        }
        response = self.session.put(
            url,
            headers=headers,
            params=params,
            data=content,
            timeout=180,
        )
        self._raise(response)
        return response.json()


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
