"""Transitional SharePoint/Graph facade.

Phase 1 chỉ cung cấp canonical import path. Implementation hiện tại vẫn ở
``sync_stock.py`` để giữ backward compatibility. Phase 2 sẽ chuyển implementation
vào đây và để ``sync_stock.py`` re-export lại API cũ.
"""

from sync_stock import (
    GRAPH,
    HOSTNAME,
    SITE_PATH,
    GraphClient,
    GraphRequestError,
    get_access_token,
    is_retryable_graph_error,
)

__all__ = [
    "GRAPH",
    "HOSTNAME",
    "SITE_PATH",
    "GraphClient",
    "GraphRequestError",
    "get_access_token",
    "is_retryable_graph_error",
]
