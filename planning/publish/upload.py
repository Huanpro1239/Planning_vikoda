"""Bounded SharePoint upload retry for authorized Planning snapshots."""

from sharepoint.retry import retry_delay_seconds as retry_wait_seconds


def _is_resource_locked(exc):
    return (
        getattr(exc, "status_code", None) == 423
        or str(getattr(exc, "error_code", "") or "").casefold()
        == "resourcelocked"
    )


def upload_authorized_snapshot(
    graph,
    drive_id,
    target_item,
    final_bytes,
    *,
    sleep_func,
    lock_attempts=4,
    lock_retry_delay_seconds=15,
):
    """Retry HTTP 423 with the exact snapshot/ETag; let 412 trigger recompute."""
    if lock_attempts < 1:
        raise ValueError("lock_attempts phải >= 1")

    for lock_attempt in range(1, lock_attempts + 1):
        try:
            return graph.upload_file(
                drive_id,
                target_item["id"],
                final_bytes,
                expected_etag=target_item["eTag"],
            )
        except Exception as exc:
            if not _is_resource_locked(exc):
                raise
            if lock_attempt >= lock_attempts:
                exc.planning_lock_retries_exhausted = True
                raise

            default_delay = min(
                lock_retry_delay_seconds * lock_attempt,
                60,
            )
            delay = min(
                retry_wait_seconds(exc, default_delay),
                60,
            )
            print(
                "[PIPELINE] SharePoint target đang bị khóa "
                f"(HTTP 423, lượt {lock_attempt}/{lock_attempts}). "
                "Giữ nguyên authorized snapshot + ETag; thử upload lại "
                f"sau {delay:g}s."
            )
            sleep_func(delay)
