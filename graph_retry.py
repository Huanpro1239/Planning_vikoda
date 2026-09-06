from datetime import datetime, timezone
from email.utils import parsedate_to_datetime

from sync_stock import GraphClient, GraphRequestError


_INSTALLED = False
_ORIGINAL_RAISE = GraphClient._raise


def parse_retry_after(value, *, now=None):
    """Parse Retry-After seconds or HTTP-date into a non-negative delay."""
    if value in (None, ""):
        return None

    text = str(value).strip()
    try:
        return max(0.0, float(text))
    except ValueError:
        pass

    try:
        target = parsedate_to_datetime(text)
        if target.tzinfo is None:
            target = target.replace(tzinfo=timezone.utc)
        current = now or datetime.now(timezone.utc)
        if current.tzinfo is None:
            current = current.replace(tzinfo=timezone.utc)
        return max(0.0, (target - current).total_seconds())
    except Exception:
        return None


def install_retry_after_support():
    """Preserve Graph Retry-After on GraphRequestError without changing callers."""
    global _INSTALLED
    if _INSTALLED:
        return

    original_raise = _ORIGINAL_RAISE

    def wrapped_raise(response):
        try:
            original_raise(response)
        except GraphRequestError as exc:
            headers = getattr(response, "headers", {}) or {}
            delay = parse_retry_after(headers.get("Retry-After"))
            if delay is not None:
                exc.retry_after_seconds = delay
            raise

    GraphClient._raise = staticmethod(wrapped_raise)
    _INSTALLED = True


def retry_delay_seconds(exc, default_seconds):
    value = getattr(exc, "retry_after_seconds", None)
    if value is None:
        return float(default_seconds)
    try:
        return max(0.0, float(value))
    except (TypeError, ValueError):
        return float(default_seconds)
