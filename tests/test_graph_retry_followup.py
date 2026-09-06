import unittest
from datetime import datetime, timezone

from graph_retry import parse_retry_after, retry_delay_seconds
from sync_stock import GraphRequestError


class GraphRetryFollowupTests(unittest.TestCase):
    def test_retry_after_seconds_is_respected(self):
        self.assertEqual(parse_retry_after("7"), 7.0)

        exc = GraphRequestError(
            "too many requests",
            status_code=429,
            error_code="tooManyRequests",
        )
        exc.retry_after_seconds = 9
        self.assertEqual(retry_delay_seconds(exc, 2), 9.0)

    def test_retry_after_http_date_is_respected(self):
        now = datetime(2026, 9, 6, 10, 0, 0, tzinfo=timezone.utc)
        self.assertEqual(
            parse_retry_after("Sun, 06 Sep 2026 10:00:12 GMT", now=now),
            12.0,
        )

    def test_invalid_retry_after_falls_back(self):
        exc = GraphRequestError("service unavailable", status_code=503)
        self.assertEqual(retry_delay_seconds(exc, 4), 4.0)
        self.assertIsNone(parse_retry_after("not-a-delay"))


if __name__ == "__main__":
    unittest.main()
