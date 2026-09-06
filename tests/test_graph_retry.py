import unittest
from unittest.mock import patch

import requests

from sync_stock import GraphClient, GraphRequestError, is_retryable_graph_error


class Response:
    def __init__(self, status_code, payload=None):
        self.status_code = status_code
        self._payload = payload or {"error": {"code": "serverError"}}
        self.ok = 200 <= status_code < 300
        self.text = str(self._payload)

    def json(self):
        return self._payload


class GraphRetryTests(unittest.TestCase):
    def test_real_upload_412_keeps_structured_status_and_is_retryable(self):
        graph = GraphClient("fake-token")
        with patch.object(graph.session, "put", return_value=Response(412)):
            with self.assertRaises(GraphRequestError) as ctx:
                graph.upload_file("drive", "item", b"data", "etag-old")

        self.assertEqual(ctx.exception.status_code, 412)
        self.assertEqual(ctx.exception.error_code, "preconditionFailed")
        self.assertTrue(is_retryable_graph_error(ctx.exception))

    def test_429_5xx_and_timeout_are_retryable(self):
        for status in (429, 500, 502, 503, 504):
            self.assertTrue(
                is_retryable_graph_error(
                    GraphRequestError("transient", status_code=status)
                )
            )
        self.assertTrue(is_retryable_graph_error(requests.Timeout("timeout")))
        self.assertFalse(
            is_retryable_graph_error(GraphRequestError("bad request", status_code=400))
        )


if __name__ == "__main__":
    unittest.main()
