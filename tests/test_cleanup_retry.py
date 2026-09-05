import unittest

import cleanup_planning


class FakeGraph:
    def __init__(self):
        self.item_calls = 0
        self.upload_calls = 0

    def get_item_by_path(self, drive_id, path):
        self.item_calls += 1
        return {
            "id": "dest-id",
            "eTag": f"etag-{self.item_calls}",
        }

    def download_file(self, drive_id, item_id):
        return b"workbook"

    def upload_file(self, drive_id, item_id, content, expected_etag):
        self.upload_calls += 1
        if self.upload_calls == 1:
            raise RuntimeError("Microsoft Graph lỗi 423: resourceLocked")
        return {"name": "Sắp kế hoạch.xlsx"}


class CleanupRetryTests(unittest.TestCase):
    def test_reloads_file_and_retries_when_sharepoint_is_locked(self):
        graph = FakeGraph()
        cleanup_calls = []

        def fake_cleanup(data, sheet_name):
            cleanup_calls.append((data, sheet_name))
            return b"updated", 40

        removed = cleanup_planning.cleanup_with_retry(
            graph,
            "drive-id",
            cleanup_func=fake_cleanup,
            sleep_func=lambda _: None,
            max_attempts=2,
        )

        self.assertEqual(removed, 40)
        self.assertEqual(graph.upload_calls, 2)
        self.assertEqual(graph.item_calls, 2)
        self.assertEqual(len(cleanup_calls), 2)


if __name__ == "__main__":
    unittest.main()
