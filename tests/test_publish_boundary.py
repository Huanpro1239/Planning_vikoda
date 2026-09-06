import unittest
from unittest.mock import patch

import sync_planning_pipeline as pipeline_runner
import sync_stock


class PublishBoundaryTests(unittest.TestCase):
    class FakeGraph:
        def __init__(self):
            self.uploads = []
            self.attempt = 0

        def get_item_by_path(self, drive_id, path):
            if path == sync_stock.DEST_PATH:
                self.attempt += 1
                return {
                    "id": f"target-r{self.attempt}",
                    "eTag": f"target-etag-r{self.attempt}",
                    "lastModifiedDateTime": f"target-time-r{self.attempt}",
                }
            key = next(
                key for key, value in pipeline_runner.SOURCES.items() if value == path
            )
            return {
                "id": f"{key}-r{self.attempt}",
                "eTag": f"{key}-etag-r{self.attempt}",
                "lastModifiedDateTime": f"{key}-time-r{self.attempt}",
            }

        def download_file(self, drive_id, item_id):
            return f"bytes:{item_id}".encode()

        def upload_file(self, drive_id, item_id, content, expected_etag):
            self.uploads.append((item_id, content, expected_etag))
            return {"name": "copy.xlsx"}

    @staticmethod
    def _review_required_report(input_revision):
        return {
            "schema_version": 2,
            "algorithm": "priority_v8_rgb_v10_galon_v9",
            "plan_month": "2026-09",
            "input_revision": input_revision,
            "output_sha256": "output-hash",
            "mass_balance": {},
            "resources": {},
            "publish_status": "review_required",
            "pipeline": {"conversion_hash": "conversion-hash"},
        }

    def test_review_required_default_run_does_not_upload_or_save_state(self):
        graph = self.FakeGraph()
        saved = []

        def prepare(target_bytes, source_bytes, *, runtime_state, input_revision):
            return (
                b"final:" + target_bytes,
                self._review_required_report(input_revision),
                {"state": "new"},
            )

        with (
            patch.object(pipeline_runner, "prepare_pipeline_output", prepare),
            patch.object(
                pipeline_runner.metrics,
                "load_runtime_state",
                lambda: {"state": "old"},
            ),
            patch.object(
                pipeline_runner,
                "_save_states_after_success",
                lambda *args: saved.append(args),
            ),
            patch.object(pipeline_runner, "print_operational_report", lambda report: None),
        ):
            result = pipeline_runner.run_pipeline_with_retry(
                graph,
                "drive",
                sleep_func=lambda _: None,
                max_attempts=1,
            )

        self.assertFalse(result["uploaded"])
        self.assertEqual(graph.uploads, [])
        self.assertEqual(saved, [])
        self.assertEqual(result["report"]["publish_status"], "review_required")


if __name__ == "__main__":
    unittest.main()
