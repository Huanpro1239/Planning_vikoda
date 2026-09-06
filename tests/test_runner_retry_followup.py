import unittest
from unittest.mock import patch

import sync_planning_calendar_all_months as calendar_all
import sync_planning_schedule_production as production
from sync_stock import GraphRequestError


class FakeGraph:
    def __init__(self, *, first_download_error=None, first_upload_error=None):
        self.first_download_error = first_download_error
        self.first_upload_error = first_upload_error
        self.item_calls = 0
        self.download_calls = 0
        self.upload_calls = 0
        self.uploads = []

    def get_item_by_path(self, drive_id, path):
        self.item_calls += 1
        return {
            "id": "dest",
            "eTag": f"etag-{self.item_calls}",
            "lastModifiedDateTime": f"rev-{self.item_calls}",
        }

    def download_file(self, drive_id, item_id):
        self.download_calls += 1
        if self.download_calls == 1 and self.first_download_error is not None:
            raise self.first_download_error
        return f"snapshot-{self.download_calls}".encode()

    def upload_file(self, drive_id, item_id, content, expected_etag):
        self.upload_calls += 1
        self.uploads.append((content, expected_etag))
        if self.upload_calls == 1 and self.first_upload_error is not None:
            raise self.first_upload_error
        return {"name": "copy.xlsx"}


class RunnerRetryFollowupTests(unittest.TestCase):
    def test_calendar_retries_transient_download_and_recomputes(self):
        err = GraphRequestError("service unavailable", status_code=503)
        err.retry_after_seconds = 6
        graph = FakeGraph(first_download_error=err)
        inputs = []
        sleeps = []

        def fake_prepare(data):
            inputs.append(data)
            return b"calendar-" + data, {
                "selector": "Tháng 9",
                "plan_month": 9,
                "plan_year": 2026,
                "start": "S1",
                "end": "AV1",
                "days": 30,
                "changed_count": 1,
            }

        with patch.object(calendar_all, "prepare_calendar_update_all_months", fake_prepare):
            calendar_all.run_calendar_with_retry(
                graph,
                "drive",
                sleep_func=sleeps.append,
                max_attempts=3,
                retry_delay_seconds=2,
            )

        self.assertEqual(graph.download_calls, 2)
        self.assertEqual(inputs, [b"snapshot-2"])
        self.assertEqual(graph.uploads, [(b"calendar-snapshot-2", "etag-2")])
        self.assertEqual(sleeps, [6])

    def test_scheduler_412_reloads_recomputes_and_uses_new_etag(self):
        err = GraphRequestError(
            "precondition failed",
            status_code=412,
            error_code="preconditionFailed",
        )
        graph = FakeGraph(first_upload_error=err)
        prepare_inputs = []
        verified_outputs = []
        sleeps = []

        def fake_prepare(data):
            prepare_inputs.append(data)
            return b"schedule-" + data, {
                "headers": [],
                "products": [],
                "schedule": {},
                "line_capacity": {},
                "line_usage": {},
                "utilization": {},
                "carryover": {},
                "inventory": {},
                "optimizer_meta": {},
                "changed_count": 1,
            }

        def fake_report(data, info, *, input_revision, algorithm):
            return {
                "schema_version": 2,
                "algorithm": algorithm,
                "plan_month": None,
                "input_revision": input_revision,
                "mass_balance": {},
                "resources": {},
                "publish_status": "feasible",
            }

        with (
            patch.object(production.priority.base, "prepare_schedule_update", fake_prepare),
            patch.object(production, "build_schedule_report", fake_report),
            patch.object(production, "attach_output_hash", lambda report, data: report),
            patch.object(production, "verify_workbook", lambda data, schedule_report=None: verified_outputs.append(data)),
            patch.object(production, "save_schedule_report", lambda report: None),
            patch.object(production, "_print_scheduler_summary", lambda info: None),
            patch.object(production, "print_operational_report", lambda report: None),
        ):
            production.run_scheduler_with_retry(
                graph,
                "drive",
                sleep_func=sleeps.append,
                max_attempts=2,
                retry_delay_seconds=3,
            )

        self.assertEqual(prepare_inputs, [b"snapshot-1", b"snapshot-2"])
        self.assertEqual(
            verified_outputs,
            [b"schedule-snapshot-1", b"schedule-snapshot-2"],
        )
        self.assertEqual(
            graph.uploads,
            [
                (b"schedule-snapshot-1", "etag-1"),
                (b"schedule-snapshot-2", "etag-2"),
            ],
        )
        self.assertEqual(sleeps, [3])

    def test_scheduler_retries_download_429_using_retry_after(self):
        err = GraphRequestError(
            "too many requests",
            status_code=429,
            error_code="tooManyRequests",
        )
        err.retry_after_seconds = 8
        graph = FakeGraph(first_download_error=err)
        sleeps = []

        def fake_prepare(data):
            return data, {
                "headers": [],
                "products": [],
                "schedule": {},
                "line_capacity": {},
                "line_usage": {},
                "utilization": {},
                "carryover": {},
                "inventory": {},
                "optimizer_meta": {},
                "changed_count": 0,
            }

        with (
            patch.object(production.priority.base, "prepare_schedule_update", fake_prepare),
            patch.object(production, "build_schedule_report", lambda *args, **kwargs: {
                "schema_version": 2,
                "algorithm": "priority_v8",
                "plan_month": None,
                "input_revision": {"etag": "copy"},
                "mass_balance": {},
                "resources": {},
                "publish_status": "feasible",
            }),
            patch.object(production, "attach_output_hash", lambda report, data: report),
            patch.object(production, "verify_workbook", lambda *args, **kwargs: {}),
            patch.object(production, "save_schedule_report", lambda report: None),
            patch.object(production, "_print_scheduler_summary", lambda info: None),
            patch.object(production, "print_operational_report", lambda report: None),
        ):
            production.run_scheduler_with_retry(
                graph,
                "drive",
                sleep_func=sleeps.append,
                max_attempts=3,
                retry_delay_seconds=2,
            )

        self.assertEqual(graph.download_calls, 2)
        self.assertEqual(graph.upload_calls, 0)
        self.assertEqual(sleeps, [8])


if __name__ == "__main__":
    unittest.main()
