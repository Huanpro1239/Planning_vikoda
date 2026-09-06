import unittest
from unittest.mock import patch

import sync_planning_pipeline as pipeline_runner
import sync_stock
from sync_stock import GraphRequestError


class PublishBoundaryTests(unittest.TestCase):
    class FakeGraph:
        def __init__(self, *, first_upload_412=False):
            self.uploads = []
            self.attempt = 0
            self.first_upload_412 = first_upload_412

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
            if self.first_upload_412 and len(self.uploads) == 1:
                raise GraphRequestError(
                    "precondition failed",
                    status_code=412,
                    error_code="preconditionFailed",
                )
            return {"name": "copy.xlsx"}

    @staticmethod
    def _report(input_revision, status):
        return {
            "schema_version": 2,
            "algorithm": "priority_v8_rgb_v10_galon_v9",
            "plan_month": "2026-09",
            "input_revision": input_revision,
            "mass_balance": {},
            "resources": {},
            "publish_status": status,
            "pipeline": {"conversion_hash": "conversion-hash"},
        }

    def _expected_proposal_id(self, status="review_required"):
        preview_graph = self.FakeGraph()
        _, target_bytes, _, _, revision = pipeline_runner._read_snapshot(
            preview_graph,
            "drive",
        )
        report = self._report(revision, status)
        return pipeline_runner._with_proposal_identity(
            report,
            b"final:" + target_bytes,
        )["proposal_id"]

    def _run(
        self,
        *,
        status,
        publish_mode="publish",
        review_approval=None,
        graph=None,
        max_attempts=2,
        retry_delay_seconds=3,
    ):
        graph = graph or self.FakeGraph()
        saved_states = []
        saved_decisions = []
        proposals = []
        sleeps = []

        def prepare(target_bytes, source_bytes, *, runtime_state, input_revision):
            return (
                b"final:" + target_bytes,
                self._report(input_revision, status),
                {"state": f"new-r{graph.attempt}"},
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
                lambda *args: saved_states.append(args),
            ),
            patch.object(
                pipeline_runner,
                "_save_publish_decision",
                lambda decision: saved_decisions.append(dict(decision)),
            ),
            patch.object(
                pipeline_runner,
                "_save_proposal_artifacts",
                lambda data, report: proposals.append((data, dict(report))),
            ),
            patch.object(pipeline_runner, "print_operational_report", lambda report: None),
        ):
            result = pipeline_runner.run_pipeline_with_retry(
                graph,
                "drive",
                publish_mode=publish_mode,
                review_approval=review_approval,
                publisher="planner",
                sleep_func=sleeps.append,
                max_attempts=max_attempts,
                retry_delay_seconds=retry_delay_seconds,
            )

        return result, graph, saved_states, saved_decisions, proposals, sleeps

    def test_cli_defaults_to_proposal_mode(self):
        args = pipeline_runner._parse_args([])
        self.assertFalse(args.publish)

    def test_ready_for_publish_explicit_publish_uploads_and_saves_state(self):
        result, graph, states, decisions, proposals, _ = self._run(
            status="ready_for_publish",
            publish_mode="publish",
        )

        self.assertTrue(result["uploaded"])
        self.assertFalse(result["publish_blocked"])
        self.assertEqual(len(graph.uploads), 1)
        self.assertEqual(graph.uploads[0][2], "target-etag-r1")
        self.assertEqual(len(states), 1)
        self.assertEqual(len(decisions), 1)
        self.assertEqual(decisions[0]["basis"], "publish_status")
        self.assertEqual(decisions[0]["state"], "published")
        self.assertGreaterEqual(len(proposals), 2)

    def test_ready_for_publish_proposal_mode_never_uploads_or_saves_state(self):
        result, graph, states, decisions, proposals, _ = self._run(
            status="ready_for_publish",
            publish_mode="proposal",
        )

        self.assertFalse(result["uploaded"])
        self.assertFalse(result["publish_blocked"])
        self.assertTrue(result["would_change"])
        self.assertEqual(graph.uploads, [])
        self.assertEqual(states, [])
        self.assertEqual(decisions, [])
        self.assertEqual(len(proposals), 1)
        self.assertEqual(
            proposals[0][1]["publish_decision"]["state"],
            "proposal_only",
        )

    def test_review_required_without_approval_is_blocked(self):
        result, graph, states, decisions, proposals, _ = self._run(
            status="review_required",
            publish_mode="publish",
        )

        self.assertFalse(result["uploaded"])
        self.assertTrue(result["publish_blocked"])
        self.assertEqual(graph.uploads, [])
        self.assertEqual(states, [])
        self.assertEqual(decisions, [])
        self.assertEqual(len(proposals), 1)
        self.assertEqual(
            proposals[0][1]["publish_decision"]["state"],
            "blocked",
        )

    def test_review_required_matching_snapshot_and_reason_can_publish(self):
        proposal_id = self._expected_proposal_id()
        approval = {
            "proposal_id": proposal_id,
            "reason": "Chấp nhận thiếu hàng bất khả kháng đã được review capacity.",
            "approved_by": "planner",
        }

        result, graph, states, decisions, _, _ = self._run(
            status="review_required",
            publish_mode="publish",
            review_approval=approval,
        )

        self.assertTrue(result["uploaded"])
        self.assertFalse(result["publish_blocked"])
        self.assertEqual(len(graph.uploads), 1)
        self.assertEqual(len(states), 1)
        self.assertEqual(len(decisions), 1)
        self.assertEqual(decisions[0]["basis"], "review_approval")
        self.assertEqual(
            decisions[0]["review_approval"]["proposal_id"],
            proposal_id,
        )
        self.assertTrue(decisions[0]["review_approval"]["reason"])

    def test_review_required_matching_snapshot_without_reason_is_blocked(self):
        proposal_id = self._expected_proposal_id()
        approval = {
            "proposal_id": proposal_id,
            "reason": "",
            "approved_by": "planner",
        }

        result, graph, states, decisions, _, _ = self._run(
            status="review_required",
            publish_mode="publish",
            review_approval=approval,
        )

        self.assertTrue(result["publish_blocked"])
        self.assertEqual(graph.uploads, [])
        self.assertEqual(states, [])
        self.assertEqual(decisions, [])

    def test_412_recompute_invalidates_review_approval_for_old_snapshot(self):
        proposal_id = self._expected_proposal_id()
        approval = {
            "proposal_id": proposal_id,
            "reason": "Chấp nhận shortage theo proposal đầu tiên.",
            "approved_by": "planner",
        }
        graph = self.FakeGraph(first_upload_412=True)

        result, graph, states, decisions, proposals, sleeps = self._run(
            status="review_required",
            publish_mode="publish",
            review_approval=approval,
            graph=graph,
            max_attempts=2,
            retry_delay_seconds=3,
        )

        self.assertTrue(result["publish_blocked"])
        self.assertFalse(result["uploaded"])
        self.assertEqual(len(graph.uploads), 1)
        self.assertEqual(graph.uploads[0][2], "target-etag-r1")
        self.assertEqual(states, [])
        self.assertEqual(decisions, [])
        self.assertEqual(sleeps, [3])
        self.assertNotEqual(result["report"]["proposal_id"], proposal_id)
        self.assertEqual(proposals[-1][1]["publish_decision"]["state"], "blocked")


if __name__ == "__main__":
    unittest.main()
