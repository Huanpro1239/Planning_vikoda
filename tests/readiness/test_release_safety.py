"""Release-safety checks beyond ordinary unit regression."""

from __future__ import annotations

import copy
from io import BytesIO
from unittest.mock import patch
import unittest

from openpyxl import load_workbook

from planning.pipeline import prepare_pipeline_output
from planning.publish import proposal as proposal_api
from planning.publish import service as publish_service
from planning.verification import verify_workbook
from tests.test_run_offline import (
    CODE,
    CODE_VKD,
    _actual_source,
    _single_value_source,
    _target_workbook,
)


def source_bundle():
    return {
        "actual_stock": _actual_source(),
        "factory_vikoda": _single_value_source(
            CODE,
            value_col_index=11,
            value=0,
            receipt=50,
        ),
        "factory_vkd": _single_value_source(
            CODE_VKD,
            value_col_index=11,
            value=0,
        ),
        "accounting_vikoda": _single_value_source(
            CODE,
            value_col_index=12,
            value=0,
        ),
        "accounting_vkd": _single_value_source(
            CODE_VKD,
            value_col_index=12,
            value=0,
        ),
    }


def revision():
    return {
        "mode": "production_readiness_fixture",
        "target": {"etag": "readiness-target"},
        "sources": {
            key: {"etag": f"readiness-{key}"}
            for key in source_bundle()
        },
    }


class ReleaseSafetyTests(unittest.TestCase):
    def test_workbook_round_trip_preserves_verified_planning(self):
        target = _target_workbook()
        sources = source_bundle()
        final_bytes, report, _ = prepare_pipeline_output(
            target,
            sources,
            runtime_state={},
            input_revision=revision(),
        )

        # Simulate an Excel/openpyxl open-save round trip.
        workbook = load_workbook(
            BytesIO(final_bytes),
            data_only=False,
        )
        try:
            output = BytesIO()
            workbook.save(output)
            round_tripped = output.getvalue()
        finally:
            workbook.close()

        verified = verify_workbook(
            round_tripped,
            schedule_report=report,
        )
        self.assertEqual(
            verified["plan_month"],
            9,
        )
        self.assertIn(
            verified["publish_status"],
            {"ready_for_publish", "review_required"},
        )

    def test_identical_snapshot_has_deterministic_stable_output(self):
        target = _target_workbook()
        sources = source_bundle()
        input_revision = revision()

        first_bytes, first_report, first_state = prepare_pipeline_output(
            target,
            sources,
            runtime_state={},
            input_revision=copy.deepcopy(input_revision),
        )
        second_bytes, second_report, second_state = prepare_pipeline_output(
            target,
            sources,
            runtime_state={},
            input_revision=copy.deepcopy(input_revision),
        )

        self.assertEqual(
            proposal_api.proposal_output_sha256(first_bytes),
            proposal_api.proposal_output_sha256(second_bytes),
        )

        first_identity = proposal_api.with_proposal_identity(
            first_report,
            first_bytes,
        )
        second_identity = proposal_api.with_proposal_identity(
            second_report,
            second_bytes,
        )
        self.assertEqual(
            first_identity["proposal_id"],
            second_identity["proposal_id"],
        )
        self.assertEqual(
            first_identity["proposal_output_sha256"],
            second_identity["proposal_output_sha256"],
        )

        # Business output and proposed state must be deterministic even if a
        # ZIP metadata timestamp changes the raw workbook hash.
        for key in (
            "algorithm",
            "plan_month",
            "mass_balance",
            "status",
            "publish_status",
            "pipeline",
        ):
            left = copy.deepcopy(first_report.get(key))
            right = copy.deepcopy(second_report.get(key))
            if key == "pipeline":
                # The pipeline verify/output sections may carry raw ZIP hashes;
                # stable proposal identity above is the release contract.
                for value in (left, right):
                    if isinstance(value, dict):
                        value.pop("output_sha256", None)
            self.assertEqual(left, right, key)

        self.assertEqual(first_state, second_state)

    def test_publish_proposal_dry_run_never_uploads_or_saves_state(self):
        target = _target_workbook()
        sources = source_bundle()

        class FakeGraph:
            def __init__(self):
                self.uploads = []

            def get_item_by_path(self, drive_id, path):
                import stock
                from planning.publish.constants import SOURCES

                if path == stock.DEST_PATH:
                    return {
                        "id": "target-id",
                        "eTag": "target-etag",
                        "lastModifiedDateTime": "2026-09-23T00:00:00Z",
                    }
                key = next(
                    key
                    for key, source_path in SOURCES.items()
                    if source_path == path
                )
                return {
                    "id": key,
                    "eTag": f"{key}-etag",
                    "lastModifiedDateTime": "2026-09-23T00:00:00Z",
                }

            def download_file(self, drive_id, item_id):
                if item_id == "target-id":
                    return target
                return sources[item_id]

            def upload_file(self, *args, **kwargs):
                self.uploads.append((args, kwargs))
                raise AssertionError(
                    "Proposal/dry-run gate must never upload SharePoint"
                )

        graph = FakeGraph()
        artifacts = []
        saved_states = []
        saved_decisions = []

        with (
            patch.object(
                publish_service.metrics,
                "load_runtime_state",
                return_value={},
            ),
            patch.object(
                publish_service,
                "save_proposal_artifacts",
                side_effect=lambda data, report: artifacts.append(
                    (data, copy.deepcopy(report))
                ),
            ),
            patch.object(
                publish_service,
                "save_states_after_success",
                side_effect=lambda *args, **kwargs: saved_states.append(
                    (args, kwargs)
                ),
            ),
            patch.object(
                publish_service,
                "save_publish_decision",
                side_effect=lambda *args, **kwargs: saved_decisions.append(
                    (args, kwargs)
                ),
            ),
            patch.object(
                publish_service,
                "print_operational_report",
                return_value=None,
            ),
        ):
            result = publish_service.run_pipeline_with_retry(
                graph,
                "drive",
                publish_mode="proposal",
                publisher="production-readiness",
                max_attempts=1,
            )

        self.assertEqual(result["publish_mode"], "proposal")
        self.assertFalse(result["uploaded"])
        self.assertFalse(result["publish_blocked"])
        self.assertEqual(graph.uploads, [])
        self.assertEqual(saved_states, [])
        self.assertEqual(saved_decisions, [])
        self.assertEqual(len(artifacts), 1)
        self.assertEqual(
            artifacts[0][1]["publish_decision"]["state"],
            "proposal_only",
        )
        self.assertTrue(artifacts[0][1].get("proposal_id"))


if __name__ == "__main__":
    unittest.main()
