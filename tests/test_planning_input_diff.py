import unittest
from unittest.mock import patch, MagicMock

import sync_planning_pipeline as pipeline_runner
import sync_stock


class PlanningInputDiffTests(unittest.TestCase):
    def test_detect_input_changes_identifies_all_components(self):
        old_state = {
            "sources": {
                "actual_stock": "etag_act_1",
                "factory_vikoda": "etag_fac_1",
                "factory_vkd": "etag_vkd_1",
                "accounting_vikoda": "etag_acc1_1",
                "accounting_vkd": "etag_acc2_1",
            },
            "conversion_hash": "conv_hash_1",
            "fc_hash": "fc_hash_1",
        }

        source_items = {k: {"eTag": v} for k, v in old_state["sources"].items()}
        pipeline_info = {"conversion_hash": "conv_hash_1", "fc_hash": "fc_hash_1"}
        self.assertEqual(
            pipeline_runner.detect_input_changes(old_state, source_items, pipeline_info),
            [],
        )

        # FC changed
        pipeline_info_fc = {"conversion_hash": "conv_hash_1", "fc_hash": "fc_hash_2"}
        self.assertEqual(
            pipeline_runner.detect_input_changes(old_state, source_items, pipeline_info_fc),
            ["sheet:FC"],
        )

        # Danh_muc changed
        pipeline_info_conv = {"conversion_hash": "conv_hash_2", "fc_hash": "fc_hash_1"}
        self.assertEqual(
            pipeline_runner.detect_input_changes(old_state, source_items, pipeline_info_conv),
            ["sheet:Danh_muc"],
        )

        # Inventory source changed
        source_items_changed = dict(source_items)
        source_items_changed["actual_stock"] = {"eTag": "etag_act_2"}
        self.assertEqual(
            pipeline_runner.detect_input_changes(old_state, source_items_changed, pipeline_info),
            ["source:actual_stock"],
        )

        # engine_version changed
        pipeline_info_engine = {
            "conversion_hash": "conv_hash_1",
            "fc_hash": "fc_hash_1",
            "engine_version": "v_new",
        }
        old_state_engine = dict(old_state)
        old_state_engine["engine_version"] = "v_old"
        self.assertEqual(
            pipeline_runner.detect_input_changes(old_state_engine, source_items, pipeline_info_engine),
            ["engine_version"],
        )

    def test_save_state_preserves_fc_hash(self):
        from pathlib import Path
        import tempfile

        with tempfile.TemporaryDirectory() as tmpdir:
            test_state_file = Path(tmpdir) / "state.json"
            with patch.object(sync_stock, "STATE_FILE", test_state_file):
                sync_stock.save_state(
                    {"actual_stock": "etag1"},
                    "conv123",
                    fc_hash="fchash456",
                )
                loaded = sync_stock.load_state()
                self.assertEqual(loaded.get("fc_hash"), "fchash456")
                self.assertEqual(loaded.get("conversion_hash"), "conv123")
                self.assertEqual(loaded.get("sources", {}).get("actual_stock"), "etag1")

    def test_skip_if_unchanged_skips_when_no_inputs_changed(self):
        fake_graph = MagicMock()
        fake_graph.get_item_by_path.return_value = {
            "id": "target_id",
            "eTag": "etag_same",
            "lastModifiedDateTime": "2026-09-07T08:00:00Z",
        }
        fake_graph.download_file.return_value = b"workbook_data"

        old_state = {
            "sources": {k: "etag_same" for k in pipeline_runner.SOURCES},
            "conversion_hash": "conv_same",
            "fc_hash": "fc_same",
        }

        mock_report = {
            "publish_status": "ready_for_publish",
            "algorithm": "test",
            "plan_month": "2026-09",
            "status": {
                "monthly_quantity": {"ok": True},
                "resource_validation": {"ok": True},
                "service": {"ok": True},
            },
            "pipeline": {
                "conversion_hash": "conv_same",
                "fc_hash": "fc_same",
            },
        }

        with (
            patch.object(pipeline_runner, "prepare_pipeline_output", return_value=(b"final_data", mock_report, {})),
            patch.object(sync_stock, "load_state", return_value=old_state),
            patch.object(pipeline_runner.metrics, "load_runtime_state", return_value={}),
            patch.object(pipeline_runner, "_save_proposal_artifacts"),
            patch.object(pipeline_runner, "_save_publish_decision"),
            patch.object(pipeline_runner, "_save_states_after_success"),
            patch.object(pipeline_runner, "print_operational_report"),
        ):
            result = pipeline_runner.run_pipeline_with_retry(
                fake_graph,
                "drive_id",
                publish_mode="publish",
                skip_if_unchanged=True,
                force=False,
            )
            self.assertTrue(result.get("skipped_unchanged"))
            self.assertFalse(result.get("uploaded"))
            fake_graph.upload_file.assert_not_called()

            result_force = pipeline_runner.run_pipeline_with_retry(
                fake_graph,
                "drive_id",
                publish_mode="publish",
                skip_if_unchanged=True,
                force=True,
            )
            self.assertFalse(result_force.get("skipped_unchanged", False))
            self.assertTrue(result_force.get("uploaded"))
            fake_graph.upload_file.assert_called_once()

    def test_skip_if_unchanged_does_not_skip_when_engine_version_changes(self):
        fake_graph = MagicMock()
        fake_graph.get_item_by_path.return_value = {
            "id": "target_id",
            "eTag": "etag_same",
            "lastModifiedDateTime": "2026-09-07T08:00:00Z",
        }
        fake_graph.download_file.return_value = b"workbook_data"

        old_state = {
            "sources": {k: "etag_same" for k in pipeline_runner.SOURCES},
            "conversion_hash": "conv_same",
            "fc_hash": "fc_same",
            "engine_version": "old_engine_v3",
        }

        mock_report = {
            "publish_status": "ready_for_publish",
            "algorithm": "new_engine_v4",
            "plan_month": "2026-09",
            "status": {
                "monthly_quantity": {"ok": True},
                "resource_validation": {"ok": True},
                "service": {"ok": True},
            },
            "pipeline": {
                "conversion_hash": "conv_same",
                "fc_hash": "fc_same",
                "engine_version": "new_engine_v4",
            },
        }

        with (
            patch.object(pipeline_runner, "prepare_pipeline_output", return_value=(b"new_data", mock_report, {})),
            patch.object(sync_stock, "load_state", return_value=old_state),
            patch.object(pipeline_runner.metrics, "load_runtime_state", return_value={}),
            patch.object(pipeline_runner, "_save_proposal_artifacts"),
            patch.object(pipeline_runner, "_save_publish_decision"),
            patch.object(pipeline_runner, "_save_states_after_success"),
            patch.object(pipeline_runner, "print_operational_report"),
        ):
            result = pipeline_runner.run_pipeline_with_retry(
                fake_graph,
                "drive_id",
                publish_mode="publish",
                skip_if_unchanged=True,
                force=False,
            )
            self.assertFalse(result.get("skipped_unchanged", False))
            self.assertTrue(result.get("uploaded"))
            fake_graph.upload_file.assert_called_once()


if __name__ == "__main__":
    unittest.main()
