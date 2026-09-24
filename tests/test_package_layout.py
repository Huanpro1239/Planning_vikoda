import unittest


class PackageLayoutTests(unittest.TestCase):
    def test_sync_stock_reexports_canonical_graph_api(self):
        import sharepoint.auth as auth
        import sharepoint.client as packaged
        import sync_stock as legacy

        self.assertIs(packaged.GraphClient, legacy.GraphClient)
        self.assertIs(packaged.GraphRequestError, legacy.GraphRequestError)
        self.assertIs(packaged.is_retryable_graph_error, legacy.is_retryable_graph_error)
        self.assertIs(packaged.get_access_token, auth.get_access_token)

    def test_sync_stock_reexports_canonical_business_api(self):
        import stock
        import sync_stock as legacy

        self.assertIs(legacy.normalize_code, stock.normalize_code)
        self.assertIs(legacy.to_number, stock.to_number)
        self.assertIs(legacy.clean_number, stock.clean_number)
        self.assertIs(legacy.read_actual_stock, stock.read_actual_stock)
        self.assertIs(
            legacy.read_single_value_source,
            stock.read_single_value_source,
        )
        self.assertIs(
            legacy.read_conversion_factors,
            stock.read_conversion_factors,
        )
        self.assertIs(
            legacy.patch_destination_workbook,
            stock.patch_destination_workbook,
        )

    def test_planning_root_entrypoints_point_to_canonical_modules(self):
        import planning.calendar as calendar
        import planning.fc as fc
        import planning.publish as publish
        import planning.stock_inputs as stock_inputs

        self.assertEqual(fc.prepare_planning_fc_update.__module__, "planning.fc")
        self.assertEqual(
            calendar.prepare_calendar_update_all_months.__module__,
            "planning.calendar",
        )
        self.assertEqual(
            stock_inputs.prepare_stock_input_update.__module__,
            "planning.stock_inputs",
        )
        self.assertEqual(
            publish.run_pipeline_with_retry.__module__,
            "planning.publish.service",
        )

    def test_nvl_facades_expose_legacy_business_api(self):
        import nvl.stock as stock
        import nvl.open_po as open_po

        self.assertEqual(
            set(stock.__all__),
            {
                "NVLConfig",
                "NVLReconcileResult",
                "load_nvl_config",
                "normalize_nvl_code",
                "parse_nvl_quantity",
                "patch_nvl_destination_workbook",
                "run_nvl_sync",
                "verify_nvl_patched_workbook",
            },
        )
        self.assertTrue(callable(stock.run_nvl_sync))
        self.assertEqual(
            set(open_po.__all__),
            {
                "OpenPOConfig",
                "load_open_po_config",
                "patch_target_workbook",
                "read_open_po",
                "reconcile_target",
                "run_online",
                "verify_target",
            },
        )
        self.assertTrue(callable(open_po.read_open_po))

    def test_sync_nvl_open_po_reexports_canonical_business_api(self):
        import nvl.open_po as packaged
        import sync_nvl_open_po as legacy

        self.assertIs(legacy.OpenPOConfig, packaged.OpenPOConfig)
        self.assertIs(legacy.load_open_po_config, packaged.load_open_po_config)
        self.assertIs(legacy.read_open_po, packaged.read_open_po)
        self.assertIs(legacy.reconcile_target, packaged.reconcile_target)
        self.assertIs(legacy.patch_target_workbook, packaged.patch_target_workbook)
        self.assertIs(legacy.verify_target, packaged.verify_target)
        self.assertIs(legacy.run_online, packaged.run_online)


    def test_planning_facade_imports(self):
        import planning.pipeline as pipeline

        self.assertEqual(
            set(pipeline.__all__),
            {"SOURCE_KEYS", "prepare_pipeline_output"},
        )
        self.assertTrue(callable(pipeline.prepare_pipeline_output))
        self.assertTrue(hasattr(pipeline, "_planning_snapshot"))


if __name__ == "__main__":
    unittest.main()
