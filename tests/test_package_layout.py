import unittest


class PackageLayoutTests(unittest.TestCase):
    def test_stock_package_exposes_canonical_business_api(self):
        import stock
        import stock.service as service

        self.assertEqual(stock.normalize_code.__module__, "stock.values")
        self.assertEqual(stock.to_number.__module__, "stock.values")
        self.assertEqual(stock.read_actual_stock.__module__, "stock.readers")
        self.assertEqual(
            stock.read_conversion_factors.__module__,
            "stock.readers",
        )
        self.assertEqual(
            stock.patch_destination_workbook.__module__,
            "stock.workbook",
        )
        self.assertTrue(callable(service.main))

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
        self.assertTrue(callable(stock.main))
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

    def test_nvl_safe_entrypoint_is_canonical_package_module(self):
        import nvl.open_po_safe as safe

        self.assertTrue(callable(safe.main))
        self.assertEqual(safe.main.__module__, "nvl.open_po_safe")


    def test_planning_support_modules_are_canonical_packages(self):
        import planning.cleanup as cleanup
        import planning.schedule_report as schedule_report

        self.assertEqual(cleanup.remove_sheet_formulas.__module__, "planning.cleanup")
        self.assertEqual(schedule_report.json_safe.__module__, "planning.schedule_report")
        self.assertEqual(
            schedule_report.save_schedule_report.__module__,
            "planning.schedule_report",
        )
        self.assertEqual(
            schedule_report.load_schedule_report.__module__,
            "planning.schedule_report",
        )

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
