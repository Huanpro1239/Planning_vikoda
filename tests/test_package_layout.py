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

    def test_planning_facade_imports(self):
        import planning.pipeline as pipeline

        self.assertEqual(
            set(pipeline.__all__),
            {"SOURCE_KEYS", "prepare_pipeline_output"},
        )
        self.assertTrue(callable(pipeline.prepare_pipeline_output))


if __name__ == "__main__":
    unittest.main()
