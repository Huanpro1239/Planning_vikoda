import unittest


class PackageLayoutTests(unittest.TestCase):
    def test_sharepoint_facade_preserves_graph_api(self):
        import sharepoint.client as packaged
        import sync_stock as legacy

        self.assertIs(packaged.GraphClient, legacy.GraphClient)
        self.assertIs(packaged.GraphRequestError, legacy.GraphRequestError)
        self.assertIs(packaged.get_access_token, legacy.get_access_token)

    def test_nvl_facades_expose_legacy_business_api(self):
        import nvl.stock as stock
        import nvl.open_po as open_po

        self.assertTrue(callable(stock.run_nvl_sync))
        self.assertTrue(callable(open_po.read_open_po_by_code))

    def test_planning_facade_imports(self):
        import planning.pipeline as pipeline

        self.assertTrue(hasattr(pipeline, "prepare_pipeline_output"))


if __name__ == "__main__":
    unittest.main()
