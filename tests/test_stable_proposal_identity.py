import io
import unittest
import zipfile

import sync_planning_pipeline as pipeline_runner


class StableProposalIdentityTests(unittest.TestCase):
    @staticmethod
    def _xlsx_bytes(core_xml: bytes, sheet_xml: bytes, *, timestamp):
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, payload in (
                ("docProps/core.xml", core_xml),
                ("xl/worksheets/sheet1.xml", sheet_xml),
                ("[Content_Types].xml", b"<Types/>")
            ):
                info = zipfile.ZipInfo(name, date_time=timestamp)
                archive.writestr(info, payload)
        return out.getvalue()

    @staticmethod
    def _report():
        return {
            "algorithm": "ke_hoach_sx_tuan_v1",
            "plan_month": "2026-09",
            "input_revision": {
                "target": {"etag": "same-etag", "sha256": "same-input"},
                "sources": {"fc": {"etag": "same-source"}},
            },
        }

    def test_core_metadata_and_zip_timestamp_do_not_change_proposal_id(self):
        first = self._xlsx_bytes(
            b"<core><modified>2026-09-07T02:51:22Z</modified></core>",
            b"<sheet><v>123</v></sheet>",
            timestamp=(2026, 9, 7, 2, 51, 22),
        )
        second = self._xlsx_bytes(
            b"<core><modified>2026-09-07T03:00:40Z</modified></core>",
            b"<sheet><v>123</v></sheet>",
            timestamp=(2026, 9, 7, 3, 0, 40),
        )

        first_report = pipeline_runner._with_proposal_identity(self._report(), first)
        second_report = pipeline_runner._with_proposal_identity(self._report(), second)

        self.assertNotEqual(first_report["output_sha256"], second_report["output_sha256"])
        self.assertEqual(
            first_report["proposal_output_sha256"],
            second_report["proposal_output_sha256"],
        )
        self.assertEqual(first_report["proposal_id"], second_report["proposal_id"])

    def test_business_payload_change_changes_proposal_id(self):
        first = self._xlsx_bytes(
            b"<core><modified>A</modified></core>",
            b"<sheet><v>123</v></sheet>",
            timestamp=(2026, 9, 7, 2, 51, 22),
        )
        second = self._xlsx_bytes(
            b"<core><modified>B</modified></core>",
            b"<sheet><v>124</v></sheet>",
            timestamp=(2026, 9, 7, 3, 0, 40),
        )

        first_report = pipeline_runner._with_proposal_identity(self._report(), first)
        second_report = pipeline_runner._with_proposal_identity(self._report(), second)

        self.assertNotEqual(
            first_report["proposal_output_sha256"],
            second_report["proposal_output_sha256"],
        )
        self.assertNotEqual(first_report["proposal_id"], second_report["proposal_id"])

    def test_non_xlsx_payload_keeps_raw_hash_behavior(self):
        report = pipeline_runner._with_proposal_identity(self._report(), b"plain-bytes")
        self.assertEqual(report["output_sha256"], report["proposal_output_sha256"])


if __name__ == "__main__":
    unittest.main()
