import unittest
from io import BytesIO

from openpyxl import Workbook

import sync_planning_metrics as metrics
from sync_planning_metrics_compat import (
    MASTER_SHEET,
    read_conversion_factors_and_leadtime,
    read_leadtime_from_master,
)


class LeadtimeFromMasterTests(unittest.TestCase):
    def make_workbook(self, leadtime=4, header="Leadtime"):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = MASTER_SHEET
        worksheet.append(
            [
                "Mã Sản Phẩm",
                "Tên Sản Phẩm",
                "Đơn vị tính",
                "Số lượng /mẻ",
                "Số lượng/ ca",
                "Chuyền",
                "Nhóm sản phẩm",
                "Phân loại SP",
                "Quy cách",
                header,
            ]
        )
        worksheet.append(
            [
                130200026,
                "SUMO",
                "Thùng",
                5200,
                4000,
                "PET 9000",
                "Sumo",
                "Có đường",
                24,
                leadtime,
            ]
        )
        output = BytesIO()
        workbook.save(output)
        return output.getvalue()

    def test_reads_leadtime_from_column_j(self):
        leadtimes = read_leadtime_from_master(self.make_workbook(leadtime=4))
        self.assertEqual(leadtimes["130200026"], 4)

    def test_runtime_mapping_is_replaced_by_master_column_j(self):
        metrics.LEADTIME_BY_CODE = {"130200026": 1}
        factors, _ = read_conversion_factors_and_leadtime(
            self.make_workbook(leadtime=4)
        )
        self.assertEqual(factors["130200026"], 24)
        self.assertEqual(metrics.LEADTIME_BY_CODE["130200026"], 4)

    def test_rejects_wrong_header(self):
        with self.assertRaisesRegex(RuntimeError, "J1 phải là 'Leadtime'"):
            read_leadtime_from_master(self.make_workbook(header="LT"))

    def test_rejects_negative_leadtime(self):
        with self.assertRaisesRegex(RuntimeError, "phải >= 0"):
            read_leadtime_from_master(self.make_workbook(leadtime=-1))


if __name__ == "__main__":
    unittest.main()
