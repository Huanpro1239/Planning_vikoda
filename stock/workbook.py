"""Ton_kho OpenXML patching for finished-goods stock."""

import zipfile
from io import BytesIO

from lxml import etree

from excel.workbook_xml import (
    find_sheet_xml_path,
    load_shared_strings,
    read_cell_text,
    set_numeric_cell,
)

from .constants import DEST_SHEET, MASTER_SHEET
from .values import clean_number, normalize_code


def patch_destination_workbook(
    dest_bytes,
    *,
    actual_stock,
    factory_vikoda,
    factory_vkd,
    accounting_vikoda,
    accounting_vkd,
    conversion_factors,
):
    source_buffer = BytesIO(dest_bytes)
    output_buffer = BytesIO()

    counters = {
        "D": 0,
        "E": 0,
        "F": 0,
        "G": 0,
        "H": 0,
    }

    with zipfile.ZipFile(source_buffer, "r") as source_zip:
        sheet_path = find_sheet_xml_path(
            source_zip,
            DEST_SHEET,
        )
        shared_strings = load_shared_strings(source_zip)
        sheet_root = etree.fromstring(
            source_zip.read(sheet_path)
        )

        seen_dest_codes = set()
        rows = sheet_root.xpath(
            '//*[local-name()="sheetData"]/*[local-name()="row"]'
        )

        for row_element in rows:
            row_number_text = row_element.get("r")
            if not row_number_text:
                continue

            row_number = int(row_number_text)
            code = None
            for cell in row_element.xpath('./*[local-name()="c"]'):
                if cell.get("r") == f"A{row_number}":
                    code = normalize_code(
                        read_cell_text(cell, shared_strings)
                    )
                    break

            if not code:
                continue
            if code in seen_dest_codes:
                raise RuntimeError(
                    f"Mã {code} bị lặp trong {DEST_SHEET}!A."
                )
            seen_dest_codes.add(code)

            if code not in conversion_factors:
                raise RuntimeError(
                    f"Không tìm thấy quy cách cho mã {code} "
                    f"trong {MASTER_SHEET}!A:I."
                )

            factor = conversion_factors[code]

            if code in actual_stock:
                value_d = actual_stock[code]
                set_numeric_cell(
                    row_element,
                    row_number,
                    "D",
                    value_d,
                )
                counters["D"] += 1

            value_e = factory_vikoda.get(code, 0)
            set_numeric_cell(
                row_element,
                row_number,
                "E",
                value_e,
            )
            counters["E"] += 1

            value_f = factory_vkd.get(code, 0)
            set_numeric_cell(
                row_element,
                row_number,
                "F",
                value_f,
            )
            counters["F"] += 1

            if code in accounting_vikoda:
                value_g = clean_number(
                    accounting_vikoda[code] / factor - value_e
                )
                set_numeric_cell(
                    row_element,
                    row_number,
                    "G",
                    value_g,
                )
                counters["G"] += 1
            else:
                value_g = "giữ cũ"

            if code in accounting_vkd:
                value_h = clean_number(
                    accounting_vkd[code] / factor - value_f
                )
                set_numeric_cell(
                    row_element,
                    row_number,
                    "H",
                    value_h,
                )
                counters["H"] += 1
            else:
                value_h = "giữ cũ"

            print(
                f"{code}: Q={factor}; "
                f"D={actual_stock.get(code, 'giữ cũ')}; "
                f"E={value_e}; F={value_f}; "
                f"G={value_g}; H={value_h}"
            )

        if not seen_dest_codes:
            raise RuntimeError(
                f"Không đọc được mã sản phẩm trong {DEST_SHEET}!A."
            )

        for column in ("D", "E", "F", "G", "H"):
            if counters[column] == 0:
                raise RuntimeError(
                    f"Không cập nhật được cột {column} của {DEST_SHEET}."
                )

        print(
            f"[{DEST_SHEET}] Số dòng cập nhật: "
            + ", ".join(
                f"{column}={count}"
                for column, count in counters.items()
            )
        )

        new_sheet_xml = etree.tostring(
            sheet_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        with zipfile.ZipFile(
            output_buffer,
            "w",
            compression=zipfile.ZIP_DEFLATED,
        ) as output_zip:
            for item in source_zip.infolist():
                data = (
                    new_sheet_xml
                    if item.filename == sheet_path
                    else source_zip.read(item.filename)
                )
                output_zip.writestr(item, data)

    return output_buffer.getvalue()
