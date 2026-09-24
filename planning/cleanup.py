"""Planning workbook cleanup primitives.

This module owns Planning-specific workbook cleanup while reusing the shared
OpenXML sheet resolver from :mod:`excel.workbook_xml`.
"""

from io import BytesIO
import zipfile

from lxml import etree

from excel.workbook_xml import find_sheet_xml_path as _find_sheet_xml_path


PLANNING_SHEET = "Ke_hoach_SX"

__all__ = [
    "PLANNING_SHEET",
    "remove_sheet_formulas",
]


def remove_sheet_formulas(workbook_bytes, sheet_name):
    """Clear formula/value payloads only in one worksheet XML."""
    source_buffer = BytesIO(workbook_bytes)
    output_buffer = BytesIO()

    with zipfile.ZipFile(source_buffer, "r") as source_zip:
        sheet_path = _find_sheet_xml_path(source_zip, sheet_name)
        sheet_root = etree.fromstring(source_zip.read(sheet_path))

        removed = 0
        for cell in sheet_root.xpath('//*[local-name()="c"]'):
            formula_nodes = cell.xpath('./*[local-name()="f"]')
            if not formula_nodes:
                continue

            for child in list(cell):
                local_name = etree.QName(child).localname
                if local_name in {"f", "v"}:
                    cell.remove(child)

            cell.attrib.pop("t", None)
            removed += 1

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

    return output_buffer.getvalue(), removed
