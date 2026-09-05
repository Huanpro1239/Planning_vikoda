import posixpath
import zipfile
from io import BytesIO

from lxml import etree


def _find_sheet_xml_path(archive, sheet_name):
    workbook_root = etree.fromstring(archive.read("xl/workbook.xml"))
    relationship_ns = (
        "http://schemas.openxmlformats.org/"
        "officeDocument/2006/relationships"
    )

    relationship_id = None
    for sheet in workbook_root.xpath('//*[local-name()="sheet"]'):
        if sheet.get("name") == sheet_name:
            relationship_id = sheet.get(f"{{{relationship_ns}}}id")
            break

    if not relationship_id:
        raise RuntimeError(
            f"Không tìm thấy sheet {sheet_name!r} trong file đích."
        )

    rels_root = etree.fromstring(
        archive.read("xl/_rels/workbook.xml.rels")
    )

    target = None
    for rel in rels_root.xpath('//*[local-name()="Relationship"]'):
        if rel.get("Id") == relationship_id:
            target = rel.get("Target")
            break

    if not target:
        raise RuntimeError(
            f"Không xác định được XML của sheet {sheet_name!r}."
        )

    if target.startswith("/"):
        return target.lstrip("/")

    return posixpath.normpath(posixpath.join("xl", target))


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
