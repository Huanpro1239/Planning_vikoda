"""NVL workbook patching and integrity verification."""

from __future__ import annotations

from io import BytesIO
import hashlib
import re
from typing import Any
import zipfile

from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries

from excel.openpyxl_io import safe_close_workbook
from excel.workbook_xml import column_number, find_sheet_xml_path, set_numeric_cell
from nvl.models import NVLConfig, NVLReconcileResult
from nvl.values import EPS, _cells_match

def patch_nvl_destination_workbook(
    target_bytes: bytes,
    reconcile_result: NVLReconcileResult,
    config: NVLConfig,
) -> bytes:
    """Cập nhật các ô số tồn kho vào file Excel đích ở cấp độ ZIP XML:

    - Chỉ giải nén và sửa file XML của sheet Ton_NVL.
    - Chèn thẻ <row> và <c> đúng vị trí tuần tự theo XML schema của OpenXML.
    - Cập nhật <dimension ref="..."> để bao phủ các ô mới tạo nếu vượt quá phạm vi cũ.
    - Bảo toàn 100% byte-for-byte các file ZIP parts còn lại.
    """
    if not reconcile_result.changes:
        return target_bytes

    source_buffer = BytesIO(target_bytes)
    output_buffer = BytesIO()

    source_zip = zipfile.ZipFile(source_buffer, "r")
    output_zip = zipfile.ZipFile(output_buffer, "w", compression=zipfile.ZIP_DEFLATED)
    try:
        sheet_path = find_sheet_xml_path(source_zip, config.target_sheet)
        sheet_root = etree.fromstring(source_zip.read(sheet_path))

        sheet_data_nodes = sheet_root.xpath('//*[local-name()="sheetData"]')
        if not sheet_data_nodes:
            raise RuntimeError(f"Không tìm thấy thẻ sheetData trong {sheet_path}.")
        sheet_data = sheet_data_nodes[0]

        row_map: dict[int, Any] = {}
        for r_elem in sheet_data.xpath('./*[local-name()="row"]'):
            r_num_text = r_elem.get("r")
            if r_num_text and r_num_text.isdigit():
                row_map[int(r_num_text)] = r_elem

        changes_by_row = {item["row"]: item["after"] for item in reconcile_result.changes}

        for row_num, new_val in changes_by_row.items():
            row_element = row_map.get(row_num)
            if row_element is None:
                namespace = etree.QName(sheet_data).namespace
                row_element = etree.Element(f"{{{namespace}}}row", r=str(row_num))
                inserted = False
                for existing_row in sheet_data.xpath('./*[local-name()="row"]'):
                    er = existing_row.get("r")
                    if er and int(er) > row_num:
                        existing_row.addprevious(row_element)
                        inserted = True
                        break
                if not inserted:
                    sheet_data.append(row_element)
                row_map[row_num] = row_element

            set_numeric_cell(
                row_element,
                row_num,
                config.target_value_col_letter,
                new_val,
            )

            spans_attr = row_element.get("spans")
            if spans_attr and ":" in spans_attr:
                try:
                    s_min, s_max = map(int, spans_attr.split(":"))
                    val_col_num = column_number(config.target_value_col_letter)
                    if val_col_num > s_max:
                        row_element.set("spans", f"{s_min}:{val_col_num}")
                    elif val_col_num < s_min:
                        row_element.set("spans", f"{val_col_num}:{s_max}")
                except Exception:
                    pass

        # Cập nhật <dimension ref="..."> nếu cần mở rộng phạm vi
        dim_nodes = sheet_root.xpath('//*[local-name()="dimension"]')
        if dim_nodes:
            dim_elem = dim_nodes[0]
            ref = dim_elem.get("ref", "")
            if ref and ":" in ref:
                try:
                    min_c, min_r, max_c, max_r = range_boundaries(ref)
                    val_col_num = column_number(config.target_value_col_letter)
                    max_c_new = max(max_c, val_col_num)
                    max_r_new = max([max_r] + list(changes_by_row.keys()))
                    if max_c_new != max_c or max_r_new != max_r:
                        dim_elem.set("ref", f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c_new)}{max_r_new}")
                except Exception:
                    pass

        new_sheet_xml = etree.tostring(
            sheet_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        for name in source_zip.namelist():
            data = new_sheet_xml if name == sheet_path else source_zip.read(name)
            output_zip.writestr(name, data)
    finally:
        output_zip.close()
        output_zip.fp = None
        source_zip.close()
        source_zip.fp = None

    return output_buffer.getvalue()


SHAREPOINT_METADATA_NAMESPACES = (
    "http://schemas.microsoft.com/office/2006/metadata/contentType",
    "http://schemas.microsoft.com/office/2006/metadata/properties",
    "http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes",
    "http://schemas.microsoft.com/sharepoint/v3/contenttype/forms",
    "http://schemas.microsoft.com/sharepoint/",
)


def identify_sharepoint_metadata_exemption(
    name: str,
    orig_bytes: bytes,
    patch_bytes: bytes,
    z_orig: zipfile.ZipFile,
    z_patch: zipfile.ZipFile,
) -> dict[str, str] | None:
    """Xác thực một phần tử trong file ZIP có phải là metadata SharePoint được máy chủ
    SharePoint tự động cập nhật khi upload hay không.

    Yêu cầu nhận diện chính xác theo:
    1. Cấu trúc XML và root element.
    2. Namespace chính thức của SharePoint metadata.
    3. Quan hệ liên kết trong file quan hệ (_rels/.rels, workbook.xml.rels hoặc customXml/_rels/*.rels).

    TUYỆT ĐỐI không bỏ qua toàn bộ customXml hay docProps:
    - Nếu phần tử không chứa cấu trúc/namespace SharePoint chuẩn -> trả về None (bị chặn).
    - Nếu phần tử nằm ngoài customXml và docProps (như xl/worksheets, xl/styles...) -> trả về None.
    """
    name_clean = name.strip("/").lower()

    # 1. Kiểm tra nếu là customXml/_rels/*.rels
    if name_clean.startswith("customxml/_rels/") and name_clean.endswith(".rels"):
        try:
            root_orig = etree.fromstring(orig_bytes)
            root_patch = etree.fromstring(patch_bytes)
            rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
            if root_orig.tag != f"{{{rel_ns}}}Relationships" or root_patch.tag != f"{{{rel_ns}}}Relationships":
                return None
            prop_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
            rels = root_orig.findall(f"{{{rel_ns}}}Relationship")
            if not rels or not all(r.get("Type") == prop_type for r in rels):
                return None
            return {
                "part": name,
                "type": "sharepoint_customxml_rels",
                "namespace": prop_type,
                "reason": "File quan hệ customXmlProps của SharePoint metadata",
            }
        except Exception:
            return None

    # 2. Kiểm tra nếu là customXml/item*.xml hoặc customXml/itemProps*.xml
    if name_clean.startswith("customxml/") and name_clean.endswith(".xml"):
        try:
            root_orig = etree.fromstring(orig_bytes)
            root_patch = etree.fromstring(patch_bytes)
        except Exception:
            return None

        # 2a. Trường hợp itemProps*.xml (ds:datastoreItem)
        ds_ns = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
        if root_orig.tag == f"{{{ds_ns}}}datastoreItem" and root_patch.tag == f"{{{ds_ns}}}datastoreItem":
            schema_refs = root_orig.findall(f".//{{{ds_ns}}}schemaRef")
            matched_ns = None
            for sref in schema_refs:
                uri = sref.get(f"{{{ds_ns}}}uri") or sref.get("uri") or ""
                if any(sp_ns in uri for sp_ns in SHAREPOINT_METADATA_NAMESPACES):
                    matched_ns = uri
                    break
            if matched_ns:
                return {
                    "part": name,
                    "type": "sharepoint_datastore_item",
                    "namespace": matched_ns,
                    "reason": f"SharePoint datastore item properties tham chiếu schema '{matched_ns}'",
                }
            return None

        # 2b. Trường hợp item*.xml (ct:contentTypeSchema, p:properties, FormTemplates...)
        orig_nsmap = root_orig.nsmap.values()
        patch_nsmap = root_patch.nsmap.values()
        root_tag_orig = root_orig.tag
        root_tag_patch = root_patch.tag

        matched_ns = None
        for sp_ns in SHAREPOINT_METADATA_NAMESPACES:
            if (
                sp_ns in root_tag_orig
                or sp_ns in root_tag_patch
                or any(sp_ns in str(v) for v in orig_nsmap)
                or any(sp_ns in str(v) for v in patch_nsmap)
            ):
                matched_ns = sp_ns
                break

        if matched_ns:
            # Kiểm tra quan hệ từ xl/_rels/workbook.xml.rels hoặc _rels/.rels
            has_valid_rel = False
            try:
                rels_checked = 0
                for rels_name in ("xl/_rels/workbook.xml.rels", "_rels/.rels"):
                    if rels_name in z_orig.namelist():
                        rels_checked += 1
                        rels_root = etree.fromstring(z_orig.read(rels_name))
                        rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
                        for rel in rels_root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                            if rel.get("Type") == rel_type:
                                target = (rel.get("Target") or "").replace("../", "").strip("/").lower()
                                if target == name_clean or target.endswith(name_clean):
                                    has_valid_rel = True
                                    break
                    if has_valid_rel:
                        break
                if rels_checked == 0:
                    has_valid_rel = True
            except Exception:
                has_valid_rel = False

            if has_valid_rel:
                tag_short = root_tag_orig.split("}")[-1] if "}" in root_tag_orig else root_tag_orig
                return {
                    "part": name,
                    "type": "sharepoint_customxml_metadata",
                    "namespace": matched_ns,
                    "reason": f"SharePoint document contentType/DIP metadata ({tag_short})",
                }
        return None

    # 3. Kiểm tra nếu là docProps/core.xml
    if name_clean == "docprops/core.xml":
        try:
            root_orig = etree.fromstring(orig_bytes)
            root_patch = etree.fromstring(patch_bytes)
            core_ns = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
            if root_orig.tag == f"{{{core_ns}}}coreProperties" and root_patch.tag == f"{{{core_ns}}}coreProperties":
                return {
                    "part": name,
                    "type": "server_core_properties",
                    "namespace": core_ns,
                    "reason": "Dublin Core properties cập nhật bởi SharePoint/Office",
                }
        except Exception:
            return None

    # 4. Kiểm tra nếu là phần tử [trash]/*.dat phát sinh do cơ chế lưu trữ phân đoạn máy chủ SharePoint
    if re.match(r"^\[trash\]/[0-9]{4,}\.dat$", name_clean):
        # 4a. Cấu trúc gói: Tuyệt đối không được khai báo trong [Content_Types].xml (nếu có là active document part)
        try:
            if "[Content_Types].xml" in z_patch.namelist():
                ct_root = etree.fromstring(z_patch.read("[Content_Types].xml"))
                for override in ct_root.findall(".//{http://schemas.openxmlformats.org/package/2006/content-types}Override"):
                    if (override.get("PartName") or "").strip("/").lower() == name_clean:
                        return None
                for default_type in ct_root.findall(".//{http://schemas.openxmlformats.org/package/2006/content-types}Default"):
                    if (default_type.get("Extension") or "").lower() == "dat":
                        return None
        except Exception:
            return None

        # 4b. Quan hệ tham chiếu: Tuyệt đối không có bất kỳ file .rels nào trong gói tham chiếu đến
        try:
            target_basename = name_clean.split("/")[-1]
            for rel_file in z_patch.namelist():
                if rel_file.endswith(".rels"):
                    rel_root = etree.fromstring(z_patch.read(rel_file))
                    for rel in rel_root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                        tgt = (rel.get("Target") or "").lower()
                        if "trash" in tgt or target_basename in tgt:
                            return None
        except Exception:
            return None

        # 4c. Định dạng nội dung: Padding nhị phân chuẩn máy chủ SharePoint (\xff\xff\xff\xff + byte 0 hoặc 255)
        #     Tuyệt đối không phải văn bản XML, HTML hay binary thực thi (không chứa header PE MZ / ELF)
        if patch_bytes.startswith(b"\xff\xff\xff\xff") and set(patch_bytes) <= {0, 255}:
            return {
                "part": name,
                "type": "server_trash_dat_part",
                "namespace": "urn:schemas-microsoft-com:sharepoint:storage:trash",
                "reason": "Phần tử [trash]/*.dat phát sinh do cơ chế lưu trữ phân đoạn máy chủ SharePoint (không tham chiếu trong OOXML package)",
            }
        return None

    return None


def verify_nvl_patched_workbook(
    original_bytes: bytes,
    patched_bytes: bytes,
    reconcile_result: NVLReconcileResult,
    config: NVLConfig,
    *,
    is_server_comparison: bool = False,
) -> dict[str, Any]:
    """Kiểm tra tính toàn vẹn của workbook sau khi patch XML:

    1. Tất cả các ô changed_cells đều nhận đúng giá trị số mới.
    2. Các ô unchanged và missing_in_source giữ nguyên giá trị ban đầu (hỗ trợ cả văn bản/blank/lỗi).
    3. Tất cả các ô không thuộc danh sách thay đổi trong chính sheet Ton_NVL giữ nguyên định dạng và nội dung XML.
    4. Tất cả các phần tử trong file ZIP ngoài sheet đích đều giống nhau từng byte (SHA-256 đối chiếu 1-1).
       Khi đối chiếu với file tải lại từ server SharePoint (is_server_comparison=True), chỉ miễn trừ các metadata
       được xác thực chính xác là do máy chủ SharePoint tự động đóng dấu (theo cấu trúc/namespace và quan hệ liên kết).
       Ghi nhận danh sách các phần tử được miễn trừ trong kết quả trả về.
    """
    if not reconcile_result.changes:
        return {"ok": True, "message": "Không có thay đổi cần xác minh.", "exempted_parts": []}

    exempted_parts: list[dict[str, str]] = []

    # 1. So sánh các ZIP parts không liên quan (byte-level SHA-256)
    orig_buf = BytesIO(original_bytes)
    patch_buf = BytesIO(patched_bytes)
    z_orig = zipfile.ZipFile(orig_buf, "r")
    z_patch = zipfile.ZipFile(patch_buf, "r")
    try:
        target_sheet_path = find_sheet_xml_path(z_orig, config.target_sheet)
        for name in z_orig.namelist():
            if name != target_sheet_path:
                if name not in z_patch.namelist():
                    raise RuntimeError(f"Phần tử '{name}' bị thiếu trong file sau khi ghi/tải lại!")
                orig_raw = z_orig.read(name)
                patch_raw = z_patch.read(name)
                orig_hash = hashlib.sha256(orig_raw).hexdigest()
                patch_hash = hashlib.sha256(patch_raw).hexdigest()
                if orig_hash != patch_hash:
                    if is_server_comparison:
                        exemption = identify_sharepoint_metadata_exemption(
                            name, orig_raw, patch_raw, z_orig, z_patch
                        )
                        if exemption is not None:
                            exempted_parts.append(exemption)
                            continue
                    raise RuntimeError(
                        f"Phần tử không liên quan '{name}' trong file ZIP bị thay đổi ngoài ý muốn!"
                    )

        # Kiểm tra không có phần tử lạ xuất hiện trong z_patch ngoài các metadata hợp lệ
        for name in z_patch.namelist():
            if name not in z_orig.namelist() and name != target_sheet_path:
                if is_server_comparison:
                    exemption = identify_sharepoint_metadata_exemption(
                        name, b"", z_patch.read(name), z_orig, z_patch
                    )
                    if exemption is not None:
                        exempted_parts.append(exemption)
                        continue
                raise RuntimeError(
                    f"Phần tử lạ ngoài ý muốn '{name}' xuất hiện trong file ZIP sau khi ghi!"
                )

        # Kiểm tra tính toàn vẹn của các ô không thay đổi trong chính sheet đích
        orig_sheet_root = etree.fromstring(z_orig.read(target_sheet_path))
        patch_sheet_root = etree.fromstring(z_patch.read(target_sheet_path))

        changed_refs = {f"{config.target_value_col_letter}{item['row']}" for item in reconcile_result.changes}
        orig_cells = {c.get("r"): etree.tostring(c) for c in orig_sheet_root.xpath('//*[local-name()="c"]') if c.get("r")}
        patch_cells = {c.get("r"): etree.tostring(c) for c in patch_sheet_root.xpath('//*[local-name()="c"]') if c.get("r")}

        for r_coord, orig_c_xml in orig_cells.items():
            if r_coord not in changed_refs:
                if r_coord not in patch_cells:
                    raise RuntimeError(f"Ô không liên quan '{r_coord}' trong sheet '{config.target_sheet}' bị mất sau khi patch!")
                if patch_cells[r_coord] != orig_c_xml:
                    raise RuntimeError(f"Ô không liên quan '{r_coord}' trong sheet '{config.target_sheet}' bị biến đổi cấu trúc XML!")

        for r_coord in patch_cells:
            if r_coord not in orig_cells and r_coord not in changed_refs:
                raise RuntimeError(
                    f"Ô lạ ngoài ý muốn '{r_coord}' xuất hiện trong sheet '{config.target_sheet}' sau khi patch!"
                )
    finally:
        z_orig.close()
        z_orig.fp = None
        z_patch.close()
        z_patch.fp = None

    # 2. Đọc lại workbook đích bằng openpyxl để xác minh dữ liệu giá trị ô
    wb = load_workbook(BytesIO(patched_bytes), data_only=True)
    try:
        ws = wb[config.target_sheet]

        # Kiểm tra các ô thay đổi
        for item in reconcile_result.changes:
            row = item["row"]
            expected = item["after"]
            actual = ws.cell(row=row, column=config.target_value_col).value
            if actual is None or abs(float(actual) - float(expected)) > EPS:
                raise RuntimeError(
                    f"Xác minh thất bại tại dòng {row}: kỳ vọng {expected}, thực tế {actual}."
                )

        # Kiểm tra các ô không đổi (unchanged)
        for item in reconcile_result.unchanged:
            row = item["row"]
            expected = item["value"]
            actual = ws.cell(row=row, column=config.target_value_col).value
            if not _cells_match(actual, expected):
                raise RuntimeError(
                    f"Ô không đổi tại dòng {row} bị thay đổi: cũ {expected!r}, mới {actual!r}."
                )

        # Kiểm tra các ô thiếu nguồn (missing_in_source)
        for item in reconcile_result.missing_in_source:
            row = item["row"]
            expected = item["current_value"]
            actual = ws.cell(row=row, column=config.target_value_col).value
            if not _cells_match(actual, expected):
                raise RuntimeError(
                    f"Ô giữ nguyên (mã thiếu nguồn) tại dòng {row} bị thay đổi: cũ {expected!r}, mới {actual!r}."
                )
    finally:
        safe_close_workbook(wb)

    return {
        "ok": True,
        "message": "Xác minh toàn vẹn thành công 100%.",
        "exempted_parts": exempted_parts,
    }


__all__ = [
    "patch_nvl_destination_workbook",
    "identify_sharepoint_metadata_exemption",
    "verify_nvl_patched_workbook",
]
