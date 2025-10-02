from __future__ import annotations

import os
import shutil
import tempfile
from pathlib import Path
from typing import Dict, Iterable, List, Tuple
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZipInfo


HEADER = ["课程名", "分数", "学分", "学时", "学时单位", "课程类别", "学期"]

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NSMAP = {"main": NS_MAIN}

# Ensure XML namespaces exist in output
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
ET.register_namespace("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("etc", "http://www.wps.cn/officeDocument/2017/etCustomData")

SHEET_NS_ATTRS = {
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xmlns:xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xmlns:etc": "http://www.wps.cn/officeDocument/2017/etCustomData",
}


def write_to_template(
    template_path: Path, records: Iterable[Dict[str, str]], output_path: Path
) -> None:
    """Render rows into the Excel template and write to ``output_path``.

    The template is copied byte-for-byte, and only the ``sharedStrings.xml``
    and ``xl/worksheets/sheet1.xml`` payloads are regenerated to include the
    provided rows while preserving styles and metadata.
    """
    rows = list(records)
    shutil.copyfile(template_path, output_path)

    sheet_xml, shared_strings_xml = _render_sheet(template_path, rows)
    _replace_zip_entries(
        output_path,
        {
            "xl/worksheets/sheet1.xml": sheet_xml,
            "xl/sharedStrings.xml": shared_strings_xml,
        },
    )


def _clone_zipinfo(info: ZipInfo, filename: str | None = None) -> ZipInfo:
    clone = ZipInfo(filename or info.filename)
    clone.date_time = info.date_time
    clone.compress_type = info.compress_type
    clone.comment = info.comment
    clone.extra = info.extra
    clone.internal_attr = info.internal_attr
    clone.external_attr = info.external_attr
    clone.create_system = info.create_system
    clone.create_version = info.create_version
    clone.extract_version = info.extract_version
    clone.flag_bits = info.flag_bits
    clone.volume = info.volume
    return clone


def _render_sheet(
    template_path: Path, records: List[Dict[str, str]]
) -> Tuple[bytes, bytes]:
    with ZipFile(template_path) as z:
        sheet_root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        shared_root = ET.fromstring(z.read("xl/sharedStrings.xml"))

    # Ensure namespaces are present
    for attr, value in SHEET_NS_ATTRS.items():
        if attr not in sheet_root.attrib:
            sheet_root.set(attr, value)

    sheet_data = sheet_root.find("main:sheetData", NSMAP)
    if sheet_data is None:
        raise ValueError("Template sheet missing sheetData node")

    rows = sheet_data.findall("main:row", NSMAP)
    if not rows:
        raise ValueError("Template sheet missing header row")

    header_row = rows[0]

    # Reset sheetData to preserve the header only
    sheet_data.clear()
    sheet_data.append(header_row)

    column_order = ["课程名", "分数", "学分", "学时", "学时单位", "课程类别", "学期"]
    col_letters = ["A", "B", "C", "D", "E", "F", "G"]
    column_styles = {"A": "5", "B": "5", "C": "5", "D": "5", "E": "6", "F": "5", "G": "7"}

    shared_strings = [
        (si.find("main:t", NSMAP).text or "") for si in shared_root.findall("main:si", NSMAP)
    ]
    string_to_index = {text: idx for idx, text in enumerate(shared_strings)}
    total_count = int(shared_root.attrib.get("count", len(shared_strings)))

    for idx, record in enumerate(records, start=2):
        row_el = ET.Element(f"{{{NS_MAIN}}}row", {"r": str(idx), "spans": "1:7"})
        for letter, key in zip(col_letters, column_order):
            value = record.get(key, "")
            if value is None:
                value = ""
            cell_attrs = {"r": f"{letter}{idx}", "s": column_styles.get(letter, "5")}
            if value == "":
                ET.SubElement(row_el, f"{{{NS_MAIN}}}c", cell_attrs)
                continue

            value_str = str(value)
            if value_str not in string_to_index:
                string_to_index[value_str] = len(shared_strings)
                shared_strings.append(value_str)
            index = string_to_index[value_str]
            total_count += 1

            cell_attrs["t"] = "s"
            cell_el = ET.SubElement(row_el, f"{{{NS_MAIN}}}c", cell_attrs)
            v_el = ET.SubElement(cell_el, f"{{{NS_MAIN}}}v")
            v_el.text = str(index)
        sheet_data.append(row_el)

    dimension = sheet_root.find("main:dimension", NSMAP)
    last_row = len(records) + 1 if records else 1
    if dimension is not None:
        dimension.set("ref", f"A1:G{last_row}")

    new_shared_root = ET.Element(
        f"{{{NS_MAIN}}}sst",
        {
            "count": str(total_count),
            "uniqueCount": str(len(shared_strings)),
        },
    )
    for text in shared_strings:
        si_el = ET.SubElement(new_shared_root, f"{{{NS_MAIN}}}si")
        t_el = ET.SubElement(si_el, f"{{{NS_MAIN}}}t")
        t_el.text = text

    sheet_fragment = ET.tostring(sheet_root, encoding="unicode", xml_declaration=False)
    shared_fragment = ET.tostring(new_shared_root, encoding="unicode", xml_declaration=False)

    prefix = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
    sheet_bytes = (prefix + sheet_fragment).encode("utf-8")
    shared_bytes = (prefix + shared_fragment).encode("utf-8")

    return sheet_bytes, shared_bytes


def _replace_zip_entries(zip_path: Path, replacements: Dict[str, bytes]) -> None:
    with ZipFile(zip_path) as existing_zip:
        infos = existing_zip.infolist()
        data = {info.filename: existing_zip.read(info.filename) for info in infos}

    data.update(replacements)

    with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
        tmp_path = Path(tmp_file.name)

    with ZipFile(tmp_path, "w") as rewritten_zip:
        for info in infos:
            name = info.filename
            if name not in data:
                continue
            rewritten_zip.writestr(_clone_zipinfo(info), data.pop(name))

        for name, payload in data.items():
            info = ZipInfo(name)
            rewritten_zip.writestr(info, payload)

    os.replace(tmp_path, zip_path)

