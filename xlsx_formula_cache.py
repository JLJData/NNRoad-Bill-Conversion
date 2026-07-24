# -*- coding: utf-8 -*-
"""
给公式单元格写入 Excel 缓存值 <v>，保留 <f>。

openpyxl 保存后公式格通常没有缓存值；LuckySheet/部分预览依赖 <v> 才能显示，
而 Excel 打开后仍可按公式重算（改 TW-L / 汇率后重算即可联动）。
"""
from __future__ import annotations

import re
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_ODR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ET.register_namespace("", NS_MAIN)


def _q(tag: str, ns: str = NS_MAIN) -> str:
    return f"{{{ns}}}{tag}"


def _col_row(cell_ref: str) -> tuple[str, int]:
    m = re.fullmatch(r"([A-Za-z]+)(\d+)", cell_ref.strip())
    if not m:
        raise ValueError(f"invalid cell ref: {cell_ref}")
    return m.group(1).upper(), int(m.group(2))


def _sheet_path_for_name(files: dict[str, bytes], sheet_name: str) -> str | None:
    wb = ET.fromstring(files["xl/workbook.xml"])
    sheets = wb.find(_q("sheets"))
    if sheets is None:
        return None
    rid = None
    for sh in sheets.findall(_q("sheet")):
        if (sh.get("name") or "") == sheet_name:
            rid = sh.get(f"{{{NS_ODR}}}id")
            break
    if not rid:
        return None
    rels = ET.fromstring(files["xl/_rels/workbook.xml.rels"])
    for rel in rels.findall(_q("Relationship", NS_REL)):
        if rel.get("Id") == rid:
            target = (rel.get("Target") or "").lstrip("/")
            if target.startswith("xl/"):
                return target
            return "xl/" + target.lstrip("./")
    return None


def _ensure_cell(sheet_root: ET.Element, cell_ref: str) -> ET.Element:
    col, row_num = _col_row(cell_ref)
    ref = f"{col}{row_num}"
    sheet_data = sheet_root.find(_q("sheetData"))
    if sheet_data is None:
        sheet_data = ET.SubElement(sheet_root, _q("sheetData"))

    row_el = None
    for r in sheet_data.findall(_q("row")):
        if r.get("r") == str(row_num):
            row_el = r
            break
    if row_el is None:
        row_el = ET.Element(_q("row"), {"r": str(row_num)})
        inserted = False
        for r in list(sheet_data.findall(_q("row"))):
            try:
                if int(r.get("r") or "0") > row_num:
                    idx = list(sheet_data).index(r)
                    sheet_data.insert(idx, row_el)
                    inserted = True
                    break
            except ValueError:
                continue
        if not inserted:
            sheet_data.append(row_el)

    for c in row_el.findall(_q("c")):
        if (c.get("r") or "").upper() == ref:
            return c
    cell_el = ET.Element(_q("c"), {"r": ref})
    row_el.append(cell_el)
    return cell_el


def inject_formula_cached_values(
    xlsx_path: Path,
    caches: dict[str, dict[str, float]],
) -> int:
    """
    caches: { sheetName: { "E16": 123.45, ... } }
    仅给已有公式的格子补 <v>；没有公式则跳过。
    """
    xlsx_path = Path(xlsx_path)
    if not caches:
        return 0

    with ZipFile(xlsx_path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    total = 0
    for sheet_name, cell_map in caches.items():
        path = _sheet_path_for_name(files, sheet_name)
        if not path or path not in files:
            continue
        root = ET.fromstring(files[path])
        changed = False
        for cell_ref, value in cell_map.items():
            cell_el = _ensure_cell(root, cell_ref)
            f_el = cell_el.find(_q("f"))
            if f_el is None:
                continue
            if "t" in cell_el.attrib:
                del cell_el.attrib["t"]
            for bad in list(cell_el.findall(_q("is"))):
                cell_el.remove(bad)
            v_el = cell_el.find(_q("v"))
            if v_el is None:
                v_el = ET.SubElement(cell_el, _q("v"))
            v_el.text = f"{float(value):.10g}"
            total += 1
            changed = True
        if changed:
            files[path] = ET.tostring(
                root,
                encoding="utf-8",
                xml_declaration=True,
            )

    if total == 0:
        return 0

    tmp = xlsx_path.with_suffix(xlsx_path.suffix + ".fcache.tmp")
    with ZipFile(tmp, "w", compression=ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    tmp.replace(xlsx_path)
    return total
