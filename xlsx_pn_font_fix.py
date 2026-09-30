# -*- coding: utf-8 -*-
"""
Windows 中文环境 openpyxl 保存后，PN Bill-to 区字体常被替换成「等线」，导致与母版 Calibri 折行不一致。

写盘后把 PN!B8/B9/B10/B11/F9/F10/F11 所用字体名统一为 Calibri（保留原字号）。
"""
from __future__ import annotations

import re
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", NS_MAIN)

_PN_BILLTO_REFS = frozenset({"B8", "B9", "B10", "B11", "F9", "F10", "F11"})
_LATIN_SUBSTITUTE_FONTS = frozenset(
    {
        "等线",
        "等线 Light",
        "DengXian",
        "DengXian Light",
        "宋体",
        "SimSun",
        "微软雅黑",
        "Microsoft YaHei",
    }
)
_CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")


def _q(tag: str) -> str:
    return f"{{{NS_MAIN}}}{tag}"


def _find_pn_sheet_path(files: dict[str, bytes]) -> str | None:
    wb = ET.fromstring(files["xl/workbook.xml"])
    rels = ET.fromstring(files["xl/_rels/workbook.xml.rels"])
    rel_map = {
        r.get("Id"): r.get("Target", "").lstrip("/")
        for r in rels
        if r.tag.endswith("Relationship")
    }
    for sh in wb.findall(_q("sheets") + "/" + _q("sheet")):
        if sh.get("name") != "PN":
            continue
        rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rel_map.get(rid, "")
        if not target:
            return None
        return target if target.startswith("xl/") else f"xl/{target}"
    return None


def _collect_pn_billto_font_ids(files: dict[str, bytes]) -> set[int]:
    pn_path = _find_pn_sheet_path(files)
    if not pn_path or pn_path not in files:
        return set()
    root = ET.fromstring(files[pn_path])
    styles = ET.fromstring(files["xl/styles.xml"])
    xfs = styles.findall(_q("cellXfs") + "/" + _q("xf"))
    font_ids: set[int] = set()
    for cell in root.findall(".//" + _q("c")):
        ref = cell.get("r") or ""
        if ref not in _PN_BILLTO_REFS:
            continue
        style_idx = int(cell.get("s", "0") or "0")
        if style_idx < 0 or style_idx >= len(xfs):
            continue
        font_ids.add(int(xfs[style_idx].get("fontId", "0") or "0"))
    return font_ids


def _should_replace_font(name: str | None) -> bool:
    text = (name or "").strip()
    if not text:
        return False
    if text in _LATIN_SUBSTITUTE_FONTS:
        return True
    return "等线" in text or text.lower().startswith("dengxian")


def normalize_pn_billto_fonts(xlsx_path: Path | str) -> int:
    """
    将 PN Bill-to 区单元格引用的 CJK 默认字体改为 Calibri。

    Returns:
        修改的 font 条目数。
    """
    path = Path(xlsx_path)
    with ZipFile(path, "r") as zin:
        files = {info.filename: zin.read(info.filename) for info in zin.infolist()}

    if "xl/styles.xml" not in files:
        return 0

    font_ids = _collect_pn_billto_font_ids(files)
    if not font_ids:
        return 0

    styles_root = ET.fromstring(files["xl/styles.xml"])
    fonts_el = styles_root.find(_q("fonts"))
    if fonts_el is None:
        return 0

    changed = 0
    for i, font_el in enumerate(fonts_el.findall(_q("font"))):
        if i not in font_ids:
            continue
        name_el = font_el.find(_q("name"))
        if name_el is None:
            continue
        current = name_el.get("val") or ""
        if not _should_replace_font(current):
            continue
        name_el.set("val", "Calibri")
        changed += 1

    if not changed:
        return 0

    files["xl/styles.xml"] = ET.tostring(styles_root, encoding="utf-8", xml_declaration=True)
    with ZipFile(path, "w", ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    return changed
