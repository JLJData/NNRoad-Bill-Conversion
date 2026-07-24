# -*- coding: utf-8 -*-
"""
openpyxl 以 rich_text=True 保存后，富文本常变成 worksheet 内的 inlineStr。
Excel 能正常打开，但 LuckyExcel 预览解析 inlineStr 时往往只读到第一个 <t>
（例如标题只剩一个蓝色「N」）。

本模块在转换保存后，把含 <r> 的 inlineStr 迁回 sharedStrings（t=\"s\"），
与原母版格式一致，兼顾下载与预览。
"""
from __future__ import annotations

import re
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

CONTENT_TYPES = "[Content_Types].xml"
SHARED_STRINGS = "xl/sharedStrings.xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"

CT_OVERRIDE = (
    '<Override PartName="/xl/sharedStrings.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
)
RELS_SST = (
    '<Relationship Id="rIdSharedStrings" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
    'Target="sharedStrings.xml"/>'
)

# <c ... t="inlineStr" ...><is>...</is></c>  （允许属性顺序变化）
_INLINE_CELL_RE = re.compile(
    r'<c\b([^>]*?\bt="inlineStr"[^>]*?)>(\s*)<is>([\s\S]*?)</is>(\s*)</c>',
    re.MULTILINE,
)
_NUMERIC_CHAR_REF_RE = re.compile(r"&#(x[0-9a-fA-F]+|\d+);")


def _decode_numeric_char_refs(xml: str) -> str:
    """
    openpyxl 常把中文等写成 &#31561;&#32447;（等线）。
    Excel 能认，但 LuckySheet 的 FontFaceSet 会把实体字面量当字体名而报错。
    只解码数字/十六进制字符引用，保留 &amp; &lt; 等命名实体以免破坏 XML。
    """

    def _repl(m: re.Match[str]) -> str:
        num = m.group(1)
        code = int(num[1:], 16) if num.lower().startswith("x") else int(num)
        try:
            return chr(code)
        except ValueError:
            return m.group(0)

    return _NUMERIC_CHAR_REF_RE.sub(_repl, xml)


def _ensure_content_types(xml: str) -> str:
    if "sharedStrings.xml" in xml:
        return xml
    if CT_OVERRIDE in xml:
        return xml
    # 插在 </Types> 前
    return xml.replace("</Types>", f"  {CT_OVERRIDE}\n</Types>")


def _ensure_workbook_rels(xml: str) -> str:
    if "sharedStrings" in xml:
        return xml
    return xml.replace("</Relationships>", f"  {RELS_SST}\n</Relationships>")


def _load_existing_sis(sst_xml: str | None) -> list[str]:
    if not sst_xml:
        return []
    return re.findall(r"<si>[\s\S]*?</si>", sst_xml)


def _build_sst(sis: list[str]) -> str:
    count = len(sis)
    body = "".join(sis)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{count}" uniqueCount="{count}">{body}</sst>'
    )


def _replace_inline_cell(match: re.Match[str], sis: list[str]) -> str:
    attrs = match.group(1)
    is_inner = _decode_numeric_char_refs(match.group(3))
    # 仅迁移带富文本 run 的；纯 <t> 也可迁，提升 LuckyExcel 兼容
    si_xml = f"<si>{is_inner}</si>"
    idx = len(sis)
    sis.append(si_xml)

    # 去掉 t="inlineStr"，改为 t="s"，内容改为 <v>index</v>
    new_attrs = re.sub(r'\s*t="inlineStr"', "", attrs)
    if re.search(r'\bt="', new_attrs):
        new_attrs = re.sub(r'\bt="[^"]*"', 't="s"', new_attrs, count=1)
    else:
        new_attrs = f'{new_attrs} t="s"'
    return f'<c{new_attrs}><v>{idx}</v></c>'


def migrate_inlinestr_richtext_to_shared_strings(xlsx_path: Path) -> int:
    """
    将 xlsx 内 inlineStr（尤其是多 run 富文本）迁回 sharedStrings。
    返回迁移的单元格数量。
    """
    xlsx_path = Path(xlsx_path)
    migrated = 0

    with ZipFile(xlsx_path, "r") as zin:
        namelist = zin.namelist()
        files: dict[str, bytes] = {name: zin.read(name) for name in namelist}

    sst_xml = (
        files.get(SHARED_STRINGS, b"").decode("utf-8", "ignore")
        if SHARED_STRINGS in files
        else None
    )
    sis = _load_existing_sis(sst_xml)

    sheet_names = [
        n for n in namelist if n.startswith("xl/worksheets/") and n.endswith(".xml")
    ]

    for sheet_name in sheet_names:
        xml = files[sheet_name].decode("utf-8", "ignore")

        def _sub(m: re.Match[str], _sis=sis) -> str:
            nonlocal migrated
            migrated += 1
            return _replace_inline_cell(m, _sis)

        new_xml, n = _INLINE_CELL_RE.subn(_sub, xml)
        if n:
            files[sheet_name] = new_xml.encode("utf-8")

    if migrated == 0:
        return 0

    files[SHARED_STRINGS] = _build_sst(sis).encode("utf-8")

    if CONTENT_TYPES in files:
        ct = files[CONTENT_TYPES].decode("utf-8", "ignore")
        files[CONTENT_TYPES] = _ensure_content_types(ct).encode("utf-8")

    if WORKBOOK_RELS in files:
        rels = files[WORKBOOK_RELS].decode("utf-8", "ignore")
        files[WORKBOOK_RELS] = _ensure_workbook_rels(rels).encode("utf-8")

    tmp = xlsx_path.with_suffix(xlsx_path.suffix + ".richfix.tmp")
    with ZipFile(tmp, "w", compression=ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    tmp.replace(xlsx_path)
    return migrated
