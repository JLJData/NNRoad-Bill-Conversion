# -*- coding: utf-8 -*-
"""
openpyxl 保存后常见问题：
1) cellXfs 丢掉 applyFill=\"1\" → Excel 按 cellStyleXfs 的空填充渲染，浅蓝底消失
2) 主题色 + tint 在部分预览引擎上不稳定

本模块把 styles.xml 里的 theme 填充落地为 RGB，并为有填充的 xf 补上 applyFill=1。
"""
from __future__ import annotations

import colorsys
import re
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

ET.register_namespace("", NS_MAIN)


def _q(tag: str, ns: str = NS_MAIN) -> str:
    return f"{{{ns}}}{tag}"


def _parse_hex_rgb(s: str) -> tuple[int, int, int] | None:
    h = (s or "").strip().lstrip("#")
    if len(h) == 8:
        h = h[2:]
    if len(h) != 6:
        return None
    try:
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except ValueError:
        return None


def _to_hex_rgb(r: int, g: int, b: int) -> str:
    return f"FF{r:02X}{g:02X}{b:02X}"


def _apply_tint(rgb: tuple[int, int, int], tint: float) -> tuple[int, int, int]:
    """与 LuckyExcel LightenDarkenColor 一致：在 HSL 明度上调 tint。"""
    r, g, b = [x / 255.0 for x in rgb]
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if tint > 0:
        l = l * (1.0 - tint) + tint
    elif tint < 0:
        l = l * (1.0 + tint)
    nr, ng, nb = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)
    return int(round(nr * 255)), int(round(ng * 255)), int(round(nb * 255))


def _load_theme_rgbs(files: dict[str, bytes]) -> list[tuple[int, int, int] | None]:
    """
    OOXML theme 顺序: dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink
    Excel theme 属性对 0/1、2/3 有交换，与 LuckyExcel getColor 一致。
    """
    theme_path = next((k for k in files if k.startswith("xl/theme/") and k.endswith(".xml")), None)
    if not theme_path:
        return [None] * 12
    root = ET.fromstring(files[theme_path])
    scheme = root.find(f".//{{{NS_A}}}clrScheme")
    if scheme is None:
        return [None] * 12
    colors: list[tuple[int, int, int] | None] = []
    for child in list(scheme):
        sys_clr = child.find(f"{{{NS_A}}}sysClr")
        srgb = child.find(f"{{{NS_A}}}srgbClr")
        hex_v = None
        if sys_clr is not None:
            hex_v = sys_clr.get("lastClr") or sys_clr.get("val")
        elif srgb is not None:
            hex_v = srgb.get("val")
        colors.append(_parse_hex_rgb(hex_v or ""))
    while len(colors) < 12:
        colors.append(None)
    return colors[:12]


def _theme_index_to_scheme(theme_num: int) -> int:
    if theme_num == 0:
        return 1
    if theme_num == 1:
        return 0
    if theme_num == 2:
        return 3
    if theme_num == 3:
        return 2
    return theme_num


def _resolve_fg_color(
    fg: ET.Element,
    theme_colors: list[tuple[int, int, int] | None],
) -> str | None:
    rgb = fg.get("rgb")
    if rgb and re.fullmatch(r"[0-9A-Fa-f]{6,8}", rgb):
        if len(rgb) == 6:
            return "FF" + rgb.upper()
        return rgb.upper()

    theme = fg.get("theme")
    if theme is None:
        return None
    try:
        theme_num = int(theme)
    except ValueError:
        return None
    idx = _theme_index_to_scheme(theme_num)
    base = theme_colors[idx] if 0 <= idx < len(theme_colors) else None
    if base is None:
        return None
    tint_s = fg.get("tint")
    tint = float(tint_s) if tint_s not in (None, "") else 0.0
    r, g, b = _apply_tint(base, tint) if tint else base
    return _to_hex_rgb(r, g, b)


def materialize_theme_fills(xlsx_path: Path) -> dict[str, int]:
    path = Path(xlsx_path)
    with ZipFile(path, "r") as zf:
        files = {name: zf.read(name) for name in zf.namelist()}

    if "xl/styles.xml" not in files:
        return {"fills": 0, "xfs": 0}

    theme_colors = _load_theme_rgbs(files)
    styles = ET.fromstring(files["xl/styles.xml"])
    fills_el = styles.find(_q("fills"))
    xfs_el = styles.find(_q("cellXfs"))
    fill_n = 0
    xf_n = 0

    if fills_el is not None:
        for fill in fills_el.findall(_q("fill")):
            pf = fill.find(_q("patternFill"))
            if pf is None:
                continue
            # openpyxl 偶发写出空 <patternFill/>，补回 none
            if pf.get("patternType") is None and pf.find(_q("fgColor")) is None:
                pf.set("patternType", "none")
                fill_n += 1
                continue
            if pf.get("patternType") != "solid":
                continue
            fg = pf.find(_q("fgColor"))
            if fg is None or fg.get("theme") is None:
                continue
            rgb = _resolve_fg_color(fg, theme_colors)
            if not rgb:
                continue
            fg.attrib.clear()
            fg.set("rgb", rgb)
            fill_n += 1

    if xfs_el is not None:
        for xf in xfs_el.findall(_q("xf")):
            fill_id = xf.get("fillId")
            if fill_id in (None, "0"):
                continue
            if xf.get("applyFill") == "1":
                continue
            xf.set("applyFill", "1")
            xf_n += 1

    files["xl/styles.xml"] = ET.tostring(styles, encoding="utf-8", xml_declaration=True)

    tmp = path.with_suffix(path.suffix + ".tmp")
    with ZipFile(tmp, "w", compression=ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    tmp.replace(path)
    return {"fills": fill_n, "xfs": xf_n}
