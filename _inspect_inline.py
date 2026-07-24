# -*- coding: utf-8 -*-
from pathlib import Path
import shutil
import tempfile
import zipfile

from openpyxl import load_workbook

src = Path(r"d:/CodeUp-工资哥All项目/转换/templates/taiwan/template.xlsx")
out = Path(tempfile.gettempdir()) / "conv_rich2.xlsx"
shutil.copy2(src, out)
wb = load_workbook(out, rich_text=True)
wb["PN"]["B31"].value = 1.0
wb.save(out)
wb.close()

with zipfile.ZipFile(out) as z:
    xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8", "ignore")
    # PN is which sheet? need map
    # list workbook sheets
    wbxml = z.read("xl/workbook.xml").decode("utf-8", "ignore")
Path(tempfile.gettempdir()).joinpath("wb.xml").write_text(wbxml, encoding="utf-8")

# find sheet for PN
import re
sheets = re.findall(r'<sheet[^>]+>', wbxml)
Path(tempfile.gettempdir()).joinpath("sheets.txt").write_text("\n".join(sheets), encoding="utf-8")

with zipfile.ZipFile(out) as z:
    for name in z.namelist():
        if not name.startswith("xl/worksheets/"):
            continue
        xml = z.read(name).decode("utf-8", "ignore")
        if "8DC63F" in xml or "Employer of Record" in xml:
            i = xml.find("8DC63F")
            if i < 0:
                i = xml.find("Employer")
            Path(tempfile.gettempdir()).joinpath("cell_snip.xml").write_text(
                xml[max(0, i - 600) : i + 800], encoding="utf-8"
            )
            Path(tempfile.gettempdir()).joinpath("which_sheet.txt").write_text(name, encoding="utf-8")
            break
print("ok")
