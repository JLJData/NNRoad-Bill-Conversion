# -*- coding: utf-8 -*-
"""openpyxl 保存时要靠 Pillow 才能把浮动图/Logo 写回去。缺包会静默丢图。"""
from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile


def require_pillow() -> None:
    try:
        import PIL.Image  # noqa: F401
    except ImportError as e:
        raise RuntimeError(
            "转换保存 Excel 会丢掉 Logo：未安装 Pillow。"
            "请在转换服务目录执行 pip install pillow 后重启。"
        ) from e
    try:
        import openpyxl.drawing.image as oxl_img
    except Exception:
        return
    if getattr(oxl_img, "PILImage", None) is None:
        raise RuntimeError(
            "openpyxl 加载时还没有 Pillow，图片仍会丢失。请安装 pillow 后重启转换服务。"
        )


def xlsx_media_names(path: Path | str) -> list[str]:
    p = Path(path)
    if not p.is_file():
        return []
    with ZipFile(p) as zf:
        return sorted(
            n for n in zf.namelist() if n.startswith("xl/media/") and not n.endswith("/")
        )


def assert_template_images_kept(template_path: Path | str, output_path: Path | str) -> None:
    """母版有 xl/media 则结果必须还在，避免再静默丢 Logo。"""
    src = xlsx_media_names(template_path)
    if not src:
        return
    dst = xlsx_media_names(output_path)
    if len(dst) >= len(src):
        return
    raise RuntimeError(
        f"转换结果丢失了母版图片/Logo（母版 {len(src)} 张，结果 {len(dst)} 张）。"
        "请确认转换服务已安装 Pillow 并已重启。"
    )
