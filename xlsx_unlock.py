# -*- coding: utf-8 -*-
"""打开可能加密的 xlsx：按 mapping / 环境变量 / 文件名服务月尝试密码。"""
from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Iterable

_YYYYMM_RE = re.compile(r"(20\d{2})(0[1-9]|1[0-2])")
_DURATION_MONTH_RE = re.compile(r"^(20\d{2})-?(0[1-9]|1[0-2])$")


def yyyymm_passwords_from_text(text: object) -> list[str]:
    found: list[str] = []
    seen: set[str] = set()
    for y, m in _YYYYMM_RE.findall(str(text or "")):
        pwd = f"nnroad{y}{m}"
        if pwd not in seen:
            seen.add(pwd)
            found.append(pwd)
    return found


def passwords_from_filename(path: Path) -> list[str]:
    return yyyymm_passwords_from_text(path.name)


def passwords_from_duration(duration: object) -> list[str]:
    text = str(duration or "").strip()
    if not text:
        return []
    m = _DURATION_MONTH_RE.match(text.replace(" ", ""))
    if m:
        return [f"nnroad{m.group(1)}{m.group(2)}"]
    return yyyymm_passwords_from_text(text)


def collect_unlock_passwords(
    path: Path,
    *,
    mapping: dict | None = None,
    extra: Iterable[str] | None = None,
) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()

    def add(raw: object) -> None:
        text = str(raw or "").strip()
        if text and text not in seen:
            seen.add(text)
            out.append(text)

    if isinstance(mapping, dict):
        add(mapping.get("sourcePassword"))
        src = mapping.get("sourceEmployeeSheet")
        if isinstance(src, dict):
            add(src.get("password"))
        for key in (
            "sourceOriginalName",
            "_sourceOriginalName",
            "sourceDiskName",
            "sourceFilename",
        ):
            for pwd in yyyymm_passwords_from_text(mapping.get(key)):
                add(pwd)
        for key in ("billDuration", "_billDuration", "duration"):
            for pwd in passwords_from_duration(mapping.get(key)):
                add(pwd)
    add(os.environ.get("CONVERT_XLSX_PASSWORD"))
    for pwd in passwords_from_filename(path):
        add(pwd)
    if extra:
        for pwd in extra:
            add(pwd)
    return out


def unlock_xlsx(path: Path, dest_dir: Path, passwords: Iterable[str] | None = None) -> Path:
    """
    未加密则返回原路径；加密则解密到 dest_dir 并返回新路径。
    解不开时抛 ValueError。
    """
    path = Path(path)
    dest_dir = Path(dest_dir)
    dest_dir.mkdir(parents=True, exist_ok=True)

    try:
        import msoffcrypto
    except ImportError:
        # 未加密文件不需要该库；加密时再报错
        msoffcrypto = None  # type: ignore

    with path.open("rb") as fh:
        if msoffcrypto is None:
            office = None
            encrypted = False
        else:
            try:
                office = msoffcrypto.OfficeFile(fh)
                encrypted = bool(office.is_encrypted())
            except Exception:
                encrypted = False
                office = None

    if not encrypted:
        return path

    if msoffcrypto is None:
        raise ValueError("源表已加密，请 pip install msoffcrypto-tool 后再转换")

    last_err: Exception | None = None
    tried = [p for p in (passwords or []) if str(p).strip()]
    if not tried:
        raise ValueError(
            f"源表已加密，但未提供密码（可在 mapping.sourcePassword 设置，"
            f"或按文件名服务月使用 nnroadYYYYMM）：{path.name}"
        )

    # 禁止写回输入路径：Windows 上边读边写会截断文件，
    # 表现为 unpack requires a buffer of 8 bytes（预览解密曾踩过）。
    out_path = dest_dir / path.name
    if out_path.resolve() == path.resolve():
        out_path = dest_dir / f"{path.stem}.unlocked{path.suffix}"
    for pwd in tried:
        try:
            with path.open("rb") as fh:
                office = msoffcrypto.OfficeFile(fh)
                office.load_key(password=str(pwd))
                with out_path.open("wb") as out_fh:
                    office.decrypt(out_fh)
            return out_path
        except Exception as exc:
            last_err = exc
            continue
    raise ValueError(f"源表解密失败（已试 {len(tried)} 个密码）: {path.name}: {last_err}")
