# -*- coding: utf-8 -*-
"""调用 Node HyperFormula 快照（与 excel_com_snapshot 输出同形）。"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
from pathlib import Path
from typing import Any


BASE_DIR = Path(__file__).resolve().parent
SCRIPT = BASE_DIR / "hf_snapshot.mjs"


def _resolve_office_ui() -> Path:
    env = os.environ.get("HRONE_OFFICE_UI", "").strip()
    if env:
        return Path(env)
    candidates = [
        BASE_DIR.parent / "GIT 其他项目" / "hrone-office-ui",
        BASE_DIR.parent / "hrone-office-ui",
    ]
    for c in candidates:
        if (c / "node_modules" / "hyperformula").is_dir():
            return c
    return candidates[0]


def snapshot_workbook_hf(path: Path, sheet_filter: str | None, max_cells: int) -> dict[str, Any]:
    if not SCRIPT.is_file():
        return {
            "ok": False,
            "engine": "hyperformula",
            "sheetName": sheet_filter or "",
            "cells": [],
            "truncated": False,
            "message": f"缺少脚本: {SCRIPT}",
        }
    node = shutil.which("node")
    if not node:
        return {
            "ok": False,
            "engine": "hyperformula",
            "sheetName": sheet_filter or "",
            "cells": [],
            "truncated": False,
            "message": "未找到 node，请安装 Node.js",
        }
    office_ui = _resolve_office_ui()
    cmd = [node, str(SCRIPT), str(path), "--max", str(max(1, max_cells))]
    if sheet_filter is None:
        cmd.extend(["--sheet", ""])
    else:
        cmd.extend(["--sheet", sheet_filter])
    env = os.environ.copy()
    env["HRONE_OFFICE_UI"] = str(office_ui)
    env["NODE_PATH"] = str(office_ui / "node_modules")
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
            env=env,
            cwd=str(BASE_DIR),
        )
    except subprocess.TimeoutExpired:
        return {
            "ok": False,
            "engine": "hyperformula",
            "sheetName": sheet_filter or "",
            "cells": [],
            "truncated": False,
            "message": "HF 快照超时（>180s）",
        }
    except Exception as exc:
        return {
            "ok": False,
            "engine": "hyperformula",
            "sheetName": sheet_filter or "",
            "cells": [],
            "truncated": False,
            "message": f"启动 Node 失败: {exc}",
        }
    text = (proc.stdout or "").strip()
    if not text:
        err = (proc.stderr or "").strip()
        return {
            "ok": False,
            "engine": "hyperformula",
            "sheetName": sheet_filter or "",
            "cells": [],
            "truncated": False,
            "message": err or f"HF 无输出 exit={proc.returncode}",
        }
    # LuckyExcel 会往 stdout 打噪音行（如 [] []），只取以 { 开头的 JSON 行
    json_lines = [ln.strip() for ln in text.splitlines() if ln.strip().startswith("{")]
    payload = json_lines[-1] if json_lines else text[text.rfind("{") :]
    try:
        return json.loads(payload)
    except Exception:
        return {
            "ok": False,
            "engine": "hyperformula",
            "sheetName": sheet_filter or "",
            "cells": [],
            "truncated": False,
            "message": f"HF 输出非 JSON: {text[:200]}",
        }
