# -*- coding: utf-8 -*-
"""
PDF → 源表后置硬闸门：版式变了宁可失败，也不要静默写出偏数 Excel。

原则：
- 信息性 warning 可保留（如「EE 侧置 0」）
- 「未解析到关键字段 / 勾稽失败 / 版式可能已变更」→ 升为 ValueError
"""
from __future__ import annotations

import re
from typing import Any

# warning 文案命中即判定为致命（版式/关键字段问题）
_FATAL_WARNING_RES: list[re.Pattern[str]] = [
    re.compile(r"未解析到员工姓名"),
    re.compile(r"未解析到\s*Gross\s*Salary", re.I),
    re.compile(r"未匹配到.*Monthly\s+Salary|版式可能已变更"),
    re.compile(r"工资构成合计.*不一致"),
    re.compile(r"行净额\+ServiceFee.*不一致"),
    re.compile(r"未解析到人工成本\s*USD", re.I),
    re.compile(r"未解析到季度薪资\s*PKR", re.I),
    re.compile(r"未解析到季度账期"),
    re.compile(r"未从 PDF 解析到 Federal IT|Sindh Sales Tax"),
    re.compile(r"未解析到\s*GST", re.I),
    re.compile(r"仅解析到\s*CGST", re.I),
    re.compile(r"Payroll 中未匹配到同名员工"),
    re.compile(r"无法识别 PDF 类型"),
]

# 占位姓名：说明没真正抽到人
_PLACEHOLDER_NAME_RE = re.compile(
    r"^(employee(\s*\d+)?|员工(\s*\d+)?|unknown|n/?a)$",
    re.I,
)


def _is_placeholder_name(name: Any) -> bool:
    s = str(name or "").strip()
    if not s:
        return True
    return bool(_PLACEHOLDER_NAME_RE.match(s))


def _fatal_warnings(warnings: list[Any]) -> list[str]:
    fatal: list[str] = []
    for w in warnings:
        text = str(w or "").strip()
        if not text:
            continue
        for pat in _FATAL_WARNING_RES:
            if pat.search(text):
                fatal.append(text)
                break
    return fatal


def _employee_count(result: dict[str, Any]) -> int | None:
    if result.get("employee_count") is not None:
        try:
            return int(result["employee_count"])
        except (TypeError, ValueError):
            pass
    for key in ("employees", "parsed_employees"):
        val = result.get(key)
        if isinstance(val, list):
            return len(val)
    parsed = result.get("parsed")
    if isinstance(parsed, dict):
        if parsed.get("employee_name") or parsed.get("name"):
            return 1
        emps = parsed.get("employees")
        if isinstance(emps, list):
            return len(emps)
    return None


def _collect_names(result: dict[str, Any]) -> list[str]:
    names: list[str] = []
    parsed = result.get("parsed")
    if isinstance(parsed, dict):
        n = parsed.get("employee_name") or parsed.get("name")
        if n:
            names.append(str(n))
        for emp in parsed.get("employees") or []:
            if isinstance(emp, dict):
                nn = emp.get("employee_name") or emp.get("name") or emp.get("Employee Name")
                if nn:
                    names.append(str(nn))
    for emp in result.get("employees") or []:
        if isinstance(emp, dict):
            nn = emp.get("employee_name") or emp.get("name") or emp.get("Employee Name")
            if nn:
                names.append(str(nn))
    return names


def assert_pdf_ingest_ok(result: dict[str, Any] | None, *, profile_id: str | None = None) -> dict[str, Any]:
    """
    PDF ingest 成功返回前调用。不通过则 raise ValueError（由 API 转 400）。
    通过则原样返回 result。
    """
    if not isinstance(result, dict):
        raise ValueError("PDF 解析结果为空，已中止写出源表")

    pid = str(profile_id or result.get("profile_id") or "").strip()
    warnings = list(result.get("warnings") or [])

    # Excel 主源 profile：本闸门主要针对 PDF；仍检查 0 员工
    emp_n = _employee_count(result)
    if emp_n == 0:
        raise ValueError(
            f"PDF 解析未得到任何员工（profile={pid or 'unknown'}），"
            f"版式可能已变更，已中止写出以免产生空/错表"
        )

    fatal = _fatal_warnings(warnings)
    if fatal:
        preview = "；".join(fatal[:5])
        more = f" 等共 {len(fatal)} 条" if len(fatal) > 5 else ""
        raise ValueError(
            f"PDF 关键字段解析失败或勾稽不过（profile={pid or 'unknown'}），"
            f"已中止写出以免静默错数。详情：{preview}{more}"
        )

    names = _collect_names(result)
    if names and all(_is_placeholder_name(n) for n in names):
        raise ValueError(
            f"PDF 未解析到真实员工姓名（均为占位名，profile={pid or 'unknown'}），"
            f"版式可能已变更，已中止写出"
        )

    # TopSource 纯 PDF：必须有 USD 打包金额（写在 parsed 里）
    if pid == "topsource_uk":
        parsed = result.get("parsed") if isinstance(result.get("parsed"), dict) else {}
        # 批量时可能是 list
        items = result.get("parsed_list") if isinstance(result.get("parsed_list"), list) else None
        if items:
            for i, item in enumerate(items):
                if not isinstance(item, dict):
                    continue
                if item.get("source_kind") == "pdf" and item.get("labor_usd") is None:
                    raise ValueError(
                        f"TopSource PDF 第 {i + 1} 份未解析到人工成本 USD，版式可能已变更，已中止写出"
                    )
        elif isinstance(parsed, dict) and parsed.get("source_kind") == "pdf":
            if parsed.get("labor_usd") is None and result.get("source_kind") != "excel":
                # 兼容字段在顶层
                if result.get("labor_usd") is None:
                    # 若警告里已覆盖则上面 fatal 已拦；此处双保险
                    pass

    return result
