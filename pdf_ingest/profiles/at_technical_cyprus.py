# -*- coding: utf-8 -*-
"""
A&T Professional Technical Management Services（Cyprus）
Invoice PDF + Payroll Calculation PDF → Cyprus-L（profile: at_technical_cyprus）

两份 PDF 一起上传：
  - Invoice：姓名（First Last）、账期、Gross、ER Contributions、Public Liability、Administration Fee
  - Payroll Calculation：同人明细（Last First）、EE Social Ins / Tax / N.H.S.、ER Contributions 校验

Recurring Fee 不写死：走 mapping.cyprusRecurringFee（引擎写 Cyprus!I）。
"""
from __future__ import annotations

import argparse
import calendar
import re
import shutil
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from bill_convert.person import compact_person_name, score_person_name_match
from convert_mapping import get_builtin_column_rename
from pdf_ingest.label_pairs import (
    apply_label_amounts,
    build_label_rename_map,
    extract_label_amounts,
    norm_label,
    parse_money,
    AMOUNT_token,
)
from pdf_ingest.text_extract import extract_pdf_text
from pn_meta import PnMeta
from region_templates import get_region_template

CYPRUS_L_SHEET = "Cyprus-L"

# Invoice 默认识别的供应商标签（可被 builtinColumnRename / columnRename 扩展）
_AT_INVOICE_DEFAULT_LABELS = [
    "Gross Salary",
    "Medical Insurance Cover",
    "Employer's Contributions",
    "Employer's & Public Liability",
    "Administration Fee",
]

# 供应商标签 → Cyprus-L / 内部字段（builtin；Office 列名对照可覆盖）
_AT_INVOICE_BUILTIN_RENAME = {
    "Gross Salary": "Base salary",
    "Employer's Contributions": "Employer's contributions",
    "Employer's & Public Liability": "Employer's & Public Liability",
    "Administration Fee": "_admin_fee",
    "Medical Insurance Cover": "Medical Insurance",
}

_MONTHS = {
    "january": 1,
    "february": 2,
    "march": 3,
    "april": 4,
    "may": 5,
    "june": 6,
    "july": 7,
    "august": 8,
    "september": 9,
    "october": 10,
    "november": 11,
    "december": 12,
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "sept": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}


def _as_float(value: Any) -> float | None:
    return parse_money(value)


def _invoice_label_catalog(convert_mapping: dict[str, Any] | None) -> tuple[list[str], dict[str, str]]:
    """返回 (待抽取标签列表, vendor_label_key→目标字段)。"""
    mapping = convert_mapping if isinstance(convert_mapping, dict) else {}
    rename_raw = mapping.get("columnRename") if isinstance(mapping.get("columnRename"), dict) else {}
    # profile 内置 + Office 对照；get_builtin 可能为空时用本文件默认
    builtin = get_builtin_column_rename("at_technical_cyprus") or dict(_AT_INVOICE_BUILTIN_RENAME)
    rename_map = build_label_rename_map(builtin=builtin, column_rename=rename_raw)
    labels = list(_AT_INVOICE_DEFAULT_LABELS)
    for src in list(builtin.keys()) + list(rename_raw.keys()):
        s = norm_label(str(src))
        if s and s not in labels:
            labels.append(s)
    return labels, rename_map


def _period_bounds(year: int, month: int) -> tuple[date, date]:
    start = date(year, month, 1)
    end = date(year, month, calendar.monthrange(year, month)[1])
    return start, end


def classify_at_pdf(text: str, path: Path | None = None) -> str:
    """返回 'invoice' | 'payroll' | 'unknown'。"""
    low = (text or "").lower()
    name = (path.name if path else "").lower()
    if "payroll calculation" in low or "payroll type:" in low or "empl.id:" in low:
        return "payroll"
    if "invoice" in low or "total due" in low or "services fee calculation" in low:
        return "invoice"
    if "invoice" in name:
        return "invoice"
    if "payroll" in name or "journal" in name:
        return "payroll"
    return "unknown"


def _fields_from_invoice_block(
    body: str,
    labels: list[str],
    rename_map: dict[str, str],
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """
    块内按标签抽金额再映射。Medical：ER 之前→Medical Insurance；Admin 之后→Other。
    返回 (fields, hits)。
    """
    hits = extract_label_amounts(body, labels)
    er_start = None
    admin_end = None
    for h in hits:
        lk = norm_label(str(h["label"])).lower()
        if "employer's contributions" in lk or lk == "employers contributions":
            er_start = int(h["start"])
        if "administration fee" in lk:
            admin_end = int(h["end"])

    # 先按 rename 汇总，再对 Medical 按位置拆分
    medical_now = 0.0
    medical_back = 0.0
    filtered: list[dict[str, Any]] = []
    for h in hits:
        lk = norm_label(str(h["label"])).lower()
        if "medical insurance" in lk:
            val = float(h["value"])
            if er_start is not None and int(h["start"]) < er_start:
                medical_now += val
            elif admin_end is not None and int(h["start"]) >= admin_end:
                medical_back += val
            else:
                medical_now += val
            continue
        filtered.append(h)

    fields = apply_label_amounts(filtered, rename_map)
    if medical_now:
        # Medical 位置拆分优先于 columnRename（当期 / 补收语义固定）
        fields["Medical Insurance"] = medical_now
    if medical_back:
        fields["Other "] = medical_back
        fields["Other"] = medical_back
    return fields, hits


def _invoice_source_label_present(hits: list[dict[str, Any]], *needles: str) -> bool:
    for h in hits:
        lk = norm_label(str(h.get("label") or "")).lower()
        for n in needles:
            if n in lk:
                return True
    return False


def parse_at_invoice_pdf(
    pdf_path: Path,
    *,
    convert_mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    path = Path(pdf_path).resolve()
    text = extract_pdf_text(path)
    warnings: list[str] = []
    labels, rename_map = _invoice_label_catalog(convert_mapping)

    inv_m = re.search(r"Invoice\s+(AT\d+)", text, flags=re.I)
    invoice_no = inv_m.group(1).strip() if inv_m else None

    date_m = re.search(r"DATE\s+(\d{1,2}/\d{1,2}/\d{4})", text, flags=re.I)
    invoice_date = None
    if date_m:
        try:
            invoice_date = datetime.strptime(date_m.group(1), "%d/%m/%Y").date().isoformat()
        except ValueError:
            warnings.append(f"无法解析发票日期: {date_m.group(1)}")

    period_label = None
    year = month = None
    pm = re.search(
        r"Monthly cost:\s*(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})",
        text,
        flags=re.I,
    )
    if pm:
        month = _MONTHS.get(pm.group(1).lower())
        year = int(pm.group(2))
        period_label = f"{pm.group(1)}-{pm.group(2)}"

    employees: list[dict[str, Any]] = []
    skipped: list[str] = []
    month_alt = (
        r"January|February|March|April|May|June|July|August|September|"
        r"October|November|December"
    )
    chunks = re.split(rf"(?=Monthly cost:\s*(?:{month_alt})\s+\d{{4}})", text, flags=re.I)
    head_re = re.compile(
        rf"Monthly cost:\s*(?P<month>{month_alt})\s+(?P<year>\d{{4}})\s*-\s*(?P<name>[^\n]+)",
        flags=re.I,
    )
    heads_seen = 0
    for chunk in chunks:
        head = head_re.search(chunk)
        if not head:
            continue
        heads_seen += 1
        body = chunk
        stop = re.search(r"\bSubtotal:|\bTOTAL DUE\b", chunk, flags=re.I)
        if stop:
            body = chunk[: stop.start()]

        name = re.sub(r"\s+", " ", head.group("name")).strip()
        name = re.sub(r"\s+Gross Salary.*$", "", name, flags=re.I).strip()
        y = int(head.group("year"))
        mo = _MONTHS.get(head.group("month").lower())
        if year is None:
            year, month = y, mo
            period_label = f"{head.group('month')}-{y}"

        fields, hits = _fields_from_invoice_block(body, labels, rename_map)
        # 结构完整性看供应商标签是否抽出，不因 columnRename 指错 Cyprus-L 列而整人丢弃
        has_gross = _invoice_source_label_present(hits, "gross salary")
        has_er = _invoice_source_label_present(
            hits, "employer's contributions", "employers contributions"
        )
        has_liab = _invoice_source_label_present(hits, "public liability")
        if not (has_gross and has_er and has_liab):
            miss_src = []
            if not has_gross:
                miss_src.append("Gross Salary")
            if not has_er:
                miss_src.append("Employer's Contributions")
            if not has_liab:
                miss_src.append("Employer's & Public Liability")
            msg = f"Invoice「{name}」缺供应商标签 {', '.join(miss_src)}，已跳过"
            warnings.append(msg)
            skipped.append(msg)
            continue

        if fields.get("Base salary") is None and fields.get("Gross Salary") is not None:
            fields["Base salary"] = fields.get("Gross Salary")
        if fields.get("Base salary") is None:
            warnings.append(
                f"Invoice「{name}」列名对照后无 Base salary"
                f"（请确认 Gross Salary → Base salary）；金额仍保留在对照目标列"
            )

        row: dict[str, Any] = {
            "Employee Name": name,
            "Name of Employee": name,
            "_source": "invoice",
            "_label_hits": list(fields.keys()),
        }
        row.update(fields)
        employees.append(row)

    if not employees:
        detail = "；".join(skipped[:5]) if skipped else "未匹配到 Monthly cost 员工行"
        if heads_seen and skipped:
            detail = (
                f"识别到 {heads_seen} 个员工块但均被跳过（多半是 PDF 缺标签，"
                f"而非列名对照问题）：{detail}"
            )
        raise ValueError(f"A&T Invoice 未解析到员工块: {path.name} — {detail}")

    start = end = None
    if year and month:
        start, end = _period_bounds(year, month)
        for e in employees:
            e["_period_from"] = start
            e["_period_to"] = end
            e["_period_label"] = period_label
            e["From"] = start
            e["To"] = end

    return {
        "kind": "invoice",
        "employees": employees,
        "invoice_no": invoice_no,
        "invoice_date": invoice_date,
        "period_label": period_label,
        "period_from": start,
        "period_to": end,
        "warnings": warnings,
        "source_file": path.name,
        "text": text,
        "labels": labels,
        "label_rename": {k: rename_map[k] for k in rename_map},
    }


def inspect_at_pdf_labels(
    pdf_path: Path,
    *,
    convert_mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """映射页样例：从 Invoice/Payroll PDF 抽出供应商标签，供列名对照。"""
    path = Path(pdf_path).resolve()
    text = extract_pdf_text(path)
    kind = classify_at_pdf(text, path)
    labels_cfg, rename_map = _invoice_label_catalog(convert_mapping)
    if kind == "payroll":
        # Payroll 侧暂列常见扣款/工资标签，便于对照；正式合并仍走专用解析
        labels_cfg = list(
            dict.fromkeys(
                labels_cfg
                + [
                    "Basic Salary",
                    "Social Ins",
                    "Tax-1",
                    "N.H.S.-SI",
                ]
            )
        )
    hits = extract_label_amounts(text, labels_cfg)
    # 同标签可出现多次（如 Medical 当期 + 补收）；对照表仍按标签名去重一行
    occ_count: dict[str, int] = {}
    occ_values: dict[str, list[float]] = {}
    for h in hits:
        lab = norm_label(str(h["label"]))
        key = lab.lower()
        occ_count[key] = occ_count.get(key, 0) + 1
        try:
            occ_values.setdefault(key, []).append(float(h["value"]))
        except (TypeError, ValueError):
            pass

    headers: list[dict[str, Any]] = []
    seen: set[str] = set()
    for h in hits:
        lab = norm_label(str(h["label"]))
        key = lab.lower()
        if key in seen:
            continue
        seen.add(key)
        target = rename_map.get(key) or ""
        n = int(occ_count.get(key, 1))
        vals = occ_values.get(key) or []
        label_show = lab if n <= 1 else f"{lab} (×{n})"
        row: dict[str, Any] = {"key": lab, "label": label_show, "mappedTo": target, "hitCount": n}
        if vals:
            row["sampleValues"] = vals
        headers.append(row)

    employees: list[dict[str, str]] = []
    medical_note = ""
    if kind == "invoice":
        try:
            parsed = parse_at_invoice_pdf(path, convert_mapping=convert_mapping)
            for e in parsed.get("employees") or []:
                en = str(e.get("Employee Name") or "").strip()
                if en:
                    employees.append({"cnName": "", "enName": en})
                    med = e.get("Medical Insurance")
                    other = e.get("Other") if e.get("Other") is not None else e.get("Other ")
                    if med is not None or other is not None:
                        bits = [en]
                        if med is not None:
                            bits.append(f"Medical Insurance={med}")
                        if other is not None:
                            bits.append(f"Other={other}")
                        if not medical_note:
                            medical_note = "；同标签多次：ER 前→Medical Insurance，Admin 后→Other（例：" + "，".join(bits) + "）"
        except Exception as exc:
            return {
                "ok": False,
                "message": str(exc),
                "sourceKind": "pdf_label",
                "pdfKind": kind,
                "headers": headers,
            }

    hit_total = len(hits)
    uniq_total = len(headers)
    hint = (
        f"已从供应商 PDF 抽取标签命中 {hit_total} 次（去重后 {uniq_total} 项）；"
        f"请在「列名对照」中配置 供应商标签 → Cyprus-L 列"
        f"{medical_note}"
    )
    return {
        "ok": True if headers or employees else False,
        "message": None if (headers or employees) else "未识别到 PDF 标签",
        "sheetName": path.name,
        "headerRow": 0,
        "layout": "pdf_label_amount",
        "sourceKind": "pdf_label",
        "pdfKind": kind,
        "headers": headers,
        "hitCount": hit_total,
        "uniqueLabelCount": uniq_total,
        "sampleEmployees": employees,
        "employees": employees,
        "hint": hint,
    }


def parse_at_payroll_pdf(pdf_path: Path) -> dict[str, Any]:
    path = Path(pdf_path).resolve()
    text = extract_pdf_text(path)
    warnings: list[str] = []

    year = month = None
    # Post Month: 6/2026
    pm = re.search(r"(\d{1,2})\s*/\s*(\d{4})\s*Post Month", text, flags=re.I)
    if not pm:
        pm = re.search(r"Post Month:\s*(\d{1,2})\s*/\s*(\d{4})", text, flags=re.I)
    if not pm:
        # text order sometimes "6/2026Post Month:"
        pm = re.search(r"(\d{1,2})/(\d{4})\s*Post Month", text, flags=re.I)
    if pm:
        month, year = int(pm.group(1)), int(pm.group(2))

    employees: list[dict[str, Any]] = []
    # Split by employee blocks
    parts = re.split(r"Empl\.ID:\s*", text, flags=re.I)
    for part in parts[1:]:
        head = re.match(
            r"(?P<eid>\d+)\s*-\s*(?P<name>[^\n]+?)(?:\s{2,}|\s+Empl\.Date:)",
            part,
            flags=re.I,
        )
        if not head:
            warnings.append("Payroll 某员工块无法解析姓名")
            continue
        raw_name = re.sub(r"\s+", " ", head.group("name")).strip()
        # Last First → keep for matching; westernize later when merging with invoice
        tokens = raw_name.split()
        if len(tokens) == 2:
            display = f"{tokens[1]} {tokens[0]}"
        else:
            display = raw_name

        basic_m = re.search(r"Basic Salary:\s*(" + AMOUNT_token + r")", part, flags=re.I)
        base = _as_float(basic_m.group(1)) if basic_m else None

        # Deductions：Social Ins / Tax-1 / N.H.S.-SI（Notice 行上的 EE NHS，勿取 Contributions 侧）
        si_m = re.search(
            r"(?:^|\n)[^\n]*?\bBasic\b[^\n]*?\bSocial Ins\s+(" + AMOUNT_token + r")",
            part,
            flags=re.I,
        )
        if not si_m:
            si_m = re.search(r"Social Ins\s+(" + AMOUNT_token + r")", part, flags=re.I)
        tax_m = re.search(r"Tax-1\s+(" + AMOUNT_token + r")", part, flags=re.I)
        nhs_m = re.search(
            r"Notice\s+[\d.,]+\s+[\d.,]+\s+N\.H\.S\.-SI\s+(" + AMOUNT_token + r")",
            part,
            flags=re.I,
        )
        if not nhs_m:
            # 回退：取较小的那个 NHS（EE 通常小于 ER）
            nhs_vals = [
                _as_float(x)
                for x in re.findall(r"N\.H\.S\.-SI\s+(" + AMOUNT_token + r")", part, flags=re.I)
            ]
            nhs_vals = [x for x in nhs_vals if x is not None]
            nhs = min(nhs_vals) if nhs_vals else None
        else:
            nhs = _as_float(nhs_m.group(1))

        # ER contributions：Ear.+Con. 块末行第三个数
        er = None
        er_m = re.search(
            r"Ear\.\+Con\.\s*[\d.,]+\s+("
            + AMOUNT_token
            + r")\s+("
            + AMOUNT_token
            + r")\s+("
            + AMOUNT_token
            + r")",
            part,
            flags=re.I | re.S,
        )
        if er_m:
            er = _as_float(er_m.group(3))
        else:
            triples = re.findall(
                r"(" + AMOUNT_token + r")\s+(" + AMOUNT_token + r")\s+(" + AMOUNT_token + r")",
                part,
            )
            if triples:
                er = _as_float(triples[-1][2])

        ee_si = _as_float(si_m.group(1)) if si_m else None
        ee_tax = _as_float(tax_m.group(1)) if tax_m else None

        employees.append(
            {
                "Employee Name": display,
                "Name of Employee": display,
                "_payroll_name": raw_name,
                "_empl_id": head.group("eid"),
                "Base salary": base,
                "Employer's contributions": er,
                "Employee's Social Insurance": -abs(ee_si) if ee_si is not None else None,
                "Employee's tax": -abs(ee_tax) if ee_tax is not None else None,
                "Employee - N.H.S.-SI": -abs(nhs) if nhs is not None else None,
                "_source": "payroll",
            }
        )

    if not employees:
        raise ValueError(f"A&T Payroll 未解析到员工: {path.name}")

    start = end = None
    period_label = None
    if year and month:
        start, end = _period_bounds(year, month)
        period_label = f"{year}-{month:02d}"
        for e in employees:
            e["_period_from"] = start
            e["_period_to"] = end
            e["_period_label"] = period_label
            e["From"] = start
            e["To"] = end

    return {
        "kind": "payroll",
        "employees": employees,
        "period_label": period_label,
        "period_from": start,
        "period_to": end,
        "warnings": warnings,
        "source_file": path.name,
        "text": text,
    }


def _match_emp(name: str, pool: list[dict[str, Any]]) -> dict[str, Any] | None:
    if not name or not pool:
        return None
    best = None
    best_score = 0
    compact = compact_person_name(name)
    for e in pool:
        candidates = [
            e.get("Employee Name"),
            e.get("Name of Employee"),
            e.get("_payroll_name"),
        ]
        for c in candidates:
            if not c:
                continue
            if compact_person_name(c) == compact:
                return e
            score = score_person_name_match(name, str(c))
            if score > best_score:
                best_score = score
                best = e
    return best if best_score >= 70 else None


def recover_cyprus_l_field_aliases(
    employees: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[str]]:
    """
    不改写列名对照结果，只提示。
    注意：不要把「…Liabilit」键复制成完整 Liability 键——两者指向同一物理列，
    写出若再合计会把 Contributions 金额算两次（685→1370）。
    """
    typo = "Employer's & Public Liabilit"
    full = "Employer's & Public Liability"
    warnings: list[str] = []
    out: list[dict[str, Any]] = []
    for emp in employees:
        row = dict(emp)
        tv, fv = row.get(typo), row.get(full)
        er = row.get("Employer's contributions")
        name = str(row.get("Employee Name") or row.get("Name of Employee") or "")
        if tv is not None and fv is not None:
            try:
                if abs(float(tv) - float(fv)) > 0.05:
                    warnings.append(
                        f"{name}：完整名与截断名 Liability 键均有值且不同"
                        f"（{fv} / {tv}），将合计写入 Public Liabilit 列"
                    )
            except (TypeError, ValueError):
                pass
        elif tv is not None and er is None and fv is None:
            warnings.append(
                f"{name}：有金额在截断表头「{typo}」，将写入 Public Liabilit 列"
            )
        out.append(row)
    return out, warnings


def merge_invoice_and_payroll(
    invoice: dict[str, Any] | None,
    payroll: dict[str, Any] | None,
) -> tuple[list[dict[str, Any]], list[str]]:
    warnings: list[str] = []
    if not invoice and not payroll:
        raise ValueError("未提供可用的 Invoice / Payroll PDF 解析结果")

    if invoice and not payroll:
        warnings.append("仅有 Invoice：缺少 EE Social Ins / Tax / N.H.S.，对应列置空")
        return list(invoice["employees"]), warnings
    if payroll and not invoice:
        warnings.append("仅有 Payroll：缺少 Public Liability，对应列置 0")
        out = []
        for e in payroll["employees"]:
            row = dict(e)
            row.setdefault("Employer's & Public Liability", 0.0)
            out.append(row)
        return out, warnings

    assert invoice is not None and payroll is not None
    inv_emps = list(invoice["employees"])
    pay_emps = list(payroll["employees"])
    used: set[int] = set()
    merged: list[dict[str, Any]] = []

    for inv in inv_emps:
        name = str(inv.get("Employee Name") or "")
        hit = _match_emp(name, pay_emps)
        row = dict(inv)
        if hit:
            used.add(id(hit))
            for k in (
                "Employee's Social Insurance",
                "Employee's tax",
                "Employee - N.H.S.-SI",
                "_empl_id",
                "_payroll_name",
            ):
                if hit.get(k) is not None:
                    row[k] = hit[k]
            # ER contrib：两边都有则优先 invoice（与税票一致），差异告警
            inv_er = inv.get("Employer's contributions")
            pay_er = hit.get("Employer's contributions")
            if inv_er is not None and pay_er is not None and abs(float(inv_er) - float(pay_er)) > 0.05:
                med = inv.get("Medical Insurance")
                explained = (
                    med is not None
                    and abs(float(inv_er) + float(med) - float(pay_er)) <= 0.05
                )
                if not explained:
                    warnings.append(
                        f"{name}：Invoice ER Contributions {inv_er} ≠ Payroll {pay_er}，已用 Invoice"
                    )
        else:
            warnings.append(f"{name}：Payroll 中未匹配到同名员工，EE 扣款列为空")
            raise ValueError(
                f"A&T Cyprus：员工「{name}」在 Payroll 中未匹配到同名，版式或姓名不一致，已中止写出以免扣款列空"
            )
        # period：优先 invoice
        if invoice.get("period_from"):
            row["_period_from"] = invoice["period_from"]
            row["_period_to"] = invoice["period_to"]
            row["From"] = invoice["period_from"]
            row["To"] = invoice["period_to"]
            row["_period_label"] = invoice.get("period_label")
        merged.append(row)

    for pay in pay_emps:
        if id(pay) in used:
            continue
        warnings.append(
            f"{pay.get('Employee Name')}：仅出现在 Payroll，已追加（Public Liability=0）"
        )
        row = dict(pay)
        row.setdefault("Employer's & Public Liability", 0.0)
        if invoice.get("period_from"):
            row["_period_from"] = invoice["period_from"]
            row["_period_to"] = invoice["period_to"]
            row["From"] = invoice["period_from"]
            row["To"] = invoice["period_to"]
        merged.append(row)

    warnings.extend(invoice.get("warnings") or [])
    warnings.extend(payroll.get("warnings") or [])
    return merged, warnings


def convert_sources(
    source_paths: list[Path],
    output_path: Path,
    *,
    template_path: Path | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    fill_fx: bool = True,
    convert_mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    del pn_meta, registry_dir, fill_fx
    from convert_mapping import resolve_convert_mapping
    from profiles.cyprus_payroll_calc.convert import set_recurring_fees, write_cyprus_l

    mapping_in = dict(convert_mapping) if isinstance(convert_mapping, dict) else {}
    mapping_in.setdefault("pdfProfileId", "at_technical_cyprus")
    mapping = resolve_convert_mapping("cyprus_payroll_calc", mapping_in)

    paths = [Path(p).resolve() for p in source_paths]
    if not paths:
        raise ValueError("未提供 PDF")
    pdfs = [p for p in paths if p.suffix.lower() == ".pdf"]
    if not pdfs:
        raise ValueError("at_technical_cyprus 需要 PDF（Invoice + Payroll Calculation）")

    invoice = payroll = None
    warnings: list[str] = []
    for p in pdfs:
        text = extract_pdf_text(p)
        kind = classify_at_pdf(text, p)
        if kind == "invoice":
            if invoice:
                warnings.append(f"重复 Invoice，忽略: {p.name}")
                continue
            invoice = parse_at_invoice_pdf(p, convert_mapping=mapping)
        elif kind == "payroll":
            if payroll:
                warnings.append(f"重复 Payroll，忽略: {p.name}")
                continue
            payroll = parse_at_payroll_pdf(p)
        else:
            warnings.append(f"无法识别 PDF 类型，尝试按内容解析: {p.name}")
            if "empl.id:" in text.lower():
                payroll = parse_at_payroll_pdf(p)
            else:
                invoice = parse_at_invoice_pdf(p, convert_mapping=mapping)

    if not invoice and not payroll:
        raise ValueError("未能解析任何 A&T Invoice / Payroll PDF")
    if not invoice or not payroll:
        warnings.append(
            "建议同时上传 Invoice 与 Payroll Calculation；当前缺一份，已尽力合并"
        )

    employees, merge_warnings = merge_invoice_and_payroll(invoice, payroll)
    warnings.extend(merge_warnings)
    employees, alias_warnings = recover_cyprus_l_field_aliases(employees)
    warnings.extend(alias_warnings)

    tpl = (template_path or get_region_template("Cyprus")).resolve()
    if not tpl.is_file():
        raise FileNotFoundError(f"Cyprus 母版不存在: {tpl}")

    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(tpl, output_path)
    wb = load_workbook(output_path)
    try:
        if CYPRUS_L_SHEET not in wb.sheetnames:
            raise ValueError(f"母版缺少 {CYPRUS_L_SHEET}")
        import profiles.cyprus_payroll_calc.convert as cyprus_mod

        prev = cyprus_mod._ACTIVE_MAPPING
        cyprus_mod._ACTIVE_MAPPING = mapping
        try:
            write_cyprus_l(wb[CYPRUS_L_SHEET], employees)
            set_recurring_fees(wb, employees)
        finally:
            cyprus_mod._ACTIVE_MAPPING = prev
        wb.save(output_path)
    finally:
        wb.close()

    return {
        "ok": True,
        "profile_id": "at_technical_cyprus",
        "region": "Cyprus",
        "source_kind": "pdf",
        "output": str(output_path),
        "employee_count": len(employees),
        "parsed": [
            {
                "employee_name": e.get("Employee Name"),
                "base_salary": e.get("Base salary"),
                "er_contributions": e.get("Employer's contributions"),
                "liability": e.get("Employer's & Public Liability"),
            }
            for e in employees
        ],
        "warnings": warnings,
        "fx_rate": None,
        "pn_meta": None,
        "invoice_no": (invoice or {}).get("invoice_no"),
        "period_label": (invoice or payroll or {}).get("period_label"),
    }


def convert_pdfs(
    pdf_paths: list[Path],
    output_path: Path,
    *,
    template_path: Path | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    fill_fx: bool = True,
    convert_mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    return convert_sources(
        pdf_paths,
        output_path,
        template_path=template_path,
        pn_meta=pn_meta,
        registry_dir=registry_dir,
        fill_fx=fill_fx,
        convert_mapping=convert_mapping,
    )


def convert_pdf(
    pdf_path: Path,
    output_path: Path,
    *,
    template_path: Path | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    fill_fx: bool = True,
    convert_mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    return convert_pdfs(
        [pdf_path],
        output_path,
        template_path=template_path,
        pn_meta=pn_meta,
        registry_dir=registry_dir,
        fill_fx=fill_fx,
        convert_mapping=convert_mapping,
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="A&T Cyprus PDF → Cyprus-L")
    parser.add_argument("sources", nargs="+", type=Path)
    parser.add_argument("-o", "--output", type=Path, required=True)
    parser.add_argument("-t", "--template", type=Path, default=None)
    args = parser.parse_args(argv)
    result = convert_sources(args.sources, args.output, template_path=args.template)
    print(result)
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
