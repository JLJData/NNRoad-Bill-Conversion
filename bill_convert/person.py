# -*- coding: utf-8 -*-
from __future__ import annotations

import re

from openpyxl.worksheet.worksheet import Worksheet


def norm_person_name(value: object) -> str:
    if value is None:
        return ""
    s = str(value).replace("\u3000", " ").strip().lower()
    return re.sub(r"\s+", " ", s)


def compact_person_name(value: object) -> str:
    """匹配用：去掉全部空白后的姓名（Kevin Will maser == Kevin Willmaser）。"""
    return re.sub(r"\s+", "", norm_person_name(value))


def _name_tokens(value: str) -> list[str]:
    """空白分词，并去掉标点（M. / I. → m / i）。"""
    out: list[str] = []
    for t in norm_person_name(value).split(" "):
        cleaned = re.sub(r"[^a-z0-9]", "", t)
        if cleaned:
            out.append(cleaned)
    return out


def _token_covered(short_tok: str, long_tokens: set[str]) -> bool:
    if short_tok in long_tokens:
        return True
    # 单字母缩写：匹配任一以该字母开头的 token（M → Muhammad）
    if len(short_tok) == 1:
        return any(lt.startswith(short_tok) for lt in long_tokens)
    return False


def score_person_name_match(excel_name: str, candidate_name: str) -> int:
    """
    精确 100；去空格后相同 100；一方包含另一方 80；较短名全部 token 在较长名中 70。
    单字母 token 可匹配首字母（M. I. Ghazi ≈ Muhammad Ismail Ghazi）。
    """
    a = norm_person_name(excel_name)
    b = norm_person_name(candidate_name)
    if not a or not b:
        return 0
    if a == b:
        return 100
    if compact_person_name(a) == compact_person_name(b):
        return 100
    if a in b or b in a:
        return 80
    ta, tb = _name_tokens(a), _name_tokens(b)
    if not ta or not tb:
        return 0

    def covered(short: list[str], long_tokens: list[str]) -> bool:
        long_set = set(long_tokens)
        return bool(short) and all(_token_covered(t, long_set) for t in short)

    # 双向：缩写↔全名（token 数相同时也要试两侧）
    if covered(ta, tb) or covered(tb, ta):
        return 70
    return 0


def person_name_labels_match(
    left_labels: list[str],
    right_labels: list[str],
    *,
    min_score: int = 70,
) -> bool:
    """任意一对姓名（可跨 CN/EN 列）达到 min_score 即视为同一人。"""
    for raw_a in left_labels:
        a = str(raw_a or "").strip()
        if not a:
            continue
        for raw_b in right_labels:
            b = str(raw_b or "").strip()
            if not b:
                continue
            if score_person_name_match(a, b) >= min_score:
                return True
    return False


def bill_employee_like_entry(emp: dict, entry: dict, *, min_score: int = 70) -> bool:
    # TW: CN/EN Name；China: 姓名；HK: Name of Employee / EE Name / Name；UAE: Employee/English Name
    bill = [
        emp.get("CN Name"),
        emp.get("EN Name"),
        emp.get("姓名"),
        emp.get("Name of Employee"),
        emp.get("EE Name"),
        emp.get("Name"),
        emp.get("Employee Name"),
        emp.get("English Name"),
    ]
    cfg = [entry.get("cnName"), entry.get("enName")]
    return person_name_labels_match(
        [str(x) for x in bill if x is not None],
        [str(x) for x in cfg if x is not None],
        min_score=min_score,
    )
