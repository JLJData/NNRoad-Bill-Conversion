# -*- coding: utf-8 -*-
from __future__ import annotations

import re

try:
    from pypinyin import Style, lazy_pinyin
except ImportError:  # pragma: no cover
    Style = None  # type: ignore[misc, assignment]
    lazy_pinyin = None

_HAN_RE = re.compile(r"[\u4e00-\u9fff]")
_COMPACT_KEEP_RE = re.compile(r"[^a-z0-9\u4e00-\u9fff]+")


def norm_person_name(value: object) -> str:
    if value is None:
        return ""
    s = str(value).replace("\u3000", " ").strip().lower()
    return re.sub(r"\s+", " ", s)


def compact_person_name(value: object) -> str:
    """匹配用：去掉空白与连字符（Kevin Will maser == Kevin Willmaser；Gui-liang == Guiliang）。"""
    return _COMPACT_KEEP_RE.sub("", norm_person_name(value))


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
    # 拼音拆写：gui ↔ guiliang
    if len(short_tok) >= 3:
        return any(lt.startswith(short_tok) or short_tok.startswith(lt) and len(lt) >= 3 for lt in long_tokens)
    return False


def _han_to_pinyin(value: str) -> str:
    """汉字转无声调拼音，拉丁片段原样保留。无汉字或未装 pypinyin 时返回空串。"""
    text = str(value or "")
    if not text or lazy_pinyin is None or Style is None or not _HAN_RE.search(text):
        return ""
    parts: list[str] = []
    latin: list[str] = []

    def flush_latin() -> None:
        if latin:
            parts.append("".join(latin))
            latin.clear()

    for ch in text:
        if _HAN_RE.match(ch):
            flush_latin()
            py = lazy_pinyin(ch, style=Style.NORMAL)
            parts.append(py[0] if py else ch)
        else:
            latin.append(ch)
    flush_latin()
    return norm_person_name(" ".join(parts))


def _name_variants(value: str) -> list[str]:
    raw = str(value or "").strip()
    if not raw:
        return []
    out = [raw]
    py = _han_to_pinyin(raw)
    if py and py not in out:
        out.append(py)
    han = "".join(_HAN_RE.findall(raw))
    if han and han != raw:
        out.append(han)
        han_py = _han_to_pinyin(han)
        if han_py and han_py not in out:
            out.append(han_py)
    return out


def _score_folded_names(a: str, b: str) -> int:
    """
    精确 100；去空格/连字符后相同 100；一方包含另一方 80；较短名全部 token 在较长名中 70。
    单字母 token 可匹配首字母（M. I. Ghazi ≈ Muhammad Ismail Ghazi）。
    """
    a = norm_person_name(a)
    b = norm_person_name(b)
    if not a or not b:
        return 0
    if a == b:
        return 100
    ca, cb = compact_person_name(a), compact_person_name(b)
    if ca and cb and ca == cb:
        return 100
    if a in b or b in a:
        return 80
    if ca and cb and (ca in cb or cb in ca) and min(len(ca), len(cb)) >= 4:
        return 80
    ta, tb = _name_tokens(a), _name_tokens(b)
    if not ta or not tb:
        return 0

    def covered(short: list[str], long_tokens: list[str]) -> bool:
        long_set = set(long_tokens)
        return bool(short) and all(_token_covered(t, long_set) for t in short)

    if covered(ta, tb) or covered(tb, ta):
        return 70
    return 0


def score_person_name_match(excel_name: str, candidate_name: str) -> int:
    """
    精确 100；去空格后相同 100；一方包含另一方 80；较短名全部 token 在较长名中 70。
    单字母 token 可匹配首字母（M. I. Ghazi ≈ Muhammad Ismail Ghazi）。
    含汉字时再按拼音比一遍（徐桂亮 Leon Xu ≈ XU Gui-liang）。
    """
    best = 0
    for left in _name_variants(excel_name):
        for right in _name_variants(candidate_name):
            best = max(best, _score_folded_names(left, right))
            if best >= 100:
                return 100
    return best


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
