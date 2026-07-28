# -*- coding: utf-8 -*-
from __future__ import annotations

import re

from openpyxl.worksheet.worksheet import Worksheet


def norm_person_name(value: object) -> str:
    if value is None:
        return ""
    s = str(value).replace("\u3000", " ").strip().lower()
    return re.sub(r"\s+", " ", s)
