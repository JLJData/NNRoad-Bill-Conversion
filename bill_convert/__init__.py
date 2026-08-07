# -*- coding: utf-8 -*-
"""多引擎账单转换公共逻辑（mapping 驱动；引擎 profile 保留地区/业务 hook）。"""

from bill_convert.formula_copy import copy_row_formulas, shift_row_formula
from bill_convert.formula_layout import (
    apply_employee_formula_styles,
    resolve_formula_rows_layout,
    tw_l_row_for_data_row,
)
from bill_convert.header_scan import find_header_row_by_markers
from bill_convert.headers import (
    build_header_cols,
    build_header_map,
    build_qualified_header_cols,
    build_qualified_header_map,
    list_qualified_header_cells,
    resolve_header_cols,
    resolve_target_col,
)
from bill_convert.mapping_spec import mapping_section
from bill_convert.meta_period import parse_period, payroll_month_start, read_summary_meta
from bill_convert.person import norm_person_name
from bill_convert.target_l_layout import (
    resolve_target_l_layout,
    resolve_target_l_sheet_name,
)
from bill_convert.template_rows import (
    clear_row_values,
    count_data_slots,
    list_formula_example_rows,
)

__all__ = [
    "apply_employee_formula_styles",
    "build_header_cols",
    "build_header_map",
    "build_qualified_header_cols",
    "build_qualified_header_map",
    "clear_row_values",
    "list_qualified_header_cells",
    "copy_row_formulas",
    "count_data_slots",
    "find_header_row_by_markers",
    "list_formula_example_rows",
    "mapping_section",
    "norm_person_name",
    "parse_period",
    "payroll_month_start",
    "read_summary_meta",
    "resolve_formula_rows_layout",
    "resolve_header_cols",
    "resolve_target_col",
    "resolve_target_l_layout",
    "resolve_target_l_sheet_name",
    "shift_row_formula",
    "tw_l_row_for_data_row",
]
