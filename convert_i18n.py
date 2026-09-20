# -*- coding: utf-8 -*-
"""转换服务用户可见文案：中文为业务源文案，英文由 API 出口按 Accept-Language 翻译。

业务代码继续 raise / warnings.append 中文即可；不要在各 profile 里散落 t()。
Office Java 须转发浏览器 Accept-Language。
"""
from __future__ import annotations

import re
from contextvars import ContextVar, Token
from typing import Any

_locale: ContextVar[str] = ContextVar("convert_locale", default="zh")

_MESSAGES: dict[str, dict[str, str]] = {
    "zh": {
        # --- common ---
        "common.no_pdf": "未提供 PDF",
        "common.no_excel": "未提供 Excel",
        "common.no_source": "未提供源文件",
        "common.pdf_not_found": "PDF 不存在: {path}",
        "common.excel_not_found": "Excel 不存在: {path}",
        "common.template_not_found": "母版不存在: {path}",
        "common.source_not_found": "源文件不存在: {path}",
        "common.template_sheet_missing": "母版缺少 sheet「{sheet}」，现有: {sheets}",
        "common.template_sheet_missing_simple": "母版缺少 {sheet}",
        "common.unsupported_file_types": "不支持的文件类型: {names}",
        "common.no_employee_rows": "未解析到员工行: {name}",
        "common.pdf_no_text": "PDF 无文本层或抽取为空: {name}",
        "common.excel_no_sheets": "Excel 无可用工作表",
        "common.no_employees_parsed": "未解析到员工",
        "common.employee": "员工",
        "common.warn.invoice_date_parse_fail": "无法解析发票日期: {date}",
        # --- text_extract ---
        "text_extract.missing_pypdf": "缺少依赖 pypdf，请执行: pip install pypdf",
        "text_extract.pdf_encrypted": "PDF 已加密，无法读取: {path}",
        "text_extract.no_text": "未能从 PDF 抽出文字（可能是扫描件，需 OCR）: {path}",
        # --- registry ---
        "registry.unknown_profile": "未知 PDF profile「{profile_id}」，已知: {known}",
        "registry.no_convert_pdf": "模块 {module} 缺少 convert_pdf()",
        # --- runner ---
        "runner.pdf_missing": "未提供 PDF",
        "runner.source_missing": "未提供源文件",
        "runner.profile_no_batch": "当前 pdf_profile「{profile_id}」不支持批量（共 {count} 份），请一次只传 1 份，或改用支持批量的 profile",
        "runner.profile_no_excel": "pdf_profile「{profile_id}」暂不支持 Excel 源，请上传 PDF 或已成型的 UK-L Excel",
        "runner.auto_detect_fail": "无法自动识别 PDF 供应商版式，请显式传入 profile_id（如 eor_uk / topsource_uk）",
        "runner.excel_need_profile_id": "Excel 源须显式传入 profile_id（如 topsource_uk / auxilium_uae）；无法从正文自动识别",
        "runner.no_result": "PDF 转换未返回结果",
        "runner.mix_prior_pdf": "同批含 PDF 与 Excel，已优先采用 PDF（忽略 {count} 个 Excel）",
        "runner.mix_prior_excel": "同批含 PDF 与 Excel，当前版式仅支持 Excel，已采用 Excel（忽略 {count} 个 PDF）",
        "runner.engine_no_convert": "引擎模块缺少 convert(): {module}",
        "runner.region_required": "未指定 template_path 时必须提供 region",
        "runner.pn_meta_must_object": "pn_meta 必须是 JSON 对象",
        "runner.convert_mapping_must_object": "convert_mapping 必须是 JSON 对象",
        "runner.employee_directory_must_array": "employee_directory 必须是 JSON 数组",
        # --- post_checks ---
        "post_checks.empty_result": "PDF 解析结果为空，已中止写出源表",
        "post_checks.no_employees": "PDF 解析未得到任何员工（profile={profile}），版式可能已变更，已中止写出以免产生空/错表",
        "post_checks.fatal_warnings": "PDF 关键字段解析失败或勾稽不过（profile={profile}），已中止写出以免静默错数。详情：{preview}{more}",
        "post_checks.fatal_warnings_more": " 等共 {count} 条",
        "post_checks.placeholder_names": "PDF 未解析到真实员工姓名（均为占位名，profile={profile}），版式可能已变更，已中止写出",
        "post_checks.topsource_no_labor_usd": "TopSource PDF 第 {index} 份未解析到人工成本 USD，版式可能已变更，已中止写出",
        # --- fatal warnings (profile warnings list) ---
        "warn.no_employee_name": "未解析到员工姓名",
        "warn.no_gross_salary": "未解析到 Gross Salary",
        "warn.monthly_salary_block_missing": "未匹配到「Monthly Salary / Empr Taxes / Empr Contributions」描述，版式可能已变更",
        "warn.labor_net_mismatch": "工资构成合计 {labor} 与行净额 {net} 不一致",
        "warn.service_fee_mismatch": "行净额+ServiceFee={expect} 与 Invoice Total {total} 不一致",
        "warn.no_labor_usd": "未解析到人工成本 USD 打包金额",
        "warn.no_quarter_salary_pkr": "未解析到季度薪资 PKR",
        "warn.no_quarter_period": "未解析到季度账期",
        "warn.no_federal_tax": "未从 PDF 解析到 Federal IT / Sindh Sales Tax 金额行",
        "warn.cgst_only": "仅解析到 CGST，Business Tax 按 ×2 暂估",
        "warn.no_gst": "未解析到 GST，Business Tax 置 0",
        "warn.payroll_name_unmatched": "{name}：Payroll 中未匹配到同名员工，EE 扣款列为空",
        "warn.unrecognized_pdf_type": "无法识别 PDF 类型，尝试按内容解析: {name}",
        # --- eor_uk ---
        "eor_uk.excel_single_only": "eor_uk Excel 暂仅支持单员工竖表（收到 {count} 份），请一次一份",
        "eor_uk.merge_pdf_one": "eor_uk 混传时 PDF 仅支持 1 份（收到 {count} 份）",
        "eor_uk.merge_excel_one": "eor_uk 混传时 Excel 仅支持 1 份（收到 {count} 份）",
        "eor_uk.pdf_no_batch": "eor_uk PDF 暂不支持批量（共 {count} 份），请一次只传 1 份",
        "eor_uk.excel_missing": "未提供 Excel",
        "eor_uk.excel_not_found": "Excel 不存在: {path}",
        "eor_uk.uk_l_multi": "多份已是 UK-L 的 Excel 无法自动合并，请只传一份",
        "eor_uk.uk_l_passthrough": "源表已是 UK-L，已原样用作转换输入",
        "eor_uk.no_source": "未提供可用的 PDF 或 Excel 源",
        "eor_uk.template_missing": "UK 母版不存在: {path}",
        "eor_uk.no_employee_name": "EOR UK PDF 未解析到员工姓名，版式可能已变更",
        "eor_uk.no_salary_fields": "EOR UK PDF 未解析到 Gross Salary / Empr Taxes / Empr Contributions，版式可能已变更",
        "eor_uk.labor_mismatch": "EOR UK PDF 工资构成合计 {labor} 与行净额 {net} 不一致，已中止写出",
        "eor_uk.service_fee_mismatch": "EOR UK PDF 行净额+ServiceFee={expect} 与 Invoice Total {total} 不一致，已中止写出",
        "eor_uk.excel_no_gross": "EOR UK Excel 未解析到 Gross Salary（Gross Pay），请检查列映射或版式",
        "eor_uk.merge_no_gross": "EOR UK 合并后未得到 Gross Salary，请检查 PDF/Excel 或列名对照",
        # --- topsource_uk ---
        "topsource_uk.no_employee_name": "TopSource PDF 未解析到员工姓名，版式可能已变更",
        "topsource_uk.no_labor_usd": "TopSource PDF 未解析到人工成本 USD，版式可能已变更",
        "topsource_uk.no_source": "未提供源文件",
        "topsource_uk.uk_l_multi": "多份已是 UK-L 的 Excel 无法自动合并，请只传一份或传供应商原始发票",
        "topsource_uk.uk_l_passthrough": "源表已是 UK-L，已原样用作转换输入",
        # --- auxilium_uae ---
        "auxilium_uae.not_recognized": "未识别为 Auxilium/UAE Payroll Draft（各 sheet 前 {scan_max} 行未见员工/薪酬表头）。sheet「{sheet}」样例: {sample}",
        "auxilium_uae.no_id_name_columns": "未识别员工工号/姓名列：请在列名对照中配置到 Emp ID / Employee Name，或保证表头与这两列同名",
        "auxilium_uae.uae_l_multi": "多份已是 UAE-L 的 Excel 无法自动合并，请只传一份",
        "auxilium_uae.no_mix_uae_l_draft": "请不要混传已成型 UAE-L 与 Payroll Draft",
        "auxilium_uae.uae_l_passthrough": "源表已是 UAE-L，已原样用作转换输入",
        "auxilium_uae.excel_only_main": "auxilium_uae 主源仅支持 Excel Payroll Draft；无法识别的 PDF: {names}。两份 Invoice 请用 INV- 编号区分：数字大的为上期 Admin Fee，数字小的为本期。",
        "auxilium_uae.need_draft_excel": "请至少上传一份 Auxilium Payroll Draft Excel（员工薪资主源）。两份 Invoice PDF 按发票号：数字大的=上期 Admin Fee，数字小的=本期。",
        "auxilium_uae.no_pdf_support": "auxilium_uae 暂不支持 PDF，请上传 Payroll Draft Excel",
        # --- biz_solutions_india ---
        "biz_solutions_india.no_employee_name": "未能解析员工姓名: {name}",
        "biz_solutions_india.no_ctc": "未能解析 Monthly CTC: {name}",
        "biz_solutions_india.cgst_only_fatal": "Biz Solutions PDF 仅解析到 CGST、缺少 SGST，版式可能已变更；已中止写出以免税额偏错",
        "biz_solutions_india.no_gst_fatal": "Biz Solutions PDF 未解析到 GST，版式可能已变更；已中止写出以免税额被置 0",
        "biz_solutions_india.single_pdf_only": "biz_solutions_india 暂只支持单份 PDF（一人一票）",
        "biz_solutions_india.need_tax_invoice_pdf": "请上传 Biz Solutions Tax Invoice PDF",
        # --- at_technical_cyprus ---
        "at_technical.invoice_no_blocks": "A&T Invoice 未解析到员工块: {name} — {detail}",
        "at_technical.payroll_no_employees": "A&T Payroll 未解析到员工: {name}",
        "at_technical.no_parse_results": "未提供可用的 Invoice / Payroll PDF 解析结果",
        "at_technical.payroll_name_mismatch": "A&T Cyprus：员工「{name}」在 Payroll 中未匹配到同名，版式或姓名不一致，已中止写出以免扣款列空",
        "at_technical.need_pdf_pair": "at_technical_cyprus 需要 PDF（Invoice + Payroll Calculation）",
        "at_technical.no_parsed_pdfs": "未能解析任何 A&T Invoice / Payroll PDF",
        # --- connect_uae ---
        "connect_uae.not_invoice": "不是 Connect Resources 税票: {name}",
        "connect_uae.split_mismatch": "Connect UAE：员工「{name}」映射拆分合计 {split_sum} ≠ PDF 月薪 {payroll}，已中止写出",
        "connect_uae.need_tax_invoice_pdf": "connect_uae 请上传 Connect Tax Invoice PDF",
        # --- safeguard_italy ---
        "safeguard_italy.no_header_row": "未找到 SafeGuard 员工表头行（Employee ID / Employee Name）",
        "safeguard_italy.italy_l_multi": "多份已是 Italy-L 的 Excel 无法自动合并，请只传一份",
        "safeguard_italy.no_mix": "请不要混传已成型 Italy-L 与 SafeGuard 源账单",
        "safeguard_italy.italy_l_passthrough": "源表已是 Italy-L，已原样用作转换输入",
        "safeguard_italy.excel_only": "safeguard_italy 目前仅支持 Excel 源账单，请上传 SGWI Payroll xlsx",
        "safeguard_italy.no_pdf_runtime": "safeguard_italy 主源为 Excel，不支持 PDF；请上传 SGWI Payroll xlsx",
        # --- panda_work_pk ---
        "panda_work_pk.no_federal_tax": "Panda Work PDF 未解析到 Federal IT / Sindh Sales Tax，版式可能已变更，已中止写出",
        "panda_work_pk.no_employee_name": "Panda Work PDF 未解析到员工姓名，版式可能已变更",
        "panda_work_pk.no_salary_pkr": "Panda Work PDF 未解析到季度薪资 PKR，版式可能已变更",
        "panda_work_pk.no_quarter_period": "Panda Work PDF 未解析到季度账期，版式可能已变更",
        "panda_work_pk.not_invoice": "不是 Panda Work Global Pakistan 发票: {name}",
        # --- engines / region ---
        "engines.unknown": "未知转换引擎「{engine_id}」，已知: {known}",
        "region.unknown": "未知地区「{region}」，已知: {known}",
        # --- api ---
        "api.pdf_only": "仅支持 .pdf，当前: {suffix}",
        "api.excel_only": "仅支持 Excel（.xlsx/.xlsm），当前: {suffix}",
        "api.excel_only_short": "仅支持 Excel，当前: {suffix}",
        "api.empty_upload": "上传文件为空",
        "api.empty_upload_named": "上传文件为空: {name}",
        "api.need_pdf": "请至少上传一个 PDF",
        "api.need_source": "请至少上传一个源文件",
        "api.pdf_excel_only": "仅支持 .pdf / .xlsx / .xlsm，当前: {name}",
        "api.pdf_excel_xlsx_only": "仅支持 .xlsx/.xlsm/.pdf，当前: {suffix}",
        "api.template_xlsx": "母版仅支持 .xlsx/.xlsm",
        "api.template_xlsx_with_suffix": "母版仅支持 .xlsx/.xlsm，当前: {suffix}",
        "api.convert_fail": "转换失败: {detail}",
        "api.vendor_batch_fail": "供应商源批量转换失败: {detail}",
        "api.pdf_convert_fail": "PDF 转换失败: {detail}",
        "api.pdf_batch_fail": "PDF 批量转换失败: {detail}",
        "api.blocked_file_type": "blocked file type: {suffix}",
        "api.engine_id_required": "engineId required",
        "api.pn_meta_invalid": "pn_meta 无效: {detail}",
        "api.convert_mapping_invalid": "convert_mapping 无效: {detail}",
        "api.employee_directory_invalid": "employee_directory 无效: {detail}",
        "api.bypass_classify_fail": "旁路文件识别失败: {detail}",
        "api.region_template_not_found": "地区母版不存在: {path}",
        "api.decrypt_fail": "解密失败: {detail}",
        "api.no_output_file": "转换完成但未生成输出文件",
        "api.excel_snapshot_fail": "Excel 快照失败: {detail}",
        "api.hf_snapshot_fail": "HF 快照失败: {detail}",
        "api.file_role_fail": "文件角色探测失败: {detail}",
        "api.region_required_no_template": "未上传母版时必须提供 region",
        # --- mapping inspect ---
        "inspect.pdf_labels_unsupported": "引擎「{engine_id}」/版式「{profile_id}」暂不支持 PDF 标签识别",
        "inspect.header_unsupported": "引擎「{engine_id}」暂不支持表头识别",
        "inspect.sheet_not_found": "未找到工作表「{sheet}」",
        "inspect.cannot_open_source": "无法打开源表: {detail}",
        "inspect.calculation_header_fail": "已找到 Calculation，但解析表头失败: {detail}",
        "inspect.no_calculation_or_italy_l": "未找到 Calculation 或 Italy-L 工作表",
        "inspect.template_not_found": "母版文件不存在: {template_path}",
        "inspect.template_sheet_not_found": "母版中未找到 sheet「{sheet}」",
        "inspect.no_sheets": "工作簿无工作表",
        "inspect.uae_l_or_draft_missing": "未找到工作表「UAE-L」，且无法识别供应商 Payroll Draft 表头。详情: {detail}",
        "inspect.uk_vertical_no_labels": "竖表未识别到标签（sheet={sheet}, labelCol={label_col}, amountCol={amount_col}），请确认是 UK-L 或 Analysis of Payroll Totals",
        "inspect.uae_vendor_header_hint": "已识别供应商原始表头，请在「列名对照」中配置 供应商列 → UAE-L 列",
        "inspect.safeguard_calculation_hint": "已识别 SafeGuard Calculation 表头。请在「列名对照」配置后点「保存映射」；未保存的对照不会用于转换。",
        # --- api (inspect/bypass) ---
        "api.no_matching_plugin": "无匹配插件",
        "api.not_bypass_file": "未识别为旁路文件",
        # --- ee_code (shared) ---
        "ee_code.no_name": "源表员工姓名为空，无法匹配 EE Code",
        "ee_code.no_directory": "未提供客户员工目录，无法匹配 EE Code",
        "ee_code.not_matched": "未匹配到 EE Code：{label}",
        "ee_code.ambiguous": "EE Code 匹配歧义（{count} 人同分）：{label}",
        "common.empty_name": "（空）",
        "common.write_uk_fx_d24_fail": "写入 UK-L!D24 汇率失败: {detail}",
        "common.write_pn_fx_fail": "写入 PN 汇率失败: {detail}",
        "common.fx_online_fallback_d24": "在线汇率失败，沿用源表 D24={rate}: {detail}",
        # --- engine (region convert) ---
        "engine.person_row": "第{index}人：{detail}",
        "engine.ee_sheet_row": "{sheet} EE 第{index}人：{detail}",
        "engine.formula_style_count": "映射员工公式样式条数: {count}",
        "engine.formula_pair": "公式配对：第{index}人 → {main_sheet}第{main_row}行 / {ee_sheet} EE第{ee_row}行",
        "engine.fx_from_bill": "汇率已取自供应商账单 {source} = {rate}",
        # --- template_rows (UI formula-row picker) ---
        "template.row_default": "第 {row} 行（默认）",
        "template.row_plain": "第 {row} 行",
        "template.row_with_marker": "第 {row} 行 · {marker}",
        # --- convert_checks ---
        "checks.header_fallback": "未找到表头「{field}」({names})，已回退固定列 {fallback}；若供应商调整了列位置请改映射/表头，勿依赖列号",
        "checks.column_missing_skip": "未找到列「{field}」({names})，该字段跳过",
        "checks.column_rename_no_hit": "columnRename 已配置但未命中任何源表头，请核对映射里的供应商列名是否与账单一致{examples_suffix}",
        "checks.column_rename_examples_suffix": "（样例未命中: {examples}）",
        "checks.column_rename_partial_miss": "columnRename 部分未命中（{count} 项），例如: {examples}",
        "checks.column_rename_idle": "columnRename 本月未使用（源表已是目标列名，{idle_aliases} 条别名闲置）",
        "checks.result_empty": "转换结果为空，请人工核对",
        "checks.zero_employees": "转换结果员工数为 0，请核对源表姓名列/表头行是否变化",
        "checks.too_many_employees": "转换结果员工数偏多（{count}），请确认是否误读表头或空行",
        "checks.many_fallback_cols": "有 {count} 个字段靠固定列号回退写入，供应商账单列位置可能已变，建议按表头核对",
        # --- eor_uk (informational) ---
        "eor_uk.not_eor_keyword": "正文未出现 EOR Services 关键字，可能不是本 profile 对应的发票",
        "eor_uk.vertical_col_detect": "竖表列探测：标签列={label_col} 金额列={amount_col}{suffix}",
        "eor_uk.vertical_col_rename_suffix": "；columnRename={count} 条",
        "eor_uk.vertical_col_default_suffix": "；使用内置默认别名",
        "eor_uk.unmapped_labels": "未映射标签（可在 Office「列名对照」补充）: {labels}",
        "eor_uk.excel_no_name_hint": "Excel 竖表通常无员工姓名，UK-L 标题暂用 Employee；请在 PN 元数据补全",
        "eor_uk.er_nic_pension_incomplete": "未完整解析 ER NIC / ER Pension，请核对列名对照（Employer NI / Employer Pension）",
        "eor_uk.ee_side_zeroed": "EE 侧 PAYE / EE NIC / EE Pension 发票未提供，已置 0（需人工或其它来源）",
        "eor_uk.service_fee_not_written": "PDF Service Fee={service_fee}（与 PN Management Fee / Recurring Fee 不同项，未自动写入 UK!H）",
        "eor_uk.pdf_excel_conflict": "PDF/Excel「{label}」不一致：PDF={pdf_val} Excel={excel_val}，已采用 Excel",
        "eor_uk.merge_note": "已合并 PDF（姓名/发票/Service Fee）与 Excel（工资明细，含 EE 侧）；金额冲突以 Excel 为准",
        "eor_uk.merge_no_name": "合并后仍无员工姓名，UK-L 标题暂用 Employee；请在 PN 元数据补全",
        # --- topsource_uk (informational) ---
        "topsource_uk.usd_bundle_manual": "TopSource PDF 仅为 USD 打包价，无 GBP 明细；Gross/Holiday/PAYE 等请按截图人工补齐",
        "topsource_uk.labor_usd_note": "发票人工打包 USD={amount}（供核对，未写入 UK-L 明细）",
        "topsource_uk.service_charge_note": "发票 Service Charge USD={amount}（与 PN Management Fee 公式不同，未自动写入）",
        "topsource_uk.not_topsource_keyword": "正文未出现 TopSource 关键字，可能不是本 profile 对应的发票",
        "topsource_uk.excel_not_topsource": "表头未出现 TopSource，可能不是本版式 Excel",
        "topsource_uk.ts_margin_note": "发票含 TS Margin/服务费 GBP={amount}（与 PN Management Fee 不同，未自动写入）",
        "topsource_uk.name_fallback": "未从 A3 解析到姓名，已用 sheet/文件名兜底",
        "topsource_uk.no_gbp_amounts": "未从 Excel 解析到任何 GBP 明细金额",
        "topsource_uk.mix_prior_pdf": "同批含 PDF 与 Excel，已按 TopSource 主源仅采用 {pdf_count} 份 PDF，忽略 Excel: {names}",
        # --- auxilium_uae (informational) ---
        "auxilium_uae.maybe_not_draft": "文件可能不是 Auxilium Payroll Draft，仍尝试解析: {name}",
        "auxilium_uae.admin_fee_no_facts": "已识别 Admin Fee PDF 但未解析到事实",
        "auxilium_uae.admin_fee_parsed_dual": "已解析 Admin Fee：上期 {prev_no} VAT={prev_vat}；本期 {curr_no} VAT={curr_vat}",
        "auxilium_uae.admin_fee_parsed_single": "已解析 Admin Fee PDF Total VAT={vat} ({source_file})",
        # --- safeguard_italy (informational) ---
        "safeguard_italy.column_rename_empty": "columnRename 为空：仅同名列匹配。请确认已在映射里配置并点「保存映射」。",
        "safeguard_italy.column_rename_loaded": "columnRename 已加载 {count} 条对照",
        "safeguard_italy.maybe_not_safeguard": "文件可能不是 SafeGuard Italy 账单，仍尝试解析: {name}",
        "safeguard_italy.italy_l_reparse": "源表含 Italy-L 且配置了列名对照，已按对照重解析: {name}",
        "safeguard_italy.column_rename_no_hit": "columnRename 已配置但未命中任何源列，请核对供应商列名是否与示例表头一致",
        "safeguard_italy.column_rename_applied": "columnRename 本批命中写入 {count} 次（按源列计）",
        "safeguard_italy.ignored_non_excel": "safeguard_italy 主源为 Excel，已忽略: {names}",
        # --- connect_uae (informational) ---
        "connect_uae.expense_split": "报销「{names}」票面合计 {total} 未拆到人，已按人均 {each} 分摊；请核对",
        "connect_uae.split_all_basic": "{name}：映射未配置 Basic/Housing/Transport，已将 PDF 月薪 {payroll} 全部计入 Basic",
        "connect_uae.split_sum_mismatch_warn": "{name}：映射拆分合计 {split_sum} ≠ PDF 月薪 {payroll}",
        "connect_uae.batch_first_pdf": "本批 {count} 份 PDF，已采用 {name}",
        # --- at_technical (informational) ---
        "at_technical.invoice_missing_labels": "Invoice「{name}」缺供应商标签 {labels}，已跳过",
        "at_technical.invoice_no_base_salary": "Invoice「{name}」列名对照后无 Base salary（请确认 Gross Salary → Base salary）；金额仍保留在对照目标列",
        "at_technical.no_monthly_cost_rows": "未匹配到 Monthly cost 员工行",
        "at_technical.invoice_blocks_skipped": "识别到 {heads_seen} 个员工块但均被跳过（多半是 PDF 缺标签，而非列名对照问题）：{detail}",
        "at_technical.inspect.labels_hint": "已从供应商 PDF 抽取标签命中 {hit_total} 次（去重后 {uniq_total} 项）；请在「列名对照」中配置 供应商标签 → Cyprus-L 列{medical_note}",
        "at_technical.inspect.no_labels": "未识别到 PDF 标签",
        "at_technical.inspect.medical_note": "；Medical 当期+补收已合计写入 Medical Insurance（例：{name} Medical Insurance={amount}）",
        "at_technical.payroll_name_unparsed": "Payroll 某员工块无法解析姓名",
        "at_technical.liability_keys_differ": "{name}：完整名与截断名 Liability 键均有值且不同（{full} / {truncated}），将合计写入 Public Liabilit 列",
        "at_technical.liability_truncated": "{name}：有金额在截断表头「{typo}」，将写入 Public Liabilit 列",
        "at_technical.invoice_only": "仅有 Invoice：缺少 EE Social Ins / Tax / N.H.S.，对应列置空",
        "at_technical.payroll_only": "仅有 Payroll：缺少 Public Liability，对应列置 0",
        "at_technical.er_contrib_mismatch": "{name}：Invoice ER Contributions {inv_er} ≠ Payroll {pay_er}，已用 Invoice",
        "at_technical.payroll_only_employee": "{name}：仅出现在 Payroll，已追加（Public Liability=0）",
        "at_technical.duplicate_invoice": "重复 Invoice，忽略: {name}",
        "at_technical.duplicate_payroll": "重复 Payroll，忽略: {name}",
        "at_technical.suggest_both_pdfs": "建议同时上传 Invoice 与 Payroll Calculation；当前缺一份，已尽力合并",
        # --- biz_solutions_india (informational) ---
        "biz_solutions_india.no_period": "未解析到账期，India-L 账期将留空",
        "biz_solutions_india.ignored_non_pdf": "biz_solutions_india 主源为 PDF，已忽略: {names}",
        # --- panda_work_pk (informational) ---
        "panda_work_pk.no_quarter_salary_row": "{name}：无季度薪资，未写入 Base Salary",
        "panda_work_pk.eobi_it_empty": "{name}：E.O.B.I / IT 若映射未配置则 Pakistan-L 对应列留空",
        # --- vendor plugins ---
        "plugin.classify_fail": "插件 {plugin_id} 分类失败 {name}: {detail}",
        "plugin.parse_fail": "插件 {plugin_id} 解析失败: {detail}",
        "plugin.write_fail": "插件 {plugin_id} 写入失败: {detail}",
        "auxilium_plugin.admin_fee_split": "同批 Admin Fee 按发票号分流：上期(较大号) {prev_no} VAT={prev_vat}；本期(较小号) {curr_no} VAT={curr_vat}",
        "auxilium_plugin.batch_two_invoices": "本批 {count} 份发票，仅采用号最小与号最大两份",
        "auxilium_plugin.no_latest_vat": "Auxilium：无最新 Admin Fee Total VAT（请先上传 Admin Fee 发票，或与 Draft 同批上传），跳过 Business Tax",
        "auxilium_plugin.invalid_vat": "Auxilium：Total VAT 无效（{currency}），跳过 Business Tax",
        "auxilium_plugin.no_uae_sheet": "Auxilium：母版缺少 UAE sheet，无法写 Business Tax",
        "auxilium_plugin.no_prior_vat": "Auxilium：无已入账上期 Admin VAT，Business Tax 暂按最新 VAT 写入；请在映射中设置期初「上期 Total VAT」",
        # --- pakistan ---
        "pakistan.fees_not_matched": "{name}：映射 pakistanEmployeeFees 未匹配到 E.O.B.I / IT",
        "pakistan.bt_no_tax_amounts": "{name}：PDF 未解析到 Sindh Sales Tax / Federal IT 金额，Business Tax 未写入",
        "pakistan.bt_column_missing": "Pakistan 主表第 {header_row} 行未找到 Business Tax 列，跳过发票推导 Business Tax",
        "pakistan.bt_missing_coeff": "{name}：映射启用了发票推导 Business Tax，但缺少 BT 系数（请用含推导列的源表/PDF 转换）",
        "pakistan.default_person": "第{index}人",
        # --- china ---
        "china.no_directory_row": "第{index}人：未传入员工库，China!B 未写入库名称（姓名 {name}）",
        "china.ee_code_fail": "第{index}人：{detail}，China!B 未填",
        "china.name_not_matched": "第{index}人：姓名 {name} 未在员工库匹配，China!B 未填",
        "china.directory_name_empty": "第{index}人：姓名 {name} 已匹配员工库，但库中姓名为空，China!B 未填",
        "china.fx_fallback": "供应商账单未读到汇率（S-Payment Notice!C49 或「汇率」标签），已回退网上 CNY；请确认付款通知页有数值汇率（不要只留未计算的公式）",
        "china.formula_miss": "公式配对未命中：映射要求 China 示例行 {rows}，但实际全部落在默认行；请检查员工库姓名是否能与账单「姓名」匹配（含拼音）",
        "china.formula_no_fields": "映射有员工公式样式，但未找到 chinaExampleRow/mainExampleRow 字段（可能未保存成功）",
        # --- china_hrone ---
        "china_hrone.fx_not_read": "供应商账单未读到汇率（S-Payment Notice!{cell} 或「汇率」标签），PN FX 格未改",
        "china_hrone.fx_nnroad_fail": "汇率应按当月1号×0.97 取自 NNRoad，但未取到（{detail}），PN FX 格未改（不用供应商账单汇率）",
        "china_hrone.no_pn_sheet": "母版没有 PN 表，汇率已读到但未写入",
        "china_hrone.no_fx_row": "母版 PN 未找到 FX rate 行，汇率已读到但未写入",
        "china_hrone.no_names": "未从源表读到员工姓名，无法匹配 EE Code",
        "china_hrone.no_ee_sheet": "母版没有 China EE 表，EE Code 未写入",
        "china_hrone.no_directory": "未提供客户员工目录，无法匹配 EE Code",
        "unlock.empty": "源表为空，不是有效 Excel: {name}",
        "unlock.source_is_pdf": "源表实际是 PDF，不是 Excel: {name}",
        "unlock.not_xlsx_header": "源表不是有效 xlsx（文件头 {header}）: {name}",
        "unlock.after_decrypt": "源表解密后仍无法打开（{name}）: {detail}",
        "unlock.template_not_xlsx": "母版不是有效 xlsx（{name}）: {detail}",
        # --- tw ---
        "tw.fx_fallback": "Summary 未找到有效 Exchange rate，已回退 API TWD 汇率",
        # --- uk ---
        "uk.gross_salary_empty": "{sheet}（{name}）Gross Salary 为空/0，请按截图人工补齐 GBP 明细",
        # --- region pn expand ---
        "cyprus.pn_expand": "Cyprus PN 多人明细行扩行暂定：已扩 Cyprus/Cyprus EE（{count} 人），请人工核对 PN",
        "italy.pn_expand": "Italy PN 多人 Labor/Service Fee 行扩行暂定：已扩 Italy/Italy EE（{count} 人），PN 明细行请人工核对",
        "india.pn_expand": "India PN 多人明细行扩行暂定：已扩 India/India EE（{count} 人），请人工核对 PN",
        "pakistan.default_employee": "员工",
    },
    "en": {
        # --- common ---
        "common.no_pdf": "No PDF provided",
        "common.no_excel": "No Excel provided",
        "common.no_source": "No source files provided",
        "common.pdf_not_found": "PDF not found: {path}",
        "common.excel_not_found": "Excel not found: {path}",
        "common.template_not_found": "Template not found: {path}",
        "common.source_not_found": "Source file not found: {path}",
        "common.template_sheet_missing": "Template missing sheet \"{sheet}\"; available: {sheets}",
        "common.template_sheet_missing_simple": "Template missing sheet {sheet}",
        "common.unsupported_file_types": "Unsupported file types: {names}",
        "common.no_employee_rows": "No employee rows parsed: {name}",
        "common.pdf_no_text": "PDF has no text layer or extraction is empty: {name}",
        "common.excel_no_sheets": "Excel has no usable worksheets",
        "common.no_employees_parsed": "No employees parsed",
        "common.employee": "employee",
        "common.warn.invoice_date_parse_fail": "Cannot parse invoice date: {date}",
        # --- text_extract ---
        "text_extract.missing_pypdf": "Missing dependency pypdf. Run: pip install pypdf",
        "text_extract.pdf_encrypted": "PDF is encrypted and cannot be read: {path}",
        "text_extract.no_text": "Could not extract text from PDF (may be a scan; OCR required): {path}",
        # --- registry ---
        "registry.unknown_profile": "Unknown PDF profile \"{profile_id}\"; known: {known}",
        "registry.no_convert_pdf": "Module {module} is missing convert_pdf()",
        # --- runner ---
        "runner.pdf_missing": "No PDF provided",
        "runner.source_missing": "No source files provided",
        "runner.profile_no_batch": "pdf_profile \"{profile_id}\" does not support batch ({count} files). Upload 1 at a time, or use a batch-capable profile.",
        "runner.profile_no_excel": "pdf_profile \"{profile_id}\" does not support Excel sources. Upload a PDF or a ready UK-L Excel.",
        "runner.auto_detect_fail": "Cannot auto-detect PDF vendor layout. Please pass profile_id explicitly (e.g. eor_uk / topsource_uk).",
        "runner.excel_need_profile_id": "Excel sources require an explicit profile_id (e.g. topsource_uk / auxilium_uae); cannot auto-detect from content.",
        "runner.no_result": "PDF conversion returned no result",
        "runner.mix_prior_pdf": "Batch contains both PDF and Excel; PDF was used (ignored {count} Excel file(s)).",
        "runner.mix_prior_excel": "Batch contains both PDF and Excel; this profile supports Excel only (ignored {count} PDF file(s)).",
        "runner.engine_no_convert": "Engine module is missing convert(): {module}",
        "runner.region_required": "region is required when template_path is not specified",
        "runner.pn_meta_must_object": "pn_meta must be a JSON object",
        "runner.convert_mapping_must_object": "convert_mapping must be a JSON object",
        "runner.employee_directory_must_array": "employee_directory must be a JSON array",
        # --- post_checks ---
        "post_checks.empty_result": "PDF parse result is empty; source sheet write aborted",
        "post_checks.no_employees": "PDF parse found no employees (profile={profile}); layout may have changed. Aborted to avoid empty/wrong sheet.",
        "post_checks.fatal_warnings": "PDF critical field parse or reconciliation failed (profile={profile}); aborted to avoid silent wrong amounts. Details: {preview}{more}",
        "post_checks.fatal_warnings_more": " ({count} total)",
        "post_checks.placeholder_names": "PDF has no real employee names (all placeholders, profile={profile}); layout may have changed. Aborted.",
        "post_checks.topsource_no_labor_usd": "TopSource PDF #{index} has no labor cost USD parsed; layout may have changed. Aborted.",
        # --- fatal warnings ---
        "warn.no_employee_name": "No employee name parsed",
        "warn.no_gross_salary": "No Gross Salary parsed",
        "warn.monthly_salary_block_missing": "No match for Monthly Salary / Empr Taxes / Empr Contributions block; layout may have changed",
        "warn.labor_net_mismatch": "Pay components total {labor} does not match line net {net}",
        "warn.service_fee_mismatch": "Line net+ServiceFee={expect} does not match Invoice Total {total}",
        "warn.no_labor_usd": "No labor cost USD bundle amount parsed",
        "warn.no_quarter_salary_pkr": "No quarterly salary PKR parsed",
        "warn.no_quarter_period": "No quarterly period parsed",
        "warn.no_federal_tax": "Federal IT / Sindh Sales Tax amounts not parsed from PDF",
        "warn.cgst_only": "Only CGST parsed; Business Tax estimated as ×2",
        "warn.no_gst": "No GST parsed; Business Tax set to 0",
        "warn.payroll_name_unmatched": "{name}: no matching employee in Payroll; EE deduction columns left empty",
        "warn.unrecognized_pdf_type": "Cannot identify PDF type, trying content parse: {name}",
        # --- eor_uk ---
        "eor_uk.excel_single_only": "eor_uk Excel currently supports only one single-employee vertical sheet (received {count}). Please upload one at a time.",
        "eor_uk.merge_pdf_one": "When mixing files for eor_uk, only 1 PDF is allowed (received {count})",
        "eor_uk.merge_excel_one": "When mixing files for eor_uk, only 1 Excel is allowed (received {count})",
        "eor_uk.pdf_no_batch": "eor_uk PDF batch is not supported yet ({count} files). Please upload 1 at a time.",
        "eor_uk.excel_missing": "No Excel provided",
        "eor_uk.excel_not_found": "Excel not found: {path}",
        "eor_uk.uk_l_multi": "Multiple UK-L Excel files cannot be merged automatically. Please upload only one.",
        "eor_uk.uk_l_passthrough": "Source is already UK-L; used as convert input as-is",
        "eor_uk.no_source": "No usable PDF or Excel source provided",
        "eor_uk.template_missing": "UK template not found: {path}",
        "eor_uk.no_employee_name": "EOR UK PDF: no employee name parsed; layout may have changed",
        "eor_uk.no_salary_fields": "EOR UK PDF: Gross Salary / Empr Taxes / Empr Contributions not parsed; layout may have changed",
        "eor_uk.labor_mismatch": "EOR UK PDF: pay components total {labor} does not match line net {net}; aborted",
        "eor_uk.service_fee_mismatch": "EOR UK PDF: line net+ServiceFee={expect} does not match Invoice Total {total}; aborted",
        "eor_uk.excel_no_gross": "EOR UK Excel: Gross Salary (Gross Pay) not parsed; check column mapping or layout",
        "eor_uk.merge_no_gross": "EOR UK merge: Gross Salary missing; check PDF/Excel or column mapping",
        # --- topsource_uk ---
        "topsource_uk.no_employee_name": "TopSource PDF: no employee name parsed; layout may have changed",
        "topsource_uk.no_labor_usd": "TopSource PDF: labor cost USD not parsed; layout may have changed",
        "topsource_uk.no_source": "No source files provided",
        "topsource_uk.uk_l_multi": "Multiple UK-L Excel files cannot be merged. Upload one, or upload vendor original invoices.",
        "topsource_uk.uk_l_passthrough": "Source is already UK-L; used as convert input as-is",
        # --- auxilium_uae ---
        "auxilium_uae.not_recognized": "Not recognized as Auxilium/UAE Payroll Draft (no employee/payroll headers in first {scan_max} rows). Sheet \"{sheet}\" sample: {sample}",
        "auxilium_uae.no_id_name_columns": "Employee ID / name columns not found. Map to Emp ID / Employee Name in column mapping, or match header names.",
        "auxilium_uae.uae_l_multi": "Multiple UAE-L Excel files cannot be merged automatically. Please upload only one.",
        "auxilium_uae.no_mix_uae_l_draft": "Do not mix ready UAE-L with Payroll Draft in one batch",
        "auxilium_uae.uae_l_passthrough": "Source is already UAE-L; used as convert input as-is",
        "auxilium_uae.excel_only_main": "auxilium_uae main source supports Excel Payroll Draft only; unrecognized PDF: {names}. Use INV- numbers on two invoices: higher = prior Admin Fee, lower = current.",
        "auxilium_uae.need_draft_excel": "Upload at least one Auxilium Payroll Draft Excel (main payroll source). Two Invoice PDFs: higher invoice number = prior Admin Fee, lower = current.",
        "auxilium_uae.no_pdf_support": "auxilium_uae does not support PDF yet; upload Payroll Draft Excel",
        # --- biz_solutions_india ---
        "biz_solutions_india.no_employee_name": "Could not parse employee name: {name}",
        "biz_solutions_india.no_ctc": "Could not parse Monthly CTC: {name}",
        "biz_solutions_india.cgst_only_fatal": "Biz Solutions PDF: only CGST parsed, SGST missing; layout may have changed. Aborted to avoid wrong tax.",
        "biz_solutions_india.no_gst_fatal": "Biz Solutions PDF: GST not parsed; layout may have changed. Aborted to avoid zero tax.",
        "biz_solutions_india.single_pdf_only": "biz_solutions_india supports only one PDF at a time (one employee per invoice)",
        "biz_solutions_india.need_tax_invoice_pdf": "Please upload Biz Solutions Tax Invoice PDF",
        # --- at_technical_cyprus ---
        "at_technical.invoice_no_blocks": "A&T Invoice: no employee blocks parsed: {name} — {detail}",
        "at_technical.payroll_no_employees": "A&T Payroll: no employees parsed: {name}",
        "at_technical.no_parse_results": "No usable Invoice / Payroll PDF parse results",
        "at_technical.payroll_name_mismatch": "A&T Cyprus: employee \"{name}\" not matched in Payroll; layout or name mismatch. Aborted to avoid empty deductions.",
        "at_technical.need_pdf_pair": "at_technical_cyprus requires PDF (Invoice + Payroll Calculation)",
        "at_technical.no_parsed_pdfs": "Could not parse any A&T Invoice / Payroll PDF",
        # --- connect_uae ---
        "connect_uae.not_invoice": "Not a Connect Resources tax invoice: {name}",
        "connect_uae.split_mismatch": "Connect UAE: employee \"{name}\" mapping split total {split_sum} ≠ PDF monthly pay {payroll}; aborted",
        "connect_uae.need_tax_invoice_pdf": "connect_uae: please upload Connect Tax Invoice PDF",
        # --- safeguard_italy ---
        "safeguard_italy.no_header_row": "SafeGuard employee header row not found (Employee ID / Employee Name)",
        "safeguard_italy.italy_l_multi": "Multiple Italy-L Excel files cannot be merged automatically. Please upload only one.",
        "safeguard_italy.no_mix": "Do not mix ready Italy-L with SafeGuard source bills",
        "safeguard_italy.italy_l_passthrough": "Source is already Italy-L; used as convert input as-is",
        "safeguard_italy.excel_only": "safeguard_italy supports Excel source bills only; upload SGWI Payroll xlsx",
        "safeguard_italy.no_pdf_runtime": "safeguard_italy main source is Excel; PDF not supported. Upload SGWI Payroll xlsx",
        # --- panda_work_pk ---
        "panda_work_pk.no_federal_tax": "Panda Work PDF: Federal IT / Sindh Sales Tax not parsed; layout may have changed. Aborted.",
        "panda_work_pk.no_employee_name": "Panda Work PDF: no employee name parsed; layout may have changed",
        "panda_work_pk.no_salary_pkr": "Panda Work PDF: quarterly salary PKR not parsed; layout may have changed",
        "panda_work_pk.no_quarter_period": "Panda Work PDF: quarterly period not parsed; layout may have changed",
        "panda_work_pk.not_invoice": "Not a Panda Work Global Pakistan invoice: {name}",
        # --- engines / region ---
        "engines.unknown": "Unknown convert engine \"{engine_id}\"; known: {known}",
        "region.unknown": "Unknown region \"{region}\"; known: {known}",
        # --- api ---
        "api.pdf_only": "Only .pdf is supported, got: {suffix}",
        "api.excel_only": "Only Excel (.xlsx/.xlsm) is supported, got: {suffix}",
        "api.excel_only_short": "Only Excel is supported, got: {suffix}",
        "api.empty_upload": "Uploaded file is empty",
        "api.empty_upload_named": "Uploaded file is empty: {name}",
        "api.need_pdf": "Please upload at least one PDF",
        "api.need_source": "Please upload at least one source file",
        "api.pdf_excel_only": "Only .pdf / .xlsx / .xlsm are supported, got: {name}",
        "api.pdf_excel_xlsx_only": "Only .xlsx/.xlsm/.pdf are supported, got: {suffix}",
        "api.template_xlsx": "Template must be .xlsx/.xlsm",
        "api.template_xlsx_with_suffix": "Template must be .xlsx/.xlsm, got: {suffix}",
        "api.convert_fail": "Convert failed: {detail}",
        "api.vendor_batch_fail": "Vendor source batch convert failed: {detail}",
        "api.pdf_convert_fail": "PDF convert failed: {detail}",
        "api.pdf_batch_fail": "PDF batch convert failed: {detail}",
        "api.blocked_file_type": "blocked file type: {suffix}",
        "api.engine_id_required": "engineId required",
        "api.pn_meta_invalid": "Invalid pn_meta: {detail}",
        "api.convert_mapping_invalid": "Invalid convert_mapping: {detail}",
        "api.employee_directory_invalid": "Invalid employee_directory: {detail}",
        "api.bypass_classify_fail": "Bypass file classification failed: {detail}",
        "api.region_template_not_found": "Region template not found: {path}",
        "api.decrypt_fail": "Decrypt failed: {detail}",
        "api.no_output_file": "Conversion finished but no output file was produced",
        "api.excel_snapshot_fail": "Excel snapshot failed: {detail}",
        "api.hf_snapshot_fail": "HF snapshot failed: {detail}",
        "api.file_role_fail": "File role detection failed: {detail}",
        "api.region_required_no_template": "region is required when no template is uploaded",
        # --- mapping inspect ---
        "inspect.pdf_labels_unsupported": "Engine \"{engine_id}\" / profile \"{profile_id}\" does not support PDF label inspection",
        "inspect.header_unsupported": "Engine \"{engine_id}\" does not support header inspection",
        "inspect.sheet_not_found": "Worksheet \"{sheet}\" not found",
        "inspect.cannot_open_source": "Cannot open source workbook: {detail}",
        "inspect.calculation_header_fail": "Found Calculation sheet but failed to parse headers: {detail}",
        "inspect.no_calculation_or_italy_l": "Calculation or Italy-L worksheet not found",
        "inspect.template_not_found": "Template file not found: {template_path}",
        "inspect.template_sheet_not_found": "Sheet \"{sheet}\" not found in template",
        "inspect.no_sheets": "Workbook has no worksheets",
        "inspect.uae_l_or_draft_missing": "Worksheet \"UAE-L\" not found and vendor Payroll Draft headers not recognized. Details: {detail}",
        "inspect.uk_vertical_no_labels": "Vertical sheet labels not recognized (sheet={sheet}, labelCol={label_col}, amountCol={amount_col}). Confirm UK-L or Analysis of Payroll Totals.",
        "inspect.uae_vendor_header_hint": "Vendor source headers recognized. Map vendor columns → UAE-L in Column Mapping.",
        "inspect.safeguard_calculation_hint": "SafeGuard Calculation headers recognized. Configure Column Mapping and click Save; unsaved mappings are not used in conversion.",
        # --- api (inspect/bypass) ---
        "api.no_matching_plugin": "No matching plugin",
        "api.not_bypass_file": "Not recognized as a bypass file",
        # --- ee_code (shared) ---
        "ee_code.no_name": "Source sheet has no employee name; cannot match EE Code",
        "ee_code.no_directory": "No client employee directory provided; cannot match EE Code",
        "ee_code.not_matched": "No EE Code matched: {label}",
        "ee_code.ambiguous": "Ambiguous EE Code match ({count} tied): {label}",
        "common.empty_name": "(empty)",
        "common.write_uk_fx_d24_fail": "Failed to write UK-L!D24 FX rate: {detail}",
        "common.write_pn_fx_fail": "Failed to write PN FX rate: {detail}",
        "common.fx_online_fallback_d24": "Online FX failed; kept source D24={rate}: {detail}",
        # --- engine (region convert) ---
        "engine.person_row": "Person #{index}: {detail}",
        "engine.ee_sheet_row": "{sheet} EE person #{index}: {detail}",
        "engine.formula_style_count": "Employee formula style mappings: {count}",
        "engine.formula_pair": "Formula pairing: person #{index} → {main_sheet} row {main_row} / {ee_sheet} EE row {ee_row}",
        "engine.fx_from_bill": "FX taken from vendor bill {source} = {rate}",
        # --- template_rows (UI formula-row picker) ---
        "template.row_default": "Row {row} (default)",
        "template.row_plain": "Row {row}",
        "template.row_with_marker": "Row {row} · {marker}",
        # --- convert_checks ---
        "checks.header_fallback": "Header \"{field}\" ({names}) not found; fell back to fixed column {fallback}. Update mapping/headers if vendor layout changed; do not rely on column numbers.",
        "checks.column_missing_skip": "Column \"{field}\" ({names}) not found; field skipped",
        "checks.column_rename_no_hit": "columnRename is configured but matched no source headers. Check vendor column names in mapping{examples_suffix}",
        "checks.column_rename_examples_suffix": " (sample misses: {examples})",
        "checks.column_rename_partial_miss": "columnRename partially missed ({count} items), e.g.: {examples}",
        "checks.column_rename_idle": "columnRename unused this month (source already uses target names; {idle_aliases} alias(es) idle)",
        "checks.result_empty": "Convert result is empty; please verify manually",
        "checks.zero_employees": "Convert result has 0 employees; check name column / header row in source",
        "checks.too_many_employees": "Convert result has unusually many employees ({count}); check for header/blank rows",
        "checks.many_fallback_cols": "{count} field(s) written via fixed-column fallback; vendor layout may have changed—verify by headers",
        # --- eor_uk (informational) ---
        "eor_uk.not_eor_keyword": "EOR Services keyword not found in body; may not be an invoice for this profile",
        "eor_uk.vertical_col_detect": "Vertical columns detected: labelCol={label_col} amountCol={amount_col}{suffix}",
        "eor_uk.vertical_col_rename_suffix": "; columnRename={count} entries",
        "eor_uk.vertical_col_default_suffix": "; using built-in default aliases",
        "eor_uk.unmapped_labels": "Unmapped labels (add in Office Column Mapping): {labels}",
        "eor_uk.excel_no_name_hint": "Vertical Excel usually has no employee name; UK-L title uses Employee—complete in PN metadata",
        "eor_uk.er_nic_pension_incomplete": "ER NIC / ER Pension not fully parsed; check column mapping (Employer NI / Employer Pension)",
        "eor_uk.ee_side_zeroed": "EE PAYE / EE NIC / EE Pension not on invoice; set to 0 (manual/other source needed)",
        "eor_uk.service_fee_not_written": "PDF Service Fee={service_fee} (different from PN Management Fee / Recurring Fee; not auto-written to UK!H)",
        "eor_uk.pdf_excel_conflict": "PDF/Excel \"{label}\" mismatch: PDF={pdf_val} Excel={excel_val}; Excel used",
        "eor_uk.merge_note": "Merged PDF (name/invoice/Service Fee) with Excel (payroll incl. EE); amount conflicts use Excel",
        "eor_uk.merge_no_name": "Still no employee name after merge; UK-L title uses Employee—complete in PN metadata",
        # --- topsource_uk (informational) ---
        "topsource_uk.usd_bundle_manual": "TopSource PDF is USD bundle only, no GBP detail; fill Gross/Holiday/PAYE etc. manually from screenshot",
        "topsource_uk.labor_usd_note": "Invoice manual USD bundle={amount} (for verification; not written to UK-L detail)",
        "topsource_uk.service_charge_note": "Invoice Service Charge USD={amount} (differs from PN Management Fee formula; not auto-written)",
        "topsource_uk.not_topsource_keyword": "TopSource keyword not found in body; may not be an invoice for this profile",
        "topsource_uk.excel_not_topsource": "TopSource not found in header; may not be this Excel layout",
        "topsource_uk.ts_margin_note": "Invoice has TS Margin/service fee GBP={amount} (differs from PN Management Fee; not auto-written)",
        "topsource_uk.name_fallback": "Name not parsed from A3; used sheet/file name as fallback",
        "topsource_uk.no_gbp_amounts": "No GBP detail amounts parsed from Excel",
        "topsource_uk.mix_prior_pdf": "Batch has PDF and Excel; TopSource main source used {pdf_count} PDF(s), ignored Excel: {names}",
        # --- auxilium_uae (informational) ---
        "auxilium_uae.maybe_not_draft": "File may not be Auxilium Payroll Draft; still attempting parse: {name}",
        "auxilium_uae.admin_fee_no_facts": "Admin Fee PDF recognized but no facts parsed",
        "auxilium_uae.admin_fee_parsed_dual": "Admin Fee parsed: prior {prev_no} VAT={prev_vat}; current {curr_no} VAT={curr_vat}",
        "auxilium_uae.admin_fee_parsed_single": "Admin Fee PDF Total VAT={vat} ({source_file})",
        # --- safeguard_italy (informational) ---
        "safeguard_italy.column_rename_empty": "columnRename empty: same-name matching only. Configure mapping and click Save Mapping.",
        "safeguard_italy.column_rename_loaded": "columnRename loaded: {count} mapping(s)",
        "safeguard_italy.maybe_not_safeguard": "File may not be SafeGuard Italy bill; still attempting parse: {name}",
        "safeguard_italy.italy_l_reparse": "Source has Italy-L with column mapping; re-parsed with mapping: {name}",
        "safeguard_italy.column_rename_no_hit": "columnRename configured but matched no source columns; check vendor names against sample headers",
        "safeguard_italy.column_rename_applied": "columnRename applied {count} write(s) this batch (by source column)",
        "safeguard_italy.ignored_non_excel": "safeguard_italy main source is Excel; ignored: {names}",
        # --- connect_uae (informational) ---
        "connect_uae.expense_split": "Reimbursement \"{names}\" total {total} not split by person; allocated {each} each—please verify",
        "connect_uae.split_all_basic": "{name}: Basic/Housing/Transport mapping missing; full PDF monthly pay {payroll} put in Basic",
        "connect_uae.split_sum_mismatch_warn": "{name}: mapping split total {split_sum} ≠ PDF monthly pay {payroll}",
        "connect_uae.batch_first_pdf": "Batch has {count} PDF(s); used {name}",
        # --- at_technical (informational) ---
        "at_technical.invoice_missing_labels": "Invoice \"{name}\" missing vendor labels {labels}; skipped",
        "at_technical.invoice_no_base_salary": "Invoice \"{name}\": no Base salary after mapping (confirm Gross Salary → Base salary); amounts kept in mapped columns",
        "at_technical.no_monthly_cost_rows": "No Monthly cost employee rows matched",
        "at_technical.invoice_blocks_skipped": "Found {heads_seen} employee block(s) but all skipped (likely missing PDF labels, not mapping): {detail}",
        "at_technical.inspect.labels_hint": "Extracted {hit_total} label hit(s) from vendor PDF ({uniq_total} unique); map vendor labels → Cyprus-L in Column Mapping{medical_note}",
        "at_technical.inspect.no_labels": "No PDF labels recognized",
        "at_technical.inspect.medical_note": "; Medical current + back-charge summed into Medical Insurance (e.g. {name} Medical Insurance={amount})",
        "at_technical.payroll_name_unparsed": "Could not parse name in a Payroll employee block",
        "at_technical.liability_keys_differ": "{name}: full and truncated Liability keys both have values ({full} / {truncated}); summed into Public Liabilit column",
        "at_technical.liability_truncated": "{name}: amount under truncated header \"{typo}\"; writing to Public Liabilit column",
        "at_technical.invoice_only": "Invoice only: missing EE Social Ins / Tax / N.H.S.; columns left empty",
        "at_technical.payroll_only": "Payroll only: missing Public Liability; column set to 0",
        "at_technical.er_contrib_mismatch": "{name}: Invoice ER Contributions {inv_er} ≠ Payroll {pay_er}; using Invoice",
        "at_technical.payroll_only_employee": "{name}: only in Payroll; appended (Public Liability=0)",
        "at_technical.duplicate_invoice": "Duplicate Invoice ignored: {name}",
        "at_technical.duplicate_payroll": "Duplicate Payroll ignored: {name}",
        "at_technical.suggest_both_pdfs": "Upload Invoice and Payroll Calculation together; one missing—merged as best effort",
        # --- biz_solutions_india (informational) ---
        "biz_solutions_india.no_period": "Period not parsed; India-L period will be empty",
        "biz_solutions_india.ignored_non_pdf": "biz_solutions_india main source is PDF; ignored: {names}",
        # --- panda_work_pk (informational) ---
        "panda_work_pk.no_quarter_salary_row": "{name}: no quarterly salary; Base Salary not written",
        "panda_work_pk.eobi_it_empty": "{name}: E.O.B.I / IT columns left empty if pakistanEmployeeFees mapping not configured",
        # --- vendor plugins ---
        "plugin.classify_fail": "Plugin {plugin_id} classify failed {name}: {detail}",
        "plugin.parse_fail": "Plugin {plugin_id} parse failed: {detail}",
        "plugin.write_fail": "Plugin {plugin_id} write failed: {detail}",
        "auxilium_plugin.admin_fee_split": "Batch Admin Fee split by invoice no.: prior (higher) {prev_no} VAT={prev_vat}; current (lower) {curr_no} VAT={curr_vat}",
        "auxilium_plugin.batch_two_invoices": "Batch has {count} invoices; only min and max invoice numbers used",
        "auxilium_plugin.no_latest_vat": "Auxilium: no latest Admin Fee Total VAT (upload Admin Fee invoice or batch with Draft); Business Tax skipped",
        "auxilium_plugin.invalid_vat": "Auxilium: Total VAT invalid ({currency}); Business Tax skipped",
        "auxilium_plugin.no_uae_sheet": "Auxilium: template missing UAE sheet; cannot write Business Tax",
        "auxilium_plugin.no_prior_vat": "Auxilium: no prior Admin VAT on books; Business Tax uses latest VAT—set opening Prior Total VAT in mapping",
        # --- pakistan ---
        "pakistan.fees_not_matched": "{name}: pakistanEmployeeFees mapping did not match E.O.B.I / IT",
        "pakistan.bt_no_tax_amounts": "{name}: Sindh Sales Tax / Federal IT not parsed from PDF; Business Tax not written",
        "pakistan.bt_column_missing": "Pakistan main sheet row {header_row}: Business Tax column not found; invoice-derived Business Tax skipped",
        "pakistan.bt_missing_coeff": "{name}: invoice-derived Business Tax enabled but BT coefficients missing (use source/PDF with derived columns)",
        "pakistan.default_person": "Person #{index}",
        # --- china ---
        "china.no_directory_row": "Person #{index}: no employee directory; China!B library name not written (name {name})",
        "china.ee_code_fail": "Person #{index}: {detail}; China!B left empty",
        "china.name_not_matched": "Person #{index}: name {name} not matched in directory; China!B left empty",
        "china.directory_name_empty": "Person #{index}: name {name} matched directory but library name empty; China!B left empty",
        "china.fx_fallback": "FX not read from vendor bill (S-Payment Notice!C49 or 汇率 label); fell back to online CNY—ensure payment notice has numeric FX (not formula-only)",
        "china.formula_miss": "Formula pairing miss: mapping expects China example rows {rows} but all fell on default row; check directory name match (incl. pinyin)",
        "china.formula_no_fields": "Employee formula styles configured but chinaExampleRow/mainExampleRow missing (save may have failed)",
        # --- china_hrone ---
        "china_hrone.fx_not_read": "FX not read from vendor bill (S-Payment Notice!{cell} or 汇率 label); PN FX cell unchanged",
        "china_hrone.fx_nnroad_fail": "FX should be NNRoad month-1 * 0.97, but it was not fetched ({detail}); PN FX unchanged (vendor bill FX not used)",
        "china_hrone.no_pn_sheet": "Template has no PN sheet; FX read but not written",
        "china_hrone.no_fx_row": "PN FX rate row not found; FX read but not written",
        "china_hrone.no_names": "No employee names read from source; cannot match EE Code",
        "china_hrone.no_ee_sheet": "Template has no China EE sheet; EE Code not written",
        "china_hrone.no_directory": "No client employee directory; cannot match EE Code",
        "unlock.empty": "Source file is empty, not a valid Excel: {name}",
        "unlock.source_is_pdf": "Source file is actually a PDF, not Excel: {name}",
        "unlock.not_xlsx_header": "Source is not a valid xlsx (header {header}): {name}",
        "unlock.after_decrypt": "Source still cannot be opened after decrypt ({name}): {detail}",
        "unlock.template_not_xlsx": "Template is not a valid xlsx ({name}): {detail}",
        # --- tw ---
        "tw.fx_fallback": "Summary has no valid Exchange rate; fell back to API TWD FX",
        # --- uk ---
        "uk.gross_salary_empty": "{sheet} ({name}): Gross Salary empty/0; fill GBP detail manually from screenshot",
        # --- region pn expand ---
        "cyprus.pn_expand": "Cyprus PN multi-employee row expand provisional: expanded Cyprus/Cyprus EE ({count} people)—verify PN manually",
        "italy.pn_expand": "Italy PN multi-employee Labor/Service Fee expand provisional: expanded Italy/Italy EE ({count} people)—verify PN detail rows",
        "india.pn_expand": "India PN multi-employee detail expand provisional: expanded India/India EE ({count} people)—verify PN",
        "pakistan.default_employee": "Employee",
    },
}


def fatal_warning_patterns() -> list[re.Pattern[str]]:
    """Bilingual regex for warnings that must abort PDF ingest."""
    return [
        re.compile(r"未解析到员工姓名|No employee name parsed", re.I),
        re.compile(r"未解析到\s*Gross\s*Salary|No Gross Salary parsed", re.I),
        re.compile(
            r"未匹配到.*Monthly\s+Salary|No match for.*Monthly Salary|layout may have changed|版式可能已变更",
            re.I,
        ),
        re.compile(r"工资构成合计.*不一致|Pay components total.*does not match", re.I),
        re.compile(r"行净额\+ServiceFee.*不一致|Line net\+ServiceFee.*does not match", re.I),
        re.compile(r"未解析到人工成本\s*USD|No labor cost USD", re.I),
        re.compile(r"未解析到季度薪资\s*PKR|No quarterly salary PKR", re.I),
        re.compile(r"未解析到季度账期|No quarterly period parsed", re.I),
        re.compile(
            r"未从 PDF 解析到 Federal IT|Sindh Sales Tax|Federal IT.*Sindh Sales Tax",
            re.I,
        ),
        re.compile(r"未解析到\s*GST|No GST parsed", re.I),
        re.compile(r"仅解析到\s*CGST|Only CGST parsed", re.I),
        re.compile(r"Payroll 中未匹配到同名员工|no matching employee in Payroll", re.I),
        re.compile(r"无法识别 PDF 类型|Cannot identify PDF type", re.I),
    ]


def parse_accept_language(header: str | None) -> str:
    """返回 'zh' 或 'en'。"""
    if not header:
        return "zh"
    primary = header.split(",")[0].strip().lower()
    if primary.startswith("en"):
        return "en"
    return "zh"


def set_locale(lang: str) -> Token:
    normalized = "en" if str(lang or "").lower().startswith("en") else "zh"
    return _locale.set(normalized)


def reset_locale(token: Token) -> None:
    _locale.reset(token)


def get_locale() -> str:
    return _locale.get()


def t(key: str, **kwargs: Any) -> str:
    lang = _locale.get()
    catalog = _MESSAGES.get(lang) or _MESSAGES["zh"]
    template = catalog.get(key) or _MESSAGES["zh"].get(key) or key
    if kwargs:
        try:
            return template.format(**kwargs)
        except (KeyError, ValueError):
            return template
    return template


# --- outbound: 业务代码可继续抛中文，API 出口按 locale 译成英文 ---

_OUTBOUND_EXACT: dict[str, str] | None = None
_OUTBOUND_TEMPLATES: list[tuple[re.Pattern[str], str]] | None = None
_FMT_FIELD_RE = re.compile(r"\{([a-zA-Z_][a-zA-Z0-9_]*)([^}]*)\}")


def _compile_outbound_tables() -> None:
    """用 zh/en 同 key 文案建：精确表 + 带占位符的正则表。"""
    global _OUTBOUND_EXACT, _OUTBOUND_TEMPLATES
    exact: dict[str, str] = {}
    templates: list[tuple[re.Pattern[str], str]] = []
    zh_cat = _MESSAGES["zh"]
    en_cat = _MESSAGES["en"]
    for key, zh in zh_cat.items():
        en = en_cat.get(key)
        if not en or not zh or zh == en:
            continue
        if "{" not in zh:
            exact[zh] = en
            continue
        # "{name}" / "{count}" → named groups；字面量做 regex escape
        parts: list[str] = []
        pos = 0
        fields: list[str] = []
        for m in _FMT_FIELD_RE.finditer(zh):
            parts.append(re.escape(zh[pos : m.start()]))
            name = m.group(1)
            fields.append(name)
            parts.append(f"(?P<{name}>.*?)")
            pos = m.end()
        parts.append(re.escape(zh[pos:]))
        try:
            pat = re.compile("^" + "".join(parts) + "$", re.S)
            templates.append((pat, en))
        except re.error:
            continue
    # 长模板优先，减少短串误伤
    templates.sort(key=lambda x: len(x[0].pattern), reverse=True)
    _OUTBOUND_EXACT = exact
    _OUTBOUND_TEMPLATES = templates


def translate_outbound(text: Any, *, target: str | None = None) -> Any:
    """把中文用户可见文案译成目标语言（默认当前 locale）。已是英文则原样返回。"""
    if not isinstance(text, str):
        if isinstance(text, list):
            return [translate_outbound(x, target=target) for x in text]
        return text
    lang = target or get_locale()
    if lang != "en" or not text:
        return text
    if _OUTBOUND_EXACT is None or _OUTBOUND_TEMPLATES is None:
        _compile_outbound_tables()
    assert _OUTBOUND_EXACT is not None and _OUTBOUND_TEMPLATES is not None
    hit = _OUTBOUND_EXACT.get(text)
    if hit is not None:
        return hit
    for pat, en_tmpl in _OUTBOUND_TEMPLATES:
        m = pat.match(text)
        if not m:
            continue
        try:
            return en_tmpl.format(**m.groupdict())
        except (KeyError, ValueError):
            return en_tmpl
    # 常见包装：\"plugin: 中文\" / \"前缀：中文\"
    for sep in (": ", "：", " - "):
        if sep in text:
            head, tail = text.split(sep, 1)
            translated_tail = translate_outbound(tail, target=target)
            if translated_tail != tail:
                return f"{head}{sep}{translated_tail}"
    return text


def translate_outbound_tree(value: Any, *, target: str | None = None) -> Any:
    """递归翻译 dict/list 里常见用户可见字段。"""
    lang = target or get_locale()
    if lang != "en":
        return value
    if isinstance(value, str):
        return translate_outbound(value, target=lang)
    if isinstance(value, list):
        return [translate_outbound_tree(v, target=lang) for v in value]
    if isinstance(value, dict):
        out: dict[str, Any] = {}
        for k, v in value.items():
            if k in ("msg", "detail", "message", "warnings", "error", "hint") or (
                isinstance(v, str) and any("\u4e00" <= ch <= "\u9fff" for ch in v)
            ):
                out[k] = translate_outbound_tree(v, target=lang)
            elif isinstance(v, (list, dict)):
                out[k] = translate_outbound_tree(v, target=lang)
            else:
                out[k] = v
        return out
    return value
