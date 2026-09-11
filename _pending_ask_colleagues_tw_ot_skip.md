# 待问同事：台湾引擎 OT 明细列是否继续 skip

**日期：** 2026-09-09  
**相关代码：** `profiles/tw_payroll_calc/convert.py` → `SKIP_SOURCE_HEADERS`

## 现状

引擎默认跳过以下源列（不写入 TW-L）：

- `Service Fee`（已还原进 skip；PN/TW 多由公式算）
- `Hr (1.34x)` / `OT Payment (1.34x)`
- `Hr (1.67x)` / `OT Payment (1.67x)`
- `Hr (1x)` / `OT Payment (1x)`
- `Hr (2.67x)` / `OT Payment (2.67x)`

加班汇总列 `加班費 / Overtime Payment` **始终会写**。

## 已核实影响（Coral Sea 2026）

- TW / TW EE 的 OT 公式只引用 `TW-L!AC`（加班費），**不引用**明细列 `AD–AK`
- 取消 skip **不会改 PN / 账单金额**
- 取消后会把供应商拆分工时/金额写入 TW-L；手工历史 PN 多数只填加班費、明细留空或 0
- 有加班月份：加班費 ≈ 各档 `OT Payment` 之和

## 待同事拍板

1. TW-L 是否需要保留加班拆分明细（跟供应商表一致）？
2. 还是继续只写汇总「加班費」、明细列留空（对齐现有手工习惯）？

- 选 1 → 可删引擎 `SKIP_SOURCE_HEADERS`（或清空）
- 选 2 → 保持现状；个别客户仍可用 Office mapping `skipSourceHeaders` 覆盖

## 备注

- `Service Fee` 已还原为引擎默认 skip（不写入 TW-L）
- 香港引擎无同类硬编码；HK 的 skip 走 mapping（如 Medical Insurance Allowance）
