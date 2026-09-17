# 只读代码结果采集器

已实现配置驱动的 Excel 结果读取、输出合同校验和 Code-vs-AI 确定性比较。此模块是 AI 核验的代码侧基础，不包含真实模型调用、HTTP 接口或审核拦截。

现有 `convert_api.py`、地区转换引擎及前端不受影响。依赖沿用项目中的 openpyxl，支持 Python 3.9+。

## 输入

1. 最终转换工作簿。
2. 标准字段 Schema；当前支持十进制金额字段。
3. 代码结果读取绑定，包括母版 SHA-256、结果 sheet 和预期表头。
4. 可信任务清单 manifest，明确该次转换的员工范围与预期身份。
5. 可选的可信重算快照，用于读取公式的已计算值。

试点配置位于 `docs/ai-validation/pilot-taiwan/`。它们仍为草案，离线使用必须明确指定 `--allow-draft`。

任务清单示例（哈希必须替换为实际值；姓名和实体编号为虚构示意）：

```json
{
  "resultSha256": "当前结果文件的64位小写SHA-256",
  "templateSha256": "本次实际使用母版的64位小写SHA-256",
  "period": "2026-03",
  "currency": "TWD",
  "employees": [
    {"row": 9, "entityKey": "employee-001", "expectedName": "Example Employee"}
  ]
}
```

manifest 应由后续 Office 转换任务集成产生。读取器会检查哈希、行范围、唯一键与实际姓名，但不能从一个成品 Excel 反向证明母版血缘，也不能证明外部提供的员工清单没有遗漏。不能从 AI 结果生成该清单或把清单交给 AI 作为匹配答案。

## 公式处理

本模块不执行公式，也不读取 XLSX 中可能过期的公式缓存。公式格没有对应的可信重算值时，输出 `unreadable` 和 `FORMULA_RESULT_UNAVAILABLE`。

未来计算服务提供的快照格式：

```json
{
  "status": "succeeded",
  "resultSha256": "被读取的最终结果文件的64位小写SHA-256",
  "engine": "实际计算服务名称及版本",
  "cells": [
    {"sheet": "TW", "cell": "B9", "type": "string", "value": "Example Employee"},
    {"sheet": "TW", "cell": "K9", "type": "number", "value": "1234.50"}
  ]
}
```

支持 number、string、blank、error；blank 的值必须为 null。公式错误或缺少所需单元格均需要复核。旧文件哈希、失败状态和重复单元格直接拒绝。

快照的成功标志是受信任计算服务的声明，并非签名证明。Office 接入时必须从受控服务生成，禁止信任浏览器或模型提交的快照。本轮尚未实现快照生产适配器；不要手工把已有缓存包装为“重算成功”。

## 使用

从转换仓库根目录运行：

```text
python -m bill_validation --workbook RESULT.xlsx --schema docs/ai-validation/pilot-taiwan/schema.v0.json --binding docs/ai-validation/pilot-taiwan/code-result-binding.v0.json --manifest MANIFEST.json --allow-draft --output COLLECTION.json
```

有成功重算快照时增加 `--snapshot SNAPSHOT.json`。不指定 `--output` 时输出到标准输出。输出文件不能覆盖任何输入，也不会覆盖已存在的输出。

Python 调用：`read_code_collection(workbook_path, schema, binding, manifest, calculation_snapshot=None, allow_draft=False)`。

## 输出与限制

- 输出员工业务键、实际读到的姓名、逐字段状态、十进制字符串、币种和文件哈希／sheet／单元格证据。
- 唯一表头决定字段位置，列号提示不作为读取依据；缺失或重复表头拒绝读取，避免列移动后静默错配。
- 空格和换行可规范化；不猜测不同标签是否同义，不模糊合并员工姓名。
- 缺失、零、负数、公式错误分别处理；货币符号、千位分隔符及非有限数值不会被静默转换。
- `ready_for_comparison` 仅表示本次清单中的数据可供后续比较，不代表金额正确、员工完整或 AI 核验通过。`automaticPassEnabled` 始终为 false。
- 特殊字段的业务比较方式保留在 Schema 中，本读取器不执行病假重算、容差判断、PN 汇总核验或任何财务规则。

## Code-vs-AI 比较

`compare_collections(code, ai, schema, documents)` 先分别校验代码与 AI 输出，再检查账期和币种。员工目前只按规范化后的姓名做唯一精确关联，不做模糊匹配；同名、缺失员工和额外员工都会进入差异报告。

普通金额字段按 Schema 的诊断容差比较。`review_recalculated_value` 等特殊模式即使值相同也要求复核。未知项目、不可读值及双方采集问题不会被吞掉。

离线命令：

```text
python -m bill_validation.compare_cli --code CODE.json --ai AI.json --schema SCHEMA.json --documents DOCUMENTS.json --output DIFFERENCES.json
```

documents 文件可直接是数组，也可为 `{ "documents": [...] }`。比较命令不修改账单和流程状态，不覆盖已有输出。退出码：0 = 字段匹配；2 = 需复核；3 = 确认存在不一致；1 = 输入或合同错误。

比较结果中的 `MATCH` 仅代表本次提供的数据按诊断规则一致。`workflowDecision` 固定为 `NOT_EVALUATED`，`automaticPassEnabled` 固定为 false，不能直接用来提交、审核或发布账单。

退出码：0 = 可比较；2 = 已生成结果但需复核；1 = 配置、输入、版本或文件错误。

## 测试

```text
python -m unittest discover -s tests -p "test_bill_validation.py" -v
```

测试使用临时合成工作簿，不修改现有账单。覆盖行列重排、重复与缺失表头、员工身份错误、缺失与零、负数、公式不可读、重算快照失效、错误公式、配置合同及输出覆盖保护。
