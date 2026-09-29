# 通用 AI 采集与母版副本

该模块提供 Provider 接口、OpenAI Responses API 适配器、结构化结果校验，以及最后一个 `-L` 工作表的动态检查和安全写入。当前 AI 工作簿链路把“原始账单 + 当前最后一个 `-L` 页”一起交给模型，由当前母版定义需要读取和填写的字段，不依赖固定字段清单或历史行列坐标。仅实例化或运行测试不会发起网络请求；只有业务入口显式调用 `OpenAIResponsesProvider` 且转换服务环境已配置 `OPENAI_API_KEY` 时才会调用外部模型。

## 独立性边界

动态母版链路只接受：运行 ID、原始文件引用及哈希、当前母版结构、少量稳定业务说明、账期和币种。它不接收 Code Result、转换结果文件、预期值或差异报告。旧的 `build_collection_request` 保留给未来的确定性字段级比较，不参与当前 AI 工作簿生成。

`sourceRef` 是服务端不透明引用。未来真实 Provider 适配器只能把它解析为本次已授权的原始供应商文件，不能解析成代码生成的中间 Excel 或最终 PN。仅靠字符串格式无法证明引用指向什么，Office 编排层必须执行文件角色、客户归属和内容哈希检查。

Provider 返回的每个写入值须包含原始文件 ID、位置和原文；金额使用严格十进制字符串。无法确定的内容必须进入 issues 并留空，不修改 Mapping。结构验证成功不代表金额正确。

## 当前能力

- `AIProvider`：所有真实／本地模型适配器的统一接口。
- `ProviderRegistry`：按全局 provider ID 注册，拒绝重复与未知 Provider。
- `MockAIProvider`：使用固定结果或回调运行测试，不读文件、不联网。
- `OpenAIResponsesProvider`：默认使用 `gpt-5.6-luna`；调用前解析原始文件引用并核对 SHA-256，通过 Responses API 文件输入和 Structured Outputs 返回采集结果。当前已按引擎默认开放 `tw/uk/uae/pakistan/india/cyprus_payroll_calc` 的 template-driven AI 对比（Office 也可在 `portal_bill_ai_profile_config_version` 挂 ACTIVE 配置）。
- `run_collection`：校验独立请求、调用 Provider、校验运行 ID、Provider 身份、Schema、证据和结构化输出。
- `AIProviderError`：超时及模型／传输失败，与 Code-vs-AI 业务差异分开处理。
- `inspect_last_l_sheet`：每次按当前母版重新选择工作簿顺序中最后一个 `-L` 表，并动态识别表头行、各列标题、表名、哈希、非空单元格、公式和合并区域。
- `plan_dynamic_template_fill`：让模型同时读取原始账单和当前母版结构，生成带来源标签和原文证据的动态写入计划；不包含 Code Result，也不使用固定字段清单。Office 当前供应商×客户转换配置中的 `columnRename` 会作为供应商列→母版列提示随本次请求传入。
- `resolve_dynamic_template_fill_targets`：目标列由程序控制。优先按 `columnRename`，其次按当前母版列标题的唯一匹配，直接替换模型给出的临时列坐标；员工行仍来自账单和母版的动态识别。无法唯一确定的列才保留模型候选并进入严格校验。
- 对 Excel 原始账单，如果 AI 证据给出了精确来源格（如 `Payroll calculation!AR4`），程序会直接读取该格上方的真实源表头和值，覆盖 AI 自报的 `sourceLabel/rawText`，再解析目标列。中英混合标题会单独提取完整英文业务码，保留 `2G/EE/ER/1T` 等限定词。
- `inspect_source_employee_layout` + `filter_inconsistent_employee_source_writes`：程序侧识别源表姓名列，丢弃无名汇总/合计行写入；同一目标行不得混用不同源员工行；同一员工不得占多行；有金额无姓名的目标行在已识别到其他员工后会被丢弃。CODE/母版已预填的人名会与源表同名行硬对齐，错人金额整笔丢弃。
- `list_named_source_employees` / `seed_missing_employee_identity_writes`：请求前把源表花名册塞进提示；模型漏人时程序按源姓名格自动补 CN/EN 到空闲 `-L` 行，并在仍不完整时强制重试。
- **两步填表（推荐）**：Office 传入 CODE 正式结果后，`prepare_template_with_code_identities` 只把 CODE `-L` 的人名写入母版副本（不拷贝金额/公式）；母版自带公式保留，模型再按已锚定行去原始账单填写其余空白格。公式格与数值 0（视同空）都不由 AI 写入。
- Service Fee / 服务费：提示词与校验都会跳过，AI 对比不填该列。
- 动态计划首次校验不合格时会把原因交给模型完整重做一次，第二次仍不合格则不生成错误工作簿。
- `write_dynamic_ai_template_copy`：程序校验值、来源和目标后，只写母版副本中的空白目标格，拒绝覆盖公式、非空格、过期母版和已有输出文件。

## 尚未实现

- 转换 API 中的异步任务入口、传输层网络重试、并发控制、用量和模型计费记录。
- PDF 图像／OCR 预处理。
- AI 运行/调用/差异表的持久异步任务落库，以及逐字段 Code-vs-AI 差异面板。当前 Office 为人工触发并保存独立工作簿，核对页可在 CODE 与 AI 标签间切换。
- 员工目录的确定性关联。本轮比较器只采用唯一的规范化姓名精确匹配。

真实调用前仍需在转换服务进程配置 `OPENAI_API_KEY`，并确认具体原始账单允许发往 OpenAI。Office 不保存该密钥。

转换服务提供 `GET /ai-validation/profiles` 和 `POST /ai-validation/run`。运行接口默认由
`AI_VALIDATION_ENABLED=false` 关闭；启用后接收原始供应商文件和当前母版，返回独立的
`AI_comparison.xlsx`。响应头 `X-AI-Formal-Result=false`，现有 `/convert` 生成的 CODE 文件仍是正式结果。
模型默认值可由 `AI_VALIDATION_MODEL` 配置，Office 也可把全局模型配置随内部请求传入；API 密钥只由
Python 转换服务进程的 `OPENAI_API_KEY` 管理。
