# 通用 AI 采集与母版副本

该模块提供 Provider 接口、OpenAI Responses API 适配器、结构化结果校验，以及最后一个 `-L` 工作表的动态检查和安全写入。当前 AI 工作簿链路把“原始账单 + 当前最后一个 `-L` 页”一起交给模型，由当前母版定义需要读取和填写的字段，不依赖固定字段清单或历史行列坐标。仅实例化或运行测试不会发起网络请求；只有业务入口显式调用 `OpenAIResponsesProvider` 且转换服务环境已配置 `OPENAI_API_KEY` 时才会调用外部模型。

## 默认流程：独立核对试验（planVersion 4）

Office 和 `run_ai_comparison_workbook` 现在默认使用 `comparisonMode=independent`。旧模式只有显式设置 `comparisonMode=legacy` 才能进入，保留用于回归和排查，不应把旧模式的一致率称为 AI 独立准确率。

- AI 接收原始账单、清理后的母版、账期/币种和特殊字段定义。不会收到 CODE 的员工名单、行顺序、普通金额、`columnRename`、程序整理的员工事实清单或差异报告。
- 原始 Excel 另附 `sourceWorkbookCells` 坐标视图：按原工作表与单元格地址原样列出内容及文件内已有公式缓存，不推断员工、不归类业务字段、不匹配目标列。它保留附件阅读时容易丢失的空行/空列位置；不提供 CODE 参考答案。该视图限 750,000 字符，超限明确要求拆分原始文件。
- AI 独立选择员工行和目标列，每格返回原始来源及 `reason`。程序检查合法位置、公式/固定内容保护、重复写入、来源 XLSX 单元格/PDF 页是否存在等。校验不改 plan；失败时原样带反馈重试一次。仍失败就结束，不输出半成品。
- 不再按源标签相似度换列，不替换 AI 数值，不从 CODE 补普通项目。错误、漏项及未知项目会保留在独立结果/差异中。
- 母版只清理可识别的样例员工行输入，保留公式、表头、备注和固定说明。布局不明确时使用目标表配置 `targetL.headerRow/dataStartRow/dataEndRow/nameColumns/protectedCells`，不把源表行号当作目标行号，也不使用 India 第 10 行等国家默认值。姓名列编号为 Excel 1-based。可用员工槽位以当前母版/显式范围为界，槽位不够必须扩充母版，程序不猜测扩展公式。
- **UK/EOR 竖表**（`targetL.layout=vertical_label_amount`，窄表也可自动识别）：按行标签清/填金额列，保留左侧字段名与公式；员工名写入标题格 `employeeNameCell`（如 `A3`），`rowFields` 列出可填金额格。特殊项按坐标从 CODE 覆盖，不做多人姓名重映射。
- 全体员工共有的 CODE 特殊列在 AI 输入中豁免。只属于部分员工的特殊项在独立提议完成后匹配；原始提议仍完整留存。员工区特殊项按唯一姓名及完整字段路径复制，重名、漏人或字段不唯一时不猜测，返回待确认事项。跨表公式也不猜依赖关系。非 `-L` 页不整页复制 CODE。
- 成功工作簿内的隐藏页 `_AI_Audit` 保存首次提议、校验反馈、重试、最终提议、哈希、响应 ID 及特殊项覆盖前后值；`_AI_CODE_Provenance` 保存特殊格最终坐标。服务端 `AI_VALIDATION_AUDIT_DIR`（默认 `output/ai_audit`，已忽略 Git）也持久保存失败与成功审计。
- 核对页面标明独立试验与 CODE 供值数，AI 表的蓝色标记指向实际复制后的坐标。特殊项及公式计算结果不计入普通独立输入差异统计，仍可显示真实差异。

### 真实账单评估

`python -m ai_collection.evaluate_independent --original bill.pdf --template master.xlsx --code reference.xlsx --period 2026-09 --currency AED --output output/ai_evaluation/run.xlsx`

可用 `--hints layout.json` 传目标表布局和特殊项坐标，`--env-file env/convert.env` 读取已有本地模型配置。该命令会真实调用模型，输出新工作簿及 `.evaluation.json`，分别比较首次/最终提议；不会修改原始文件、母版或 CODE。统计的是与已确认 CODE 参考的一致性，不是所有供应商客户的准确率保证。

### 边界

当前独立核对范围是最后一个 `-L` 页的员工输入。其他工作表沿用母版，不能当作 AI 独立生成的正式 PN。结构校验不验证 PDF 文字的真实性或薪资业务判断，不能替代人工审核。重名员工待人工确认，不自动按模糊姓名同步特殊金额。

## 旧接口与独立性边界

动态母版链路只接受：运行 ID、原始文件引用及哈希、当前母版结构、少量稳定业务说明、账期和币种。它不接收 Code Result、转换结果文件、预期值或差异报告。旧的 `build_collection_request` 保留给未来的确定性字段级比较，不参与当前 AI 工作簿生成。

`sourceRef` 是服务端不透明引用。未来真实 Provider 适配器只能把它解析为本次已授权的原始供应商文件，不能解析成代码生成的中间 Excel 或最终 PN。仅靠字符串格式无法证明引用指向什么，Office 编排层必须执行文件角色、客户归属和内容哈希检查。

Provider 返回的每个写入值须包含原始文件 ID、位置和原文；金额使用严格十进制字符串。无法确定的内容必须进入 issues 并留空，不修改 Mapping。结构验证成功不代表金额正确。

## 旧模式能力（仅 legacy）

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
- **旧两步填表**：Office 传入 CODE 正式结果后，`prepare_template_with_code_identities` 先清空员工数据区（`dataStartRow` 起）的母版样例字面值（公式保留），再把 CODE `-L` 人名写入；模型只填空白格。公式格与数值 0（视同空）都不由 AI 写入。员工行从 `targetL.dataStartRow`（或 profile 默认，如 India=10）起扫。落盘后再同步非 `-L` 页及员工区上方的元数据带（账期等）。
- **CODE 特殊来源格（蓝色 ⓘ）**：Office 把 `convert_source_snapshot.cellProvenance` 坐标作为 `codeProvenanceCells` 传入；AI 跳过这些格，覆盖检查也豁免这些格。程序在语义定位后、校验前剔除落到特殊坐标的 AI 写入，落盘后复制 CODE 的值或公式。员工区普通公式不会从 CODE 覆盖，以免掩盖 AI 差异；母版原有公式保留。`codeProvenanceCopyCount` 只计实际修改的特殊格，`codeOwnedCopyCount` 计全部同步修改，其余分类单独记录。特殊格超过 500 个时明确报错，不静默截断。
- 预填姓名后仍保留原始账单花名册、源行号和非零单元格事实（`detectedSourceEmployees.inputFacts`），供模型按姓名对应目标行；这些事实仅来自原始账单。逐员工检查最终保留的写入：原始账单存在可填入空白目标列的非零数值，但该员工没有任何此类写入时，重试一次，仍为空则报错，不返回仅含姓名的成功文件。零值、母版公式和服务费不触发此检查。
- 对源表头与目标表头唯一精确匹配、或由已配置列映射明确对应的非零数值，再逐字段检查遗漏，避免只填工资就放过 `Total` 等空列。禁止填写的是无名汇总行，不是员工自己的 `Total` 列；原始账单已给出的员工 Total 应原样采集，即使其中包含被跳过的 Service Fee，也不得扣减或重算。
- Service Fee / 服务费：提示词与校验都会跳过，AI 对比不填该列；也不得把跳过的费用金额挪到相邻费用列。
- 相近列互串：Excel 源列的父子表头若与母版 `pathLabel` 完全一致，程序与提示都把它当作最高优先目标列；否则再按完整标签与限定词匹配，模糊则留空。
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
