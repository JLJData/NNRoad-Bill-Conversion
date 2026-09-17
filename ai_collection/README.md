# 通用 AI 采集骨架

该模块已经提供 Provider 接口、注册表、独立采集请求构造、AI 结构化结果校验和离线 Mock Provider。目前没有配置或实现任何外部模型，不会发起网络请求。

## 独立性边界

`build_collection_request` 只接受：运行 ID、原始文件引用及哈希、Collection Schema 的语义字段、采集说明、账期和币种。它会从模型 Schema 中排除比较规则，并拒绝 Code Result、转换结果文件、预期值及差异报告等输入键。

`sourceRef` 是服务端不透明引用。未来真实 Provider 适配器只能把它解析为本次已授权的原始供应商文件，不能解析成代码生成的中间 Excel 或最终 PN。仅靠字符串格式无法证明引用指向什么，Office 编排层必须执行文件角色、客户归属和内容哈希检查。

Provider 返回的每个字段须包含原始文件 ID和位置；金额为十进制字符串，缺失、无法辨认和不适用须明确区分。未知项目单独返回，不修改 Mapping。结构验证成功不代表金额正确。

## 当前能力

- `AIProvider`：所有真实／本地模型适配器的统一接口。
- `ProviderRegistry`：按全局 provider ID 注册，拒绝重复与未知 Provider。
- `MockAIProvider`：使用固定结果或回调运行测试，不读文件、不联网。
- `run_collection`：校验独立请求、调用 Provider、校验运行 ID、Provider 身份、Schema、证据和结构化输出。
- `AIProviderError`：超时及模型／传输失败，与 Code-vs-AI 业务差异分开处理。

## 尚未实现

- 真实 Provider、API Key、网络重试、并发控制和模型计费。
- PDF 图像／OCR 预处理。
- 持久任务、审计表、Office API 和前端。
- 员工目录的确定性关联。本轮比较器只采用唯一的规范化姓名精确匹配。

真实 Provider 开发前需要确认模型、密钥托管方式及真实账单是否允许发往该服务。
