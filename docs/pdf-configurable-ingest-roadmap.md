# 供应商源文件可配置接入（后续要做）

> 记录日期：2026-08-07（同日补充 TopSource Excel / 竖表可配置）  
> 状态：构思 / 未开工  
> 原则：**老供应商冻结求稳，新供应商走可配置，多测再上线**

---

## 背景

- EOR SERVICES 等现有 **PDF** 转换：Python `pdf_ingest/profiles/*` **硬编码正则**，稳定可用。
- **TopSource** 实际有两条源：
  - **PDF**：一人一票，硬编码文本解析 → UK-L
  - **Excel**：一人一表发票（竖表 A 标签 / B 金额），硬编码 `parse_topsource_uk_excel` → UK-L；若已是 UK-L 则原样进引擎
- 映射 UI（`billConvertProfile`）的横向 `columnRename` / inspect（如 Auxilium）**不等于** TopSource 这种竖表抽取，也 **改不了** PDF 字段怎么抽。
- 配置轨要对齐的体验：上传样例 → 识别候选字段/标签 → 配置映射 → 试跑；**PDF 与「专用发票 Excel」都要能配**，但子类型不同。

---

## 目标

**新供应商**（PDF 或同类专用 Excel）：不改引擎代码，只配规则包即可接入。  
**老供应商**（`eor_uk`、现网 `topsource_uk` 硬编码路径等）：不动、不强制迁移。

非目标（至少一期不做）：
- 任意扫描件 / 纯图片 PDF 零配置万能解析
- 版式每次都变、毫无结构的账单
- 生产直接用 LLM 出账（最多辅助「建议映射」，人确认后落规则）
- 为统一架构强行改写已稳定的 EOR / TopSource 硬编码

---

## 策略：双轨并存

| 轨道 | 对象 | 行为 |
|------|------|------|
| **冻结轨** | 已上线 profile（`eor_uk`、现网 TopSource PDF/Excel 硬编码等） | 保持硬编码；只修明确 bug；不迁规则包 |
| **配置轨** | **新** 供应商（及可选：新接的同类 Excel/PDF） | 规则包 + 标准中间模型；多样例试跑、勾稽通过后再挂生产 |

识别优先级建议：
1. 配置显式指定 `pdfProfileId` / 规则包 ID（代传时已知供应商）
2. 冻结轨关键字命中 → 走老代码
3. 配置轨规则包命中 → 走解释器
4. 都未命中 → **失败并提示**，禁止误进老 profile

---

## TopSource 现状与配置化关系（专记）

| 来源 | 现状 | 配置化态度 |
|------|------|------------|
| TopSource **PDF** | `parse_topsource_uk_pdf` 硬编码 | **冻结**；新 PDF 供应商另走配置轨 |
| TopSource **Excel 发票** | `parse_topsource_uk_excel` 硬编码竖表 | **可配置，且优先于 PDF 配置化**（结构更稳、更好测）；老路径可先冻结，或做成「内置默认规则包」与现行为一致、允许覆盖 |
| 已是 **UK-L** 的 Excel | 原样拷贝进 `uk_payroll_calc` | 继续；无需抽取配置 |

> TopSource Excel **不是** Auxilium 那种「横表列名 + columnRename」。  
> 要对齐的是：**竖表「源标签 → 标准字段 / UK-L 标签」** 可配置；业务例外（如 TS Margin / 服务费解析到了但不自动写入）也要能在规则里声明。

---

## 架构分层

```
[供应商 PDF 或 专用发票 Excel]
    → ① 版式识别（冻结 profile 或 新规则包）
    → ② 字段抽取（文本/表格/竖表标签 → 标准中间模型）
    → ③ 标准中间模型（与供应商无关）
    → ④ 写入区域母版（UK-L / UAE-L / TW-L…）+ 现有正式转换 / 映射 UI
```

- ①②③：新供应商要配置化  
- ④：复用现有 Excel 映射、PN、公式、横表 columnRename  
- 老 profile：继续「源文件 → 直接写区域-L」也可，不必强行插入③  

配置轨按 **源类型** 分子能力（同一中间模型，不同抽取器）：

1. **PDF 文本/锚点**（EOR 类）  
2. **Excel 竖表标签**（TopSource 发票类）← 建议先做  
3. **Excel 横表列名**（Auxilium 已基本具备，可复用/收拢）  
4. **PDF 表格**（后期）

---

## 标准中间模型（配置轨专用，初稿方向）

与供应商无关的语义字段，例如：

- **单据头**：invoiceNo, invoiceDate, period, currency, vendorName  
- **员工行**：employeeName, employeeId, country, 分项金额…  
- **汇总**：subtotal, tax, total  
- **扩展**：`extra.*` 兜底  
- **写入策略**：某字段 `write: false`（如 TS Margin 仅告警不入库）

区域母版只认中间模型（或认映射后的标签），不认某个供应商的正则条文。

---

## 规则包内容（一个新供应商一份）

建议落在转换配置 / `convert_mapping`（或独立 `rulePack`），大致包括：

1. **识别**：关键字、页眉、域名、优先级；可强制指定；`sourceKinds: pdf | excel_vertical | excel_horizontal`  
2. **抽取策略**（可组合）：  
   - **Excel 竖表**：标签列/金额列、姓名标题行、账期格、标签→字段映射（对齐 TopSource）  
   - **Excel 横表**：表头行 + columnRename（对齐 Auxilium）  
   - **PDF**：标签锚定、正则、行切分、表格、区域块；bbox 二期  
3. **行模型**：一人一文件 / 一人一 sheet / 多行一表  
4. **校验**：必填、合计勾稽、人数上下限；失败指明哪条规则未命中  
5. **样例与回归**：绑样例文件 + 期望中间结果；改配置可一键试跑  

---

## 产品体验（映射页）

按源类型分段，不要糊成一张大表：

1. **源抽取配置（新）**  
   - PDF：上传样例 → 文本预览 → 绑标准字段 → 试跑  
   - Excel 竖表：上传样例 → 扫 A 列候选标签 → 下拉绑 UK-L/标准字段 → 试跑（**优先做**）  
   - Excel 横表：沿用现有 inspect + columnRename  
2. **区域母版映射（现有）**  
   中间字段 / 源表标签 → PN、公式、横表 rename 等  

运营心智：先教会「源文件里什么是什么」，再教「写到我们模板哪」。

---

## 建议分期

| 阶段 | 做什么 | 说明 |
|------|--------|------|
| **P0** | 配置轨解释器骨架 + 规则包格式；**不迁移** eor / 现网 topsource | 老路径零回归 |
| **P1a（优先）** | **Excel 竖表**标签映射可配 + UI 扫标签/试跑 | 覆盖 TopSource 类发票 Excel、同类新供应商；老 TS Excel 可默认规则包或继续冻结 |
| **P1b** | PDF 标签锚定 + 正则 + 行切分可配；UI 试跑 | 新 PDF 供应商；**多测**；不动 EOR/TS PDF 硬编码 |
| **P2** | 表格型 PDF；与横表 Excel 能力收拢文档化 | 接近完整「任意结构化源」 |
| **P3** | 规则包导入导出、样例回归集、变更审计/回滚 | 规模化 |
| **P4（可选）** | OCR / 模型辅助「提议规则」，人确认后落配置 | 降人工，仍非免配置 |

版式大改：复制规则包出 v2，避免在原包上无限打补丁。

---

## 验收与测试（新供应商上线前）

- [ ] 多样例（不同账期、人数、边界金额；Excel/PDF 按类型）试跑通过  
- [ ] 中间字段与人工核对一致  
- [ ] 写入区域母版后正式转换 / 金额勾稽通过  
- [ ] 业务例外字段（如不自动写入的费用）行为符合规则声明  
- [ ] 故意改文案/缺字段时，错误信息可读、可定位到规则  
- [ ] 确认不会命中冻结轨老 profile  
- [ ] 灰度：测试配置 / 代传验证后再挂生产  

---

## 明确不做什么（提醒）

- 不为「架构统一」去改写已稳定的 `eor_uk` / 现网 TopSource 硬编码（除非单独立项迁移且回归充分）  
- 不把 PDF/竖表 inspect 做成必须先动老供应商才能用  
- 不把 TopSource Excel 误当成横表 columnRename 去配  
- 生产不以模型直接出最终账单数字  

---

## 相关代码位置（现状，便于开工时对照）

- Python：`pdf_ingest/registry.py`、`pdf_ingest/profiles/eor_uk.py`、`pdf_ingest/profiles/topsource_uk.py`（含 `convert_pdfs` / `convert_excels`）、`pdf_ingest/runner.py`  
- 横表参考：`pdf_ingest/profiles/auxilium_uae.py`（吃 `columnRename`）  
- 映射 inspect：`mapping_inspect.py`（目前偏 Excel 横表/UK 竖表扫标签）  
- Office UI：`hrone-office-ui` → `billConvertProfile`  
- 后端：`hrone-office-abp` → `PortalBillConvertServiceImpl`（`pdfProfileId` / vendorToSource）  

---

## 一句话

**老的冻结；新的规则包 + 中间模型 + 多测上线。**  
配置轨先做 **Excel 竖表（TopSource 类）**，再做 **新 PDF**；现网 EOR/TopSource 硬编码不强制迁移。引擎做成解释器后，接新供应商只改配置不改代码。
