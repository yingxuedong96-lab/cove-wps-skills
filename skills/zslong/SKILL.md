---
name: zslong
description: >
  技术报告文档校对。按照设计文档编制规范，
  对中文技术报告(.docx)进行自动校对与格式修正。

  支持全文校对排版、编号校对、内容规范校对、格式规范排版等场景。
  快速触发词与 scope 的唯一对照以正文"快速模式（单一元素）"总表为准。
---

# 技术报告校对

## ⚠️ 第一优先级：判断执行模式

**按以下顺序判断：**

### 1. 全文校对排版模式

用户输入包含"全文校对排版"、"对全文进行校对排版"、"校对并排版全文"、"全文校对并排版"、"根据设计文档编制规范，对全文进行校对排版"时：

**⚠️ 长文档禁止单次执行 `scope: "full_proofread"`。必须按以下顺序串行执行多个 `executeFile`，每一步完成后再执行下一步：**

```
executeFile:
  filePath: "skills/zslong/scripts/scan-structure.js"
  params: { "scope": "numbering" }
```

```
executeFile:
  filePath: "skills/zslong/scripts/run-all-checks.js"
  params: { "fixMode": "aggressive", "isSelection": false, "scope": "table_content" }
```

```
executeFile:
  filePath: "skills/zslong/scripts/run-all-checks.js"
  params: { "fixMode": "aggressive", "isSelection": false, "scope": "value" }
```

```
executeFile:
  filePath: "skills/zslong/scripts/run-all-checks.js"
  params: { "fixMode": "aggressive", "isSelection": false, "scope": "page_setup" }
```

```
executeFile:
  filePath: "skills/zslong/scripts/run-all-checks.js"
  params: { "fixMode": "aggressive", "isSelection": false, "scope": "font" }
```

```
executeFile:
  filePath: "skills/zslong/scripts/run-all-checks.js"
  params: { "fixMode": "aggressive", "isSelection": false, "scope": "figure_table_layout" }
```

```
executeFile:
  filePath: "skills/zslong/scripts/run-all-checks.js"
  params: { "fixMode": "aggressive", "isSelection": false, "scope": "formula_layout" }
```

```
executeFile:
  filePath: "skills/zslong/scripts/run-all-checks.js"
  params: { "fixMode": "aggressive", "isSelection": false, "scope": "header_footer" }
```

**⚠️ 不要调用 `run-all-checks.js` with `scope: "full_proofread"` 处理长文档。**

**顺序原因：**
- 编号必须最先做，否则后面的图表题注、公式编号、标题层级都会基于旧编号排版。
- 表格内容和数值格式属于文本修订，必须先于排版。
- 页面设置要先于图表/公式排版，因为表格等宽和公式右对齐都依赖页面可用宽度。
- 公式排版必须放在编号之后，否则编号重写会把公式对齐结构打散。

### 2. 快速模式（单一元素）

检查用户输入是否包含以下特定元素关键词：

**⚠️ 关键：编号类scope走两阶段处理，内容/格式类scope走run-all-checks.js！**

**⚠️ 高优先级特判：如果用户输入包含 `同上`、`同左`、`单元格`、`相邻单元格` 中任意两个，或直接包含“将单元格内的‘同上’和‘同左’替换为实际数值”这类表述，必须优先命中 `table_content`，不要落到通用 `content`。**

**⚠️ 高优先级特判：如果用户输入包含 `数值格式`、`温度单位`、`补零`、`波浪号`、`范围`、`幂次`、`尺寸单位`、`分数`、`比例` 中任意两个，或直接包含“对文档进行数值格式规范校对”这类表述，必须优先命中 `value`，不要落到通用 `content`。**

**⚠️ 高优先级特判：如果用户输入包含 `各级标题`、`正文`、`排版`、`宋体小四`、`20磅行距`、`0.5行段间距`、`首行缩进2字` 中任意两个，或直接包含“根据设计文档编制规范对全文各级标题及正文进行排版”这类表述，必须优先命中 `font`，不要落到通用 `format`。**

**⚠️ 高优先级特判：如果用户输入包含 `图片`、`表格`、`图名`、`表名`、`图片居中`、`页面等宽`、`重复首行表头`、`表格文字居中` 中任意两个，或直接包含“根据设计文档编制规范对全文图片和表格进行排版”这类表述，必须优先命中 `figure_table_layout`，不要落到通用 `format`。**

**⚠️ 高优先级特判：如果用户输入包含 `公式`、`居中书写`、`编号右对齐`、`公式居中`、`公式编号右对齐` 中任意两个，或直接包含“根据设计文档编制规范将全文公式本身居中书写，其编号则需右对齐”这类表述，必须优先命中 `formula_layout`，不要落到通用 `format`。**

**⚠️ 高优先级特判：如果用户输入包含 `页眉`、`页脚`、`小五号`、`细线`、`双细线`、`页眉页脚排版` 中任意两个，或直接包含“根据设计文档编制规范对页眉页脚进行排版”这类表述，必须优先命中 `header_footer`，不要落到通用 `format`。**

**⚠️ 高优先级特判：如果用户输入包含 `页面设置`、`页边距`、`2.54cm`、`3.17cm`、`1.5cm`、`页眉和页脚距离边界` 中任意两个，或直接包含“根据设计文档编制规范对页面按照规范进行设置”这类表述，必须优先命中 `page_setup`，不要落到通用 `format`。**

**⚠️ 精确映射：如果用户输入接近以下表述，直接执行 `scope: "value"`，不要再做二次判断：**

```
根据设计院编制规范对文档进行数值格式规范校对。去除温度单位的冗余符号、小数前补零、统一使用波浪号（～）表示范围并确保单位和幂次在每个数值后重复书写、尺寸数据需逐项标注单位，以及分数和比例应采用标准数学记号。
```

**⚠️ 精确映射：如果用户输入接近以下表述，直接执行 `scope: "font"`，不要再做二次判断：**

```
根据设计文档编制规范对全文各级标题及正文进行排版。标题与正文统一为宋体小四、20磅行距及0.5行段间距；一二级标题加粗，正文明文首行缩进2字，其余无缩进。
```

**⚠️ 精确映射：如果用户输入接近以下表述，直接执行 `scope: "figure_table_layout"`，不要再做二次判断：**

```
根据设计文档编制规范对全文图片和表格进行排版。图名与表名均采用小四号宋体居中（图名段前后0.5行，表名段前0.5行段后0行），图片需居中对齐，表格宽度应与页面等宽且跨页时重复首行表头，表格内文字居中。
```

**⚠️ 精确映射：如果用户输入接近以下表述，直接执行 `scope: "formula_layout"`，不要再做二次判断：**

```
根据设计文档编制规范将全文公式本身居中书写，其编号则需右对齐。
```

**⚠️ 精确映射：如果用户输入接近以下表述，直接执行 `scope: "header_footer"`，不要再做二次判断：**

```
根据设计文档编制规范对页眉页脚进行排版。统一采用小五号字，页眉下设0.5磅通长细线，页脚上设0.5磅通长双细线。
```

**⚠️ 精确映射：如果用户输入接近以下表述，直接执行 `scope: "page_setup"`，不要再做二次判断：**

```
根据设计文档编制规范对页面按照规范进行设置，页边距上下为2.54cm、左右为3.17cm，同时页眉和页脚距离边界均为1.5cm。
```

#### 编号类触发词 → 直接执行 scan-structure.js（一步完成扫描+修复）

| 触发词示例 | scope | 执行路径 |
|-----------|-------|---------|
| 校对编号、校对所有编号、标题编号表编号图编号公式编号 | `numbering` | scan-structure.js |
| 校对标题编号、标题编号 | `heading` | scan-structure.js |
| 校对表编号、表编号 | `table` | scan-structure.js |
| 校对图编号、图编号 | `figure` | scan-structure.js |
| 校对公式编号、公式编号 | `formula` | scan-structure.js |

**编号类执行（一步完成，不调用其他脚本）：**

```
executeFile:
  filePath: "skills/zslong/scripts/scan-structure.js"
  params: { "scope": "<匹配的scope>" }
```

**⚠️ scan-structure.js 已完成扫描+修复，返回空 fixPlan，禁止再调用 apply-numbering-fixes.js！**

**⚠️ 编号类scope禁止使用调度器（submitScheduler）！编号校对需要全局上下文，不能分片处理！**

#### 内容/格式类触发词 → 走run-all-checks.js

| 关键词类型 | 触发词示例 | scope |
|-----------|-----------|-------|
| 数值类 | 校对数值、数值格式、校对所有数值、数值格式规范校对、去除温度单位的冗余符号、小数前补零、统一使用波浪号表示范围、确保单位和幂次在每个数值后重复书写、尺寸数据逐项标注单位、分数和比例采用标准数学记号 | `value` |
| 标点类 | 校对标点、校对所有标点 | `punctuation` |
| 内容校对 | 校对内容、校对所有内容 | `content` |
| 排版类 | 规范排版、全文排版、格式排版 | `format` |
| 标题正文类 | 各级标题及正文排版、标题正文排版、标题正文、根据设计文档编制规范对全文各级标题及正文进行排版、标题与正文统一为宋体小四20磅行距0.5行段间距 | `font` |
| 页眉页脚排版 | 页眉、页脚、页眉页脚、页眉排版、页脚排版、根据设计文档编制规范对页眉页脚进行排版、小五号字、页眉下设0.5磅通长细线、页脚上设0.5磅通长双细线 | `header_footer` |
| 页面设置 | 页面设置、页面排版、对页面进行设置、页边距、页面边距、根据设计文档编制规范对页面按照规范进行设置、页边距上下为2.54cm左右为3.17cm、页眉和页脚距离边界均为1.5cm | `page_setup` |
| 图表排版 | 图表排版、根据设计文档编制规范对全文图片和表格进行排版、图名与表名小四号宋体居中、图片需居中对齐、表格宽度与页面等宽、跨页重复首行表头、表格内文字居中 | `figure_table_layout` |
| 图片排版 | 图片排版 | `figure_layout` |
| 表格排版 | 表格排版 | `table_layout` |
| 公式排版 | 公式排版、对公式排版、根据设计文档编制规范将全文公式本身居中书写其编号则需右对齐、公式本身居中书写编号右对齐 | `formula_layout` |
| 图名排版 | 图名排版、图名格式 | `figure_caption` |
| 图片居中 | 图片居中 | `figure_center` |
| 表名排版 | 表名排版、表名格式 | `table_caption` |
| 表格内容 | 校对表格内容、表格内容同上同左、同上同左替换、将单元格内的“同上”和“同左”替换为实际数值、根据设计文档编制规范将单元格内的“同上”和“同左”文字分别替换为上方或左侧相邻单元格的实际数值 | `table_content` |

**内容/格式类执行：**

```
executeFile:
  filePath: "skills/zslong/scripts/run-all-checks.js"
  params: { "fixMode": "aggressive", "isSelection": false, "scope": "<匹配的scope>" }
```

**数值格式长句示例固定执行：**

```
executeFile:
  filePath: "skills/zslong/scripts/run-all-checks.js"
  params: { "fixMode": "aggressive", "isSelection": false, "scope": "value" }
```

**⚠️ 禁止执行 detect-mode.js！禁止询问修复策略！直接执行！**

---

## Stage 划分

```
Stage 1: 编号校对（文本修改）← scan-structure.js（使用段落扫描避免 Find.Execute 漏检）
    ↓
Stage 2: 内容规范（文本修改）← run-all-checks.js
    ↓
Stage 4: 格式规范排版（字体字号行距页面设置）← run-all-checks.js
```

**重要**：
- 编号校对走 `scan-structure.js`，使用段落扫描构建章节上下文后统一修复
- 内容/格式校对走 `run-all-checks.js`
- 格式规范排版放在最后执行，确保所有文本内容修改完成后再统一排版

---

## ⚡ 快速模式（按阶段归类）

### Stage 1: 编号校对（scan-structure.js）

| scope | 说明 | 执行脚本 |
|-------|------|---------|
| `numbering` | 标题+表+图+公式编号 | scan-structure.js |
| `heading` | 仅标题编号（N-002~008） | scan-structure.js |
| `table` | 仅表编号（T-001） | scan-structure.js |
| `figure` | 仅图编号（G-001） | scan-structure.js |
| `formula` | 仅公式编号（E-001） | scan-structure.js |

### Stage 2: 内容规范（run-all-checks.js）

| scope | 说明 |
|-------|------|
| `content` | 数值格式+图表标点+表格内容（全部） |
| `value` | V-006~013 数值格式 |
| `punctuation` | M-001~005 图表标点 |
| `table_content` | T-005~006 同上同左替换 |

### Stage 4: 格式规范排版（run-all-checks.js）

| scope | 说明 |
|-------|------|
| `format` | 字体字号行距页面设置 |
| `font` | 标题与正文规范（F-001~F-005） |
| `figure_caption` | 图名格式规范（G-002） |
| `figure_center` | 图片居中对齐（G-004） |
| `figure_layout` | 图片排版（G-002 + G-004） |
| `table_caption` | 表名格式规范（T-002） |
| `table_layout` | 表格排版（T-002 + T-004 + T-007） |
| `figure_table_layout` | 图表排版（G-002 + G-004 + T-002 + T-004 + T-007） |
| `formula_layout` | 公式排版（E-002, E-003） |
| `header_footer` | 页眉页脚排版（HF-001~003） |
| `page_setup` | 页面设置（PG-001~002） |

---

## 编号校对详细流程

编号校对通过 `scan-structure.js` 执行，使用段落扫描：

```text
filePath: "skills/zslong/scripts/scan-structure.js"
params: { "scope": "<numbering/heading/table/figure/formula>" }
```

**技术实现**：
- 单次遍历段落，记录当前章号/节号
- 图/表编号：识别旧格式 `图1` / `表1`，转换为 `图X.Y-Z` / `表X.Y-Z`
- 公式编号：识别旧格式 `(1-2)`，转换为 `(X.Y-Z)`

**输出**：
```json
{
  "success": true,
  "fixed": 50,
  "commented": 0,
  "revisionLog": [...],
  "summary": { "totalIssues": 50 }
}
```

---

## 内容/格式校对流程（run-all-checks.js）

调用 `executeFile`：

```text
filePath: "skills/zslong/scripts/run-all-checks.js"
params: {
  "fixMode": "aggressive",
  "isSelection": false,
  "scope": "<content/value/punctuation/format/font/...>"
}
```

**返回值**：
```json
{
  "fixed": number,
  "commented": number,
  "revisionLog": [...],
  "summary": {
    "totalIssues": number,
    "byRule": { "V-006": 5, "M-001": 3 },
    "scope": "content"
  }
}
```

---

## 规则详情

### Stage 1 编号规则

| 规则ID | 检查项 | 示例 |
|--------|--------|------|
| N-002 | 一级标题编号连续性 | 第一章→第三章，应为第二章 |
| N-003 | 二级标题章节号正确 | 3.1 在第二章下，应为 2.1 |
| N-004 | 二级标题节序号连续 | 2.1 → 2.3，应为 2.2 |
| N-005 | 三级标题编号 |
| N-006 | 四级标题编号 |
| N-007 | 附录标题编号格式 | 附录内应为 A1、A2 格式 |
| G-001 | 图编号连续性 | 图2-1 → 图2-3，应为 图2-2 |
| T-001 | 表编号连续性 | 表1-1 → 表1-3，应为 表1-2 |
| E-001 | 公式编号连续性 | (2-1) → (2-3)，应为 (2-2) |

### Stage 2 规则列表

#### 数值格式（V系列）

| 规则ID | 检查项 | 示例 | 修复方式 |
|--------|--------|------|----------|
| V-006 | 温度偏差格式 | 20℃±2℃ → 20±2℃ | 自动修复 |
| V-007 | 小数补零 | .5 → 0.5 | 自动修复 |
| V-008 | 数值范围格式 | 10-15 → 10～15 | 自动修复 |
| V-009 | 单位范围格式 | 10mm-20mm → 10～20mm | 自动修复 |
| V-010 | 百分比范围 | 10～20% → 10%～20% | 自动修复 |
| V-011 | 幂次范围 | 1～5×10³ → 每个数值后应写出幂次 | 批注提醒 |
| V-012 | 体积尺寸 | 240×240×60mm → 240mm×240mm×60mm | 自动修复 |
| V-013 | 中文数值书写 | 四分之三→3/4 | 批注提醒 |

#### 图表标点（M系列）

| 规则ID | 检查项 | 示例 |
|--------|--------|------|
| M-001 | 图名标点 | 图1 大坝剖面图。 → 图1 大坝剖面图 |
| M-002 | 表名标点 | 表2-1 材料参数表。 → 表2-1 材料参数表 |
| M-005 | 中文括号 | (混凝土面板) → （混凝土面板） |

#### 表格内容规范（T系列）

| 规则ID | 检查项 | 示例 |
|--------|--------|------|
| T-005 | 同上替换 | 表格中"同上"替换为上方单元格的数值 |
| T-006 | 同左替换 | 表格中"同左"替换为左侧单元格的数值 |

### Stage 4 规则列表

| 规则ID | 检查项 | 要求 |
|--------|--------|------|
| F-001 | 一级标题 | 宋体加粗，小四号（12磅），20磅行间距 |
| F-002 | 二级标题 | 宋体加粗，小四号（12磅），20磅行间距 |
| F-003 | 三级标题 | 宋体，小四号（12磅），20磅行间距 |
| F-004 | 四级标题 | 宋体，小四号（12磅），20磅行间距 |
| F-005 | 正文 | 宋体，小四号，首行缩进2字，20磅行间距 |
| G-002 | 图名字体字号 | 小四号宋体，居中 |
| G-004 | 图片居中 | 图片居中对齐 |
| T-002 | 表名字体字号 | 小四号宋体，居中 |
| T-004 | 表格宽度 | 表格与页面等宽 |
| T-007 | 跨页重复表头 | 表格首行跨页时重复 |
| E-002 | 公式居中 | 公式应居中书写 |
| E-003 | 公式编号位置 | 编号右对齐 |
| HF-001 | 页眉页脚字体字号 | 宋体+Arial，小五号(9磅) |
| HF-002 | 页眉线 | 通长细线，0.5磅 |
| HF-003 | 页脚线 | 通长双细线，0.5磅 |
| PG-001 | 页边距 | 上下2.54cm，左右3.17cm |
| PG-002 | 页眉页脚距边界 | 1.5cm |

---

## 完整模式（交付建议）

**触发条件**：用户说"校对这个文档"、"帮我检查一下"，未指定具体范围。

**⚠️ 交付时不要再使用 `detect-mode.js`、不要询问修复策略、不要走调度器。**

**推荐做法**：默认使用 `aggressive`，并按下面顺序串行执行快路径：

1. `scan-structure.js` with `scope: "numbering"`
2. `run-all-checks.js` with `scope: "table_content"`
3. `run-all-checks.js` with `scope: "value"`
4. `run-all-checks.js` with `scope: "page_setup"`
5. `run-all-checks.js` with `scope: "font"`
6. `run-all-checks.js` with `scope: "figure_table_layout"`
7. `run-all-checks.js` with `scope: "formula_layout"`
8. `run-all-checks.js` with `scope: "header_footer"`

**不推荐**：
- `scope: "full_proofread"` 直接处理长文档
- `scope: "content"` / `scope: "format"` 作为一体化入口交给宿主自由拆分
- `detect-mode.js`
- 调度器链路

---

## 异常处理

**调度提交失败时**：

> 任务提交失败：{error 原因}。请检查后重试。

**脚本报错时**：

> 操作失败：{错误信息}。请确认文档已打开后重试。

**文档为空时**：

> 文档内容为空，无法执行校对。

---

## 规则分类速查

| 类别 | 规则ID | 执行脚本 | 处理方式 |
|------|--------|---------|---------|
| 编号 | N-002~010, G-001, T-001, E-001 | run-all-checks.js | 自动修复 |
| 内容 | V-006~013, M-001~005, T-005~006 | run-all-checks.js | 批注/修复 |
| 格式 | F-001~005, G-002~004, T-002~007, E-002~003, HF-001~003, PG-001~002 | run-all-checks.js | 直接修复 |
