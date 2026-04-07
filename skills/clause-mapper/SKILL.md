---
name: clause-mapper
description: >
  条款智能映射。将用户草稿条款自动映射为标准协议条款并生成新文档。
---

# 条款智能映射

## 执行步骤

### Step 1 — 提取草稿条款

调用 `executeFile`：

```text
filePath: "skills/clause-mapper/scripts/extract-clauses.js"
params: {}
```

返回值：`{ clauses: [{ index: number, title: string, content: string }], hasSelection: boolean }`

- `clauses` 为空数组：回复"未检测到条款内容，请确认文档已打开且包含条款"，终止。
- 继续执行 Step 2。

### Step 2 — 加载映射规则库

调用 `skill_resource`：

```text
skillName: "clause-mapper"
resourcePath: "references/mapping-rules.md"
```

阅读映射规则库，理解以下映射模式：
- 一对一映射：单条款直接对应
- 多对一映射：多条合并到一条标准条款
- 一对多映射：一条拆分到多条标准条款
- 整合映射：多条信息整合到前言
- 拆分映射：不同条款合并到同一标准条款

### Step 3 — 加载标准条款库

调用 `skill_resource`：

```text
skillName: "clause-mapper"
resourcePath: "references/standard-clauses.md"
```

阅读标准条款库，获取 15 条标准条款模板。

### Step 4 — 智能映射分析

**重要说明：此步骤不需要调用任何工具。**

**直接使用 Step 1 返回的 clauses 数据，结合 Step 2 和 Step 3 加载的规则库和标准条款库，进行智能映射分析。**

**执行智能映射**：

1. **关键词匹配**：识别条款标题中的关键词（如"回购"、"清算"、"反稀释"）
2. **语义理解**：理解条款意图（如"交割前提条件"对应"交割先决条件"）
3. **参数提取**：从草稿条款中提取关键参数（投资额、估值、董事人数等）
4. **条款整合**：将多个相关条款合并（如投资人权利合并到第七条）
5. **条款拆分**：将常规条款拆分到对应标准条款

**生成输出**：

根据映射分析结果，生成以下内容：

1. **preamble**（前言内容）：整合项目名称、签署日期、投资标的、创始股东、投资人等信息
2. **clauses**（标准条款数组）：每条包含以下字段：
   - `title`（条款标题）：完整标题，如 "第一条 定义"、"第二条 本次交易安排"
   - `content`（条款内容）：条款的具体内容文本
   
   **注意**：不要生成 `number` 字段，标题中已包含中文编号。

### Step 5 — 创建新文档

调用 `executeFile`：

```text
filePath: "skills/clause-mapper/scripts/create-new-doc.js"
params: { 
  "preamble": "<Step 4 生成的 preamble 内容>",
  "clauses": "<Step 4 生成的 clauses 数组>"
}
```

返回值：`{ success: boolean, docName: string }`

告知用户：已生成新文档 `{docName}`，包含 {条款数量} 条标准条款。

## 异常处理

**脚本报错时，回复：**

> 操作失败，请确认文档已打开后重试。

**映射失败时，回复：**

> 部分条款无法自动映射，已生成包含可映射条款的文档，请人工审核未映射部分。
