---
name: jd-jswj
description: 对Word文档进行全面校对（标题编号、数值符号、错别字、标点规范），以修订模式（Track Changes）写回源文件，方便逐条审阅。触发词：校对、公文校对、文档校对、修订模式校对、错别字检查。
---

# 公文校对（极速模式）

## ⚡ 执行铁律

**除 Step 4 的最终汇报外，所有步骤禁止输出任何文字。**
分析、判断全部内部完成，不输出过程、不输出 JSON、不输出任何中间内容。

---

## 校对规则（内部参考，不输出）

1. **标题编号**：一级用阿拉伯数字，编号后两个半角空格，禁止末尾句号
2. **数值符号**：乘号用×，小数补前导零（.5→0.5），1-9非物理量用汉字，≥10用阿拉伯数字，物理量必须阿拉伯+单位
3. **标点规范**：时间范围用—，数值范围用～，年份括号用〔〕，并列书名号不加顿号，去掉"约…左右"冗余
4. **错别字**：纠正字形错误，英文品牌大小写（deepseek→DeepSeek）
5. **列项格式**：各条前加字母序号 a) b) c)

---

## 执行步骤

### Step 1：读取文档

```
executeFile: skills/jd-jswj/scripts/get-doc-info.js
```

- `success: false` → 只输出：`❌ {error}`，终止
- 记录 `path`、`charCount`、`paragraphs`、`docName`，**不输出任何内容**

### Step 2：生成修订清单（静默）

根据 `paragraphs` 对照校对规则，**内部生成** corrections 数组，**不输出任何文字**：

```json
[{"original":"含前后3字的原文片段","corrected":"修正后","type":"HEADING|NUMBER|TYPO|PUNCTUATION"}]
```

约束：`original` 必须是文档实际存在的原文；`original` 与 `corrected` 不能相同；无需修改则 corrections = []。

若 `path = 'schedule'`（字符数 ≥ 15000），改为执行：

```
executeFile: skills/jd-jswj/scripts/init-schedule.js
params:
  systemPrompt: <skill.json 中的 subAgentSystemPrompt>
  paragraphs: <Step 1 返回的 paragraphs>
```

用返回的 corrections 继续 Step 3。

### Step 3：写入修订（静默）

```
executeFile: skills/jd-jswj/scripts/apply-corrections.js
params:
  corrections: <Step 2 的 corrections 数组>
```

- `success: false` → 只输出：`❌ {error}`，终止
- 记录 `applied`、`notFound`，**不输出任何内容**

### Step 4：汇报（唯一允许输出的步骤）

corrections 为空时输出：`文档规范，无需修改。`

否则固定格式输出，不添加任何多余说明：

```
✅ 「{docName}」校对完成，{applied} 处修订已写入。
标题 {n1} | 数值 {n2} | 标点 {n3} | 错别字 {n4}
```

仅当 `notFound` 不为空时追加：
```
⚠️ {notFound.length} 处未匹配，请手动核查：{notFound 用顿号连接}
```
