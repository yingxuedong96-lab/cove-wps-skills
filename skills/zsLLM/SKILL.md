---
name: zsLLM
description: >
  技术报告语义检查。对中文技术报告(.docx)进行 LLM 语义分析，
  检查术语一致性、模糊引用、图表引用完整性、繁体字等问题，并添加批注。

  【触发词】语义检查、深度检查、对全文进行语义检查、语义分析
---

# 语义检查技能

## 重要规则

**禁止编写或执行任何代码！禁止调用 executeCode！**

只能使用 executeFile 工具调用本技能目录下的脚本文件。

## 执行流程（必须严格按顺序）

### 第一步：读取文档内容

**立即调用 executeFile 工具，参数如下：**

- filePath: `skills/zsLLM/scripts/get-document-content.js`
- params: `{"maxLength": 15000}`

等待工具返回结果，你会收到类似：
```
{"content": "【第1段】第一章...\n【第2段】...", "paraCount": 50, "charCount": 5000}
```

### 第二步：分析文档内容

**在脑中分析第一步返回的 content 字段内容，不要在对话中输出任何 JSON 或中间结果。**

检查以下问题：

| 规则ID | 检查项 | 示例 |
|--------|--------|------|
| L-001 | 术语一致性 | "混凝土"与"砼"混用 |
| L-002 | 模糊引用 | "如上图所示"但未指明具体图 |
| L-003 | 图表引用缺失 | 引用了不存在的图1或表2-1 |
| L-004 | 繁体字 | "計算"应为"计算" |

分析完成后，**直接进入第三步**，不要在对话中展示 JSON 数据。

### 第三步：添加批注

**将分析结果直接作为参数调用 executeFile，不要先在对话中展示 JSON：**

- filePath: `skills/zsLLM/scripts/semantic-check.js`
- params: `{"issues": <分析得到的问题数组>}`

示例：
```
filePath: skills/zsLLM/scripts/semantic-check.js
params: {"issues": [{"rule":"L-004","name":"繁体字检测","location":"第5段","original":"計算數據","suggestion":"计算数据"}]}
```

如果没有发现任何问题，传入空数组：`{"issues": []}`

**重要：工具会返回检查摘要。直接将 summary 字段内容呈现给用户，不要重复展示 JSON 数据。**

## 错误示例（禁止这样做）

❌ 错误：调用不存在的脚本 `init.js` 或 `main.js`
❌ 错误：使用 executeCode 编写代码
❌ 错误：在代码中调用 `chat()` 函数
❌ 错误：在对话中重复展示工具返回的摘要或 JSON 数据

## 正确示例

1. 用户："对文档进行语义检查"
2. 你调用 executeFile(filePath: "skills/zsLLM/scripts/get-document-content.js", params: {"maxLength": 15000})
3. 工具返回文档内容
4. 你分析内容，输出 JSON 数组
5. 你调用 executeFile(filePath: "skills/zsLLM/scripts/semantic-check.js", params: {"issues": [...]})
6. 工具添加批注，返回摘要
7. 你简单确认：**语义检查已完成，批注已添加到文档中。**