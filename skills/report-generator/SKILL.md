---
name: report-generator
description: >
  生成各类研究报告。根据用户输入的主题，生成完整的结构化研究报告，包含章节、段落和表格。
---

# 研究报告生成

## 执行步骤

### Step 1 — 获取用户输入

从用户消息中提取报告主题，记为 `topic`。

- 若用户没有明确主题 → 回复"请提供报告主题，例如：生成一篇关于人工智能发展趋势的研究报告"
- 继续执行 Step 2

### Step 2 — 调用生成脚本

调用 `executeFile`：

```text
filePath: "skills/report-generator/scripts/generate-report.js"
params: { "topic": "<Step 1 的 topic>" }
```

返回值：`{ success: boolean, content: string, error?: string }`

- `success = false`：回复"生成失败：{error}"，终止
- 继续执行 Step 3

### Step 3 — 写入文档

调用 `executeFile`：

```text
filePath: "skills/report-generator/scripts/generate-report.js"
params: { "action": "write", "content": "<Step 2 的 content>" }
```

返回值：`{ success: boolean, error?: string }`

告知用户：研究报告已生成完毕。

## 异常处理

**脚本报错时，回复：**

> 操作失败，请确认文档已打开后重试。

**服务端调用失败时，回复：**

> 报告生成失败，请稍后重试。