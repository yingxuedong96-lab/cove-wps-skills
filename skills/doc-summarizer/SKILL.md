---
name: doc-summarizer
description: >
  文档内容总结提炼。对文档内容进行分析，生成结构化总结并插入到文档中。
---

# 文档内容总结提炼

## 执行步骤

### Step 1 — 识别功能类型并检测文档

调用 `executeFile`：

```
filePath: "skills/doc-summarizer/scripts/detect-size.js"
params: { "userMessage": "<用户输入的完整消息>" }
```

返回值：`{ charCount: number, isSelection: boolean, mode: 'direct' | 'schedule', funcType: 'extract' | 'minutes' | 'weekly' }`

- `charCount = 0`：回复"文档内容为空"，终止。
- 根据 `funcType` 和 `mode` 字段决定执行路径。

### Step 2 — 按 funcType 分功能执行

**必须严格按照 Step 1 返回的 `funcType` 字段决定执行对应功能，不得自行判断。**

#### 功能 A：`funcType = 'extract'`（内容提炼）

告知用户："正在为您生成内容提炼..."

**Step 2.1 — 读取内容**

若 `mode = 'direct'`，调用 `executeFile`：

```
filePath: "skills/doc-summarizer/scripts/read-content.js"
params: { "isSelection": "<Step 1 的 isSelection>" }
```

返回值：`{ content: string, title: string }`

若 `mode = 'schedule'`，调用 `submitScheduler`：

```
skill: "doc-summarizer"
scheduleParams: { "isSelection": "<Step 1 的 isSelection>", "funcType": "extract" }
```

**Step 2.2 — 生成内容提炼**

基于文档内容进行分析：
- 生成简短摘要，保留核心观点和要点
- 语言简明，保持客观中立
- 不添加主观评论，仅总结文章内容本身
- 仅输出摘要内容，不需要解释
- 不要使用 # 或 * 等符号

**Step 2.3 — 插入内容提炼**

调用 `executeFile`：

```
filePath: "skills/doc-summarizer/scripts/insert-summary.js"
params: { "content": "<内容提炼结果>" }
```

#### 功能 B：`funcType = 'minutes'`（会议纪要）

告知用户："正在为您生成会议纪要..."

**Step 2.1 — 读取内容**

若 `mode = 'direct'`，调用 `executeFile`：

```
filePath: "skills/doc-summarizer/scripts/read-content.js"
params: { "isSelection": "<Step 1 的 isSelection>" }
```

返回值：`{ content: string, title: string }`

若 `mode = 'schedule'`，调用 `submitScheduler`：

```
skill: "doc-summarizer"
scheduleParams: { "isSelection": "<Step 1 的 isSelection>", "funcType": "minutes" }
```

**Step 2.2 — 生成会议纪要**

基于文档内容进行分析，生成结构化会议纪要：
- 会议主题
- 讨论要点
- 决策事项
- 待办事项（如有）
- 责任人及截止日期（如有）
- 语言客观简明，仅输出纪要内容
- 不要使用 # 或 * 等符号

如果文档内容不是会议记录，回复"该文档未检测到会议相关内容，无法生成会议纪要"，终止。

**Step 2.3 — 插入会议纪要**

调用 `executeFile`：

```
filePath: "skills/doc-summarizer/scripts/insert-summary.js"
params: { "content": "<会议纪要结果>" }
```

#### 功能 C：`funcType = 'weekly'`（周报助手）

告知用户："正在为您生成工作周报..."

**Step 2.1 — 读取内容**

若 `mode = 'direct'`，调用 `executeFile`：

```
filePath: "skills/doc-summarizer/scripts/read-content.js"
params: { "isSelection": "<Step 1 的 isSelection>" }
```

返回值：`{ content: string, title: string }`

若 `mode = 'schedule'`，调用 `submitScheduler`：

```
skill: "doc-summarizer"
scheduleParams: { "isSelection": "<Step 1 的 isSelection>", "funcType": "weekly" }
```

**Step 2.2 — 生成工作周报**

基于文档内容进行分析，生成工作周报：
- 本周完成工作
- 本周亮点
- 下周计划（如果原文未提及，不要强行生成）
- 需协调事项
- 语言简洁，仅输出周报正文
- 不要使用 # 或 * 等符号

如果文档内容与工作周报无关，回复"该文档未检测到工作周报相关内容，无法生成周报"，终止。

**Step 2.3 — 插入工作周报**

调用 `executeFile`：

```
filePath: "skills/doc-summarizer/scripts/insert-summary.js"
params: { "content": "<周报结果>" }
```

### Step 3 — 告知用户

告知用户：已成功生成并在文档末尾插入内容，字体已设置为深蓝色以便识别。

## 异常处理

**脚本报错时，回复：**

> 操作失败，请确认文档已打开后重试。

**调度提交失败（busy / init_failed）时，回复：**

> 任务提交失败：{error 原因}。请检查后重试。

**内容为空时，回复：**

> 未检测到有效内容，请确认文档内容后重试。
