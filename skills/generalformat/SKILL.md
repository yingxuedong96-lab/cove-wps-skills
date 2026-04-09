---
name: generalformat
description: >
  文档格式排版。触发词：排版图、排版表格。
---

# 通用文档排版

## ⚠️ 重要：执行规则

**检测到触发词后，立即调用脚本，忽略用户输入中的其他内容。**

用户可能输入：`排版图。图名用黑体小五号...` 或 `排版表格。表名用黑体...`

**忽略触发词后面的内容，直接按以下映射调用脚本：**

---

## 执行方式

### 排版图

触发词：`排版图`

**直接调用（忽略后续内容）：**
```text
filePath: "skills/generalformat/scripts/format-figure.js"
params: {}
```

### 排版表格

触发词：`排版表格`

**直接调用（忽略后续内容）：**
```text
filePath: "skills/generalformat/scripts/format-table.js"
params: {}
```

---

## 返回值

`{ success: boolean, applied: number }`

告知用户：排版完成，处理 {applied} 个元素。