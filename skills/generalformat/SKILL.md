---
name: generalformat
description: >
  文档格式排版。触发词：排版图、排版表格。
---

# 通用文档排版

## ⚠️ 重要：执行规则

**检测到触发词后，立即调用脚本。**

把用户输入中触发词后面的内容原文放入 `specText` 字段，**不要理解、不要解析、不要修改**。

---

## 执行方式

### 排版图

触发词：`排版图`

```text
filePath: "skills/generalformat/scripts/format-figure.js"
params:
  specText: "<用户输入中'排版图'后面的全部内容>"
```

### 排版表格

触发词：`排版表格`

```text
filePath: "skills/generalformat/scripts/format-table.js"
params:
  specText: "<用户输入中'排版表格'后面的全部内容>"
```

---

## 示例

**用户输入**：`generalformat：排版表格。表名用宋体小五号居中，表格与页面等宽，跨页重复表头，表头用黑体五号居中加粗，表格内容用宋体五号靠左对齐。`

**调用**：
```yaml
filePath: skills/generalformat/scripts/format-table.js
params:
  specText: 表名用宋体小五号居中，表格与页面等宽，跨页重复表头，表头用黑体五号居中加粗，表格内容用宋体五号靠左对齐
```

**注意**：`specText` 的值就是用户输入原文，脚本会自动解析字体、字号、对齐方式。

---

## 返回值

`{ success: boolean, applied: number }`

告知用户：排版完成，处理 {applied} 个元素。