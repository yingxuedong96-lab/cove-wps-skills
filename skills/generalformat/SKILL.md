---
name: generalformat
description: >
  文档格式排版。触发词：排版图、排版表格、排版正文、排版标题。
---

# 通用文档排版

## 执行方式

**⚠️ 触发词匹配后，直接调用对应脚本，params 传空对象 {}。**

### 排版图

```
executeFile:
  filePath: "skills/generalformat/scripts/format-figure.js"
  params: {}
```

### 排版表格

```
executeFile:
  filePath: "skills/generalformat/scripts/format-table.js"
  params: {}
```

---

## 返回值

`{ success: boolean, applied: number }`

告知用户：排版完成，处理 {applied} 个元素。