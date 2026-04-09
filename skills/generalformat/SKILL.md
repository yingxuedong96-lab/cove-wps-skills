---
name: generalformat
description: >
  文档格式排版。触发词：排版图、排版表格、排版正文、排版标题、排版页面。
---

# 通用文档排版

## ⚠️ 执行规则

**触发词匹配后，直接调用脚本，不做任何处理。**

用户输入中"排版图"、"排版表格"等触发词之后的内容，**原文直接作为specText传入脚本**，不需要理解、不需要解析。

---

## 执行方式

```
executeFile:
  filePath: "skills/generalformat/scripts/format-engine.js"
  params:
    config:
      specText: "<用户输入中触发词后的全部内容>"
      paragraphRules: {}
```

**注意**：paragraphRules 传空对象 `{}`，脚本会自动从 specText 解析规则。

---

## 示例

**用户输入**：`generalformat：排版图。图名用黑体小五号居中，图片居中对齐。`

**调用**：
```yaml
filePath: skills/generalformat/scripts/format-engine.js
params:
  config:
    specText: 图名用黑体小五号居中，图片居中对齐
    paragraphRules: {}
```

**用户输入**：`generalformat：排版表格。表名用黑体小五号居中，表格与页面等宽，跨页重复表头，表头用黑体五号居中加粗，表格内容用宋体五号靠左对齐。`

**调用**：
```yaml
filePath: skills/generalformat/scripts/format-engine.js
params:
  config:
    specText: 表名用黑体小五号居中，表格与页面等宽，跨页重复表头，表头用黑体五号居中加粗，表格内容用宋体五号靠左对齐
    paragraphRules: {}
```

---

## 返回值

`{ success: boolean, applied: number, elapsedMs: number }`

告知用户：排版完成，处理 {applied} 个元素。