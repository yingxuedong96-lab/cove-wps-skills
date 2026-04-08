---
name: templatecreator
description: >
  样式模板提取与应用。

  【触发词】样式复制、提取模板、应用模板
---

# 样式模板技能

## 使用方式

执行 `scripts/copy-from-linked.js`：

**有关联文档时**（从消息中提取路径）：
```
executeFile:
  filePath: skills/templatecreator/scripts/copy-from-linked.js
  params:
    sourceDocPath: "/完整路径/源文档.docx"
```

关联文档格式：
```
[引用文件：样式提取.docx]
文件路径：/Users/cassia/Desktop/dyx/wpsjs/模版生成/样式提取.docx
```

**无关联文档时**：
```
executeFile:
  filePath: skills/templatecreator/scripts/copy-from-linked.js
  params: {}
```
从当前文档提取样式。

---

## 重要规则

**禁止使用 askUser！**

直接执行脚本，脚本会自动判断模式。