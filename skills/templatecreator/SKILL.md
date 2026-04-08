---
name: templatecreator
description: >
  样式模板提取与应用。

  【触发词】提取模板、应用模板、样式模板
---

# 样式模板技能

## 使用方式

### 提取模板
直接执行 `scripts/extract-template.js`

### 应用模板
直接执行 `scripts/apply-template.js`（无需参数，脚本自动读取最近模板）

---

## 重要规则

**禁止调用 executeCode！禁止使用 askUser！**

只能使用 executeFile 工具直接执行脚本。

**正确的执行方式：**
```
executeFile:
  filePath: skills/templatecreator/scripts/extract-template.js
  params: {}

executeFile:
  filePath: skills/templatecreator/scripts/apply-template.js
  params: {}
```

**禁止的操作：**
- ❌ 不要查找 artifact
- ❌ 不要使用 askUser 询问用户
- ❌ 不要尝试从其他文档提取

直接执行脚本即可！脚本会处理所有逻辑。