---
name: templatecreator
description: >
  样式模板提取与应用。

  【触发词】提取模板、应用模板、样式模板、样式复制
---

# 样式模板技能

## 使用方式

### 有关联文档时（一键复制）
如果用户关联了文档，从消息中找到关联文档的路径，执行：
```
executeFile:
  filePath: skills/templatecreator/scripts/copy-from-linked.js
  params:
    sourceDocPath: "/完整路径/源文档.docx"
```

关联文档格式示例：
```
[引用文件：样式提取.docx]
文件路径：/Users/cassia/Desktop/dyx/wpsjs/模版生成/样式提取.docx
```

提取 `文件路径` 作为 `sourceDocPath` 参数。

### 无关联文档时
- 提取模板：执行 `scripts/extract-template.js`
- 应用模板：执行 `scripts/apply-template.js`

---

## 重要规则

**禁止调用 executeCode！禁止使用 askUser！**

只能使用 executeFile 工具直接执行脚本。

**禁止的操作：**
- ❌ 不要查找 artifact
- ❌ 不要使用 askUser 询问用户
- ❌ 不要尝试从其他文档提取

直接执行脚本即可！