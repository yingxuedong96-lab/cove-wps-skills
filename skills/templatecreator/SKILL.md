---
name: templatecreator
description: >
  样式模板提取与应用。从A文档提取格式模板，保存为JSON，然后应用到B文档。

  【触发词】提取模板、应用模板、创建模板、样式模板、从文档提取样式
---

# 样式模板技能

## 功能说明

本技能用于：
1. **提取模板**：从排版好的文档（A文档）中提取样式规则，保存为JSON模板
2. **应用模板**：将保存的模板应用到新文档（B文档），自动格式化

支持的样式元素：
- 主标题、一级~五级标题
- 正文（字体、字号、行距、首行缩进）
- 图名、表名
- 列表项（多级缩进）
- 附录标题、附录节题
- 页眉页脚、页面设置

## 重要规则

**禁止编写或执行任何代码！禁止调用 executeCode！**

只能使用 executeFile 工具调用本技能目录下的脚本文件。

## 执行流程

### 场景1：提取模板

用户请求示例："从当前文档提取样式模板"、"提取排版模板"

#### 步骤1：执行提取脚本

调用 executeFile 工具：

```
filePath: skills/templatecreator/scripts/extract-template.js
params: {}
```

#### 步骤2：处理提取结果

脚本返回结果可能包含：

**情况A：提取成功，无需用户确认**
```json
{
  "success": true,
  "template": { ... },
  "message": "已提取X种样式",
  "templatePath": "templates/xxx.json"
}
```
→ 直接告知用户提取完成

**情况B：检测到不确定的格式，需要用户确认**
```json
{
  "success": true,
  "needUserInput": true,
  "uncertainFormats": [
    {"formatKey": "22pt黑体居中", "count": 1, "samples": ["XX报告"]},
    {"formatKey": "16pt黑体左对齐", "count": 5, "samples": ["1 范围", "2 设计"]}
  ],
  "question": "检测到多种大字号格式，请确认哪个是主标题？"
}
```
→ **调用 askUser 工具询问用户**

#### 步骤3：askUser 调用示例

```
question: "检测到以下格式，请确认元素类型映射：\n\n1. 22pt黑体居中（1处）：XX报告\n2. 16pt黑体左对齐（5处）：1 范围、2 设计...\n\n请选择主标题对应的格式："
options: ["22pt黑体居中", "16pt黑体左对齐", "都不是，需要手动指定"]
allowCustom: true
```

#### 步骤4：根据用户回答继续

用户选择后，再次调用 executeFile 传递确认信息：

```
filePath: skills/templatecreator/scripts/extract-template.js
params: {"confirmMapping": {"docTitle": "22pt黑体居中", "heading1": "16pt黑体左对齐"}}
```

---

### 场景2：应用模板

用户请求示例："把模板应用到当前文档"、"用xxx模板格式化"

#### 步骤1：列出可用模板（如用户未指定）

调用 executeFile：

```
filePath: skills/templatecreator/scripts/list-templates.js
params: {}
```

返回可用模板列表，让用户选择。

#### 步骤2：应用模板

```
filePath: skills/templatecreator/scripts/apply-template.js
params: {"templateName": "中航机载报告模板"}
```

或使用模板路径：

```
params: {"templatePath": "templates/中航机载.json"}
```

#### 步骤3：处理应用结果

脚本返回应用统计：

```json
{
  "success": true,
  "applied": {
    "docTitle": 1,
    "heading1": 5,
    "heading2": 12,
    "body": 150
  },
  "message": "已应用模板，共处理168个段落"
}
```

---

### 场景3：提取并立即应用

用户请求："用A文档的样式格式化B文档"

#### 步骤1：确认文档

使用 askUser 确认：

```
question: "请确认操作流程：\n1. 当前打开的文档是A文档（样式源）还是B文档（待格式化）？"
options: ["当前是A文档，我将打开B文档后继续", "当前是B文档，请先打开A文档提取样式"]
```

#### 步骤2：按场景1提取模板

#### 步骤3：用户切换到B文档后，按场景2应用模板

---

## 模板JSON结构参考

```json
{
  "name": "中航机载报告模板",
  "version": "1.0",
  "extractedFrom": "样式提取.docx",
  "styles": [
    {
      "type": "docTitle",
      "name": "主标题",
      "detect": {"pattern": "firstNonEmpty", "description": "第一个非空段落"},
      "format": {
        "fontCN": "黑体",
        "fontSize": 22,
        "bold": true,
        "alignment": "center"
      }
    },
    {
      "type": "heading1",
      "name": "一级标题",
      "detect": {"pattern": "^\\d+\\s", "description": "数字开头如 1 范围"},
      "format": {
        "fontCN": "黑体",
        "fontSize": 16,
        "bold": true,
        "alignment": "left",
        "outlineLevel": 1
      }
    },
    {
      "type": "body",
      "name": "正文",
      "detect": {"pattern": "default", "description": "默认段落"},
      "format": {
        "fontCN": "宋体",
        "fontSize": 12,
        "alignment": "justify",
        "firstLineIndent": 24,
        "lineSpacing": 22
      }
    }
  ],
  "pageSetup": {
    "topMargin": 71,
    "bottomMargin": 71,
    "leftMargin": 71,
    "rightMargin": 71
  }
}
```

## 错误处理

- 如果没有打开文档，脚本返回 `{"success": false, "error": "没有打开的文档"}`
- 如果模板不存在，返回 `{"success": false, "error": "模板不存在: xxx"}`
- 如果文档结构无法识别，返回 `needUserInput: true` 请求用户确认

## 示例对话

**用户**：从当前文档提取样式模板

**你**：好的，我来提取样式模板。[调用 executeFile: extract-template.js]

**工具返回**：
```json
{
  "success": true,
  "needUserInput": true,
  "uncertainFormats": [...],
  "question": "检测到多种格式，请确认..."
}
```

**你**：[调用 askUser] 检测到以下格式，请确认主标题是哪个...

**用户选择**：22pt黑体居中

**你**：[调用 executeFile: extract-template.js, params: {confirmMapping: {...}}]

**工具返回**：`{"success": true, "templatePath": "templates/模板_20260407.json"}`

**你**：样式模板已提取并保存，包含5种标题样式和正文样式。模板保存在：templates/模板_20260407.json