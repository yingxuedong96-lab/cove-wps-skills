---
name: templatecreator
description: >
  样式模板提取与应用。基于预定义的样式规范表，从A文档提取格式模板，保存为JSON，然后应用到B文档。

  【触发词】提取模板、应用模板、创建模板、样式模板、从文档提取样式
---

# 样式模板技能

## 功能说明

本技能基于**样式规范表**进行结构化的模板提取与应用：

1. **规范表定义**：预先定义公文和论文的所有标签（标签ID+名称）及其属性（字体、字号、缩进等）
2. **解析匹配**：从文档扫描段落，将格式匹配到规范表的标签
3. **对话确认**：不确定的格式通过askUser让用户确认，支持"帮我找一下"
4. **模板应用**：将确认后的模板应用到新文档

**支持的文档类型：**
- 论文/技术报告：主标题、作者、摘要、关键词、1-5级标题、正文、图表名、参考文献、附录
- 公文：发文机关、文号、标题、1-3级标题、正文、附件说明、落款

## 重要规则

**禁止编写或执行任何代码！禁止调用 executeCode！**

只能使用 executeFile 工具调用本技能目录下的脚本文件。

## 执行流程

### 场景1：提取模板

用户请求示例："从当前文档提取样式模板"、"提取排版模板"

#### 步骤1：选择文档类型

调用 executeFile：

```
filePath: skills/templatecreator/scripts/extract-template.js
params: {}
```

脚本返回：

```json
{
  "success": true,
  "needUserInput": true,
  "stage": "selectDocType",
  "question": "请选择文档类型，以便使用对应的样式规范表：",
  "options": ["论文/技术报告", "公文"]
}
```

→ **调用 askUser 让用户选择**

#### 步骤2：继续提取

用户选择后，再次调用：

```
filePath: skills/templatecreator/scripts/extract-template.js
params: {"docType": "论文/技术报告"}
```

#### 步骤3：确认未匹配格式（如有）

脚本可能返回：

```json
{
  "success": true,
  "needUserInput": true,
  "stage": "confirmUnmatched",
  "unmatchedFormats": [
    {"formatSignature": "22pt黑体加粗居中", "count": 1, "samples": ["XX系统报告"]}
  ],
  "availableTags": [...],
  "question": "检测到以下格式未能自动识别，请帮助确认..."
}
```

→ **调用 askUser**

用户可能的回复：
- 直接指定："22pt黑体加粗居中 是 主标题"
- 请求帮助："帮我找一下类似的格式"

#### 步骤4：处理"帮我找一下"

当用户说"帮我找一下"时：

**LLM应当**：在当前文档中搜索类似格式的段落，提供更多示例帮助用户判断

可以再次调用脚本搜索：

```
filePath: skills/templatecreator/scripts/extract-template.js
params: {"docType": "论文", "searchFormat": "22pt黑体加粗居中"}
```

#### 步骤5：生成模板并展示结果

用户确认后：

```
filePath: skills/templatecreator/scripts/extract-template.js
params: {"docType": "论文", "confirmMapping": {"22pt黑体加粗居中": "docTitle"}}
```

返回格式：

```json
{
  "success": true,
  "title": "样式模板提取完成",
  "docName": "样式提取.docx",
  "docType": "论文报告样式",
  "totalStyles": 10,
  "totalParagraphs": 200,
  "stylesTable": [
    {
      "样式名称": "论文标题",
      "出现次数": "1处",
      "字体": "黑体",
      "字号": "22pt",
      "加粗": "是",
      "对齐": "居中",
      "首行缩进": "-",
      "行距": "-"
    }
  ],
  "pageSetup": {...},
  "templateJson": {...}
}
```

---

## ⚠️ 展示格式要求（必须遵守）

**脚本返回的 `message` 字段已经包含了格式化好的详细信息，直接把这个 message 内容展示给用户即可！**

**禁止重新格式化或简化 message 内容！禁止忽略 message 字段！**

返回结构示例：
```json
{
  "success": true,
  "message": "✅ 样式模板提取完成！\n\n📄 源文档：xxx.docx\n...\n### 一级标题（5处）\n字体: 黑体 | 字号: 16pt | 加粗 | 对齐: 左对齐\n...",
  "templateJson": {...},
  "stylesTable": [...]
}
```

**正确做法**：直接把 message 字段内容展示给用户。

**错误做法**：忽略 message，只用 stylesTable 重新格式化，或只展示简单摘要。

---

## 📁 JSON模板文件获取

模板JSON保存在服务器端skill目录，用户无法直接访问文件系统。

**正确做法**：
1. 从返回的 `templateJson` 字段获取完整JSON内容
2. 告知用户：模板已保存，JSON内容可在对话中查看或复制
3. 如果用户需要保存：提供 templateJson 内容让用户复制保存

---

### 场景2：应用模板

用户请求："把模板应用到当前文档"、"用xxx模板格式化"

#### 步骤1：列出可用模板

```
filePath: skills/templatecreator/scripts/list-templates.js
params: {}
```

#### 步骤2：应用模板

```
filePath: skills/templatecreator/scripts/apply-template.js
params: {"templateName": "中航机载报告模板"}
```

---

### 场景3：提取并立即应用

用户请求："用A文档的样式格式化B文档"

使用 askUser 确认流程，先提取A文档模板，再应用到B文档。

---

## 样式规范表结构

### 论文报告样式标签（共20个）

| 标签ID | 名称 | 检测提示 | 默认属性 |
|--------|------|----------|----------|
| docTitle | 论文标题 | 首段大字居中 | 22pt黑体加粗居中 |
| heading1 | 一级标题 | 数字开头如'1 引言' | 16pt黑体加粗 |
| heading2 | 二级标题 | 如'1.1 背景' | 15pt黑体加粗 |
| heading3 | 三级标题 | 如'1.1.1' | 14pt黑体加粗 |
| heading4 | 四级标题 | 如'1.1.1.1' | 12pt黑体加粗 |
| heading5 | 五级标题 | 如'1.1.1.1.1' | 12pt黑体加粗 |
| body | 正文 | 默认类型 | 12pt宋体两端缩进2字符 |
| figureCaption | 图名 | '图'开头 | 9pt黑体居中 |
| tableCaption | 表名 | '表'开头 | 9pt黑体居中 |
| listItem | 列表项 | 列表符号 | 12pt宋体 |
| reference | 参考文献 | 文献编号 | 10.5pt宋体悬挂缩进 |
| referenceTitle | 参考文献标题 | '参考文献' | 14pt黑体居中加粗 |
| abstractTitle | 摘要标题 | '摘要' | 14pt黑体居中加粗 |
| keywords | 关键词 | '关键词' | 10.5pt宋体 |
| appendixTitle | 附录标题 | '附录' | 16pt黑体居中加粗 |
| appendixSection | 附录节题 | 如'A.1' | 14pt黑体加粗 |

### 公文样式标签（共10个）

| 标签ID | 名称 | 检测提示 | 默认属性 |
|--------|------|----------|----------|
| docTitle | 公文标题 | 主标题居中大字 | 22pt方正小标宋居中 |
| docNumber | 发文字号 | 如'国发〔2024〕1号' | 16pt仿宋居中红色 |
| issuer | 发文机关 | 发文单位名称 | 22pt方正小标宋居中红色 |
| heading1 | 一级标题 | 汉字数字加顿号'一、' | 16pt黑体 |
| heading2 | 二级标题 | 括号加汉字'(一)' | 16pt楷体 |
| heading3 | 三级标题 | 阿拉伯数字加点 | 16pt仿宋加粗 |
| body | 正文 | 默认类型 | 16pt仿宋两端缩进2字符 |
| attachment | 附件说明 | '附件'开头 | 16pt仿宋 |
| signature | 落款 | 右下方署名 | 16pt仿宋右齐 |

### 属性定义

每个标签包含以下属性（根据类型可选）：

```javascript
{
  fontCN: "中文字体",
  fontEN: "英文字体",
  fontSize: 22,          // pt
  bold: true,            // boolean
  italic: false,
  color: "auto",         // 或 "red"
  alignment: "center",   // left/center/right/justify
  firstLineIndent: 2,    // 字符数
  leftIndent: 0,
  hangingIndent: 0,      // 悬挂缩进
  lineSpacing: 28,       // pt
  spaceBefore: 12,       // pt
  spaceAfter: 6          // pt
}
```

---

## 对话示例

**用户**：从当前文档提取样式模板

**你**：[executeFile: extract-template.js]

**返回**：`{"needUserInput": true, "stage": "selectDocType", "options": ["论文/技术报告", "公文"]}`

**你**：[askUser] 请选择文档类型...

**用户**：论文/技术报告

**你**：[executeFile: extract-template.js, params: {docType: "论文/技术报告"}]

**返回**：`{"needUserInput": true, "stage": "confirmUnmatched", "unmatchedFormats": [{"formatSignature": "22pt黑体加粗居中", ...}]}`

**你**：[askUser] 检测到22pt黑体加粗居中格式未识别，请确认它属于哪种标签？

**用户**：帮我找一下类似的格式

**你**：让我在文档中搜索更多示例... 发现该格式出现在文档开头"XX系统报告"，这通常是主标题的位置。建议映射为"论文标题"。

**用户**：好的，是论文标题

**你**：[executeFile: extract-template.js, params: {docType: "论文", confirmMapping: {"22pt黑体加粗居中": "docTitle"}}]

**返回**：
```json
{
  "success": true,
  "styleDetails": [...],
  "pageSetupInfo": {...},
  "templateFileName": "模板_paper_20260407.json",
  "fullTemplatePath": "/Users/.../skills/templatecreator/templates/模板_paper_20260407.json"
}
```

**你**：（按格式展示）
✅ 样式模板提取完成！

📄 源文档：样式提取.docx
📁 模板保存位置：/Users/.../skills/templatecreator/templates/模板_paper_20260407.json

📋 提取结果（共10种样式）：

| 样式名称 | 出现次数 | 详细参数 |
|---------|---------|---------|
| 论文标题 | 1处 | 字体: 黑体 \| 字号: 22pt \| 加粗 \| 对齐: 居中 |
| 一级标题 | 5处 | 字体: 黑体 \| 字号: 16pt \| 加粗 \| 对齐: 左对齐 |
| 正文 | 150处 | 字体: 宋体 \| 字号: 12pt \| 对齐: 两端对齐 \| 首行缩进: 2.0字符 \| 行距: 22.0pt |
| ... | ... | ... |

📐 页面设置：A4纸张，边距 上下2.54cm 左右3.17cm