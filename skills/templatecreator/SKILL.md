---
name: templatecreator
description: >
  样式模板提取与应用。从A文档提取样式模板，直接应用到B文档，无需中间保存文件。

  【触发词】提取模板、应用模板、样式模板、从文档提取样式、用A文档样式格式化B文档

  【核心流程】
  1. A文档「提取模板」 → 样式数据自动保存到服务器
  2. B文档「应用模板」 → 自动读取最近的模板数据并应用

  无需用户手动保存文件！
---

# 样式模板技能

## 功能说明

本技能实现**A文档提取 → B文档应用**的无缝流程：

1. **提取模板**：从当前文档扫描样式，自动识别类型（论文报告/公文）
2. **应用模板**：将最近的模板数据应用到当前文档

**支持的文档类型：**
- 论文/技术报告：主标题、1-5级标题、正文、图表名、参考文献、附录等（26种元素）
- 公文：发文机关、标题、1-3级标题、正文、附件说明、落款等（20种元素）

## 重要规则

**禁止编写或执行任何代码！禁止调用 executeCode！**

只能使用 executeFile 工具调用本技能目录下的脚本文件。

---

## ⚠️ 核心展示规则

### 脚本返回结构

```json
{
  "success": true,
  "message": "✅ 样式模板提取完成！\n\n📄 源文档：xxx.docx\n### 一级标题（5处）\n字体: 黑体 | 字号: 16pt | 加粗...",
  "styleCount": 7,
  "templateJson": "{...完整样式数据...}"
}
```

### ✅ 正确做法

**直接原样展示 message 字段内容！**

不要重新总结、不要简化参数、不要用"等"省略。

---

## 执行流程

### 流程1：提取模板（自动检测类型）

用户请求："提取模板"、"从当前文档提取样式"

#### 单步执行

```
executeFile:
  filePath: skills/templatecreator/scripts/extract-template.js
  params: {}
```

脚本会：
1. 自动检测文档类型（扫描前30段统计特征得分）
2. 提取所有样式参数
3. 返回完整样式详情

**返回后直接展示 message 内容！无需 askUser 选择文档类型！**

---

### 流程2：应用模板

用户请求："应用模板"、"用刚才提取的样式格式化这个文档"

#### 单步执行

```
executeFile:
  filePath: skills/templatecreator/scripts/apply-template.js
  params: {}
```

脚本会：
1. 自动读取最近提取的模板数据（从上下文或artifact）
2. 检测当前文档段落类型
3. 应用对应样式格式

---

### 流程3：A文档提取 → B文档应用

用户请求："用A文档的样式格式化B文档"

#### 步骤指导

1. **打开A文档** → 说「提取模板」
2. **切换到B文档** → 说「应用模板」

**无需中间保存步骤！模板数据自动保存在服务器端。**

---

## 样式规范表参考

### 论文报告元素（26种）

| 元素ID | 名称 | 检测模式 |
|--------|------|----------|
| chapterTitle | 章标题 | 第X章 |
| heading1 | 一级标题 | 1 XXX |
| heading2 | 二级标题 | 1.1 XXX |
| heading3 | 三级标题 | 1.1.1 XXX |
| heading4 | 四级标题 | 1.1.1.1 XXX |
| heading5 | 五级标题 | 1.1.1.1.1 XXX |
| body | 正文 | 默认 |
| figureCaption | 图名 | 图 X-X |
| tableCaption | 表名 | 表 X-X |
| referenceTitle | 参考文献标题 | 参考文献 |
| appendixTitle | 附录标题 | 附录 A |

### 公文元素（20种）

| 元素ID | 名称 | 检测模式 |
|--------|------|----------|
| docTitle | 公文标题 | 关于... |
| docNumber | 发文字号 | 国发〔2024〕1号 |
| heading1 | 一级标题 | 一、 |
| heading2 | 二级标题 | （一） |
| heading3 | 三级标题 | 1. |
| attachment | 附件说明 | 附件 |

---

## 参数说明（45个）

| 类别 | 参数 |
|------|------|
| 字体（7） | fontCN, fontEN, fontSize, fontSizeName, bold, italic, underline, color |
| 段落（5） | alignment, firstLineIndent, leftIndent, rightIndent, characterUnitFirstLine |
| 间距（4） | spaceBefore, spaceAfter, lineSpacing, lineSpacingRule |
| 页面（9） | paperSize, orientation, margins, header/footer distance |
| 其他（20） | 检测模式、样本文本等 |

---

## 常见问题

### Q: 模板保存在哪里？

A: 保存在服务器端，用户无需关心文件路径。应用模板时会自动读取最近提取的数据。

### Q: 能否指定某个模板？

A: 当前版本自动使用最近提取的模板。如需管理多个模板，请说「列出模板」查看。

### Q: 提取的样式不准确怎么办？

A: 提取结果会显示具体参数和样本，请检查样本是否符合预期。如有问题可反馈调整检测逻辑。