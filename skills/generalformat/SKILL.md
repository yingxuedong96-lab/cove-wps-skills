---
name: generalformat
description: >
  文档格式排版。根据排版规则自动识别段落类型并应用格式。
  【强制要求】必须调用executeFile执行scripts/format-engine.js，禁止自己写代码！
  触发词：排版标题、排版正文、排版图表、排版表格、排版页面设置、排版页眉页脚、完整排版。
compatibility:
  runtime: WPS JS 宏 / WPS 加载项（JSAPI）
---

# 通用文档排版

## 使用方式

直接在 prompt 中提供排版规则，脚本自动识别段落类型并应用格式。

---

## 执行步骤

**⚠️⚠️⚠️ 重要：必须调用脚本执行，禁止自己写代码！**

本技能的所有功能都已实现在 `scripts/format-engine.js` 中，Agent只需：
1. 解析用户输入生成配置JSON
2. 调用 executeFile 执行脚本

**禁止行为**：
- ❌ 禁止自己编写WPS JS代码
- ❌ 禁止直接操作Application.ActiveDocument
- ❌ 禁止返回代码片段

**正确流程**：
```
用户输入 → 解析规则 → 生成JSON配置 → 调用format-engine.js → 返回执行结果
```

---

### Step 1 — 解析排版规则

从用户输入或快捷指令 prompt 中提取排版规则，生成配置 JSON。

**⚠️ 关键规则：**

1. **specText 必须完整**：将排版规则**完整**放入 specText 字段
2. **类型必须完整覆盖**：规则提到多少种类型，配置就必须有多少种
3. **属性必须完整提取**：字体、字号、对齐、段前段后、行距、缩进等，全部提取
4. **不要添加未提及的类型**：规则没说四级标题，就不要添加 heading4
5. **加粗规则**：说"加粗"设置 `bold: true`，说"不加粗"设置 `bold: false`，都没说则不设置
6. 直接输出 JSON，不要写代码，不要解释

---

### Step 2 — 调用脚本执行

```yaml
executeFile:
  filePath: "skills/generalformat/scripts/format-engine.js"
  params:
    config: <生成的JSON配置>
```

返回值：
```json
{
  "success": true,
  "applied": 5656,
  "elapsedMs": 2882,
  "typeCounts": { "zhangTitle": 20, "heading2": 421, "body": 5207 }
}
```

---

## 类型映射表

| 规范中的叫法 | 配置中的类型名 | 说明 |
|-------------|---------------|------|
| 主标题、文档标题、报告标题 | docTitle | 文档第一个段落 |
| 章标题、一级标题、标题一 | zhangTitle | 如"1 引言"、"第1章" |
| 附录标题 | appendixTitle | 如"附录一 主要设计规范"、"附录A" |
| 二级标题、标题二 | heading2 | 如"1.1 设计依据" |
| 三级标题、标题三 | heading3 | 如"1.1.1 具体内容" |
| 四级标题、标题四 | heading4 | 如"1.1.1.1 详细说明" |
| 五级标题、标题五 | heading5 | 如"1.1.1.1.1 补充说明" |
| 正文、正文格式 | body | 普通段落 |
| 图名、图标题、图号 | figureCaption | 图标题 |
| 表名、表标题、表号 | tableCaption | 表标题 |
| 参考文献 | ref | 参考文献 |

---

## 字号对照表

| 中文字号 | 磅值(fontSize) |
|---------|---------------|
| 初号 | 42 |
| 小初 | 36 |
| 一号 | 26 |
| 小一 | 24 |
| 二号 | 22 |
| 小二 | 18 |
| 三号 | 16 |
| 小三 | 15 |
| 四号 | 14 |
| 小四 | 12 |
| 五号 | 10.5 |
| 小五 | 9 |

---

## 字段说明

| 字段 | 类型 | 说明 |
|------|------|------|
| fontCN | string | 中文字体（宋体、黑体、楷体、仿宋） |
| fontEN | string | 西文字体（Times New Roman） |
| fontSize | number | 字号磅值 |
| bold | boolean | 是否加粗 |
| alignment | number | 对齐：0=左对齐，1=居中，2=右对齐，3=两端对齐 |
| spaceBefore | number | 段前间距（磅） |
| spaceAfter | number | 段后间距（磅） |
| firstLineIndent | number | 首行缩进（磅，2字符≈24磅） |
| lineSpacingRule | number | 行距：0=单倍，1=1.5倍，2=2倍，4=固定值 |

---

## 配置格式

```json
{
  "version": "1.0",
  "specText": "排版规则原文（必须完整）",
  "fontDefaults": { "fontCN": "宋体", "fontEN": "Times New Roman" },
  "paragraphRules": {
    "zhangTitle": { "fontCN": "黑体", "fontSize": 16, "bold": true, "alignment": 1, ... },
    "body": { "fontCN": "宋体", "fontSize": 12, "alignment": 3, "firstLineIndent": 24, ... }
  }
}
```

---

## 特殊规则（从 specText 自动识别）

| 关键词 | 功能 |
|-------|------|
| 图片居中 | 图片段落居中 |
| 表格等宽 | 表格宽度等于页面宽度 |
| 跨页重复表头 | 表格首行跨页重复 |
| 公式居中 | 公式段落居中 |
| 公式编号右对齐 | 公式编号右对齐 |
| 页码居中/左对齐/右对齐 | 页码位置 |
| 页码阿拉伯/罗马/中文 | 页码格式 |

---

## 快捷指令示例

**排版标题：**
```
generalformat：排版主标题和各级标题。
主标题用黑体二号居中加粗，
一级标题用黑体三号靠左加粗，
二级标题用黑体小三号靠左加粗，
三级标题用黑体四号靠左加粗，
四级标题用黑体四号靠左加粗，
五级标题用黑体四号靠左对齐，
附录标题用黑体三号居中，
所有标题无缩进。
```

**排版正文：**
```
generalformat：排版正文段落格式。
正文用宋体小四号两端对齐首行缩进2字符行距固定值22磅。
```

**排版图表：**
```
generalformat：排版图名表名。
图名用黑体小五号居中，表名用黑体小五号居中，图片居中对齐，表格与页面等宽。
```

---

## 段落类型自动识别

脚本自动识别段落类型并设置大纲级别：

| 类型 | 匹配规则 | 大纲级别 |
|------|---------|---------|
| zhangTitle | 第X章、数字+空格+汉字 | 1级 |
| heading2 | X.X 二级编号 | 2级 |
| heading3 | X.X.X 三级编号 | 3级 |
| heading4 | X.X.X.X 四级编号 | 4级 |
| figureCaption | 图X... | - |
| tableCaption | 表X... | - |
| body | 其他正文 | - |