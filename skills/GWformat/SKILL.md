---
name: GWformat
description: >
  公文格式排版。按XX集团公文排版格式规范自动识别文档结构并应用样式。触发词：公文排版、格式排版、排版、GWformat。
---

# 公文格式排版

## 规范依据

XX集团公文排版格式规范。

## 执行步骤

### Step 1 — 创建样式

调用 `executeFile`：

```text
filePath: "skills/GWformat/scripts/setup-styles.js"
params: {}
```

返回值：`{ success: boolean, message: string }`

确保文档中存在所需的公文样式。

### Step 2 — 应用排版

调用 `executeFile`：

```text
filePath: "skills/GWformat/scripts/apply-format.js"
params: {}
```

返回值：`{ success: boolean, paragraphCount: number, appliedStyles: object }`

自动识别段落类型并应用对应样式：

| 元素类型 | 样式名 | 字体 | 字号 | 其他 |
|---------|--------|------|------|------|
| 标题 | 集团1标题 | 方正小标宋简体 | 22pt | 居中，加粗 |
| 一级标题 | 集团2级标题黑体 | 黑体 | 16pt | 左对齐 |
| 二级标题 | 集团3级段落重点 | 楷体_GB2312 | 16pt | 首行缩进 |
| 正文 | 集团正文文本缩进 | 仿宋_GB2312 | 16pt | 首行缩进2字符 |
| 落款 | 集团落款 | 仿宋_GB2312 | 16pt | 右对齐 |

**元素识别规则**：

| 类型 | 识别规则 | 示例 |
|------|---------|------|
| 标题 | 第一段，≤50字 | "关于OfficeAI的采购申请" |
| 一级标题 | 以"一、""二、"等开头 | "一、采购背景" |
| 二级标题 | 以"（一）""（二）"等开头 | "（一）提升办公效率" |
| 落款 | 末尾段落，含日期或公司名 | "深圳机智引擎科技有限公司" |
| 正文 | 其他段落 | - |

同时设置页面格式：A4纸张，标准页边距。

告知用户：排版完成，共处理 {paragraphCount} 个段落，应用样式：{appliedStyles}。

## 异常处理

**脚本报错时，回复：**

> 排版失败：{error}。请确认文档已打开后重试。