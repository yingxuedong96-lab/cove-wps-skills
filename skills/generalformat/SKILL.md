---
name: generalformat
description: >
  文档格式排版。根据排版规则自动识别段落类型并应用格式。
  触发词：排版标题、排版正文、排版图表、排版图、排版表格、排版页面、排版页眉页脚、完整排版。
compatibility:
  runtime: WPS JS 宏 / WPS 加载项（JSAPI）
---

# 通用文档排版

**⚠️ 核心原则：直接调用脚本，不探索文档，不写代码**

---

## Step 1 — 调用脚本执行排版

**直接调用：**
```text
filePath: "skills/generalformat/scripts/format-engine.js"
params:
  config:
    specText: "<用户输入的排版规则原文>"
    paragraphRules: <从规则中提取的类型配置>
```

**示例：**

排版图：
```json
{
  "specText": "图名用黑体小五号居中，图片居中对齐",
  "paragraphRules": {
    "figureCaption": {"fontCN":"黑体","fontSize":9,"alignment":1}
  }
}
```

排版表格：
```json
{
  "specText": "表名用黑体小五号居中，表格与页面等宽、跨页重复表头，表头用黑体五号居中加粗，表格内容用宋体五号靠左对齐",
  "paragraphRules": {
    "tableCaption": {"fontCN":"黑体","fontSize":9,"alignment":1},
    "tableHeader": {"fontCN":"黑体","fontSize":10.5,"alignment":1,"bold":true},
    "tableContent": {"fontCN":"宋体","fontSize":10.5,"alignment":0}
  }
}
```

排版正文：
```json
{
  "specText": "正文用宋体小四号两端对齐首行缩进2字符行距固定值22磅",
  "paragraphRules": {
    "body": {"fontCN":"宋体","fontSize":12,"alignment":3,"firstLineIndent":24,"lineSpacingRule":4}
  }
}
```

---

## 类型名映射

| 规范叫法 | 配置类型名 |
|---------|-----------|
| 主标题 | docTitle |
| 章标题/一级标题 | zhangTitle |
| 二级标题 | heading2 |
| 三级标题 | heading3 |
| 四级标题 | heading4 |
| 五级标题 | heading5 |
| 正文 | body |
| 图名/图号 | figureCaption |
| 表名/表号 | tableCaption |
| 表头 | tableHeader |
| 表格内容 | tableContent |
| 页眉页脚 | headerFooter |

---

## 字号对照

| 中文 | 磅值 |
|-----|-----|
| 小五 | 9 |
| 五号 | 10.5 |
| 小四 | 12 |
| 四号 | 14 |

---

## 对齐值

| 对齐方式 | 值 |
|---------|---|
| 左对齐 | 0 |
| 居中 | 1 |
| 右对齐 | 2 |
| 两端对齐 | 3 |

---

## 特殊关键词（自动处理）

| specText包含 | 自动处理 |
|-------------|---------|
| 图片居中/图片居中对齐 | 图片段落居中 |
| 图片左对齐/图片靠左 | 图片段落左对齐 |
| 图片右对齐/图片靠右 | 图片段落右对齐 |
| 表格等宽/与页面等宽 | 表格宽度=页面宽度 |
| 跨页重复表头 | 表格首行跨页重复 |