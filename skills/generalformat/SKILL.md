---
name: generalformat
description: >
  文档格式排版。根据排版规则自动识别段落类型并应用格式。触发词：排版、排版图、排版表格、排版正文、排版标题。
---

# 通用文档排版

## 执行步骤

### Step 1 — 调用排版脚本

调用 `executeFile`：

```text
filePath: "skills/generalformat/scripts/format-engine.js"
params:
  config:
    specText: "图名用黑体小五号居中，图片居中对齐"
    paragraphRules:
      figureCaption:
        fontCN: 黑体
        fontSize: 9
        alignment: 1
```

**排版图示例**：
```yaml
specText: "图名用黑体小五号居中，图片居中对齐"
paragraphRules:
  figureCaption: {fontCN: 黑体, fontSize: 9, alignment: 1}
```

**排版表格示例**：
```yaml
specText: "表格与页面等宽、跨页重复表头，表头用黑体五号居中加粗，表格内容用宋体五号靠左"
paragraphRules:
  tableHeader: {fontCN: 黑体, fontSize: 10.5, alignment: 1, bold: true}
  tableContent: {fontCN: 宋体, fontSize: 10.5, alignment: 0}
```

**排版正文示例**：
```yaml
specText: "正文用宋体小四号两端对齐首行缩进2字符"
paragraphRules:
  body: {fontCN: 宋体, fontSize: 12, alignment: 3, firstLineIndent: 24}
```

---

### 类型名

| 规范叫法 | 类型名 |
|---------|-------|
| 图名 | figureCaption |
| 表名 | tableCaption |
| 表头 | tableHeader |
| 表格内容 | tableContent |
| 正文 | body |
| 章标题 | zhangTitle |

---

### 字号

| 中文 | 磅值 |
|-----|-----|
| 小五 | 9 |
| 五号 | 10.5 |
| 小四 | 12 |
| 四号 | 14 |

---

### 对齐

| 方式 | 值 |
|-----|---|
| 左 | 0 |
| 居中 | 1 |
| 右 | 2 |
| 两端 | 3 |

---

### 自动关键词

| specText包含 | 自动处理 |
|-------------|---------|
| 图片居中 | 图片居中 |
| 表格等宽 | 表格=页面宽度 |
| 跨页重复表头 | 表头跨页重复 |

---

## 返回值

`{ success: boolean, applied: number, elapsedMs: number }`

告知用户：排版完成，处理 {applied} 个元素。