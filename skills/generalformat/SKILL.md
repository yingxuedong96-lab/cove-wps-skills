---
name: generalformat
description: >
  文档格式排版。触发词：排版图、排版表格、排版正文、排版标题、排版页面。
---

# 通用文档排版

## ⚠️ 执行流程

根据用户输入的触发词，直接执行对应配置的脚本。

### 触发词 → 配置映射表

| 触发词 | specText | paragraphRules |
|-------|----------|----------------|
| 排版图 | 图名用黑体小五号居中，图片居中对齐 | `{figureCaption:{fontCN:黑体,fontSize:9,alignment:1}}` |
| 排版表格 | 表名用黑体小五号居中，表格与页面等宽、跨页重复表头，表头用黑体五号居中加粗，表格内容用宋体五号靠左对齐 | `{tableCaption:{fontCN:黑体,fontSize:9,alignment:1},tableHeader:{fontCN:黑体,fontSize:10.5,alignment:1,bold:true},tableContent:{fontCN:宋体,fontSize:10.5,alignment:0}}` |
| 排版正文 | 正文用宋体小四号两端对齐首行缩进2字符 | `{body:{fontCN:宋体,fontSize:12,alignment:3,firstLineIndent:24}}` |
| 排版标题 | 章标题用黑体三号居中加粗，二级标题用黑体小三左对齐 | `{zhangTitle:{fontCN:黑体,fontSize:16,alignment:1,bold:true},heading2:{fontCN:黑体,fontSize:15,alignment:0}}` |

---

## 执行方式

**⚠️ 直接调用脚本，不要做任何额外处理。**

```
executeFile:
  filePath: "skills/generalformat/scripts/format-engine.js"
  params:
    config:
      specText: "<匹配的specText>"
      paragraphRules: <匹配的paragraphRules>
```

---

## 完整示例

**排版图**：
```yaml
filePath: skills/generalformat/scripts/format-engine.js
params:
  config:
    specText: 图名用黑体小五号居中，图片居中对齐
    paragraphRules:
      figureCaption: {fontCN: 黑体, fontSize: 9, alignment: 1}
```

**排版表格**：
```yaml
filePath: skills/generalformat/scripts/format-engine.js
params:
  config:
    specText: 表名用黑体小五号居中，表格与页面等宽、跨页重复表头，表头用黑体五号居中加粗，表格内容用宋体五号靠左对齐
    paragraphRules:
      tableCaption: {fontCN: 黑体, fontSize: 9, alignment: 1}
      tableHeader: {fontCN: 黑体, fontSize: 10.5, alignment: 1, bold: true}
      tableContent: {fontCN: 宋体, fontSize: 10.5, alignment: 0}
```

---

## 类型名

| 规范叫法 | 类型名 |
|---------|-------|
| 图名 | figureCaption |
| 表名 | tableCaption |
| 表头 | tableHeader |
| 表格内容 | tableContent |
| 正文 | body |
| 章标题 | zhangTitle |

---

## 字号对照

| 中文 | 磅值 |
|-----|-----|
| 小五 | 9 |
| 五号 | 10.5 |
| 小四 | 12 |
| 四号 | 14 |
| 小三 | 15 |
| 三号 | 16 |

---

## 对齐值

| 方式 | 值 |
|-----|---|
| 左 | 0 |
| 居中 | 1 |
| 右 | 2 |
| 两端 | 3 |

---

## 自动关键词

| specText包含 | 自动处理 |
|-------------|---------|
| 图片居中 | 图片居中 |
| 表格等宽 | 表格=页面宽度 |
| 跨页重复表头 | 表头跨页重复 |

---

## 返回值

`{ success: boolean, applied: number, elapsedMs: number }`

告知用户：排版完成，处理 {applied} 个元素。