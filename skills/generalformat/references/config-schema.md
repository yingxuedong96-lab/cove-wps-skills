# 排版配置格式说明

## 完整配置结构

```json
{
  "version": "1.0",
  "fontDefaults": { ... },
  "numberingPatterns": { ... },
  "paragraphRules": { ... },
  "documentInfo": { ... }
}
```

---

## 一、fontDefaults - 默认字体

```json
{
  "fontDefaults": {
    "fontCN": "宋体",
    "fontEN": "Times New Roman"
  }
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| fontCN | string | 默认中文字体 |
| fontEN | string | 默认西文字体 |

---

## 二、numberingPatterns - 编号识别规则

```json
{
  "numberingPatterns": {
    "zhangTitle": ["^第[一二三四五六七八九十]+章", "^\\d{1,2}\\s+[\\u4e00-\\u9fff]"],
    "heading2": ["^\\d+\\.\\d+(\\s|$)"],
    "heading3": ["^\\d+\\.\\d+\\.\\d+(\\s|$)"],
    "tableCaption": ["^表\\s*[\\d\\.A-Za-z\\-]+"],
    "figureCaption": ["^图\\s*[\\d\\.A-Za-z\\-]+"]
  }
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| zhangTitle | string[] | 章标题的正则数组 |
| heading2 | string[] | 二级标题的正则数组 |
| heading3 | string[] | 三级标题的正则数组 |
| heading3Plus | string[] | 四级及以下的正则数组 |
| tableCaption | string[] | 表名的正则数组 |
| figureCaption | string[] | 图名的正则数组 |

**正则编写要点**：
- 使用 `^` 锚定行首
- 使用 `\\d` 表示数字（JSON 中需双转义）
- 使用 `[\\u4e00-\\u9fff]` 匹配中文字符

---

## 三、paragraphRules - 段落排版规则

```json
{
  "paragraphRules": {
    "zhangTitle": {
      "fontCN": "黑体",
      "fontSize": 16,
      "bold": true,
      "alignment": 1,
      "spaceBefore": 12,
      "spaceAfter": 6
    },
    "body": {
      "fontCN": "宋体",
      "fontSize": 12,
      "alignment": 3,
      "firstLineIndent": 24,
      "lineSpacing": 20,
      "lineSpacingRule": 4
    }
  }
}
```

### 支持的字段

| 字段 | 类型 | 说明 |
|------|------|------|
| fontCN | string | 中文字体 |
| fontEN | string | 西文字体 |
| fontSize | number | 字号（磅值） |
| bold | boolean | 是否加粗 |
| italic | boolean | 是否斜体 |
| color | number | 字体颜色（RGB数值） |
| alignment | number | 对齐方式 |
| firstLineIndent | number | 首行缩进（磅值） |
| spaceBefore | number | 段前间距（磅值） |
| spaceAfter | number | 段后间距（磅值） |
| lineSpacing | number | 行距值 |
| lineSpacingRule | number | 行距类型 |

### 对齐方式

| 值 | 说明 |
|---|------|
| 0 | 左对齐 |
| 1 | 居中 |
| 2 | 右对齐 |
| 3 | 两端对齐 |

### 行距类型

| 值 | 说明 |
|---|------|
| 0 | 单倍行距 |
| 1 | 1.5 倍行距 |
| 2 | 2 倍行距 |
| 4 | 固定值 |

---

## 四、documentInfo - 文档信息

```json
{
  "documentInfo": {
    "pageMargins": {
      "top": 56.7,
      "bottom": 56.7,
      "left": 70.9,
      "right": 70.9
    }
  }
}
```

**单位换算**：1 厘米 ≈ 28.35 磅

---

## 速查表

### 字号对照

| 中文字号 | 磅值 |
|---------|------|
| 小五号 | 9 |
| 五号 | 10.5 |
| 小四号 | 12 |
| 四号 | 14 |
| 三号 | 16 |
| 二号 | 22 |

### 对齐方式

| 值 | 说明 | 适用场景 |
|---|------|---------|
| 0 | 左对齐 | 二级以下标题 |
| 1 | 居中 | 章标题、图表名 |
| 2 | 右对齐 | 落款、日期 |
| 3 | 两端对齐 | 正文 |