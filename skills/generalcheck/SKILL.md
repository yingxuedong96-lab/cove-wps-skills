---
name: generalcheck
description: >
  通用编号校对。对文档中的标题编号、图编号、表编号、公式编号进行自动校对与修正。
  触发词：校对编号、校对标题编号、校对图编号、校对表编号、校对公式编号、校对引用。
compatibility:
  runtime: WPS JS 宏 / WPS 加载项（JSAPI）
---

# 通用编号校对

## ⚠️ 执行流程

根据用户输入的触发词，直接执行对应 scope 的脚本：

### 触发词 → 参数映射表

| 触发词 | scope | figureFormat | tableFormat | 执行脚本 |
|-------|-------|--------------|-------------|---------|
| 校对编号、校对所有编号、编号校对 | `numbering` | `chapter` | `chapter` | scan-structure.js |
| 校对标题编号、标题编号校对 | `heading` | - | - | scan-structure.js |
| 校对图编号、图编号校对、校对图编号（章节式） | `figure` | `chapter` | - | scan-structure.js |
| 校对图编号（顺序式）、顺序式图编号 | `figure` | `simple` | - | scan-structure.js |
| 校对表编号、表编号校对、校对表编号（章节式） | `table` | - | `chapter` | scan-structure.js |
| 校对表编号（顺序式）、顺序式表编号 | `table` | - | `simple` | scan-structure.js |
| 校对公式编号、公式编号校对 | `formula` | - | - | scan-structure.js |
| 校对引用、引用校对 | `reference` | - | - | scan-structure.js |

---

## 执行方式

**⚠️ 直接调用脚本，不要做任何额外处理。**

```
executeFile:
  filePath: "skills/generalcheck/scripts/scan-structure.js"
  params: { "scope": "<匹配的scope>", "figureFormat": "<匹配的figureFormat>", "tableFormat": "<匹配的tableFormat>" }
```

**参数说明**：
- `scope`: 必需，校对范围（heading/figure/table/formula/numbering）
- `figureFormat`: 可选，仅 scope=figure 时有效
  - `chapter` = 章节式编号（图X.Y-Z），默认值
  - `simple` = 顺序式编号（图1、图2...）
- `tableFormat`: 可选，仅 scope=table 时有效
  - `chapter` = 章节式编号（表X.Y-Z），默认值
  - `simple` = 顺序式编号（表1、表2...）

---

## 规则说明

脚本自动识别并修复以下编号问题：

| 规则ID | 检查项 | 说明 |
|--------|--------|------|
| N-002 | 一级标题编号连续性 | 第一章→第三章，应为第二章 |
| N-003 | 二级标题章节号正确 | 3.1 在第二章下，应为 2.1 |
| N-004 | 二级标题节序号连续 | 2.1 → 2.3，应为 2.2 |
| N-005 | 三级标题编号 | 同一节内编号连续 |
| N-006 | 四级标题编号 | 同一条内编号连续 |
| N-008 | 五级标题编号 | 同一四级标题内编号连续 |
| N-007 | 附录标题编号格式 | 附录内应为 A1、A2 格式 |
| G-001 | 图编号（章节式） | 图X.Y-Z 格式，同一章节内连续 |
| G-002 | 图编号（顺序式） | 图1、图2... 全文递增 |
| G-003 | 附录图编号 | 图A1、图A2... 按附录内顺序 |
| T-001 | 表编号（章节式） | 表X.Y-Z 格式，同一章节内连续 |
| T-002 | 表编号（顺序式） | 表1、表2... 全文递增 |
| T-003 | 附录表编号 | 表A1、表A2... 按附录内顺序 |
| E-001 | 公式编号连续性 | (2-1) → (2-3)，应为 (2-2) |

---

## 技术实现

脚本使用单次遍历段落方式：
- 记录当前章号/节号上下文
- 识别旧格式编号并转换为规范格式
- 自动开启修订模式，所有修改可追溯

**输出格式**：
```json
{
  "success": true,
  "fixed": 50,
  "structure": { "headings": 20, "figures": 10, "tables": 15, "formulas": 5 },
  "details": [...]
}
```

---

## 注意事项

1. **编号校对需要全局上下文**，不能分片处理
2. **所有修复自动开启修订模式**，用户可接受或拒绝
3. **脚本已包含扫描+修复**，返回空 fixPlan，禁止再调用 apply-numbering-fixes.js