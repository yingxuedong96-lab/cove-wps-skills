/**
 * 样式元素规范表 - 定义排版元素及其参数（不含默认值）
 *
 * 用途：
 * 1. 作为提取样式的"检查清单"，确保不漏掉任何参数
 * 2. 定义每个元素需要提取哪些OOXML参数
 * 3. 提供检测模式用于识别元素类型
 *
 * 注意：此表不包含默认值，所有参数值从文档中提取
 */

const ELEMENT_SPEC_TABLE = {

  // ==================== 公文元素（GB/T 9704-2012）====================
  official: {
    name: "党政机关公文",
    standard: "GB/T 9704-2012",
    elementCount: 20,

    elements: [
      // ====== 版头部分 ======
      {
        id: "issuer",
        name: "发文机关标志",
        category: "版头",
        description: "发文机关名称，如'XX市人民政府文件'",
        position: "版头红色区域",
        detectHint: "位于文档开头，红色大字",
        params: ["fontCN", "fontEN", "fontSize", "bold", "color", "alignment"]
      },
      {
        id: "dividerLine",
        name: "版头分隔线",
        category: "版头",
        description: "版头与主体之间的分隔线",
        position: "版头底部",
        params: ["color", "lineWidth", "lineStyle"]
      },
      {
        id: "docNumber",
        name: "发文字号",
        category: "版头",
        description: "如'国发〔2024〕1号'",
        detectPattern: "[\\d]{4}[号]|〔[\\d]{4}〕",
        params: ["fontCN", "fontEN", "fontSize", "bold", "color", "alignment"]
      },

      // ====== 主体部分 - 标题 ======
      {
        id: "docTitle",
        name: "公文标题",
        category: "主体",
        description: "公文主标题，如'关于XXX的通知'",
        position: "发文字号下方",
        detectHint: "字号较大，居中",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "color", "alignment", "spaceBefore", "spaceAfter", "lineSpacing", "lineSpacingRule"]
      },
      {
        id: "mainSender",
        name: "主送机关",
        category: "主体",
        description: "主要受文单位",
        position: "标题下方",
        detectHint: "左侧顶格，后接冒号",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },
      {
        id: "heading1",
        name: "一级标题",
        category: "主体",
        description: "如'一、XXX'",
        detectPattern: "^[一二三四五六七八九十]+、",
        detectHint: "汉字数字加顿号",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "color", "alignment", "firstLineIndent", "leftIndent", "spaceBefore", "spaceAfter", "lineSpacing", "lineSpacingRule"]
      },
      {
        id: "heading2",
        name: "二级标题",
        category: "主体",
        description: "如'(一)XXX'",
        detectPattern: "^\\([一二三四五六七八九十]+\\)",
        detectHint: "括号加汉字数字",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "color", "alignment", "firstLineIndent", "leftIndent", "spaceBefore", "spaceAfter", "lineSpacing", "lineSpacingRule"]
      },
      {
        id: "heading3",
        name: "三级标题",
        category: "主体",
        description: "如'1. XXX'",
        detectPattern: "^\\d+\\.\\s",
        detectHint: "阿拉伯数字加点",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "color", "alignment", "firstLineIndent", "leftIndent", "spaceBefore", "spaceAfter", "lineSpacing", "lineSpacingRule"]
      },

      // ====== 主体部分 - 正文 ======
      {
        id: "body",
        name: "正文",
        category: "主体",
        description: "公文主体内容",
        isDefault: true,
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "underline", "color", "alignment", "firstLineIndent", "leftIndent", "rightIndent", "hangingIndent", "spaceBefore", "spaceAfter", "lineSpacing", "lineSpacingRule"]
      },
      {
        id: "listItem",
        name: "附件列表",
        category: "主体",
        description: "附件名称列表项",
        detectHint: "列表编号或符号开头",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment", "firstLineIndent", "leftIndent", "lineSpacing", "numId", "ilvl"]
      },

      // ====== 结尾部分 ======
      {
        id: "attachment",
        name: "附件说明",
        category: "结尾",
        description: "如'附件：1. XXX'",
        detectPattern: "^附件",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },
      {
        id: "signature",
        name: "发文机关署名",
        category: "结尾",
        description: "落款单位名称",
        position: "正文右下方",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },
      {
        id: "signDate",
        name: "成文日期",
        category: "结尾",
        description: "发文日期",
        detectPattern: "\\d{4}年\\d{1,2}月\\d{1,2}日",
        position: "署名下方",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },
      {
        id: "sealPosition",
        name: "印章位置",
        category: "结尾",
        description: "印章盖章位置标记",
        params: ["width", "height", "offsetX", "offsetY"]
      },

      // ====== 版记部分 ======
      {
        id: "copySender",
        name: "抄送机关",
        category: "版记",
        description: "抄送单位列表",
        detectPattern: "^抄送",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },
      {
        id: "issuerDept",
        name: "印发机关",
        category: "版记",
        description: "印发单位名称",
        position: "版记左下",
        params: ["fontCN", "fontEN", "fontSize", "bold"]
      },
      {
        id: "issueDate",
        name: "印发日期",
        category: "版记",
        description: "印发日期",
        position: "版记右下",
        params: ["fontCN", "fontEN", "fontSize", "bold"]
      },

      // ====== 页面元素 ======
      {
        id: "header",
        name: "页眉",
        category: "页面",
        description: "页眉内容",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment", "lineWidth"]
      },
      {
        id: "footer",
        name: "页码",
        category: "页面",
        description: "页码",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },
      {
        id: "pageSetup",
        name: "页面设置",
        category: "页面",
        description: "纸张、边距等",
        params: ["paperSize", "orientation", "topMargin", "bottomMargin", "leftMargin", "rightMargin", "headerMargin", "footerMargin", "gutter"]
      }
    ]
  },

  // ==================== 论文报告元素（Q/BIDR-G-JS00-101-002-2017）====================
  paper: {
    name: "技术报告",
    standard: "Q/BIDR-G-JS00-101-002-2017",
    elementCount: 26,

    elements: [
      // ====== 前置部分 ======
      {
        id: "coverTitle",
        name: "封面标题",
        category: "封面",
        description: "报告名称",
        position: "封面上部居中",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "color", "alignment"]
      },
      {
        id: "coverSubtitle",
        name: "封面副标题",
        category: "封面",
        description: "报告副标题或英文名称",
        position: "标题下方",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "alignment"]
      },
      {
        id: "coverUnit",
        name: "封面单位",
        category: "封面",
        description: "编制单位名称",
        position: "封面下部",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },
      {
        id: "coverDate",
        name: "封面日期",
        category: "封面",
        description: "发布日期",
        position: "封面底部",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },
      {
        id: "titlePage",
        name: "扉页署名",
        category: "前置",
        description: "批准、审定、编写人员署名",
        position: "封面后",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },
      {
        id: "tocTitle",
        name: "目录标题",
        category: "前置",
        description: "'目录'二字",
        detectPattern: "^目\\s*录$|^目次$",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment", "spaceBefore", "spaceAfter", "lineSpacing"]
      },
      {
        id: "tocChapter",
        name: "目录章题",
        category: "前置",
        description: "目录一级条目",
        params: ["fontCN", "fontEN", "fontSize", "bold", "firstLineIndent", "tabStop"]
      },
      {
        id: "tocSection",
        name: "目录节题",
        category: "前置",
        description: "目录节条目",
        params: ["fontCN", "fontEN", "fontSize", "bold", "firstLineIndent", "tabStop"]
      },

      // ====== 主体 - 标题 ======
      {
        id: "chapterTitle",
        name: "章标题",
        category: "标题",
        description: "如'第一章 范围'、'第1章'",
        detectPattern: "^第[一二三四五六七八九十\\d]+章",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "color", "alignment", "firstLineIndent", "spaceBefore", "spaceAfter", "lineSpacing", "outlineLevel"]
      },
      {
        id: "heading1",
        name: "一级标题",
        category: "标题",
        description: "如'1 范围'",
        detectPattern: "^\\d+\\s+[^\\d\\.]",
        detectHint: "阿拉伯数字开头加空格",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "color", "alignment", "firstLineIndent", "leftIndent", "spaceBefore", "spaceAfter", "lineSpacing", "lineSpacingRule", "outlineLevel"]
      },
      {
        id: "heading2",
        name: "二级标题",
        category: "标题",
        description: "如'1.1 基本要求'",
        detectPattern: "^\\d+\\.\\d+\\s",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "color", "alignment", "firstLineIndent", "leftIndent", "spaceBefore", "spaceAfter", "lineSpacing", "lineSpacingRule", "outlineLevel"]
      },
      {
        id: "heading3",
        name: "三级标题",
        category: "标题",
        description: "如'1.1.1 内容'",
        detectPattern: "^\\d+\\.\\d+\\.\\d+\\s",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "color", "alignment", "firstLineIndent", "leftIndent", "spaceBefore", "spaceAfter", "lineSpacing", "lineSpacingRule", "outlineLevel"]
      },
      {
        id: "heading4",
        name: "四级标题",
        category: "标题",
        description: "如'1.1.1.1'",
        detectPattern: "^\\d+\\.\\d+\\.\\d+\\.\\d+\\s",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "color", "alignment", "firstLineIndent", "leftIndent", "spaceBefore", "spaceAfter", "lineSpacing", "lineSpacingRule"]
      },
      {
        id: "heading5",
        name: "五级标题",
        category: "标题",
        description: "如'(1)'、'(a)'、'①'",
        detectPattern: "^\\(\\d+\\)|^\\([a-z]\\)|^[①②③④⑤⑥⑦⑧⑨⑩]",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "alignment", "firstLineIndent", "leftIndent", "spaceBefore", "spaceAfter", "lineSpacing"]
      },

      // ====== 主体 - 正文 ======
      {
        id: "body",
        name: "正文",
        category: "正文",
        description: "正文段落",
        isDefault: true,
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "underline", "color", "alignment", "firstLineIndent", "leftIndent", "rightIndent", "hangingIndent", "spaceBefore", "spaceAfter", "lineSpacing", "lineSpacingRule"]
      },
      {
        id: "listItem",
        name: "列表项",
        category: "正文",
        description: "编号列表或项目符号",
        detectHint: "列表编号或符号",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "alignment", "firstLineIndent", "leftIndent", "lineSpacing", "numId", "ilvl"]
      },

      // ====== 主体 - 图表公式 ======
      {
        id: "figureCaption",
        name: "图名",
        category: "图表",
        description: "如'图 1.2-1 系统架构'",
        detectPattern: "^图\\s*\\d+",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "alignment", "spaceBefore", "spaceAfter", "lineSpacing"]
      },
      {
        id: "tableCaption",
        name: "表名",
        category: "图表",
        description: "如'表 1.2-1 技术参数'",
        detectPattern: "^表\\s*\\d+",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "alignment", "spaceBefore", "spaceAfter", "lineSpacing"]
      },
      {
        id: "tableBody",
        name: "表内文字",
        category: "图表",
        description: "表格单元格文字",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "alignment", "verticalAlign", "cellPadding"]
      },
      {
        id: "tableBorder",
        name: "表格边框",
        category: "图表",
        description: "表格线框样式",
        params: ["borderTop", "borderBottom", "borderLeft", "borderRight", "borderInsideH", "borderInsideV", "borderStyle", "borderWidth", "borderColor"]
      },
      {
        id: "formulaCaption",
        name: "公式编号",
        category: "图表",
        description: "如'(1)'、'(3.2.1-1)'",
        detectPattern: "^\\([\\d.-]+\\)$",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },
      {
        id: "formulaNote",
        name: "公式说明",
        category: "图表",
        description: "'式中'符号说明",
        detectPattern: "^式中",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },

      // ====== 补充部分 ======
      {
        id: "appendixTitle",
        name: "附录标题",
        category: "附录",
        description: "如'附录 A'",
        detectPattern: "^附录\\s*[A-Z]",
        params: ["fontCN", "fontEN", "fontSize", "bold", "italic", "alignment"]
      },
      {
        id: "appendixSection",
        name: "附录节题",
        category: "附录",
        description: "如'A.1 详细说明'",
        detectPattern: "^[A-Z]\\.\\d+\\s",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },
      {
        id: "referenceTitle",
        name: "参考文献标题",
        category: "附录",
        description: "'参考文献'",
        detectPattern: "^参考文献",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },
      {
        id: "reference",
        name: "参考文献条目",
        category: "附录",
        description: "文献引用条目",
        detectPattern: "^\\[\\d+\\]",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment", "firstLineIndent", "leftIndent", "hangingIndent"]
      },
      {
        id: "note",
        name: "注释说明",
        category: "附录",
        description: "脚注、尾注",
        detectPattern: "^注\\s*\\d*",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment"]
      },

      // ====== 页面元素 ======
      {
        id: "header",
        name: "页眉",
        category: "页面",
        description: "页眉内容",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment", "lineWidth", "lineStyle"]
      },
      {
        id: "footer",
        name: "页脚",
        category: "页面",
        description: "页码",
        params: ["fontCN", "fontEN", "fontSize", "bold", "alignment", "lineWidth", "lineStyle"]
      },
      {
        id: "pageSetup",
        name: "页面设置",
        category: "页面",
        description: "纸张、边距等",
        params: ["paperSize", "orientation", "topMargin", "bottomMargin", "leftMargin", "rightMargin", "headerMargin", "footerMargin", "gutter"]
      }
    ]
  }
};

// ==================== 参数定义表 ====================
const PARAM_SPEC = {
  // ====== 字体相关 ======
  fontCN: {
    name: "中文字体",
    ooxml: "w:rFonts/@w:eastAsia",
    type: "string",
    description: "段落的中文字体名称"
  },
  fontEN: {
    name: "英文字体",
    ooxml: "w:rFonts/@w:ascii",
    type: "string",
    description: "段落的英文字体名称"
  },
  fontSize: {
    name: "字号",
    ooxml: "w:sz/@w:val",
    type: "number",
    unit: "half-pt",
    description: "字号，单位为半点（22pt = 44）"
  },
  bold: {
    name: "加粗",
    ooxml: "w:rPr/w:b",
    type: "boolean",
    description: "是否加粗"
  },
  italic: {
    name: "斜体",
    ooxml: "w:rPr/w:i",
    type: "boolean",
    description: "是否斜体"
  },
  underline: {
    name: "下划线",
    ooxml: "w:rPr/w:u/@w:val",
    type: "enum",
    values: ["none", "single", "double", "thick"],
    description: "下划线类型"
  },
  color: {
    name: "颜色",
    ooxml: "w:color/@w:val",
    type: "string",
    description: "字体颜色，十六进制如FF0000"
  },

  // ====== 段落格式 ======
  alignment: {
    name: "对齐方式",
    ooxml: "w:jc/@w:val",
    type: "enum",
    values: ["left", "center", "right", "justify", "distribute"],
    description: "段落对齐方式"
  },
  firstLineIndent: {
    name: "首行缩进",
    ooxml: "w:ind/@w:firstLine",
    type: "number",
    unit: "twips",
    description: "首行缩进量（1字符 ≈ 240 twips）"
  },
  leftIndent: {
    name: "左缩进",
    ooxml: "w:ind/@w:left",
    type: "number",
    unit: "twips",
    description: "左缩进量"
  },
  rightIndent: {
    name: "右缩进",
    ooxml: "w:ind/@w:right",
    type: "number",
    unit: "twips",
    description: "右缩进量"
  },
  hangingIndent: {
    name: "悬挂缩进",
    ooxml: "w:ind/@w:hanging",
    type: "number",
    unit: "twips",
    description: "悬挂缩进量（除首行外的缩进）"
  },

  // ====== 间距 ======
  spaceBefore: {
    name: "段前间距",
    ooxml: "w:spacing/@w:before",
    type: "number",
    unit: "twips",
    description: "段前间距"
  },
  spaceAfter: {
    name: "段后间距",
    ooxml: "w:spacing/@w:after",
    type: "number",
    unit: "twips",
    description: "段后间距"
  },
  lineSpacing: {
    name: "行距",
    ooxml: "w:spacing/@w:line",
    type: "number",
    unit: "twips",
    description: "行距值"
  },
  lineSpacingRule: {
    name: "行距规则",
    ooxml: "w:spacing/@w:lineRule",
    type: "enum",
    values: ["auto", "exact", "atLeast"],
    description: "auto=单倍行距的倍数，exact=固定值，atLeast=最小值"
  },

  // ====== 大纲/列表 ======
  outlineLevel: {
    name: "大纲级别",
    ooxml: "w:outlineLvl/@w:val",
    type: "number",
    description: "文档结构中的级别（0=正文，1=一级标题...）"
  },
  numId: {
    name: "列表编号ID",
    ooxml: "w:numPr/w:numId/@w:val",
    type: "number",
    description: "列表样式的ID"
  },
  ilvl: {
    name: "列表级别",
    ooxml: "w:numPr/w:ilvl/@w:val",
    type: "number",
    description: "列表的级别（0开始）"
  },

  // ====== 表格相关 ======
  verticalAlign: {
    name: "垂直对齐",
    ooxml: "w:vAlign/@w:val",
    type: "enum",
    values: ["top", "center", "bottom"],
    description: "单元格内垂直对齐"
  },
  cellPadding: {
    name: "单元格边距",
    type: "number",
    unit: "twips",
    description: "单元格内边距"
  },
  borderTop: { name: "上边框", type: "object" },
  borderBottom: { name: "下边框", type: "object" },
  borderLeft: { name: "左边框", type: "object" },
  borderRight: { name: "右边框", type: "object" },
  borderInsideH: { name: "内部横线", type: "object" },
  borderInsideV: { name: "内部竖线", type: "object" },
  borderStyle: { name: "边框样式", type: "enum", values: ["none", "single", "double"] },
  borderWidth: { name: "边框宽度", type: "number", unit: "pt" },
  borderColor: { name: "边框颜色", type: "string" },

  // ====== 页面设置 ======
  paperSize: {
    name: "纸张大小",
    ooxml: "w:pgSz/@w:w,w:h",
    type: "enum",
    values: ["A4", "A3", "B5", "Letter", "Legal"],
    description: "纸张规格"
  },
  orientation: {
    name: "纸张方向",
    ooxml: "w:pgSz/@w:orient",
    type: "enum",
    values: ["portrait", "landscape"],
    description: "纵向或横向"
  },
  topMargin: {
    name: "上边距",
    ooxml: "w:pgMar/@w:top",
    type: "number",
    unit: "twips",
    description: "页面上边距"
  },
  bottomMargin: {
    name: "下边距",
    ooxml: "w:pgMar/@w:bottom",
    type: "number",
    unit: "twips",
    description: "页面下边距"
  },
  leftMargin: {
    name: "左边距",
    ooxml: "w:pgMar/@w:left",
    type: "number",
    unit: "twips",
    description: "页面左边距"
  },
  rightMargin: {
    name: "右边距",
    ooxml: "w:pgMar/@w:right",
    type: "number",
    unit: "twips",
    description: "页面右边距"
  },
  headerMargin: {
    name: "页眉边距",
    ooxml: "w:pgMar/@w:header",
    type: "number",
    unit: "twips",
    description: "页眉距页面顶部的距离"
  },
  footerMargin: {
    name: "页脚边距",
    ooxml: "w:pgMar/@w:footer",
    type: "number",
    unit: "twips",
    description: "页脚距页面底部的距离"
  },
  gutter: {
    name: "装订线",
    ooxml: "w:pgMar/@w:gutter",
    type: "number",
    unit: "twips",
    description: "装订线宽度"
  },

  // ====== 其他 ======
  lineWidth: { name: "线条宽度", type: "number", unit: "pt" },
  lineStyle: { name: "线条样式", type: "enum", values: ["none", "single", "double"] },
  tabStop: { name: "制表位", type: "array", description: "制表位设置" },
  width: { name: "宽度", type: "number", unit: "mm" },
  height: { name: "高度", type: "number", unit: "mm" },
  offsetX: { name: "X偏移", type: "number", unit: "mm" },
  offsetY: { name: "Y偏移", type: "number", unit: "mm" }
};

// ==================== 单位转换 ====================
const UNIT_CONVERT = {
  twips: {
    toPt: (v) => v / 20,
    toMm: (v) => v / 56.692,
    toCm: (v) => v / 566.92,
    fromPt: (v) => v * 20,
    fromMm: (v) => v * 56.692,
    fromCm: (v) => v * 566.92
  },
  halfPt: {
    toPt: (v) => v / 2,
    fromPt: (v) => v * 2
  },
  char: {
    toTwips: (chars, fontSize = 12) => chars * fontSize * 20,
    fromTwips: (twips, fontSize = 12) => twips / (fontSize * 20)
  }
};

// 导出
if (typeof module !== 'undefined') {
  module.exports = { ELEMENT_SPEC_TABLE, PARAM_SPEC, UNIT_CONVERT };
}