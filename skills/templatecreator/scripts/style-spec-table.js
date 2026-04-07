/**
 * 样式元素规范表 - 定义排版元素及其OOXML参数
 *
 * 用途：
 * 1. 解析时：从文档提取指定参数
 * 2. 匹配时：按参数特征识别元素类型
 * 3. 应用时：设置指定参数到文档
 *
 * 参数说明：
 * - w:rFonts -> fontCN, fontEN
 * - w:sz -> fontSize (半点)
 * - w:b -> bold
 * - w:i -> italic
 * - w:color -> color
 * - w:jc -> alignment
 * - w:ind -> firstLineIndent, leftIndent, hangingIndent
 * - w:spacing -> spaceBefore, spaceAfter, lineSpacing
 */

const ELEMENT_SPEC_TABLE = {

  // ==================== 公文元素（GB/T 9704-2012）====================
  official: {
    name: "党政机关公文",
    standard: "GB/T 9704-2012",

    elements: [
      // ------ 版头部分 ------
      {
        id: "issuer",
        name: "发文机关标志",
        category: "版头",
        description: "发文机关名称，如'XX市人民政府文件'",
        position: "版头红色区域",
        ooxml: {
          // 字体
          fontCN: { ooxml: "w:rFonts/@w:eastAsia", type: "string", description: "中文字体" },
          fontSize: { ooxml: "w:sz/@w:val", type: "number", unit: "half-pt", description: "字号" },
          color: { ooxml: "w:color/@w:val", type: "string", description: "颜色(RED=FF0000)" },
          alignment: { ooxml: "w:jc/@w:val", type: "enum", values: ["left", "center", "right"], description: "对齐" }
        }
      },
      {
        id: "dividerLine",
        name: "版头分隔线",
        category: "版头",
        description: "版头与主体之间的红色分隔线",
        ooxml: {
          // 这是一个图形元素，不是段落
          color: { type: "string", default: "FF0000" },
          width: { type: "number", unit: "pt" }
        }
      },
      {
        id: "docNumber",
        name: "发文字号",
        category: "版头",
        description: "如'国发〔2024〕1号'",
        detectPattern: "[\\d]{4}[号]|〔[\\d]{4}〕",
        ooxml: {
          fontCN: { ooxml: "w:rFonts/@w:eastAsia", type: "string" },
          fontSize: { ooxml: "w:sz/@w:val", type: "number", unit: "half-pt" },
          alignment: { ooxml: "w:jc/@w:val", type: "enum" }
        }
      },

      // ------ 主体部分 - 标题 ------
      {
        id: "docTitle",
        name: "公文标题",
        category: "主体",
        description: "公文主标题，如'关于XXX的通知'",
        position: "发文字号下方居中",
        ooxml: {
          fontCN: { ooxml: "w:rFonts/@w:eastAsia", type: "string" },
          fontEN: { ooxml: "w:rFonts/@w:ascii", type: "string" },
          fontSize: { ooxml: "w:sz/@w:val", type: "number", unit: "half-pt" },
          bold: { ooxml: "w:rPr/w:b", type: "boolean" },
          color: { ooxml: "w:color/@w:val", type: "string" },
          alignment: { ooxml: "w:jc/@w:val", type: "enum" },
          lineSpacing: { ooxml: "w:spacing/@w:line", type: "number", unit: "twips" }
        }
      },
      {
        id: "mainSender",
        name: "主送机关",
        category: "主体",
        description: "主要受文单位",
        position: "标题下左顶格",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum", default: "left" }
        }
      },
      {
        id: "heading1",
        name: "一级标题",
        category: "主体",
        description: "如'一、XXX'",
        detectPattern: "^[一二三四五六七八九十]+、",
        ooxml: {
          fontCN: { ooxml: "w:rFonts/@w:eastAsia", type: "string" },
          fontSize: { ooxml: "w:sz/@w:val", type: "number", unit: "half-pt" },
          bold: { ooxml: "w:rPr/w:b", type: "boolean" },
          alignment: { ooxml: "w:jc/@w:val", type: "enum" },
          firstLineIndent: { ooxml: "w:ind/@w:firstLine", type: "number", unit: "twips", description: "首行缩进" },
          spaceBefore: { ooxml: "w:spacing/@w:before", type: "number", unit: "twips" },
          spaceAfter: { ooxml: "w:spacing/@w:after", type: "number", unit: "twips" },
          lineSpacing: { ooxml: "w:spacing/@w:line", type: "number", unit: "twips" }
        }
      },
      {
        id: "heading2",
        name: "二级标题",
        category: "主体",
        description: "如'(一)XXX'",
        detectPattern: "^\\([一二三四五六七八九十]+\\)",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          bold: { type: "boolean" },
          alignment: { type: "enum" },
          firstLineIndent: { type: "number", unit: "twips" },
          spaceBefore: { type: "number", unit: "twips" },
          spaceAfter: { type: "number", unit: "twips" },
          lineSpacing: { type: "number", unit: "twips" }
        }
      },
      {
        id: "heading3",
        name: "三级标题",
        category: "主体",
        description: "如'1. XXX'",
        detectPattern: "^\\d+\\.\\s",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          bold: { type: "boolean" },
          alignment: { type: "enum" },
          firstLineIndent: { type: "number", unit: "twips" },
          spaceBefore: { type: "number", unit: "twips" },
          spaceAfter: { type: "number", unit: "twips" },
          lineSpacing: { type: "number", unit: "twips" }
        }
      },

      // ------ 主体部分 - 正文 ------
      {
        id: "body",
        name: "正文",
        category: "主体",
        description: "公文主体内容",
        isDefault: true,
        ooxml: {
          fontCN: { ooxml: "w:rFonts/@w:eastAsia", type: "string" },
          fontEN: { ooxml: "w:rFonts/@w:ascii", type: "string" },
          fontSize: { ooxml: "w:sz/@w:val", type: "number", unit: "half-pt" },
          bold: { ooxml: "w:rPr/w:b", type: "boolean" },
          italic: { ooxml: "w:rPr/w:i", type: "boolean" },
          alignment: { ooxml: "w:jc/@w:val", type: "enum" },
          firstLineIndent: { ooxml: "w:ind/@w:firstLine", type: "number", unit: "twips" },
          leftIndent: { ooxml: "w:ind/@w:left", type: "number", unit: "twips" },
          rightIndent: { ooxml: "w:ind/@w:right", type: "number", unit: "twips" },
          spaceBefore: { ooxml: "w:spacing/@w:before", type: "number", unit: "twips" },
          spaceAfter: { ooxml: "w:spacing/@w:after", type: "number", unit: "twips" },
          lineSpacing: { ooxml: "w:spacing/@w:line", type: "number", unit: "twips" },
          lineSpacingRule: { ooxml: "w:spacing/@w:lineRule", type: "enum", values: ["auto", "exact", "atLeast"] }
        }
      },
      {
        id: "listItem",
        name: "附件列表",
        category: "主体",
        description: "附件名称列表项",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" },
          leftIndent: { type: "number", unit: "twips" }
        }
      },

      // ------ 结尾部分 ------
      {
        id: "attachment",
        name: "附件说明",
        category: "结尾",
        description: "如'附件：1. XXX'",
        detectPattern: "^附件",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" }
        }
      },
      {
        id: "signature",
        name: "发文机关署名",
        category: "结尾",
        description: "落款单位",
        position: "正文右下方",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum", default: "right" }
        }
      },
      {
        id: "signDate",
        name: "成文日期",
        category: "结尾",
        description: "发文日期",
        detectPattern: "\\d{4}年\\d{1,2}月\\d{1,2}日",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum", default: "right" }
        }
      },
      {
        id: "sealPosition",
        name: "印章位置",
        category: "结尾",
        description: "印章盖章位置标记",
        ooxml: {
          // 印章是图形元素
          width: { type: "number", unit: "mm" },
          height: { type: "number", unit: "mm" }
        }
      },

      // ------ 版记部分 ------
      {
        id: "copySender",
        name: "抄送机关",
        category: "版记",
        description: "抄送单位列表",
        detectPattern: "^抄送",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" }
        }
      },
      {
        id: "issuerDept",
        name: "印发机关",
        category: "版记",
        description: "印发单位名称",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" }
        }
      },
      {
        id: "issueDate",
        name: "印发日期",
        category: "版记",
        description: "印发日期",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" }
        }
      },

      // ------ 页面元素 ------
      {
        id: "header",
        name: "页眉",
        category: "页面",
        description: "页眉内容",
        ooxmlPath: "w:hdr/w:p",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" }
        }
      },
      {
        id: "footer",
        name: "页码",
        category: "页面",
        description: "页码",
        ooxmlPath: "w:ftr/w:p",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" }
        }
      },
      {
        id: "pageSetup",
        name: "页面设置",
        category: "页面",
        description: "纸张、边距等",
        ooxmlPath: "w:sectPr",
        ooxml: {
          paperSize: { ooxml: "w:pgSz/@w:w,w:h", type: "enum", values: ["A4", "A3", "B5"] },
          orientation: { ooxml: "w:pgSz/@w:orient", type: "enum", values: ["portrait", "landscape"] },
          topMargin: { ooxml: "w:pgMar/@w:top", type: "number", unit: "twips" },
          bottomMargin: { ooxml: "w:pgMar/@w:bottom", type: "number", unit: "twips" },
          leftMargin: { ooxml: "w:pgMar/@w:left", type: "number", unit: "twips" },
          rightMargin: { ooxml: "w:pgMar/@w:right", type: "number", unit: "twips" },
          headerMargin: { ooxml: "w:pgMar/@w:header", type: "number", unit: "twips" },
          footerMargin: { ooxml: "w:pgMar/@w:footer", type: "number", unit: "twips" }
        }
      }
    ]
  },

  // ==================== 论文报告元素（Q/BIDR-G-JS00-101-002-2017）====================
  paper: {
    name: "技术报告",
    standard: "Q/BIDR-G-JS00-101-002-2017",

    elements: [
      // ------ 前置部分 ------
      {
        id: "coverTitle",
        name: "封面标题",
        category: "封面",
        description: "报告名称",
        ooxml: {
          fontCN: { ooxml: "w:rFonts/@w:eastAsia", type: "string" },
          fontSize: { ooxml: "w:sz/@w:val", type: "number", unit: "half-pt" },
          bold: { ooxml: "w:rPr/w:b", type: "boolean" },
          alignment: { ooxml: "w:jc/@w:val", type: "enum" }
        }
      },
      {
        id: "coverUnit",
        name: "封面单位",
        category: "封面",
        description: "编制单位",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" }
        }
      },
      {
        id: "coverDate",
        name: "封面日期",
        category: "封面",
        description: "发布日期",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" }
        }
      },
      {
        id: "titlePage",
        name: "扉页署名",
        category: "前置",
        description: "批准、审定、编写人员",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" }
        }
      },
      {
        id: "tocTitle",
        name: "目录标题",
        category: "前置",
        description: "'目录'二字",
        detectPattern: "^目\\s*录$",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" },
          spaceBefore: { type: "number", unit: "twips" },
          spaceAfter: { type: "number", unit: "twips" },
          lineSpacing: { type: "number", unit: "twips" }
        }
      },
      {
        id: "tocChapter",
        name: "目录章题",
        category: "前置",
        description: "目录一级条目",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          firstLineIndent: { type: "number", unit: "twips" }
        }
      },
      {
        id: "tocSection",
        name: "目录节题",
        category: "前置",
        description: "目录节条目",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          firstLineIndent: { type: "number", unit: "twips" }
        }
      },

      // ------ 主体 - 标题 ------
      {
        id: "chapterTitle",
        name: "章标题",
        category: "标题",
        description: "如'第一章'、'第1章'",
        detectPattern: "^第[一二三四五六七八九十\\d]+章",
        ooxml: {
          fontCN: { ooxml: "w:rFonts/@w:eastAsia", type: "string" },
          fontSize: { ooxml: "w:sz/@w:val", type: "number", unit: "half-pt" },
          bold: { ooxml: "w:rPr/w:b", type: "boolean" },
          alignment: { ooxml: "w:jc/@w:val", type: "enum" },
          spaceBefore: { ooxml: "w:spacing/@w:before", type: "number", unit: "twips" },
          spaceAfter: { ooxml: "w:spacing/@w:after", type: "number", unit: "twips" },
          lineSpacing: { ooxml: "w:spacing/@w:line", type: "number", unit: "twips" }
        }
      },
      {
        id: "heading1",
        name: "一级标题",
        category: "标题",
        description: "如'1 范围'",
        detectPattern: "^\\d+\\s+[^\\d\\.]",
        ooxml: {
          fontCN: { ooxml: "w:rFonts/@w:eastAsia", type: "string" },
          fontEN: { ooxml: "w:rFonts/@w:ascii", type: "string" },
          fontSize: { ooxml: "w:sz/@w:val", type: "number", unit: "half-pt" },
          bold: { ooxml: "w:rPr/w:b", type: "boolean" },
          alignment: { ooxml: "w:jc/@w:val", type: "enum" },
          firstLineIndent: { ooxml: "w:ind/@w:firstLine", type: "number", unit: "twips" },
          spaceBefore: { ooxml: "w:spacing/@w:before", type: "number", unit: "twips" },
          spaceAfter: { ooxml: "w:spacing/@w:after", type: "number", unit: "twips" },
          lineSpacing: { ooxml: "w:spacing/@w:line", type: "number", unit: "twips" },
          outlineLevel: { ooxml: "w:outlineLvl/@w:val", type: "number", description: "大纲级别" }
        }
      },
      {
        id: "heading2",
        name: "二级标题",
        category: "标题",
        description: "如'1.1 基本要求'",
        detectPattern: "^\\d+\\.\\d+\\s",
        ooxml: {
          fontCN: { type: "string" },
          fontEN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          bold: { type: "boolean" },
          alignment: { type: "enum" },
          firstLineIndent: { type: "number", unit: "twips" },
          spaceBefore: { type: "number", unit: "twips" },
          spaceAfter: { type: "number", unit: "twips" },
          lineSpacing: { type: "number", unit: "twips" },
          outlineLevel: { type: "number" }
        }
      },
      {
        id: "heading3",
        name: "三级标题",
        category: "标题",
        description: "如'1.1.1 内容'",
        detectPattern: "^\\d+\\.\\d+\\.\\d+\\s",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          bold: { type: "boolean" },
          alignment: { type: "enum" },
          firstLineIndent: { type: "number", unit: "twips" },
          spaceBefore: { type: "number", unit: "twips" },
          spaceAfter: { type: "number", unit: "twips" },
          lineSpacing: { type: "number", unit: "twips" },
          outlineLevel: { type: "number" }
        }
      },
      {
        id: "heading4",
        name: "四级标题",
        category: "标题",
        description: "如'1.1.1.1'",
        detectPattern: "^\\d+\\.\\d+\\.\\d+\\.\\d+\\s",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          bold: { type: "boolean" },
          alignment: { type: "enum" },
          firstLineIndent: { type: "number", unit: "twips" },
          spaceBefore: { type: "number", unit: "twips" },
          spaceAfter: { type: "number", unit: "twips" },
          lineSpacing: { type: "number", unit: "twips" }
        }
      },
      {
        id: "heading5",
        name: "五级标题",
        category: "标题",
        description: "如'(1)'、'(a)'",
        detectPattern: "^\\(\\d+\\)|^\\([a-z]\\)|^[①②③④⑤]",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          bold: { type: "boolean" },
          alignment: { type: "enum" },
          firstLineIndent: { type: "number", unit: "twips" },
          spaceBefore: { type: "number", unit: "twips" },
          spaceAfter: { type: "number", unit: "twips" },
          lineSpacing: { type: "number", unit: "twips" }
        }
      },

      // ------ 主体 - 正文 ------
      {
        id: "body",
        name: "正文",
        category: "正文",
        description: "正文段落",
        isDefault: true,
        ooxml: {
          fontCN: { ooxml: "w:rFonts/@w:eastAsia", type: "string" },
          fontEN: { ooxml: "w:rFonts/@w:ascii", type: "string" },
          fontSize: { ooxml: "w:sz/@w:val", type: "number", unit: "half-pt" },
          bold: { ooxml: "w:rPr/w:b", type: "boolean" },
          italic: { ooxml: "w:rPr/w:i", type: "boolean" },
          color: { ooxml: "w:color/@w:val", type: "string" },
          alignment: { ooxml: "w:jc/@w:val", type: "enum" },
          firstLineIndent: { ooxml: "w:ind/@w:firstLine", type: "number", unit: "twips" },
          leftIndent: { ooxml: "w:ind/@w:left", type: "number", unit: "twips" },
          rightIndent: { ooxml: "w:ind/@w:right", type: "number", unit: "twips" },
          hangingIndent: { ooxml: "w:ind/@w:hanging", type: "number", unit: "twips" },
          spaceBefore: { ooxml: "w:spacing/@w:before", type: "number", unit: "twips" },
          spaceAfter: { ooxml: "w:spacing/@w:after", type: "number", unit: "twips" },
          lineSpacing: { ooxml: "w:spacing/@w:line", type: "number", unit: "twips" },
          lineSpacingRule: { ooxml: "w:spacing/@w:lineRule", type: "enum" }
        }
      },
      {
        id: "listItem",
        name: "列表项",
        category: "正文",
        description: "编号列表或项目符号",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" },
          firstLineIndent: { type: "number", unit: "twips" },
          leftIndent: { type: "number", unit: "twips" },
          lineSpacing: { type: "number", unit: "twips" },
          numId: { ooxml: "w:numPr/w:numId/@w:val", type: "number", description: "列表编号ID" },
          ilvl: { ooxml: "w:numPr/w:ilvl/@w:val", type: "number", description: "列表级别" }
        }
      },

      // ------ 主体 - 图表公式 ------
      {
        id: "figureCaption",
        name: "图名",
        category: "图表",
        description: "如'图 1.2-1 系统架构'",
        detectPattern: "^图\\s*\\d+",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" },
          spaceBefore: { type: "number", unit: "twips" },
          spaceAfter: { type: "number", unit: "twips" },
          lineSpacing: { type: "number", unit: "twips" }
        }
      },
      {
        id: "tableCaption",
        name: "表名",
        category: "图表",
        description: "如'表 1.2-1 技术参数'",
        detectPattern: "^表\\s*\\d+",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" },
          spaceBefore: { type: "number", unit: "twips" },
          spaceAfter: { type: "number", unit: "twips" },
          lineSpacing: { type: "number", unit: "twips" }
        }
      },
      {
        id: "tableBody",
        name: "表内文字",
        category: "图表",
        description: "表格单元格文字",
        ooxmlPath: "w:tbl/w:tr/w:tc/w:p",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" },
          verticalAlign: { ooxml: "w:vAlign/@w:val", type: "enum", values: ["top", "center", "bottom"] }
        }
      },
      {
        id: "tableBorder",
        name: "表格边框",
        category: "图表",
        description: "表格线框样式",
        ooxmlPath: "w:tbl/w:tblPr/w:tblBorders",
        ooxml: {
          top: { ooxml: "w:top", type: "object" },
          bottom: { ooxml: "w:bottom", type: "object" },
          left: { ooxml: "w:left", type: "object" },
          right: { ooxml: "w:right", type: "object" },
          insideH: { ooxml: "w:insideH", type: "object", description: "内部横线" },
          insideV: { ooxml: "w:insideV", type: "object", description: "内部竖线" }
        }
      },
      {
        id: "formulaCaption",
        name: "公式编号",
        category: "图表",
        description: "如'(1)'、'(3.2.1-1)'",
        detectPattern: "^\\([\\d.-]+\\)$",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum", default: "right" }
        }
      },
      {
        id: "formulaNote",
        name: "公式说明",
        category: "图表",
        description: "'式中'符号说明",
        detectPattern: "^式中",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" }
        }
      },

      // ------ 补充部分 ------
      {
        id: "appendixTitle",
        name: "附录标题",
        category: "附录",
        description: "如'附录 A'",
        detectPattern: "^附录\\s*[A-Z]",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          bold: { type: "boolean" },
          alignment: { type: "enum" }
        }
      },
      {
        id: "appendixSection",
        name: "附录节题",
        category: "附录",
        description: "如'A.1 详细说明'",
        detectPattern: "^[A-Z]\\.\\d+\\s",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          bold: { type: "boolean" },
          alignment: { type: "enum" }
        }
      },
      {
        id: "referenceTitle",
        name: "参考文献标题",
        category: "附录",
        description: "'参考文献'",
        detectPattern: "^参考文献",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          bold: { type: "boolean" },
          alignment: { type: "enum" }
        }
      },
      {
        id: "reference",
        name: "参考文献条目",
        category: "附录",
        description: "文献引用条目",
        detectPattern: "^\\[\\d+\\]",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" },
          firstLineIndent: { type: "number", unit: "twips" },
          leftIndent: { type: "number", unit: "twips" },
          hangingIndent: { type: "number", unit: "twips" }
        }
      },
      {
        id: "note",
        name: "注释说明",
        category: "附录",
        description: "脚注、尾注",
        detectPattern: "^注\\s*\\d*",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" }
        }
      },

      // ------ 页面元素 ------
      {
        id: "header",
        name: "页眉",
        category: "页面",
        description: "页眉内容",
        ooxmlPath: "w:hdr/w:p",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" }
        }
      },
      {
        id: "footer",
        name: "页脚",
        category: "页面",
        description: "页码",
        ooxmlPath: "w:ftr/w:p",
        ooxml: {
          fontCN: { type: "string" },
          fontSize: { type: "number", unit: "half-pt" },
          alignment: { type: "enum" }
        }
      },
      {
        id: "pageSetup",
        name: "页面设置",
        category: "页面",
        description: "纸张、边距等",
        ooxmlPath: "w:sectPr",
        ooxml: {
          paperSize: { ooxml: "w:pgSz/@w:w,w:h", type: "enum" },
          orientation: { ooxml: "w:pgSz/@w:orient", type: "enum" },
          topMargin: { ooxml: "w:pgMar/@w:top", type: "number", unit: "twips" },
          bottomMargin: { ooxml: "w:pgMar/@w:bottom", type: "number", unit: "twips" },
          leftMargin: { ooxml: "w:pgMar/@w:left", type: "number", unit: "twips" },
          rightMargin: { ooxml: "w:pgMar/@w:right", type: "number", unit: "twips" },
          headerMargin: { ooxml: "w:pgMar/@w:header", type: "number", unit: "twips" },
          footerMargin: { ooxml: "w:pgMar/@w:footer", type: "number", unit: "twips" }
        }
      }
    ]
  }
};

// OOXML单位转换
const UNIT_CONVERT = {
  // twips <-> 其他单位
  twips: {
    toPt: (twips) => twips / 20,
    toMm: (twips) => twips / 56.692,
    toCm: (twips) => twips / 566.92,
    toInch: (twips) => twips / 1440,
    fromPt: (pt) => pt * 20,
    fromMm: (mm) => mm * 56.692,
    fromCm: (cm) => cm * 566.92
  },
  // half-pt (w:sz) <-> pt
  halfPt: {
    toPt: (halfPt) => halfPt / 2,
    fromPt: (pt) => pt * 2
  },
  // 字符数 <-> twips (假设标准字号12pt)
  char: {
    toTwips: (chars, fontSize = 12) => chars * fontSize * 20,
    fromTwips: (twips, fontSize = 12) => twips / (fontSize * 20)
  }
};

// 导出
if (typeof module !== 'undefined') {
  module.exports = { ELEMENT_SPEC_TABLE, UNIT_CONVERT };
}