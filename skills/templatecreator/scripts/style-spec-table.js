/**
 * 样式规范表 - 定义所有支持的标签及其属性
 *
 * 分为两大类：公文样式、论文样式
 */

const STYLE_SPEC_TABLE = {
  // ========== 公文样式 ==========
  official: {
    name: "公文样式",
    tags: [
      {
        id: "docTitle",
        name: "公文标题",
        category: "标题类",
        description: "公文的主标题，如'关于XXX的通知'",
        detectHint: "通常位于文号下方，字号较大居中",
        properties: {
          fontCN: { type: "string", required: true, default: "方正小标宋简体", description: "中文字体" },
          fontEN: { type: "string", required: false, description: "英文字体" },
          fontSize: { type: "number", required: true, default: 22, unit: "pt", description: "字号" },
          bold: { type: "boolean", default: false, description: "加粗" },
          color: { type: "string", default: "auto", description: "颜色" },
          alignment: { type: "enum", values: ["left", "center", "right", "justify"], default: "center", description: "对齐" },
          lineSpacing: { type: "number", default: 28, unit: "pt", description: "行距" }
        }
      },
      {
        id: "docNumber",
        name: "发文字号",
        category: "标识类",
        description: "如'国发〔2024〕1号'",
        detectHint: "位于标题上方或下方，包含年份和序号",
        properties: {
          fontCN: { type: "string", default: "仿宋" },
          fontSize: { type: "number", default: 16, unit: "pt" },
          alignment: { type: "enum", default: "center" },
          color: { type: "string", default: "red" }
        }
      },
      {
        id: "issuer",
        name: "发文机关",
        category: "标识类",
        description: "发文单位名称",
        detectHint: "位于文号上方或红头部分",
        properties: {
          fontCN: { type: "string", default: "方正小标宋简体" },
          fontSize: { type: "number", default: 22 },
          alignment: { type: "enum", default: "center" },
          color: { type: "string", default: "red" }
        }
      },
      {
        id: "heading1",
        name: "一级标题",
        category: "标题类",
        description: "如'一、XXX'",
        detectHint: "汉字数字开头加顿号",
        properties: {
          fontCN: { type: "string", default: "黑体" },
          fontSize: { type: "number", default: 16 },
          bold: { type: "boolean", default: false },
          alignment: { type: "enum", default: "left" },
          firstLineIndent: { type: "number", default: 0, unit: "char", description: "首行缩进(字符数)" }
        }
      },
      {
        id: "heading2",
        name: "二级标题",
        category: "标题类",
        description: "如'(一)XXX'",
        detectHint: "括号加汉字数字",
        properties: {
          fontCN: { type: "string", default: "楷体" },
          fontSize: { type: "number", default: 16 },
          bold: { type: "boolean", default: false },
          alignment: { type: "enum", default: "left" },
          firstLineIndent: { type: "number", default: 0 }
        }
      },
      {
        id: "heading3",
        name: "三级标题",
        category: "标题类",
        description: "如'1. XXX'",
        detectHint: "阿拉伯数字加点",
        properties: {
          fontCN: { type: "string", default: "仿宋" },
          fontSize: { type: "number", default: 16 },
          bold: { type: "boolean", default: true },
          alignment: { type: "enum", default: "left" },
          firstLineIndent: { type: "number", default: 0 }
        }
      },
      {
        id: "body",
        name: "正文",
        category: "正文类",
        description: "公文主体内容",
        detectHint: "默认段落类型",
        properties: {
          fontCN: { type: "string", default: "仿宋" },
          fontSize: { type: "number", default: 16, unit: "pt" },
          alignment: { type: "enum", default: "justify" },
          firstLineIndent: { type: "number", default: 2, unit: "char" },
          lineSpacing: { type: "number", default: 28, unit: "pt" }
        }
      },
      {
        id: "attachment",
        name: "附件说明",
        category: "附属类",
        description: "如'附件：1. XXX'",
        detectHint: "正文末尾，'附件'二字开头",
        properties: {
          fontCN: { type: "string", default: "仿宋" },
          fontSize: { type: "number", default: 16 },
          alignment: { type: "enum", default: "left" },
          firstLineIndent: { type: "number", default: 2 }
        }
      },
      {
        id: "signature",
        name: "落款",
        category: "结尾类",
        description: "发文机关署名和日期",
        detectHint: "正文右下方",
        properties: {
          fontCN: { type: "string", default: "仿宋" },
          fontSize: { type: "number", default: 16 },
          alignment: { type: "enum", default: "right" }
        }
      },
      {
        id: "sealPosition",
        name: "印章位置",
        category: "结尾类",
        description: "印章盖章位置标记",
        detectHint: "落款下方或骑缝",
        properties: {}
      },
      {
        id: "header",
        name: "页眉",
        category: "页面类",
        description: "页眉内容",
        properties: {
          fontCN: { type: "string", default: "宋体" },
          fontSize: { type: "number", default: 10.5 },
          alignment: { type: "enum", default: "center" }
        }
      },
      {
        id: "footer",
        name: "页脚",
        category: "页面类",
        description: "页码等",
        properties: {
          fontCN: { type: "string", default: "宋体" },
          fontSize: { type: "number", default: 10.5 },
          alignment: { type: "enum", default: "center" }
        }
      }
    ]
  },

  // ========== 论文/报告样式 ==========
  paper: {
    name: "论文报告样式",
    tags: [
      {
        id: "docTitle",
        name: "论文标题",
        category: "标题类",
        description: "论文/报告的主标题",
        detectHint: "文档首段，字号最大居中",
        properties: {
          fontCN: { type: "string", required: true, default: "黑体" },
          fontSize: { type: "number", required: true, default: 22 },
          bold: { type: "boolean", default: true },
          alignment: { type: "enum", default: "center" }
        }
      },
      {
        id: "author",
        name: "作者信息",
        category: "标识类",
        description: "作者姓名、单位",
        detectHint: "标题下方",
        properties: {
          fontCN: { type: "string", default: "宋体" },
          fontSize: { type: "number", default: 12 },
          alignment: { type: "enum", default: "center" }
        }
      },
      {
        id: "abstract",
        name: "摘要",
        category: "内容类",
        description: "摘要正文",
        detectHint: "'摘要'关键词之后",
        properties: {
          fontCN: { type: "string", default: "宋体" },
          fontSize: { type: "number", default: 10.5 },
          alignment: { type: "enum", default: "justify" },
          firstLineIndent: { type: "number", default: 2 }
        }
      },
      {
        id: "abstractTitle",
        name: "摘要标题",
        category: "标题类",
        description: "'摘要'二字",
        detectHint: "标题下方或正文前",
        properties: {
          fontCN: { type: "string", default: "黑体" },
          fontSize: { type: "number", default: 14 },
          bold: { type: "boolean", default: true },
          alignment: { type: "enum", default: "center" }
        }
      },
      {
        id: "keywords",
        name: "关键词",
        category: "标识类",
        description: "关键词列表",
        detectHint: "'关键词'开头",
        properties: {
          fontCN: { type: "string", default: "宋体" },
          fontSize: { type: "number", default: 10.5 },
          alignment: { type: "enum", default: "left" }
        }
      },
      {
        id: "heading1",
        name: "一级标题",
        category: "标题类",
        description: "如'1 引言'",
        detectHint: "阿拉伯数字开头加空格",
        detectPattern: "^\\d+\\s",
        properties: {
          fontCN: { type: "string", default: "黑体" },
          fontSize: { type: "number", default: 16 },
          bold: { type: "boolean", default: true },
          alignment: { type: "enum", default: "left" },
          spaceBefore: { type: "number", default: 12, unit: "pt" },
          spaceAfter: { type: "number", default: 6, unit: "pt" }
        }
      },
      {
        id: "heading2",
        name: "二级标题",
        category: "标题类",
        description: "如'1.1 背景'",
        detectHint: "两级数字编号",
        detectPattern: "^\\d+\\.\\d+\\s",
        properties: {
          fontCN: { type: "string", default: "黑体" },
          fontSize: { type: "number", default: 15 },
          bold: { type: "boolean", default: true },
          alignment: { type: "enum", default: "left" }
        }
      },
      {
        id: "heading3",
        name: "三级标题",
        category: "标题类",
        description: "如'1.1.1 概述'",
        detectHint: "三级数字编号",
        detectPattern: "^\\d+\\.\\d+\\.\\d+\\s",
        properties: {
          fontCN: { type: "string", default: "黑体" },
          fontSize: { type: "number", default: 14 },
          bold: { type: "boolean", default: true },
          alignment: { type: "enum", default: "left" }
        }
      },
      {
        id: "heading4",
        name: "四级标题",
        category: "标题类",
        description: "如'1.1.1.1'",
        detectHint: "四级数字编号",
        detectPattern: "^\\d+\\.\\d+\\.\\d+\\.\\d+\\s",
        properties: {
          fontCN: { type: "string", default: "黑体" },
          fontSize: { type: "number", default: 12 },
          bold: { type: "boolean", default: true },
          alignment: { type: "enum", default: "left" }
        }
      },
      {
        id: "heading5",
        name: "五级标题",
        category: "标题类",
        description: "如'1.1.1.1.1'",
        detectHint: "五级数字编号",
        detectPattern: "^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s",
        properties: {
          fontCN: { type: "string", default: "黑体" },
          fontSize: { type: "number", default: 12 },
          bold: { type: "boolean", default: true },
          alignment: { type: "enum", default: "left" }
        }
      },
      {
        id: "body",
        name: "正文",
        category: "正文类",
        description: "正文段落",
        detectHint: "默认类型",
        properties: {
          fontCN: { type: "string", default: "宋体" },
          fontSize: { type: "number", default: 12 },
          alignment: { type: "enum", default: "justify" },
          firstLineIndent: { type: "number", default: 2, unit: "char" },
          lineSpacing: { type: "number", default: 22, unit: "pt" }
        }
      },
      {
        id: "figureCaption",
        name: "图名",
        category: "图表类",
        description: "如'图 1 系统架构'",
        detectHint: "'图'开头加编号",
        detectPattern: "^图\\s*\\d+",
        properties: {
          fontCN: { type: "string", default: "黑体" },
          fontSize: { type: "number", default: 9 },
          alignment: { type: "enum", default: "center" }
        }
      },
      {
        id: "tableCaption",
        name: "表名",
        category: "图表类",
        description: "如'表 1 技术参数'",
        detectHint: "'表'开头加编号",
        detectPattern: "^表\\s*\\d+",
        properties: {
          fontCN: { type: "string", default: "黑体" },
          fontSize: { type: "number", default: 9 },
          alignment: { type: "enum", default: "center" }
        }
      },
      {
        id: "formulaCaption",
        name: "公式编号",
        category: "图表类",
        description: "如'(1)'、'(2-1)'",
        detectHint: "公式右侧编号",
        properties: {
          fontCN: { type: "string", default: "宋体" },
          fontSize: { type: "number", default: 12 },
          alignment: { type: "enum", default: "right" }
        }
      },
      {
        id: "listItem",
        name: "列表项",
        category: "正文类",
        description: "编号列表或项目符号",
        detectHint: "列表符号开头或连续编号",
        properties: {
          fontCN: { type: "string", default: "宋体" },
          fontSize: { type: "number", default: 12 },
          alignment: { type: "enum", default: "left" },
          leftIndent: { type: "number", default: 1, unit: "char" }
        }
      },
      {
        id: "reference",
        name: "参考文献",
        category: "附属类",
        description: "文献引用条目",
        detectHint: "'参考文献'标题后，编号开头",
        properties: {
          fontCN: { type: "string", default: "宋体" },
          fontSize: { type: "number", default: 10.5 },
          alignment: { type: "enum", default: "justify" },
          firstLineIndent: { type: "number", default: 0 },
          leftIndent: { type: "number", default: 2, unit: "char" },
          hangingIndent: { type: "number", default: 2, unit: "char", description: "悬挂缩进" }
        }
      },
      {
        id: "referenceTitle",
        name: "参考文献标题",
        category: "标题类",
        description: "'参考文献'",
        detectHint: "正文末尾",
        properties: {
          fontCN: { type: "string", default: "黑体" },
          fontSize: { type: "number", default: 14 },
          bold: { type: "boolean", default: true },
          alignment: { type: "enum", default: "center" }
        }
      },
      {
        id: "appendixTitle",
        name: "附录标题",
        category: "附属类",
        description: "如'附录 A'",
        detectHint: "'附录'开头",
        detectPattern: "^附录\\s*[A-Z]|^附录",
        properties: {
          fontCN: { type: "string", default: "黑体" },
          fontSize: { type: "number", default: 16 },
          bold: { type: "boolean", default: true },
          alignment: { type: "enum", default: "center" }
        }
      },
      {
        id: "appendixSection",
        name: "附录节题",
        category: "附属类",
        description: "如'A.1 详细说明'",
        detectHint: "附录编号开头",
        detectPattern: "^[A-Z]\\.[A-Z]?\\s|^[A-Z]\\s",
        properties: {
          fontCN: { type: "string", default: "黑体" },
          fontSize: { type: "number", default: 14 },
          bold: { type: "boolean", default: true },
          alignment: { type: "enum", default: "left" }
        }
      },
      {
        id: "header",
        name: "页眉",
        category: "页面类",
        description: "页眉内容",
        properties: {
          fontCN: { type: "string", default: "宋体" },
          fontSize: { type: "number", default: 10.5 }
        }
      },
      {
        id: "footer",
        name: "页脚",
        category: "页面类",
        description: "页码",
        properties: {
          fontCN: { type: "string", default: "宋体" },
          fontSize: { type: "number", default: 10.5 },
          alignment: { type: "enum", default: "center" }
        }
      }
    ]
  },

  // ========== 页面设置属性 ==========
  pageSetup: {
    id: "pageSetup",
    name: "页面设置",
    properties: {
      paperSize: { type: "enum", values: ["A4", "A3", "B5", "Letter"], default: "A4", description: "纸张大小" },
      orientation: { type: "enum", values: ["portrait", "landscape"], default: "portrait", description: "方向" },
      topMargin: { type: "number", default: 2.54, unit: "cm", description: "上边距" },
      bottomMargin: { type: "number", default: 2.54, unit: "cm", description: "下边距" },
      leftMargin: { type: "number", default: 3.17, unit: "cm", description: "左边距" },
      rightMargin: { type: "number", default: 3.17, unit: "cm", description: "右边距" },
      headerMargin: { type: "number", default: 1.5, unit: "cm", description: "页眉边距" },
      footerMargin: { type: "number", default: 1.75, unit: "cm", description: "页脚边距" }
    }
  }
};

// 属性值映射表（WPS API值 <-> 规范表值）
const PROPERTY_VALUE_MAP = {
  alignment: {
    0: "left",
    1: "center",
    2: "right",
    3: "justify",
    "left": 0,
    "center": 1,
    "right": 2,
    "justify": 3
  },
  paperSize: {
    1: "Letter",
    2: "LetterSmall",
    3: "Tabloid",
    4: "Ledger",
    5: "Legal",
    6: "Statement",
    7: "Executive",
    8: "A3",
    9: "A4",
    10: "A4Small",
    11: "A5",
    "A4": 9,
    "A3": 8,
    "B5": 13
  },
  orientation: {
    0: "portrait",
    1: "landscape"
  }
};

// 导出
if (typeof module !== 'undefined') {
  module.exports = { STYLE_SPEC_TABLE, PROPERTY_VALUE_MAP };
}