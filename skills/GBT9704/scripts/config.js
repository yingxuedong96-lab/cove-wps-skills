/**
 * config.js
 * GB/T 9704-2012 国标公文格式配置常量
 */

// ========================================
// 单位转换常量
// ========================================
var MM_TO_POINTS = 2.835;  // 1mm ≈ 2.835 磅

// ========================================
// 页面设置（GB/T 9704-2012）
// ========================================
var PAGE_SETUP = {
  pageWidth: 210 * MM_TO_POINTS,   // A4 宽 210mm ≈ 595.3 磅
  pageHeight: 297 * MM_TO_POINTS,  // A4 高 297mm ≈ 841.9 磅
  topMargin: 37 * MM_TO_POINTS,    // 天头 37mm ≈ 104.9 磅
  bottomMargin: 35 * MM_TO_POINTS, // 地脚 35mm ≈ 99.2 磅
  leftMargin: 28 * MM_TO_POINTS,   // 订口 28mm ≈ 79.4 磅
  rightMargin: 26 * MM_TO_POINTS,  // 翻口 26mm ≈ 73.7 磅
  // 版心尺寸
  contentWidth: 156 * MM_TO_POINTS,  // 版心宽 156mm
  contentHeight: 225 * MM_TO_POINTS, // 版心高 225mm
  // 行数字数
  linesPerPage: 22,
  charsPerLine: 28
};

// ========================================
// 字号对照表（号数 → 磅值）
// ========================================
var FONT_SIZES = {
  '初号': 42,
  '小初': 36,
  '一号': 26,
  '小一': 24,
  '二号': 22,
  '小二': 18,
  '三号': 16,
  '小三': 15,
  '四号': 14,
  '小四': 12,
  '五号': 10.5,
  '小五': 9
};

// ========================================
// 样式定义（21个公文要素）
// ========================================
var STYLE_RULES = {
  // ===== 版头 =====
  'fenzhen': {
    name: 'GBT9704-份号',
    fontCN: '仿宋_GB2312',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: false,
    alignment: 0,  // 左对齐
    firstLineIndent: 0,
    color: 0x000000
  },
  'miji': {
    name: 'GBT9704-密级',
    fontCN: '黑体',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: false,
    alignment: 0,
    firstLineIndent: 0,
    color: 0x000000
  },
  'jinji': {
    name: 'GBT9704-紧急程度',
    fontCN: '黑体',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: false,
    alignment: 0,
    firstLineIndent: 0,
    color: 0x000000
  },
  'documentFlag': {
    name: 'GBT9704-发文机关标志',
    fontCN: '方正小标宋简体',
    fontEN: 'Times New Roman',
    fontSize: 0,   // 通常由用户设定或根据机关名长度自适应
    bold: true,
    alignment: 1,  // 居中
    firstLineIndent: 0,
    color: 0xFF0000  // 红色
  },
  'docNumber': {
    name: 'GBT9704-发文字号',
    fontCN: '仿宋_GB2312',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: false,
    alignment: 1,  // 居中
    firstLineIndent: 0,
    color: 0x000000
  },
  'qianfaren': {
    name: 'GBT9704-签发人',
    fontCN: '仿宋_GB2312',  // "签发人"三字
    fontCN2: '楷体_GB2312', // 姓名用楷体
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: false,
    alignment: 2,  // 右对齐
    firstLineIndent: 0,
    color: 0x000000
  },

  // ===== 主体 =====
  'title': {
    name: 'GBT9704-标题',
    fontCN: '方正小标宋简体',
    fontEN: 'Times New Roman',
    fontSize: 22,  // 2号
    bold: true,
    alignment: 1,  // 居中
    firstLineIndent: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    color: 0x000000
  },
  'zhushong': {
    name: 'GBT9704-主送机关',
    fontCN: '仿宋_GB2312',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: false,
    alignment: 0,  // 左对齐
    firstLineIndent: 0,
    color: 0x000000
  },
  'heading1': {
    name: 'GBT9704-一级标题',
    fontCN: '黑体',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: false,
    alignment: 0,
    firstLineIndent: 0,
    color: 0x000000
  },
  'heading2': {
    name: 'GBT9704-二级标题',
    fontCN: '楷体_GB2312',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: false,
    alignment: 0,
    firstLineIndent: 0,
    color: 0x000000
  },
  'heading3': {
    name: 'GBT9704-三级标题',
    fontCN: '仿宋_GB2312',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: true,    // 加粗
    alignment: 0,
    firstLineIndent: 0,
    color: 0x000000
  },
  'heading4': {
    name: 'GBT9704-四级标题',
    fontCN: '仿宋_GB2312',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: true,    // 加粗
    alignment: 0,
    firstLineIndent: 0,
    color: 0x000000
  },
  'body': {
    name: 'GBT9704-正文',
    fontCN: '仿宋_GB2312',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: false,
    alignment: 0,  // 两端对齐（WPS中0表示左对齐，3表示两端对齐）
    firstLineIndent: 32,  // 约2字符
    lineSpacing: 28,  // 固定值28磅
    color: 0x000000
  },
  'fujian': {
    name: 'GBT9704-附件说明',
    fontCN: '仿宋_GB2312',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: false,
    alignment: 0,
    firstLineIndent: 0,
    color: 0x000000
  },
  'signature': {
    name: 'GBT9704-落款',
    fontCN: '仿宋_GB2312',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: false,
    alignment: 2,  // 右对齐
    firstLineIndent: 0,
    color: 0x000000
  },
  'date': {
    name: 'GBT9704-成文日期',
    fontCN: '仿宋_GB2312',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: false,
    alignment: 2,  // 右对齐
    firstLineIndent: 0,
    color: 0x000000
  },
  'fuzhu': {
    name: 'GBT9704-附注',
    fontCN: '仿宋_GB2312',
    fontEN: 'Times New Roman',
    fontSize: 16,  // 3号
    bold: false,
    alignment: 0,  // 左对齐
    firstLineIndent: 32,
    color: 0x000000
  },

  // ===== 版记 =====
  'chaosong': {
    name: 'GBT9704-抄送',
    fontCN: '仿宋_GB2312',
    fontEN: 'Times New Roman',
    fontSize: 14,  // 4号
    bold: false,
    alignment: 0,
    firstLineIndent: 0,
    color: 0x000000
  },
  'yinfa': {
    name: 'GBT9704-印发',
    fontCN: '仿宋_GB2312',
    fontEN: 'Times New Roman',
    fontSize: 14,  // 4号
    bold: false,
    alignment: 0,
    firstLineIndent: 0,
    color: 0x000000
  },

  // ===== 页码 =====
  'pagenumber': {
    name: 'GBT9704-页码',
    fontCN: '宋体',
    fontEN: 'Times New Roman',
    fontSize: 14,  // 4号
    bold: false,
    alignment: 1,  // 居中（实际通过奇偶页调整）
    color: 0x000000
  }
};

// ========================================
// 元素识别规则
// ========================================
var DETECT_RULES = {
  // 版头区域
  fenzhen: {
    patterns: [/^\d{6}$/],
    positionRange: [1, 3],  // 前3段内
    alignment: 'left'
  },
  miji: {
    patterns: [/秘密|机密|绝密/],
    positionRange: [1, 5],
    alignment: 'left'
  },
  jinji: {
    patterns: [/特急|加急/],
    positionRange: [1, 5],
    alignment: 'left'
  },
  documentFlag: {
    patterns: [/文件$/, /政府$/, /办公室$/, /委员会$/, /厅$/, /局$/, /委$/],
    alignment: 'center',
    keywords: ['文件']
  },
  docNumber: {
    patterns: [/〔\d{4}〕\d+号/, /\[\d{4}\]\d+号/],
    alignment: 'center'
  },
  qianfaren: {
    patterns: [/^签发人[：:]/],
    alignment: 'right'
  },

  // 主体区域
  title: {
    maxLength: 50,
    position: 'afterSeparator',
    alignment: 'center',
    keywords: ['关于', '通知', '决定', '请示', '批复', '函', '报告', '意见', '方案', '规定']
  },
  zhushong: {
    patterns: [/^[^，。！？\n]+[：:]$/],
    position: 'afterTitle',
    keywords: ['政府', '办公室', '委员会', '厅', '局', '委', '公司', '各']
  },
  heading1: {
    patterns: [/^[一二三四五六七八九十]+、/]
  },
  heading2: {
    patterns: [/^[（(][一二三四五六七八九十]+[）)]/]
  },
  heading3: {
    patterns: [/^\d+\./]
  },
  heading4: {
    patterns: [/^[（(]\d+[）)]/]
  },
  fujian: {
    patterns: [/^附件[：:]/]
  },
  date: {
    patterns: [/\d{4}年\d{1,2}月\d{1,2}日/],
    position: 'end'
  },
  signature: {
    position: 'end',
    noPunctuation: true,
    maxLength: 30
  },
  fuzhu: {
    patterns: [/^\（[^）]+\）$/]
  },

  // 版记区域
  chaosong: {
    patterns: [/^抄送[：:]/]
  },
  yinfa: {
    patterns: [/印发$/, /印发\s*$/],
    hasDate: true
  }
};

// ========================================
// 分隔线配置
// ========================================
var SEPARATOR_CONFIG = {
  header: {
    color: 0xFF0000,      // 红色
    position: 'afterDocNumber',
    distance: 4 * MM_TO_POINTS,  // 发文字号之下4mm
    width: PAGE_SETUP.contentWidth
  },
  footerFirst: {
    color: 0x000000,      // 黑色
    type: 'thick',        // 粗线
    height: 0.35          // 0.35mm
  },
  footerMiddle: {
    color: 0x000000,
    type: 'thin',         // 细线
    height: 0.25          // 0.25mm
  },
  footerLast: {
    color: 0x000000,
    type: 'thick',
    height: 0.35
  }
};

// ========================================
// 页码配置
// ========================================
var PAGENUMBER_CONFIG = {
  fontSize: 14,           // 4号
  fontCN: '宋体',
  format: '- {PAGE} -',   // 一字线格式
  distanceFromContent: 7 * MM_TO_POINTS,  // 距版心7mm
  oddAlignment: 2,        // 奇数页右对齐
  evenAlignment: 0,       // 偶数页左对齐
  firstPageVisible: false // 首页不显示
};