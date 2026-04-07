/**
 * apply-format.js
 * 按XX集团公文排版格式规范识别段落类型并应用样式。
 *
 * 设计思路：
 * 1. 定义样式规范（STYLE_RULES）
 * 2. 遍历段落，识别元素类型
 * 3. 根据元素类型应用样式规范
 *
 * 出参: { success: boolean, paragraphCount: number, appliedStyles: object }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, paragraphCount: 0, appliedStyles: {} };
  }

  // ========================================
  // 样式规范配置（按XX集团公文排版格式）
  // ========================================
  var STYLE_RULES = {
    'title': {
      name: '集团1标题',
      fontCN: '方正小标宋简体',
      fontEN: 'Times New Roman',
      fontSize: 22,
      bold: true,
      alignment: 1,
      firstLineIndent: 0,
      spaceAfter: 14
    },
    'recipient': {
      name: '集团主送单位抬头',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    'heading1': {
      name: '集团2级标题黑体',
      fontCN: '黑体',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    'heading2': {
      name: '集团3级段落重点',
      fontCN: '楷体_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 10
    },
    'numbered': {
      name: '集团4级数字编号',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 10
    },
    'body': {
      name: '集团正文文本缩进',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 32
    },
    'ending': {
      name: '集团结尾语',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 32
    },
    'attachment': {
      name: '集团附件',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    'signature': {
      name: '集团落款',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 2,
      firstLineIndent: 0
    }
  };

  // ========================================
  // 元素识别规则
  // ========================================

  function isTitle(text, index) {
    if (!text || index !== 1) return false;
    return text.length <= 50;
  }

  function isRecipient(text, index) {
    if (!text || index !== 2) return false;
    return /[：:]$/.test(text);
  }

  function isHeading1(text) {
    if (!text) return false;
    return /^[一二三四五六七八九十百]+、/.test(text);
  }

  function isHeading2(text) {
    if (!text) return false;
    return /^[（(][一二三四五六七八九十百]+[）)]/.test(text);
  }

  function isNumbered(text) {
    if (!text) return false;
    return /^第[一二三四五六七八九十百]+[项条步阶段]/.test(text);
  }

  function isEnding(text) {
    if (!text) return false;
    return /^特此[报告通知请示批复函]/.test(text);
  }

  function isSignature(text, totalParas, currentIndex) {
    if (!text) return false;
    if (currentIndex < totalParas - 2) return false;
    if (/\d{4}年.*月.*日/.test(text)) return true;
    if (text.length < 30 && !/[，。、；：！？]$/.test(text)) return true;
    return false;
  }

  // ========================================
  // 应用样式规范
  // ========================================
  function applyStyleRule(para, rule) {
    var range = para.Range;
    var fmt = para.Format;

    if (range && range.Font) {
      if (rule.fontCN) range.Font.NameFarEast = rule.fontCN;
      if (rule.fontEN) range.Font.Name = rule.fontEN;
      if (rule.fontSize) range.Font.Size = rule.fontSize;
      if (rule.bold) range.Font.Bold = true;
    }

    if (fmt) {
      if (rule.alignment !== undefined) fmt.Alignment = rule.alignment;
      if (rule.firstLineIndent !== undefined) fmt.FirstLineIndent = rule.firstLineIndent;
      if (rule.spaceAfter !== undefined) fmt.SpaceAfter = rule.spaceAfter;
    }

    if (rule.name) {
      try {
        para.Style = rule.name;
      } catch (e) {}
    }
  }

  // ========================================
  // 设置页面格式
  // ========================================
  try {
    var ps = doc.PageSetup;
    ps.PageWidth = 595.3;
    ps.PageHeight = 841.9;
    ps.TopMargin = 105;
    ps.BottomMargin = 99;
    ps.LeftMargin = 79;
    ps.RightMargin = 74;
  } catch (e) {
    console.warn('[apply-format] 页面设置失败', e);
  }

  // ========================================
  // 主流程
  // ========================================
  var appliedStyles = {};
  var totalParas = doc.Paragraphs.Count;
  var inAttachmentSection = false;

  for (var i = 1; i <= totalParas; i++) {
    var para = doc.Paragraphs.Item(i);

    var text = '';
    try {
      text = para.Range.Text;
      if (text) text = text.replace(/[\r\n]/g, '').trim();
    } catch (e) {
      text = '';
    }

    if (!text) continue;

    // 检测是否进入附件区域（"附件："开头）
    if (/^附件[：:]/.test(text)) {
      inAttachmentSection = true;
    }

    // 检测是否离开附件区域（遇到落款）
    if (isSignature(text, totalParas, i)) {
      inAttachmentSection = false;
    }

    // 识别元素类型
    var elementType = 'body';

    if (isTitle(text, i)) {
      elementType = 'title';
    } else if (isRecipient(text, i)) {
      elementType = 'recipient';
    } else if (isEnding(text)) {
      elementType = 'ending';
    } else if (inAttachmentSection) {
      // 附件区域内的段落都算附件
      elementType = 'attachment';
    } else if (isHeading1(text)) {
      elementType = 'heading1';
    } else if (isHeading2(text)) {
      elementType = 'heading2';
    } else if (isNumbered(text)) {
      elementType = 'numbered';
    } else if (isSignature(text, totalParas, i)) {
      elementType = 'signature';
    }

    var rule = STYLE_RULES[elementType];
    if (!rule) {
      rule = STYLE_RULES['body'];
      elementType = 'body';
    }

    applyStyleRule(para, rule);
    appliedStyles[elementType] = (appliedStyles[elementType] || 0) + 1;
  }

  return {
    success: true,
    paragraphCount: totalParas,
    appliedStyles: appliedStyles
  };

} catch (e) {
  console.warn('[apply-format]', e);
  return { success: false, paragraphCount: 0, appliedStyles: {}, error: String(e) };
}