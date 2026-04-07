/**
 * detect-elements.js
 * 识别公文各要素类型及格式类型
 *
 * 支持四种公文格式：
 * - general: 通用公文格式（默认）
 * - letter: 信函格式
 * - order: 命令(令)格式
 * - minutes: 纪要格式
 *
 * 出参: { success: boolean, formatType: string, elements: array, summary: object }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, formatType: 'general', elements: [], summary: {} };
  }

  var totalParas = doc.Paragraphs.Count;
  var elements = [];

  // ========================================
  // 常量定义
  // ========================================
  var MM_TO_POINTS = 2.835;
  var FORMAT_TYPES = {
    GENERAL: 'general',    // 通用公文格式
    LETTER: 'letter',      // 信函格式
    ORDER: 'order',        // 命令(令)格式
    MINUTES: 'minutes'     // 纪要格式
  };

  // ========================================
  // 辅助函数
  // ========================================

  /**
   * 获取段落文本
   */
  function getParaText(para) {
    try {
      var text = para.Range.Text;
      return text ? text.replace(/[\r\n]/g, '').trim() : '';
    } catch (e) {
      return '';
    }
  }

  /**
   * 获取段落对齐方式
   */
  function getAlignment(para) {
    try {
      return para.Format.Alignment;
    } catch (e) {
      return -1;
    }
  }

  /**
   * 判断是否为空段落
   */
  function isEmpty(text) {
    return !text || text.length === 0;
  }

  /**
   * 获取段落距页面顶部的距离（近似，通过段落位置估算）
   */
  function getDistanceFromTop(para, pageIndex) {
    try {
      // WPS JS API 没有直接获取段落位置的方法
      // 通过段落索引和页面设置估算
      var topMargin = doc.PageSetup.TopMargin;
      return topMargin + (pageIndex - 1) * 20;  // 粗略估算
    } catch (e) {
      return 0;
    }
  }

  // ========================================
  // 格式类型识别函数
  // ========================================

  /**
   * 检测公文格式类型
   * 规则：
   * - 命令(令)格式：发文机关标志含"命令"或"令"字
   * - 纪要格式：发文机关标志含"纪要"字样
   * - 信函格式：发文机关标志不含"文件"且标题含"函"
   * - 通用公文：默认
   */
  function detectFormatType() {
    var formatType = FORMAT_TYPES.GENERAL;

    // 检查前10个段落
    for (var i = 1; i <= Math.min(10, totalParas); i++) {
      var para = doc.Paragraphs.Item(i);
      var text = getParaText(para);
      var alignment = getAlignment(para);

      if (isEmpty(text)) continue;

      // 检查发文机关标志（通常是前几个居中的段落）
      if (alignment === 1) {  // 居中
        // 命令(令)格式：含"命令"或"令"
        if (/命令|令$/.test(text)) {
          return FORMAT_TYPES.ORDER;
        }

        // 纪要格式：含"纪要"
        if (/纪要/.test(text)) {
          return FORMAT_TYPES.MINUTES;
        }

        // 信函格式特征：发文机关标志不含"文件"
        if (!/文件$/.test(text)) {
          // 继续检查是否有"函"字样的标题
          // 标记可能是信函格式，需要进一步确认
          if (/政府$|办公室$|委员会$|厅$|局$|委$/.test(text)) {
            formatType = FORMAT_TYPES.LETTER;
          }
        }
      }

      // 检查发文字号位置
      // 信函格式：发文字号在右上角（右对齐）
      // 通用公文：发文字号居中
      if (/〔\d{4}〕\d+号/.test(text) || /\[\d{4}\]\d+号/.test(text)) {
        if (alignment === 2) {  // 右对齐，信函格式特征
          return FORMAT_TYPES.LETTER;
        }
      }
    }

    // 检查标题是否含"函"
    for (var i = 1; i <= Math.min(20, totalParas); i++) {
      var text = getParaText(doc.Paragraphs.Item(i));
      if (/函$|复函$/.test(text) && text.length <= 30) {
        // 标题含"函"，可能是信函格式
        if (formatType === FORMAT_TYPES.LETTER) {
          return FORMAT_TYPES.LETTER;
        }
      }
    }

    return formatType;
  }

  // ========================================
  // 识别规则函数
  // ========================================

  /**
   * 识别份号：6位纯数字
   */
  function isFenzhen(text, index) {
    if (index > 5) return false;
    return /^\d{6}$/.test(text);
  }

  /**
   * 识别密级
   */
  function isMiji(text, index) {
    if (index > 5) return false;
    return /秘密|机密|绝密/.test(text);
  }

  /**
   * 识别紧急程度
   */
  function isJinji(text, index) {
    if (index > 5) return false;
    return /特急|加急/.test(text);
  }

  /**
   * 识别发文机关标志（根据格式类型有不同规则）
   */
  function isDocumentFlag(text, alignment, formatType) {
    if (alignment !== 1) return false;  // 居中

    // 命令格式：含"命令"或"令"
    if (formatType === FORMAT_TYPES.ORDER) {
      return /命令|令$/.test(text);
    }

    // 纪要格式：含"纪要"
    if (formatType === FORMAT_TYPES.MINUTES) {
      return /纪要/.test(text);
    }

    // 信函格式：不含"文件"
    if (formatType === FORMAT_TYPES.LETTER) {
      if (/政府$|办公室$|委员会$|厅$|局$|委$/.test(text)) {
        return true;
      }
      return false;
    }

    // 通用公文：含"文件"或机关名
    if (/文件$/.test(text)) return true;
    if (/政府$|办公室$|委员会$|厅$|局$|委$/.test(text)) return true;
    return false;
  }

  /**
   * 识别发文字号
   */
  function isDocNumber(text, alignment) {
    return /〔\d{4}〕\d+号/.test(text) || /\[\d{4}\]\d+号/.test(text);
  }

  /**
   * 识别令号（命令格式专用）
   */
  function isOrderNumber(text, formatType) {
    if (formatType !== FORMAT_TYPES.ORDER) return false;
    // 令号格式：第X号
    return /^第\d+号$/.test(text);
  }

  /**
   * 识别签发人
   */
  function isQianfaren(text) {
    return /^签发人[：:]/.test(text);
  }

  /**
   * 识别标题
   */
  function isTitle(text, afterSeparator) {
    if (!text || text.length > 50) return false;
    if (afterSeparator && text.length <= 50) {
      if (/关于|通知|决定|请示|批复|函|报告|意见|方案|规定|纪要|命令/.test(text)) {
        return true;
      }
    }
    return false;
  }

  /**
   * 识别主送机关
   */
  function isZhushong(text) {
    if (!text) return false;
    if (/[：:]$/.test(text)) {
      if (/政府|办公室|委员会|厅|局|委|公司|各|有关/.test(text)) {
        return true;
      }
    }
    return false;
  }

  /**
   * 识别一级标题
   */
  function isHeading1(text) {
    return /^[一二三四五六七八九十]+、/.test(text);
  }

  /**
   * 识别二级标题
   */
  function isHeading2(text) {
    return /^[（(][一二三四五六七八九十]+[）)]/.test(text);
  }

  /**
   * 识别三级标题
   */
  function isHeading3(text) {
    return /^\d+\./.test(text);
  }

  /**
   * 识别四级标题
   */
  function isHeading4(text) {
    return /^[（(]\d+[）)]/.test(text);
  }

  /**
   * 识别附件说明
   */
  function isFujian(text) {
    return /^附件[：:]/.test(text);
  }

  /**
   * 识别成文日期
   */
  function isDate(text) {
    return /\d{4}年\d{1,2}月\d{1,2}日/.test(text);
  }

  /**
   * 识别落款（署名）
   */
  function isSignature(text, index, total) {
    if (!text) return false;
    if (index < total - 3) return false;
    if (text.length <= 30 && !/[，。！？、；：]$/.test(text)) {
      return true;
    }
    return false;
  }

  /**
   * 识别附注
   */
  function isFuzhu(text) {
    return /^\（[^）]+\）$/.test(text);
  }

  /**
   * 识别抄送机关
   */
  function isChaosong(text) {
    return /^抄送[：:]/.test(text);
  }

  /**
   * 识别印发
   */
  function isYinfa(text) {
    return /印发\s*$/.test(text);
  }

  /**
   * 识别出席名单（纪要格式专用）
   */
  function isAttendeeList(text, formatType) {
    if (formatType !== FORMAT_TYPES.MINUTES) return false;
    return /^出席[：:]|^请假[：:]|^列席[：:]/.test(text);
  }

  // ========================================
  // 主识别流程
  // ========================================

  // 1. 检测格式类型
  var formatType = detectFormatType();

  var separatorIndex = -1;  // 版头分隔线位置
  var titleIndex = -1;      // 标题位置
  var dateIndex = -1;       // 成文日期位置
  var chaosongIndex = -1;   // 抄送位置
  var yinfaIndex = -1;      // 印发位置
  var orderNumberIndex = -1; // 令号位置（命令格式）
  var attendeeIndex = -1;   // 出席名单位置（纪要格式）

  // 2. 第一遍：识别关键位置
  for (var i = 1; i <= totalParas; i++) {
    var para = doc.Paragraphs.Item(i);
    var text = getParaText(para);

    if (isEmpty(text)) continue;

    // 查找发文字号（分隔线在其后）
    if (isDocNumber(text)) {
      separatorIndex = i;  // 分隔线在发文字号之后
    }

    // 查找令号（命令格式）
    if (isOrderNumber(text, formatType)) {
      orderNumberIndex = i;
      separatorIndex = i;  // 令号后也是分隔位置
    }

    // 查找标题
    if (titleIndex < 0 && separatorIndex > 0 && i > separatorIndex) {
      if (text.length > 0 && text.length <= 50) {
        titleIndex = i;
      }
    }

    // 查找成文日期
    if (isDate(text)) {
      dateIndex = i;
    }

    // 查找抄送
    if (isChaosong(text)) {
      chaosongIndex = i;
    }

    // 查找印发
    if (isYinfa(text)) {
      yinfaIndex = i;
    }

    // 查找出席名单（纪要格式）
    if (isAttendeeList(text, formatType)) {
      attendeeIndex = i;
    }
  }

  // 3. 第二遍：识别每个段落的类型
  for (var i = 1; i <= totalParas; i++) {
    var para = doc.Paragraphs.Item(i);
    var text = getParaText(para);
    var alignment = getAlignment(para);

    if (isEmpty(text)) {
      elements.push({ index: i, type: 'empty', text: text });
      continue;
    }

    var type = 'body';  // 默认为正文

    // 版头区域（分隔线之前）
    if (separatorIndex < 0 || i <= separatorIndex) {
      if (isFenzhen(text, i)) {
        type = 'fenzhen';
      } else if (isMiji(text, i)) {
        type = 'miji';
      } else if (isJinji(text, i)) {
        type = 'jinji';
      } else if (isDocumentFlag(text, alignment, formatType)) {
        type = 'documentFlag';
      } else if (isDocNumber(text)) {
        type = 'docNumber';
      } else if (isQianfaren(text)) {
        type = 'qianfaren';
      } else if (isOrderNumber(text, formatType)) {
        type = 'orderNumber';
      }
    }
    // 主体区域（分隔线之后，抄送之前）
    else if (chaosongIndex < 0 || i < chaosongIndex) {
      if (i === titleIndex) {
        type = 'title';
      } else if (isAttendeeList(text, formatType)) {
        type = 'attendeeList';
      } else if (isZhushong(text)) {
        type = 'zhushong';
      } else if (isHeading1(text)) {
        type = 'heading1';
      } else if (isHeading2(text)) {
        type = 'heading2';
      } else if (isHeading3(text)) {
        type = 'heading3';
      } else if (isHeading4(text)) {
        type = 'heading4';
      } else if (isFujian(text)) {
        type = 'fujian';
      } else if (isDate(text)) {
        type = 'date';
      } else if (isFuzhu(text)) {
        type = 'fuzhu';
      } else if (isSignature(text, i, totalParas)) {
        type = 'signature';
      }
    }
    // 版记区域（抄送及之后）
    else {
      if (isChaosong(text)) {
        type = 'chaosong';
      } else if (isYinfa(text)) {
        type = 'yinfa';
      }
    }

    elements.push({ index: i, type: type, text: text.substring(0, 50) });
  }

  // ========================================
  // 统计结果
  // ========================================
  var summary = {};
  for (var i = 0; i < elements.length; i++) {
    var t = elements[i].type;
    summary[t] = (summary[t] || 0) + 1;
  }

  // 添加关键位置信息
  summary.separatorIndex = separatorIndex;
  summary.titleIndex = titleIndex;
  summary.dateIndex = dateIndex;
  summary.chaosongIndex = chaosongIndex;
  summary.yinfaIndex = yinfaIndex;
  summary.orderNumberIndex = orderNumberIndex;
  summary.attendeeIndex = attendeeIndex;

  return {
    success: true,
    formatType: formatType,
    elements: elements,
    summary: summary
  };

} catch (e) {
  console.warn('[detect-elements]', e);
  return { success: false, formatType: 'general', elements: [], summary: {}, error: String(e) };
}