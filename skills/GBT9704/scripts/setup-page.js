/**
 * setup-page.js
 * 设置 GB/T 9704-2012 国标公文页面格式
 *
 * 支持四种公文格式：
 * - general: 通用公文格式（默认）
 * - letter: 信函格式（上边距30mm）
 * - order: 命令(令)格式
 * - minutes: 纪要格式
 *
 * 出参: { success: boolean, formatType: string, message: string }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, formatType: 'general', message: '未找到活动文档' };
  }

  // ========================================
  // 单位转换常量
  // ========================================
  var MM_TO_POINTS = 2.835;

  // ========================================
  // 格式类型检测
  // ========================================
  var FORMAT_TYPES = {
    GENERAL: 'general',
    LETTER: 'letter',
    ORDER: 'order',
    MINUTES: 'minutes'
  };

  function detectFormatType() {
    var totalParas = doc.Paragraphs.Count;
    var detectedInfo = {
      type: FORMAT_TYPES.GENERAL,
      reasons: []
    };

    for (var i = 1; i <= Math.min(10, totalParas); i++) {
      var para = doc.Paragraphs.Item(i);
      var text = para.Range.Text ? para.Range.Text.replace(/[\r\n]/g, '').trim() : '';
      var alignment = para.Format.Alignment;

      if (!text) continue;

      // 命令(令)格式
      // 注意：不要求居中，原文档可能格式不规范
      if (/命令|令$/.test(text)) {
        detectedInfo.type = FORMAT_TYPES.ORDER;
        detectedInfo.reasons.push('第' + i + '段: 命令格式 "' + text + '"');
        console.log('[setup-page] 检测到命令格式: ' + text);
        return detectedInfo.type;
      }

      // 纪要格式
      // 注意：不要求居中，原文档可能格式不规范
      if (/纪要/.test(text)) {
        detectedInfo.type = FORMAT_TYPES.MINUTES;
        detectedInfo.reasons.push('第' + i + '段: 纪要格式 "' + text + '"');
        console.log('[setup-page] 检测到纪要格式: ' + text);
        return detectedInfo.type;
      }

      // 信函格式：发文字号右对齐
      if (/〔\d{4}〕\d+号/.test(text) || /\[\d{4}\]\d+号/.test(text)) {
        if (alignment === 2) {  // 右对齐
          detectedInfo.type = FORMAT_TYPES.LETTER;
          detectedInfo.reasons.push('第' + i + '段: 信函格式(发文字号右对齐) "' + text + '"');
          console.log('[setup-page] 检测到信函格式: ' + text);
          return detectedInfo.type;
        }
      }

      // 信函格式：发文机关标志不含"文件"且后续标题含"函"
      if (alignment === 1 && !/文件$/.test(text)) {
        if (/政府$|办公室$|委员会$|厅$|局$|委$/.test(text)) {
          for (var j = i + 1; j <= Math.min(20, totalParas); j++) {
            var nextText = doc.Paragraphs.Item(j).Range.Text;
            if (nextText && /函$|复函$/.test(nextText.replace(/[\r\n]/g, '').trim())) {
              detectedInfo.type = FORMAT_TYPES.LETTER;
              detectedInfo.reasons.push('第' + i + '段: 信函格式(标志+函标题)');
              console.log('[setup-page] 检测到信函格式: ' + text);
              return detectedInfo.type;
            }
          }
        }
      }

      // 信函格式：标题含"函"且回溯发文机关标志不含"文件"
      if (text.length > 5 && text.length <= 50 && /函$|复函$/.test(text)) {
        for (var k = 1; k < i; k++) {
          var prevText = doc.Paragraphs.Item(k).Range.Text;
          if (prevText) {
            var cleanPrev = prevText.replace(/[\r\n]/g, '').trim();
            if (/政府$|办公室$|委员会$|厅$|局$|委$/.test(cleanPrev) && !/文件$/.test(cleanPrev)) {
              detectedInfo.type = FORMAT_TYPES.LETTER;
              detectedInfo.reasons.push('第' + i + '段: 信函格式(函标题+标志)');
              console.log('[setup-page] 检测到信函格式: ' + text);
              return detectedInfo.type;
            }
          }
        }
      }
    }

    console.log('[setup-page] 使用默认格式: 通用公文');
    return FORMAT_TYPES.GENERAL;
  }

  var formatType = detectFormatType();

  // ========================================
  // 页面设置（根据格式类型）
  // ========================================
  var ps = doc.PageSetup;

  // 纸张大小：A4 (210mm × 297mm)
  ps.PageWidth = 210 * MM_TO_POINTS;   // ≈ 595.3 磅
  ps.PageHeight = 297 * MM_TO_POINTS;  // ≈ 841.9 磅

  // 页边距（根据格式类型调整）
  var topMargin, bottomMargin, leftMargin, rightMargin;

  switch (formatType) {
    case FORMAT_TYPES.LETTER:
      // 信函格式：发文机关标志距上页边30mm
      topMargin = 30 * MM_TO_POINTS;    // 上边距 30mm
      bottomMargin = 35 * MM_TO_POINTS; // 地脚 35mm
      leftMargin = 28 * MM_TO_POINTS;   // 订口 28mm
      rightMargin = 26 * MM_TO_POINTS;  // 翻口 26mm
      break;

    case FORMAT_TYPES.ORDER:
      // 命令格式：发文机关标志距版心20mm（版心到上边距37mm，所以发文机关标志距上边距约17mm）
      topMargin = 37 * MM_TO_POINTS;
      bottomMargin = 35 * MM_TO_POINTS;
      leftMargin = 28 * MM_TO_POINTS;
      rightMargin = 26 * MM_TO_POINTS;
      break;

    case FORMAT_TYPES.MINUTES:
      // 纪要格式：与通用公文相同
      topMargin = 37 * MM_TO_POINTS;
      bottomMargin = 35 * MM_TO_POINTS;
      leftMargin = 28 * MM_TO_POINTS;
      rightMargin = 26 * MM_TO_POINTS;
      break;

    default:
      // 通用公文格式（默认）
      topMargin = 37 * MM_TO_POINTS;    // 天头 37mm
      bottomMargin = 35 * MM_TO_POINTS; // 地脚 35mm
      leftMargin = 28 * MM_TO_POINTS;   // 订口 28mm
      rightMargin = 26 * MM_TO_POINTS;  // 翻口 26mm
  }

  ps.TopMargin = topMargin;
  ps.BottomMargin = bottomMargin;
  ps.LeftMargin = leftMargin;
  ps.RightMargin = rightMargin;

  // 页面方向：纵向
  ps.Orientation = 0;  // 0 = 纵向, 1 = 横向

  // ========================================
  // 页眉页脚距离
  // ========================================
  try {
    ps.HeaderDistance = 15;  // 约 5mm
    ps.FooterDistance = 28 * MM_TO_POINTS;  // 一字线距版心7mm
  } catch (e) {}

  // ========================================
  // 奇偶页不同
  // ========================================
  try {
    ps.OddAndEvenPagesHeaderFooter = true;
  } catch (e) {}

  // ========================================
  // 首页不同
  // 仅信函格式首页无页码，其他格式首页显示页码
  // ========================================
  try {
    ps.DifferentFirstPageHeaderFooter = (formatType === FORMAT_TYPES.LETTER);
  } catch (e) {}

  // ========================================
  // 返回结果
  // ========================================
  var formatNames = {
    'general': '通用公文格式',
    'letter': '信函格式',
    'order': '命令(令)格式',
    'minutes': '纪要格式'
  };

  return {
    success: true,
    formatType: formatType,
    formatName: formatNames[formatType] || '通用公文格式',
    message: '页面设置完成（' + (formatNames[formatType] || '通用公文格式') + '）',
    pageSetup: {
      pageWidth: Math.round(ps.PageWidth / MM_TO_POINTS) + 'mm',
      pageHeight: Math.round(ps.PageHeight / MM_TO_POINTS) + 'mm',
      topMargin: Math.round(ps.TopMargin / MM_TO_POINTS) + 'mm',
      bottomMargin: Math.round(ps.BottomMargin / MM_TO_POINTS) + 'mm',
      leftMargin: Math.round(ps.LeftMargin / MM_TO_POINTS) + 'mm',
      rightMargin: Math.round(ps.RightMargin / MM_TO_POINTS) + 'mm'
    }
  };

} catch (e) {
  console.warn('[setup-page]', e);
  return { success: false, formatType: 'general', message: String(e) };
}
