/**
 * detect-size.js
 * 识别用户意图（内容提炼/会议纪要/周报助手）并检测文档字数，判断执行模式。
 *
 * 入参: userMessage (string) - 用户输入的完整消息
 * 出参: { charCount: number, isSelection: boolean, mode: 'direct' | 'schedule', funcType: 'extract' | 'minutes' | 'weekly' }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc || !doc.Content) {
    return { charCount: 0, isSelection: false, mode: 'direct', funcType: 'extract' };
  }

  // 检测是否有选区
  var isSelection = false;
  try {
    var sel = wps.Application.Selection ? wps.Application.Selection.Range : null;
    isSelection = !!(sel && sel.Start !== sel.End);
  } catch (e) {}

  // 获取内容字数
  var charCount = 0;
  if (isSelection) {
    var text = sel && sel.Range && sel.Range.Text ? sel.Range.Text : '';
    charCount = text.length;
  } else {
    var fullText = doc.Content && doc.Content.Text ? doc.Content.Text : '';
    charCount = fullText.length;
  }

  // 判断执行模式：超过15000字使用调度器模式
  var mode = charCount > 15000 ? 'schedule' : 'direct';

  // 识别功能类型（从用户消息中判断）
  var funcType = 'extract'; // 默认为内容提炼
  var userMessage = typeof userMessage !== 'undefined' ? userMessage : '';

  if (userMessage) {
    var lowerMsg = userMessage.toLowerCase();

    // 会议纪要关键词（更全面）
    var minutesKeywords = [
      '会议纪要', '会议记录', '生成纪要', '会议总结', '会议笔记',
      '会议内容', '会议文档', '会议汇报', '开会记录', '会议要点',
      '整理会议', '汇总会议', '纪要', '会议'
    ];

    // 周报助手关键词（更全面）
    var weeklyKeywords = [
      '周报', '工作周报', '写周报', '生成周报', '本周总结',
      '周总结', '一周工作', '周报生成', '工作总结', '本周工作',
      '下周计划', '周报助手', '每周总结', '周工作'
    ];

    // 内容提炼关键词
    var extractKeywords = [
      '内容提炼', '总结文档', '文档摘要', '生成摘要', '生成要点',
      '文档概览', '提炼要点', '总结一下', '概括一下', '摘要',
      '要点总结', '内容总结', '文档总结', '总结', '提炼'
    ];

    // 判断会议纪要
    var isMinutes = false;
    for (var i = 0; i < minutesKeywords.length; i++) {
      if (lowerMsg.indexOf(minutesKeywords[i]) !== -1) {
        isMinutes = true;
        break;
      }
    }

    // 判断周报助手
    var isWeekly = false;
    for (var j = 0; j < weeklyKeywords.length; j++) {
      if (lowerMsg.indexOf(weeklyKeywords[j]) !== -1) {
        isWeekly = true;
        break;
      }
    }

    // 判断内容提炼
    var isExtract = false;
    for (var k = 0; k < extractKeywords.length; k++) {
      if (lowerMsg.indexOf(extractKeywords[k]) !== -1) {
        isExtract = true;
        break;
      }
    }

    // 优先级判断：周报 > 会议纪要 > 内容提炼
    if (isWeekly) {
      funcType = 'weekly';
    } else if (isMinutes) {
      funcType = 'minutes';
    } else if (isExtract) {
      funcType = 'extract';
    }
  }

  return {
    charCount: charCount,
    isSelection: isSelection,
    mode: mode,
    funcType: funcType
  };
} catch (e) {
  console.warn('[detect-size]', e);
  return { charCount: 0, isSelection: false, mode: 'direct', funcType: 'extract' };
}
