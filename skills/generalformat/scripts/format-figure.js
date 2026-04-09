/**
 * 排版图 - 支持动态解析用户输入
 * 默认：图名黑体小五号居中，图片居中对齐
 */
try {
  var doc = Application.ActiveDocument;
  if (!doc) return { success: false, error: '无活动文档' };

  console.log('[format-图] 开始处理');

  // 字号映射
  var FONT_SIZE_MAP = {
    '初号': 42, '小初': 36, '一号': 26, '小一': 24,
    '二号': 22, '小二': 18, '三号': 16, '小三': 15,
    '四号': 14, '小四': 12, '五号': 10.5, '小五': 9,
    '六号': 7.5, '小六': 6.5
  };

  // 解析字号
  function parseSize(text) {
    var match = text.match(/(初号|小初|一号|小一|二号|小二|三号|小三|四号|小四|五号|小五|六号|小六)/);
    return match ? FONT_SIZE_MAP[match[1]] : 0;
  }

  // 解析字体
  function parseFont(text) {
    if (text.indexOf('黑体') !== -1) return '黑体';
    if (text.indexOf('宋体') !== -1) return '宋体';
    if (text.indexOf('楷体') !== -1) return '楷体';
    if (text.indexOf('仿宋') !== -1) return '仿宋';
    return null;
  }

  // 解析对齐
  function parseAlign(text) {
    if (text.indexOf('居中') !== -1) return 1;
    if (text.indexOf('右对齐') !== -1 || text.indexOf('靠右') !== -1) return 2;
    if (text.indexOf('两端对齐') !== -1) return 3;
    if (text.indexOf('左对齐') !== -1 || text.indexOf('靠左') !== -1) return 0;
    return -1;
  }

  // 从用户输入解析规则
  var userInput = specText || '';
  console.log('[format-图] 用户输入: ' + userInput);

  // 默认配置
  var figureCaptionRule = { fontCN: '黑体', fontSize: 9, alignment: 1 };

  // 尝试从用户输入解析图名规则
  if (userInput.indexOf('图名') !== -1) {
    var captionMatch = userInput.match(/图名[^。；,，]*/);
    if (captionMatch) {
      var ct = captionMatch[0];
      if (parseFont(ct)) figureCaptionRule.fontCN = parseFont(ct);
      if (parseSize(ct)) figureCaptionRule.fontSize = parseSize(ct);
      if (parseAlign(ct) >= 0) figureCaptionRule.alignment = parseAlign(ct);
      console.log('[format-图] 解析图名规则: ' + JSON.stringify(figureCaptionRule));
    }
  }

  console.log('[format-图] 最终规则: ' + JSON.stringify(figureCaptionRule));

  var applied = 0;

  // 辅助函数
  function applyRuleToRange(range, rule) {
    if (!range) return false;
    try {
      if (range.Font) {
        if (rule.fontCN) range.Font.NameFarEast = rule.fontCN;
        if (rule.fontSize) range.Font.Size = rule.fontSize;
      }
      if (range.ParagraphFormat) {
        if (rule.alignment !== undefined) range.ParagraphFormat.Alignment = rule.alignment;
      }
      return true;
    } catch (e) { return false; }
  }

  // 处理图名段落
  var figPattern = /^图\s*\d+/;
  var fullText = doc.Content.Text || '';
  var lines = fullText.split('\r');
  var captionCount = 0;

  for (var i = 0; i < lines.length; i++) {
    var text = String(lines[i]).replace(/[\r\u0007]/g, '').trim();
    if (figPattern.test(text)) {
      try {
        var para = doc.Paragraphs.Item(i + 1);
        if (para && para.Range) {
          if (applyRuleToRange(para.Range, figureCaptionRule)) {
            captionCount++;
          }
        }
      } catch (e) {}
    }
  }
  applied += captionCount;
  console.log('[format-图] 图名处理: ' + captionCount + '个');

  // 处理图片对齐
  var imgAlign = 1;  // 默认居中
  if (userInput.indexOf('图片左对齐') !== -1 || userInput.indexOf('图片靠左') !== -1) {
    imgAlign = 0;
  } else if (userInput.indexOf('图片右对齐') !== -1 || userInput.indexOf('图片靠右') !== -1) {
    imgAlign = 2;
  } else if (userInput.indexOf('图片居中') !== -1) {
    imgAlign = 1;
  }

  var inlineCount = doc.InlineShapes ? doc.InlineShapes.Count : 0;
  var imgCount = 0;
  for (var imgIdx = 1; imgIdx <= inlineCount; imgIdx++) {
    try {
      var inlineShape = doc.InlineShapes.Item(imgIdx);
      if (inlineShape && inlineShape.Range && inlineShape.Range.ParagraphFormat) {
        inlineShape.Range.ParagraphFormat.Alignment = imgAlign;
        imgCount++;
      }
    } catch (e) {}
  }
  applied += imgCount;
  console.log('[format-图] 图片对齐(' + (imgAlign===0?'左':imgAlign===1?'居中':'右') + '): ' + imgCount + '个');

  return { success: true, applied: applied };
} catch (e) {
  return { success: false, error: String(e) };
}