/**
 * 排版图 - 图名黑体小五号居中，图片居中对齐
 */
try {
  var config = {
    specText: "图名用黑体小五号居中，图片居中对齐",
    paragraphRules: {
      figureCaption: { fontCN: "黑体", fontSize: 9, alignment: 1 }
    }
  };

  // 调用主引擎
  var result = (function() {
    var doc = Application.ActiveDocument;
    if (!doc) return { success: false, error: '无活动文档' };

    var paraCount = doc.Paragraphs.Count || 0;
    console.log('[format-图] 开始处理，段落数: ' + paraCount);

    var FONT_SIZE_MAP = {
      '初号': 42, '小初': 36, '一号': 26, '小一': 24,
      '二号': 22, '小二': 18, '三号': 16, '小三': 15,
      '四号': 14, '小四': 12, '五号': 10.5, '小五': 9,
      '六号': 7.5, '小六': 6.5, '七号': 5.5, '八号': 5
    };

    var rules = config.paragraphRules || {};
    var specText = config.specText || '';
    var applied = 0;

    // 辅助函数
    function applyRuleToRange(range, rule) {
      if (!range) return false;
      try {
        if (range.Font) {
          if (rule.fontCN) range.Font.NameFarEast = rule.fontCN;
          if (rule.fontSize) range.Font.Size = rule.fontSize;
          if (rule.bold !== undefined) range.Font.Bold = rule.bold ? -1 : 0;
        }
        if (range.ParagraphFormat) {
          if (rule.alignment !== undefined) range.ParagraphFormat.Alignment = rule.alignment;
        }
        return true;
      } catch (e) { return false; }
    }

    // 处理图名段落
    if (rules.figureCaption) {
      var figPattern = /^图\s*\d+/;
      var fullText = doc.Content.Text || '';
      var lines = fullText.split('\r');

      for (var i = 0; i < lines.length; i++) {
        var text = String(lines[i]).replace(/[\r\u0007]/g, '').trim();
        if (figPattern.test(text)) {
          try {
            var para = doc.Paragraphs.Item(i + 1);
            if (para && para.Range) {
              if (applyRuleToRange(para.Range, rules.figureCaption)) {
                applied++;
              }
            }
          } catch (e) {}
        }
      }
      console.log('[format-图] 图名处理: ' + applied + '个');
    }

    // 处理图片居中
    if (specText.indexOf('图片居中') !== -1) {
      var inlineCount = doc.InlineShapes ? doc.InlineShapes.Count : 0;
      var imgCount = 0;
      for (var imgIdx = 1; imgIdx <= inlineCount; imgIdx++) {
        try {
          var inlineShape = doc.InlineShapes.Item(imgIdx);
          if (inlineShape && inlineShape.Range && inlineShape.Range.ParagraphFormat) {
            inlineShape.Range.ParagraphFormat.Alignment = 1;  // 居中
            imgCount++;
          }
        } catch (e) {}
      }
      applied += imgCount;
      console.log('[format-图] 图片居中: ' + imgCount + '个');
    }

    return { success: true, applied: applied, elapsedMs: 0 };
  })();

  return result;
} catch (e) {
  return { success: false, error: String(e) };
}