// 提取文档中的编号标题
try {
  var useSelection = typeof useSelection !== 'undefined' ? useSelection : false;
  var startPara = typeof startPara !== 'undefined' ? startPara : 1;
  var endPara = typeof endPara !== 'undefined' ? endPara : 999999;

  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, error: '无活动文档' };
  }

  var paras = doc.Paragraphs;
  var total = paras.Count;

  // 确定处理的段落范围
  var startIndex = startPara;
  var endIndex = endPara;
  if (endIndex > total) {
    endIndex = total;
  }

  var titles = [];
  var titleIndex = 0;

  // 定义编号样式正则表达式
  var patterns = {
    chineseLevel1: /^一[、．.]\s*/,
    chineseLevel2: /^（[一二三四五六七八九十]+）\s*/,
    chineseLevel3: /^(\d+)\.[^0-9.]/,
    chineseLevel4: /^（\d+）\s*/,
    chineseLevel5: /^①\s*/,

    // 数字层次式：支持 "1 标题" 和 "1. 标题" 两种格式
    digitalLevel1: /^(\d{1,2})(?:\.\s|\s+)(?!\d)/,  // 匹配 "1 标题" 或 "1. 标题"
    digitalLevel2: /^(\d+\.\d+)\s/,
    digitalLevel3: /^(\d+\.\d+\.\d+)\s/,
    digitalLevel4: /^(\d+\.\d+\.\d+\.\d+)\s/
  };

  // 遍历段落
  for (var i = startIndex; i <= endIndex; i++) {
    var para = paras.Item(i);
    if (!para || !para.Range) { continue; }

    var text = para.Range.Text || '';
    text = text.replace(/[\r\n]/g, '').trim();

    if (text.length === 0) { continue; }

    var level = 0;
    var matchedPattern = '';

    // 检测中文公文式编号
    if (patterns.chineseLevel1.test(text)) {
      level = 1;
      matchedPattern = 'chinese-1';
    } else if (patterns.chineseLevel2.test(text)) {
      level = 2;
      matchedPattern = 'chinese-2';
    } else if (patterns.chineseLevel3.test(text)) {
      level = 3;
      matchedPattern = 'chinese-3';
    } else if (patterns.chineseLevel4.test(text)) {
      level = 4;
      matchedPattern = 'chinese-4';
    } else if (patterns.chineseLevel5.test(text)) {
      level = 5;
      matchedPattern = 'chinese-5';
    }

    // 检测数字层次式编号
    else if (patterns.digitalLevel1.test(text)) {
      level = 1;
      matchedPattern = 'digital-1';
    } else if (patterns.digitalLevel2.test(text)) {
      level = 2;
      matchedPattern = 'digital-2';
    } else if (patterns.digitalLevel3.test(text)) {
      level = 3;
      matchedPattern = 'digital-3';
    } else if (patterns.digitalLevel4.test(text)) {
      level = 4;
      matchedPattern = 'digital-4';
    }

    // 如果匹配到编号，记录标题
    if (level > 0) {
      titleIndex++;
      titles.push({
        index: titleIndex,
        text: text,
        level: level,
        paraIndex: i,
        pattern: matchedPattern,
        fullText: text
      });
    }
  }

  return {
    success: true,
    titles: titles,
    totalTitles: titles.length
  };

} catch (e) {
  console.warn('[extract-titles]', e);
  return { success: false, error: String(e) };
}
