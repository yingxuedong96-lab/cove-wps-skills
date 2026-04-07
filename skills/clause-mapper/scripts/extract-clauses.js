/**
 * extract-clauses.js
 * 从当前 WPS 文档中提取草稿条款内容。
 *
 * 出参: { clauses: [{ index: number, title: string, content: string }], hasSelection: boolean }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc || !doc.Content) {
    return { clauses: [], hasSelection: false };
  }

  // 检测是否有选区
  var hasSelection = false;
  try {
    var sel = Application.Selection;
    if (sel && sel.Range && sel.Range.Start !== sel.Range.End) {
      hasSelection = true;
    }
  } catch (e) {}

  var clauses = [];
  var text = doc.Content.Text;
  
  if (!text || text.trim().length === 0) {
    return { clauses: [], hasSelection: hasSelection };
  }

  // 按段落分割
  var paragraphs = doc.Paragraphs;
  var total = paragraphs.Count;
  
  var currentIndex = 0;
  var currentTitle = '';
  var currentContent = '';
  var inClause = false;
  
  for (var i = 1; i <= total; i++) {
    var para = paragraphs.Item(i);
    var paraText = para.Range.Text.trim();
    
    if (!paraText) {
      continue;
    }
    
    // 检测条款标题（数字开头，如"1. xxx"、"第1条 xxx"）
    var clauseMatch = paraText.match(/^(\d+)[.、\s]+(.+)/);
    var articleMatch = paraText.match(/^第([一二三四五六七八九十\d]+)条\s*(.*)/);
    
    if (clauseMatch || articleMatch) {
      // 保存上一个条款
      if (inClause && currentContent) {
        clauses.push({
          index: currentIndex,
          title: currentTitle,
          content: currentContent.trim()
        });
      }
      
      // 开始新条款
      if (clauseMatch) {
        currentIndex = parseInt(clauseMatch[1], 10);
        currentTitle = clauseMatch[2].trim();
      } else if (articleMatch) {
        currentIndex = parseInt(articleMatch[1], 10);
        currentTitle = articleMatch[2].trim();
      }
      currentContent = paraText;
      inClause = true;
    } else if (inClause) {
      // 继续当前条款内容
      currentContent += '\n' + paraText;
    } else {
      // 非条款内容（如项目名称、签署日期等）
      clauses.push({
        index: 0,
        title: paraText.split('：')[0] || paraText.substring(0, 20),
        content: paraText
      });
    }
  }
  
  // 保存最后一个条款
  if (inClause && currentContent) {
    clauses.push({
      index: currentIndex,
      title: currentTitle,
      content: currentContent.trim()
    });
  }

  return { 
    clauses: clauses, 
    hasSelection: hasSelection
  };
} catch (e) {
  console.warn('[extract-clauses]', e);
  return { clauses: [], hasSelection: false };
}
