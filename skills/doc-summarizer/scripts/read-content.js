/**
 * read-content.js
 * 读取文档内容（直接模式）。
 *
 * 入参: isSelection (boolean)
 * 出参: { content: string, title: string }
 */

try {
  var isSelection = typeof isSelection !== 'undefined' ? isSelection : false;
  var doc = Application.ActiveDocument;
  if (!doc || !doc.Content) {
    return { content: '', title: '' };
  }

  var content = '';
  var title = '';

  // 读取选区或全文
  if (isSelection) {
    var sel = wps.Application.Selection ? wps.Application.Selection.Range : null;
    content = sel && sel.Range && sel.Range.Text ? sel.Range.Text : '';
  } else {
    content = doc.Content && doc.Content.Text ? doc.Content.Text : '';
    // 提取文档标题（第一段非空内容）
    if (doc.Paragraphs && doc.Paragraphs.Count > 0) {
      var firstPara = doc.Paragraphs.Item(1);
      title = firstPara && firstPara.Range && firstPara.Range.Text ? firstPara.Range.Text.trim() : '';
    }
  }

  return {
    content: content,
    title: title
  };
} catch (e) {
  console.warn('[read-content]', e);
  return { content: '', title: '' };
}
