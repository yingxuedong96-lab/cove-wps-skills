/**
 * read-source.js
 * 读取当前文档内容，供 LLM 分析投资条款。
 *
 * 出参: { text: string, isEmpty: boolean }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc || !doc.Content) {
    return { text: '', isEmpty: true };
  }

  var text = '';
  try {
    text = doc.Content.Text || '';
  } catch (e) {
    text = '';
  }

  // 清理文本（去除多余空白字符）
  text = text.replace(/\r/g, '\n').replace(/\n+/g, '\n').trim();

  var isEmpty = text.length === 0;

  return { text: text, isEmpty: isEmpty };
} catch (e) {
  console.warn('[read-source]', e);
  return { text: '', isEmpty: true, error: String(e) };
}
