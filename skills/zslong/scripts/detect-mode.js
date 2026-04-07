/**
 * detect-mode.js
 * 检测文档大小，决定执行模式。
 *
 * 出参: { charCount: number, isSelection: boolean, mode: 'direct' | 'schedule' }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { charCount: 0, isSelection: false, mode: 'direct' };
  }

  // 检测选区
  var isSelection = false;
  var charCount = 0;

  try {
    var sel = Application.Selection;
    if (sel && sel.Range && sel.Range.Start !== sel.Range.End) {
      isSelection = true;
      charCount = sel.Range.Text ? sel.Range.Text.length : 0;
    }
  } catch (e) {
    isSelection = false;
  }

  // 若无选区，统计全文
  if (!isSelection) {
    try {
      charCount = doc.Content && doc.Content.Text ? doc.Content.Text.length : 0;
    } catch (e) {
      charCount = 0;
    }
  }

  // 根据字符数决定模式
  var THRESHOLD = 15000;
  var mode = charCount > THRESHOLD ? 'schedule' : 'direct';

  return {
    charCount: charCount,
    isSelection: isSelection,
    mode: mode
  };
} catch (e) {
  console.warn('[detect-mode]', e);
  return { charCount: 0, isSelection: false, mode: 'direct' };
}