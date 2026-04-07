// 初始化调度器：检测文档大小，决定使用直接模式还是调度器模式
try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, error: '无活动文档' };
  }

  var text = (doc.Content && doc.Content.Text) ? doc.Content.Text : '';
  var charCount = text.replace(/\r/g, '').length;

  // 检测是否有选区
  var isSelection = false;
  try {
    var sel = Application.Selection;
    if (sel && sel.Type !== 1) {
      isSelection = true;
    }
  } catch (e) {}

  // 阈值：15000字
  var THRESHOLD = 15000;
  var path = 'direct';

  if (charCount > THRESHOLD && !isSelection) {
    path = 'schedule';
  }

  return {
    success: true,
    charCount: charCount,
    isSelection: isSelection,
    path: path,
    message: path === 'direct' ? '使用直接模式' : '使用调度器模式'
  };

} catch (e) {
  console.warn('[init-schedule]', e);
  return { success: false, error: String(e) };
}
