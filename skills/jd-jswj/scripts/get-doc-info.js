// get-doc-info.js
// 读取当前 WPS 活动文档的段落内容，判断文档规模，返回校对所需数据

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, error: '无活动文档，请先在 WPS 中打开需要校对的文档' };
  }

  var docName = doc.Name ? doc.Name : '未知文档';

  // ── 检测选区 ──────────────────────────────────────────────────────────────
  var isSelection = false;
  var selText = '';
  try {
    var sel = Application.Selection;
    if (sel && sel.Type !== 1) {
      isSelection = true;
      selText = sel.Text ? sel.Text : '';
    }
  } catch (selErr) {}

  // ── 选区模式：只处理选中文字 ───────────────────────────────────────────────
  if (isSelection) {
    var selClean = selText.replace(/\r/g, '');
    return {
      success: true,
      path: 'direct',
      charCount: selClean.length,
      isSelection: true,
      docName: docName,
      paragraphs: [{ index: 0, text: selClean, style: '' }]
    };
  }

  // ── 全文模式：计算字符数决定路径 ──────────────────────────────────────────
  var fullText = (doc.Content && doc.Content.Text) ? doc.Content.Text : '';
  var charCount = fullText.replace(/\r/g, '').length;
  var THRESHOLD = 15000;
  var path = (charCount > THRESHOLD) ? 'schedule' : 'direct';

  // ── 遍历段落 ──────────────────────────────────────────────────────────────
  var paras = doc.Paragraphs;
  var totalParas = (paras && paras.Count) ? paras.Count : 0;
  var paraList = [];

  for (var i = 1; i <= totalParas; i++) {
    var para = paras.Item(i);
    if (!para) { continue; }

    var range = para.Range;
    var text = (range && range.Text) ? range.Text.replace(/\r/g, '') : '';
    if (text.trim().length === 0) { continue; }

    var styleName = '';
    try {
      var st = para.Style;
      styleName = (st && st.NameLocal) ? st.NameLocal : '';
    } catch (stErr) {}

    paraList.push({ index: i, text: text, style: styleName });
  }

  return {
    success: true,
    path: path,
    charCount: charCount,
    isSelection: false,
    docName: docName,
    paragraphs: paraList
  };

} catch (e) {
  return { success: false, error: String(e) };
}
