// init-schedule.js
// 大文档模式：将全文段落切分为若干片段，返回 ScheduleInitResult
// 每个 item 包含一批段落文本，子 Agent 分析后输出 corrections JSON 数组

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, error: '无活动文档' };
  }

  var docName = doc.Name ? doc.Name : '未知文档';
  var paras = doc.Paragraphs;
  var totalParas = (paras && paras.Count) ? paras.Count : 0;

  var CHUNK_SIZE = 3000;  // 每片约 3000 字
  var chunks = [];
  var bufText = '';
  var bufItems = [];

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

    bufItems.push({ index: i, text: text, style: styleName });
    bufText += text;

    // 达到分片大小或遍历到末尾时，生成一个 chunk
    if (bufText.length >= CHUNK_SIZE || i === totalParas) {
      var chunkIndex = chunks.length;
      chunks.push({
        id: 'chunk_' + chunkIndex,
        payload: {
          chunkIndex: chunkIndex,
          paragraphs: bufItems
        }
      });
      bufText = '';
      bufItems = [];
    }
  }

  // 构造 ScheduleInitResult
  return {
    success: true,
    taskType: 'docx-proofreader',
    docName: docName,
    totalChunks: chunks.length,
    items: chunks
  };

} catch (e) {
  return { success: false, error: String(e) };
}
