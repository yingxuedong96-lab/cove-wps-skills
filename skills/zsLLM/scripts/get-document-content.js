/**
 * get-document-content.js
 * 读取文档全文内容，用于传递给 LLM 进行语义分析
 *
 * 入参:
 *   maxLength (number) - 最大字符数，默认 10000
 *
 * 出参: { content, paraCount, charCount }
 */

try {
  var maxLen = typeof maxLength !== 'undefined' ? maxLength : 10000;

  var doc = Application.ActiveDocument;
  if (!doc) {
    console.log('[get-document-content] 无活动文档');
    return { error: '无活动文档', content: '', paraCount: 0, charCount: 0 };
  }

  console.log('[get-document-content] 开始读取文档，最大长度: ' + maxLen);

  var content = '';
  var paraCount = doc.Paragraphs.Count;
  var maxParas = 500;  // 最多读取500个段落
  var limit = paraCount < maxParas ? paraCount : maxParas;

  // 使用连续的非空段落编号
  var nonEmptyIndex = 0;

  for (var i = 1; i <= limit; i++) {
    try {
      var para = doc.Paragraphs.Item(i);
      if (!para || !para.Range) continue;

      var text = para.Range.Text || '';
      text = String(text).replace(/\r/g, '').replace(/\n/g, '').trim();

      if (text.length === 0) continue;

      // 使用连续编号，方便 LLM 定位
      nonEmptyIndex++;
      content += '【第' + nonEmptyIndex + '段】' + text + '\n';

      // 同时记录原始段落索引，用于精确定位
      // 格式：段落索引|内容

      // 检查是否超过最大长度
      if (content.length >= maxLen) {
        content = content.substring(0, maxLen) + '...(内容已截断)';
        break;
      }
    } catch (e) {}
  }

  var charCount = content.length;

  console.log('[get-document-content] 读取完成，非空段落数: ' + nonEmptyIndex + '，字符数: ' + charCount);

  return {
    content: content,
    paraCount: nonEmptyIndex,
    charCount: charCount
  };

} catch (e) {
  console.warn('[get-document-content]', e);
  return { error: String(e), content: '', paraCount: 0, charCount: 0 };
}