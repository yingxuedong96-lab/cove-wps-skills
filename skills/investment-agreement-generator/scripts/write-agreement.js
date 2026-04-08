/**
 * write-agreement.js
 * 创建新文档并写入增资协议内容（简版模式）。
 * 仅输出15条正文条款，不包含封面、签署页等。
 *
 * 入参: content (string) — 完整的增资协议文本
 * 出参: { success: boolean, message: string, docName: string }
 */

try {
  var agreementContent = typeof content !== 'undefined' ? content : '';
  
  if (!agreementContent) {
    return { success: false, message: '协议内容为空' };
  }

  // 创建新文档
  var newDoc = Application.Documents.Add();
  if (!newDoc) {
    return { success: false, message: '无法创建新文档' };
  }

  // 先整体写入所有内容（不加粗）
  newDoc.Content.InsertAfter(agreementContent);
  
  // 设置全文字体为微软雅黑、小四号（12pt），并取消加粗
  try {
    newDoc.Content.Font.Name = '微软雅黑';
    newDoc.Content.Font.Size = 12;
    newDoc.Content.Font.Bold = false;
  } catch (e) {
    console.warn('[write-agreement] set font failed:', e);
  }
  
  // 识别章节标题并单独设置加粗
  var lines = agreementContent.split('\n');
  var totalLength = 0;
  
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var lineLen = line.length;
    
    // 判断是否是章节标题（以"第"开头且包含"条"，格式如"第一条 定义"或"第十五条 其他"）
    var isTitle = false;
    if (line.indexOf('第') === 0 && line.indexOf('条') > 0) {
      // 进一步判断：标题通常较短，且"条"字后跟空格或冒号
      var tiaoPos = line.indexOf('条');
      var afterTiao = line.substring(tiaoPos + 1);
      // 标题格式："第X条 标题名"，"条"后通常是空格或直接是内容
      if (afterTiao.length > 0 && (afterTiao.charAt(0) === ' ' || afterTiao.charAt(0) === '　')) {
        isTitle = true;
      } else if (lineLen <= 20) {
        // 短行且符合格式，也视为标题
        isTitle = true;
      }
    }
    
    if (isTitle) {
      try {
        // 计算该行在文档中的起始和结束位置
        var start = totalLength;
        var end = totalLength + lineLen;
        var titleRange = newDoc.Range(start, end);
        titleRange.Font.Bold = true;
      } catch (e) {
        console.warn('[write-agreement] set title bold failed:', e);
      }
    }
    
    // 累加长度（包括换行符）
    totalLength += lineLen + 1; // +1 for newline
  }

  // 获取新文档名称
  var docName = newDoc.Name || '增资协议.docx';

  return { success: true, message: '增资协议已生成新文档', docName: docName };
} catch (e) {
  console.warn('[write-agreement]', e);
  return { success: false, message: String(e) };
}
