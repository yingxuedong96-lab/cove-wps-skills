/**
 * insert-summary.js
 * 在文档末尾插入内容，使用公文格式（换页符 + 仿宋_GB2312 + 5号字 + 深蓝色）。
 * 不包含标题，直接插入正文内容。自动清理特殊符号。
 * 强制所有字符（包括数字）使用统一字体格式。
 *
 * 入参: content (string)
 * 出参: { success: boolean }
 */

try {
  var content = typeof content !== 'undefined' ? content : '';

  if (!content) {
    return { success: false };
  }

  // 清理函数：移除 # 和 * 等符号
  function cleanText(text) {
    if (!text) return '';
    return text.replace(/[#*`]/g, '').trim();
  }

  // 清理内容
  content = cleanText(content);

  if (!content) {
    return { success: false };
  }

  var doc = Application.ActiveDocument;
  if (!doc || !doc.Range) {
    return { success: false };
  }

  // 定位到文档结尾
  var endRange = doc.Content;
  endRange.Collapse(0); // wdCollapseEnd

  // 插入分页符（换页符）
  endRange.InsertBreak(7); // wdPageBreak = 7

  // 重新定位到新页面的开头
  var newRange = doc.Content;
  newRange.Collapse(0); // wdCollapseEnd

  // 记录插入前的位置
  var startPos = newRange.Start;

  // 直接插入正文内容，不包含标题
  newRange.InsertAfter(content);

  // 重新选中插入的完整内容范围
  var formatRange = doc.Range(startPos, doc.Content.End);

  // 设置字体格式：仿宋_GB2312，5号字（10.5pt），深蓝色
  if (formatRange && formatRange.Font) {
    // 强制设置所有字体（包括中文字体和西文字体）为统一格式
    formatRange.Font.Name = '仿宋_GB2312';
    formatRange.Font.NameFarEast = '仿宋_GB2312';
    formatRange.Font.NameAscii = '仿宋_GB2312';      // ASCII 字符（包括数字）
    formatRange.Font.NameOther = '仿宋_GB2312';      // 其他西文字符
    formatRange.Font.Size = 10.5; // 5号字对应10.5磅
    // 使用更深蓝色：RGB(0, 0, 80) - 更深
    // BGR 格式：80 * 65536 + 0 * 256 + 0 = 5242880
    formatRange.Font.Color = 5242880;
  }

  // 设置段落格式
  if (formatRange && formatRange.ParagraphFormat) {
    // 首行缩进2字符
    formatRange.ParagraphFormat.FirstLineIndent = 2 * 10.5; // 2个字符 * 字号
  }

  return { success: true };
} catch (e) {
  console.warn('[insert-summary]', e);
  return { success: false };
}
