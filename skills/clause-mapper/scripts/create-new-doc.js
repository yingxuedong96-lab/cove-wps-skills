/**
 * create-new-doc.js
 * 创建新文档并写入标准条款内容。
 *
 * 入参: 
 *   preamble (string) - 协议前言内容
 *   clauses (array) - 标准条款数组，格式：[{ title: string, content: string }]
 * 
 * 出参: { success: boolean, docName: string }
 */

try {
  // 参数接收
  var preambleText = typeof preamble !== 'undefined' ? preamble : '';
  var clausesList = typeof clauses !== 'undefined' ? clauses : [];
  
  if (!clausesList || clausesList.length === 0) {
    return { success: false, docName: '' };
  }

  // 创建新文档
  var newDoc = Application.Documents.Add();
  if (!newDoc) {
    return { success: false, docName: '' };
  }

  var range = newDoc.Content;
  range.Collapse(1); // wdCollapseStart

  // 写入前言
  if (preambleText && preambleText.length > 0) {
    range.InsertAfter(preambleText);
    range.InsertAfter('\r\n\r\n');
    range = newDoc.Content;
    range.Collapse(0); // wdCollapseEnd
  }

  // 写入条款
  for (var i = 0; i < clausesList.length; i++) {
    var clause = clausesList[i];
    
    if (!clause || !clause.content) {
      continue;
    }

    // 写入条款标题（标题已包含中文编号，如"第一条 定义"）
    if (clause.title) {
      range.InsertAfter(clause.title + '\r\n');
    }

    // 写入条款内容
    range.InsertAfter(clause.content + '\r\n\r\n');
    
    // 移动到文档末尾
    range = newDoc.Content;
    range.Collapse(0);
  }

  // 生成文档名称
  var now = new Date();
  var timestamp = now.getFullYear() + 
    String(now.getMonth() + 1).padStart(2, '0') + 
    String(now.getDate()).padStart(2, '0') + '_' +
    String(now.getHours()).padStart(2, '0') + 
    String(now.getMinutes()).padStart(2, '0');
  
  var docName = '标准条款_' + timestamp;
  
  // 保存文档
  try {
    // 尝试获取默认文档路径
    var docPath = Application.DefaultFilePath;
    if (docPath) {
      newDoc.SaveAs2(docPath + '\\' + docName + '.docx');
    }
  } catch (saveErr) {
    console.warn('[create-new-doc] SaveAs failed:', saveErr);
    // 保存失败不影响返回
  }

  return { 
    success: true, 
    docName: docName 
  };
} catch (e) {
  console.warn('[create-new-doc]', e);
  return { success: false, docName: '' };
}
