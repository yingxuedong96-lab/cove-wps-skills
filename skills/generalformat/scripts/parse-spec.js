/**
 * parse-spec.js - 读取规范文档内容（带大文档保护）
 *
 * 入参:
 *   specDocPath: 规范文档路径（可选，不提供则使用当前文档）
 *   specDocContent: 规范文档内容文本（可选，直接提供内容时使用）
 *
 * 出参:
 *   { success, content, error, truncated, length }
 *
 * 用途：读取规范文档全文内容，供 LLM 提取排版规则
 *
 * ⚠️ 大文档保护：超过 50000 字符时截断，防止 UI 卡顿
 */

try {
  var MAX_CONTENT_LENGTH = 50000;  // 最大返回长度，防止卡顿
  var content = '';

  // 方式1：直接提供内容
  if (typeof specDocContent !== 'undefined' && specDocContent) {
    content = String(specDocContent);
    console.log('[parse-spec] 使用直接提供的内容，长度=' + content.length);
  } else {

    // 方式2：从文件路径读取（兼容多种参数名）
    var docPath = '';
    if (typeof specDocPath !== 'undefined' && specDocPath) {
      docPath = specDocPath;
    } else if (typeof specFilePath !== 'undefined' && specFilePath) {
      docPath = specFilePath;
    } else if (typeof specPath !== 'undefined' && specPath) {
      docPath = specPath;
    } else if (typeof path !== 'undefined' && path) {
      docPath = path;
    }

    if (docPath) {
      console.log('[parse-spec] 尝试打开文档: ' + docPath);

      var docs = Application.Documents;
      var wasOpen = false;
      var doc = null;

      // 检查文档是否已打开
      for (var i = 1; i <= docs.Count; i++) {
        var d = docs.Item(i);
        if (d.FullName === docPath || d.Name === docPath) {
          doc = d;
          wasOpen = true;
          break;
        }
      }

      // 未打开则尝试打开
      if (!doc) {
        try {
          doc = docs.Open(docPath, true);  // true = 只读模式
          console.log('[parse-spec] 文档已打开');
        } catch (e) {
          return { success: false, error: '无法打开文档: ' + docPath + ' (' + e + ')' };
        }
      }

      // 读取文档内容
      if (doc && doc.Content) {
        content = doc.Content.Text;
        // 关闭文档（如果是我们打开的）
        if (!wasOpen) {
          try { doc.Close(false); } catch (e) {}  // false = 不保存
          console.log('[parse-spec] 文档已关闭');
        }
      }

      console.log('[parse-spec] 读取完成，长度=' + content.length);
    } else {
      // 方式3：使用当前活动文档
      var activeDoc = Application.ActiveDocument;
      if (activeDoc && activeDoc.Content) {
        content = activeDoc.Content.Text;
        console.log('[parse-spec] 使用当前文档，长度=' + content.length);

        // ⚠️ 检测大文档：提示用户这可能不是规范文档
        if (content.length > 50000) {
          console.log('[parse-spec] ⚠️ 警告：当前文档超过50000字符，可能是目标文档而非规范文档');
          console.log('[parse-spec] ⚠️ 建议：明确指定 specDocPath 参数指向规范文档');
        }
      }
    }
  }

  if (!content) {
    return { success: false, error: '未提供规范文档路径或内容，且无活动文档' };
  }

  // ⚠️ 大文档保护：截断过长内容，防止 UI 卡顿
  var truncated = false;
  var originalLength = content.length;

  if (content.length > MAX_CONTENT_LENGTH) {
    console.log('[parse-spec] ⚠️ 内容过长(' + content.length + ')，截断至 ' + MAX_CONTENT_LENGTH);
    content = content.substring(0, MAX_CONTENT_LENGTH);
    truncated = true;
  }

  // 清理控制字符
  content = content.replace(/\u0007/g, '').replace(/\r/g, '\n');

  return {
    success: true,
    content: content,
    truncated: truncated,
    originalLength: originalLength,
    length: content.length
  };

} catch (e) {
  console.error('[parse-spec] 错误: ' + e);
  return { success: false, error: String(e) };
}