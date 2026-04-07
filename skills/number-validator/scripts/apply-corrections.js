// 应用编号修正（修订模式）
try {
  var corrections = typeof corrections !== 'undefined' ? corrections : [];
  var mode = typeof mode !== 'undefined' ? mode : 'apply';

  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, error: '无活动文档' };
  }

  var paras = doc.Paragraphs;

  // 报告模式：在文档末尾插入校对报告
  if (mode === 'report') {
    var reportContent = '\n\n======== 编号校对报告 ========\n';
    reportContent += '校对时间：' + new Date().toLocaleString() + '\n';
    reportContent += '检测标题数：' + corrections.length + ' 个\n';
    reportContent += '--------------------------------\n';

    if (corrections.length === 0) {
      reportContent += '未检测到编号问题，文档编号规范。\n';
    } else {
      var errorTypes = {};
      for (var i = 0; i < corrections.length; i++) {
        var err = corrections[i];
        reportContent += '标题' + err.index + '：' + err.original + '\n';
        reportContent += '  问题：' + err.issue + '\n';
        reportContent += '  建议：' + err.suggested + '\n\n';

        // 统计错误类型
        var type = err.issue;
        if (!errorTypes[type]) {
          errorTypes[type] = 0;
        }
        errorTypes[type]++;
      }

      reportContent += '--------------------------------\n';
      reportContent += '错误类型统计：\n';
      for (var typeKey in errorTypes) {
        reportContent += '  - ' + typeKey + '：' + errorTypes[typeKey] + ' 个\n';
      }
    }

    reportContent += '================================\n';

    // 在文档末尾插入报告
    var endRange = doc.Content.Duplicate;
    endRange.Collapse(0); // wdCollapseEnd
    endRange.InsertAfter(reportContent);

    return {
      success: true,
      reportGenerated: true
    };
  }

  // 修正模式：应用修正建议
  if (corrections.length === 0) {
    return {
      success: true,
      appliedCount: 0,
      skippedCount: 0,
      message: '无需修正'
    };
  }

  // 保存并开启修订模式
  var originalTrackRevisions = doc.TrackRevisions;
  doc.TrackRevisions = true;

  var appliedCount = 0;
  var skippedCount = 0;

  // 从后往前修正，避免位置偏移
  for (var i = corrections.length - 1; i >= 0; i--) {
    var correction = corrections[i];
    var paraIndex = correction.paraIndex;

    if (!paraIndex || paraIndex < 1 || paraIndex > paras.Count) {
      skippedCount++;
      continue;
    }

    try {
      var para = paras.Item(paraIndex);
      if (!para || !para.Range) {
        skippedCount++;
        continue;
      }

      var fullText = para.Range.Text || '';
      var originalText = correction.original;
      var suggestedText = correction.suggested;

      // 替换标题编号
      var newFullText = fullText.replace(originalText, suggestedText);

      if (newFullText !== fullText) {
        para.Range.Text = newFullText;
        appliedCount++;
      } else {
        skippedCount++;
      }
    } catch (e2) {
      skippedCount++;
      console.warn('[apply-corrections] item error:', e2);
    }
  }

  // 恢复修订模式
  doc.TrackRevisions = originalTrackRevisions;

  return {
    success: true,
    appliedCount: appliedCount,
    skippedCount: skippedCount,
    message: '已修正 ' + appliedCount + ' 个标题，跳过 ' + skippedCount + ' 个'
  };

} catch (e) {
  console.warn('[apply-corrections]', e);
  return { success: false, error: String(e) };
}
