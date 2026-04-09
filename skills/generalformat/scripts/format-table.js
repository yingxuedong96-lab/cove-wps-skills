/**
 * 排版表格 - 表名黑体小五号居中，表格等宽，跨页重复表头，表头黑体五号居中加粗，表格内容宋体五号靠左
 */
try {
  var config = {
    specText: "表名用黑体小五号居中，表格与页面等宽，跨页重复表头，表头用黑体五号居中加粗，表格内容用宋体五号靠左对齐",
    paragraphRules: {
      tableCaption: { fontCN: "黑体", fontSize: 9, alignment: 1 },
      tableHeader: { fontCN: "黑体", fontSize: 10.5, alignment: 1, bold: true },
      tableContent: { fontCN: "宋体", fontSize: 10.5, alignment: 0 }
    }
  };

  var doc = Application.ActiveDocument;
  if (!doc) return { success: false, error: '无活动文档' };

  console.log('[format-表格] 开始处理');
  var rules = config.paragraphRules || {};
  var applied = 0;

  // 辅助函数
  function applyRuleToRange(range, rule) {
    if (!range) return false;
    try {
      if (range.Font) {
        if (rule.fontCN) range.Font.NameFarEast = rule.fontCN;
        if (rule.fontSize) range.Font.Size = rule.fontSize;
        if (rule.bold !== undefined) range.Font.Bold = rule.bold ? -1 : 0;
      }
      if (range.ParagraphFormat) {
        if (rule.alignment !== undefined) range.ParagraphFormat.Alignment = rule.alignment;
      }
      return true;
    } catch (e) { return false; }
  }

  // 处理表名段落
  if (rules.tableCaption) {
    var tablePattern = /^表\s*\d+/;
    var fullText = doc.Content.Text || '';
    var lines = fullText.split('\r');
    var captionCount = 0;

    for (var i = 0; i < lines.length; i++) {
      var text = String(lines[i]).replace(/[\r\u0007]/g, '').trim();
      if (tablePattern.test(text)) {
        try {
          var para = doc.Paragraphs.Item(i + 1);
          if (para && para.Range) {
            if (applyRuleToRange(para.Range, rules.tableCaption)) {
              captionCount++;
            }
          }
        } catch (e) {}
      }
    }
    applied += captionCount;
    console.log('[format-表格] 表名处理: ' + captionCount + '个');
  }

  // 处理表格
  var tableCount = doc.Tables ? doc.Tables.Count : 0;
  var tableApplied = 0;

  for (var tblIdx = 1; tblIdx <= tableCount; tblIdx++) {
    try {
      var table = doc.Tables.Item(tblIdx);
      if (!table || !table.Rows) continue;

      var rowCount = table.Rows.Count;

      // 表格等宽
      try {
        var section = doc.Sections.Item(1);
        var pageWidth = section.PageSetup.PageWidth;
        var leftMargin = section.PageSetup.LeftMargin;
        var rightMargin = section.PageSetup.RightMargin;
        var usableWidth = pageWidth - leftMargin - rightMargin;

        table.PreferredWidthType = 3;  // wdPreferredWidthPoints
        table.PreferredWidth = usableWidth;
        table.AllowAutoFit = false;
      } catch (e) {}

      // 跨页重复表头
      if (rowCount > 0) {
        try {
          table.Rows.Item(1).HeadingFormat = true;
        } catch (e) {}
      }

      // 表头格式
      if (rules.tableHeader && rowCount > 0) {
        try {
          var headerRow = table.Rows.Item(1);
          if (headerRow && headerRow.Range) {
            applyRuleToRange(headerRow.Range, rules.tableHeader);
          }
        } catch (e) {}
      }

      // 表格内容格式
      if (rules.tableContent && rowCount > 1) {
        try {
          var startRow = table.Rows.Item(2);
          var endRow = table.Rows.Item(rowCount);
          if (startRow && startRow.Range && endRow && endRow.Range) {
            var contentRange = doc.Range(startRow.Range.Start, endRow.Range.End);
            applyRuleToRange(contentRange, rules.tableContent);
          }
        } catch (e) {
          // 逐行处理
          for (var r = 2; r <= rowCount; r++) {
            try {
              var row = table.Rows.Item(r);
              if (row && row.Range) {
                applyRuleToRange(row.Range, rules.tableContent);
              }
            } catch (e2) {}
          }
        }
      }

      tableApplied++;
    } catch (e) {}
  }

  applied += tableApplied;
  console.log('[format-表格] 表格处理: ' + tableApplied + '个');

  return { success: true, applied: applied, tables: tableApplied };
} catch (e) {
  return { success: false, error: String(e) };
}