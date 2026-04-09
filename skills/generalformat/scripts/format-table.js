/**
 * 排版表格 - 支持动态解析用户输入
 * 默认：表名黑体小五号居中，表格等宽，跨页重复表头，表头黑体五号居中加粗，表格内容宋体五号靠左
 */
try {
  var doc = Application.ActiveDocument;
  if (!doc) return { success: false, error: '无活动文档' };

  console.log('[format-表格] 开始处理');

  // 字号映射
  var FONT_SIZE_MAP = {
    '初号': 42, '小初': 36, '一号': 26, '小一': 24,
    '二号': 22, '小二': 18, '三号': 16, '小三': 15,
    '四号': 14, '小四': 12, '五号': 10.5, '小五': 9,
    '六号': 7.5, '小六': 6.5
  };

  // 解析字号
  function parseSize(text) {
    var match = text.match(/(初号|小初|一号|小一|二号|小二|三号|小三|四号|小四|五号|小五|六号|小六)/);
    var result = match ? FONT_SIZE_MAP[match[1]] : 0;
    console.log('[format-表格] parseSize: text=' + text + ', match=' + (match ? match[1] : 'null') + ', result=' + result);
    return result;
  }

  // 解析字体
  function parseFont(text) {
    if (text.indexOf('黑体') !== -1) return '黑体';
    if (text.indexOf('宋体') !== -1) return '宋体';
    if (text.indexOf('楷体') !== -1) return '楷体';
    if (text.indexOf('仿宋') !== -1) return '仿宋';
    return null;
  }

  // 解析对齐
  function parseAlign(text) {
    if (text.indexOf('居中') !== -1) return 1;
    if (text.indexOf('右对齐') !== -1 || text.indexOf('靠右') !== -1) return 2;
    if (text.indexOf('两端对齐') !== -1) return 3;
    if (text.indexOf('左对齐') !== -1 || text.indexOf('靠左') !== -1) return 0;
    return -1;
  }

  // 解析加粗
  function parseBold(text) {
    return text.indexOf('加粗') !== -1;
  }

  // 从用户输入解析规则
  var userInput = specText || '';
  console.log('[format-表格] 用户输入: ' + userInput);

  // 默认配置
  var rules = {
    tableCaption: { fontCN: '黑体', fontSize: 9, alignment: 1 },
    tableHeader: { fontCN: '黑体', fontSize: 10.5, alignment: 1, bold: true },
    tableContent: { fontCN: '宋体', fontSize: 10.5, alignment: 0 }
  };

  // 尝试从用户输入解析表名规则
  if (userInput.indexOf('表名') !== -1) {
    var captionMatch = userInput.match(/表名[^。；,，]*/);
    if (captionMatch) {
      var ct = captionMatch[0];
      if (parseFont(ct)) rules.tableCaption.fontCN = parseFont(ct);
      if (parseSize(ct)) rules.tableCaption.fontSize = parseSize(ct);
      if (parseAlign(ct) >= 0) rules.tableCaption.alignment = parseAlign(ct);
      console.log('[format-表格] 解析表名规则: ' + JSON.stringify(rules.tableCaption));
    }
  }

  // 尝试从用户输入解析表头规则
  if (userInput.indexOf('表头') !== -1) {
    var headerMatch = userInput.match(/表头[^。；,，]*/);
    if (headerMatch) {
      var ht = headerMatch[0];
      console.log('[format-表格] 表头匹配文本: ' + ht);
      var parsedFont = parseFont(ht);
      var parsedSize = parseSize(ht);
      var parsedAlign = parseAlign(ht);
      console.log('[format-表格] 解析结果: font=' + parsedFont + ', size=' + parsedSize + ', align=' + parsedAlign);
      if (parsedFont) rules.tableHeader.fontCN = parsedFont;
      if (parsedSize) rules.tableHeader.fontSize = parsedSize;
      if (parsedAlign >= 0) rules.tableHeader.alignment = parsedAlign;
      if (parseBold(ht)) rules.tableHeader.bold = true;
      console.log('[format-表格] 解析表头规则: ' + JSON.stringify(rules.tableHeader));
    }
  }

  // 尝试从用户输入解析表格内容规则
  if (userInput.indexOf('表格内容') !== -1) {
    var contentMatch = userInput.match(/表格内容[^。；,，]*/);
    if (contentMatch) {
      var cnt = contentMatch[0];
      if (parseFont(cnt)) rules.tableContent.fontCN = parseFont(cnt);
      if (parseSize(cnt)) rules.tableContent.fontSize = parseSize(cnt);
      if (parseAlign(cnt) >= 0) rules.tableContent.alignment = parseAlign(cnt);
      console.log('[format-表格] 解析表格内容规则: ' + JSON.stringify(rules.tableContent));
    }
  }

  console.log('[format-表格] 最终规则: ' + JSON.stringify(rules));

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

  // 处理表格
  var tableCount = doc.Tables ? doc.Tables.Count : 0;
  var tableApplied = 0;
  var needFullWidth = userInput.indexOf('等宽') !== -1 || userInput.indexOf('页面等宽') !== -1;
  var needHeadingRepeat = userInput.indexOf('跨页重复') !== -1 || userInput.indexOf('重复表头') !== -1;

  for (var tblIdx = 1; tblIdx <= tableCount; tblIdx++) {
    try {
      var table = doc.Tables.Item(tblIdx);
      if (!table || !table.Rows) continue;

      var rowCount = table.Rows.Count;

      // 表格等宽
      if (needFullWidth) {
        try {
          var section = doc.Sections.Item(1);
          var pageWidth = section.PageSetup.PageWidth;
          var leftMargin = section.PageSetup.LeftMargin;
          var rightMargin = section.PageSetup.RightMargin;
          var usableWidth = pageWidth - leftMargin - rightMargin;

          table.PreferredWidthType = 3;
          table.PreferredWidth = usableWidth;
          table.AllowAutoFit = false;
        } catch (e) {}
      }

      // 跨页重复表头
      if (needHeadingRepeat && rowCount > 0) {
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