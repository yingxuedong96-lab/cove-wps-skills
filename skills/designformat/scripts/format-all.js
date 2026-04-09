/**
 * format-all.js - 设计文档短文档一键排版
 */
try {
  var startTime = Date.now();
  var doc = Application.ActiveDocument;
  if (!doc) return { success: false, error: '没有打开的文档' };

  var paraCount = doc.Paragraphs.Count || 0;
  console.log('[designformat] 开始一键排版，段落数: ' + paraCount);

  // ========================================
  // 内置格式标准
  // ========================================
  var FONT_SIZE = {
    '二号': 22, '小二': 18, '三号': 16, '小三': 15,
    '四号': 14, '小四': 12, '五号': 10.5, '小五': 9
  };

  var FORMAT_RULES = {
    docTitle: { font: '黑体', size: FONT_SIZE['二号'], align: 1, bold: true },
    zhangTitle: { font: '黑体', size: FONT_SIZE['三号'], align: 0, bold: true },
    heading2: { font: '黑体', size: FONT_SIZE['小三'], align: 0, bold: true },
    heading3: { font: '黑体', size: FONT_SIZE['四号'], align: 0, bold: true },
    heading4: { font: '黑体', size: FONT_SIZE['四号'], align: 0, bold: false },
    heading5: { font: '黑体', size: FONT_SIZE['四号'], align: 0, bold: false },
    body: { font: '宋体', size: FONT_SIZE['小四'], align: 3, indent: 24, lineSpacing: 22 },
    tableCaption: { font: '黑体', size: FONT_SIZE['小五'], align: 1, bold: false },
    figureCaption: { font: '黑体', size: FONT_SIZE['小五'], align: 1, bold: false }
  };

  var HEADING_PATTERNS = {
    zhangTitle: [/^第[一二三四五六七八九十百零]+章/, /^第\s*\d{1,3}\s*章/, /^\d{1,3}\s+[^\d]/],
    heading2: [/^\d+\.\d+\s/],
    heading3: [/^\d+\.\d+\.\d+\s/],
    heading4: [/^\d+\.\d+\.\d+\.\d+\s/],
    heading5: [/^\d+\.\d+\.\d+\.\d+\.\d+\s/]
  };

  function cleanText(t) {
    return String(t || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function classifyPara(text) {
    text = cleanText(text);
    if (!text) return 'empty';
    if (/^图\s*\d/.test(text) || /^图\s*[A-Z]/.test(text)) return 'figureCaption';
    if (/^表\s*\d/.test(text) || /^表\s*[A-Z]/.test(text)) return 'tableCaption';
    for (var level in HEADING_PATTERNS) {
      var patterns = HEADING_PATTERNS[level];
      for (var i = 0; i < patterns.length; i++) {
        if (patterns[i].test(text)) return level;
      }
    }
    return text.length <= 2 ? 'short' : 'body';
  }

  function applyFormat(range, rule) {
    if (!range) return false;
    try {
      if (range.Font) {
        range.Font.NameFarEast = rule.font;
        range.Font.Name = rule.font;
        range.Font.Size = rule.size;
        range.Font.Bold = rule.bold ? -1 : 0;
      }
      if (range.ParagraphFormat) {
        range.ParagraphFormat.Alignment = rule.align;
        if (rule.indent) range.ParagraphFormat.FirstLineIndent = rule.indent;
        if (rule.lineSpacing) {
          range.ParagraphFormat.LineSpacingRule = 4;
          range.ParagraphFormat.LineSpacing = rule.lineSpacing;
        }
      }
      return true;
    } catch (e) { return false; }
  }

  // ========================================
  // 1. 先设置页面（这样表格宽度计算才能用正确的值）
  // ========================================
  var applied = 0;
  var details = { headings: 0, bodyParas: 0, figures: 0, tables: 0, pages: 0 };

  // 固定的页面参数
  var PAGE_WIDTH = 595.35;   // A4宽度（磅）
  var PAGE_HEIGHT = 841.95;  // A4高度（磅）
  var MARGIN_TOP = 72;       // 2.54cm
  var MARGIN_BOTTOM = 72;
  var MARGIN_LEFT = 90;      // 约3.17cm
  var MARGIN_RIGHT = 90;

  try {
    var sections = doc.Sections;
    if (sections && sections.Count > 0) {
      for (var si = 1; si <= sections.Count; si++) {
        try {
          var sec = sections.Item(si);
          if (!sec || !sec.PageSetup) continue;

          var ps = sec.PageSetup;
          ps.TopMargin = MARGIN_TOP;
          ps.BottomMargin = MARGIN_BOTTOM;
          ps.LeftMargin = MARGIN_LEFT;
          ps.RightMargin = MARGIN_RIGHT;
          ps.PageWidth = PAGE_WIDTH;
          ps.PageHeight = PAGE_HEIGHT;
          ps.Orientation = 0;

          applied++;
          details.pages++;
        } catch (e) {}
      }
      console.log('[designformat] 页面设置完成: ' + details.pages + '节');
    }
  } catch (e) {}

  // 计算表格可用宽度
  var usableWidth = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT;  // 595.35 - 90 - 90 = 415.35
  console.log('[designformat] 表格可用宽度: ' + usableWidth.toFixed(1) + '磅');

  // ========================================
  // 2. 分类段落
  // ========================================
  var docText = doc.Content ? String(doc.Content.Text) : '';
  var paras = docText.split('\r');
  var classifications = [];
  var typeIndices = { docTitle: [], zhangTitle: [], heading2: [], heading3: [], heading4: [], heading5: [], body: [], tableCaption: [], figureCaption: [] };

  for (var i = 0; i < Math.min(paraCount, paras.length); i++) {
    var type = classifyPara(paras[i]);
    classifications.push(type);
    if (type !== 'empty' && type !== 'short' && typeIndices[type]) {
      typeIndices[type].push(i + 1);
    }
  }

  // 文档标题
  if (typeIndices.docTitle && typeIndices.body.length > 0) {
    var firstBody = typeIndices.body[0];
    var isTitle = true;
    for (var t in typeIndices) {
      if (t !== 'docTitle' && t !== 'body' && typeIndices[t].indexOf(firstBody) !== -1) {
        isTitle = false;
        break;
      }
    }
    if (isTitle && typeIndices.zhangTitle.indexOf(firstBody) === -1) {
      typeIndices.docTitle.push(firstBody);
      typeIndices.body = typeIndices.body.filter(function(idx) { return idx !== firstBody; });
    }
  }

  // ========================================
  // 3. 应用格式
  // ========================================
  var origTrack = false;
  try { origTrack = doc.TrackRevisions; doc.TrackRevisions = false; } catch (e) {}

  try {
    // 3.1 标题
    var headingTypes = ['docTitle', 'zhangTitle', 'heading2', 'heading3', 'heading4', 'heading5'];
    var outlineLevelMap = { docTitle: 0, zhangTitle: 1, heading2: 2, heading3: 3, heading4: 4, heading5: 5 };

    for (var h = 0; h < headingTypes.length; h++) {
      var hType = headingTypes[h];
      var indices = typeIndices[hType];
      var rule = FORMAT_RULES[hType];
      if (!rule || indices.length === 0) continue;

      for (var i = 0; i < indices.length; i++) {
        try {
          var para = doc.Paragraphs.Item(indices[i]);
          if (para && para.Range) {
            if (applyFormat(para.Range, rule)) {
              applied++;
              details.headings++;
            }
            if (outlineLevelMap[hType] && para.Range.ParagraphFormat) {
              try { para.Range.ParagraphFormat.OutlineLevel = outlineLevelMap[hType]; } catch (e) {}
            }
          }
        } catch (e) {}
      }
    }

    // 3.2 正文
    var bodyIndices = typeIndices.body;
    if (bodyIndices.length > 0 && FORMAT_RULES.body) {
      var segments = [];
      var segStart = bodyIndices[0], segEnd = bodyIndices[0];
      for (var i = 1; i < bodyIndices.length; i++) {
        if (bodyIndices[i] === segEnd + 1) {
          segEnd = bodyIndices[i];
        } else {
          segments.push({ start: segStart, end: segEnd });
          segStart = bodyIndices[i];
          segEnd = bodyIndices[i];
        }
      }
      segments.push({ start: segStart, end: segEnd });

      for (var s = 0; s < segments.length; s++) {
        try {
          var startPara = doc.Paragraphs.Item(segments[s].start);
          var endPara = doc.Paragraphs.Item(segments[s].end);
          if (startPara && startPara.Range && endPara && endPara.Range) {
            var segRange = doc.Range(startPara.Range.Start, endPara.Range.End);
            if (applyFormat(segRange, FORMAT_RULES.body)) {
              applied += segments[s].end - segments[s].start + 1;
              details.bodyParas += segments[s].end - segments[s].start + 1;
            }
          }
        } catch (e) {}
      }
    }

    // 3.3 图名
    var figIndices = typeIndices.figureCaption;
    if (figIndices.length > 0 && FORMAT_RULES.figureCaption) {
      for (var i = 0; i < figIndices.length; i++) {
        try {
          var para = doc.Paragraphs.Item(figIndices[i]);
          if (para && para.Range && applyFormat(para.Range, FORMAT_RULES.figureCaption)) {
            applied++;
            details.figures++;
          }
        } catch (e) {}
      }
    }

    // 3.4 表名
    var tblCapIndices = typeIndices.tableCaption;
    if (tblCapIndices.length > 0 && FORMAT_RULES.tableCaption) {
      for (var i = 0; i < tblCapIndices.length; i++) {
        try {
          var para = doc.Paragraphs.Item(tblCapIndices[i]);
          if (para && para.Range && applyFormat(para.Range, FORMAT_RULES.tableCaption)) {
            applied++;
            details.tables++;
          }
        } catch (e) {}
      }
    }

    // 3.5 图片居中
    try {
      var inlineCount = doc.InlineShapes ? doc.InlineShapes.Count : 0;
      for (var imgIdx = 1; imgIdx <= inlineCount; imgIdx++) {
        try {
          var shape = doc.InlineShapes.Item(imgIdx);
          if (shape && shape.Range && shape.Range.ParagraphFormat) {
            shape.Range.ParagraphFormat.Alignment = 1;
            shape.Range.ParagraphFormat.LineSpacingRule = 0;
            applied++;
            details.figures++;
          }
        } catch (e) {}
      }
    } catch (e) {}

    // 3.6 表格（使用固定计算的宽度）
    try {
      var tableCount = doc.Tables ? doc.Tables.Count : 0;
      for (var tIdx = 1; tIdx <= tableCount; tIdx++) {
        try {
          var table = doc.Tables.Item(tIdx);
          if (!table) continue;

          // 设置表格宽度为可用宽度
          try {
            table.PreferredWidthType = 3;  // wdPreferredWidthPoints
            table.PreferredWidth = usableWidth;
          } catch (e) {}
          try { table.AllowAutoFit = false; } catch (e) {}

          // 跨页重复表头
          try {
            if (table.Rows && table.Rows.Count > 0) {
              table.Rows.Item(1).HeadingFormat = true;
            }
          } catch (e) {}

          applied++;
          details.tables++;
        } catch (e) {}
      }
      console.log('[designformat] 表格: ' + tableCount + '个, 宽度=' + usableWidth.toFixed(1) + '磅');
    } catch (e) {}

    // 3.6.1 表格内容格式（逐单元格处理）
    try {
      var tableCount = doc.Tables ? doc.Tables.Count : 0;
      for (var tIdx = 1; tIdx <= tableCount; tIdx++) {
        try {
          var table = doc.Tables.Item(tIdx);
          if (!table || !table.Rows || !table.Columns) continue;

          var rowCount = table.Rows.Count;
          var colCount = table.Columns.Count;

          // 表头（第一行）：逐单元格，黑体五号加粗居中
          for (var cIdx = 1; cIdx <= colCount; cIdx++) {
            try {
              var cell = table.Cell(1, cIdx);
              if (cell && cell.Range) {
                cell.Range.Font.Reset();
                cell.Range.Font.NameFarEast = '黑体';
                cell.Range.Font.Name = '黑体';
                cell.Range.Font.Size = 10.5;
                cell.Range.Font.Bold = -1;
                if (cell.Range.ParagraphFormat) {
                  cell.Range.ParagraphFormat.Alignment = 1;  // 居中
                }
              }
            } catch (e) {}
          }

          // 表格内容（第2行起）：逐单元格，宋体五号不加粗靠左
          if (rowCount > 1) {
            for (var rIdx = 2; rIdx <= rowCount; rIdx++) {
              for (var cIdx = 1; cIdx <= colCount; cIdx++) {
                try {
                  var cell = table.Cell(rIdx, cIdx);
                  if (cell && cell.Range) {
                    cell.Range.Font.Reset();
                    cell.Range.Font.NameFarEast = '宋体';
                    cell.Range.Font.Name = '宋体';
                    cell.Range.Font.Size = 10.5;
                    cell.Range.Font.Bold = 0;
                    if (cell.Range.ParagraphFormat) {
                      cell.Range.ParagraphFormat.Alignment = 0;  // 靠左
                    }
                  }
                } catch (e) {}
              }
            }
          }
        } catch (e) {}
      }
      console.log('[designformat] 表格内容格式(单元格级): ' + tableCount + '个');
    } catch (e) {}

    // 3.7 页眉页脚
    try {
      var sections = doc.Sections;
      if (sections && sections.Count > 0) {
        for (var si = 1; si <= sections.Count; si++) {
          try {
            var sec = sections.Item(si);
            if (!sec) continue;

            // 页眉：设置格式（不清除内容）
            try {
              var header = sec.Headers.Item(1);
              if (header && header.Range) {
                // 设置字体
                if (header.Range.Font) {
                  header.Range.Font.NameFarEast = '宋体';
                  header.Range.Font.Name = 'Arial';
                  header.Range.Font.Size = 9;
                }
                // 居中对齐
                if (header.Range.ParagraphFormat) {
                  header.Range.ParagraphFormat.Alignment = 1;  // 居中
                }
                // 页眉线
                try {
                  var borders = header.Range.ParagraphFormat.Borders;
                  if (borders) {
                    var bottom = borders.Item(-3);
                    bottom.LineStyle = 1;
                    bottom.LineWidth = 6;
                  }
                } catch (e) {}
              }
            } catch (e) {}

            // 页脚：设置页码居中
            try {
              var footer = sec.Footers.Item(1);
              if (footer && footer.Range) {
                // 清除原有内容
                footer.Range.Text = '';
                // 设置字体
                if (footer.Range.Font) {
                  footer.Range.Font.NameFarEast = '宋体';
                  footer.Range.Font.Name = 'Arial';
                  footer.Range.Font.Size = 9;
                }
                // 添加页码
                try {
                  var pn = footer.PageNumbers;
                  if (pn) {
                    pn.NumberStyle = 0;
                    pn.Add(1);  // 居中
                  }
                } catch (e) {}
              }
            } catch (e) {}

            applied++;
          } catch (e) {}
        }
      }
    } catch (e) {}

  } finally {
    try { doc.TrackRevisions = origTrack; } catch (e) {}
  }

  var elapsed = Date.now() - startTime;
  console.log('[designformat] 完成: applied=' + applied + ', time=' + elapsed + 'ms');
  console.log('[designformat] 详情: 标题' + details.headings + ' 正文' + details.bodyParas + ' 图' + details.figures + ' 表' + details.tables + ' 页面' + details.pages);

  return {
    success: true,
    applied: applied,
    details: details,
    elapsedMs: elapsed
  };

} catch (e) {
  console.error('[designformat] 错误: ' + e);
  return { success: false, error: String(e) };
}