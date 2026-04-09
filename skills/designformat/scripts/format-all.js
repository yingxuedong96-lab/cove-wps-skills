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
  // 1. 先设置页面
  // ========================================
  var applied = 0;
  var details = { headings: 0, bodyParas: 0, figures: 0, tables: 0, pages: 0 };

  var PAGE_WIDTH = 595.35;
  var PAGE_HEIGHT = 841.95;
  var MARGIN_TOP = 72;
  var MARGIN_BOTTOM = 72;
  var MARGIN_LEFT = 90;
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

  var usableWidth = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT;
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

    // 3.6 表格格式（参考generalformat的实现）
    try {
      var tableCount = doc.Tables ? doc.Tables.Count : 0;
      console.log('[designformat] 开始处理表格: ' + tableCount + '个');

      for (var tIdx = 1; tIdx <= tableCount; tIdx++) {
        try {
          var table = doc.Tables.Item(tIdx);
          if (!table || !table.Rows) continue;

          var rowCount = table.Rows.Count;
          console.log('[designformat] 表格' + tIdx + ': ' + rowCount + '行');

          // 设置表格宽度
          try {
            table.PreferredWidthType = 3;
            table.PreferredWidth = usableWidth;
          } catch (e) {}
          try { table.AllowAutoFit = false; } catch (e) {}

          // 跨页重复表头
          try {
            if (rowCount > 0) {
              table.Rows.Item(1).HeadingFormat = true;
            }
          } catch (e) {}

          // 【表头】第一行：黑体五号加粗居中（用整行Range）
          try {
            var headerRow = table.Rows.Item(1);
            if (headerRow && headerRow.Range) {
              headerRow.Range.Font.NameFarEast = '黑体';
              headerRow.Range.Font.Name = '黑体';
              headerRow.Range.Font.Size = 10.5;
              headerRow.Range.Font.Bold = -1;
              if (headerRow.Range.ParagraphFormat) {
                headerRow.Range.ParagraphFormat.Alignment = 1;
              }
              console.log('[designformat] 表格' + tIdx + '表头设置完成');
            }
          } catch (e) {}

          // 【内容】第2行起：宋体五号靠左（用Range覆盖多行）
          if (rowCount > 1) {
            try {
              var startRow = table.Rows.Item(2);
              var endRow = table.Rows.Item(rowCount);
              if (startRow && startRow.Range && endRow && endRow.Range) {
                var contentRange = doc.Range(startRow.Range.Start, endRow.Range.End);
                contentRange.Font.NameFarEast = '宋体';
                contentRange.Font.Name = '宋体';
                contentRange.Font.Size = 10.5;
                if (contentRange.ParagraphFormat) {
                  contentRange.ParagraphFormat.Alignment = 0;
                }
                console.log('[designformat] 表格' + tIdx + '内容设置完成(整体Range)');
              }
            } catch (e) {
              // 备用：逐行处理
              console.log('[designformat] 表格' + tIdx + '整体Range失败，逐行处理');
              for (var rIdx = 2; rIdx <= rowCount; rIdx++) {
                try {
                  var row = table.Rows.Item(rIdx);
                  if (row && row.Range) {
                    row.Range.Font.NameFarEast = '宋体';
                    row.Range.Font.Name = '宋体';
                    row.Range.Font.Size = 10.5;
                    if (row.Range.ParagraphFormat) {
                      row.Range.ParagraphFormat.Alignment = 0;
                    }
                  }
                } catch (e2) {}
              }
            }
          }

          applied++;
          details.tables++;
        } catch (e) {
          console.log('[designformat] 表格' + tIdx + '处理异常: ' + e);
        }
      }
      console.log('[designformat] 表格处理完成: ' + details.tables + '个');
    } catch (e) {
      console.log('[designformat] 表格处理整体异常: ' + e);
    }

    // 3.7 页眉页脚
    try {
      var sections = doc.Sections;
      if (sections && sections.Count > 0) {
        for (var si = 1; si <= sections.Count; si++) {
          try {
            var sec = sections.Item(si);
            if (!sec) continue;

            // 页眉
            try {
              var header = sec.Headers.Item(1);
              if (header && header.Range) {
                if (header.Range.Font) {
                  header.Range.Font.NameFarEast = '宋体';
                  header.Range.Font.Name = 'Arial';
                  header.Range.Font.Size = 9;
                }
                if (header.Range.ParagraphFormat) {
                  header.Range.ParagraphFormat.Alignment = 1;
                }
              }
            } catch (e) {}

            // 页脚
            try {
              var footer = sec.Footers.Item(1);
              if (footer && footer.Range) {
                footer.Range.Text = '';
                if (footer.Range.Font) {
                  footer.Range.Font.NameFarEast = '宋体';
                  footer.Range.Font.Name = 'Arial';
                  footer.Range.Font.Size = 9;
                }
                try {
                  var pn = footer.PageNumbers;
                  if (pn) {
                    pn.NumberStyle = 0;
                    pn.Add(1);
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