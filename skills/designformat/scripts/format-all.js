/**
 * format-all.js - 设计文档短文档一键排版
 * 内置标准格式，无需参数，一键完成全部排版
 */
try {
  var startTime = Date.now();
  var doc = Application.ActiveDocument;
  if (!doc) return { success: false, error: '没有打开的文档' };

  var paraCount = doc.Paragraphs.Count || 0;
  console.log('[designformat] 开始一键排版，段落数: ' + paraCount);

  // ========================================
  // 内置格式标准（设计文档规范）
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

  // 标题识别正则
  var HEADING_PATTERNS = {
    zhangTitle: [/^第[一二三四五六七八九十百零]+章/, /^第\s*\d{1,3}\s*章/, /^\d{1,3}\s+[^\d]/],
    heading2: [/^\d+\.\d+\s/],
    heading3: [/^\d+\.\d+\.\d+\s/],
    heading4: [/^\d+\.\d+\.\d+\.\d+\s/],
    heading5: [/^\d+\.\d+\.\d+\.\d+\.\d+\s/]
  };

  // ========================================
  // 辅助函数
  // ========================================
  function cleanText(t) {
    return String(t || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function classifyPara(text) {
    text = cleanText(text);
    if (!text) return 'empty';

    // 图名/表名优先
    if (/^图\s*\d/.test(text) || /^图\s*[A-Z]/.test(text)) return 'figureCaption';
    if (/^表\s*\d/.test(text) || /^表\s*[A-Z]/.test(text)) return 'tableCaption';

    // 标题（从高级到低级检测）
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
      // 字体 - 必须先清除原有格式再设置
      if (range.Font) {
        range.Font.NameFarEast = rule.font;
        range.Font.Name = rule.font;
        range.Font.Size = rule.size;
        // 明确设置加粗/不加粗
        range.Font.Bold = rule.bold ? -1 : 0;
      }
      // 段落
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
  // 1. 分类段落
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

  // 文档标题：第一个非空段落
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

  console.log('[designformat] 分类结果: 标题' + (typeIndices.zhangTitle.length + typeIndices.heading2.length + typeIndices.heading3.length + typeIndices.heading4.length) + ' 正文' + typeIndices.body.length + ' 图' + typeIndices.figureCaption.length + ' 表' + typeIndices.tableCaption.length);

  // ========================================
  // 2. 应用格式
  // ========================================
  var origTrack = false;
  try { origTrack = doc.TrackRevisions; doc.TrackRevisions = false; } catch (e) {}

  var applied = 0;
  var details = { headings: 0, bodyParas: 0, figures: 0, tables: 0, pages: 0 };

  try {
    // 2.1 标题排版
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

    // 2.2 正文排版
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

    // 2.3 图名排版
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

    // 2.4 表名排版
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

    // 2.5 图片居中
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

    // 2.6 表格处理（等宽 + 跨页重复表头）
    try {
      var tableCount = doc.Tables ? doc.Tables.Count : 0;
      // 获取页面宽度（单位：磅）
      var pageWidth = 595.35, leftMargin = 90, rightMargin = 90;
      try {
        var ps = doc.PageSetup;
        if (ps) {
          pageWidth = ps.PageWidth || 595.35;
          leftMargin = ps.LeftMargin || 90;
          rightMargin = ps.RightMargin || 90;
        }
      } catch (e) {}
      // 计算可用宽度（磅）
      var usableWidth = pageWidth - leftMargin - rightMargin;
      console.log('[designformat] 表格可用宽度: ' + usableWidth.toFixed(1) + '磅');

      for (var tIdx = 1; tIdx <= tableCount; tIdx++) {
        try {
          var table = doc.Tables.Item(tIdx);
          if (!table) continue;

          // 等宽设置
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
    } catch (e) {}

    // 2.7 页面设置
    try {
      var sections = doc.Sections;
      if (sections && sections.Count > 0) {
        for (var si = 1; si <= sections.Count; si++) {
          try {
            var sec = sections.Item(si);
            if (!sec || !sec.PageSetup) continue;

            var ps = sec.PageSetup;
            // 页边距（磅）：上下72磅=2.54cm，左右90磅≈3.17cm
            ps.TopMargin = 72;
            ps.BottomMargin = 72;
            ps.LeftMargin = 90;
            ps.RightMargin = 90;
            ps.PageWidth = 595.35;  // A4宽度
            ps.PageHeight = 841.95; // A4高度
            ps.Orientation = 0;     // 纵向

            applied++;
            details.pages++;
          } catch (e) {}
        }
      }
    } catch (e) {}

    // 2.8 页眉页脚
    try {
      var sections = doc.Sections;
      if (sections && sections.Count > 0) {
        for (var si = 1; si <= sections.Count; si++) {
          try {
            var sec = sections.Item(si);
            if (!sec) continue;

            // 页眉：清除原内容，设置格式
            try {
              var header = sec.Headers.Item(1);  // wdHeaderFooterPrimary
              if (header && header.Range) {
                // 清除原有内容
                header.Range.Delete();
                // 设置字体格式（为后续可能的页眉内容准备）
                if (header.Range.Font) {
                  header.Range.Font.NameFarEast = '宋体';
                  header.Range.Font.Name = 'Arial';
                  header.Range.Font.Size = 9;
                }
                // 页眉线
                try {
                  var borders = header.Range.ParagraphFormat.Borders;
                  if (borders) {
                    var bottom = borders.Item(-3);  // wdBorderBottom
                    bottom.LineStyle = 1;
                    bottom.LineWidth = 6;  // 0.5磅 ≈ 8 eighths of a point
                  }
                } catch (e) {}
              }
            } catch (e) {}

            // 页脚：清除原内容，添加页码
            try {
              var footer = sec.Footers.Item(1);  // wdHeaderFooterPrimary
              if (footer && footer.Range) {
                // 清除原有内容
                footer.Range.Delete();
                // 设置字体
                if (footer.Range.Font) {
                  footer.Range.Font.NameFarEast = '宋体';
                  footer.Range.Font.Name = 'Arial';
                  footer.Range.Font.Size = 9;
                }
                // 添加居中页码
                try {
                  var pn = footer.PageNumbers;
                  if (pn) {
                    pn.NumberStyle = 0;  // 阿拉伯数字
                    pn.Add(1);           // 居中
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