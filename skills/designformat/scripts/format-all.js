/**
 * format-all.js - 设计文档短文档一键排版
 * 内置标准格式，无需参数，一键完成全部排版
 * 适合10页以内的设计报告
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
    docTitle: { font: '黑体', size: FONT_SIZE['二号'], align: 1, bold: true },      // 文档标题
    zhangTitle: { font: '黑体', size: FONT_SIZE['三号'], align: 0, bold: true },     // 一级标题
    heading2: { font: '黑体', size: FONT_SIZE['小三'], align: 0, bold: true },       // 二级标题
    heading3: { font: '黑体', size: FONT_SIZE['四号'], align: 0, bold: true },       // 三级标题
    heading4: { font: '黑体', size: FONT_SIZE['四号'], align: 0, bold: false },      // 四级标题
    heading5: { font: '黑体', size: FONT_SIZE['四号'], align: 0, bold: false },      // 五级标题
    body: { font: '宋体', size: FONT_SIZE['小四'], align: 3, indent: 24, lineSpacing: 22 }, // 正文
    tableCaption: { font: '黑体', size: FONT_SIZE['小五'], align: 1, bold: false },  // 表名
    figureCaption: { font: '黑体', size: FONT_SIZE['小五'], align: 1, bold: false }  // 图名
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
      // 字体
      if (range.Font) {
        range.Font.NameFarEast = rule.font;
        range.Font.Name = rule.font;
        range.Font.Size = rule.size;
        if (rule.bold !== undefined) range.Font.Bold = rule.bold ? -1 : 0;
      }
      // 段落
      if (range.ParagraphFormat) {
        range.ParagraphFormat.Alignment = rule.align;
        if (rule.indent) range.ParagraphFormat.FirstLineIndent = rule.indent;
        if (rule.lineSpacing) {
          range.ParagraphFormat.LineSpacingRule = 4; // 固定行距
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
    // 检查是否被其他标题类型捕获
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

  var counts = {};
  for (var i = 0; i < classifications.length; i++) {
    counts[classifications[i]] = (counts[classifications[i]] || 0) + 1;
  }
  console.log('[designformat] 分类结果: ' + JSON.stringify(counts));

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
            // 设置大纲级别
            if (outlineLevelMap[hType] && para.Range.ParagraphFormat) {
              try { para.Range.ParagraphFormat.OutlineLevel = outlineLevelMap[hType]; } catch (e) {}
            }
          }
        } catch (e) {}
      }
      console.log('[designformat] ' + hType + ': ' + indices.length + '段');
    }

    // 2.2 正文排版（批量Range优化）
    var bodyIndices = typeIndices.body;
    if (bodyIndices.length > 0 && FORMAT_RULES.body) {
      // 构建连续段落段
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

      console.log('[designformat] 正文分段数: ' + segments.length);

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
      console.log('[designformat] 正文: ' + details.bodyParas + '段');
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
      console.log('[designformat] 图名: ' + details.figures + '个');
    }

    // 2.4 表名排版
    var tblIndices = typeIndices.tableCaption;
    if (tblIndices.length > 0 && FORMAT_RULES.tableCaption) {
      for (var i = 0; i < tblIndices.length; i++) {
        try {
          var para = doc.Paragraphs.Item(tblIndices[i]);
          if (para && para.Range && applyFormat(para.Range, FORMAT_RULES.tableCaption)) {
            applied++;
            details.tables++;
          }
        } catch (e) {}
      }
      console.log('[designformat] 表名: ' + details.tables + '个');
    }

    // 2.5 图片居中
    try {
      var inlineCount = doc.InlineShapes ? doc.InlineShapes.Count : 0;
      for (var imgIdx = 1; imgIdx <= inlineCount; imgIdx++) {
        try {
          var shape = doc.InlineShapes.Item(imgIdx);
          if (shape && shape.Range && shape.Range.ParagraphFormat) {
            shape.Range.ParagraphFormat.Alignment = 1; // 居中
            // 图片段落用单倍行距避免裁剪
            shape.Range.ParagraphFormat.LineSpacingRule = 0;
            applied++;
            details.figures++;
          }
        } catch (e) {}
      }
      console.log('[designformat] 图片居中: ' + inlineCount + '个');
    } catch (e) {}

    // 2.6 表格处理（等宽 + 跨页重复表头）
    try {
      var tableCount = doc.Tables ? doc.Tables.Count : 0;
      // 获取页面宽度
      var pageWidth = 595.35, leftMargin = 72, rightMargin = 72;
      try {
        var ps = doc.PageSetup;
        if (ps) {
          pageWidth = ps.PageWidth || 595.35;
          leftMargin = ps.LeftMargin || 72;
          rightMargin = ps.RightMargin || 72;
        }
      } catch (e) {}
      var usableWidth = pageWidth - leftMargin - rightMargin;

      for (var tIdx = 1; tIdx <= tableCount; tIdx++) {
        try {
          var table = doc.Tables.Item(tIdx);
          if (!table) continue;

          // 等宽
          try {
            table.PreferredWidthType = 3;
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
      console.log('[designformat] 表格: ' + tableCount + '个');
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
            // cm转磅: 1cm = 72/2.54 磅
            ps.TopMargin = 2.54 * 72 / 2.54;    // 72磅 = 2.54cm
            ps.BottomMargin = 72;
            ps.LeftMargin = 3.17 * 72 / 2.54;   // 约90磅
            ps.RightMargin = 3.17 * 72 / 2.54;
            ps.PageWidth = 595.35;  // A4
            ps.PageHeight = 841.95;
            ps.Orientation = 0;  // 纵向

            applied++;
            details.pages++;
          } catch (e) {}
        }
        console.log('[designformat] 页面设置: ' + details.pages + '节');
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

            // 页眉
            try {
              var header = sec.Headers.Item(1);
              if (header && header.Range && header.Range.Font) {
                header.Range.Font.NameFarEast = '宋体';
                header.Range.Font.Name = 'Arial';
                header.Range.Font.Size = 9; // 小五号
                // 页眉线
                try {
                  var borders = header.Range.ParagraphFormat.Borders;
                  if (borders) {
                    var bottom = borders.Item(-3); // wdBorderBottom
                    bottom.LineStyle = 1;
                    bottom.LineWidth = 0.5;
                  }
                } catch (e) {}
              }
            } catch (e) {}

            // 页脚（页码）
            try {
              var footer = sec.Footers.Item(1);
              if (footer) {
                if (footer.Range && footer.Range.Font) {
                  footer.Range.Font.NameFarEast = '宋体';
                  footer.Range.Font.Name = 'Arial';
                  footer.Range.Font.Size = 9;
                }
                // 添加页码
                try {
                  var pn = footer.PageNumbers;
                  if (pn) {
                    pn.NumberStyle = 0; // 阿拉伯数字
                    pn.Add(1); // 居中
                  }
                } catch (e) {}
              }
            } catch (e) {}

            applied++;
          } catch (e) {}
        }
        console.log('[designformat] 页眉页脚: ' + sections.Count + '节');
      }
    } catch (e) {}

  } finally {
    try { doc.TrackRevisions = origTrack; } catch (e) {}
  }

  var elapsed = Date.now() - startTime;
  console.log('[designformat] 完成: applied=' + applied + ', time=' + elapsed + 'ms');

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