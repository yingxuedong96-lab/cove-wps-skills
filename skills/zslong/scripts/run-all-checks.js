/**
 * run-all-checks.js
 * 单一脚本完成所有检查和修复工作
 * 数据在脚本内部流转，不经过 LLM 中转
 *
 * 入参:
 *   fixMode (string) - 'aggressive' | 'standard' | 'conservative'
 *     - aggressive: 格式+内容+编号全部自动修复
 *     - standard: 格式自动修复，内容类批注，编号自动修复
 *     - conservative: 全部只批注不修改
 *   isSelection (boolean) - 是否只处理选区
 *   scope (string) - 执行范围
 *     - all: 执行所有检查
 *     - numbering: 章节+表+图+公式编号（全部编号）
 *     - heading: 只执行标题编号检查（N-002~008）
 *     - table: 只执行表编号检查（T-001）
 *     - figure: 只执行图编号检查（G-001）
 *     - formula: 只执行公式编号检查（E-001）
 *
 * 出参: { fixed, commented, revisionLog, summary: { byRule, totalIssues }, paraCount, tableCount }
 */

try {
  // ========== 参数接收 ==========
  var mode = typeof fixMode !== 'undefined' ? fixMode : 'standard';
  var isSel = typeof isSelection !== 'undefined' ? isSelection : false;
  var checkScope = typeof scope !== 'undefined' ? scope : 'all';

  console.log('[run-all-checks] 参数: mode=' + mode + ', scope=' + checkScope);

  var doc = Application.ActiveDocument;
  if (!doc) {
    console.log('[run-all-checks] 无活动文档');
    return { error: '无活动文档', fixed: 0, commented: 0, revisionLog: [] };
  }

  // 保存原始修订模式设置
  var originalTrackRevisions = doc.TrackRevisions;

  // ========== 编号类scope走两阶段处理 ==========
  var isPureNumberingScope = checkScope === 'numbering' || checkScope === 'heading' ||
                              checkScope === 'table' || checkScope === 'figure' || checkScope === 'formula';
  var isDirectTableContentScope = checkScope === 'table_content';
  var isDirectValueScope = checkScope === 'value';
  var isDirectContentScope = checkScope === 'content';
  var isDirectPunctuationScope = checkScope === 'punctuation';
  var isDirectFontScope = checkScope === 'font';
  var isDirectFigureTableScope = checkScope === 'figure_table_layout' || checkScope === 'figure_layout' ||
                                 checkScope === 'table_layout' || checkScope === 'figure_caption' ||
                                 checkScope === 'table_caption' || checkScope === 'figure_center';
  var isDirectFormulaLayoutScope = checkScope === 'formula_layout';
  var isDirectHeaderFooterScope = checkScope === 'header_footer';
  var isDirectPageSetupScope = checkScope === 'page_setup';
  var isDirectFormatScope = checkScope === 'format';
  var isDirectFullProofreadScope = checkScope === 'full_proofread';

  if (isDirectFullProofreadScope) {
    console.log('[run-all-checks] 全文校对排版scope，走完整流水线快路径');
    var fullProofreadResult = processFullProofreadFast(doc, mode);
    return {
      fixed: fullProofreadResult.fixed,
      commented: fullProofreadResult.commented,
      revisionLog: fullProofreadResult.revisionLog,
      summary: {
        totalIssues: fullProofreadResult.fixed + fullProofreadResult.commented,
        byRule: fullProofreadResult.byRule,
        scope: checkScope
      },
      paraCount: fullProofreadResult.paraCount || 0,
      tableCount: doc.Tables ? doc.Tables.Count : 0
    };
  }
  
  if (isPureNumberingScope) {
    console.log('[run-all-checks] 编号类scope，走两阶段处理');
    
    // 直接在 scanStructureForNumbering 中完成扫描+修复
    var scanResult = scanStructureForNumbering(doc, checkScope);
    if (!scanResult.success) {
      return { error: scanResult.error, fixed: 0, commented: 0, revisionLog: [] };
    }
    
    return {
      fixed: scanResult.totalFixed,
      commented: 0,
      revisionLog: scanResult.details,
      summary: {
        totalIssues: scanResult.totalFixed,
        byRule: {},
        scope: checkScope
      },
      paraCount: 0,
      tableCount: 0
    };
  }

  if (isDirectTableContentScope) {
    console.log('[run-all-checks] 表格内容scope，走表格快路径');
    var tableContentResult = processTableContentFast(doc, mode);
    return {
      fixed: tableContentResult.fixed,
      commented: tableContentResult.commented,
      revisionLog: tableContentResult.revisionLog,
      summary: {
        totalIssues: tableContentResult.fixed + tableContentResult.commented,
        byRule: tableContentResult.byRule,
        scope: checkScope
      },
      paraCount: 0,
      tableCount: doc.Tables ? doc.Tables.Count : 0
    };
  }

  if (isDirectValueScope) {
    console.log('[run-all-checks] 数值scope，走数值快路径');
    var valueResult = processValueFast(doc, mode);
    return {
      fixed: valueResult.fixed,
      commented: valueResult.commented,
      revisionLog: valueResult.revisionLog,
      summary: {
        totalIssues: valueResult.fixed + valueResult.commented,
        byRule: valueResult.byRule,
        scope: checkScope
      },
      paraCount: valueResult.paraCount,
      tableCount: 0
    };
  }

  if (isDirectPunctuationScope) {
    console.log('[run-all-checks] 标点scope，走标点快路径');
    var punctuationResult = processPunctuationFast(doc, mode);
    return {
      fixed: punctuationResult.fixed,
      commented: punctuationResult.commented,
      revisionLog: punctuationResult.revisionLog,
      summary: {
        totalIssues: punctuationResult.fixed + punctuationResult.commented,
        byRule: punctuationResult.byRule,
        scope: checkScope
      },
      paraCount: punctuationResult.paraCount,
      tableCount: 0
    };
  }

  if (isDirectContentScope) {
    console.log('[run-all-checks] 内容scope，走内容快路径');
    var contentResult = processContentFast(doc, mode);
    return {
      fixed: contentResult.fixed,
      commented: contentResult.commented,
      revisionLog: contentResult.revisionLog,
      summary: {
        totalIssues: contentResult.fixed + contentResult.commented,
        byRule: contentResult.byRule,
        scope: checkScope
      },
      paraCount: contentResult.paraCount || 0,
      tableCount: doc.Tables ? doc.Tables.Count : 0
    };
  }

  if (isDirectFontScope) {
    console.log('[run-all-checks] 标题正文字体scope，走排版快路径');
    var fontResult = processFontFast(doc, mode);
    return {
      fixed: fontResult.fixed,
      commented: fontResult.commented,
      revisionLog: fontResult.revisionLog,
      summary: {
        totalIssues: fontResult.fixed + fontResult.commented,
        byRule: fontResult.byRule,
        scope: checkScope
      },
      paraCount: fontResult.paraCount,
      tableCount: 0
    };
  }

  if (isDirectFigureTableScope) {
    console.log('[run-all-checks] 图表排版scope，走图表快路径');
    var figureTableResult = processFigureTableLayoutFast(doc, mode, checkScope);
    return {
      fixed: figureTableResult.fixed,
      commented: figureTableResult.commented,
      revisionLog: figureTableResult.revisionLog,
      summary: {
        totalIssues: figureTableResult.fixed + figureTableResult.commented,
        byRule: figureTableResult.byRule,
        scope: checkScope
      },
      paraCount: figureTableResult.paraCount,
      tableCount: doc.Tables ? doc.Tables.Count : 0
    };
  }

  if (isDirectFormulaLayoutScope) {
    console.log('[run-all-checks] 公式排版scope，走公式快路径');
    var formulaLayoutResult = processFormulaLayoutFast(doc, mode);
    return {
      fixed: formulaLayoutResult.fixed,
      commented: formulaLayoutResult.commented,
      revisionLog: formulaLayoutResult.revisionLog,
      summary: {
        totalIssues: formulaLayoutResult.fixed + formulaLayoutResult.commented,
        byRule: formulaLayoutResult.byRule,
        scope: checkScope
      },
      paraCount: formulaLayoutResult.paraCount,
      tableCount: 0
    };
  }

  if (isDirectHeaderFooterScope) {
    console.log('[run-all-checks] 页眉页脚scope，走页眉页脚快路径');
    var headerFooterResult = processHeaderFooterFast(doc, mode);
    return {
      fixed: headerFooterResult.fixed,
      commented: headerFooterResult.commented,
      revisionLog: headerFooterResult.revisionLog,
      summary: {
        totalIssues: headerFooterResult.fixed + headerFooterResult.commented,
        byRule: headerFooterResult.byRule,
        scope: checkScope
      },
      paraCount: 0,
      tableCount: 0
    };
  }

  if (isDirectPageSetupScope) {
    console.log('[run-all-checks] 页面设置scope，走页面设置快路径');
    var pageSetupResult = processPageSetupFast(doc, mode);
    return {
      fixed: pageSetupResult.fixed,
      commented: pageSetupResult.commented,
      revisionLog: pageSetupResult.revisionLog,
      summary: {
        totalIssues: pageSetupResult.fixed + pageSetupResult.commented,
        byRule: pageSetupResult.byRule,
        scope: checkScope
      },
      paraCount: 0,
      tableCount: 0
    };
  }

  if (isDirectFormatScope) {
    console.log('[run-all-checks] 格式scope，走格式快路径');
    var formatResult = processFormatFast(doc, mode);
    return {
      fixed: formatResult.fixed,
      commented: formatResult.commented,
      revisionLog: formatResult.revisionLog,
      summary: {
        totalIssues: formatResult.fixed + formatResult.commented,
        byRule: formatResult.byRule,
        scope: checkScope
      },
      paraCount: formatResult.paraCount || 0,
      tableCount: doc.Tables ? doc.Tables.Count : 0
    };
  }

  // ========== Step 1: 读取文档结构 ==========
  console.log('[run-all-checks] 开始读取文档结构...');
  var structure = readStructure(doc, isSel);
  console.log('[run-all-checks] 读取完成，段落数: ' + structure.paragraphs.length);

  // 根据 scope 决定执行哪些检查
  var isFullProofreadScope = checkScope === 'full_proofread'; // 全文校对排版（编号→内容→格式）
  var isNumberingScope = checkScope === 'numbering';  // 全部编号
  var isHeadingScope = checkScope === 'heading';       // 仅标题编号
  var isTableScope = checkScope === 'table';           // 仅表编号
  var isFigureScope = checkScope === 'figure';         // 仅图编号
  var isFormulaScope = checkScope === 'formula';       // 仅公式编号

  // Stage 2 内容规范的细分 scope
  var isValueScope = checkScope === 'value';           // 仅数值格式（V系列）
  var isPunctuationScope = checkScope === 'punctuation'; // 仅标点（M系列）
  var isTableContentScope = checkScope === 'table_content'; // 仅表格内容（T-005, T-006）
  var isContentScope = checkScope === 'content';       // 全部内容规范

  // Stage 4 格式规范排版
  var isFormatScope = checkScope === 'format';         // 全部排版
  var isFontScope = checkScope === 'font';             // 仅字体字号（F-001~F-005）
  var isFigureCaptionScope = checkScope === 'figure_caption'; // 图名排版（G-002）
  var isFigureCenterScope = checkScope === 'figure_center';   // 图片居中（G-004）
  var isTableCaptionScope = checkScope === 'table_caption';   // 表名排版（T-002）
  var isTableFormatScope = checkScope === 'table_format';     // 表格排版（T-004, T-007）
  var isFigureLayoutScope = checkScope === 'figure_layout';   // 图片排版（G-002, G-004）
  var isTableLayoutScope = checkScope === 'table_layout';     // 表格排版（T-002, T-004, T-007）
  var isFigureTableLayoutScope = checkScope === 'figure_table_layout'; // 图表排版
  var isFormulaLayoutScope = checkScope === 'formula_layout'; // 公式排版（E-002, E-003）
  var isHeaderFooterScope = checkScope === 'header_footer';   // 页眉页脚排版（HF-001~003）
  var isPageSetupScope = checkScope === 'page_setup';         // 页面设置（PG-001~002）

  // 是否需要执行标题编号检查
  var needHeadingCheck = checkScope === 'all' || isFullProofreadScope || isNumberingScope || isHeadingScope;
  // 是否需要执行表编号检查
  var needTableCheck = checkScope === 'all' || isFullProofreadScope || isNumberingScope || isTableScope;
  // 是否需要执行图编号检查
  var needFigureCheck = checkScope === 'all' || isFullProofreadScope || isNumberingScope || isFigureScope;
  // 是否需要执行公式编号检查
  var needFormulaCheck = checkScope === 'all' || isFullProofreadScope || isNumberingScope || isFormulaScope;

  // 是否需要执行内容检查（Stage 2）
  var needContentCheck = checkScope === 'all' || isFullProofreadScope || isContentScope || isValueScope || isPunctuationScope || isTableContentScope;

  // 是否需要执行格式检查（Stage 4）
  // all 仅表示“校对全文”，不包含排版；full_proofread 才包含完整排版流程
  var needFormatCheck = isFullProofreadScope || isFormatScope || isFontScope || isFigureCaptionScope || isFigureCenterScope || isTableCaptionScope || isTableFormatScope || isFigureLayoutScope || isTableLayoutScope || isFigureTableLayoutScope || isFormulaLayoutScope || isHeaderFooterScope || isPageSetupScope;

  // 记录总修复数
  var totalFixed = 0;
  var totalCommented = 0;
  var allRevisionLog = [];
  var allByRule = {};

  // ========== 分阶段处理 ==========
  // Stage 1A：标题编号
  // Stage 1B：图/表/公式编号（基于修复后的标题编号）
  // Stage 2：内容规范
  // Stage 4：格式规范排版

  if (needHeadingCheck) {
    // Stage 1A：章节编号检查
    var structureIssues = checkStructure(structure);
    console.log('[run-all-checks] 标题编号检查发现 ' + structureIssues.length + ' 个问题');

    // 只修复章节编号问题
    var chapterIssues = structureIssues.filter(function(issue) {
      return issue.rule.indexOf('N-') === 0;
    });

    if (chapterIssues.length > 0) {
      // 开启修订模式
      doc.TrackRevisions = true;

      var chapterResult = applyFixesAndComments(doc, chapterIssues, mode);
      console.log('[run-all-checks] 标题编号修复：' + chapterResult.fixed + ' 处');

      // 恢复修订模式
      doc.TrackRevisions = originalTrackRevisions;

      // 记录结果
      totalFixed += chapterResult.fixed;
      totalCommented += chapterResult.commented;
      allRevisionLog = allRevisionLog.concat(chapterResult.revisionLog);
      for (var key in chapterResult.byRule) {
        allByRule[key] = (allByRule[key] || 0) + chapterResult.byRule[key];
      }
    }
  }

  // Stage 1B 依赖最新章节编号，进入前重新读取一次结构
  if (needTableCheck || needFigureCheck || needFormulaCheck) {
    structure = readStructure(doc, isSel);
    console.log('[run-all-checks] Stage 1B：重新读取文档结构');
  }

  // Stage 1B 问题列表（图/表/公式编号）
  var stage1bIssues = [];

  if (needTableCheck) {
    // 表编号检查
    var tableIssues = checkTableNumbering(structure);
    console.log('[run-all-checks] 表编号检查发现 ' + tableIssues.length + ' 个问题');
    stage1bIssues = stage1bIssues.concat(tableIssues);
  }

  if (needFigureCheck) {
    // 图编号检查
    var figureIssues = checkFigureNumbering(structure);
    console.log('[run-all-checks] 图编号检查发现 ' + figureIssues.length + ' 个问题');
    stage1bIssues = stage1bIssues.concat(figureIssues);
  }

  if (needFormulaCheck) {
    // 公式编号检查
    var formulaIssues = checkFormulaNumbering(structure);
    console.log('[run-all-checks] 公式编号检查发现 ' + formulaIssues.length + ' 个问题');
    stage1bIssues = stage1bIssues.concat(formulaIssues);
  }

  if (stage1bIssues.length > 0) {
    console.log('[run-all-checks] Stage 1B 总共发现 ' + stage1bIssues.length + ' 个问题');

    doc.TrackRevisions = true;
    var stage1bResult = applyFixesAndComments(doc, stage1bIssues, mode);
    doc.TrackRevisions = originalTrackRevisions;

    totalFixed += stage1bResult.fixed;
    totalCommented += stage1bResult.commented;
    allRevisionLog = allRevisionLog.concat(stage1bResult.revisionLog);
    for (var key in stage1bResult.byRule) {
      allByRule[key] = (allByRule[key] || 0) + stage1bResult.byRule[key];
    }
  }

  if (needContentCheck) {
    // Stage 2 需要基于最新编号和最新文本重新分析
    structure = readStructure(doc, isSel);
    console.log('[run-all-checks] Stage 2：重新读取文档结构');

    var contentScope = checkScope;
    var contentIssues = checkContent(structure, contentScope);
    console.log('[run-all-checks] 内容检查发现 ' + contentIssues.length + ' 个问题');

    if (contentIssues.length > 0) {
      doc.TrackRevisions = true;
      var contentResult = applyFixesAndComments(doc, contentIssues, mode);
      doc.TrackRevisions = originalTrackRevisions;

      totalFixed += contentResult.fixed;
      totalCommented += contentResult.commented;
      allRevisionLog = allRevisionLog.concat(contentResult.revisionLog);
      for (var key in contentResult.byRule) {
        allByRule[key] = (allByRule[key] || 0) + contentResult.byRule[key];
      }
    }
  }

  if (needFormatCheck) {
    // Stage 4 依赖最终文本和编号，进入前再次读取结构
    structure = readStructure(doc, isSel);
    console.log('[run-all-checks] Stage 4：重新读取文档结构');

    console.log('[run-all-checks] 执行格式检查，scope=' + checkScope);
    var formatIssues = checkFormat(structure, checkScope);
    console.log('[run-all-checks] 格式检查发现 ' + formatIssues.length + ' 个问题');

    if (formatIssues.length > 0) {
      doc.TrackRevisions = true;
      var formatResult = applyFixesAndComments(doc, formatIssues, mode);
      doc.TrackRevisions = originalTrackRevisions;

      totalFixed += formatResult.fixed;
      totalCommented += formatResult.commented;
      allRevisionLog = allRevisionLog.concat(formatResult.revisionLog);
      for (var key in formatResult.byRule) {
        allByRule[key] = (allByRule[key] || 0) + formatResult.byRule[key];
      }
    }
  }

  console.log('[run-all-checks] 总计修复：' + totalFixed + ' 处，批注：' + totalCommented + ' 处');

  return {
    fixed: totalFixed,
    commented: totalCommented,
    revisionLog: allRevisionLog,
    summary: {
      totalIssues: totalFixed + totalCommented,
      byRule: allByRule,
      scope: checkScope
    },
    paraCount: structure.paragraphs.length,
    tableCount: structure.tables.length
  };

} catch (e) {
  console.warn('[run-all-checks]', e);
  return { error: String(e), fixed: 0, commented: 0, revisionLog: [] };
}

// ========== 内部函数：读取文档结构 ==========

function readStructure(doc, isSelection) {
  var result = { paragraphs: [], tables: [], images: [], charCount: 0, isSelection: isSelection };

  // 确定处理范围
  var targetRange;
  if (isSelection === true) {
    // 只处理选区
    try {
      var sel = Application.Selection;
      if (sel && sel.Range && sel.Range.Start !== sel.Range.End) {
        targetRange = sel.Range;
      } else {
        // 无有效选区，返回空结果
        return result;
      }
    } catch (e) {
      return result;
    }
  } else {
    targetRange = doc.Content;
  }

  // 获取目标范围内的段落
  // 注意：选区模式下需要通过 Range 的 Paragraphs 获取
  var paraCount = targetRange.Paragraphs.Count;
  // 长文档处理：移除段落限制，允许处理所有段落
  // 编号校对需要全文上下文，不能截断
  var limit = paraCount;

  console.log('[readStructure] 开始读取 ' + paraCount + ' 个段落');

  for (var i = 1; i <= limit; i++) {
    try {
      var para = targetRange.Paragraphs.Item(i);
      if (!para || !para.Range) continue;

      var range = para.Range;
      // 保存原始文本（包含开头空白），用于检测开头空格/Tab
      var rawText = range.Text ? String(range.Text).replace(/\r/g, '').replace(/\n/g, '') : '';
      // trim 后的文本用于标题识别等
      var text = rawText.trim();

      // 获取 Word 自动编号
      var listString = '';
      try {
        if (range.ListFormat && range.ListFormat.ListString) {
          listString = String(range.ListFormat.ListString).trim();
          // 如果自动编号存在且文本不以编号开头，则将编号加到文本前面
          if (listString && !new RegExp('^' + listString.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).test(text)) {
            text = listString + ' ' + text;
          }
        }
      } catch (e) {}

      if (text.length === 0) continue;

      // 调试：检测表/图开头的段落
      if (/^表\s*\d/.test(text) || /^图\s*\d/.test(text)) {
        console.log('[readStructure] 检测到表/图段落: 段落' + i + ' = "' + text.substring(0, 50) + '"');
      }

      // 检测段落是否包含图片（嵌入式或浮动式）
      var hasImage = false;
      try {
        // 检测嵌入式图片
        if (range.InlineShapes && range.InlineShapes.Count > 0) {
          hasImage = true;
        }
        // 检测浮动图片（通过 ShapeRange）
        if (!hasImage && range.ShapeRange && range.ShapeRange.Count > 0) {
          hasImage = true;
        }
      } catch (e) {}

      var info = {
        index: i,
        text: text.length > 100 ? text.substring(0, 100) + '...' : text,
        fullText: text,  // 保存完整文本，用于检查
        rawText: rawText,  // 保存原始文本（含开头空白），用于检测开头空格
        styleName: para.Style ? para.Style.NameLocal : '',
        fontCN: range.Font && range.Font.NameFarEast ? range.Font.NameFarEast : '',
        fontEN: range.Font && range.Font.Name ? range.Font.Name : '',
        fontSize: range.Font && range.Font.Size ? range.Font.Size : 0,
        bold: range.Font && range.Font.Bold === -1,
        alignment: para.Format && para.Format.Alignment !== undefined ? para.Format.Alignment : 0,
        lineSpacing: para.Format && para.Format.LineSpacing ? para.Format.LineSpacing : 0,
        firstLineIndent: para.Format && para.Format.FirstLineIndent ? para.Format.FirstLineIndent : 0,
        spaceBefore: para.Format && para.Format.SpaceBefore ? para.Format.SpaceBefore : 0,
        spaceAfter: para.Format && para.Format.SpaceAfter ? para.Format.SpaceAfter : 0,
        listString: listString,  // 保存自动编号
        hasImage: hasImage  // 标记是否包含图片
      };

      result.paragraphs.push(info);
      result.charCount += text.length;
    } catch (e) {}
  }

  // 表格处理
  try {
    if (isSelection === true) {
      // 选区模式：只统计选区内的表格
      var tableCount = doc.Tables.Count;
      for (var j = 1; j <= tableCount; j++) {
        try {
          var table = doc.Tables.Item(j);
          if (!table || !table.Range) continue;
          // 检查表格是否与选区有交集
          if (targetRange.InRange(table.Range) || table.Range.InRange(targetRange)) {
            result.tables.push({
              index: j,
              rows: table.Rows ? table.Rows.Count : 0,
              cols: table.Columns ? table.Columns.Count : 0
            });
          }
        } catch (e) {}
      }
    } else {
      var tableCount = doc.Tables.Count;
      for (var j = 1; j <= tableCount; j++) {
        try {
          var table = doc.Tables.Item(j);
          if (!table) continue;
          result.tables.push({
            index: j,
            rows: table.Rows ? table.Rows.Count : 0,
            cols: table.Columns ? table.Columns.Count : 0
          });
        } catch (e) {}
      }
    }
  } catch (e) {}

  // 图片处理
  try {
    if (isSelection === true) {
      // 选区模式：只统计选区内的图片
      var imageCount = doc.InlineShapes.Count;
      for (var k = 1; k <= imageCount; k++) {
        try {
          var shape = doc.InlineShapes.Item(k);
          if (shape && shape.Range && targetRange.InRange(shape.Range)) {
            result.images.push({ index: k });
          }
        } catch (e) {}
      }
    } else {
      var imageCount = doc.InlineShapes.Count;
      for (var k = 1; k <= imageCount; k++) {
        result.images.push({ index: k });
      }
    }
  } catch (e) {}

  return result;
}

// ========== 页眉页脚辅助函数 ==========

function getHeaderFooterBorderParagraph(hf, preferLast) {
  if (!hf || !hf.Range || !hf.Range.Paragraphs || hf.Range.Paragraphs.Count <= 0) return null;

  var paragraphs = hf.Range.Paragraphs;
  var count = paragraphs.Count;
  var start = preferLast ? count : 1;
  var end = preferLast ? 1 : count;
  var step = preferLast ? -1 : 1;

  for (var i = start; preferLast ? i >= end : i <= end; i += step) {
    try {
      var para = paragraphs.Item(i);
      var text = para && para.Range && para.Range.Text ? String(para.Range.Text).replace(/[\r\u0007]/g, '').trim() : '';
      if (text) return para;
    } catch (e) {}
  }

  try {
    return preferLast ? paragraphs.Item(count) : paragraphs.Item(1);
  } catch (e) {
    return null;
  }
}

function getHeaderFooterEffectiveParagraphs(hf) {
  var result = [];
  if (!hf || !hf.Range || !hf.Range.Paragraphs || hf.Range.Paragraphs.Count <= 0) return result;

  var paragraphs = hf.Range.Paragraphs;
  for (var i = 1; i <= paragraphs.Count; i++) {
    try {
      var para = paragraphs.Item(i);
      var text = para && para.Range && para.Range.Text ? String(para.Range.Text).replace(/[\r\u0007]/g, '').trim() : '';
      if (text) result.push(para);
    } catch (e) {}
  }

  if (result.length === 0) {
    try {
      result.push(paragraphs.Item(1));
    } catch (e) {}
  }

  return result;
}

function setBorderVisibleSafe(borders, borderIndex, visible, lineStyle, lineWidth) {
  try {
    var border = borders ? borders.Item(borderIndex) : null;
    if (!border) return;
    border.Visible = visible;
    if (visible) {
      if (lineStyle !== undefined) border.LineStyle = lineStyle;
      if (lineWidth !== undefined) border.LineWidth = lineWidth;
    }
  } catch (e) {}
}

function clearHeaderFooterParagraphBorders(paragraphs) {
  for (var i = 0; i < paragraphs.length; i++) {
    try {
      var para = paragraphs[i];
      var borders = para && para.Range ? para.Range.ParagraphFormat.Borders : null;
      if (!borders) continue;
      // WdBorderType: top=-1, left=-2, bottom=-3, right=-4
      setBorderVisibleSafe(borders, -1, false);
      setBorderVisibleSafe(borders, -2, false);
      setBorderVisibleSafe(borders, -3, false);
      setBorderVisibleSafe(borders, -4, false);
    } catch (e) {}
  }
}

function applyHeaderFooterLine(paragraph, borderIndex, lineStyle, lineWidth) {
  if (!paragraph || !paragraph.Range) return false;
  try {
    var para = paragraph;
    var borders = para.Range.ParagraphFormat.Borders;
    if (!borders) return false;

    // 清掉该段其他常见边框，避免出现拐角或残留竖线
    setBorderVisibleSafe(borders, -1, false);
    setBorderVisibleSafe(borders, -2, false);
    setBorderVisibleSafe(borders, -3, false);
    setBorderVisibleSafe(borders, -4, false);

    setBorderVisibleSafe(borders, borderIndex, true, lineStyle, lineWidth);
    return true;
  } catch (e) {
    return false;
  }
}

// ========== 内部函数：格式检查 ==========

function checkFormat(structure, formatScope) {
  var issues = [];
  var scope = formatScope || 'format';  // 默认执行全部格式检查

  // 是否只检查字体字号
  var fontOnly = scope === 'font';
  // 是否只检查图名格式（G-002）
  var figureCaptionOnly = scope === 'figure_caption';
  // 是否只检查图片居中（G-004）
  var figureCenterOnly = scope === 'figure_center';
  // 是否只检查表名格式（T-002）
  var tableCaptionOnly = scope === 'table_caption';
  // 是否只检查表格格式（T-004, T-007）
  var tableFormatOnly = scope === 'table_format';
  // 图片排版：G-002 + G-004
  var figureLayoutOnly = scope === 'figure_layout';
  // 表格排版：T-002 + T-004 + T-007
  var tableLayoutOnly = scope === 'table_layout';
  // 图表排版：G-002 + G-004 + T-002 + T-004 + T-007
  var figureTableLayoutOnly = scope === 'figure_table_layout';
  // 公式排版：E-002 + E-003
  var formulaLayoutOnly = scope === 'formula_layout';
  // 页眉页脚排版：HF-001~003
  var headerFooterOnly = scope === 'header_footer';
  // 页面设置：PG-001~002
  var pageSetupOnly = scope === 'page_setup';

  // 统一的格式规则定义（合并原 F 系列和 P 系列）
  var FORMAT_RULES = {
    'F-001': {
      name: '一级标题',
      expectedFont: '宋体',
      expectedSize: 12,
      expectedBold: true,
      expectedLineSpacing: 20,
      expectedSpaceBefore: 6,
      expectedSpaceAfter: 6,
      expectedFirstIndent: 0
    },
    'F-002': {
      name: '二级标题',
      expectedFont: '宋体',
      expectedSize: 12,
      expectedBold: true,
      expectedLineSpacing: 20,
      expectedSpaceBefore: 6,
      expectedSpaceAfter: 6,
      expectedFirstIndent: 0
    },
    'F-003': {
      name: '三级标题',
      expectedFont: '宋体',
      expectedSize: 12,
      expectedBold: false,
      expectedLineSpacing: 20,
      expectedSpaceBefore: 6,
      expectedSpaceAfter: 6,
      expectedFirstIndent: 0
    },
    'F-004': {
      name: '四级标题',
      expectedFont: '宋体',
      expectedSize: 12,
      expectedBold: false,
      expectedLineSpacing: 20,
      expectedSpaceBefore: 6,
      expectedSpaceAfter: 6,
      expectedFirstIndent: 0
    },
    'F-005': {
      name: '正文',
      expectedFont: '宋体',
      expectedSize: 12,
      expectedBold: false,
      expectedLineSpacing: 20,
      expectedSpaceBefore: 6,
      expectedSpaceAfter: 6,
      expectedFirstIndent: 24  // 首行缩进2字符
    },
    // 图名格式规则
    'G-002': {
      name: '图名',
      expectedFont: '宋体',
      expectedSize: 12,
      expectedBold: false,
      expectedLineSpacing: 20,
      expectedSpaceBefore: 6,
      expectedSpaceAfter: 6,
      expectedAlignment: 1  // 居中对齐
    },
    // 图片格式规则
    'G-004': {
      name: '图片',
      expectedAlignment: 1,  // 居中对齐
      expectedLineSpacing: 20,
      expectedSpaceBefore: 6,
      expectedSpaceAfter: 6
    }
  };

  // 判断是否是图名段落（如 "图 1-1 大坝剖面图"、"图B1 调压室结构图"）
  function isFigureCaption(text) {
    // 匹配：图+数字（如 图1-1）、图+字母+数字（如图B1）、图+空格+编号
    return /^图\s*[\dA-Za-z]/.test(text);
  }

  function isTableCaption(text) {
    return /^表\s*[\dA-Za-z]/.test(text);
  }

  function isFormulaParagraph(text) {
    return /\([A-Z]?\d+(?:\.\d+)*(?:-\d+)?\)\s*$/.test(text) || /\t\([A-Z]?\d+(?:\.\d+)*(?:-\d+)?\)\s*$/.test(text);
  }

  function getHeadingLevel(styleName, text) {
    // 附录标题（如 "附录 A"）
    if (/^附录\s*[A-Z一二三四五六七八九十]/i.test(text)) return 10;

    // 附录章节编号（如 A1、B1.1、C1.1.1）- 根据规范附录应顶格书写
    if (/^[A-Z]\d+\s/.test(text)) return 1;        // A1 标题
    if (/^[A-Z]\d+\.\d+\s/.test(text)) return 2;   // A1.1 标题
    if (/^[A-Z]\d+\.\d+\.\d+\s/.test(text)) return 3; // A1.1.1 标题

    // 优先检查文本格式
    // 中文章节编号（第一章、第二章...）
    if (/^第[一二三四五六七八九十百]+章/.test(text)) return 1;
    // 阿拉伯数字编号 - 使用负向前瞻避免匹配二级标题 "1.1"
    if (/^(\d+)\s+(?!\.)/.test(text)) return 1;
    if (/^\d+\.\d+\s/.test(text)) return 2;
    if (/^\d+\.\d+\.\d+\s/.test(text)) return 3;
    if (/^\d+\.\d+\.\d+\.\d+\s/.test(text)) return 4;

    // 再检查样式名称
    if (!styleName) return 0;
    var upperName = styleName.toUpperCase();

    if (upperName.indexOf('HEADING 1') >= 0 || upperName.indexOf('标题 1') >= 0 || upperName === '1') return 1;
    if (upperName.indexOf('HEADING 2') >= 0 || upperName.indexOf('标题 2') >= 0 || upperName === '2') return 2;
    if (upperName.indexOf('HEADING 3') >= 0 || upperName.indexOf('标题 3') >= 0 || upperName === '3') return 3;
    if (upperName.indexOf('HEADING 4') >= 0 || upperName.indexOf('标题 4') >= 0 || upperName === '4') return 4;

    return 0;
  }

  // 统一的段落格式检查（合并原 checkParaFormat 和 checkLineSpacing）
  function checkParagraphFormat(paragraph) {
    var text = paragraph.text || '';
    var fontCN = paragraph.fontCN || '';
    var fontSize = paragraph.fontSize || 0;
    var styleName = paragraph.styleName || '';
    var bold = paragraph.bold || false;
    var lineSpacing = paragraph.lineSpacing || 0;
    var spaceBefore = paragraph.spaceBefore || 0;
    var spaceAfter = paragraph.spaceAfter || 0;
    var firstLineIndent = paragraph.firstLineIndent || 0;
    var alignment = paragraph.alignment || 0;

    // 跳过包含图片的段落，避免影响图片布局
    if (paragraph.hasImage) {
      return;
    }

    var level = getHeadingLevel(styleName, text);
    var ruleKey = '';

    // 标题正文排版只处理标题和正文，不处理图名、表名、公式段
    if (fontOnly && (isFigureCaption(text) || isTableCaption(text) || isFormulaParagraph(text))) {
      return;
    }

    // 先检查是否是图名段落
    if (isFigureCaption(text)) {
      ruleKey = 'G-002';
    } else if (level === 1) ruleKey = 'F-001';
    else if (level === 2) ruleKey = 'F-002';
    else if (level === 3) ruleKey = 'F-003';
    else if (level === 4) ruleKey = 'F-004';
    else if (level === 0 && text.length > 0 && !isTableCaption(text) && !isFormulaParagraph(text)) ruleKey = 'F-005';  // 正文：只要有内容就检查

    if (!ruleKey) return;

    var rule = FORMAT_RULES[ruleKey];
    if (!rule) return;

    var problems = [];
    var fixSpec = {
      font: rule.expectedFont,
      size: rule.expectedSize,
      bold: rule.expectedBold,
      lineSpacing: rule.expectedLineSpacing,
      spaceBefore: rule.expectedSpaceBefore,
      spaceAfter: rule.expectedSpaceAfter
    };

    // 检查字体
    if (fontCN && rule.expectedFont && fontCN !== rule.expectedFont) {
      problems.push('字体应为' + rule.expectedFont);
    }

    // 检查字号
    if (fontSize > 0 && rule.expectedSize && Math.abs(fontSize - rule.expectedSize) > 0.5) {
      problems.push('字号应为' + rule.expectedSize + '磅');
    }

    // 检查加粗
    if (rule.expectedBold !== undefined && level >= 1 && level <= 4) {
      if (bold !== rule.expectedBold) {
        problems.push(rule.expectedBold ? '应加粗' : '不应加粗');
      }
    }

    // 检查行距
    if (rule.expectedLineSpacing && lineSpacing > 0 && Math.abs(lineSpacing - rule.expectedLineSpacing) > 2) {
      problems.push('行距应为' + rule.expectedLineSpacing + '磅');
      fixSpec.lineSpacing = rule.expectedLineSpacing;
    }

    // 检查段间距
    if (rule.expectedSpaceBefore !== undefined && Math.abs(spaceBefore - rule.expectedSpaceBefore) > 1) {
      problems.push('段前间距应为' + rule.expectedSpaceBefore + '磅');
    }
    if (rule.expectedSpaceAfter !== undefined && Math.abs(spaceAfter - rule.expectedSpaceAfter) > 1) {
      problems.push('段后间距应为' + rule.expectedSpaceAfter + '磅');
    }

    // 检查首行缩进（仅正文，图名不检查）
    if (rule.expectedFirstIndent !== undefined && ruleKey !== 'G-002') {
      // 使用 rawText 检测开头是否有手动空白字符
      var rawText = paragraph.rawText || text;

      // 检测各种空白字符：空格、全角空格、Tab、以及其他Unicode空白
      var leadingSpaces = rawText.match(/^([ \t　\u00A0\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200A\u202F\u205F\u3000]+)/);
      if (leadingSpaces) {
        var spaceCount = leadingSpaces[1].length;
        var spaceStr = leadingSpaces[1];
        var spaceType = '空白字符';
        if (spaceStr.indexOf('\t') >= 0) spaceType = 'Tab缩进';
        else if (spaceStr.indexOf('　') >= 0 || spaceStr.indexOf('\u3000') >= 0) spaceType = '全角空格';
        else if (spaceStr.indexOf(' ') >= 0) spaceType = '空格';
        problems.push('开头有' + spaceCount + '个' + spaceType + '应删除');
        fixSpec.removeLeadingSpaces = true;
        fixSpec.firstIndent = rule.expectedFirstIndent;
      } else {
        // 确保 firstLineIndent 有默认值，避免 NaN 问题
        var actualIndent = firstLineIndent || 0;
        if (Math.abs(actualIndent - rule.expectedFirstIndent) > 0.5) {
          problems.push(rule.expectedFirstIndent > 0 ? '首行缩进应为2字符' : '首行缩进应为0');
          fixSpec.firstIndent = rule.expectedFirstIndent;
        }
      }
    }

    // 检查图名居中对齐
    if (ruleKey === 'G-002' && rule.expectedAlignment !== undefined) {
      // alignment: 0=左对齐, 1=居中, 2=右对齐
      if (alignment !== rule.expectedAlignment) {
        problems.push('图名应居中对齐');
        fixSpec.alignment = rule.expectedAlignment;
      }
    }

    if (problems.length > 0) {
      issues.push({
        index: paragraph.index,
        rule: ruleKey,
        name: rule.name,
        original: text.substring(0, 50),
        message: problems.join('，'),
        autoFix: true,
        fixSpec: fixSpec
      });
    }
  }

  var paragraphs = structure.paragraphs || [];

  // F-001~F-005: 检查标题与正文格式
  // 在 font 或 format scope 时检查
  if (fontOnly || scope === 'format' || scope === 'full_proofread' || scope === 'all') {
    for (var i = 0; i < paragraphs.length; i++) {
      checkParagraphFormat(paragraphs[i]);
    }
  }

  // G-002: 检查图名格式
  // 在 figure_caption, figure_layout, figure_table_layout, format scope 时检查
  if (figureCaptionOnly || figureLayoutOnly || figureTableLayoutOnly || scope === 'format' || scope === 'full_proofread' || scope === 'all') {
    for (var i = 0; i < paragraphs.length; i++) {
      // 如果只检查图名，跳过非图名段落
      if ((figureCaptionOnly || figureLayoutOnly || figureTableLayoutOnly) && !isFigureCaption(paragraphs[i].text || '')) {
        continue;
      }
      checkParagraphFormat(paragraphs[i]);
    }
  }

  // G-004: 检查图片居中
  // 在 figure_center, figure_layout, figure_table_layout, format scope 时检查
  if (figureCenterOnly || figureLayoutOnly || figureTableLayoutOnly || scope === 'format' || scope === 'full_proofread' || scope === 'all') {
    for (var i = 0; i < paragraphs.length; i++) {
      var p = paragraphs[i];
      var text = p.text || '';

      if (p.hasImage) {
        var alignment = p.alignment || 0;
        console.log('[G-004] 检测到图片段落 ' + p.index + ', 对齐方式=' + alignment);
        if (alignment !== 1) {  // 1 = 居中
          issues.push({
            index: p.index,
            rule: 'G-004',
            name: '图片居中',
            original: text.substring(0, 30) || '图片段落',
            message: '图片应居中对齐',
            autoFix: true,
            fixSpec: {
              alignment: 1,
              // 图片段落使用单倍行距，不使用固定行距（避免遮挡图片）
              lineSpacingRule: 0,  // wdLineSpaceSingle = 0 (单倍行距)
              spaceBefore: 6,
              spaceAfter: 6
            }
          });
        }
      }
    }
  }

  // ========== 表格检查（T-002~005） ==========

  // 判断是否是表名段落（如 "表 1-1 参数表"、"表B1 测试数据"）
  // T-002: 表名格式检查
  // 在 table_caption, table_layout, figure_table_layout, format scope 时检查
  if (tableCaptionOnly || tableLayoutOnly || figureTableLayoutOnly || scope === 'format' || scope === 'full_proofread' || scope === 'all') {
    for (var i = 0; i < paragraphs.length; i++) {
      var p = paragraphs[i];
      var text = p.text || '';

      if (isTableCaption(text)) {
        console.log('[T-002] 检测到表名段落 ' + p.index + ': ' + text.substring(0, 30));

        // T-002: 表名字体字号检查
        var font = p.fontName || '';
        var size = p.fontSize || 0;
        var alignment = p.alignment || 0;
        var spaceBefore = p.spaceBefore || 0;
        var spaceAfter = p.spaceAfter || 0;

        var needsFix = false;
        var fixSpec = {};
        var messages = [];

        // 检查字体（应为宋体）
        if (font.indexOf('宋体') < 0 && font.indexOf('SimSun') < 0) {
          needsFix = true;
          fixSpec.fontName = '宋体';
          messages.push('字体应为宋体');
        }

        // 检查字号（应为小四号=12磅）
        if (Math.abs(size - 12) > 0.5) {
          needsFix = true;
          fixSpec.fontSize = 12;
          messages.push('字号应为小四号(12磅)');
        }

        // 检查对齐方式（应为居中）
        if (alignment !== 1) {
          needsFix = true;
          fixSpec.alignment = 1;
          messages.push('应居中对齐');
        }

        // 检查段前间距（应为6磅=0.5行）
        if (Math.abs(spaceBefore - 6) > 0.5) {
          needsFix = true;
          fixSpec.spaceBefore = 6;
          messages.push('段前应为0.5行');
        }

        // 检查段后间距（应为0磅）
        if (Math.abs(spaceAfter - 0) > 0.5) {
          needsFix = true;
          fixSpec.spaceAfter = 0;
          messages.push('段后应为0');
        }

        if (needsFix) {
          issues.push({
            index: p.index,
            rule: 'T-002',
            name: '表名字体字号',
            original: text.substring(0, 50),
            message: messages.join('，'),
            autoFix: true,
            fixSpec: fixSpec
          });
        }
      }
    }
  }

  // T-004, T-007: 表格格式检查
  // 在 table_format, table_layout, figure_table_layout, format scope 时检查
  if (tableFormatOnly || tableLayoutOnly || figureTableLayoutOnly || scope === 'format' || scope === 'full_proofread' || scope === 'all') {
    var tables = doc.Tables;
    if (tables && tables.Count > 0) {
      for (var t = 1; t <= tables.Count; t++) {
        var table = tables.Item(t);

        // T-004: 表格宽度（与页面等宽）
        try {
          var pageWidth = doc.PageSetup.PageWidth;
          var leftMargin = doc.PageSetup.LeftMargin;
          var rightMargin = doc.PageSetup.RightMargin;
          var contentWidth = pageWidth - leftMargin - rightMargin;

          // 转换为磅（WPS API 可能返回不同单位）
          // PageWidth 通常以磅为单位
          var tableWidth = table.PreferredWidth;
          var widthType = table.PreferredWidthType;

          console.log('[T-004] 表格 ' + t + ' 宽度=' + tableWidth + ', 类型=' + widthType + ', 页面内容宽度=' + contentWidth);

          // 检查表格宽度是否与页面等宽
          // WdPreferredWidthType:
          //   1 = wdPreferredWidthAuto
          //   2 = wdPreferredWidthPercent
          //   3 = wdPreferredWidthPoints
          // 只要不是精确的磅值类型，或者宽度与页面不等，都需要修复
          var needsFix = false;
          if (widthType !== 3) {
            needsFix = true;
            console.log('[T-004] 需要修复：宽度类型不是磅值（type=' + widthType + '）');
          } else if (Math.abs(tableWidth - contentWidth) > 10) {
            needsFix = true;
            console.log('[T-004] 需要修复：宽度差异超过10磅（diff=' + Math.abs(tableWidth - contentWidth) + '）');
          }

          if (needsFix) {
            console.log('[T-004] 添加修复问题，issues.length=' + (issues.length + 1));
            issues.push({
              rule: 'T-004',
              name: '表格宽度',
              tableIndex: t,
              original: '表格宽度: ' + Math.round(tableWidth) + '磅',
              suggested: '页面宽度: ' + Math.round(contentWidth) + '磅',
              message: '表格应与页面等宽',
              autoFix: true,
              fixSpec: {
                preferredWidth: contentWidth,
                preferredWidthType: 3  // wdPreferredWidthPoints
              }
            });
          }
        } catch (e) {
          console.log('[T-004] 检查表格宽度出错: ' + e.message);
        }

        // T-007: 跨页重复表头
        try {
          var rows = table.Rows;
          if (rows.Count > 0) {
            var firstRow = rows.Item(1);
            var headingFormat = firstRow.HeadingFormat;

            console.log('[T-007] 表格 ' + t + ' 表头格式=' + headingFormat);

            // HeadingFormat = -1 表示标题行会重复
            if (headingFormat !== -1 && headingFormat !== true) {
              console.log('[T-007] 需要修复：表头未设置重复（headingFormat=' + headingFormat + '）');
              issues.push({
                rule: 'T-007',
                name: '跨页重复表头',
                tableIndex: t,
                original: '表头未设置重复',
                message: '表格首行应设置为跨页重复表头',
                autoFix: true,
                fixSpec: {
                  headingFormat: true
                }
              });
            }
          }
        } catch (e) {
          console.log('[T-007] 检查表头重复出错: ' + e.message);
        }
      }
    }
  }

  // ========== 公式检查（E-002, E-003） ==========

  function getFormulaNumberMatch(text) {
    if (!text) return null;
    return text.match(/\(([A-Z]\d+|\d+(?:\.\d+){0,2}[-—–]\d+)\)\s*$/);
  }

  function getHeaderFooterBorderParagraph(hf, preferLast) {
    if (!hf || !hf.Range || !hf.Range.Paragraphs || hf.Range.Paragraphs.Count <= 0) return null;

    var paragraphs = hf.Range.Paragraphs;
    var count = paragraphs.Count;
    var start = preferLast ? count : 1;
    var end = preferLast ? 1 : count;
    var step = preferLast ? -1 : 1;

    for (var i = start; preferLast ? i >= end : i <= end; i += step) {
      try {
        var para = paragraphs.Item(i);
        var text = para && para.Range && para.Range.Text ? String(para.Range.Text).replace(/[\r\u0007]/g, '').trim() : '';
        if (text) return para;
      } catch (e) {}
    }

    try {
      return preferLast ? paragraphs.Item(count) : paragraphs.Item(1);
    } catch (e) {
      return null;
    }
  }

  function getHeaderFooterEffectiveParagraphs(hf) {
    var result = [];
    if (!hf || !hf.Range || !hf.Range.Paragraphs || hf.Range.Paragraphs.Count <= 0) return result;

    var paragraphs = hf.Range.Paragraphs;
    for (var i = 1; i <= paragraphs.Count; i++) {
      try {
        var para = paragraphs.Item(i);
        var text = para && para.Range && para.Range.Text ? String(para.Range.Text).replace(/[\r\u0007]/g, '').trim() : '';
        if (text) result.push(para);
      } catch (e) {}
    }

    if (result.length === 0) {
      try {
        result.push(paragraphs.Item(1));
      } catch (e) {}
    }

    return result;
  }

  function countChineseChars(text) {
    var m = String(text || '').match(/[\u4e00-\u9fa5]/g);
    return m ? m.length : 0;
  }

  function isFormulaExplanationText(text) {
    if (!text) return false;
    return /^\s*式中\s*[：:]/.test(text);
  }

  function isReferenceStandardText(text) {
    if (!text) return false;
    if (/《.+》/.test(text)) return true;
    if (/\b(?:GB|GB\/T|DL\/T|NB\/T|SL|SDJ|IEC|ISO)\s*\d+(?:[-—–]\d+)+/i.test(text)) return true;
    return false;
  }

  function isUnitOnlyText(text) {
    if (!text) return false;
    var normalized = String(text).replace(/\u0007/g, '').trim();
    return /^(\d+(?:\.\d+)?)?\s*[A-Za-z%℃°/·²³⁴⁵⁶⁷⁸⁹₀₁₂₃₄₅₆₇₈₉]+$/.test(normalized);
  }

  function containsStrongMathFeatures(text) {
    if (!text) return false;
    if (/[=±×÷∑ΣΔ√∫≈≠≤≥]/.test(text)) return true;
    if (/[λρμτωφψβγαεδ]/.test(text)) return true;
    if (/[A-Za-zΑ-Ωα-ω][A-Za-z0-9Α-Ωα-ω₀₁₂₃₄₅₆₇₈₉²³]*\s*=\s*[^=]/.test(text)) return true;
    if (/[A-Za-zΑ-Ωα-ω]\s*\/\s*[A-Za-z0-9Α-Ωα-ω]/.test(text)) return true;
    if (/[A-Za-zΑ-Ωα-ω]\s*[\+\-*]\s*[A-Za-z0-9Α-Ωα-ω]/.test(text)) return true;
    return false;
  }

  function isLikelyFormulaText(text) {
    if (!text) return false;
    var normalized = String(text).replace(/\u0007/g, '').trim();
    if (!normalized) return false;
    if (isFormulaExplanationText(normalized)) return false;
    if (isReferenceStandardText(normalized)) return false;
    if (isUnitOnlyText(normalized)) return false;

    var chineseCount = countChineseChars(normalized);
    var hasStrongMath = containsStrongMathFeatures(normalized);

    if (!hasStrongMath) return false;

    // 正文叙述里偶尔会嵌入 Re=... 这类表达，不应整体按公式段落处理
    if (chineseCount >= 8 && /[，。；：:]/.test(normalized) && !getFormulaNumberMatch(normalized)) {
      return false;
    }

    return true;
  }

  function getFormulaParagraphInfo(p) {
    var fullText = p.fullText || p.text || '';
    var text = String(fullText).trim();
    var info = {
      isFormula: false,
      hasOMath: false,
      hasFormulaNumber: false,
      formulaNumber: '',
      formulaBody: text
    };

    try {
      var para = doc.Paragraphs.Item(p.index);
      if (para && para.Range && para.Range.OMaths && para.Range.OMaths.Count > 0) {
        info.isFormula = true;
        info.hasOMath = true;
      }
    } catch (e) {}

    var formulaNumMatch = getFormulaNumberMatch(text);
    if (formulaNumMatch) {
      var body = text.substring(0, text.lastIndexOf(formulaNumMatch[0])).trim();
      info.hasFormulaNumber = true;
      info.formulaNumber = formulaNumMatch[0].trim();
      info.formulaBody = body;

      if (isLikelyFormulaText(body)) {
        info.isFormula = true;
      }
    }

    // 无公式编号但具有明显数学特征，也视为公式
    if (!info.isFormula && isLikelyFormulaText(text)) {
      info.isFormula = true;
      info.formulaBody = text;
    }

    return info;
  }

  // E-002, E-003: 公式格式检查
  // 在 formula_layout 或 format scope 时检查
  if (formulaLayoutOnly || scope === 'format' || scope === 'full_proofread' || scope === 'all') {
    for (var i = 0; i < paragraphs.length; i++) {
      var p = paragraphs[i];
      var formulaInfo = getFormulaParagraphInfo(p);
      if (formulaInfo.isFormula) {
        console.log('[E-002] 检测到公式段落 ' + p.index + ': ' + (p.text || '').substring(0, 30));

        // E-002: 公式居中
        var alignment = p.alignment || 0;
        // 带编号的公式通过 E-003 用“居中制表位 + 右对齐制表位”处理，不直接整段居中
        if (!formulaInfo.hasFormulaNumber && alignment !== 1) {  // 1 = 居中
          issues.push({
            index: p.index,
            rule: 'E-002',
            name: '公式居中',
            original: (p.text || '').substring(0, 50),
            message: '公式应居中对齐',
            autoFix: true,
            fixSpec: {
              alignment: 1
            }
          });
        }

        // E-003: 公式编号位置（检测编号是否右对齐）
        // 公式编号通常在段落末尾，格式为 (X-Y)
        if (formulaInfo.hasFormulaNumber) {
          var rawText = p.rawText || p.fullText || p.text || '';
          // 标准格式：左对齐段落 + 居中制表位 + 公式 + 右对齐制表位 + 编号
          if (!/^\t/.test(rawText) || rawText.indexOf('\t' + formulaInfo.formulaNumber) === -1) {
            issues.push({
              index: p.index,
              rule: 'E-003',
              name: '公式编号位置',
              original: rawText.substring(0, 50),
              message: '公式本体应居中，编号应右对齐',
              autoFix: true,
              fixSpec: {
                formatWithTabStops: true,
                formulaBody: formulaInfo.formulaBody,
                numberText: formulaInfo.formulaNumber
              }
            });
          }
        }
      }
    }
  }

  // ========== 页眉页脚检查（HF-001~003） ==========

  // HF-001~003: 页眉页脚格式检查
  // 在 header_footer 或 format scope 时检查
  if (headerFooterOnly || scope === 'format' || scope === 'full_proofread' || scope === 'all') {
    try {
      var sections = doc.Sections;
      if (sections && sections.Count > 0) {
        for (var s = 1; s <= sections.Count; s++) {
          var section = sections.Item(s);

          // HF-001: 页眉页脚字体字号检查
          // 检查页眉
          var header = section.Headers.Item(1);  // 1 = wdHeaderFooterPrimary
          if (header && header.Range) {
            var headerParas = getHeaderFooterEffectiveParagraphs(header);
            var headerNeedsFix = false;
            var headerObservedSize = 0;
            var headerObservedCnFont = '';
            var headerObservedEnFont = '';
            for (var hp = 0; hp < headerParas.length; hp++) {
              var headerPara = headerParas[hp];
              var headerFont = headerPara && headerPara.Range ? headerPara.Range.Font : null;
              var headerSize = headerFont && headerFont.Size ? headerFont.Size : 0;
              var headerCnFont = headerFont && headerFont.NameFarEast ? String(headerFont.NameFarEast) : '';
              var headerEnFont = headerFont && headerFont.Name ? String(headerFont.Name) : '';
              if (!headerObservedSize && headerSize) headerObservedSize = headerSize;
              if (!headerObservedCnFont && headerCnFont) headerObservedCnFont = headerCnFont;
              if (!headerObservedEnFont && headerEnFont) headerObservedEnFont = headerEnFont;
              if (!headerSize || Math.abs(headerSize - 9) > 0.5 ||
                  (headerCnFont && headerCnFont.indexOf('宋体') < 0 && headerCnFont.indexOf('SimSun') < 0) ||
                  (headerEnFont && headerEnFont.indexOf('Arial') < 0)) {
                headerNeedsFix = true;
                if (headerSize) headerObservedSize = headerSize;
                if (headerCnFont) headerObservedCnFont = headerCnFont;
                if (headerEnFont) headerObservedEnFont = headerEnFont;
                break;
              }
            }

            if (headerNeedsFix) {
              console.log('[HF-001] 页眉字体字号: cn=' + headerObservedCnFont + ', en=' + headerObservedEnFont + ', size=' + headerObservedSize + '，应为宋体/Arial/9磅');
              issues.push({
                rule: 'HF-001',
                name: '页眉字体字号',
                sectionIndex: s,
                original: '页眉: 中文字体=' + headerObservedCnFont + '，西文字体=' + headerObservedEnFont + '，字号=' + headerObservedSize + '磅',
                suggested: '中文字体宋体，西文字体Arial，小五号(9磅)',
                message: '页眉中文应为宋体，西文应为Arial，字号应为小五号(9磅)',
                autoFix: true,
                fixSpec: {
                  type: 'header',
                  sectionIndex: s,
                  fontSize: 9,
                  fontNameCn: '宋体',
                  fontNameEn: 'Arial'
                }
              });
            }
          }

          // 检查页脚
          var footer = section.Footers.Item(1);  // 1 = wdHeaderFooterPrimary
          if (footer && footer.Range) {
            var footerParas = getHeaderFooterEffectiveParagraphs(footer);
            var footerNeedsFix = false;
            var footerObservedSize = 0;
            var footerObservedCnFont = '';
            var footerObservedEnFont = '';
            for (var fp = 0; fp < footerParas.length; fp++) {
              var footerPara = footerParas[fp];
              var footerFont = footerPara && footerPara.Range ? footerPara.Range.Font : null;
              var footerSize = footerFont && footerFont.Size ? footerFont.Size : 0;
              var footerCnFont = footerFont && footerFont.NameFarEast ? String(footerFont.NameFarEast) : '';
              var footerEnFont = footerFont && footerFont.Name ? String(footerFont.Name) : '';
              if (!footerObservedSize && footerSize) footerObservedSize = footerSize;
              if (!footerObservedCnFont && footerCnFont) footerObservedCnFont = footerCnFont;
              if (!footerObservedEnFont && footerEnFont) footerObservedEnFont = footerEnFont;
              if (!footerSize || Math.abs(footerSize - 9) > 0.5 ||
                  (footerCnFont && footerCnFont.indexOf('宋体') < 0 && footerCnFont.indexOf('SimSun') < 0) ||
                  (footerEnFont && footerEnFont.indexOf('Arial') < 0)) {
                footerNeedsFix = true;
                if (footerSize) footerObservedSize = footerSize;
                if (footerCnFont) footerObservedCnFont = footerCnFont;
                if (footerEnFont) footerObservedEnFont = footerEnFont;
                break;
              }
            }

            if (footerNeedsFix) {
              console.log('[HF-001] 页脚字体字号: cn=' + footerObservedCnFont + ', en=' + footerObservedEnFont + ', size=' + footerObservedSize + '，应为宋体/Arial/9磅');
              issues.push({
                rule: 'HF-001',
                name: '页脚字体字号',
                sectionIndex: s,
                original: '页脚: 中文字体=' + footerObservedCnFont + '，西文字体=' + footerObservedEnFont + '，字号=' + footerObservedSize + '磅',
                suggested: '中文字体宋体，西文字体Arial，小五号(9磅)',
                message: '页脚中文应为宋体，西文应为Arial，字号应为小五号(9磅)',
                autoFix: true,
                fixSpec: {
                  type: 'footer',
                  sectionIndex: s,
                  fontSize: 9,
                  fontNameCn: '宋体',
                  fontNameEn: 'Arial'
                }
              });
            }
          }

          // HF-002: 页眉线检查（通长细线，线宽0.5磅）
          // 页眉线通过页眉段落下边框实现
          if (header && header.Range) {
            try {
              var headerPara = getHeaderFooterBorderParagraph(header, true);
              var headerBorders = headerPara && headerPara.Range ? headerPara.Range.ParagraphFormat.Borders : null;
              if (headerBorders) {
                var bottomBorder = headerBorders.Item(-3);  // wdBorderBottom = -3
                var borderEnabled = bottomBorder ? bottomBorder.Visible : false;
                var borderLineWidth = bottomBorder ? bottomBorder.LineWidth : 0;
                var borderLineStyle = bottomBorder ? bottomBorder.LineStyle : 0;

                // wdLineWidth050pt = 4, wdLineStyleSingle = 1
                console.log('[HF-002] 页眉线: visible=' + borderEnabled + ', style=' + borderLineStyle + ', lineWidth=' + borderLineWidth);

                // 检查是否需要修复：需要有下边框，且为单线、0.5磅
                // 如果没有边框或样式不对，添加修复
                if (!borderEnabled || borderLineStyle !== 1 || borderLineWidth !== 4) {
                  issues.push({
                    rule: 'HF-002',
                    name: '页眉线',
                    sectionIndex: s,
                    original: '页眉线: ' + (borderEnabled ? '样式不符' : '无'),
                    suggested: '通长细线，线宽0.5磅',
                    message: '页眉线应为通长细线，线宽0.5磅',
                    autoFix: true,
                    fixSpec: {
                      type: 'header_border',
                      sectionIndex: s,
                      borderType: 'bottom',
                      lineStyle: 1,  // wdLineStyleSingle
                      lineWidth: 4   // wdLineWidth050pt
                    }
                  });
                }
              }
            } catch (e) {
              console.log('[HF-002] 检查页眉线出错: ' + e.message);
            }
          }

          // HF-003: 页脚线检查（通长双细线，线宽0.5磅）
          // 页脚线通过页脚段落上边框实现
          if (footer && footer.Range) {
            try {
              var footerPara = getHeaderFooterBorderParagraph(footer, false);
              var footerBorders = footerPara && footerPara.Range ? footerPara.Range.ParagraphFormat.Borders : null;
              if (footerBorders) {
                var topBorder = footerBorders.Item(-1);  // wdBorderTop = -1
                var borderEnabled = topBorder ? topBorder.Visible : false;
                var borderLineWidth = topBorder ? topBorder.LineWidth : 0;
                var borderLineStyle = topBorder ? topBorder.LineStyle : 0;

                // wdLineWidth050pt = 4, wdLineStyleDouble = 7
                console.log('[HF-003] 页脚线: visible=' + borderEnabled + ', style=' + borderLineStyle + ', lineWidth=' + borderLineWidth);

                // 检查是否需要修复：需要有上边框，且为双线、0.5磅
                if (!borderEnabled || borderLineStyle !== 7 || borderLineWidth !== 4) {
                  issues.push({
                    rule: 'HF-003',
                    name: '页脚线',
                    sectionIndex: s,
                    original: '页脚线: ' + (borderEnabled ? '样式不符' : '无'),
                    suggested: '通长双细线，线宽0.5磅',
                    message: '页脚线应为通长双细线，线宽0.5磅',
                    autoFix: true,
                    fixSpec: {
                      type: 'footer_border',
                      sectionIndex: s,
                      borderType: 'top',
                      lineStyle: 7,  // wdLineStyleDouble
                      lineWidth: 4   // wdLineWidth050pt
                    }
                  });
                }
              }
            } catch (e) {
              console.log('[HF-003] 检查页脚线出错: ' + e.message);
            }
          }
        }
      }
    } catch (e) {
      console.log('[HF-001~003] 检查页眉页脚出错: ' + e.message);
    }
  }

  // ========== 页面设置检查（PG-001~002） ==========

  // PG-001~002: 页面设置检查
  // 在 page_setup 或 format scope 时检查
  if (pageSetupOnly || scope === 'format' || scope === 'full_proofread' || scope === 'all') {
    try {
      var sections = doc.Sections;
      if (sections && sections.Count > 0) {
        for (var s = 1; s <= sections.Count; s++) {
          var section = sections.Item(s);
          var pageSetup = section.PageSetup;

          if (pageSetup) {
            // PG-001: 页边距检查
            // 上下2.54cm，左右3.17cm
            // 1cm ≈ 28.35磅
            var topMargin = pageSetup.TopMargin;
            var bottomMargin = pageSetup.BottomMargin;
            var leftMargin = pageSetup.LeftMargin;
            var rightMargin = pageSetup.RightMargin;

            // 2.54cm = 72磅 (2.54 * 28.35 ≈ 72)
            // 3.17cm = 90磅 (3.17 * 28.35 ≈ 90)
            var expectedTopBottom = 72;  // 2.54cm in points
            var expectedLeftRight = 90;  // 3.17cm in points
            var tolerance = 5;  // 容差5磅

            console.log('[PG-001] 页边距: 上=' + topMargin + ', 下=' + bottomMargin + ', 左=' + leftMargin + ', 右=' + rightMargin);

            // 检查上下页边距
            if (Math.abs(topMargin - expectedTopBottom) > tolerance || Math.abs(bottomMargin - expectedTopBottom) > tolerance) {
              issues.push({
                rule: 'PG-001',
                name: '页边距(上下)',
                sectionIndex: s,
                original: '上: ' + (topMargin / 28.35).toFixed(2) + 'cm, 下: ' + (bottomMargin / 28.35).toFixed(2) + 'cm',
                suggested: '上下: 2.54cm',
                message: '页边距上下应为2.54cm',
                autoFix: true,
                fixSpec: {
                  type: 'page_margin_vertical',
                  sectionIndex: s,
                  topMargin: expectedTopBottom,
                  bottomMargin: expectedTopBottom
                }
              });
            }

            // 检查左右页边距
            if (Math.abs(leftMargin - expectedLeftRight) > tolerance || Math.abs(rightMargin - expectedLeftRight) > tolerance) {
              issues.push({
                rule: 'PG-001',
                name: '页边距(左右)',
                sectionIndex: s,
                original: '左: ' + (leftMargin / 28.35).toFixed(2) + 'cm, 右: ' + (rightMargin / 28.35).toFixed(2) + 'cm',
                suggested: '左右: 3.17cm',
                message: '页边距左右应为3.17cm',
                autoFix: true,
                fixSpec: {
                  type: 'page_margin_horizontal',
                  sectionIndex: s,
                  leftMargin: expectedLeftRight,
                  rightMargin: expectedLeftRight
                }
              });
            }

            // PG-002: 页眉页脚距边界检查
            // 页眉页脚距离边界1.5cm
            var headerDistance = pageSetup.HeaderDistance;
            var footerDistance = pageSetup.FooterDistance;

            // 1.5cm = 42.5磅 (1.5 * 28.35 ≈ 42.5)
            var expectedHeaderFooter = 42.5;  // 1.5cm in points

            console.log('[PG-002] 页眉页脚距边界: 页眉=' + headerDistance + ', 页脚=' + footerDistance);

            if (Math.abs(headerDistance - expectedHeaderFooter) > tolerance || Math.abs(footerDistance - expectedHeaderFooter) > tolerance) {
              issues.push({
                rule: 'PG-002',
                name: '页眉页脚距边界',
                sectionIndex: s,
                original: '页眉: ' + (headerDistance / 28.35).toFixed(2) + 'cm, 页脚: ' + (footerDistance / 28.35).toFixed(2) + 'cm',
                suggested: '均为1.5cm',
                message: '页眉页脚距离边界应为1.5cm',
                autoFix: true,
                fixSpec: {
                  type: 'header_footer_distance',
                  sectionIndex: s,
                  headerDistance: expectedHeaderFooter,
                  footerDistance: expectedHeaderFooter
                }
              });
            }
          }
        }
      }
    } catch (e) {
      console.log('[PG-001~002] 检查页面设置出错: ' + e.message);
    }
  }

  console.log('[checkFormat] 检查完成，发现 ' + issues.length + ' 个问题');
  return issues;
}

// ========== 内部函数：内容检查 ==========

function checkContent(structure, contentScope) {
  var issues = [];
  var doc = Application.ActiveDocument;

  // 默认 scope 为 'content'（执行全部）
  var scope = contentScope || 'content';
  var isValueScope = scope === 'value';
  var isPunctuationScope = scope === 'punctuation';
  var isTableContentScope = scope === 'table_content';
  var isFullContentScope = scope === 'content' || scope === 'all' || scope === 'full_proofread';

  // 是否执行各类规则
  var doValueCheck = isFullContentScope || isValueScope;
  var doPunctuationCheck = isFullContentScope || isPunctuationScope;
  var doTableContentCheck = isFullContentScope || isTableContentScope;

  console.log('[checkContent] scope=' + scope + ', doValueCheck=' + doValueCheck + ', doPunctuationCheck=' + doPunctuationCheck + ', doTableContentCheck=' + doTableContentCheck);

  // V-006: 温度偏差格式
  function checkTempBias(text, idx) {
    var re = /(\d+(?:\.\d+)?)\s*℃\s*±\s*(\d+(?:\.\d+)?)\s*℃/g;
    var m;
    while ((m = re.exec(text)) !== null) {
      issues.push({
        index: idx,
        rule: 'V-006',
        name: '温度偏差格式',
        original: m[0],
        suggested: m[1] + '±' + m[2] + '℃',
        message: '温度偏差应写作"数值±偏差℃"格式',
        autoFix: true
      });
    }
  }

  // V-007: 小数补零
  function checkDecimalZero(text, idx) {
    var re = /(?<![0-9])\.(\d+)/g;
    var m;
    while ((m = re.exec(text)) !== null) {
      var before = text.substring(Math.max(0, m.index - 5), m.index);
      if (/[0-9]\.$/.test(before)) continue;
      if (/\d\.\d$/.test(before)) continue;

      issues.push({
        index: idx,
        rule: 'V-007',
        name: '小数补零',
        original: m[0],
        suggested: '0' + m[0],
        message: '小数点前应补零',
        autoFix: true
      });
    }
  }

  // V-008: 数值范围格式
  function checkRangeFormat(text, idx) {
    var re = /(\d+(?:\.\d+)?)\s*([-—–])\s*(\d+(?:\.\d+)?)/g;
    var m;
    while ((m = re.exec(text)) !== null) {
      var fullMatch = m[0];
      var sep = m[2];
      var beforeIdx = m.index;
      var afterIdx = beforeIdx + fullMatch.length;

      // 跳过日期格式
      if (/\d{4}[-—–]\d{1,2}[-—–]\d{1,2}/.test(fullMatch)) continue;
      if (/\d{1,2}[-—–]\d{1,2}[-—–]\d{4}/.test(fullMatch)) continue;

      // 【重要】跳过括号内的公式编号 (4-5) 或 (4.3-1)
      var beforeChar = beforeIdx > 0 ? text.charAt(beforeIdx - 1) : '';
      var afterChar = afterIdx < text.length ? text.charAt(afterIdx) : '';
      if (beforeChar === '(' || beforeChar === '（') continue;

      // 跳过图表编号格式 图 4-1、表 1-2
      var context = text.substring(Math.max(0, beforeIdx - 10), beforeIdx);
      if (/[图表]\s*\d*\.?\d*$/u.test(context)) continue;

      // 跳过小数点后的编号格式 4.3-1（图表编号的一部分）
      if (beforeIdx >= 2 && text.charAt(beforeIdx - 1) === '.' && /\d/.test(text.charAt(beforeIdx - 2))) continue;

      // 【重要】跳过范围后紧跟单位的情况，交给 V-009 处理
      // 例如：1.5—2.0m 应该变成 1.5～2.0m
      var units = ['mm', 'm', 'cm', 'km', 'N', 'kN', 'Pa', 'kPa', 'MPa', '℃', '%'];
      var afterText = text.substring(afterIdx).trim();
      var hasUnitAfter = false;
      for (var u = 0; u < units.length; u++) {
        if (afterText.indexOf(units[u]) === 0) {
          hasUnitAfter = true;
          break;
        }
      }
      if (hasUnitAfter) continue;

      var num1 = parseFloat(m[1]);
      var num2 = parseFloat(m[3]);
      if (num1 < num2 && num2 - num1 < 1000) {
        issues.push({
          index: idx,
          rule: 'V-008',
          name: '数值范围格式',
          original: m[0],
          suggested: m[1] + '～' + m[3],
          message: '数值范围应使用波浪号"～"连接',
          autoFix: true
        });
      }
    }
  }

  // V-009: 单位范围格式
  function checkUnitRange(text, idx) {
    var units = ['mm', 'm', 'cm', 'km', 'N', 'kN', 'Pa', 'kPa', 'MPa', '℃'];
    var unitPattern = units.join('|');

    // 格式1: 数值单位-数值单位 (如 10mm-20mm → 10～20mm)
    var re1 = new RegExp('(\\d+(?:\\.\\d+)?)\\s*(' + unitPattern + ')\\s*([-—～])\\s*(\\d+(?:\\.\\d+)?)\\s*(' + unitPattern + ')', 'gi');
    var m;
    while ((m = re1.exec(text)) !== null) {
      if (m[2].toLowerCase() === m[5].toLowerCase() && m[3] !== '～') {
        issues.push({
          index: idx,
          rule: 'V-009',
          name: '单位范围格式',
          original: m[0],
          suggested: m[1] + '～' + m[4] + m[2],
          message: '单位范围应写作"数值～数值单位"格式',
          autoFix: true
        });
      }
    }

    // 格式2: 数值-数值单位 (如 1.5—2.0m → 1.5～2.0m)
    var re2 = new RegExp('(\\d+(?:\\.\\d+)?)\\s*([-—])\\s*(\\d+(?:\\.\\d+)?)\\s*(' + unitPattern + ')', 'gi');
    while ((m = re2.exec(text)) !== null) {
      // 确保不是格式1已经匹配过的情况（即第一个数值后面没有单位）
      var beforeIdx = m.index;
      var checkUnitBefore = text.substring(Math.max(0, beforeIdx - 10), beforeIdx);
      // 如果前面已经有一个单位，则跳过（格式1会处理）
      var unitBeforePattern = new RegExp('(' + unitPattern + ')\\s*$', 'i');
      if (unitBeforePattern.test(checkUnitBefore)) continue;

      issues.push({
        index: idx,
        rule: 'V-009',
        name: '单位范围格式',
        original: m[0],
        suggested: m[1] + '～' + m[3] + m[4],
        message: '单位范围应写作"数值～数值单位"格式',
        autoFix: true
      });
    }
  }

  // V-010: 百分比范围
  function checkPercentRange(text, idx) {
    var re = /(\d+(?:\.\d+)?)\s*～\s*(\d+(?:\.\d+)?)\s*%/g;
    var m;
    while ((m = re.exec(text)) !== null) {
      issues.push({
        index: idx,
        rule: 'V-010',
        name: '百分比范围',
        original: m[0],
        suggested: m[1] + '%～' + m[2] + '%',
        message: '百分比范围每个数值都应带百分号',
        autoFix: true
      });
    }
  }

  // V-011: 幂次范围
  function checkPowerRange(text, idx) {
    var re = /(\d+(?:\.\d+)?)\s*～\s*(\d+(?:\.\d+)?)\s*[×x]\s*10\^?\d*/gi;
    var m;
    while ((m = re.exec(text)) !== null) {
      issues.push({
        index: idx,
        rule: 'V-011',
        name: '幂次范围',
        original: m[0],
        message: '幂次范围应在每个数值后都写出幂次，建议人工修改',
        autoFix: false
      });
    }
  }

  // V-012: 体积尺寸
  // 规范要求：240mm×240mm×60mm，不应写成240×240×60mm
  function checkVolumeSize(text, idx) {
    // 匹配 "数字×数字×数字单位" 格式（只有最后有单位）
    var re = /(\d+(?:\.\d+)?)\s*[×xX]\s*(\d+(?:\.\d+)?)\s*[×xX]\s*(\d+(?:\.\d+)?)\s*(mm|cm|m)/gi;
    var m;
    while ((m = re.exec(text)) !== null) {
      var num1 = m[1];
      var num2 = m[2];
      var num3 = m[3];
      var unit = m[4];

      // 生成正确的格式：每个数字都带单位
      var suggested = num1 + unit + '×' + num2 + unit + '×' + num3 + unit;

      issues.push({
        index: idx,
        rule: 'V-012',
        name: '体积尺寸',
        original: m[0],
        suggested: suggested,
        message: '体积尺寸应在每个数值后都标明单位，如：' + suggested,
        autoFix: true
      });
    }
  }

  // V-013: 中文数值书写规范
  // 规范要求：分数、百分数和比例数应采用数学记号
  // 如：四分之三→3/4，百分之四十→40%，一比一点五→1:1.5
  function checkChineseNumber(text, idx) {
    var chineseNumMap = {
      '零': 0, '一': 1, '二': 2, '三': 3, '四': 4,
      '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
      '两': 2, '半': 0.5
    };

    // 匹配 "百分之X" 格式
    var percentRe = /百分之([零一二三四五六七八九十百千万]+)/g;
    var m;
    while ((m = percentRe.exec(text)) !== null) {
      var chineseNum = m[1];
      var arabicNum = convertChineseToArabic(chineseNum, chineseNumMap);
      if (arabicNum !== null) {
        issues.push({
          index: idx,
          rule: 'V-013',
          name: '中文数值书写',
          original: m[0],
          suggested: arabicNum + '%',
          message: '百分数应采用数学记号',
          autoFix: false
        });
      }
    }

    // 匹配 "X分之Y" 格式（分数）
    var fractionRe = /([零一二三四五六七八九十百千万]+)分之([零一二三四五六七八九十百千万]+)/g;
    while ((m = fractionRe.exec(text)) !== null) {
      var denominator = convertChineseToArabic(m[1], chineseNumMap);
      var numerator = convertChineseToArabic(m[2], chineseNumMap);
      if (denominator !== null && numerator !== null) {
        issues.push({
          index: idx,
          rule: 'V-013',
          name: '中文数值书写',
          original: m[0],
          suggested: numerator + '/' + denominator,
          message: '分数应采用数学记号（也可写为 ' + (numerator/denominator) + '）',
          autoFix: false
        });
      }
    }

    // 匹配 "X比Y" 格式（比例）
    var ratioRe = /([零一二三四五六七八九十百千万点]+)比([零一二三四五六七八九十百千万点]+)/g;
    while ((m = ratioRe.exec(text)) !== null) {
      var left = convertChineseToArabic(m[1], chineseNumMap);
      var right = convertChineseToArabic(m[2], chineseNumMap);
      if (left !== null && right !== null) {
        issues.push({
          index: idx,
          rule: 'V-013',
          name: '中文数值书写',
          original: m[0],
          suggested: left + ':' + right,
          message: '比例数应采用数学记号',
          autoFix: false
        });
      }
    }
  }

  // 辅助函数：中文数字转阿拉伯数字
  function convertChineseToArabic(chinese, map) {
    if (!chinese) return null;

    // 简单情况：直接映射
    if (map[chinese] !== undefined) {
      return map[chinese];
    }

    // 检查是否包含小数点（点）
    if (chinese.indexOf('点') !== -1) {
      var parts = chinese.split('点');
      if (parts.length === 2) {
        var intPart = convertChineseToArabic(parts[0], map);
        var decPart = convertChineseToArabic(parts[1], map);
        if (intPart !== null && decPart !== null) {
          // 计算小数部分的位数
          var decStr = String(decPart);
          var divisor = Math.pow(10, decStr.length);
          return intPart + decPart / divisor;
        }
      }
      return null;
    }

    // 处理组合数字（如"四十"、"一百二十五"）
    var result = 0;
    var temp = 0;
    for (var i = 0; i < chinese.length; i++) {
      var char = chinese[i];
      if (char === '十') {
        temp = temp === 0 ? 10 : temp * 10;
        result += temp;
        temp = 0;
      } else if (char === '百') {
        temp = temp === 0 ? 100 : temp * 100;
        result += temp;
        temp = 0;
      } else if (char === '千') {
        temp = temp === 0 ? 1000 : temp * 1000;
        result += temp;
        temp = 0;
      } else if (map[char] !== undefined) {
        temp = map[char];
      } else {
        return null; // 无法解析
      }
    }
    result += temp;
    return result > 0 ? result : null;
  }

  // M-001: 图名标点
  function checkCaptionPunct(text, idx) {
    // 支持阿拉伯数字和中文数字（一、二、三...十、百、千）
    // 例如：图1、图2-1、图一、图二、表1、表一
    if (/^[图表]\s*[\d\.一二三四五六七八九十百千]+/.test(text) && /[。！？]$/.test(text.trim())) {
      issues.push({
        index: idx,
        rule: 'M-001',
        name: '图表名标点',
        original: text.trim(),
        suggested: text.trim().replace(/[。！？]$/, ''),
        message: '图表名称末尾不应有句号等标点',
        autoFix: true
      });
    }
  }

  // M-002: 表名标点
  function checkTableCaptionPunct(text, idx) {
    if (/^表\s*[\d\.]+/.test(text) && /[。！？]$/.test(text.trim())) {
      issues.push({
        index: idx,
        rule: 'M-002',
        name: '表格名标点',
        original: text.trim(),
        suggested: text.trim().replace(/[。！？]$/, ''),
        message: '表格名称末尾不应有句号等标点',
        autoFix: true
      });
    }
  }

  // M-005: 中文括号
  function checkChineseBracket(text, idx) {
    var re = /\(([^)]*[\u4e00-\u9fff][^)]*)\)/g;
    var m;
    while ((m = re.exec(text)) !== null) {
      issues.push({
        index: idx,
        rule: 'M-005',
        name: '中文括号',
        original: m[0],
        suggested: '（' + m[1] + '）',
        message: '中文内容应使用中文括号"（）"',
        autoFix: true
      });
    }
  }

  // M-006: 范围符号
  function checkRangeSymbol(text, idx) {
    // 【重要】如果整行是图表编号格式，跳过检查
    // 例如："图 2.3-1  厂房剖面图"、"表 1.2-3  材料参数表"
    if (/^[图表]\s*\d+/.test(text)) {
      return;
    }

    // 匹配数值范围，但排除：
    // 1. 公式编号 (X-Y) 或 (X.Y-Z)
    // 2. 标准/规范编号 GB50201-2014、DL/T 5212-2005
    // 3. 日期格式
    var numRe = /(\d+)\s*[-—–]\s*(\d+)/g;
    var m;
    while ((m = numRe.exec(text)) !== null) {
      // 检查是否在括号内（公式编号）
      var beforeIdx = m.index;
      var afterIdx = m.index + m[0].length;
      var beforeChar = beforeIdx > 0 ? text.charAt(beforeIdx - 1) : '';
      var afterChar = afterIdx < text.length ? text.charAt(afterIdx) : '';

      // 跳过公式编号格式 (X-Y)
      if (beforeChar === '(' || beforeChar === '（') continue;

      // 跳过日期格式
      if (afterChar === '日' || afterChar === '月' || afterChar === '年') continue;

      // 跳过标准编号（前面有字母，如 GB50201-2014）
      if (beforeIdx > 0 && /[A-Za-z]/.test(text.charAt(beforeIdx - 1))) continue;

      // 跳过小数点后的数字（如 4.2-3 中的 2-3）
      // 如果前面是小数点加数字，说明是图表编号的一部分
      if (beforeIdx >= 2 && text.charAt(beforeIdx - 1) === '.' && /\d/.test(text.charAt(beforeIdx - 2))) continue;

      // 检查前后是否有单位（表示真正的数值范围）
      // 例如 "10-20m" 应改为 "10～20m"
      // 如果后面跟着单位或文字，说明是数值范围
      if (/[a-zA-Z\u4e00-\u9fa5]/.test(text.charAt(afterIdx)) || /\s/.test(text.charAt(afterIdx))) {
        issues.push({
          index: idx,
          rule: 'M-006',
          name: '范围符号',
          original: m[0],
          suggested: m[1] + '～' + m[2],
          message: '数值范围应使用波浪号"～"',
          autoFix: true
        });
      }
    }
  }

  // T-006: 表格同上/同左
  function checkSameAs(text, idx) {
    if (/同上|同左/.test(text)) {
      issues.push({
        index: idx,
        rule: 'T-006',
        name: '表格同上同左',
        original: text,
        message: '表格中不应使用"同上"或"同左"，应填写具体内容',
        autoFix: false
      });
    }
  }

  // T-001/G-001/E-001: 编号格式
  function checkNumberFormat(text, idx) {
    if (/^表\s*\d+/.test(text)) {
      var tableRe = /^表\s*(\d+(?:\.\d+)?)\s*[-—–]\s*(\d+)/;
      if (tableRe.test(text)) {
        issues.push({
          index: idx,
          rule: 'T-001',
          name: '表编号格式',
          original: text,
          message: '表格编号格式建议检查是否符合规范（如"表1.1-1"）',
          autoFix: false
        });
      }
    }

    if (/^图\s*\d+/.test(text)) {
      var figRe = /^图\s*(\d+(?:\.\d+)?)\s*[-—–]\s*(\d+)/;
      if (figRe.test(text)) {
        issues.push({
          index: idx,
          rule: 'G-001',
          name: '图编号格式',
          original: text,
          message: '图片编号格式建议检查是否符合规范（如"图1.1-1"）',
          autoFix: false
        });
      }
    }

    if (/^\(\d+\.?\d*-?\d+\)/.test(text)) {
      issues.push({
        index: idx,
        rule: 'E-001',
        name: '公式编号格式',
        original: text,
        message: '公式编号格式建议检查是否符合规范',
        autoFix: false
      });
    }
  }

  // E-004: 式中格式
  function checkFormulaNotation(text, idx) {
    if (/^式中[：:]/.test(text)) {
      issues.push({
        index: idx,
        rule: 'E-004',
        name: '式中格式',
        original: text,
        message: '公式说明"式中"后的符号说明格式建议检查',
        autoFix: false
      });
    }
  }

  var paragraphs = structure.paragraphs || [];
  for (var i = 0; i < paragraphs.length; i++) {
    var p = paragraphs[i];
    var text = p.fullText || p.text || '';  // 优先使用完整文本
    var idx = p.index;

    // V系列：数值格式检查
    if (doValueCheck) {
      checkTempBias(text, idx);
      checkDecimalZero(text, idx);
      checkRangeFormat(text, idx);
      checkUnitRange(text, idx);
      checkPercentRange(text, idx);
      checkPowerRange(text, idx);    // V-011 幂次范围
      checkVolumeSize(text, idx);    // V-012 体积尺寸
      checkChineseNumber(text, idx); // V-013 中文数值书写
    }

    // M系列：图表标点检查
    if (doPunctuationCheck) {
      checkCaptionPunct(text, idx);
      checkTableCaptionPunct(text, idx);
      checkChineseBracket(text, idx);
      // 注：M-006 范围符号检查已合并到 V-008/V-009 中，在 value scope 执行
    }
  }

  // T系列：表格内容检查（需要遍历表格）
  if (doTableContentCheck && doc && doc.Tables) {
    var tableIssues = checkTableContent(doc);
    issues = issues.concat(tableIssues);
  }

  return issues;
}

// ========== 内部函数：表格内容检查（T-005, T-006） ==========

/**
 * T-005: 同上替换
 * T-006: 同左替换
 * 将表格中的"同上"/"同左"替换为对应单元格的数值
 */
function checkTableContent(doc) {
  var issues = [];

  if (!doc || !doc.Tables || doc.Tables.Count === 0) {
    console.log('[checkTableContent] 无表格');
    return issues;
  }

  var tableCount = doc.Tables.Count;
  console.log('[checkTableContent] 开始检查，共 ' + tableCount + ' 个表格');

  for (var t = 1; t <= tableCount; t++) {
    try {
      var table = doc.Tables.Item(t);
      if (!table || !table.Rows) continue;

      var rowCount = table.Rows.Count;
      var colCount = table.Columns.Count;

      for (var r = 1; r <= rowCount; r++) {
        for (var c = 1; c <= colCount; c++) {
          try {
            var cell = table.Cell(r, c);
            if (!cell || !cell.Range) continue;

            var cellText = cell.Range.Text || '';
            cellText = String(cellText).replace(/\r/g, '').replace(/\n/g, '').replace(/[\x00-\x1f]/g, '').trim();

            // T-005: 检测"同上"
            if (cellText === '同上' || cellText === '同  上') {
              // 获取上方单元格的值
              if (r > 1) {
                try {
                  var aboveCell = table.Cell(r - 1, c);
                  var aboveText = aboveCell.Range.Text || '';
                  aboveText = String(aboveText).replace(/\r/g, '').replace(/\n/g, '').replace(/[\x00-\x1f]/g, '').trim();

                  // 确保上方单元格有内容且不是"同上"/"同左"
                  if (aboveText && aboveText !== '同上' && aboveText !== '同左' && aboveText !== '同  上' && aboveText !== '同  左') {
                    issues.push({
                      tableIndex: t,
                      rowIndex: r,
                      colIndex: c,
                      rule: 'T-005',
                      name: '同上替换',
                      original: cellText,
                      suggested: aboveText,
                      message: '表格中的"同上"应替换为上方单元格的值：' + aboveText,
                      autoFix: true
                    });
                  }
                } catch (e) {}
              }
            }

            // T-006: 检测"同左"
            if (cellText === '同左' || cellText === '同  左') {
              // 获取左侧单元格的值
              if (c > 1) {
                try {
                  var leftCell = table.Cell(r, c - 1);
                  var leftText = leftCell.Range.Text || '';
                  leftText = String(leftText).replace(/\r/g, '').replace(/\n/g, '').replace(/[\x00-\x1f]/g, '').trim();

                  // 确保左侧单元格有内容且不是"同上"/"同左"
                  if (leftText && leftText !== '同上' && leftText !== '同左' && leftText !== '同  上' && leftText !== '同  左') {
                    issues.push({
                      tableIndex: t,
                      rowIndex: r,
                      colIndex: c,
                      rule: 'T-006',
                      name: '同左替换',
                      original: cellText,
                      suggested: leftText,
                      message: '表格中的"同左"应替换为左侧单元格的值：' + leftText,
                      autoFix: true
                    });
                  }
                } catch (e) {}
              }
            }
          } catch (e) {}
        }
      }
    } catch (e) {}
  }

  console.log('[checkTableContent] 发现 ' + issues.length + ' 个问题');
  return issues;
}

function processTableContentFast(doc, mode) {
  var result = {
    fixed: 0,
    commented: 0,
    byRule: {},
    revisionLog: []
  };

  if (!doc || !doc.Tables || doc.Tables.Count === 0) {
    console.log('[processTableContentFast] 无表格');
    return result;
  }

  function cleanCellText(text) {
    return String(text || '').replace(/\r/g, '').replace(/\n/g, '').replace(/[\x00-\x1f]/g, '').replace(/\u0007/g, '').trim();
  }

  function normalizeMarker(text) {
    var normalized = cleanCellText(text).replace(/[\s　]+/g, '');
    if (normalized === '同上') return 'T-005';
    if (normalized === '同左') return 'T-006';
    return '';
  }

  function shouldAutoFix(ruleId) {
    if (mode === 'conservative') return false;
    return ruleId === 'T-005' || ruleId === 'T-006';
  }

  function getCellText(table, row, col, cache) {
    var key = row + ':' + col;
    if (cache[key] !== undefined) return cache[key];
    try {
      var cell = table.Cell(row, col);
      if (!cell || !cell.Range) {
        cache[key] = null;
        return cache[key];
      }
      cache[key] = cleanCellText(cell.Range.Text || '');
      return cache[key];
    } catch (e) {
      cache[key] = null;
      return cache[key];
    }
  }

  function resolveCellValue(table, row, col, direction, textCache, resolvedCache, visiting) {
    var cacheKey = direction + ':' + row + ':' + col;
    if (resolvedCache[cacheKey] !== undefined) return resolvedCache[cacheKey];
    if (visiting[cacheKey]) return '';
    visiting[cacheKey] = true;

    var sourceRow = row;
    var sourceCol = col;
    var step = 1;
    var value = '';

    while (true) {
      if (direction === 'up') {
        sourceRow = row - step;
        if (sourceRow < 1) break;
      } else {
        sourceCol = col - step;
        if (sourceCol < 1) break;
      }

      var sourceText = getCellText(table, sourceRow, sourceCol, textCache);
      var marker = normalizeMarker(sourceText);
      if (!sourceText) {
        step++;
        continue;
      }

      if (!marker) {
        value = sourceText;
        break;
      }

      if (marker === 'T-005') {
        value = resolveCellValue(table, sourceRow, sourceCol, 'up', textCache, resolvedCache, visiting);
      } else if (marker === 'T-006') {
        value = resolveCellValue(table, sourceRow, sourceCol, 'left', textCache, resolvedCache, visiting);
      }

      if (value) break;
      step++;
    }

    visiting[cacheKey] = false;
    resolvedCache[cacheKey] = value;
    return value;
  }

  var originalTrackRevisions = doc.TrackRevisions;
  doc.TrackRevisions = true;

  try {
    var tableCount = doc.Tables.Count;
    console.log('[processTableContentFast] 开始处理，共 ' + tableCount + ' 个表格');

    for (var t = 1; t <= tableCount; t++) {
      try {
        var table = doc.Tables.Item(t);
        if (!table || !table.Rows || !table.Columns) continue;

        var rowCount = table.Rows.Count;
        var colCount = table.Columns.Count;
        var textCache = {};
        var resolvedCache = {};

        for (var r = 1; r <= rowCount; r++) {
          for (var c = 1; c <= colCount; c++) {
            var cellText = getCellText(table, r, c, textCache);
            var rule = normalizeMarker(cellText);
            if (!rule) continue;

            var replacement = '';
            if (rule === 'T-005') {
              replacement = resolveCellValue(table, r, c, 'up', textCache, resolvedCache, {});
            } else if (rule === 'T-006') {
              replacement = resolveCellValue(table, r, c, 'left', textCache, resolvedCache, {});
            }

            if (!replacement) continue;

            try {
              var targetCell = table.Cell(r, c);
              if (!targetCell || !targetCell.Range) continue;

              if (shouldAutoFix(rule)) {
                targetCell.Range.Text = replacement;
                textCache[r + ':' + c] = replacement;
                result.fixed++;
              } else {
                addCommentToDoc(doc, targetCell.Range, {
                  rule: rule,
                  name: rule === 'T-005' ? '同上替换' : '同左替换',
                  original: cellText,
                  suggested: replacement,
                  message: (rule === 'T-005' ? '表格中的"同上"应替换为上方单元格的值：' : '表格中的"同左"应替换为左侧单元格的值：') + replacement
                });
                result.commented++;
              }

              result.byRule[rule] = (result.byRule[rule] || 0) + 1;
              result.revisionLog.push({
                rule: rule,
                name: rule === 'T-005' ? '同上替换' : '同左替换',
                tableIndex: t,
                rowIndex: r,
                colIndex: c,
                original: cellText,
                suggested: replacement,
                message: (rule === 'T-005' ? '表格中的"同上"应替换为上方单元格的值：' : '表格中的"同左"应替换为左侧单元格的值：') + replacement
              });
            } catch (cellErr) {}
          }
        }
      } catch (tableErr) {}
    }
  } finally {
    doc.TrackRevisions = originalTrackRevisions;
  }

  console.log('[processTableContentFast] 完成，修复: ' + result.fixed + '，批注: ' + result.commented);
  return result;
}

function processFontFast(doc, mode) {
  var result = { fixed: 0, commented: 0, byRule: {}, revisionLog: [], paraCount: 0 };
  if (!doc || !doc.Paragraphs) {
    console.log('[processFontFast] 无段落内容');
    return result;
  }

  function cleanText(text) {
    return String(text || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function normalizeTextForMatch(text) {
    return cleanText(text).replace(/\s+/g, ' ');
  }

  function normalizeFormulaCore(text) {
    return cleanText(text).replace(/\s+/g, '').replace(/[（）]/g, function(ch) {
      return ch === '（' ? '(' : ')';
    });
  }

  function normalizeStyleName(styleName) {
    return String(styleName || '').toUpperCase();
  }

  function getHeadingLevel(styleName, text) {
    if (/^附录\s*[A-Z一二三四五六七八九十]/i.test(text)) return 1;
    if (/^[A-Z]\d+\s/.test(text)) return 1;
    if (/^[A-Z]\d+\.\d+\s/.test(text)) return 2;
    if (/^[A-Z]\d+\.\d+\.\d+\s/.test(text)) return 3;
    if (/^第[一二三四五六七八九十百]+章/.test(text)) return 1;
    if (/^(\d+)\s+(?!\.)/.test(text)) return 1;
    if (/^\d+\.\d+\s/.test(text)) return 2;
    if (/^\d+\.\d+\.\d+\s/.test(text)) return 3;
    if (/^\d+\.\d+\.\d+\.\d+\s/.test(text)) return 4;

    var upperName = normalizeStyleName(styleName);
    if (upperName.indexOf('HEADING 1') >= 0 || upperName.indexOf('标题 1') >= 0 || upperName === '1') return 1;
    if (upperName.indexOf('HEADING 2') >= 0 || upperName.indexOf('标题 2') >= 0 || upperName === '2') return 2;
    if (upperName.indexOf('HEADING 3') >= 0 || upperName.indexOf('标题 3') >= 0 || upperName === '3') return 3;
    if (upperName.indexOf('HEADING 4') >= 0 || upperName.indexOf('标题 4') >= 0 || upperName === '4') return 4;
    return 0;
  }

  function isFigureCaption(text) {
    return /^图\s*[\dA-Za-z]/.test(text);
  }

  function isTableCaption(text) {
    return /^表\s*[\dA-Za-z]/.test(text);
  }

  function isFormulaParagraph(text) {
    return /\([A-Z]?\d+(?:\.\d+)*(?:-\d+)?\)\s*$/.test(text) || /\t\([A-Z]?\d+(?:\.\d+)*(?:-\d+)?\)\s*$/.test(text);
  }

  function hasInlineImage(paraRange) {
    try {
      if (paraRange && paraRange.InlineShapes && paraRange.InlineShapes.Count > 0) return true;
    } catch (e) {}
    try {
      if (paraRange && paraRange.ShapeRange && paraRange.ShapeRange.Count > 0) return true;
    } catch (e2) {}
    return false;
  }

  function isParagraphInTable(paraRange) {
    try {
      if (paraRange && paraRange.Tables && paraRange.Tables.Count > 0) return true;
    } catch (e) {}
    try {
      if (paraRange && paraRange.Cells && paraRange.Cells.Count > 0) return true;
    } catch (e2) {}
    return false;
  }

  function buildHeadingSpec(level) {
    return {
      rule: level === 1 ? 'F-001' : level === 2 ? 'F-002' : level === 3 ? 'F-003' : 'F-004',
      bold: level === 1 || level === 2
    };
  }

  function applyCommonFormat(targetRange, targetParaFormat, bold, firstIndent) {
    if (targetRange && targetRange.Font) {
      targetRange.Font.NameFarEast = '宋体';
      targetRange.Font.Name = '宋体';
      targetRange.Font.Size = 12;
      targetRange.Font.Bold = bold ? -1 : 0;
    }

    if (targetParaFormat) {
      if (targetParaFormat.LineSpacingRule !== undefined) {
        targetParaFormat.LineSpacingRule = 4;
        targetParaFormat.LineSpacing = 20;
      }
      if (targetParaFormat.SpaceBefore !== undefined) targetParaFormat.SpaceBefore = 6;
      if (targetParaFormat.SpaceAfter !== undefined) targetParaFormat.SpaceAfter = 6;
      if (targetParaFormat.FirstLineIndent !== undefined) targetParaFormat.FirstLineIndent = firstIndent;
    }
  }

  var paraCount = doc.Paragraphs.Count || 0;
  result.paraCount = paraCount;
  console.log('[processFontFast] 开始，段落数: ' + paraCount);

  var docText = doc.Content && doc.Content.Text ? String(doc.Content.Text) : '';
  var logicalParas = docText.split('\r');
  var maxParaCount = Math.min(paraCount, logicalParas.length);
  var headingPlans = [];
  var bodySegments = [];
  var bodyStartIndex = 0;
  var excludedParaMap = {};
  var useUltraFastMode = paraCount > 6000;

  try {
    var tableCount = doc.Tables ? doc.Tables.Count : 0;
    for (var tblIdx = 1; tblIdx <= tableCount; tblIdx++) {
      try {
        var table = doc.Tables.Item(tblIdx);
        if (!table || !table.Range || !table.Range.Paragraphs) continue;
        for (var tp = 1; tp <= table.Range.Paragraphs.Count; tp++) {
          try {
            var tablePara = table.Range.Paragraphs.Item(tp);
            if (tablePara && tablePara.Index) excludedParaMap[tablePara.Index] = 'table';
          } catch (tableParaErr) {}
        }
      } catch (tableErr) {}
    }
  } catch (tableOuterErr) {}

  try {
    var inlineCount = doc.InlineShapes ? doc.InlineShapes.Count : 0;
    for (var inIdx = 1; inIdx <= inlineCount; inIdx++) {
      try {
        var inlineShape = doc.InlineShapes.Item(inIdx);
        if (!inlineShape || !inlineShape.Range || !inlineShape.Range.Paragraphs || inlineShape.Range.Paragraphs.Count <= 0) continue;
        var imagePara = inlineShape.Range.Paragraphs.Item(1);
        if (imagePara && imagePara.Index) excludedParaMap[imagePara.Index] = 'image';
      } catch (inlineErr) {}
    }
  } catch (inlineOuterErr) {}

  for (var firstPassIndex = 1; firstPassIndex <= maxParaCount; firstPassIndex++) {
    var firstPassText = cleanText(logicalParas[firstPassIndex - 1]);
    if (!firstPassText) continue;
    if (getHeadingLevel('', firstPassText) === 1) {
      bodyStartIndex = firstPassIndex;
      break;
    }
  }

  if (!bodyStartIndex) {
    console.log('[processFontFast] 未找到正文起点，跳过排版');
    return result;
  }

  var currentBodyStart = 0;
  var currentBodyEnd = 0;
  var bodyParaCount = 0;

  for (var logicalIndex = bodyStartIndex; logicalIndex <= maxParaCount; logicalIndex++) {
    var text = cleanText(logicalParas[logicalIndex - 1]);
    if (!text) continue;

    var level = getHeadingLevel('', text);
    if (level >= 1 && level <= 4) {
      headingPlans.push({ index: logicalIndex, level: level, text: text });
      if (!useUltraFastMode && currentBodyStart) {
        bodySegments.push({ start: currentBodyStart, end: currentBodyEnd });
        currentBodyStart = 0;
        currentBodyEnd = 0;
      }
      continue;
    }

    if (excludedParaMap[logicalIndex]) {
      if (currentBodyStart) {
        bodySegments.push({ start: currentBodyStart, end: currentBodyEnd });
        currentBodyStart = 0;
        currentBodyEnd = 0;
      }
      continue;
    }

    if (isFigureCaption(text) || isTableCaption(text) || isFormulaParagraph(text)) {
      if (currentBodyStart) {
        bodySegments.push({ start: currentBodyStart, end: currentBodyEnd });
        currentBodyStart = 0;
        currentBodyEnd = 0;
      }
      continue;
    }

    if (!currentBodyStart) currentBodyStart = logicalIndex;
    currentBodyEnd = logicalIndex;
    bodyParaCount++;
  }

  if (!useUltraFastMode && currentBodyStart) {
    bodySegments.push({ start: currentBodyStart, end: currentBodyEnd });
  }

  var useWideBodyApply = bodySegments.length > 80;

  var originalTrackRevisions = doc.TrackRevisions;
  doc.TrackRevisions = true;

  try {
    if (bodyParaCount > 0 && mode !== 'conservative') {
      if (useWideBodyApply) {
        try {
          var bodyStartPara = doc.Paragraphs.Item(bodyStartIndex);
          if (bodyStartPara && bodyStartPara.Range) {
            var wideBodyRange = doc.Range(bodyStartPara.Range.Start, doc.Content.End);
            if (wideBodyRange) {
              applyCommonFormat(wideBodyRange, wideBodyRange.ParagraphFormat ? wideBodyRange.ParagraphFormat : null, false, 24);
              result.fixed += bodyParaCount;
            }
          }
        } catch (wideApplyErr) {}
      } else {
        for (var bs = 0; bs < bodySegments.length; bs++) {
          try {
            var segment = bodySegments[bs];
            var startPara = doc.Paragraphs.Item(segment.start);
            var endPara = doc.Paragraphs.Item(segment.end);
            if (!startPara || !startPara.Range || !endPara || !endPara.Range) continue;

            var segmentRange = doc.Range(startPara.Range.Start, endPara.Range.End);
            if (!segmentRange) continue;
            applyCommonFormat(segmentRange, segmentRange.ParagraphFormat ? segmentRange.ParagraphFormat : null, false, 24);
            result.fixed += (segment.end - segment.start + 1);
          } catch (bodyApplyErr) {}
        }
      }
      if (bodyParaCount > 0) {
        result.byRule['F-005'] = (result.byRule['F-005'] || 0) + bodyParaCount;
        result.revisionLog.push({
          rule: 'F-005',
          index: bodyStartIndex,
          original: '正文区段',
          suggested: '正文统一为宋体小四、20磅行距、0.5行段间距、首行缩进2字'
        });
      }
    } else if (bodyParaCount > 0) {
      result.commented++;
      result.byRule['F-005'] = (result.byRule['F-005'] || 0) + 1;
      result.revisionLog.push({
        rule: 'F-005',
        index: bodyStartIndex,
        original: '正文区段',
        suggested: '正文统一为宋体小四、20磅行距、0.5行段间距、首行缩进2字'
      });
    }

    for (var h = 0; h < headingPlans.length; h++) {
      try {
        var headingPlan = headingPlans[h];
        var para = doc.Paragraphs.Item(headingPlan.index);
        if (!para || !para.Range) continue;
        if (mode === 'conservative') {
          var headingSpec = buildHeadingSpec(headingPlan.level);
          addCommentToDoc(doc, para.Range, {
            rule: headingSpec.rule,
            name: '标题排版',
            original: headingPlan.text.substring(0, 50),
            suggested: '宋体小四、20磅行距、段前段后0.5行' + (headingSpec.bold ? '、加粗、无缩进' : '、无缩进'),
            message: '该段排版不符合标题/正文规范，建议按统一格式调整'
          });
          result.commented++;
        } else {
          headingSpec = buildHeadingSpec(headingPlan.level);
          applyCommonFormat(para.Range, para.Format, headingSpec.bold, 0);
          result.fixed++;
        }

        result.byRule[headingSpec.rule] = (result.byRule[headingSpec.rule] || 0) + 1;
        result.revisionLog.push({
          rule: headingSpec.rule,
          index: headingPlan.index,
          original: headingPlan.text.substring(0, 80),
          suggested: '标题统一为宋体小四、20磅行距、0.5行段间距'
        });
      } catch (headingErr) {}
    }
  } finally {
    doc.TrackRevisions = originalTrackRevisions;
  }

  console.log('[processFontFast] 完成，修复: ' + result.fixed + '，批注: ' + result.commented);
  return result;
}

function processFigureTableLayoutFast(doc, mode, scope) {
  var result = { fixed: 0, commented: 0, byRule: {}, revisionLog: [], paraCount: 0 };
  if (!doc) {
    console.log('[processFigureTableLayoutFast] 无文档内容');
    return result;
  }

  function cleanText(text) {
    return String(text || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function isFigureCaption(text) {
    return /^图\s*[\dA-Za-z]/.test(text);
  }

  function isTableCaption(text) {
    return /^表\s*[\dA-Za-z]/.test(text);
  }

  function canHandle(rule) {
    if (scope === 'figure_caption') return rule === 'G-002';
    if (scope === 'figure_center') return rule === 'G-004';
    if (scope === 'table_caption') return rule === 'T-002';
    if (scope === 'figure_layout') return rule === 'G-002' || rule === 'G-004';
    if (scope === 'table_layout') return rule === 'T-002' || rule === 'T-004' || rule === 'T-007' || rule === 'T-008';
    return true;
  }

  function track(rule, payload) {
    result.byRule[rule] = (result.byRule[rule] || 0) + 1;
    result.revisionLog.push(payload);
  }

  var paraCount = doc.Paragraphs ? doc.Paragraphs.Count : 0;
  result.paraCount = paraCount;
  console.log('[processFigureTableLayoutFast] 开始，段落数: ' + paraCount + ', 表格数: ' + (doc.Tables ? doc.Tables.Count : 0) + ', scope=' + scope);

  var originalTrackRevisions = doc.TrackRevisions;
  doc.TrackRevisions = true;

  try {
    if (canHandle('G-002') || canHandle('T-002')) {
      var docText = doc.Content && doc.Content.Text ? String(doc.Content.Text) : '';
      var logicalParas = docText.split('\r');
      var maxParaCount = Math.min(paraCount, logicalParas.length);

      for (var i = 1; i <= maxParaCount; i++) {
        var text = cleanText(logicalParas[i - 1]);
        if (!text) continue;

        var rule = '';
        var spaceBefore = 6;
        var spaceAfter = 6;
        if (isFigureCaption(text) && canHandle('G-002')) {
          rule = 'G-002';
        } else if (isTableCaption(text) && canHandle('T-002')) {
          rule = 'T-002';
          spaceAfter = 0;
        }
        if (!rule) continue;

        try {
          var para = doc.Paragraphs.Item(i);
          if (!para || !para.Range) continue;
          var range = para.Range;
          if (mode === 'conservative') {
            addCommentToDoc(doc, range, {
              rule: rule,
              name: rule === 'G-002' ? '图名排版' : '表名排版',
              original: text.substring(0, 50),
              suggested: '小四号宋体居中',
              message: rule === 'G-002' ? '图名应为小四号宋体居中，段前后0.5行' : '表名应为小四号宋体居中，段前0.5行段后0行'
            });
            result.commented++;
          } else {
            if (range.Font) {
              range.Font.NameFarEast = '宋体';
              range.Font.Name = '宋体';
              range.Font.Size = 12;
              range.Font.Bold = 0;
            }
            if (para.Format) {
              if (para.Format.Alignment !== undefined) para.Format.Alignment = 1;
              if (para.Format.SpaceBefore !== undefined) para.Format.SpaceBefore = spaceBefore;
              if (para.Format.SpaceAfter !== undefined) para.Format.SpaceAfter = spaceAfter;
            }
            result.fixed++;
          }

          track(rule, {
            rule: rule,
            index: i,
            original: text.substring(0, 80),
            suggested: rule === 'G-002' ? '图名统一为小四号宋体居中，段前后0.5行' : '表名统一为小四号宋体居中，段前0.5行段后0行'
          });
        } catch (captionErr) {}
      }
    }

    if (canHandle('G-004')) {
      try {
        var inlineCount = doc.InlineShapes ? doc.InlineShapes.Count : 0;
        for (var inlineIndex = 1; inlineIndex <= inlineCount; inlineIndex++) {
          try {
            var inlineShape = doc.InlineShapes.Item(inlineIndex);
            if (!inlineShape || !inlineShape.Range || !inlineShape.Range.Paragraphs || inlineShape.Range.Paragraphs.Count <= 0) continue;
            var imagePara = inlineShape.Range.Paragraphs.Item(1);
            if (!imagePara || !imagePara.Format) continue;

            if (mode === 'conservative') {
              addCommentToDoc(doc, imagePara.Range, {
                rule: 'G-004',
                name: '图片居中',
                original: '图片段落',
                suggested: '图片段落居中',
                message: '图片应居中对齐'
              });
              result.commented++;
            } else {
              imagePara.Format.Alignment = 1;
              try {
                if (imagePara.Format.FirstLineIndent !== undefined) imagePara.Format.FirstLineIndent = 0;
                if (imagePara.Format.LeftIndent !== undefined) imagePara.Format.LeftIndent = 0;
                if (imagePara.Format.RightIndent !== undefined) imagePara.Format.RightIndent = 0;
                if (imagePara.Format.SpaceBefore !== undefined) imagePara.Format.SpaceBefore = 6;
                if (imagePara.Format.SpaceAfter !== undefined) imagePara.Format.SpaceAfter = 6;
                if (imagePara.Format.LineSpacingRule !== undefined) imagePara.Format.LineSpacingRule = 0;
              } catch (imageFormatErr) {}
              try {
                if (imagePara.Range && imagePara.Range.ParagraphFormat) {
                  if (imagePara.Range.ParagraphFormat.FirstLineIndent !== undefined) imagePara.Range.ParagraphFormat.FirstLineIndent = 0;
                  if (imagePara.Range.ParagraphFormat.LeftIndent !== undefined) imagePara.Range.ParagraphFormat.LeftIndent = 0;
                  if (imagePara.Range.ParagraphFormat.RightIndent !== undefined) imagePara.Range.ParagraphFormat.RightIndent = 0;
                  if (imagePara.Range.ParagraphFormat.SpaceBefore !== undefined) imagePara.Range.ParagraphFormat.SpaceBefore = 6;
                  if (imagePara.Range.ParagraphFormat.SpaceAfter !== undefined) imagePara.Range.ParagraphFormat.SpaceAfter = 6;
                  if (imagePara.Range.ParagraphFormat.LineSpacingRule !== undefined) imagePara.Range.ParagraphFormat.LineSpacingRule = 0;
                }
              } catch (imageRangeFormatErr) {}
              result.fixed++;
            }
            track('G-004', { rule: 'G-004', index: imagePara.Range.Start, original: '图片段落', suggested: '图片居中对齐' });
          } catch (inlineErr) {}
        }
      } catch (imageErr) {}
    }

    if (canHandle('T-004') || canHandle('T-007') || canHandle('T-008')) {
      try {
        var tableCount = doc.Tables ? doc.Tables.Count : 0;
        var contentWidth = doc.PageSetup ? (doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin) : 0;
        for (var t = 1; t <= tableCount; t++) {
          try {
            var table = doc.Tables.Item(t);
            if (!table) continue;

            if (canHandle('T-004')) {
              if (mode === 'conservative') {
                result.commented++;
              } else {
                try { table.PreferredWidthType = 3; } catch (widthTypeErr) {}
                try { table.PreferredWidth = contentWidth; } catch (widthErr) {}
                result.fixed++;
              }
              track('T-004', { rule: 'T-004', tableIndex: t, original: '表格宽度', suggested: '表格与页面等宽' });
            }

            if (canHandle('T-007') && table.Rows && table.Rows.Count > 0) {
              var firstRow = table.Rows.Item(1);
              if (mode === 'conservative') {
                result.commented++;
              } else {
                try { firstRow.HeadingFormat = -1; } catch (headingErr) {}
                result.fixed++;
              }
              track('T-007', { rule: 'T-007', tableIndex: t, original: '首行表头', suggested: '跨页重复首行表头' });
            }

            if (canHandle('T-008')) {
              if (mode === 'conservative') {
                result.commented++;
              } else {
                try {
                  try {
                    if (table.TopPadding !== undefined) table.TopPadding = 0;
                    if (table.BottomPadding !== undefined) table.BottomPadding = 0;
                    if (table.LeftPadding !== undefined) table.LeftPadding = 0;
                    if (table.RightPadding !== undefined) table.RightPadding = 0;
                    if (table.Spacing !== undefined) table.Spacing = 0;
                  } catch (tablePaddingErr) {}

                  try {
                    if (table.Rows) {
                      for (var rowIdx = 1; rowIdx <= table.Rows.Count; rowIdx++) {
                        try {
                          var row = table.Rows.Item(rowIdx);
                          if (!row) continue;
                          if (row.HeightRule !== undefined) row.HeightRule = 0;
                          if (row.Height !== undefined) row.Height = 0;
                        } catch (rowErr) {}
                      }
                    }
                  } catch (rowsOuterErr) {}

                  if (table.Range && table.Range.ParagraphFormat && table.Range.ParagraphFormat.Alignment !== undefined) {
                    table.Range.ParagraphFormat.Alignment = 1;
                    try {
                      if (table.Range.ParagraphFormat.FirstLineIndent !== undefined) table.Range.ParagraphFormat.FirstLineIndent = 0;
                      if (table.Range.ParagraphFormat.LeftIndent !== undefined) table.Range.ParagraphFormat.LeftIndent = 0;
                      if (table.Range.ParagraphFormat.RightIndent !== undefined) table.Range.ParagraphFormat.RightIndent = 0;
                      if (table.Range.ParagraphFormat.SpaceBefore !== undefined) table.Range.ParagraphFormat.SpaceBefore = 0;
                      if (table.Range.ParagraphFormat.SpaceAfter !== undefined) table.Range.ParagraphFormat.SpaceAfter = 0;
                      if (table.Range.ParagraphFormat.LineSpacingRule !== undefined) table.Range.ParagraphFormat.LineSpacingRule = 0;
                    } catch (tableRangeFormatErr) {}
                  } else if (table.Range && table.Range.Paragraphs) {
                    for (var pIdx = 1; pIdx <= table.Range.Paragraphs.Count; pIdx++) {
                      try {
                        var tablePara = table.Range.Paragraphs.Item(pIdx);
                        if (tablePara && tablePara.Format && tablePara.Format.Alignment !== undefined) {
                          tablePara.Format.Alignment = 1;
                          if (tablePara.Format.FirstLineIndent !== undefined) tablePara.Format.FirstLineIndent = 0;
                          if (tablePara.Format.LeftIndent !== undefined) tablePara.Format.LeftIndent = 0;
                          if (tablePara.Format.RightIndent !== undefined) tablePara.Format.RightIndent = 0;
                          if (tablePara.Format.SpaceBefore !== undefined) tablePara.Format.SpaceBefore = 0;
                          if (tablePara.Format.SpaceAfter !== undefined) tablePara.Format.SpaceAfter = 0;
                          if (tablePara.Format.LineSpacingRule !== undefined) tablePara.Format.LineSpacingRule = 0;
                        }
                      } catch (tableParaErr) {}
                    }
                  }

                  try {
                    if (table.Range && table.Range.Cells) {
                      for (var cellIdx = 1; cellIdx <= table.Range.Cells.Count; cellIdx++) {
                        try {
                          var cell = table.Range.Cells.Item(cellIdx);
                          if (!cell) continue;
                          if (cell.VerticalAlignment !== undefined) cell.VerticalAlignment = 1;
                          if (cell.Range && cell.Range.ParagraphFormat) {
                            if (cell.Range.ParagraphFormat.Alignment !== undefined) cell.Range.ParagraphFormat.Alignment = 1;
                            if (cell.Range.ParagraphFormat.FirstLineIndent !== undefined) cell.Range.ParagraphFormat.FirstLineIndent = 0;
                            if (cell.Range.ParagraphFormat.LeftIndent !== undefined) cell.Range.ParagraphFormat.LeftIndent = 0;
                            if (cell.Range.ParagraphFormat.RightIndent !== undefined) cell.Range.ParagraphFormat.RightIndent = 0;
                            if (cell.Range.ParagraphFormat.SpaceBefore !== undefined) cell.Range.ParagraphFormat.SpaceBefore = 0;
                            if (cell.Range.ParagraphFormat.SpaceAfter !== undefined) cell.Range.ParagraphFormat.SpaceAfter = 0;
                            if (cell.Range.ParagraphFormat.LineSpacingRule !== undefined) cell.Range.ParagraphFormat.LineSpacingRule = 0;
                          }
                        } catch (cellErr) {}
                      }
                    }
                  } catch (cellsOuterErr) {}
                } catch (tableAlignErr) {}
                result.fixed++;
              }
              track('T-008', { rule: 'T-008', tableIndex: t, original: '表格文字对齐', suggested: '表格内文字居中' });
            }
          } catch (tableErr) {}
        }
      } catch (tableOuterErr) {}
    }
  } finally {
    doc.TrackRevisions = originalTrackRevisions;
  }

  console.log('[processFigureTableLayoutFast] 完成，修复: ' + result.fixed + '，批注: ' + result.commented);
  return result;
}

function processFormulaLayoutFast(doc, mode) {
  var result = { fixed: 0, commented: 0, byRule: {}, revisionLog: [], paraCount: 0 };
  if (!doc || !doc.Paragraphs) {
    console.log('[processFormulaLayoutFast] 无段落内容');
    return result;
  }

  function cleanText(text) {
    return String(text || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function normalizeTextForMatch(text) {
    return cleanText(text).replace(/\s+/g, ' ');
  }

  function normalizeFormulaCore(text) {
    return cleanText(text)
      .replace(/\s+/g, '')
      .replace(/[（）]/g, function(ch) {
        return ch === '（' ? '(' : ')';
      })
      .replace(/[－—–]/g, '-');
  }

  function getFormulaNumberMatch(text) {
    if (!text) return null;
    return String(text).match(/\(([A-Z]?\d+(?:\.\d+)*(?:[-—–]\d+)?)\)\s*$/);
  }

  function containsStrongMathFeatures(text) {
    if (!text) return false;
    if (/[=±×÷∑ΣΔ√∫≈≠≤≥]/.test(text)) return true;
    if (/[λρμτωφψβγαεδ]/.test(text)) return true;
    if (/[A-Za-zΑ-Ωα-ω][A-Za-z0-9Α-Ωα-ω₀₁₂₃₄₅₆₇₈₉²³]*\s*=\s*[^=]/.test(text)) return true;
    if (/[A-Za-zΑ-Ωα-ω]\s*\/\s*[A-Za-z0-9Α-Ωα-ω]/.test(text)) return true;
    if (/[A-Za-zΑ-Ωα-ω]\s*[\+\-*]\s*[A-Za-z0-9Α-Ωα-ω]/.test(text)) return true;
    return false;
  }

  function isLikelyFormulaText(text) {
    var normalized = cleanText(text);
    if (!normalized) return false;
    if (/^\s*式中\s*[：:]/.test(normalized)) return false;
    if (/《.+》/.test(normalized)) return false;
    if (/\b(?:GB|GB\/T|DL\/T|NB\/T|SL|SDJ|IEC|ISO)\s*\d+(?:[-—–]\d+)+/i.test(normalized)) return false;
    if (/^(\d+(?:\.\d+)?)?\s*[A-Za-z%℃°/·²³⁴⁵⁶⁷⁸⁹₀₁₂₃₄₅₆₇₈₉]+$/.test(normalized)) return false;
    if (!containsStrongMathFeatures(normalized)) return false;
    if ((normalized.match(/[\u4e00-\u9fa5]/g) || []).length >= 8 && /[，。；：:]/.test(normalized) && !getFormulaNumberMatch(normalized)) {
      return false;
    }
    return true;
  }

  function isLikelyNumberedFormulaBody(text) {
    var normalized = cleanText(text);
    if (!normalized) return false;
    if (/^\s*式中\s*[：:]/.test(normalized)) return false;
    if (/《.+》/.test(normalized)) return false;
    if ((normalized.match(/[\u4e00-\u9fa5]/g) || []).length > 12 && /[，。；：:]/.test(normalized)) return false;
    if (normalized.length <= 40 && /[A-Za-zΑ-Ωα-ω]/.test(normalized) && /=/.test(normalized)) return true;
    if (/[=±×÷∑ΣΔ√∫≈≠≤≥]/.test(normalized)) return true;
    if (/[λρμτωφψβγαεδσ]/.test(normalized)) return true;
    if (/[A-Za-zΑ-Ωα-ω]/.test(normalized) && /[\/\+\-*]/.test(normalized)) return true;
    if (/[A-Za-zΑ-Ωα-ω]\s*=\s*[^=]/.test(normalized)) return true;
    return false;
  }

  var paraCount = doc.Paragraphs.Count || 0;
  result.paraCount = paraCount;
  console.log('[processFormulaLayoutFast] 开始，段落数: ' + paraCount);

  var docText = doc.Content && doc.Content.Text ? String(doc.Content.Text) : '';
  var logicalParas = docText.split('\r');
  var maxParaCount = Math.min(paraCount, logicalParas.length);
  var plans = [];
  var paragraphTextCache = {};

  function getParagraphTextByIndex(index) {
    if (paragraphTextCache[index] !== undefined) return paragraphTextCache[index];
    try {
      var para = doc.Paragraphs.Item(index);
      paragraphTextCache[index] = para && para.Range ? normalizeTextForMatch(para.Range.Text || '') : '';
    } catch (e) {
      paragraphTextCache[index] = '';
    }
    return paragraphTextCache[index];
  }

  function getParagraphCoreByIndex(index) {
    return normalizeFormulaCore(getParagraphTextByIndex(index));
  }

  function paragraphLooksLikeFormula(index) {
    try {
      var para = doc.Paragraphs.Item(index);
      if (!para || !para.Range) return false;
      var raw = String(para.Range.Text || '');
      if (getFormulaNumberMatch(raw)) return true;
      if (para.Range.OMaths && para.Range.OMaths.Count > 0) return true;
      return isLikelyNumberedFormulaBody(raw) || isLikelyFormulaText(raw);
    } catch (e) {
      return false;
    }
  }

  function findParagraphByFormulaNumber(numberText) {
    if (!numberText || !doc.Content || !doc.Content.Find) return null;
    try {
      var searchRange = doc.Content.Duplicate ? doc.Content.Duplicate : doc.Content;
      searchRange.Find.ClearFormatting();
      searchRange.Find.Forward = true;
      searchRange.Find.Wrap = 0;
      searchRange.Find.MatchWildcards = false;
      var found = searchRange.Find.Execute(numberText, false, false, false, false, false, true, 1, false);
      if (!found) return null;
      if (searchRange.Paragraphs && searchRange.Paragraphs.Count > 0) {
        return searchRange.Paragraphs.Item(1);
      }
    } catch (e) {}
    return null;
  }

  function tryApplyFormulaTabStops(target, centerPos, rightPos) {
    if (!target || !target.TabStops) return false;
    try {
      target.TabStops.ClearAll();
      target.TabStops.Add(centerPos, 1, 0);
      target.TabStops.Add(rightPos, 2, 0);
      return true;
    } catch (e) {
      return false;
    }
  }

  function applyFormulaParagraphLayout(para, centerPos, rightPos) {
    var applied = false;
    try {
      if (para && para.Range && para.Range.ParagraphFormat) {
        if (para.Range.ParagraphFormat.Alignment !== undefined) para.Range.ParagraphFormat.Alignment = 0;
        applied = tryApplyFormulaTabStops(para.Range.ParagraphFormat, centerPos, rightPos) || applied;
      }
    } catch (e1) {}
    try {
      if (para && para.Range && para.Range.Paragraphs && para.Range.Paragraphs.Count > 0) {
        var firstPara = para.Range.Paragraphs.Item(1);
        if (firstPara && firstPara.Format) {
          if (firstPara.Format.Alignment !== undefined) firstPara.Format.Alignment = 0;
          applied = tryApplyFormulaTabStops(firstPara.Format, centerPos, rightPos) || applied;
        }
      }
    } catch (e2) {}
    try {
      if (para && para.Format) {
        if (para.Format.Alignment !== undefined) para.Format.Alignment = 0;
        applied = tryApplyFormulaTabStops(para.Format, centerPos, rightPos) || applied;
      }
    } catch (e3) {}
    return applied;
  }

  function resolveParagraphIndex(plan) {
    var logicalIndex = plan.index;
    var normalizedExpected = normalizeTextForMatch(plan.expectedText || plan.original || '');
    var normalizedBody = normalizeTextForMatch(plan.formulaBody || '');
    var normalizedNumber = normalizeTextForMatch(plan.numberText || '');
    var coreBody = normalizeFormulaCore(plan.formulaBody || '');
    var coreNumber = normalizeFormulaCore(plan.numberText || '');
    var offsets = [0, -1, 1, -2, 2, -3, 3, -5, 5, -8, 8, -12, 12, -16, 16, -20, 20, -30, 30, -40, 40];
    var weakCandidate = 0;
    for (var oi = 0; oi < offsets.length; oi++) {
      var candidate = logicalIndex + offsets[oi];
      if (candidate < 1 || candidate > paraCount) continue;
      var candidateText = getParagraphTextByIndex(candidate);
      if (!candidateText) continue;
      var candidateCore = normalizeFormulaCore(candidateText);

      if (plan.rule === 'E-003') {
        if (normalizedNumber && candidateText.indexOf(normalizedNumber) < 0 && (!coreNumber || candidateCore.indexOf(coreNumber) < 0)) continue;
        if (normalizedBody && candidateText.indexOf(normalizedBody) < 0 && (!coreBody || candidateCore.indexOf(coreBody) < 0)) {
          if (!weakCandidate && paragraphLooksLikeFormula(candidate)) weakCandidate = candidate;
          continue;
        }
        return candidate;
      }

      if (candidateText === normalizedExpected) return candidate;
      if (candidateText && normalizedExpected && (candidateText.indexOf(normalizedExpected) >= 0 || normalizedExpected.indexOf(candidateText) >= 0)) return candidate;
    }
    if (plan.rule === 'E-003' && weakCandidate) return weakCandidate;
    if (logicalIndex >= 1 && logicalIndex <= paraCount && paragraphLooksLikeFormula(logicalIndex)) return logicalIndex;
    return 0;
  }

  for (var i = 1; i <= maxParaCount; i++) {
    var text = cleanText(logicalParas[i - 1]);
    if (!text) continue;

    var numMatch = getFormulaNumberMatch(text);
    if (numMatch) {
      var formulaBody = cleanText(text.substring(0, text.lastIndexOf(numMatch[0])));
      if (isLikelyNumberedFormulaBody(formulaBody) || isLikelyFormulaText(formulaBody)) {
        plans.push({ index: i, rule: 'E-003', formulaBody: formulaBody, numberText: numMatch[0].trim(), original: text, expectedText: text });
      }
      continue;
    }

    if (isLikelyFormulaText(text)) {
      plans.push({ index: i, rule: 'E-002', original: text, expectedText: text });
    }
  }

  console.log('[processFormulaLayoutFast] 识别到候选公式: ' + plans.length);

  var originalTrackRevisions = doc.TrackRevisions;
  doc.TrackRevisions = true;

  try {
    for (var p = 0; p < plans.length; p++) {
      try {
        var plan = plans[p];
        var para = null;
        var resolvedIndex = 0;
        if (plan.rule === 'E-003') {
          para = findParagraphByFormulaNumber(plan.numberText);
          if (para && para.Range) {
            try {
              resolvedIndex = para.Range.Paragraphs && para.Range.Paragraphs.Count > 0 ? para.Range.Paragraphs.Item(1).Index : 0;
            } catch (indexErr) {
              resolvedIndex = 0;
            }
          }
        }

        if (!para) {
          resolvedIndex = resolveParagraphIndex(plan);
          if (!resolvedIndex) continue;
          para = doc.Paragraphs.Item(resolvedIndex);
        }

        if (!para || !para.Range) continue;
        if (!resolvedIndex) {
          try { resolvedIndex = para.Index || 0; } catch (paraIndexErr) { resolvedIndex = 0; }
        }
        var actualText = cleanText(para.Range.Text || '');

        if (mode === 'conservative') {
          addCommentToDoc(doc, para.Range, {
            rule: plan.rule,
            name: plan.rule === 'E-003' ? '公式编号位置' : '公式居中',
            original: actualText.substring(0, 50),
            suggested: plan.rule === 'E-003' ? '公式本身居中，编号右对齐' : '公式居中',
            message: plan.rule === 'E-003' ? '公式本体应居中，编号应右对齐' : '公式应居中对齐'
          });
          result.commented++;
        } else if (plan.rule === 'E-003') {
          var range = para.Range;
          var rawText = String(range.Text || '');
          var actualNumMatch = getFormulaNumberMatch(rawText);
          var actualBody = actualNumMatch ? cleanText(rawText.substring(0, rawText.lastIndexOf(actualNumMatch[0]))) : plan.formulaBody;
          var actualNumber = actualNumMatch ? actualNumMatch[0].trim() : plan.numberText;
          var paraMark = /\r$/.test(rawText) ? '\r' : '';
          var newText = '\t' + actualBody + '\t' + actualNumber + paraMark;

          if (newText !== rawText) {
            range.Text = newText;
          }

          var pageWidth = doc.PageSetup.PageWidth;
          var leftMargin = doc.PageSetup.LeftMargin;
          var rightMargin = doc.PageSetup.RightMargin;
          var contentWidth = pageWidth - leftMargin - rightMargin;
          var centerPos = contentWidth / 2;
          var rightPos = contentWidth;

          var formatPara = para;
          try {
            if (range.Paragraphs && range.Paragraphs.Count > 0) {
              formatPara = range.Paragraphs.Item(1);
            }
          } catch (rangeParaErr) {}
          applyFormulaParagraphLayout(formatPara, centerPos, rightPos);
          if (resolvedIndex && (!formatPara || !formatPara.Range)) {
            try {
              applyFormulaParagraphLayout(doc.Paragraphs.Item(resolvedIndex), centerPos, rightPos);
            } catch (fallbackFormatErr) {}
          }
          result.fixed++;
        } else {
          if (para.Format && para.Format.Alignment !== undefined) {
            para.Format.Alignment = 1;
          }
          result.fixed++;
        }

        result.byRule[plan.rule] = (result.byRule[plan.rule] || 0) + 1;
        result.revisionLog.push({
          rule: plan.rule,
          index: resolvedIndex,
          original: actualText.substring(0, 80),
          suggested: plan.rule === 'E-003' ? '公式本身居中，编号右对齐' : '公式居中'
        });
      } catch (planErr) {
        console.log('[processFormulaLayoutFast] 跳过候选，原因: ' + planErr);
      }
    }
  } finally {
    doc.TrackRevisions = originalTrackRevisions;
  }

  console.log('[processFormulaLayoutFast] 完成，修复: ' + result.fixed + '，批注: ' + result.commented);
  return result;
}

function processHeaderFooterFast(doc, mode) {
  var result = { fixed: 0, commented: 0, byRule: {}, revisionLog: [], sectionCount: 0 };
  if (!doc || !doc.Sections) {
    console.log('[processHeaderFooterFast] 无节对象');
    return result;
  }

  function pushLog(rule, sectionIndex, message) {
    result.byRule[rule] = (result.byRule[rule] || 0) + 1;
    result.revisionLog.push({
      rule: rule,
      sectionIndex: sectionIndex,
      message: message
    });
  }

  var sectionCount = doc.Sections.Count || 0;
  result.sectionCount = sectionCount;
  console.log('[processHeaderFooterFast] 开始，节数: ' + sectionCount);

  var originalTrackRevisions = doc.TrackRevisions;
  doc.TrackRevisions = true;

  try {
    for (var s = 1; s <= sectionCount; s++) {
      try {
        var section = doc.Sections.Item(s);
        if (!section) continue;

        var header = section.Headers ? section.Headers.Item(1) : null;
        var footer = section.Footers ? section.Footers.Item(1) : null;

        if (mode === 'conservative') {
          result.commented += 3;
          pushLog('HF-001', s, '页眉页脚统一为宋体/Arial，小五号');
          pushLog('HF-002', s, '页眉下设0.5磅通长细线');
          pushLog('HF-003', s, '页脚上设0.5磅通长双细线');
          continue;
        }

        if (header && header.Range) {
          var headerParas = getHeaderFooterEffectiveParagraphs(header);
          for (var hp = 0; hp < headerParas.length; hp++) {
            try {
              var headerPara = headerParas[hp];
              if (headerPara && headerPara.Range && headerPara.Range.Font) {
                headerPara.Range.Font.NameFarEast = '宋体';
                headerPara.Range.Font.Name = 'Arial';
                headerPara.Range.Font.Size = 9;
              }
            } catch (headerFontErr) {}
          }
          clearHeaderFooterParagraphBorders(headerParas);
          applyHeaderFooterLine(getHeaderFooterBorderParagraph(header, true), -3, 1, 4);
        }

        if (footer && footer.Range) {
          var footerParas = getHeaderFooterEffectiveParagraphs(footer);
          for (var fp = 0; fp < footerParas.length; fp++) {
            try {
              var footerPara = footerParas[fp];
              if (footerPara && footerPara.Range && footerPara.Range.Font) {
                footerPara.Range.Font.NameFarEast = '宋体';
                footerPara.Range.Font.Name = 'Arial';
                footerPara.Range.Font.Size = 9;
              }
            } catch (footerFontErr) {}
          }
          clearHeaderFooterParagraphBorders(footerParas);
          applyHeaderFooterLine(getHeaderFooterBorderParagraph(footer, false), -1, 7, 4);
        }

        result.fixed += 3;
        pushLog('HF-001', s, '页眉页脚已统一为宋体/Arial，小五号');
        pushLog('HF-002', s, '页眉下已设置0.5磅通长细线');
        pushLog('HF-003', s, '页脚上已设置0.5磅通长双细线');
      } catch (sectionErr) {
        console.log('[processHeaderFooterFast] 跳过节 ' + s + '，原因: ' + sectionErr);
      }
    }
  } finally {
    doc.TrackRevisions = originalTrackRevisions;
  }

  console.log('[processHeaderFooterFast] 完成，修复: ' + result.fixed + '，批注: ' + result.commented);
  return result;
}

function processPageSetupFast(doc, mode) {
  var result = { fixed: 0, commented: 0, byRule: {}, revisionLog: [], sectionCount: 0 };
  if (!doc || !doc.Sections) {
    console.log('[processPageSetupFast] 无节对象');
    return result;
  }

  function pushLog(rule, sectionIndex, message) {
    result.byRule[rule] = (result.byRule[rule] || 0) + 1;
    result.revisionLog.push({
      rule: rule,
      sectionIndex: sectionIndex,
      message: message
    });
  }

  var sectionCount = doc.Sections.Count || 0;
  result.sectionCount = sectionCount;
  console.log('[processPageSetupFast] 开始，节数: ' + sectionCount);

  var topBottom = 72;
  var leftRight = 90;
  var headerFooterDistance = 42.5;

  for (var s = 1; s <= sectionCount; s++) {
    try {
      var section = doc.Sections.Item(s);
      if (!section || !section.PageSetup) continue;
      var pageSetup = section.PageSetup;

      if (mode === 'conservative') {
        result.commented += 2;
        pushLog('PG-001', s, '页边距应为上下2.54cm、左右3.17cm');
        pushLog('PG-002', s, '页眉页脚距边界均应为1.5cm');
        continue;
      }

      pageSetup.TopMargin = topBottom;
      pageSetup.BottomMargin = topBottom;
      pageSetup.LeftMargin = leftRight;
      pageSetup.RightMargin = leftRight;
      pageSetup.HeaderDistance = headerFooterDistance;
      pageSetup.FooterDistance = headerFooterDistance;

      result.fixed += 2;
      pushLog('PG-001', s, '页边距已设置为上下2.54cm、左右3.17cm');
      pushLog('PG-002', s, '页眉页脚距边界已设置为1.5cm');
    } catch (sectionErr) {
      console.log('[processPageSetupFast] 跳过节 ' + s + '，原因: ' + sectionErr);
    }
  }

  console.log('[processPageSetupFast] 完成，修复: ' + result.fixed + '，批注: ' + result.commented);
  return result;
}

function processFullProofreadFast(doc, mode) {
  var result = { fixed: 0, commented: 0, byRule: {}, revisionLog: [], paraCount: 0 };
  if (!doc) {
    console.log('[processFullProofreadFast] 无活动文档');
    return result;
  }

  function mergeStage(stageName, stageResult) {
    if (!stageResult) return;
    var stageFixed = stageResult.fixed || stageResult.totalFixed || 0;
    var stageCommented = stageResult.commented || 0;
    result.fixed += stageFixed;
    result.commented += stageCommented;
    if (stageResult.byRule) {
      for (var key in stageResult.byRule) {
        result.byRule[key] = (result.byRule[key] || 0) + stageResult.byRule[key];
      }
    }
    if (stageResult.revisionLog && stageResult.revisionLog.length) {
      result.revisionLog = result.revisionLog.concat(stageResult.revisionLog);
    } else if (stageResult.details && stageResult.details.length) {
      result.revisionLog = result.revisionLog.concat(stageResult.details);
    }
    console.log('[processFullProofreadFast] ' + stageName + ' 完成，修复: ' + stageFixed + '，批注: ' + stageCommented);
  }

  console.log('[processFullProofreadFast] 开始执行完整流水线');

  mergeStage('编号', scanStructureForNumbering(doc, 'numbering'));
  mergeStage('同上同左', processTableContentFast(doc, mode));
  mergeStage('数值格式', processValueFast(doc, mode));
  mergeStage('页面设置', processPageSetupFast(doc, mode));
  mergeStage('标题正文', processFontFast(doc, mode));
  mergeStage('图表排版', processFigureTableLayoutFast(doc, mode, 'figure_table_layout'));
  mergeStage('公式排版', processFormulaLayoutFast(doc, mode));
  mergeStage('页眉页脚', processHeaderFooterFast(doc, mode));

  try {
    result.paraCount = doc.Paragraphs ? doc.Paragraphs.Count : 0;
  } catch (e) {
    result.paraCount = 0;
  }

  console.log('[processFullProofreadFast] 完成，总修复: ' + result.fixed + '，总批注: ' + result.commented);
  return result;
}

function processContentFast(doc, mode) {
  var result = { fixed: 0, commented: 0, byRule: {}, revisionLog: [], paraCount: 0 };
  function merge(stageResult) {
    if (!stageResult) return;
    result.fixed += stageResult.fixed || 0;
    result.commented += stageResult.commented || 0;
    if (stageResult.byRule) {
      for (var key in stageResult.byRule) {
        result.byRule[key] = (result.byRule[key] || 0) + stageResult.byRule[key];
      }
    }
    if (stageResult.revisionLog && stageResult.revisionLog.length) {
      result.revisionLog = result.revisionLog.concat(stageResult.revisionLog);
    }
    if (!result.paraCount && stageResult.paraCount) result.paraCount = stageResult.paraCount;
  }

  merge(processTableContentFast(doc, mode));
  merge(processValueFast(doc, mode));
  merge(processPunctuationFast(doc, mode));
  return result;
}

function processFormatFast(doc, mode) {
  var result = { fixed: 0, commented: 0, byRule: {}, revisionLog: [], paraCount: 0 };
  function merge(stageResult) {
    if (!stageResult) return;
    result.fixed += stageResult.fixed || 0;
    result.commented += stageResult.commented || 0;
    if (stageResult.byRule) {
      for (var key in stageResult.byRule) {
        result.byRule[key] = (result.byRule[key] || 0) + stageResult.byRule[key];
      }
    }
    if (stageResult.revisionLog && stageResult.revisionLog.length) {
      result.revisionLog = result.revisionLog.concat(stageResult.revisionLog);
    }
    if (!result.paraCount && stageResult.paraCount) result.paraCount = stageResult.paraCount;
  }

  merge(processPageSetupFast(doc, mode));
  merge(processFontFast(doc, mode));
  merge(processFigureTableLayoutFast(doc, mode, 'figure_table_layout'));
  merge(processFormulaLayoutFast(doc, mode));
  merge(processHeaderFooterFast(doc, mode));
  return result;
}

function processPunctuationFast(doc, mode) {
  var result = { fixed: 0, commented: 0, byRule: {}, revisionLog: [], paraCount: 0 };
  if (!doc || !doc.Paragraphs) {
    console.log('[processPunctuationFast] 无段落内容');
    return result;
  }

  function cleanText(text) {
    return String(text || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function addRule(rule, index, original, suggested, message) {
    result.byRule[rule] = (result.byRule[rule] || 0) + 1;
    result.revisionLog.push({ rule: rule, index: index, original: original, suggested: suggested, message: message });
  }

  var paraCount = doc.Paragraphs.Count || 0;
  result.paraCount = paraCount;
  console.log('[processPunctuationFast] 开始，段落数: ' + paraCount);

  var docText = doc.Content && doc.Content.Text ? String(doc.Content.Text) : '';
  var logicalParas = docText.split('\r');
  var maxParaCount = Math.min(paraCount, logicalParas.length);
  var originalTrackRevisions = doc.TrackRevisions;
  doc.TrackRevisions = true;

  try {
    for (var i = 1; i <= maxParaCount; i++) {
      try {
        var text = cleanText(logicalParas[i - 1]);
        if (!text) continue;
        var para = doc.Paragraphs.Item(i);
        if (!para || !para.Range) continue;
        var raw = String(para.Range.Text || '');
        var changed = false;

        if (/^[图表]\s*[\dA-Za-z一二三四五六七八九十百千\.]+/.test(text) && /[。！？]$/.test(text)) {
          var replacedCaption = raw.replace(/([。！？])(\r?)$/, '$2');
          if (replacedCaption !== raw && mode !== 'conservative') {
            para.Range.Text = replacedCaption;
            raw = String(para.Range.Text || '');
            changed = true;
            result.fixed++;
          } else if (mode === 'conservative') {
            result.commented++;
          }
          addRule(/^图/.test(text) ? 'M-001' : 'M-002', i, text, text.replace(/[。！？]$/, ''), '图表名称末尾不应有句号等标点');
        }

        var normalized = cleanText(raw);
        if (/\(([^)]*[\u4e00-\u9fff][^)]*)\)/.test(normalized)) {
          var bracketFixed = raw.replace(/\(([^)\r]*[\u4e00-\u9fff][^)\r]*)\)/g, '（$1）');
          if (bracketFixed !== raw && mode !== 'conservative') {
            para.Range.Text = bracketFixed;
            raw = String(para.Range.Text || '');
            changed = true;
            result.fixed++;
          } else if (bracketFixed !== raw && mode === 'conservative') {
            result.commented++;
          }
          addRule('M-005', i, normalized, cleanText(bracketFixed), '中文内容应使用中文括号');
        }
      } catch (paraErr) {}
    }
  } finally {
    doc.TrackRevisions = originalTrackRevisions;
  }

  console.log('[processPunctuationFast] 完成，修复: ' + result.fixed + '，批注: ' + result.commented);
  return result;
}

function processValueFast(doc, mode) {
  var result = { fixed: 0, commented: 0, byRule: {}, revisionLog: [], paraCount: 0 };
  if (!doc || !doc.Content) {
    console.log('[processValueFast] 无文档内容');
    return result;
  }

  function cleanText(text) {
    return String(text || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function normalizeTextForMatch(text) {
    return cleanText(text).replace(/\s+/g, ' ');
  }

  function dedupeCommentKey(issue) {
    return [issue.index, issue.rule, issue.original, issue.suggested || ''].join('||');
  }

  function dedupeOps(ops) {
    var deduped = [];
    var seen = {};
    for (var i = 0; i < ops.length; i++) {
      var key = ops[i].rule + '||' + ops[i].original + '||' + ops[i].suggested;
      if (seen[key]) continue;
      seen[key] = true;
      deduped.push(ops[i]);
    }
    return deduped;
  }

  function collectValueOps(text, idx) {
    var ops = [];
    var m;

    var tempRe = /(\d+(?:\.\d+)?)\s*℃\s*±\s*(\d+(?:\.\d+)?)\s*℃/g;
    while ((m = tempRe.exec(text)) !== null) {
      ops.push({ index: idx, rule: 'V-006', original: m[0], suggested: m[1] + '±' + m[2] + '℃' });
    }

    var decimalRe = /(^|[^0-9])\.(\d+)/g;
    while ((m = decimalRe.exec(text)) !== null) {
      if (/\d\.$/.test(m[1])) continue;
      ops.push({ index: idx, rule: 'V-007', original: m[0], suggested: m[1] + '0.' + m[2] });
    }

    var unitRe1 = /(\d+(?:\.\d+)?)\s*(mm|cm|m|km|N|kN|Pa|kPa|MPa|℃)\s*[-—–]\s*(\d+(?:\.\d+)?)\s*\2/gi;
    while ((m = unitRe1.exec(text)) !== null) {
      ops.push({ index: idx, rule: 'V-009', original: m[0], suggested: m[1] + '～' + m[3] + m[2] });
    }

    var unitRe2 = /(\d+(?:\.\d+)?)\s*[-—–]\s*(\d+(?:\.\d+)?)\s*(mm|cm|m|km|N|kN|Pa|kPa|MPa|℃)/gi;
    while ((m = unitRe2.exec(text)) !== null) {
      ops.push({ index: idx, rule: 'V-009', original: m[0], suggested: m[1] + '～' + m[2] + m[3] });
    }

    var percentRe = /(\d+(?:\.\d+)?)\s*～\s*(\d+(?:\.\d+)?)\s*%/g;
    while ((m = percentRe.exec(text)) !== null) {
      ops.push({ index: idx, rule: 'V-010', original: m[0], suggested: m[1] + '%～' + m[2] + '%' });
    }

    var sizeRe = /(\d+(?:\.\d+)?)\s*[×xX]\s*(\d+(?:\.\d+)?)\s*[×xX]\s*(\d+(?:\.\d+)?)\s*(mm|cm|m)/gi;
    while ((m = sizeRe.exec(text)) !== null) {
      ops.push({ index: idx, rule: 'V-012', original: m[0], suggested: m[1] + m[4] + '×' + m[2] + m[4] + '×' + m[3] + m[4] });
    }

    return dedupeOps(ops);
  }

  function collectValueComments(text, idx, comments, commentSeen) {
    var m;
    var powerRe = /(\d+(?:\.\d+)?)\s*～\s*(\d+(?:\.\d+)?)\s*[×x]\s*10\^?\d*/gi;
    while ((m = powerRe.exec(text)) !== null) {
      var powerIssue = { index: idx, rule: 'V-011', name: '幂次范围', original: m[0], message: '幂次范围应在每个数值后都写出幂次，建议人工修改', autoFix: false };
      var powerKey = dedupeCommentKey(powerIssue);
      if (!commentSeen[powerKey]) {
        commentSeen[powerKey] = true;
        comments.push(powerIssue);
      }
    }

    var chineseNumMap = { '零': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10, '两': 2, '半': 0.5 };
    function convertChineseToArabic(chinese) {
      if (!chinese) return null;
      if (chineseNumMap[chinese] !== undefined) return chineseNumMap[chinese];
      if (chinese.indexOf('点') !== -1) {
        var parts = chinese.split('点');
        if (parts.length === 2) {
          var intPart = convertChineseToArabic(parts[0]);
          var decPart = convertChineseToArabic(parts[1]);
          if (intPart !== null && decPart !== null) {
            return intPart + decPart / Math.pow(10, String(decPart).length);
          }
        }
        return null;
      }
      var total = 0;
      var temp = 0;
      for (var ci = 0; ci < chinese.length; ci++) {
        var ch = chinese.charAt(ci);
        if (ch === '十') {
          temp = temp === 0 ? 10 : temp * 10;
          total += temp;
          temp = 0;
        } else if (ch === '百') {
          temp = temp === 0 ? 100 : temp * 100;
          total += temp;
          temp = 0;
        } else if (ch === '千') {
          temp = temp === 0 ? 1000 : temp * 1000;
          total += temp;
          temp = 0;
        } else if (chineseNumMap[ch] !== undefined) {
          temp = chineseNumMap[ch];
        } else {
          return null;
        }
      }
      total += temp;
      return total > 0 ? total : null;
    }

    var percentRe = /百分之([零一二三四五六七八九十百千万]+)/g;
    while ((m = percentRe.exec(text)) !== null) {
      var percentValue = convertChineseToArabic(m[1]);
      if (percentValue !== null) {
        var percentIssue = { index: idx, rule: 'V-013', name: '中文数值书写', original: m[0], suggested: percentValue + '%', message: '百分数应采用数学记号', autoFix: false };
        var percentKey = dedupeCommentKey(percentIssue);
        if (!commentSeen[percentKey]) {
          commentSeen[percentKey] = true;
          comments.push(percentIssue);
        }
      }
    }

    var fractionRe = /([零一二三四五六七八九十百千万]+)分之([零一二三四五六七八九十百千万]+)/g;
    while ((m = fractionRe.exec(text)) !== null) {
      var denominator = convertChineseToArabic(m[1]);
      var numerator = convertChineseToArabic(m[2]);
      if (denominator !== null && numerator !== null) {
        var fractionIssue = { index: idx, rule: 'V-013', name: '中文数值书写', original: m[0], suggested: numerator + '/' + denominator, message: '分数应采用数学记号', autoFix: false };
        var fractionKey = dedupeCommentKey(fractionIssue);
        if (!commentSeen[fractionKey]) {
          commentSeen[fractionKey] = true;
          comments.push(fractionIssue);
        }
      }
    }

    var ratioRe = /([零一二三四五六七八九十百千万点]+)比([零一二三四五六七八九十百千万点]+)/g;
    while ((m = ratioRe.exec(text)) !== null) {
      var leftValue = convertChineseToArabic(m[1]);
      var rightValue = convertChineseToArabic(m[2]);
      if (leftValue !== null && rightValue !== null) {
        var ratioIssue = { index: idx, rule: 'V-013', name: '中文数值书写', original: m[0], suggested: leftValue + ':' + rightValue, message: '比例数应采用数学记号', autoFix: false };
        var ratioKey = dedupeCommentKey(ratioIssue);
        if (!commentSeen[ratioKey]) {
          commentSeen[ratioKey] = true;
          comments.push(ratioIssue);
        }
      }
    }
  }

  var ops = [];
  var comments = [];
  var commentSeen = {};
  var docText = doc.Content.Text ? String(doc.Content.Text) : '';
  var paras = docText.split('\r');
  var paragraphTextCache = {};
  result.paraCount = paras.length;
  console.log('[processValueFast] 开始，段落数: ' + result.paraCount);

  function getParagraphTextByIndex(index) {
    if (paragraphTextCache[index] !== undefined) return paragraphTextCache[index];
    try {
      var para = doc.Paragraphs.Item(index);
      paragraphTextCache[index] = para && para.Range ? normalizeTextForMatch(para.Range.Text || '') : '';
    } catch (e) {
      paragraphTextCache[index] = '';
    }
    return paragraphTextCache[index];
  }

  function resolveParagraphIndex(logicalIndex, expectedText) {
    var normalizedExpected = normalizeTextForMatch(expectedText);
    if (!normalizedExpected) return 0;
    var offsets = [0, -1, 1, -2, 2, -3, 3, -5, 5, -8, 8];
    for (var oi = 0; oi < offsets.length; oi++) {
      var candidate = logicalIndex + offsets[oi];
      if (candidate < 1 || candidate > result.paraCount) continue;
      if (getParagraphTextByIndex(candidate) === normalizedExpected) return candidate;
    }
    return 0;
  }

  for (var i = 0; i < paras.length; i++) {
    var paragraphText = cleanText(paras[i]);
    if (!paragraphText) continue;
    var paraOps = collectValueOps(paragraphText, i + 1);
    if (paraOps.length) ops = ops.concat(paraOps);
    collectValueComments(paragraphText, i + 1, comments, commentSeen);
  }

  var originalTrackRevisions = doc.TrackRevisions;
  doc.TrackRevisions = true;
  try {
    for (i = 0; i < ops.length; i++) {
      var op = ops[i];
      try {
        var resolvedIndex = resolveParagraphIndex(op.index, paras[op.index - 1] || '');
        if (!resolvedIndex) continue;
        var targetPara = doc.Paragraphs.Item(resolvedIndex);
        if (!targetPara || !targetPara.Range) continue;

        var paraRange = targetPara.Range;
        var rawParaText = String(paraRange.Text || '');
        var searchText = /\r$/.test(rawParaText) ? rawParaText.slice(0, -1) : rawParaText;
        var matchOffset = searchText.indexOf(op.original);
        if (matchOffset < 0) continue;

        var exactRange = doc.Range(paraRange.Start + matchOffset, paraRange.Start + matchOffset + op.original.length);
        if (!exactRange || exactRange.Start === exactRange.End) continue;

        exactRange.Text = op.suggested;
        paragraphTextCache[resolvedIndex] = undefined;
        result.fixed++;
        result.byRule[op.rule] = (result.byRule[op.rule] || 0) + 1;
        result.revisionLog.push({ rule: op.rule, index: resolvedIndex, original: op.original, suggested: op.suggested });
      } catch (replaceErr) {}
    }

    for (i = 0; i < comments.length; i++) {
      var issue = comments[i];
      if (mode === 'conservative' || mode === 'standard' || mode === 'aggressive') {
        try {
          var resolvedCommentIndex = resolveParagraphIndex(issue.index, paras[issue.index - 1] || '');
          if (!resolvedCommentIndex) continue;
          var para = doc.Paragraphs.Item(resolvedCommentIndex);
          if (!para || !para.Range) continue;
          issue.index = resolvedCommentIndex;
          addCommentToDoc(doc, para.Range, issue);
          result.commented++;
          result.byRule[issue.rule] = (result.byRule[issue.rule] || 0) + 1;
          result.revisionLog.push(issue);
        } catch (commentErr) {}
      }
    }
  } finally {
    doc.TrackRevisions = originalTrackRevisions;
  }

  console.log('[processValueFast] 完成，修复: ' + result.fixed + '，批注: ' + result.commented);
  return result;
}

// ========== 内部函数：表编号检查 ==========

function checkTableNumbering(structure) {
  var issues = [];
  var paragraphs = structure.paragraphs || [];

  // 中文数字映射
  var chineseNumMap = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
                        '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
                        '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15 };
  var chineseToLetter = { '一': 'A', '二': 'B', '三': 'C', '四': 'D', '五': 'E',
                          '六': 'F', '七': 'G', '八': 'H', '九': 'I', '十': 'J' };

  // 按章节跟踪表格编号
  var currentChapter = 0;
  var currentSection = 0;
  var tableCounters = {};  // key: "chapter.section" -> count
  var inAppendix = false;  // 是否在附录区域
  var currentAppendix = '';  // 当前附录编号 (A, B, C...)
  var appendixTableCounter = 0;  // 附录内表格计数
  var lastMainChapter = 0;  // 最后一个正文章节号

  console.log('[checkTableNumbering] 开始检查，共 ' + paragraphs.length + ' 个段落');

  // 输出前10个段落的摘要，用于调试
  for (var debugIdx = 0; debugIdx < Math.min(10, paragraphs.length); debugIdx++) {
    var debugP = paragraphs[debugIdx];
    console.log('[checkTableNumbering] 段落' + debugIdx + ': ' + (debugP.fullText || debugP.text || '').substring(0, 40));
  }

  for (var i = 0; i < paragraphs.length; i++) {
    var p = paragraphs[i];
    var text = p.fullText || p.text || '';

    // 检测是否进入附录
    var appMatch = text.match(/^附录\s*([A-Z一二三四五六七八九十])/i);
    if (appMatch) {
      inAppendix = true;
      var letter = appMatch[1].toUpperCase();
      currentAppendix = chineseToLetter[letter] || letter;
      appendixTableCounter = 0;
      console.log('[checkTableNumbering] 进入附录: ' + currentAppendix);
      continue;
    }

    // 检测中文章节标题
    var cnCh = text.match(/^第([一二三四五六七八九十百]+)章/);
    if (cnCh) {
      if (inAppendix) {
        // 中文章节标题意味着离开附录
        inAppendix = false;
        currentAppendix = '';
        console.log('[checkTableNumbering] 离开附录，遇到中文章节');
      }
      currentChapter = chineseNumMap[cnCh[1]] || 0;
      lastMainChapter = currentChapter;
      currentSection = 0;
      tableCounters = {};
      console.log('[checkTableNumbering] 中文章节: ' + currentChapter);
      continue;
    }

    // 检测阿拉伯数字一级标题
    var arCh = text.match(/^(\d+)\s+(?!\.)/);
    if (arCh) {
      var arNum = parseInt(arCh[1], 10);
      if (!inAppendix) {
        // 正文中：更新章节
        currentChapter = arNum;
        lastMainChapter = arNum;
        currentSection = 0;
        tableCounters = {};
        console.log('[checkTableNumbering] 阿拉伯章节: ' + currentChapter);
      } else {
        // 附录中：检查是否是正文章节的延续
        if (lastMainChapter > 0 && arNum === lastMainChapter + 1) {
          inAppendix = false;
          currentAppendix = '';
          currentChapter = arNum;
          lastMainChapter = arNum;
          currentSection = 0;
          tableCounters = {};
          console.log('[checkTableNumbering] 从附录返回正文，章节: ' + currentChapter);
        }
        // 否则保持在附录状态
      }
      continue;
    }

    // 更新当前节
    var l2Match = text.match(/^(\d+)\.(\d+)\s/);
    if (l2Match) {
      var l2Chapter = parseInt(l2Match[1], 10);
      if (l2Chapter === currentChapter) {
        var newSection = parseInt(l2Match[2], 10);
        // 如果节号变化，重置该章节的表格计数器
        if (newSection !== currentSection) {
          currentSection = newSection;
          // 重置当前节的表格计数器
          var key = currentChapter + '.' + currentSection;
          tableCounters[key] = 0;
        }
      }
    }

    // 检测表格编号
    // 格式1: "表一" 或 "表一."（中文数字表格编号，可能有句号）
    var tableCnMatch = text.match(/^表([一二三四五六七八九十]+)[.。\s]+(.*)$/);
    // 格式2: "表 1.2-1" 或 "表 1-1"（阿拉伯数字格式，规范格式）
    var tableArMatch = text.match(/^表\s*(\d+)\.(\d+)[-—–](\d+)\s+(.*)$/);
    // 格式3: "表 X-Y"（简单章节-序号格式）
    var tableSimpleChapterMatch = text.match(/^表\s*(\d+)[-—–](\d+)\s+(.*)$/);
    // 格式4: "表 X 表名"（全局顺序号格式，需要转换为规范格式）
    var tableSimpleMatch = text.match(/^表\s*(\d+)\s+(.*)$/);

    // 调试：检测到表格
    if (tableCnMatch || tableArMatch || tableSimpleMatch) {
      console.log('[checkTableNumbering] 检测到表格: "' + text.substring(0, 50) + '"' +
        ', inAppendix=' + inAppendix + ', currentChapter=' + currentChapter + ', currentSection=' + currentSection +
        ', tableCnMatch=' + !!tableCnMatch + ', tableArMatch=' + !!tableArMatch + ', tableSimpleMatch=' + !!tableSimpleMatch);
    }

    if (tableCnMatch) {
      // 中文数字表格编号 - 需要根据当前章节/节生成正确编号
      var tableName = tableCnMatch[2];
      var suggestedText;

      if (inAppendix && currentAppendix) {
        // 附录中：表A1、表A2 格式
        appendixTableCounter++;
        suggestedText = '表' + currentAppendix + appendixTableCounter + '  ' + tableName;
      } else {
        // 正文中：使用正确的计数器
        var key = currentChapter + '.' + currentSection;
        tableCounters[key] = (tableCounters[key] || 0) + 1;
        var tableNum = tableCounters[key];

        if (currentSection > 0) {
          suggestedText = '表 ' + currentChapter + '.' + currentSection + '-' + tableNum + '  ' + tableName;
        } else {
          suggestedText = '表 ' + currentChapter + '.' + (currentSection || 1) + '-' + tableNum + '  ' + tableName;
        }
      }

      issues.push({
        index: p.index,
        rule: 'T-001',
        name: '表编号格式不规范',
        original: text,
        suggested: suggestedText,
        message: inAppendix ? '附录中表编号应为 表' + currentAppendix + 'N 格式' : '表编号格式不规范',
        autoFix: true
      });
    } else if (tableArMatch) {
      // 规范格式：表X.Y-Z
      var numChapter = parseInt(tableArMatch[1], 10);
      var numSection = parseInt(tableArMatch[2], 10);
      var numTable = parseInt(tableArMatch[3], 10);
      var tableName = tableArMatch[4];

      // 检查章节号是否正确
      if (inAppendix && currentAppendix) {
        // 附录中：应该改为 表A1 格式
        appendixTableCounter++;
        var suggestedText = '表' + currentAppendix + appendixTableCounter + '  ' + tableName;

        issues.push({
          index: p.index,
          rule: 'T-001',
          name: '附录表编号格式错误',
          original: text,
          suggested: suggestedText,
          message: '附录' + currentAppendix + '中的表编号应为 表' + currentAppendix + 'N 格式',
          autoFix: true
        });
      } else if (numChapter !== currentChapter || numSection !== (currentSection || 1)) {
        // 正文中：章节号或节号错误
        var key = currentChapter + '.' + (currentSection || 1);
        tableCounters[key] = (tableCounters[key] || 0) + 1;
        var newTableNum = tableCounters[key];

        var suggestedText = '表 ' + currentChapter + '.' + (currentSection || 1) + '-' + newTableNum + '  ' + tableName;

        issues.push({
          index: p.index,
          rule: 'T-001',
          name: '表编号章节号错误',
          original: text,
          suggested: suggestedText,
          message: '表编号章节号 ' + numChapter + '.' + numSection + ' → ' + currentChapter + '.' + (currentSection || 1),
          autoFix: true
        });
      }
    } else if (tableSimpleMatch && !tableSimpleChapterMatch) {
      // 简单格式：表X 表名 - 需要转换为规范格式 表章.节-序号
      var tableName = tableSimpleMatch[2];
      console.log('[checkTableNumbering] 简单格式表格检测: currentChapter=' + currentChapter + ', tableName=' + tableName);

      if (inAppendix && currentAppendix) {
        // 附录中：表A1、表A2 格式
        appendixTableCounter++;
        var suggestedText = '表' + currentAppendix + appendixTableCounter + '  ' + tableName;

        issues.push({
          index: p.index,
          rule: 'T-001',
          name: '表编号需转换为规范格式',
          original: text,
          suggested: suggestedText,
          message: '表编号应转换为规范格式',
          autoFix: true
        });
      } else if (currentChapter > 0) {
        // 正文中：转换为 表章.节-序号
        var key = currentChapter + '.' + (currentSection || 1);
        tableCounters[key] = (tableCounters[key] || 0) + 1;
        var tableNum = tableCounters[key];

        var suggestedText = '表 ' + currentChapter + '.' + (currentSection || 1) + '-' + tableNum + '  ' + tableName;

        issues.push({
          index: p.index,
          rule: 'T-001',
          name: '表编号需转换为规范格式',
          original: text,
          suggested: suggestedText,
          message: '表编号应转换为规范格式 表' + currentChapter + '.' + (currentSection || 1) + '-N',
          autoFix: true
        });
      }
    }
  }

  return issues;
}

// ========== 内部函数：图编号检查 ==========

function checkFigureNumbering(structure) {
  var issues = [];
  var paragraphs = structure.paragraphs || [];

  // 中文数字映射
  var chineseNumMap = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
                        '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
                        '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15 };
  var chineseToLetter = { '一': 'A', '二': 'B', '三': 'C', '四': 'D', '五': 'E',
                          '六': 'F', '七': 'G', '八': 'H', '九': 'I', '十': 'J' };

  // 按章节跟踪图编号
  var currentChapter = 0;
  var currentSection = 0;
  var figCounters = {};  // key: "chapter.section" -> count
  var inAppendix = false;  // 是否在附录区域
  var currentAppendix = '';  // 当前附录编号 (A, B, C...)
  var appendixFigCounter = 0;  // 附录内图计数
  var lastMainChapter = 0;  // 最后一个正文章节号

  console.log('[checkFigureNumbering] 开始检查，共 ' + paragraphs.length + ' 个段落');

  for (var i = 0; i < paragraphs.length; i++) {
    var p = paragraphs[i];
    var text = p.fullText || p.text || '';

    // 检测是否进入附录
    var appMatch = text.match(/^附录\s*([A-Z一二三四五六七八九十])/i);
    if (appMatch) {
      inAppendix = true;
      var letter = appMatch[1].toUpperCase();
      currentAppendix = chineseToLetter[letter] || letter;
      appendixFigCounter = 0;
      console.log('[checkFigureNumbering] 进入附录: ' + currentAppendix);
      continue;
    }

    // 检测中文章节标题
    var cnCh = text.match(/^第([一二三四五六七八九十百]+)章/);
    if (cnCh) {
      if (inAppendix) {
        inAppendix = false;
        currentAppendix = '';
        console.log('[checkFigureNumbering] 离开附录，遇到中文章节');
      }
      currentChapter = chineseNumMap[cnCh[1]] || 0;
      lastMainChapter = currentChapter;
      currentSection = 0;
      figCounters = {};
      console.log('[checkFigureNumbering] 中文章节: ' + currentChapter);
      continue;
    }

    var arCh = text.match(/^(\d+)\s+(?!\.)/);
    if (arCh) {
      var arNum = parseInt(arCh[1], 10);
      if (!inAppendix) {
        currentChapter = arNum;
        lastMainChapter = arNum;
        currentSection = 0;
        figCounters = {};
        console.log('[checkFigureNumbering] 阿拉伯章节: ' + currentChapter);
      } else {
        if (lastMainChapter > 0 && arNum === lastMainChapter + 1) {
          inAppendix = false;
          currentAppendix = '';
          currentChapter = arNum;
          lastMainChapter = arNum;
          currentSection = 0;
          figCounters = {};
          console.log('[checkFigureNumbering] 从附录返回正文，章节: ' + currentChapter);
        }
      }
      continue;
    }

    // 更新当前节
    var l2Match = text.match(/^(\d+)\.(\d+)\s/);
    if (l2Match) {
      var l2Chapter = parseInt(l2Match[1], 10);
      if (l2Chapter === currentChapter) {
        var newSection = parseInt(l2Match[2], 10);
        // 如果节号变化，重置该章节的图计数器
        if (newSection !== currentSection) {
          currentSection = newSection;
          // 重置当前节的图计数器
          var key = currentChapter + '.' + currentSection;
          figCounters[key] = 0;
        }
      }
    }

    // 检测图编号
    // 格式1: "图一" 或 "图一."（中文数字图编号，可能有句号）
    var figCnMatch = text.match(/^图([一二三四五六七八九十]+)[.。\s]+(.*)$/);
    // 格式2: "图 1.2-1"（阿拉伯数字格式，规范格式）
    var figArMatch = text.match(/^图\s*(\d+)\.(\d+)[-—–](\d+)\s+(.*)$/);
    // 格式3: "图 X-Y"（简单章节-序号格式）
    var figSimpleChapterMatch = text.match(/^图\s*(\d+)[-—–](\d+)\s+(.*)$/);
    // 格式4: "图 X 图名"（全局顺序号格式，需要转换为规范格式）
    var figSimpleMatch = text.match(/^图\s*(\d+)\s+(.*)$/);

    // 调试：检测到图
    if (figCnMatch || figArMatch || figSimpleMatch) {
      console.log('[checkFigureNumbering] 检测到图: "' + text.substring(0, 50) + '"' +
        ', inAppendix=' + inAppendix + ', currentChapter=' + currentChapter + ', currentSection=' + currentSection +
        ', figCnMatch=' + !!figCnMatch + ', figArMatch=' + !!figArMatch + ', figSimpleMatch=' + !!figSimpleMatch);
    }

    if (figCnMatch) {
      // 中文数字图编号 - 需要根据当前章节/节生成正确编号
      var figName = figCnMatch[2];
      var suggestedText;

      if (inAppendix && currentAppendix) {
        // 附录中：图A1、图A2 格式
        appendixFigCounter++;
        suggestedText = '图' + currentAppendix + appendixFigCounter + '  ' + figName;
      } else {
        // 正文中：使用正确的计数器
        var key = currentChapter + '.' + (currentSection || 1);
        figCounters[key] = (figCounters[key] || 0) + 1;
        var figNum = figCounters[key];

        suggestedText = '图 ' + currentChapter + '.' + (currentSection || 1) + '-' + figNum + '  ' + figName;
      }

      issues.push({
        index: p.index,
        rule: 'G-001',
        name: '图编号格式不规范',
        original: text,
        suggested: suggestedText,
        message: inAppendix ? '附录中图编号应为 图' + currentAppendix + 'N 格式' : '图编号格式不规范',
        autoFix: true
      });
    } else if (figArMatch) {
      // 规范格式：图X.Y-Z
      var numChapter = parseInt(figArMatch[1], 10);
      var numSection = parseInt(figArMatch[2], 10);
      var numFig = parseInt(figArMatch[3], 10);
      var figName = figArMatch[4];

      // 检查章节号是否正确
      if (inAppendix && currentAppendix) {
        // 附录中：应该改为 图A1 格式
        appendixFigCounter++;
        var suggestedText = '图' + currentAppendix + appendixFigCounter + '  ' + figName;

        issues.push({
          index: p.index,
          rule: 'G-001',
          name: '附录图编号格式错误',
          original: text,
          suggested: suggestedText,
          message: '附录' + currentAppendix + '中的图编号应为 图' + currentAppendix + 'N 格式',
          autoFix: true
        });
      } else if (numChapter !== currentChapter || numSection !== (currentSection || 1)) {
        // 正文中：章节号或节号错误
        var key = currentChapter + '.' + (currentSection || 1);
        figCounters[key] = (figCounters[key] || 0) + 1;
        var newFigNum = figCounters[key];

        var suggestedText = '图 ' + currentChapter + '.' + (currentSection || 1) + '-' + newFigNum + '  ' + figName;

        issues.push({
          index: p.index,
          rule: 'G-001',
          name: '图编号章节号错误',
          original: text,
          suggested: suggestedText,
          message: '图编号章节号 ' + numChapter + '.' + numSection + ' → ' + currentChapter + '.' + (currentSection || 1),
          autoFix: true
        });
      }
    } else if (figSimpleMatch && !figSimpleChapterMatch) {
      // 简单格式：图X 图名 - 需要转换为规范格式 图章.节-序号
      var figName = figSimpleMatch[2];
      console.log('[checkFigureNumbering] 简单格式图检测: currentChapter=' + currentChapter + ', figName=' + figName);

      if (inAppendix && currentAppendix) {
        // 附录中：图A1、图A2 格式
        appendixFigCounter++;
        var suggestedText = '图' + currentAppendix + appendixFigCounter + '  ' + figName;

        issues.push({
          index: p.index,
          rule: 'G-001',
          name: '图编号需转换为规范格式',
          original: text,
          suggested: suggestedText,
          message: '图编号应转换为规范格式',
          autoFix: true
        });
      } else if (currentChapter > 0) {
        // 正文中：转换为 图章.节-序号
        var key = currentChapter + '.' + (currentSection || 1);
        figCounters[key] = (figCounters[key] || 0) + 1;
        var figNum = figCounters[key];

        var suggestedText = '图 ' + currentChapter + '.' + (currentSection || 1) + '-' + figNum + '  ' + figName;

        issues.push({
          index: p.index,
          rule: 'G-001',
          name: '图编号需转换为规范格式',
          original: text,
          suggested: suggestedText,
          message: '图编号应转换为规范格式 图' + currentChapter + '.' + (currentSection || 1) + '-N',
          autoFix: true
        });
      }
    }
  }

  return issues;
}

// ========== 内部函数：公式编号检查 ==========
/**
 * 公式编号规范：
 * - 格式：(X.Y.Z-N) - 章号.节号.条号-顺序号
 * - 示例：(3.2.1-1) 表示第三章第二节第一条的第一个公式
 * - 当同一条内有多个公式时，加顺序号，如 (3.2.1-1) 和 (3.2.1-2)
 * - 附录公式：附录中公式编号方法与正文相似，或直接用附录号如 (1)
 */

function checkFormulaNumbering(structure) {
  var issues = [];
  var paragraphs = structure.paragraphs || [];

  // 中文数字转阿拉伯数字
  var chineseToNumber = function(cn) {
    var map = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
                '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
                '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15 };
    return map[cn] || 0;
  };

  var chineseToLetter = { '一': 'A', '二': 'B', '三': 'C', '四': 'D', '五': 'E',
                          '六': 'F', '七': 'G', '八': 'H', '九': 'I', '十': 'J' };

  // 跟踪当前章节、节、条
  var currentChapter = 0;      // 章号
  var currentSection = 0;      // 节号
  var currentSubsection = 0;   // 条号
  var inAppendix = false;      // 是否在附录区域
  var currentAppendix = '';    // 当前附录编号 (A, B, C...)
  var appendixFormulaCounter = 0;  // 附录内公式计数
  var lastMainChapter = 0;     // 最后一个正文章节号

  // 公式计数器（用于生成建议编号）
  var formulaCounters = {};

  console.log('[checkFormulaNumbering] 开始检查，共 ' + paragraphs.length + ' 个段落');

  for (var i = 0; i < paragraphs.length; i++) {
    var p = paragraphs[i];
    var text = p.fullText || p.text || '';

    // 检测是否进入附录
    var appMatch = text.match(/^附录\s*([A-Z一二三四五六七八九十])/i);
    if (appMatch) {
      inAppendix = true;
      var letter = appMatch[1].toUpperCase();
      currentAppendix = chineseToLetter[letter] || letter;
      appendixFormulaCounter = 0;
      console.log('[checkFormulaNumbering] 进入附录: ' + currentAppendix);
      continue;
    }

    // 检测中文章节编号（第一章、第二章...）
    var cnChapterMatch = text.match(/^第([一二三四五六七八九十百]+)章/);
    if (cnChapterMatch) {
      if (inAppendix) {
        inAppendix = false;
        currentAppendix = '';
        console.log('[checkFormulaNumbering] 离开附录，遇到中文章节');
      }
      currentChapter = chineseToNumber(cnChapterMatch[1]);
      lastMainChapter = currentChapter;
      currentSection = 0;
      currentSubsection = 0;
      console.log('[checkFormulaNumbering] 中文章节: ' + currentChapter);
      continue;
    }

    // 检测阿拉伯数字一级标题（1 xxx，但不是 1.1 xxx）
    var arChapterMatch = text.match(/^(\d+)\s+(?!\.)/);
    if (arChapterMatch) {
      var arNum = parseInt(arChapterMatch[1], 10);
      if (!inAppendix) {
        currentChapter = arNum;
        lastMainChapter = arNum;
        currentSection = 0;
        currentSubsection = 0;
        console.log('[checkFormulaNumbering] 阿拉伯章节: ' + currentChapter);
      } else {
        if (lastMainChapter > 0 && arNum === lastMainChapter + 1) {
          inAppendix = false;
          currentAppendix = '';
          currentChapter = arNum;
          lastMainChapter = arNum;
          currentSection = 0;
          currentSubsection = 0;
          console.log('[checkFormulaNumbering] 从附录返回正文，章节: ' + currentChapter);
        }
      }
      continue;
    }

    // 检测二级标题（X.X xxx）- 更新节号
    var sectionMatch = text.match(/^(\d+)\.(\d+)\s/);
    if (sectionMatch) {
      var sectionChapter = parseInt(sectionMatch[1], 10);
      var sectionNum = parseInt(sectionMatch[2], 10);
      console.log('[checkFormulaNumbering] 检测到二级标题: ' + text.substring(0, 30) +
        ', sectionChapter=' + sectionChapter + ', currentChapter=' + currentChapter);
      if (sectionChapter === currentChapter) {
        currentSection = sectionNum;
        currentSubsection = 0;
        console.log('[checkFormulaNumbering] 更新 currentSection=' + currentSection);
      } else {
        console.log('[checkFormulaNumbering] 二级标题章节号不匹配，不更新节号');
      }
      continue;
    }

    // 检测三级标题（X.X.X xxx）- 更新条号
    var subsectionMatch = text.match(/^(\d+)\.(\d+)\.(\d+)\s/);
    if (subsectionMatch) {
      var subChapter = parseInt(subsectionMatch[1], 10);
      var subSection = parseInt(subsectionMatch[2], 10);
      if (subChapter === currentChapter && subSection === currentSection) {
        currentSubsection = parseInt(subsectionMatch[3], 10);
      }
      continue;
    }

    // 检测公式编号
    // 格式1: (3.2.1-1) - 完整格式：章.节.条-顺序号
    var formulaMatch3 = text.match(/\((\d+)\.(\d+)\.(\d+)[-—–](\d+)\)\s*$/);
    // 格式2: (3.2-1) - 简化格式：章.节-顺序号
    var formulaMatch2 = text.match(/\((\d+)\.(\d+)[-—–](\d+)\)\s*$/);
    // 格式3: (3-1) - 最简格式：章-顺序号
    var formulaMatch1 = text.match(/\((\d+)[-—–](\d+)\)\s*$/);
    // 格式4: 附录中简单编号 (1)、(2)
    var formulaMatchSimple = text.match(/\((\d+)\)\s*$/);
    // 格式5: "公式1" - 不规范格式（无括号）
    var formulaMatchNoParen = text.match(/公式(\d+)\s*$/);

    // 调试：检测到可能的公式
    if (formulaMatch3 || formulaMatch2 || formulaMatch1 || formulaMatchSimple || formulaMatchNoParen) {
      console.log('[checkFormulaNumbering] 检测到公式编号: ' + text.substring(0, 50) +
        ', inAppendix=' + inAppendix + ', currentChapter=' + currentChapter);
    }

    // 处理不规范格式 "公式1"
    if (formulaMatchNoParen && !formulaMatch3 && !formulaMatch2 && !formulaMatch1 && !formulaMatchSimple) {
      var formulaNum = parseInt(formulaMatchNoParen[1], 10);
      // 【修复】统一使用 currentChapter.currentSection 作为计数器 key
      var key = currentChapter + '.' + currentSection;
      formulaCounters[key] = (formulaCounters[key] || 0) + 1;
      var expectedNum = formulaCounters[key];

      var suggestedText = text.replace(/公式(\d+)\s*$/, '(' + currentChapter + '.' + currentSection + '-' + expectedNum + ')');

      console.log('[checkFormulaNumbering] 公式1格式: 段落' + p.index + ', key=' + key + ', expectedNum=' + expectedNum);

      issues.push({
        index: p.index,
        rule: 'E-001',
        name: '公式编号格式不规范',
        original: text,
        suggested: suggestedText,
        message: '公式编号应使用括号格式如(章.节-序号)，当前格式不规范',
        autoFix: true
      });
      continue;
    }

    // 在附录区域，检测到公式编号
    if (inAppendix && currentAppendix) {
      if (formulaMatch1) {
        // 附录中已有格式 (X-N)，应该改为 (AN) 格式
        appendixFormulaCounter++;
        var suggestedText = text.replace(
          /\((\d+)[-—–](\d+)\)\s*$/,
          '(' + currentAppendix + appendixFormulaCounter + ')'
        );

        issues.push({
          index: p.index,
          rule: 'E-001',
          name: '附录公式编号格式错误',
          original: text,
          suggested: suggestedText,
          message: '附录' + currentAppendix + '中的公式编号应为 (' + currentAppendix + 'N) 格式',
          autoFix: true
        });
      } else if (formulaMatchSimple && !formulaMatch3 && !formulaMatch2) {
        // 附录中简单编号 (1)、(2)，需要加附录号
        appendixFormulaCounter++;
        var suggestedText = text.replace(
          /\((\d+)\)\s*$/,
          '(' + currentAppendix + appendixFormulaCounter + ')'
        );

        issues.push({
          index: p.index,
          rule: 'E-001',
          name: '附录公式编号格式不规范',
          original: text,
          suggested: suggestedText,
          message: '附录' + currentAppendix + '中的公式编号建议改为 (' + currentAppendix + 'N) 格式',
          autoFix: true
        });
      } else if (!formulaMatch3 && !formulaMatch2 && !formulaMatch1 && !formulaMatchSimple && !formulaMatchNoParen) {
        // 附录中没有编号的公式（检测是否像公式）
        // 公式特征：包含 = 且长度较短，不包含中文（排除普通文本）
        if (text.indexOf('=') > 0 && text.length < 50 && !/[\u4e00-\u9fa5]/.test(text)) {
          appendixFormulaCounter++;
          issues.push({
            index: p.index,
            rule: 'E-001',
            name: '附录公式缺少编号',
            original: text,
            suggested: text + '      (' + currentAppendix + appendixFormulaCounter + ')',
            message: '附录' + currentAppendix + '中的公式缺少编号，建议添加: (' + currentAppendix + appendixFormulaCounter + ')',
            autoFix: true
          });
        }
      }
      continue;
    }

    // 正文中公式编号检查
    if (formulaMatch3) {
      // 完整格式 (X.Y.Z-N)
      var numChapter = parseInt(formulaMatch3[1], 10);
      var numSection = parseInt(formulaMatch3[2], 10);
      var numSubsection = parseInt(formulaMatch3[3], 10);
      var numFormula = parseInt(formulaMatch3[4], 10);

      // 更新计数器
      var key = currentChapter + '.' + currentSection + '.' + currentSubsection;
      formulaCounters[key] = (formulaCounters[key] || 0) + 1;
      var expectedNum = formulaCounters[key];

      // 检查章节号或编号是否正确
      if (numChapter !== currentChapter || numSection !== currentSection ||
          numSubsection !== currentSubsection || numFormula !== expectedNum) {
        var suggestedText = text.replace(
          new RegExp('\\(' + numChapter + '\\.' + numSection + '\\.' + numSubsection + '[-—–]' + numFormula + '\\)\\s*$'),
          '(' + currentChapter + '.' + currentSection + '.' + currentSubsection + '-' + expectedNum + ')'
        );

        issues.push({
          index: p.index,
          rule: 'E-001',
          name: '公式编号错误',
          original: text,
          suggested: suggestedText,
          message: '公式编号 ' + numChapter + '.' + numSection + '.' + numSubsection + '-' + numFormula + ' → ' + currentChapter + '.' + currentSection + '.' + currentSubsection + '-' + expectedNum,
          autoFix: true
        });
      }
    } else if (formulaMatch2) {
      // 简化格式 (X.Y-N)
      var numChapter2 = parseInt(formulaMatch2[1], 10);
      var numSection2 = parseInt(formulaMatch2[2], 10);
      var numFormula2 = parseInt(formulaMatch2[3], 10);

      // 更新计数器
      var key2 = currentChapter + '.' + currentSection;
      formulaCounters[key2] = (formulaCounters[key2] || 0) + 1;
      var expectedNum2 = formulaCounters[key2];

      // 检查章节号或编号是否正确
      if (numChapter2 !== currentChapter || numSection2 !== currentSection || numFormula2 !== expectedNum2) {
        var suggestedText2 = text.replace(
          new RegExp('\\(' + numChapter2 + '\\.' + numSection2 + '[-—–]' + numFormula2 + '\\)\\s*$'),
          '(' + currentChapter + '.' + currentSection + '-' + expectedNum2 + ')'
        );

        issues.push({
          index: p.index,
          rule: 'E-001',
          name: '公式编号错误',
          original: text,
          suggested: suggestedText2,
          message: '公式编号 ' + numChapter2 + '.' + numSection2 + '-' + numFormula2 + ' → ' + currentChapter + '.' + currentSection + '-' + expectedNum2,
          autoFix: true
        });
      }
    } else if (formulaMatch1) {
      // 最简格式 (X-N)
      var numChapter1 = parseInt(formulaMatch1[1], 10);
      var numFormula1 = parseInt(formulaMatch1[2], 10);

      console.log('[checkFormulaNumbering] 检测到(X-N)格式公式: 段落' + p.index +
        ', 编号=' + numChapter1 + '-' + numFormula1 +
        ', currentChapter=' + currentChapter + ', currentSection=' + currentSection);

      // 检查是否在附录中
      if (inAppendix && currentAppendix) {
        // 附录中的公式，使用附录编号
        appendixFormulaCounter++;
        var suggestedTextApp = text.replace(
          /\((\d+)[-—–](\d+)\)\s*$/,
          '(' + currentAppendix + appendixFormulaCounter + ')'
        );
        issues.push({
          index: p.index,
          rule: 'E-001',
          name: '附录公式编号格式错误',
          original: text,
          suggested: suggestedTextApp,
          message: '附录' + currentAppendix + '中的公式编号应为 (' + currentAppendix + 'N) 格式',
          autoFix: true
        });
        continue;
      }

      // 【重要】如果当前有节号（currentSection > 0），则 (X-N) 格式不规范
      // 应该建议改为 (X.Y-N) 格式
      if (currentSection > 0) {
        // 更新计数器（使用 X.Y 作为 key）
        var key1 = currentChapter + '.' + currentSection;
        formulaCounters[key1] = (formulaCounters[key1] || 0) + 1;
        var expectedNum1 = formulaCounters[key1];

        var suggestedText1 = text.replace(
          new RegExp('\\(' + numChapter1 + '[-—–]' + numFormula1 + '\\)\\s*$'),
          '(' + currentChapter + '.' + currentSection + '-' + expectedNum1 + ')'
        );

        issues.push({
          index: p.index,
          rule: 'E-001',
          name: '公式编号格式不规范',
          original: text,
          suggested: suggestedText1,
          message: '公式编号格式不规范，当前在章节 ' + currentChapter + '.' + currentSection + '，建议改为 (' + currentChapter + '.' + currentSection + '-' + expectedNum1 + ')',
          autoFix: true
        });
        continue;
      }

      // 只有章节号，无节号的情况
      formulaCounters[currentChapter] = (formulaCounters[currentChapter] || 0) + 1;
      var expectedNum1 = formulaCounters[currentChapter];

      // 检查章节号是否正确
      if (numChapter1 !== currentChapter || numFormula1 !== expectedNum1) {
        console.log('[checkFormulaNumbering] 发现问题! 添加到 issues 列表');
        var suggestedText1 = text.replace(
          new RegExp('\\(' + numChapter1 + '[-—–]' + numFormula1 + '\\)\\s*$'),
          '(' + currentChapter + '-' + expectedNum1 + ')'
        );

        issues.push({
          index: p.index,
          rule: 'E-001',
          name: '公式编号错误',
          original: text,
          suggested: suggestedText1,
          message: '公式编号 ' + numChapter1 + '-' + numFormula1 + ' → ' + currentChapter + '-' + expectedNum1,
          autoFix: true
        });
      }
    }
  }

  return issues;
}

function checkStructure(structure) {
  var issues = [];
  var chapters = { level1: [], level2: [], level3: [], level4: [], appendix: [] };

  // 中文数字转阿拉伯数字
  var chineseToNumber = function(cn) {
    var map = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
                '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
                '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15,
                '十六': 16, '十七': 17, '十八': 18, '十九': 19, '二十': 20 };
    return map[cn] || 0;
  };

  // 阿拉伯数字转中文
  var numberToChinese = function(num) {
    var map = ['', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十',
               '十一', '十二', '十三', '十四', '十五', '十六', '十七', '十八', '十九', '二十'];
    return map[num] || String(num);
  };

  function getHeadingLevel(styleName, text) {
    // 附录标题
    if (/^附录\s*[A-Z一二三四五六七八九十]/i.test(text)) return 10;  // 特殊标记

    // 优先检查文本格式（因为很多文档没有应用标题样式）
    // 中文章节编号（第一章、第二章...）
    if (/^第[一二三四五六七八九十百]+章/.test(text)) return 1;
    // 阿拉伯数字编号 - 使用负向前瞻避免匹配二级标题 "1.1"
    if (/^(\d+)\s+(?!\.)/.test(text)) return 1;
    if (/^\d+\.\d+\s/.test(text)) return 2;
    if (/^\d+\.\d+\.\d+\s/.test(text)) return 3;
    if (/^\d+\.\d+\.\d+\.\d+\s/.test(text)) return 4;

    // 再检查样式名称
    if (!styleName) return 0;
    var upperName = styleName.toUpperCase();

    if (upperName.indexOf('HEADING 1') >= 0 || upperName.indexOf('标题 1') >= 0) return 1;
    if (upperName.indexOf('HEADING 2') >= 0 || upperName.indexOf('标题 2') >= 0) return 2;
    if (upperName.indexOf('HEADING 3') >= 0 || upperName.indexOf('标题 3') >= 0) return 3;
    if (upperName.indexOf('HEADING 4') >= 0 || upperName.indexOf('标题 4') >= 0) return 4;

    return 0;
  }

  function extractChapterNumber(text, level) {
    if (level === 10) {
      // 附录编号
      var appMatch = text.match(/^附录\s*([A-Z一二三四五六七八九十])/i);
      if (appMatch) return { letter: appMatch[1].toUpperCase() };
      return null;
    }
    if (level === 1) {
      // 中文章节编号：第一章、第二章...
      var cnMatch = text.match(/^第([一二三四五六七八九十百]+)章/);
      if (cnMatch) {
        return { number: chineseToNumber(cnMatch[1]), type: 'chinese' };
      }
      // 阿拉伯数字编号：1 设计依据
      var m = text.match(/^(\d+)\s/);
      return m ? { number: parseInt(m[1], 10), type: 'arabic' } : null;
    }
    if (level === 2) {
      var m2 = text.match(/^(\d+)\.(\d+)\s/);
      return m2 ? { chapter: parseInt(m2[1], 10), section: parseInt(m2[2], 10) } : null;
    }
    if (level === 3) {
      var m3 = text.match(/^(\d+)\.(\d+)\.(\d+)\s/);
      return m3 ? {
        chapter: parseInt(m3[1], 10),
        section: parseInt(m3[2], 10),
        subsection: parseInt(m3[3], 10)
      } : null;
    }
    if (level === 4) {
      var m4 = text.match(/^(\d+)\.(\d+)\.(\d+)\.(\d+)\s/);
      return m4 ? {
        chapter: parseInt(m4[1], 10),
        section: parseInt(m4[2], 10),
        subsection: parseInt(m4[3], 10),
        item: parseInt(m4[4], 10)
      } : null;
    }
    return null;
  }

  // 收集章节信息，同时跟踪当前章节
  var paragraphs = structure.paragraphs || [];
  var inAppendix = false;  // 是否进入附录区域
  var currentAppendix = '';  // 当前附录编号（A, B, C...）
  var currentLevel1 = { num: 0, type: 'arabic' };  // 当前一级标题

  // 先遍历一遍，标记所有附录区域内的段落
  // 关键：附录区域只会在遇到下一个附录标题或中文章节标题时结束
  // 阿拉伯数字标题如 "1 xxx" 在附录内应视为附录标题，不触发退出
  var appendixParagraphs = {};
  var tempInAppendix = false;
  var lastMainChapter = 0;  // 记录最后一个正文章节号

  for (var i = 0; i < paragraphs.length; i++) {
    var text = paragraphs[i].text || '';

    // 检测附录标题
    if (/^附录\s*[A-Z一二三四五六七八九十]/i.test(text)) {
      tempInAppendix = true;
    }
    // 只在中文章节标题时退出附录（因为中文章节标题不会出现在附录内）
    else if (/^第[一二三四五六七八九十百]+章/.test(text)) {
      if (tempInAppendix) {
        tempInAppendix = false;
      }
      // 更新正文章节计数
      var cnMatch = text.match(/^第([一二三四五六七八九十百]+)章/);
      if (cnMatch) {
        lastMainChapter = chineseToNumber(cnMatch[1]);
      }
    }
    // 阿拉伯数字一级标题：只有在确定是正文章节时才退出附录
    // 判断标准：章节号应该是上一个正文章节+1，或者从1开始但前面没有正文章节
    else if (/^(\d+)\s+(?!\.)/.test(text) && tempInAppendix) {
      var arMatch = text.match(/^(\d+)\s+(?!\.)/);
      var arNum = parseInt(arMatch[1], 10);
      // 如果章节号明显是正文的延续（如前一章是3，现在是4），则退出附录
      // 否则视为附录内的标题，保持在附录状态
      if (lastMainChapter > 0 && arNum === lastMainChapter + 1) {
        tempInAppendix = false;
        lastMainChapter = arNum;
      }
      // 否则保持在附录状态（附录内的 "1 xxx", "2 xxx" 等）
    }

    if (tempInAppendix) {
      appendixParagraphs[paragraphs[i].index] = true;
    }
  }

  var appendixTitleCounters = {};  // 每个附录的标题计数器
  for (var i = 0; i < paragraphs.length; i++) {
    var p = paragraphs[i];
    var text = p.fullText || p.text || '';
    var level = getHeadingLevel(p.styleName, text);

    // 检测是否进入附录（更新 currentAppendix）
    if (level === 10) {
      inAppendix = true;
      var appMatch = text.match(/^附录\s*([A-Z一二三四五六七八九十])/i);
      if (appMatch) {
        var letter = appMatch[1].toUpperCase();
        var letterMap = { '一': 'A', '二': 'B', '三': 'C', '四': 'D', '五': 'E',
                          '六': 'F', '七': 'G', '八': 'H', '九': 'I', '十': 'J' };
        currentAppendix = letterMap[letter] || letter;
      }
      // 【修复】在 continue 之前收集附录标题
      chapters.appendix.push({ index: p.index, text: text, letter: currentAppendix });
      console.log('[checkStructure] 收集附录标题: ' + text.substring(0, 30) + ', letter=' + currentAppendix);
      continue;  // 附录标题本身不收集到其他列表
    }

    // 【关键修改】使用预扫描结果判断是否在附录中
    // 如果这个段落被标记为附录段落，直接处理附录逻辑并跳过正文收集
    if (appendixParagraphs[p.index] && level > 0 && level < 10) {
      // 附录区域的标题，检查是否需要添加附录前缀
      var appendixTitleMatch = text.match(/^(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:\.(\d+))?\s/);
      if (appendixTitleMatch) {
        var titleText = text.replace(/^[\d\.]+\s+/, '');
        var newPrefix = currentAppendix || 'A';

        // 【修复】附录中标题应该是简单编号：A1, A2, A3 或 B1, B2, B3
        // 不保留原来的数字结构，而是重新从 1 开始编号
        var counterKey = newPrefix;
        if (!appendixTitleCounters[counterKey]) {
          appendixTitleCounters[counterKey] = 0;
        }
        appendixTitleCounters[counterKey]++;
        var newNum = appendixTitleCounters[counterKey];

        // 生成新的标题：A1, A2, B1, B2, C1, C2...
        var suggestedText = newPrefix + newNum + ' ' + titleText;
        console.log('[N-007 附录标题] 原文: ' + text + ', 修改为: ' + suggestedText);

        issues.push({
          index: p.index,
          rule: 'N-007',
          name: '附录标题编号格式错误',
          original: text,
          suggested: suggestedText,
          message: '附录标题编号应为 ' + newPrefix + newNum + ' 格式',
          autoFix: true
        });
      }
      continue;  // 跳过附录区域内的标题，不收集到 chapters
    }

    // 检测是否离开附录
    if (level === 1 && inAppendix) {
      inAppendix = false;
      currentAppendix = '';
    }

    var num = extractChapterNumber(text, level);

    if (level > 0 && level < 10 && num) {
      var info = { index: p.index, text: text, number: num, level: level, inAppendix: false };  // 正文标题，inAppendix始终为false

      // 对于二级及以下标题，附加当前章节信息
      if (level === 2 || level === 3 || level === 4) {
        info.parentChapter = currentLevel1;
      }

      if (level === 1) {
        chapters.level1.push(info);
        currentLevel1 = { num: (typeof num === 'object') ? num.number : num, type: (typeof num === 'object') ? num.type : 'arabic' };
      }
      else if (level === 2) chapters.level2.push(info);
      else if (level === 3) chapters.level3.push(info);
      else if (level === 4) chapters.level4.push(info);
    }
    // 注意：附录标题已在前面收集，这里不再重复
  }

  // N-002: 一级标题连续性（不包括附录区域内的标题）
  console.log('[checkStructure] 收集到的一级标题数量: ' + chapters.level1.length);
  for (var i = 0; i < chapters.level1.length; i++) {
    console.log('[checkStructure] 一级标题[' + i + ']: ' + chapters.level1[i].text.substring(0, 30) + ', inAppendix=' + chapters.level1[i].inAppendix);
  }

  // 【重要】先构建一级标题的预期编号映射
  var level1ExpectedMap = {};  // paraIndex -> expected chapter number
  var level1Counter = 0;
  for (var i = 0; i < chapters.level1.length; i++) {
    var info = chapters.level1[i];
    if (info.inAppendix) continue;

    level1Counter++;
    var expected = level1Counter;
    var numObj = info.number;
    var actual = (typeof numObj === 'object') ? numObj.number : numObj;
    var numType = (typeof numObj === 'object') ? numObj.type : 'arabic';

    // 记录这个段落应该对应的章节号
    level1ExpectedMap[info.index] = { expected: expected, actual: actual, numType: numType };
    console.log('[N-002] 一级标题映射: index=' + info.index + ', actual=' + actual + ', expected=' + expected);

    if (actual !== expected) {
      var originalText = info.text;
      var suggestedText;

      if (numType === 'chinese') {
        var expectedCn = numberToChinese(expected);
        var actualCn = numberToChinese(actual);
        suggestedText = originalText.replace(
          new RegExp('第' + actualCn + '章'),
          '第' + expectedCn + '章'
        );
      } else {
        suggestedText = originalText.replace(
          new RegExp('^(第\\s*)?' + actual + '(\\s+|\\s*章)'),
          function(m, p1, p2) {
            return (p1 || '') + expected + (p2 || ' ');
          }
        );
        if (suggestedText === originalText) {
          suggestedText = originalText.replace(new RegExp('^' + actual + '\\s'), expected + ' ');
        }
      }

      issues.push({
        index: info.index,
        rule: 'N-002',
        name: '一级标题编号不连续',
        original: originalText,
        suggested: suggestedText,
        message: '编号 ' + (numType === 'chinese' ? numberToChinese(actual) : actual) + ' → ' + (numType === 'chinese' ? numberToChinese(expected) : expected),
        autoFix: true
      });
    }
  }

  // 【新增】N-007: 附录标题编号（附录一 → 附录 A）
  console.log('[checkStructure] 收集到的附录标题数量: ' + chapters.appendix.length);
  for (var i = 0; i < chapters.appendix.length; i++) {
    var appInfo = chapters.appendix[i];
    console.log('[checkStructure] 附录标题[' + i + ']: ' + appInfo.text.substring(0, 30));
  }
  for (var i = 0; i < chapters.appendix.length; i++) {
    var appInfo = chapters.appendix[i];
    var appText = appInfo.text;
    var letterMap = { '一': 'A', '二': 'B', '三': 'C', '四': 'D', '五': 'E',
                      '六': 'F', '七': 'G', '八': 'H', '九': 'I', '十': 'J' };

    // 检查是否是中文数字格式
    var cnMatch = appText.match(/^附录\s*([一二三四五六七八九十])/);
    console.log('[N-007] 检查附录标题: ' + appText.substring(0, 30) + ', 匹配结果: ' + (cnMatch ? cnMatch[1] : '无'));
    if (cnMatch) {
      var oldLetter = cnMatch[1];
      var newLetter = letterMap[oldLetter] || oldLetter;
      var suggestedText = appText.replace(/^附录\s*[一二三四五六七八九十]/, '附录 ' + newLetter);
      console.log('[N-007] 生成修复: "' + oldLetter + '" → "' + newLetter + '", 原文: ' + appText.substring(0, 30) + ', 建议: ' + suggestedText.substring(0, 30));

      issues.push({
        index: appInfo.index,
        rule: 'N-007',
        name: '附录编号格式',
        original: appText,
        suggested: suggestedText,
        message: '附录编号 "附录' + oldLetter + '" → "附录 ' + newLetter + '"',
        autoFix: true
      });
    }
  }

  // N-003: 二级标题连续性及章节号正确性
  // 【重要】使用预期的一级标题编号来设置 currentCh
  var currentCh = 0;  // 当前章节编号（预期值）
  var sectionCounts = {};  // 每章的二级标题计数
  var inAppendixN3 = false;  // 是否在附录区域

  for (var i = 0; i < paragraphs.length; i++) {
    var text = paragraphs[i].text || '';

    // 检测是否进入/离开附录
    if (/^附录\s*[A-Z一二三四五六七八九十]/i.test(text)) {
      inAppendixN3 = true;
      continue;
    }
    if (/^第[一二三四五六七八九十百]+章/.test(text)) {
      inAppendixN3 = false;
    }

    // 跳过附录区域内的段落
    if (inAppendixN3) continue;

    // 【关键修改】使用预期的一级标题编号
    var pIndex = paragraphs[i].index;
    if (level1ExpectedMap[pIndex]) {
      // 这是一个一级标题，更新 currentCh 为预期值
      currentCh = level1ExpectedMap[pIndex].expected;
      sectionCounts[currentCh] = 0;
      console.log('[N-003] 一级标题: ' + text.substring(0, 20) + ', 预期章节号=' + currentCh);
    }

    // 检测二级标题
    var l2Match = text.match(/^(\d+)\.(\d+)\s/);
    if (l2Match && currentCh > 0) {
      var titleCh = parseInt(l2Match[1], 10);  // 标题中的章节号
      var titleSec = parseInt(l2Match[2], 10);  // 标题中的节号
      console.log('[N-003] 二级标题: ' + text.substring(0, 30) + ', titleCh=' + titleCh + ', currentCh=' + currentCh);

      // 检查章节号是否正确（使用预期的章节号）
      if (titleCh !== currentCh) {
        // 章节号错误，需要修正
        sectionCounts[currentCh] = (sectionCounts[currentCh] || 0) + 1;
        var newSec = sectionCounts[currentCh];

        var suggestedText = text.replace(
          new RegExp('^' + titleCh + '\\.' + titleSec + '\\s'),
          currentCh + '.' + newSec + ' '
        );

        issues.push({
          index: paragraphs[i].index,
          rule: 'N-003',
          name: '二级标题章节号错误',
          original: text,
          suggested: suggestedText,
          message: '章节号错误：' + titleCh + '.' + titleSec + ' → ' + currentCh + '.' + newSec,
          autoFix: true
        });
      } else {
        // 章节号正确，检查连续性
        sectionCounts[currentCh] = (sectionCounts[currentCh] || 0) + 1;
        var expectedSec = sectionCounts[currentCh];

        if (titleSec !== expectedSec) {
          var suggestedText = text.replace(
            new RegExp('^' + currentCh + '\\.' + titleSec + '\\s'),
            currentCh + '.' + expectedSec + ' '
          );

          issues.push({
            index: paragraphs[i].index,
            rule: 'N-003',
            name: '二级标题编号不连续',
            original: text,
            suggested: suggestedText,
            message: '编号 ' + currentCh + '.' + titleSec + ' → ' + currentCh + '.' + expectedSec,
            autoFix: true
          });
        }
      }
    }
  }

  // N-004: 三级标题连续性及章节号正确性
  var currentCh3 = 0;
  var currentSec = 0;
  var subsectionCounts = {};
  var inAppendixN4 = false;  // 是否在附录区域

  for (var i = 0; i < paragraphs.length; i++) {
    var text = paragraphs[i].text || '';

    // 检测是否进入/离开附录
    if (/^附录\s*[A-Z一二三四五六七八九十]/i.test(text)) {
      inAppendixN4 = true;
      continue;
    }
    if (/^第[一二三四五六七八九十百]+章/.test(text)) {
      inAppendixN4 = false;
    }

    // 跳过附录区域内的段落
    if (inAppendixN4) continue;

    // 【关键修改】使用预期的一级标题编号
    var pIndex4 = paragraphs[i].index;
    if (level1ExpectedMap[pIndex4]) {
      currentCh3 = level1ExpectedMap[pIndex4].expected;
      currentSec = 0;
    }

    // 更新当前节（使用预期章节号）
    var l2Match = text.match(/^(\d+)\.(\d+)\s/);
    if (l2Match) {
      currentSec = parseInt(l2Match[2], 10);
      var key = currentCh3 + '.' + currentSec;
      subsectionCounts[key] = 0;
    }

    // 检测三级标题
    var l3Match = text.match(/^(\d+)\.(\d+)\.(\d+)\s/);
    if (l3Match && currentCh3 > 0 && currentSec > 0) {
      var titleCh3 = parseInt(l3Match[1], 10);
      var titleSec3 = parseInt(l3Match[2], 10);
      var titleSub = parseInt(l3Match[3], 10);

      var key = currentCh3 + '.' + currentSec;

      // 先增加计数
      subsectionCounts[key] = (subsectionCounts[key] || 0) + 1;
      var expectedSub = subsectionCounts[key];

      // 检查章节号和节号是否正确，以及编号是否连续
      if (titleCh3 !== currentCh3 || titleSec3 !== currentSec || titleSub !== expectedSub) {
        // 需要修正
        var suggestedText = text.replace(
          new RegExp('^' + titleCh3 + '\\.' + titleSec3 + '\\.' + titleSub + '\\s'),
          currentCh3 + '.' + currentSec + '.' + expectedSub + ' '
        );

        var originalText = text;

        issues.push({
          index: paragraphs[i].index,
          rule: 'N-004',
          name: '三级标题编号错误',
          original: originalText,
          suggested: suggestedText,
          message: '编号 ' + titleCh3 + '.' + titleSec3 + '.' + titleSub + ' → ' + currentCh3 + '.' + currentSec + '.' + expectedSub,
          autoFix: true
        });
      }
    }
  }

  // N-005: 四级标题编号连续性
  var currentCh4 = 0;
  var currentSec4 = 0;
  var currentSub4 = 0;
  var itemCounts = {};
  var inAppendixN5 = false;  // 是否在附录区域

  for (var i = 0; i < paragraphs.length; i++) {
    var text = paragraphs[i].text || '';

    // 检测是否进入/离开附录
    if (/^附录\s*[A-Z一二三四五六七八九十]/i.test(text)) {
      inAppendixN5 = true;
      continue;
    }
    if (/^第[一二三四五六七八九十百]+章/.test(text)) {
      inAppendixN5 = false;
    }

    // 跳过附录区域内的段落
    if (inAppendixN5) continue;

    // 【关键修改】使用预期的一级标题编号
    var pIndex5 = paragraphs[i].index;
    if (level1ExpectedMap[pIndex5]) {
      currentCh4 = level1ExpectedMap[pIndex5].expected;
      currentSec4 = 0;
      currentSub4 = 0;
    }

    // 更新当前节
    var l2Match = text.match(/^(\d+)\.(\d+)\s/);
    if (l2Match) {
      currentSec4 = parseInt(l2Match[2], 10);
      currentSub4 = 0;
    }

    // 更新当前小节
    var l3Match = text.match(/^(\d+)\.(\d+)\.(\d+)\s/);
    if (l3Match) {
      currentSub4 = parseInt(l3Match[3], 10);
      var key = currentCh4 + '.' + currentSec4 + '.' + currentSub4;
      itemCounts[key] = 0;
    }

    // 检测四级标题
    var l4Match = text.match(/^(\d+)\.(\d+)\.(\d+)\.(\d+)\s/);
    if (l4Match && currentCh4 > 0 && currentSec4 > 0 && currentSub4 > 0) {
      var titleCh4 = parseInt(l4Match[1], 10);
      var titleSec4 = parseInt(l4Match[2], 10);
      var titleSub4 = parseInt(l4Match[3], 10);
      var titleItem = parseInt(l4Match[4], 10);

      var key = currentCh4 + '.' + currentSec4 + '.' + currentSub4;

      // 先增加计数
      itemCounts[key] = (itemCounts[key] || 0) + 1;
      var expectedItem = itemCounts[key];

      // 检查编号是否正确
      if (titleCh4 !== currentCh4 || titleSec4 !== currentSec4 || titleSub4 !== currentSub4 || titleItem !== expectedItem) {
        var suggestedText = text.replace(
          new RegExp('^' + titleCh4 + '\\.' + titleSec4 + '\\.' + titleSub4 + '\\.' + titleItem + '\\s'),
          currentCh4 + '.' + currentSec4 + '.' + currentSub4 + '.' + expectedItem + ' '
        );

        issues.push({
          index: paragraphs[i].index,
          rule: 'N-005',
          name: '四级标题编号错误',
          original: text,
          suggested: suggestedText,
          message: '编号 ' + titleCh4 + '.' + titleSec4 + '.' + titleSub4 + '.' + titleItem + ' → ' + currentCh4 + '.' + currentSec4 + '.' + currentSub4 + '.' + expectedItem,
          autoFix: true
        });
      }
    }
  }

  // N-006: 标题层级过深（五级及以上）
  for (var i = 0; i < paragraphs.length; i++) {
    var p = paragraphs[i];
    var text = p.fullText || p.text || '';

    // 检测五级及以上标题，如 1.1.1.1.1 标题
    var match = text.match(/^((\d+\.){4,})(\d+)\s+(.+)$/);
    if (match) {
      var parts = match[1].split('.').filter(Boolean);
      parts.pop();
      var newPrefix = parts.join('.') + ' ';
      var suggestedText = newPrefix + match[4];

      issues.push({
        index: p.index,
        rule: 'N-006',
        name: '标题层级过深',
        original: text,
        suggested: suggestedText,
        message: '标题层级超过四级，已提升为四级标题',
        autoFix: true
      });
    }
  }

  // N-007: 附录编号（改为字母）
  for (var i = 0; i < chapters.appendix.length; i++) {
    var info = chapters.appendix[i];
    var letter = info.letter;
    var letterMap = { '一': 'A', '二': 'B', '三': 'C', '四': 'D', '五': 'E',
                      '六': 'F', '七': 'G', '八': 'H', '九': 'I', '十': 'J' };

    if (letterMap[letter]) {
      var suggestedText = info.text.replace(
        new RegExp('附录\\s*' + letter),
        '附录 ' + letterMap[letter]
      );
      issues.push({
        index: info.index,
        rule: 'N-007',
        name: '附录编号',
        original: info.text,
        suggested: suggestedText,
        message: '附录编号 "' + letter + '" → "' + letterMap[letter] + '"',
        autoFix: true
      });
    }
  }

  // N-008: 章节编号唯一性
  var seenLevel1 = {};
  var l1Counter = 0;

  for (var i = 0; i < chapters.level1.length; i++) {
    var info = chapters.level1[i];
    if (info.inAppendix) continue;

    var numObj = info.number;
    var numValue = (typeof numObj === 'object') ? numObj.number : numObj;
    var numType = (typeof numObj === 'object') ? numObj.type : 'arabic';

    if (seenLevel1[numValue + '_' + numType]) {
      // 重复，生成新编号
      l1Counter++;
      var newNum = l1Counter;
      var suggestedText;

      if (numType === 'chinese') {
        suggestedText = info.text.replace(
          new RegExp('第' + numberToChinese(numValue) + '章'),
          '第' + numberToChinese(newNum) + '章'
        );
      } else {
        suggestedText = info.text.replace(new RegExp('^' + numValue + '\\s'), newNum + ' ');
      }

      issues.push({
        index: info.index,
        rule: 'N-008',
        name: '章节编号重复',
        original: info.text,
        suggested: suggestedText,
        message: '一级标题编号重复，已调整为 ' + (numType === 'chinese' ? '第' + numberToChinese(newNum) + '章' : newNum),
        autoFix: true
      });
    } else {
      seenLevel1[numValue + '_' + numType] = true;
      l1Counter = numValue;
    }
  }

  // N-009: 并列编号连续性（a、b、c...）- 检测同一行内的并列编号
  for (var i = 0; i < paragraphs.length; i++) {
    var p = paragraphs[i];
    var text = p.fullText || p.text || '';

    // 检测同一行内的并列编号，如 "a、xxx；b、xxx；c、xxx"
    var listItems = text.match(/[a-z][、．.]\s*[^；。\n]+[；。]?/gi);
    if (listItems && listItems.length >= 2) {
      // 检查编号是否连续
      var expectedLetter = 'a'.charCodeAt(0);
      for (var j = 0; j < listItems.length; j++) {
        var item = listItems[j];
        var letterMatch = item.match(/^([a-z])[、．.]/i);
        if (letterMatch) {
          var actualLetter = letterMatch[1].toLowerCase();
          var expected = String.fromCharCode(expectedLetter + j);
          if (actualLetter !== expected) {
            // 生成建议文本
            var suggestedText = text.replace(
              new RegExp('(^|[^a-zA-Z])' + actualLetter + '([、．.])', 'gi'),
              function(m, before, sep) {
                return before + expected + '、';
              }
            );
            issues.push({
              index: p.index,
              rule: 'N-009',
              name: '并列编号不连续',
              original: text.substring(0, 50),
              suggested: suggestedText.substring(0, 50),
              message: '并列编号 ' + actualLetter + ' → ' + expected,
              autoFix: false  // 同一行内的替换较复杂，建议人工确认
            });
            break;  // 只报告第一个问题
          }
        }
      }
    }
  }

  // N-010: 参考文献编号连续性
  var refs = [];
  for (var i = 0; i < paragraphs.length; i++) {
    var text = paragraphs[i].text || '';
    var m = text.match(/^\[(\d+)\]/);
    if (m) {
      refs.push({ index: paragraphs[i].index, num: parseInt(m[1], 10), text: text });
    }
  }

  // 检查参考文献编号连续性和唯一性
  var seenRef = {};
  var expectedRef = 0;
  for (var i = 0; i < refs.length; i++) {
    var ref = refs[i];
    expectedRef++;

    // 检查重复
    if (seenRef[ref.num]) {
      var suggestedText = ref.text.replace(
        new RegExp('^\\[' + ref.num + '\\]'),
        '[' + expectedRef + ']'
      );
      issues.push({
        index: ref.index,
        rule: 'N-010',
        name: '参考文献编号重复',
        original: ref.text,
        suggested: suggestedText,
        message: '参考文献编号 [' + ref.num + '] 重复，已调整为 [' + expectedRef + ']',
        autoFix: true
      });
      seenRef[expectedRef] = true;
    } else {
      seenRef[ref.num] = true;

      // 检查连续性
      if (ref.num !== expectedRef) {
        var suggestedText = ref.text.replace(
          new RegExp('^\\[' + ref.num + '\\]'),
          '[' + expectedRef + ']'
        );
        issues.push({
          index: ref.index,
          rule: 'N-010',
          name: '参考文献编号不连续',
          original: ref.text,
          suggested: suggestedText,
          message: '参考文献编号 [' + ref.num + '] → [' + expectedRef + ']',
          autoFix: true
        });
      }
    }
  }

  return issues;
}

// ========== 内部函数：应用修复和批注 ==========

function applyFixesAndComments(doc, allIssues, mode) {
  var fixed = 0;
  var commented = 0;
  var byRule = {};
  var revisionLog = [];

  // 可自动修复的规则
  var AUTO_FIX_RULES = {
    'F-001': true, 'F-002': true, 'F-003': true, 'F-004': true, 'F-005': true,
    'G-002': true, 'G-004': true,  // 图名格式、图片居中
    // G-003 图名位置只批注提醒，不自动修改
    'V-006': true, 'V-007': true, 'V-008': true, 'V-009': true, 'V-010': true, 'V-012': true,
    'M-001': true, 'M-002': true, 'M-005': true, 'M-006': true,
    // 编号类规则（Stage 1）
    'N-002': true, 'N-003': true, 'N-004': true, 'N-005': true, 'N-006': true, 'N-007': true, 'N-008': true,
    'N-009': true, 'N-010': true,  // 并列编号、参考文献编号
    'T-001': true, 'G-001': true, 'E-001': true,
    // 表格格式规则（Stage 4）
    'T-002': true,  // 表名字体字号
    'T-004': true,  // 表格宽度
    'T-007': true,  // 跨页重复表头
    // 公式格式规则（Stage 4）
    'E-002': true,  // 公式居中
    'E-003': true,  // 公式编号位置
    // 页眉页脚规则（Stage 4）
    'HF-001': true,  // 页眉页脚字号
    'HF-002': true,  // 页眉线
    'HF-003': true,  // 页脚线
    // 页面设置规则（Stage 4）
    'PG-001': true,  // 页边距
    'PG-002': true,  // 页眉页脚距边界
    // 表格内容规则（Stage 2）
    'T-005': true, 'T-006': true
  };

  // 从后往前排序，避免位置偏移（仅对段落问题）
  var paraIssues = allIssues.filter(function(issue) { return issue.index && !issue.tableIndex; });
  var tableIssues = allIssues.filter(function(issue) { return issue.tableIndex; });
  var sectionIssues = allIssues.filter(function(issue) {
    return issue.sectionIndex && !issue.index && !issue.tableIndex;
  });

  paraIssues.sort(function(a, b) {
    return (b.index || 0) - (a.index || 0);
  });

  // 处理段落问题
  for (var i = 0; i < paraIssues.length; i++) {
    var issue = paraIssues[i];
    var rule = issue.rule || '';
    var canAutoFix = issue.autoFix === true;

    // 根据修复模式决定处理方式
    var shouldFix = false;
    if (mode === 'aggressive') {
      shouldFix = canAutoFix && AUTO_FIX_RULES[rule];
    } else if (mode === 'standard') {
      // 标准模式：格式类 + 编号类自动修复，内容类批注
      shouldFix = canAutoFix && AUTO_FIX_RULES[rule] &&
        (rule.indexOf('F-') === 0 || rule.indexOf('G-') === 0 || rule.indexOf('N-') === 0 ||
         rule === 'T-001' || rule === 'T-002' || rule === 'T-004' || rule === 'T-007' ||
         rule === 'E-002' || rule === 'E-003' || rule === 'E-001');
    }
    // conservative 模式下 shouldFix 始终为 false

    try {
      var para = doc.Paragraphs.Item(issue.index);
      if (!para || !para.Range) continue;

      var range = para.Range;

      if (shouldFix && issue.suggested) {
        // 直接替换段落文本（更可靠）
        try {
          var originalText = range.Text;
          var newText = String(issue.suggested);

          console.log('[applyFixes] 尝试替换段落 ' + issue.index + ': ' + rule);
          console.log('[applyFixes] 原文: ' + (originalText || '').substring(0, 50));
          console.log('[applyFixes] 建议修改为: ' + newText.substring(0, 50));

          // 【关键修复】保存段落标记
          var hadParaMark = originalText && /\r$/.test(originalText);

          // 尝试精确匹配替换
          if (originalText && originalText.indexOf(issue.original) >= 0) {
            // 替换匹配的部分
            newText = originalText.replace(issue.original, issue.suggested);
          }

          // 【关键修复】确保保留段落标记，防止段落合并
          if (hadParaMark && newText && !/\r$/.test(newText)) {
            newText = newText + '\r';
            console.log('[applyFixes] 添加段落标记');
          }

          // 设置新文本
          range.Text = newText;
          fixed++;
          byRule[rule] = (byRule[rule] || 0) + 1;
          console.log('[applyFixes] 替换成功');
          // 记录修订日志
          revisionLog.push({
            rule: rule,
            name: issue.name,
            paraIndex: issue.index,
            original: issue.original,
            suggested: issue.suggested,
            message: issue.message
          });
        } catch (e) {
          console.warn('[applyFixes] 替换失败: ' + e);
          // 替换失败，改为批注
          addCommentToDoc(doc, range, issue);
          commented++;
          byRule[rule] = (byRule[rule] || 0) + 1;
        }
      } else if (shouldFix && rule === 'T-002') {
        // T-002: 表名字体字号修复
        try {
          console.log('[applyFixes] 修复表名字体字号 段落 ' + issue.index);
          var para = doc.Paragraphs.Item(issue.index);
          if (!para) continue;

          var spec = issue.fixSpec || {};
          var range = para.Range;

          // 设置字体
          if (spec.fontName && range && range.Font) {
            range.Font.NameFarEast = spec.fontName;
            range.Font.Name = spec.fontName;
            console.log('[applyFixes] 设置表名字体: ' + spec.fontName);
          }

          // 设置字号
          if (spec.fontSize && range && range.Font) {
            range.Font.Size = spec.fontSize;
            console.log('[applyFixes] 设置表名字号: ' + spec.fontSize + '磅');
          }

          // 设置对齐方式
          if (spec.alignment !== undefined && para.Format && para.Format.Alignment !== undefined) {
            para.Format.Alignment = spec.alignment;
            console.log('[applyFixes] 设置表名对齐: ' + (spec.alignment === 1 ? '居中' : spec.alignment));
          }

          // 设置段前间距
          if (spec.spaceBefore !== undefined && para.Format && para.Format.SpaceBefore !== undefined) {
            para.Format.SpaceBefore = spec.spaceBefore;
            console.log('[applyFixes] 设置表名段前: ' + spec.spaceBefore + '磅');
          }

          // 设置段后间距
          if (spec.spaceAfter !== undefined && para.Format && para.Format.SpaceAfter !== undefined) {
            para.Format.SpaceAfter = spec.spaceAfter;
            console.log('[applyFixes] 设置表名段后: ' + spec.spaceAfter + '磅');
          }

          fixed++;
          byRule[rule] = (byRule[rule] || 0) + 1;
          revisionLog.push({
            rule: rule,
            name: issue.name,
            paraIndex: issue.index,
            original: issue.original,
            message: issue.message
          });
        } catch (e) {
          console.warn('[applyFixes] 表名字体字号修复失败: ' + e);
          var para = doc.Paragraphs.Item(issue.index);
          if (para) addCommentToDoc(doc, para.Range, issue);
          commented++;
          byRule[rule] = (byRule[rule] || 0) + 1;
        }
      } else if (shouldFix && rule === 'E-002') {
        // E-002: 公式居中修复
        try {
          console.log('[applyFixes] 修复公式居中 段落 ' + issue.index);
          var para = doc.Paragraphs.Item(issue.index);
          if (!para) continue;

          var spec = issue.fixSpec || {};
          if (spec.alignment !== undefined && para.Format && para.Format.Alignment !== undefined) {
            para.Format.Alignment = spec.alignment;
            console.log('[applyFixes] 设置公式居中');
          }

          fixed++;
          byRule[rule] = (byRule[rule] || 0) + 1;
          revisionLog.push({
            rule: rule,
            name: issue.name,
            paraIndex: issue.index,
            original: issue.original,
            message: issue.message
          });
        } catch (e) {
          console.warn('[applyFixes] 公式居中修复失败: ' + e);
          var para = doc.Paragraphs.Item(issue.index);
          if (para) addCommentToDoc(doc, para.Range, issue);
          commented++;
          byRule[rule] = (byRule[rule] || 0) + 1;
        }
      } else if (shouldFix && rule === 'E-003') {
        // E-003: 公式编号位置修复（居中制表位 + 右对齐制表位）
        try {
          console.log('[applyFixes] 修复公式编号位置 段落 ' + issue.index);
          var para = doc.Paragraphs.Item(issue.index);
          if (!para) continue;

          var spec = issue.fixSpec || {};
          if (spec.formatWithTabStops && spec.numberText) {
            var range = para.Range;
            var text = range.Text || '';
            var formulaBody = (spec.formulaBody || '').replace(/\t/g, ' ').trim();
            var numberText = (spec.numberText || '').trim();

            if (formulaBody && numberText) {
              var pageWidth = doc.PageSetup.PageWidth;
              var leftMargin = doc.PageSetup.LeftMargin;
              var rightMargin = doc.PageSetup.RightMargin;
              var contentWidth = pageWidth - leftMargin - rightMargin;
              var centerPos = contentWidth / 2;
              var rightPos = contentWidth;
              var paraMark = /\r$/.test(text) ? '\r' : '';
              var newText = '\t' + formulaBody + '\t' + numberText + paraMark;

              if (newText !== text) {
                range.Text = newText;
              }

              para = doc.Paragraphs.Item(issue.index);
              if (para && para.Format && para.Format.Alignment !== undefined) {
                para.Format.Alignment = 0;
              }

              try {
                if (para && para.Format && para.Format.TabStops) {
                  para.Format.TabStops.ClearAll();
                  // 1 = wdAlignTabCenter, 2 = wdAlignTabRight, 0 = wdTabLeaderSpaces
                  para.Format.TabStops.Add(centerPos, 1, 0);
                  para.Format.TabStops.Add(rightPos, 2, 0);
                }
              } catch (tabErr) {
                console.warn('[applyFixes] 设置公式制表位失败: ' + tabErr);
              }

              console.log('[applyFixes] 已设置居中制表位和右对齐制表位');
            }
          }

          fixed++;
          byRule[rule] = (byRule[rule] || 0) + 1;
          revisionLog.push({
            rule: rule,
            name: issue.name,
            paraIndex: issue.index,
            original: issue.original,
            message: issue.message
          });
        } catch (e) {
          console.warn('[applyFixes] 公式编号位置修复失败: ' + e);
          var para = doc.Paragraphs.Item(issue.index);
          if (para) addCommentToDoc(doc, para.Range, issue);
          commented++;
          byRule[rule] = (byRule[rule] || 0) + 1;
        }
      } else if (shouldFix && (rule.indexOf('F-') === 0 || rule.indexOf('G-') === 0)) {
        // F-001~F-005 格式修复 + G-002~G-004 图名图片格式修复
        try {
          console.log('[applyFixes] 修复格式 段落 ' + issue.index + ': ' + rule);

          var para = doc.Paragraphs.Item(issue.index);
          if (!para) continue;

          var spec = issue.fixSpec || {};
          var range = para.Range;

          // 设置字体
          if (spec.font && range && range.Font) {
            range.Font.NameFarEast = spec.font;
            range.Font.Name = spec.font;
            console.log('[applyFixes] 设置字体: ' + spec.font);
          }

          // 设置字号
          if (spec.size && range && range.Font) {
            range.Font.Size = spec.size;
            console.log('[applyFixes] 设置字号: ' + spec.size + '磅');
          }

          // 设置加粗
          if (spec.bold !== undefined && range && range.Font) {
            range.Font.Bold = spec.bold ? -1 : 0;
            console.log('[applyFixes] 设置加粗: ' + spec.bold);
          }

          // 设置行距
          if (para.Format && para.Format.LineSpacingRule !== undefined) {
            if (spec.lineSpacingRule !== undefined) {
              // 使用行距规则（如单倍行距）
              para.Format.LineSpacingRule = spec.lineSpacingRule;
              console.log('[applyFixes] 设置行距规则: ' + (spec.lineSpacingRule === 0 ? '单倍行距' : spec.lineSpacingRule));
            } else if (spec.lineSpacing !== undefined) {
              // 使用固定行距
              para.Format.LineSpacingRule = 4;  // wdLineSpaceExactly
              para.Format.LineSpacing = spec.lineSpacing;
              console.log('[applyFixes] 设置行距: ' + spec.lineSpacing + '磅');
            }
          }

          // 设置段间距
          if (para.Format) {
            if (spec.spaceBefore !== undefined && para.Format.SpaceBefore !== undefined) {
              para.Format.SpaceBefore = spec.spaceBefore;
              console.log('[applyFixes] 设置段前间距: ' + spec.spaceBefore + '磅');
            }
            if (spec.spaceAfter !== undefined && para.Format.SpaceAfter !== undefined) {
              para.Format.SpaceAfter = spec.spaceAfter;
              console.log('[applyFixes] 设置段后间距: ' + spec.spaceAfter + '磅');
            }
          }

          // 设置首行缩进
          if (spec.firstIndent !== undefined && para.Format && para.Format.FirstLineIndent !== undefined) {
            para.Format.FirstLineIndent = spec.firstIndent;
            console.log('[applyFixes] 设置首行缩进: ' + spec.firstIndent + '磅');
          }

          // 设置对齐方式
          if (spec.alignment !== undefined && para.Format && para.Format.Alignment !== undefined) {
            para.Format.Alignment = spec.alignment;
            var alignName = spec.alignment === 0 ? '左对齐' : (spec.alignment === 1 ? '居中' : '右对齐');
            console.log('[applyFixes] 设置对齐方式: ' + alignName);
          }

          // 修复图片环绕方式，避免被文字遮挡
          if (spec.fixImageWrap) {
            console.log('[applyFixes] 开始处理图片环绕方式, fixImageWrap=' + spec.fixImageWrap);
            try {
              // 检查是否有嵌入式图片
              var inlineCount = range.InlineShapes ? range.InlineShapes.Count : 0;
              // 检查是否有浮动图片
              var shapeCount = 0;
              try {
                shapeCount = range.ShapeRange ? range.ShapeRange.Count : 0;
              } catch (e) {
                shapeCount = 0;
              }
              console.log('[applyFixes] 嵌入式图片数=' + inlineCount + ', 浮动图片数=' + shapeCount);

              // 处理浮动图片（ShapeRange）- 转换为嵌入式或设置环绕方式
              if (shapeCount > 0) {
                for (var k = 1; k <= shapeCount; k++) {
                  try {
                    var shape = range.ShapeRange.Item(k);
                    // 方法1: 尝试转换为嵌入式图片（最安全）
                    try {
                      shape.ConvertToInlineShape();
                      console.log('[applyFixes] 浮动图片已转换为嵌入式');
                    } catch (convertErr) {
                      // 方法2: 如果转换失败，设置上下型环绕
                      if (shape.WrapFormat) {
                        shape.WrapFormat.Type = 4;  // wdWrapTopBottom
                        console.log('[applyFixes] 设置图片环绕方式: 上下型');
                      }
                    }
                  } catch (shapeErr) {
                    console.warn('[applyFixes] 处理单个图片失败: ' + shapeErr);
                  }
                }
              } else {
                console.log('[applyFixes] 无浮动图片需要处理');
              }
            } catch (e) {
              console.warn('[applyFixes] 设置图片环绕方式失败: ' + e);
            }
          }

          // 删除开头空白字符（不破坏图片）
          if (spec.removeLeadingSpaces && range) {
            var originalText = range.Text || '';
            var leadingMatch = originalText.match(/^([ \t　\u00A0\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200A\u202F\u205F\u3000]+)/);
            if (leadingMatch) {
              var spaceLen = leadingMatch[1].length;
              // 只删除开头的空白字符，保留后面内容（包括图片）
              var delRange = para.Range;
              delRange.SetRange(delRange.Start, delRange.Start + spaceLen);
              delRange.Delete();
              console.log('[applyFixes] 删除开头空白字符: ' + spaceLen + '个');
            }
          }

          fixed++;
          byRule[rule] = (byRule[rule] || 0) + 1;
          console.log('[applyFixes] 格式修复成功');

          revisionLog.push({
            rule: rule,
            name: issue.name,
            paraIndex: issue.index,
            original: issue.original,
            message: issue.message
          });
        } catch (e) {
          console.warn('[applyFixes] 格式修复失败: ' + e);
          addCommentToDoc(doc, range, issue);
          commented++;
          byRule[rule] = (byRule[rule] || 0) + 1;
        }
      } else {
        // 添加批注
        addCommentToDoc(doc, range, issue);
        commented++;
        byRule[rule] = (byRule[rule] || 0) + 1;
      }
    } catch (e) {}
  }

  // 处理页眉页脚等 section 级问题（HF-001~003）
  for (var s = 0; s < sectionIssues.length; s++) {
    var sIssue = sectionIssues[s];
    var sRule = sIssue.rule || '';
    var sCanAutoFix = sIssue.autoFix === true;

    var sShouldFix = false;
    if (mode === 'aggressive') {
      sShouldFix = sCanAutoFix && AUTO_FIX_RULES[sRule];
    } else if (mode === 'standard') {
      sShouldFix = sCanAutoFix && AUTO_FIX_RULES[sRule] &&
        (sRule === 'HF-001' || sRule === 'HF-002' || sRule === 'HF-003' || sRule === 'PG-001' || sRule === 'PG-002');
    }

    try {
      if (sShouldFix && sRule === 'HF-001') {
        console.log('[applyFixes] 修复页眉页脚字号');
        var sSpec = sIssue.fixSpec || {};
        var sSectionIndex = sSpec.sectionIndex;
        var sFontSize = sSpec.fontSize || 9;
        var sFontNameCn = sSpec.fontNameCn || '宋体';
        var sFontNameEn = sSpec.fontNameEn || 'Arial';
        var sType = sSpec.type;

        if (sSectionIndex) {
          var sSection = doc.Sections.Item(sSectionIndex);
          if (sSection) {
            if (sType === 'header') {
              var sHeader = sSection.Headers.Item(1);
              var sHeaderParas = getHeaderFooterEffectiveParagraphs(sHeader);
              for (var shp = 0; shp < sHeaderParas.length; shp++) {
                var sHeaderPara = sHeaderParas[shp];
                if (sHeaderPara && sHeaderPara.Range && sHeaderPara.Range.Font) {
                  sHeaderPara.Range.Font.NameFarEast = sFontNameCn;
                  sHeaderPara.Range.Font.Name = sFontNameEn;
                  sHeaderPara.Range.Font.Size = sFontSize;
                }
              }
              console.log('[applyFixes] 设置页眉字体字号为 ' + sFontNameCn + '/' + sFontNameEn + '/' + sFontSize + '磅');
            } else if (sType === 'footer') {
              var sFooter = sSection.Footers.Item(1);
              var sFooterParas = getHeaderFooterEffectiveParagraphs(sFooter);
              for (var sfp = 0; sfp < sFooterParas.length; sfp++) {
                var sFooterPara = sFooterParas[sfp];
                if (sFooterPara && sFooterPara.Range && sFooterPara.Range.Font) {
                  sFooterPara.Range.Font.NameFarEast = sFontNameCn;
                  sFooterPara.Range.Font.Name = sFontNameEn;
                  sFooterPara.Range.Font.Size = sFontSize;
                }
              }
              console.log('[applyFixes] 设置页脚字体字号为 ' + sFontNameCn + '/' + sFontNameEn + '/' + sFontSize + '磅');
            }

            fixed++;
            byRule[sRule] = (byRule[sRule] || 0) + 1;
            revisionLog.push({
              rule: sRule,
              name: sIssue.name,
              sectionIndex: sSectionIndex,
              original: sIssue.original,
              suggested: sIssue.suggested,
              message: sIssue.message
            });
          }
        }
      } else if (sShouldFix && (sRule === 'HF-002' || sRule === 'HF-003')) {
        console.log('[applyFixes] 修复' + (sRule === 'HF-002' ? '页眉线' : '页脚线'));
        var sSpec = sIssue.fixSpec || {};
        var sSectionIndex = sSpec.sectionIndex;
        var sType = sSpec.type;
        var sLineStyle = sSpec.lineStyle || 1;
        var sLineWidth = sSpec.lineWidth || 4;

        if (sSectionIndex) {
          var sSection = doc.Sections.Item(sSectionIndex);
          if (sSection) {
            if (sType === 'header_border') {
              var sHeader = sSection.Headers.Item(1);
              if (sHeader && sHeader.Range) {
                var sHeaderParas = getHeaderFooterEffectiveParagraphs(sHeader);
                clearHeaderFooterParagraphBorders(sHeaderParas);
                var sHeaderPara = getHeaderFooterBorderParagraph(sHeader, true);
                if (applyHeaderFooterLine(sHeaderPara, -3, sLineStyle, sLineWidth)) {
                  console.log('[applyFixes] 设置页眉线: border=bottom, style=' + sLineStyle + ', width=' + sLineWidth);
                }
              }
            } else if (sType === 'footer_border') {
              var sFooter = sSection.Footers.Item(1);
              if (sFooter && sFooter.Range) {
                var sFooterParas = getHeaderFooterEffectiveParagraphs(sFooter);
                clearHeaderFooterParagraphBorders(sFooterParas);
                var sFooterPara = getHeaderFooterBorderParagraph(sFooter, false);
                if (applyHeaderFooterLine(sFooterPara, -1, sLineStyle, sLineWidth)) {
                  console.log('[applyFixes] 设置页脚线: border=top, style=' + sLineStyle + ', width=' + sLineWidth);
                }
              }
            }

            fixed++;
            byRule[sRule] = (byRule[sRule] || 0) + 1;
            revisionLog.push({
              rule: sRule,
              name: sIssue.name,
              sectionIndex: sSectionIndex,
              original: sIssue.original,
              suggested: sIssue.suggested,
              message: sIssue.message
            });
          }
        }
      } else if (sShouldFix && (sRule === 'PG-001' || sRule === 'PG-002')) {
        // PG-001, PG-002: 页面设置修复
        console.log('[applyFixes] 修复页面设置: ' + sRule);
        var pgSpec = sIssue.fixSpec || {};
        var pgSectionIndex = sIssue.sectionIndex || pgSpec.sectionIndex;

        if (pgSectionIndex) {
          var pgSection = doc.Sections.Item(pgSectionIndex);
          if (pgSection && pgSection.PageSetup) {
            var pgPageSetup = pgSection.PageSetup;

            if (sRule === 'PG-001') {
              // PG-001: 页边距
              if (pgSpec.type === 'page_margin_vertical') {
                pgPageSetup.TopMargin = pgSpec.topMargin || 72;
                pgPageSetup.BottomMargin = pgSpec.bottomMargin || 72;
                console.log('[applyFixes] 设置页边距(上下)为 72磅(2.54cm)');
              } else if (pgSpec.type === 'page_margin_horizontal') {
                pgPageSetup.LeftMargin = pgSpec.leftMargin || 90;
                pgPageSetup.RightMargin = pgSpec.rightMargin || 90;
                console.log('[applyFixes] 设置页边距(左右)为 90磅(3.17cm)');
              }
            } else if (sRule === 'PG-002') {
              // PG-002: 页眉页脚距边界
              pgPageSetup.HeaderDistance = pgSpec.headerDistance || 42.5;
              pgPageSetup.FooterDistance = pgSpec.footerDistance || 42.5;
              console.log('[applyFixes] 设置页眉页脚距边界为 42.5磅(1.5cm)');
            }

            fixed++;
            byRule[sRule] = (byRule[sRule] || 0) + 1;
            revisionLog.push({
              rule: sRule,
              name: sIssue.name,
              sectionIndex: pgSectionIndex,
              original: sIssue.original,
              suggested: sIssue.suggested,
              message: sIssue.message
            });
          }
        }
      }
    } catch (e) {
      console.warn('[applyFixes] 页眉页脚/页面设置修复失败(' + sRule + '): ' + e);
      commented++;
      byRule[sRule] = (byRule[sRule] || 0) + 1;
    }
  }

  // 处理表格问题（T-004, T-005, T-006, T-007）
  for (var j = 0; j < tableIssues.length; j++) {
    var tIssue = tableIssues[j];
    var tRule = tIssue.rule || '';
    var tCanAutoFix = tIssue.autoFix === true;

    var tShouldFix = false;
    if (mode === 'aggressive') {
      tShouldFix = tCanAutoFix && AUTO_FIX_RULES[tRule];
    } else if (mode === 'standard') {
      tShouldFix = tCanAutoFix && AUTO_FIX_RULES[tRule];
    }

    try {
      var targetTable = doc.Tables.Item(tIssue.tableIndex);
      if (!targetTable) continue;

      // T-004: 表格宽度
      if (tRule === 'T-004' && tShouldFix) {
        var spec = tIssue.fixSpec || {};
        if (spec.preferredWidth !== undefined) {
          targetTable.PreferredWidthType = spec.preferredWidthType || 3;  // wdPreferredWidthPoints
          targetTable.PreferredWidth = spec.preferredWidth;
          console.log('[applyFixes] 设置表格 ' + tIssue.tableIndex + ' 宽度为 ' + spec.preferredWidth + '磅');
          fixed++;
          byRule[tRule] = (byRule[tRule] || 0) + 1;
          revisionLog.push({
            rule: tRule,
            name: tIssue.name,
            tableIndex: tIssue.tableIndex,
            original: tIssue.original,
            suggested: tIssue.suggested,
            message: tIssue.message
          });
        }
        continue;
      }

      // T-007: 跨页重复表头
      if (tRule === 'T-007' && tShouldFix) {
        var rows = targetTable.Rows;
        if (rows.Count > 0) {
          var firstRow = rows.Item(1);
          firstRow.HeadingFormat = -1;  // -1 表示标题行会重复
          console.log('[applyFixes] 设置表格 ' + tIssue.tableIndex + ' 首行为重复表头');
          fixed++;
          byRule[tRule] = (byRule[tRule] || 0) + 1;
          revisionLog.push({
            rule: tRule,
            name: tIssue.name,
            tableIndex: tIssue.tableIndex,
            original: tIssue.original,
            message: tIssue.message
          });
        }
        continue;
      }

      // T-005, T-006: 表格内容（需要 rowIndex 和 colIndex）
      if (!tIssue.rowIndex || !tIssue.colIndex) continue;

      var targetCell = targetTable.Cell(tIssue.rowIndex, tIssue.colIndex);
      if (!targetCell || !targetCell.Range) continue;

      if (tShouldFix && tIssue.suggested) {
        // 替换单元格内容
        var cellRange = targetCell.Range;
        var newText = String(tIssue.suggested);

        console.log('[applyFixes] 替换表格[' + tIssue.tableIndex + '](' + tIssue.rowIndex + ',' + tIssue.colIndex + '): ' + tRule);
        console.log('[applyFixes] "' + tIssue.original + '" → "' + newText + '"');

        cellRange.Text = newText;
        fixed++;
        byRule[tRule] = (byRule[tRule] || 0) + 1;
        revisionLog.push({
          rule: tRule,
          name: tIssue.name,
          tableIndex: tIssue.tableIndex,
          rowIndex: tIssue.rowIndex,
          colIndex: tIssue.colIndex,
          original: tIssue.original,
          suggested: tIssue.suggested,
          message: tIssue.message
        });
      } else {
        // 添加批注到单元格
        addCommentToDoc(doc, targetCell.Range, tIssue);
        commented++;
        byRule[tRule] = (byRule[tRule] || 0) + 1;
      }
    } catch (e) {
      console.warn('[applyFixes] 表格处理失败: ' + e);
    }
  }

  return { fixed: fixed, commented: commented, byRule: byRule, revisionLog: revisionLog };
}

function addCommentToDoc(doc, range, issue) {
  try {
    var commentText = '[' + issue.rule + '] ' + (issue.name || '问题');
    if (issue.message) {
      commentText += '：' + issue.message;
    }
    if (issue.suggested) {
      commentText += '\n建议修改为：' + issue.suggested;
    }

    doc.Comments.Add(range, commentText);
  } catch (e) {}
}

// ========== 编号校对专用函数 ==========

/**
 * 编号校对：一次遍历完成所有操作
 * 支持一级标题（第X章）和二级标题（X.Y）
 */
function scanStructureForNumbering(doc, scope) {
  try {
    console.log('[scanStructureForNumbering] 开始，scope=' + scope);
    var origTrack = doc.TrackRevisions;
    var needHeading = scope === 'numbering' || scope === 'heading';
    var needFigure = scope === 'numbering' || scope === 'figure';
    var needTable = scope === 'numbering' || scope === 'table';
    var needFormula = scope === 'numbering' || scope === 'formula';
    var details = [];

    var cn2num = {
      '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
      '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
      '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15,
      '十六': 16, '十七': 17, '十八': 18, '十九': 19, '二十': 20,
      '二十一': 21, '二十二': 22, '二十三': 23, '二十四': 24, '二十五': 25,
      '二十六': 26, '二十七': 27, '二十八': 28, '二十九': 29, '三十': 30
    };
    var num2cn = {
      1: '一', 2: '二', 3: '三', 4: '四', 5: '五',
      6: '六', 7: '七', 8: '八', 9: '九', 10: '十',
      11: '十一', 12: '十二', 13: '十三', 14: '十四', 15: '十五',
      16: '十六', 17: '十七', 18: '十八', 19: '十九', 20: '二十',
      21: '二十一', 22: '二十二', 23: '二十三', 24: '二十四', 25: '二十五',
      26: '二十六', 27: '二十七', 28: '二十八', 29: '二十九', 30: '三十'
    };

    function parseCN(cn) { return cn2num[cn] || 0; }
    function toCN(num) { return num2cn[num] || String(num); }
    function cleanText(text) { return String(text || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim(); }
    function normalizeFormulaSuffix(text) {
      return String(text || '').replace(/[\(（][A-Z]?\d+(?:\.\d+){0,3}(?:\s*[-－—]\s*\d+)?[\)）]\s*$/, '');
    }
    function pushPlan(plans, oldText, newText, rule) {
      if (!oldText || !newText || oldText === newText) return;
      plans.push({ oldText: oldText, newText: newText, rule: rule });
    }
    function dedupePlans(plans) {
      var result = [];
      var seen = {};
      for (var di = 0; di < plans.length; di++) {
        var key = plans[di].rule + '||' + plans[di].oldText + '||' + plans[di].newText;
        if (seen[key]) continue;
        seen[key] = true;
        result.push(plans[di]);
      }
      return result;
    }
    function appendixLetterFromIndex(index) { return String.fromCharCode(64 + index); }
    function getCurrentFormulaAnchor() {
      var parts = [String(currentChapter)];
      if (currentSection > 0) parts.push(String(currentSection));
      if (currentSubsection > 0) parts.push(String(currentSubsection));
      if (currentItem > 0) parts.push(String(currentItem));
      return parts.join('.');
    }

    var docText = doc.Content && doc.Content.Text ? String(doc.Content.Text) : '';
    var paras = docText.split('\r');
    var plans = [];
    var counts = { headings: 0, figures: 0, tables: 0, formulas: 0 };
    var currentChapter = 1;
    var currentSection = 0;
    var currentSubsection = 0;
    var currentItem = 0;
    var expectedChapter = 0;
    var expectedSection = 0;
    var expectedSubsection = 0;
    var expectedItem = 0;
    var figureCounters = {};
    var tableCounters = {};
    var formulaCounters = {};
    var inAppendix = false;
    var appendixIndex = 0;
    var currentAppendix = '';
    var appendixTitle1 = 0;
    var appendixTitle2 = 0;
    var appendixTitle3 = 0;
    var appendixFigureCounter = 0;
    var appendixTableCounter = 0;
    var appendixFormulaCounter = 0;
    var attachedTableCounter = 0;

    function resetForChapter() {
      expectedSection = 0;
      expectedSubsection = 0;
      expectedItem = 0;
      currentSection = 0;
      currentSubsection = 0;
      currentItem = 0;
    }
    function resetForSection() {
      expectedSubsection = 0;
      expectedItem = 0;
      currentSubsection = 0;
      currentItem = 0;
    }
    function resetForSubsection() {
      expectedItem = 0;
      currentItem = 0;
    }
    function resetAppendixCounters() {
      appendixTitle1 = 0;
      appendixTitle2 = 0;
      appendixTitle3 = 0;
      appendixFigureCounter = 0;
      appendixTableCounter = 0;
      appendixFormulaCounter = 0;
    }

    for (var i = 0; i < paras.length; i++) {
      var text = cleanText(paras[i]);
      if (!text) continue;

      var appendixMatch = text.match(/^附录\s*([A-Z一二三四五六七八九十]?)[\s　]*(.*)$/i);
      if (appendixMatch && appendixMatch[1]) {
        appendixIndex++;
        currentAppendix = appendixLetterFromIndex(appendixIndex);
        inAppendix = true;
        resetAppendixCounters();
        counts.headings++;
        if (needHeading) pushPlan(plans, text, '附录 ' + currentAppendix + (appendixMatch[2] ? ' ' + appendixMatch[2] : ''), 'N-007');
        continue;
      }

      if (inAppendix) {
        var appendixChapterMatch = text.match(/^第([一二三四五六七八九十]+)章\s*(.*)$/);
        if (appendixChapterMatch) {
          inAppendix = false;
        }
      }

      var m1 = text.match(/^第([一二三四五六七八九十]+)章\s*(.*)$/);
      if (m1) {
        expectedChapter++;
        currentChapter = expectedChapter;
        resetForChapter();
        counts.headings++;
        if (needHeading) pushPlan(plans, text, '第' + toCN(expectedChapter) + '章 ' + m1[2], 'N-002');
        continue;
      }

      if (inAppendix) {
        var appM1 = text.match(/^(?:[A-Z])?(\d+)\s+(.+)$/);
        if (appM1 && text.indexOf('表') !== 0 && text.indexOf('图') !== 0) {
          appendixTitle1++;
          appendixTitle2 = 0;
          appendixTitle3 = 0;
          counts.headings++;
          if (needHeading) pushPlan(plans, text, currentAppendix + appendixTitle1 + ' ' + appM1[2], 'N-007');
          continue;
        }

        var appM2 = text.match(/^(?:[A-Z])?(\d+)\.(\d+)\s+(.+)$/);
        if (appM2 && text.indexOf('表') !== 0 && text.indexOf('图') !== 0) {
          if (appendixTitle1 <= 0) appendixTitle1 = 1;
          appendixTitle2++;
          appendixTitle3 = 0;
          counts.headings++;
          if (needHeading) pushPlan(plans, text, currentAppendix + appendixTitle1 + '.' + appendixTitle2 + ' ' + appM2[3], 'N-007');
          continue;
        }

        var appM3 = text.match(/^(?:[A-Z])?(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
        if (appM3) {
          if (appendixTitle1 <= 0) appendixTitle1 = 1;
          if (appendixTitle2 <= 0) appendixTitle2 = 1;
          appendixTitle3++;
          counts.headings++;
          if (needHeading) pushPlan(plans, text, currentAppendix + appendixTitle1 + '.' + appendixTitle2 + '.' + appendixTitle3 + ' ' + appM3[4], 'N-007');
          continue;
        }

        if (needFigure) {
          var appFigOld = text.match(/^图\s*(\d+)\s+(.+)$/);
          var appFigNew = text.match(/^图\s*([A-Z])(\d+)\s+(.+)$/);
          if (appFigOld || appFigNew) {
            appendixFigureCounter++;
            counts.figures++;
            pushPlan(plans, text, '图' + currentAppendix + appendixFigureCounter + ' ' + (appFigOld ? appFigOld[2] : appFigNew[3]), 'G-001-APP');
            continue;
          }
        }

        if (needTable) {
          var appTableOld = text.match(/^表\s*(\d+)\s+(.+)$/);
          var appTableNew = text.match(/^表\s*([A-Z])(\d+)\s+(.+)$/);
          if (appTableOld || appTableNew) {
            appendixTableCounter++;
            counts.tables++;
            pushPlan(plans, text, '表' + currentAppendix + appendixTableCounter + ' ' + (appTableOld ? appTableOld[2] : appTableNew[3]), 'T-001-APP');
            continue;
          }
        }

        if (needFormula) {
          var appendixFormulaMatch = text.match(/^(.*?)[\(（]([A-Z]?\d+(?:\.\d+){0,3}(?:\s*[-－—]\s*\d+)*)[\)）]\s*$/);
          if (appendixFormulaMatch) {
            appendixFormulaCounter++;
            counts.formulas++;
            pushPlan(plans, text, normalizeFormulaSuffix(text) + ' (' + currentAppendix + appendixFormulaCounter + ')', 'E-001-APP');
            continue;
          }
        }
      }

      var m2 = text.match(/^(\d+)\.(\d+)\s+(.+)$/);
      if (m2 && text.indexOf('表') !== 0 && text.indexOf('图') !== 0) {
        if (expectedChapter <= 0) {
          expectedChapter = 1;
          currentChapter = 1;
        }
        expectedSection++;
        currentSection = expectedSection;
        resetForSection();
        counts.headings++;
        if (needHeading) pushPlan(plans, text, currentChapter + '.' + expectedSection + ' ' + m2[3], 'N-003');
        continue;
      }

      var m3 = text.match(/^(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
      if (m3) {
        if (expectedChapter <= 0) {
          expectedChapter = 1;
          currentChapter = 1;
        }
        if (currentSection <= 0) {
          currentSection = 1;
          expectedSection = 1;
        }
        expectedSubsection++;
        currentSubsection = expectedSubsection;
        resetForSubsection();
        counts.headings++;
        if (needHeading) pushPlan(plans, text, currentChapter + '.' + currentSection + '.' + expectedSubsection + ' ' + m3[4], 'N-004');
        continue;
      }

      var m4 = text.match(/^(\d+)\.(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
      if (m4) {
        if (expectedChapter <= 0) {
          expectedChapter = 1;
          currentChapter = 1;
        }
        if (currentSection <= 0) {
          currentSection = 1;
          expectedSection = 1;
        }
        if (currentSubsection <= 0) {
          currentSubsection = 1;
          expectedSubsection = 1;
        }
        expectedItem++;
        currentItem = expectedItem;
        counts.headings++;
        if (needHeading) pushPlan(plans, text, currentChapter + '.' + currentSection + '.' + currentSubsection + '.' + expectedItem + ' ' + m4[5], 'N-005');
        continue;
      }

      if (needFigure) {
        var figOld = text.match(/^图\s*(\d+)\s+(.+)$/);
        var figNew = text.match(/^图\s*(\d+)\.(\d+)-(\d+)\s+(.+)$/);
        if (figOld || figNew) {
          var figCaption = figOld ? figOld[2] : figNew[4];
          var figKey = currentChapter + '.' + (currentSection > 0 ? currentSection : 1);
          figureCounters[figKey] = (figureCounters[figKey] || 0) + 1;
          counts.figures++;
          pushPlan(plans, text, '图' + currentChapter + '.' + (currentSection > 0 ? currentSection : 1) + '-' + figureCounters[figKey] + ' ' + figCaption, 'G-001');
          continue;
        }
      }

      if (needTable) {
        var attachedTable = text.match(/^附表\s*(\d+)\s+(.+)$/);
        if (attachedTable) {
          attachedTableCounter++;
          counts.tables++;
          pushPlan(plans, text, '附表' + attachedTableCounter + ' ' + attachedTable[2], 'T-002');
          continue;
        }

        var tableOld = text.match(/^表\s*(\d+)\s+(.+)$/);
        var tableNew = text.match(/^表\s*(\d+)\.(\d+)-(\d+)\s+(.+)$/);
        if (tableOld || tableNew) {
          var tableCaption = tableOld ? tableOld[2] : tableNew[4];
          var tableKey = currentChapter + '.' + (currentSection > 0 ? currentSection : 1);
          tableCounters[tableKey] = (tableCounters[tableKey] || 0) + 1;
          counts.tables++;
          pushPlan(plans, text, '表' + currentChapter + '.' + (currentSection > 0 ? currentSection : 1) + '-' + tableCounters[tableKey] + ' ' + tableCaption, 'T-001');
          continue;
        }
      }

      if (needFormula) {
        var formulaMatch = text.match(/^(.*?)[\(（](\d+(?:\.\d+){0,3})(?:\s*[-－—]\s*(\d+))?[\)）]\s*$/);
        if (formulaMatch) {
          var formulaPrefix = normalizeFormulaSuffix(text);
          var formulaKey = getCurrentFormulaAnchor();
          formulaCounters[formulaKey] = (formulaCounters[formulaKey] || 0) + 1;
          counts.formulas++;
          pushPlan(plans, text, formulaPrefix + ' (' + formulaKey + '-' + formulaCounters[formulaKey] + ')', 'E-001');
          continue;
        }
      }
    }

    plans = dedupePlans(plans);
    doc.TrackRevisions = true;

    var totalFixed = 0;
    var cursor = 0;
    var docEnd = doc.Content.End;

    for (i = 0; i < plans.length; i++) {
      var plan = plans[i];
      try {
        var searchRange = doc.Range(cursor, docEnd);
        searchRange.Find.ClearFormatting();
        searchRange.Find.Forward = true;
        searchRange.Find.Wrap = 0;
        searchRange.Find.MatchWildcards = false;
        var found = searchRange.Find.Execute(plan.oldText, false, false, false, false, false, true, 1, false);
        if (found) {
          var foundStart = searchRange.Start;
          var foundEnd = searchRange.End;
          searchRange.Text = plan.newText;
          totalFixed++;
          cursor = foundStart + String(plan.newText).length;
          if (cursor < foundEnd) cursor = foundEnd;
          details.push({ rule: plan.rule, original: plan.oldText, suggested: plan.newText });
        }
      } catch (findErr) {}
    }

    doc.TrackRevisions = origTrack;

    console.log('[scanStructureForNumbering] 完成：标题' + counts.headings + ' 图' + counts.figures + ' 表' + counts.tables + ' 公式' + counts.formulas + '，修复' + totalFixed + '处');
    return {
      success: true,
      totalFixed: totalFixed,
      details: details,
      structure: { headings: counts.headings, figures: counts.figures, tables: counts.tables, formulas: counts.formulas },
      fixPlan: { headingFixes: [], figureFixes: [], tableFixes: [], formulaFixes: [] },
      summary: { totalIssues: totalFixed }
    };
  } catch (e) {
    console.warn('[scanStructureForNumbering] 错误: ' + e);
    return { success: false, error: String(e) };
  }
}
