/**
 * scan-structure.js
 * 全文规划标题/图/表/公式编号，然后按顺序精确替换，避免段落索引错位。
 */
try {
  var doc = Application.ActiveDocument;
  if (!doc) return { success: false, error: '没有打开的文档' };

  var scopeType = typeof scope !== 'undefined' ? scope : 'numbering';
  var needHeading = scopeType === 'numbering' || scopeType === 'heading';
  var needFigure = scopeType === 'numbering' || scopeType === 'figure';
  var needTable = scopeType === 'numbering' || scopeType === 'table';
  var needFormula = scopeType === 'numbering' || scopeType === 'formula';

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

  function parseCN(cn) {
    return cn2num[cn] || 0;
  }

  function toCN(num) {
    return num2cn[num] || String(num);
  }

  function cleanText(text) {
    return String(text || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function normalizeText(text) {
    return cleanText(text).replace(/\s+/g, ' ');
  }

  function normalizeFormulaSuffix(text) {
    return String(text || '').replace(/[\(（][A-Z]?\d+(?:\.\d+){0,3}(?:\s*[-－—]\s*\d+)?[\)）]\s*$/, '');
  }

  function pushPlan(plans, oldText, newText, rule) {
    if (!oldText || !newText || oldText === newText) return;
    plans.push({
      oldText: oldText,
      newText: newText,
      rule: rule
    });
  }

  function dedupePlans(plans) {
    var result = [];
    var seen = {};
    for (var i = 0; i < plans.length; i++) {
      var key = plans[i].rule + '||' + plans[i].oldText + '||' + plans[i].newText;
      if (seen[key]) continue;
      seen[key] = true;
      result.push(plans[i]);
    }
    return result;
  }

  function appendixLetterFromIndex(index) {
    return String.fromCharCode(64 + index);
  }

  function getCurrentFormulaAnchor() {
    var parts = [String(currentChapter)];
    if (currentSection > 0) parts.push(String(currentSection));
    if (currentSubsection > 0) parts.push(String(currentSubsection));
    if (currentItem > 0) parts.push(String(currentItem));
    return parts.join('.');
  }

  var docText = doc.Content && doc.Content.Text ? String(doc.Content.Text) : '';
  var paras = docText.split('\r');
  var totalParas = paras.length;

  console.log('[scan] 开始规划，总段落数: ' + totalParas + ', scope=' + scopeType);

  var plans = [];
  var counts = { headings: 0, figures: 0, tables: 0, formulas: 0 };

  var currentChapter = 1;
  var currentSection = 0;
  var currentSubsection = 0;
  var currentItem = 0;
  var currentSubItem = 0;
  var expectedChapter = 0;
  var expectedSection = 0;
  var expectedSubsection = 0;
  var expectedItem = 0;
  var expectedSubItem = 0;
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
    expectedSubItem = 0;
    currentItem = 0;
    currentSubItem = 0;
  }

  function resetAppendixCounters() {
    appendixTitle1 = 0;
    appendixTitle2 = 0;
    appendixTitle3 = 0;
    appendixFigureCounter = 0;
    appendixTableCounter = 0;
    appendixFormulaCounter = 0;
  }

  // 辅助函数：检查是否是中文标题（排除表格中的数字+单位格式）
  function isChineseTitle(str) {
    if (!str) return false;
    // 标题必须以中文字符开头
    return /^[\u4e00-\u9fa5]/.test(str);
  }

  for (var i = 0; i < totalParas; i++) {
    var text = cleanText(paras[i]);
    if (!text) continue;

    var appendixMatch = text.match(/^附\s*录\s*([A-Z一二三四五六七八九十]?)[\s　]*(.*)$/i);
    if (appendixMatch && appendixMatch[1]) {
      appendixIndex++;
      currentAppendix = appendixLetterFromIndex(appendixIndex);
      inAppendix = true;
      resetAppendixCounters();
      counts.headings++;
      console.log('[scan] 检测到附录: ' + text + ' → 附录 ' + currentAppendix);
      if (needHeading) {
        pushPlan(plans, text, '附录 ' + currentAppendix + (appendixMatch[2] ? ' ' + appendixMatch[2] : ''), 'N-007');
      }
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
      if (needHeading) {
        pushPlan(plans, text, '第' + toCN(expectedChapter) + '章 ' + m1[2], 'N-002');
      }
      continue;
    }

    // 识别数字格式的一级标题：如 "3 系统范围"、"5 系统设计"
    // 特征：单个数字开头，后面是空格和中文标题
    var m1num = text.match(/^(\d+)\s+([^\d\s].*)$/);
    if (m1num && !text.match(/^\d+\.\d/) && isChineseTitle(m1num[2])) {
      // 这是数字格式的一级标题
      var detectedChapter = parseInt(m1num[1], 10);
      if (detectedChapter > 0) {
        // 如果检测到的章节号比预期大1以内，认为是正确的顺序
        // 如果跳跃较大，则按顺序修正
        if (detectedChapter === expectedChapter + 1 || expectedChapter === 0) {
          expectedChapter = detectedChapter;
        } else if (detectedChapter > expectedChapter + 1) {
          // 跳跃了，按顺序修正
          expectedChapter++;
        } else if (detectedChapter <= expectedChapter) {
          // 编号重复或倒退，修正为预期值
          expectedChapter++;
        }
        currentChapter = expectedChapter;
        resetForChapter();
        counts.headings++;
        if (needHeading) {
          pushPlan(plans, text, expectedChapter + ' ' + m1num[2], 'N-002');
        }
        continue;
      }
    }

    if (inAppendix) {
      // 支持两种格式：A1 标题 或 A.1 标题
      var appM1 = text.match(/^(?:[A-Z]\.?)?(\d+)\s+([^\d].*)$/);
      if (appM1 && text.indexOf('表') !== 0 && text.indexOf('图') !== 0 && isChineseTitle(appM1[2])) {
        appendixTitle1++;
        appendixTitle2 = 0;
        appendixTitle3 = 0;
        counts.headings++;
        console.log('[scan] 附录一级标题: ' + text + ' → ' + currentAppendix + appendixTitle1 + ' ' + appM1[2]);
        if (needHeading) {
          pushPlan(plans, text, currentAppendix + appendixTitle1 + ' ' + appM1[2], 'N-007');
        }
        continue;
      }

      var appM2 = text.match(/^(?:[A-Z]\.?)?(\d+)\.(\d+)\s+(.+)$/);
      if (appM2 && text.indexOf('表') !== 0 && text.indexOf('图') !== 0 && isChineseTitle(appM2[3])) {
        if (appendixTitle1 <= 0) appendixTitle1 = 1;
        appendixTitle2++;
        appendixTitle3 = 0;
        counts.headings++;
        console.log('[scan] 附录二级标题: ' + text + ' → ' + currentAppendix + appendixTitle1 + '.' + appendixTitle2 + ' ' + appM2[3]);
        if (needHeading) {
          pushPlan(plans, text, currentAppendix + appendixTitle1 + '.' + appendixTitle2 + ' ' + appM2[3], 'N-007');
        }
        continue;
      }

      var appM3 = text.match(/^(?:[A-Z]\.?)?(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
      if (appM3 && isChineseTitle(appM3[4])) {
        if (appendixTitle1 <= 0) appendixTitle1 = 1;
        if (appendixTitle2 <= 0) appendixTitle2 = 1;
        appendixTitle3++;
        counts.headings++;
        console.log('[scan] 附录三级标题: ' + text + ' → ' + currentAppendix + appendixTitle1 + '.' + appendixTitle2 + '.' + appendixTitle3 + ' ' + appM3[4]);
        if (needHeading) {
          pushPlan(plans, text, currentAppendix + appendixTitle1 + '.' + appendixTitle2 + '.' + appendixTitle3 + ' ' + appM3[4], 'N-007');
        }
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
    if (m2 && text.indexOf('表') !== 0 && text.indexOf('图') !== 0 && isChineseTitle(m2[3])) {
      if (expectedChapter <= 0) {
        expectedChapter = 1;
        currentChapter = 1;
      }
      expectedSection++;
      currentSection = expectedSection;
      resetForSection();
      counts.headings++;
      if (needHeading) {
        pushPlan(plans, text, currentChapter + '.' + expectedSection + ' ' + m2[3], 'N-003');
      }
      continue;
    }

    var m3 = text.match(/^(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
    if (m3 && isChineseTitle(m3[4])) {
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
      if (needHeading) {
        pushPlan(plans, text, currentChapter + '.' + currentSection + '.' + expectedSubsection + ' ' + m3[4], 'N-004');
      }
      continue;
    }

    var m4 = text.match(/^(\d+)\.(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
    if (m4 && isChineseTitle(m4[5])) {
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
      if (needHeading) {
        pushPlan(plans, text, currentChapter + '.' + currentSection + '.' + currentSubsection + '.' + expectedItem + ' ' + m4[5], 'N-005');
      }
      continue;
    }

    // 五级标题：1.1.1.1.1 标题
    var m5 = text.match(/^(\d+)\.(\d+)\.(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
    if (m5 && isChineseTitle(m5[6])) {
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
      if (currentItem <= 0) {
        currentItem = 1;
        expectedItem = 1;
      }
      expectedSubItem++;
      currentSubItem = expectedSubItem;
      counts.headings++;
      if (needHeading) {
        pushPlan(plans, text, currentChapter + '.' + currentSection + '.' + currentSubsection + '.' + currentItem + '.' + expectedSubItem + ' ' + m5[6], 'N-008');
      }
      continue;
    }

    if (needFigure) {
      var figOld = text.match(/^图\s*(\d+)\s+(.+)$/);
      var figNew = text.match(/^图\s*(\d+)\.(\d+)-(\d+)\s+(.+)$/);
      if (figOld || figNew) {
        var figCaption = figOld ? figOld[2] : figNew[4];
        var figKey = currentChapter + '.' + (currentSection > 0 ? currentSection : 1);
        figureCounters[figKey] = (figureCounters[figKey] || 0) + 1;
        var expectedFig = '图' + currentChapter + '.' + (currentSection > 0 ? currentSection : 1) + '-' + figureCounters[figKey] + ' ' + figCaption;
        counts.figures++;
        pushPlan(plans, text, expectedFig, 'G-001');
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
        var expectedTable = '表' + currentChapter + '.' + (currentSection > 0 ? currentSection : 1) + '-' + tableCounters[tableKey] + ' ' + tableCaption;
        counts.tables++;
        pushPlan(plans, text, expectedTable, 'T-001');
        continue;
      }
    }

    if (needFormula) {
      var formulaMatch = text.match(/^(.*?)[\(（](\d+(?:\.\d+){0,3})(?:\s*[-－—]\s*(\d+))?[\)）]\s*$/);
      if (formulaMatch) {
        var formulaPrefix = normalizeFormulaSuffix(text);
        var formulaKey = getCurrentFormulaAnchor();
        formulaCounters[formulaKey] = (formulaCounters[formulaKey] || 0) + 1;
        var expectedFormula = formulaPrefix + ' (' + formulaKey + '-' + formulaCounters[formulaKey] + ')';
        counts.formulas++;
        pushPlan(plans, text, expectedFormula, 'E-001');
        continue;
      }
    }
  }

  plans = dedupePlans(plans);
  console.log('[scan] 规划完成：标题' + counts.headings + ' 图' + counts.figures + ' 表' + counts.tables + ' 公式' + counts.formulas + '，待修复' + plans.length + '处');

  var origTrack = doc.TrackRevisions;
  doc.TrackRevisions = true;

  var revisionLog = [];
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
        revisionLog.push({
          rule: plan.rule,
          original: plan.oldText,
          suggested: plan.newText
        });
      }
    } catch (e) {}
  }

  doc.TrackRevisions = origTrack;

  console.log('[scan] 完成，修复: ' + totalFixed);
  return {
    success: true,
    totalFixed: totalFixed,
    fixed: totalFixed,
    details: revisionLog,
    structure: {
      headings: counts.headings,
      figures: counts.figures,
      tables: counts.tables,
      formulas: counts.formulas
    },
    summary: { totalIssues: totalFixed },
    fixPlan: { headingFixes: [], figureFixes: [], tableFixes: [], formulaFixes: [] }
  };
} catch (e) {
  console.warn('[scan] 错误: ' + e);
  return { success: false, error: String(e) };
}
