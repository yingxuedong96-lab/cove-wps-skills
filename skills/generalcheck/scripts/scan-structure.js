/**
 * scan-structure.js
 * 校对标题编号、图编号、表编号、公式编号
 * 支持：自动编号检测与转换为手动编号
 * scope: heading（标题）, figure（图）, table（表）, formula（公式）, numbering（全部）
 * figureFormat: chapter（图X.Y-Z）, simple（图1、图2...）
 * tableFormat: chapter（表X.Y-Z）, simple（表1、表2...）
 * formulaFormat: chapter（(X.Y-Z)）, simple（(1)、(2)...）
 */
try {
  var doc = Application.ActiveDocument;
  if (!doc) return { success: false, error: '没有打开的文档' };

  var scopeType = typeof scope !== 'undefined' ? scope : 'heading';
  var needHeading = scopeType === 'numbering' || scopeType === 'heading';
  var needFigure = scopeType === 'numbering' || scopeType === 'figure';
  var needTable = scopeType === 'numbering' || scopeType === 'table';
  var needFormula = scopeType === 'numbering' || scopeType === 'formula';

  var figFormat = typeof figureFormat !== 'undefined' ? figureFormat : 'chapter';
  var tblFormat = typeof tableFormat !== 'undefined' ? tableFormat : 'chapter';
  var frmFormat = typeof formulaFormat !== 'undefined' ? formulaFormat : 'chapter';

  var cn2num = { '一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10,'十一':11,'十二':12,'十三':13,'十四':14,'十五':15,'十六':16,'十七':17,'十八':18,'十九':19,'二十':20 };
  var num2cn = { 1:'一',2:'二',3:'三',4:'四',5:'五',6:'六',7:'七',8:'八',9:'九',10:'十',11:'十一',12:'十二',13:'十三',14:'十四',15:'十五',16:'十六',17:'十七',18:'十八',19:'十九',20:'二十' };

  function cleanText(text) {
    return String(text || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function isChineseTitle(str) {
    return str && /^[\u4e00-\u9fa5]/.test(str);
  }

  /**
   * 检测并转换自动编号为手动编号
   * 返回：{ isAuto: boolean, listString: string, converted: boolean }
   */
  function processAutoNumbering(para) {
    var result = { isAuto: false, listString: '', converted: false };
    try {
      var listFormat = para.Range.ListFormat;
      if (listFormat && listFormat.ListType !== 0) {
        // ListType: 0=无, 1=项目符号, 2=编号列表, 3=多级编号
        result.isAuto = true;
        result.listString = listFormat.ListString || '';

        if (result.listString) {
          // 将自动编号转为手动文本
          var rng = para.Range;
          // 在段落开头插入编号文本
          rng.InsertBefore(result.listString + ' ');
          // 清除 ListFormat（移除自动编号）
          listFormat.RemoveNumbers();
          result.converted = true;
        }
      }
    } catch (e) {
      // 某些段落可能不支持 ListFormat
    }
    return result;
  }

  console.log('[scan] 开始规划, scope=' + scopeType + ', figureFormat=' + figFormat);

  var plans = [];
  var counts = { headings: 0, figures: 0, tables: 0, formulas: 0, autoConverted: 0 };
  var autoNumberLog = [];

  var state = { ch: 0, sec: 0, sub: 0, item: 0, subItem: 0 };
  var appState = { letter: '', letterIndex: 0, l1: 0, l2: 0, l3: 0 };
  var inAppendix = false;

  var figureCounters = {};
  var simpleFigureCounter = 0;
  var appendixFigureCounter = 0;

  var tableCounters = {};
  var simpleTableCounter = 0;
  var appendixTableCounter = 0;

  var formulaCounters = {};
  var simpleFormulaCounter = 0;
  var appendixFormulaCounter = 0;

  // 开启修订模式
  var origTrack = doc.TrackRevisions;
  doc.TrackRevisions = true;

  // 遍历段落
  var paras = doc.Paragraphs;
  for (var i = 1; i <= paras.Count; i++) {
    var para = paras.Item(i);
    var rawText = para.Range.Text;
    var text = cleanText(rawText);

    // 处理自动编号
    var autoResult = processAutoNumbering(para);
    if (autoResult.isAuto) {
      counts.autoConverted++;
      autoNumberLog.push({
        original: text,
        listString: autoResult.listString,
        converted: autoResult.converted
      });
      // 更新 text（已插入编号）
      text = cleanText(para.Range.Text);
      console.log('[scan] 自动编号转换: "' + autoResult.listString + '" → "' + text.substring(0, 30) + '"');
    }

    if (!text) continue;

    // === 检测附录 ===
    var appMatch = text.match(/^附\s*录\s*([A-Z一二三四五六七八九十]+)[\s　]*(.*)$/i);
    if (appMatch) {
      inAppendix = true;
      appState.l1 = 0; appState.l2 = 0; appState.l3 = 0;
      appendixFigureCounter = 0; appendixTableCounter = 0; appendixFormulaCounter = 0;
      appState.letterIndex++;
      var letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
      appState.letter = letters.charAt((appState.letterIndex - 1) % 26);
      if (needHeading) {
        counts.headings++;
        var newText = '附录' + appState.letter + (appMatch[2] ? ' ' + appMatch[2] : '');
        if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
      }
      continue;
    }

    // === 附录内处理 ===
    if (inAppendix) {
      if (/^第[一二三四五六七八九十]+章/.test(text)) {
        inAppendix = false;
      } else {
        if (needHeading) {
          var m1 = text.match(/^[A-Z]\.(\d+)\s+(.+)$/);
          if (m1 && isChineseTitle(m1[2])) {
            appState.l1++; appState.l2 = 0; appState.l3 = 0;
            counts.headings++;
            var newText = appState.letter + '.' + appState.l1 + ' ' + m1[2];
            if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
            continue;
          }
          var m2 = text.match(/^[A-Z]\.(\d+)\.(\d+)\s+(.+)$/);
          if (m2 && isChineseTitle(m2[3])) {
            if (appState.l1 === 0) appState.l1 = 1;
            appState.l2++; appState.l3 = 0;
            counts.headings++;
            var newText = appState.letter + '.' + appState.l1 + '.' + appState.l2 + ' ' + m2[3];
            if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
            continue;
          }
          var m3 = text.match(/^[A-Z]\.(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
          if (m3 && isChineseTitle(m3[4])) {
            if (appState.l1 === 0) appState.l1 = 1;
            if (appState.l2 === 0) appState.l2 = 1;
            appState.l3++;
            counts.headings++;
            var newText = appState.letter + '.' + appState.l1 + '.' + appState.l2 + '.' + appState.l3 + ' ' + m3[4];
            if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
            continue;
          }
        }

        if (needFigure) {
          var appFig = text.match(/^图\s*([A-Z])?\.?(\d+)\s+(.+)$/i);
          if (appFig) {
            appendixFigureCounter++;
            counts.figures++;
            var newText = '图' + appState.letter + appendixFigureCounter + ' ' + appFig[3];
            if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
            continue;
          }
        }

        if (needTable) {
          var appTbl = text.match(/^表\s*([A-Z])?\.?(\d+)\s+(.+)$/i);
          if (appTbl) {
            appendixTableCounter++;
            counts.tables++;
            var newText = '表' + appState.letter + appendixTableCounter + ' ' + appTbl[3];
            if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
            continue;
          }
        }

        if (needFormula) {
          var appFrm = text.match(/\(([A-Z])?\.?(\d+)\)$/);
          if (appFrm) {
            appendixFormulaCounter++;
            counts.formulas++;
            var newText = text.replace(/\(([A-Z])?\.?(\d+)\)$/, '(' + appState.letter + appendixFormulaCounter + ')');
            if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
            continue;
          }
        }
        continue;
      }
    }

    // === 正文标题 ===
    var h1 = text.match(/^第([一二三四五六七八九十]+)章\s*(.*)$/);
    if (h1) {
      state.ch++;
      state.sec = 0; state.sub = 0; state.item = 0; state.subItem = 0;
      if (needHeading) {
        counts.headings++;
        var newText = '第' + (num2cn[state.ch] || state.ch) + '章 ' + h1[2];
        if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
      }
      continue;
    }

    var h1n = text.match(/^(\d+)\s+([^\d\s].*)$/);
    if (h1n && isChineseTitle(h1n[2]) && !text.match(/^\d+\.\d/)) {
      state.ch++;
      state.sec = 0; state.sub = 0; state.item = 0; state.subItem = 0;
      if (needHeading) {
        counts.headings++;
        var newText = state.ch + ' ' + h1n[2];
        if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
      }
      continue;
    }

    var h2 = text.match(/^(\d+)\.(\d+)\s+(.+)$/);
    if (h2 && isChineseTitle(h2[3]) && text.indexOf('表') !== 0 && text.indexOf('图') !== 0) {
      if (state.ch === 0) state.ch = 1;
      state.sec++; state.sub = 0; state.item = 0; state.subItem = 0;
      if (needHeading) {
        counts.headings++;
        var newText = state.ch + '.' + state.sec + ' ' + h2[3];
        if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
      }
      continue;
    }

    var h3 = text.match(/^(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
    if (h3 && isChineseTitle(h3[4])) {
      if (state.ch === 0) state.ch = 1;
      if (state.sec === 0) state.sec = 1;
      state.sub++; state.item = 0; state.subItem = 0;
      if (needHeading) {
        counts.headings++;
        var newText = state.ch + '.' + state.sec + '.' + state.sub + ' ' + h3[4];
        if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
      }
      continue;
    }

    var h4 = text.match(/^(\d+)\.(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
    if (h4 && isChineseTitle(h4[5])) {
      if (state.ch === 0) state.ch = 1;
      if (state.sec === 0) state.sec = 1;
      if (state.sub === 0) state.sub = 1;
      state.item++; state.subItem = 0;
      if (needHeading) {
        counts.headings++;
        var newText = state.ch + '.' + state.sec + '.' + state.sub + '.' + state.item + ' ' + h4[5];
        if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
      }
      continue;
    }

    var h5 = text.match(/^(\d+)\.(\d+)\.(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
    if (h5 && isChineseTitle(h5[6])) {
      if (state.ch === 0) state.ch = 1;
      if (state.sec === 0) state.sec = 1;
      if (state.sub === 0) state.sub = 1;
      if (state.item === 0) state.item = 1;
      state.subItem++;
      if (needHeading) {
        counts.headings++;
        var newText = state.ch + '.' + state.sec + '.' + state.sub + '.' + state.item + '.' + state.subItem + ' ' + h5[6];
        if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
      }
      continue;
    }

    // === 图编号 ===
    if (needFigure) {
      var figMatch = text.match(/^图\s*(\d+)(?:\.(\d+))?(?:-(\d+))?\s+(.+)$/);
      if (figMatch) {
        counts.figures++;
        if (figFormat === 'simple') {
          simpleFigureCounter++;
          var newText = '图' + simpleFigureCounter + ' ' + figMatch[4];
          if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
        } else {
          if (state.ch === 0) state.ch = 1;
          var figCh = state.ch;
          var figSec = state.sec > 0 ? state.sec : 1;
          var figKey = figCh + '.' + figSec;
          figureCounters[figKey] = (figureCounters[figKey] || 0) + 1;
          var figNum = figureCounters[figKey];
          var newText = '图' + figCh + '.' + figSec + '-' + figNum + ' ' + figMatch[4];
          if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
        }
        continue;
      }
    }

    // === 表编号 ===
    if (needTable) {
      var tblMatch = text.match(/^表\s*(\d+)(?:\.(\d+))?(?:-(\d+))?\s+(.+)$/);
      if (tblMatch) {
        counts.tables++;
        if (tblFormat === 'simple') {
          simpleTableCounter++;
          var newText = '表' + simpleTableCounter + ' ' + tblMatch[4];
          if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
        } else {
          if (state.ch === 0) state.ch = 1;
          var tblCh = state.ch;
          var tblSec = state.sec > 0 ? state.sec : 1;
          var tblKey = tblCh + '.' + tblSec;
          tableCounters[tblKey] = (tableCounters[tblKey] || 0) + 1;
          var tblNum = tableCounters[tblKey];
          var newText = '表' + tblCh + '.' + tblSec + '-' + tblNum + ' ' + tblMatch[4];
          if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
        }
        continue;
      }
    }

    // === 公式编号 ===
    if (needFormula) {
      var frmMatch = text.match(/\((\d+)(?:\.(\d+))?(?:-(\d+))?\)$/);
      if (frmMatch) {
        counts.formulas++;
        if (frmFormat === 'simple') {
          simpleFormulaCounter++;
          var newText = text.replace(/\((\d+)(?:\.(\d+))?(?:-(\d+))?\)$/, '(' + simpleFormulaCounter + ')');
          if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
        } else {
          if (state.ch === 0) state.ch = 1;
          var frmCh = state.ch;
          var frmSec = state.sec > 0 ? state.sec : 1;
          var frmKey = frmCh + '.' + frmSec;
          formulaCounters[frmKey] = (formulaCounters[frmKey] || 0) + 1;
          var frmNum = formulaCounters[frmKey];
          var newText = text.replace(/\((\d+)(?:\.(\d+))?(?:-(\d+))?\)$/, '(' + frmCh + '.' + frmSec + '-' + frmNum + ')');
          if (text !== newText) plans.push({ para: para, oldText: text, newText: newText });
        }
        continue;
      }
    }
  }

  console.log('[scan] 规划完成：标题' + counts.headings + ' 图' + counts.figures + ' 表' + counts.tables + ' 公式' + counts.formulas + '，自动编号转换' + counts.autoConverted + '，待修复' + plans.length + '处');

  // 执行修复
  var totalFixed = 0;
  var log = [];

  for (var p = 0; p < plans.length; p++) {
    try {
      var planPara = plans[p].para;
      var rng = planPara.Range;
      // 保留段落标记（最后一位是 \r 或 \u0007）
      var rngEnd = rng.End;
      var contentEnd = rngEnd;
      try {
        // 检查最后是否是段落标记
        var lastChar = doc.Range(rngEnd - 1, rngEnd).Text;
        if (lastChar === '\r' || lastChar === '\u0007' || lastChar.charCodeAt(0) === 13) {
          contentEnd = rngEnd - 1;
        }
      } catch (e) {}
      // 只替换内容部分，保留段落标记
      var contentRng = doc.Range(rng.Start, contentEnd);
      contentRng.Text = plans[p].newText;
      totalFixed++;
      log.push({ original: plans[p].oldText, suggested: plans[p].newText });
    } catch (e) {
      console.warn('[scan] 修复失败: ' + e);
    }
  }

  doc.TrackRevisions = origTrack;

  console.log('[scan] 完成，修复: ' + totalFixed);

  var resultMsg = '编号校对完成：修复 ' + totalFixed + ' 处';
  if (counts.autoConverted > 0) {
    resultMsg += '，转换自动编号 ' + counts.autoConverted + ' 处';
  }

  return {
    success: true,
    totalFixed: totalFixed,
    fixed: totalFixed,
    details: log,
    structure: counts,
    autoNumberConversions: autoNumberLog,
    summary: { totalIssues: totalFixed, autoNumberDetected: counts.autoConverted },
    message: resultMsg
  };
} catch (e) {
  return { success: false, error: String(e) };
}