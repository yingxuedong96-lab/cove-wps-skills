/**
 * scan-structure.js
 * 校对标题编号、图编号、表编号
 * scope: heading（标题）, figure（图）, table（表）, numbering（全部）
 * figureFormat: chapter（图X.Y-Z）, simple（图1、图2...）
 * tableFormat: chapter（表X.Y-Z）, simple（表1、表2...）
 */
try {
  var doc = Application.ActiveDocument;
  if (!doc) return { success: false, error: '没有打开的文档' };

  var scopeType = typeof scope !== 'undefined' ? scope : 'heading';
  var needHeading = scopeType === 'numbering' || scopeType === 'heading';
  var needFigure = scopeType === 'numbering' || scopeType === 'figure';
  var needTable = scopeType === 'numbering' || scopeType === 'table';

  // 图编号格式：chapter=图X.Y-Z（章节式），simple=图1、图2...（顺序式）
  var figFormat = typeof figureFormat !== 'undefined' ? figureFormat : 'chapter';
  // 表编号格式：chapter=表X.Y-Z（章节式），simple=表1、表2...（顺序式）
  var tblFormat = typeof tableFormat !== 'undefined' ? tableFormat : 'chapter';

  var cn2num = { '一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10,'十一':11,'十二':12,'十三':13,'十四':14,'十五':15,'十六':16,'十七':17,'十八':18,'十九':19,'二十':20 };
  var num2cn = { 1:'一',2:'二',3:'三',4:'四',5:'五',6:'六',7:'七',8:'八',9:'九',10:'十',11:'十一',12:'十二',13:'十三',14:'十四',15:'十五',16:'十六',17:'十七',18:'十八',19:'十九',20:'二十' };

  function cleanText(text) {
    return String(text || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function isChineseTitle(str) {
    return str && /^[\u4e00-\u9fa5]/.test(str);
  }

  var docText = doc.Content && doc.Content.Text ? String(doc.Content.Text) : '';
  var paras = docText.split('\r');
  console.log('[scan] 开始规划，总段落数: ' + paras.length + ', scope=' + scopeType + ', figureFormat=' + figFormat);

  var plans = [];
  var counts = { headings: 0, figures: 0, tables: 0 };

  // 当前活跃的各级编号状态
  var state = { ch: 0, sec: 0, sub: 0, item: 0, subItem: 0 };
  var appState = { letter: '', letterIndex: 0, l1: 0, l2: 0, l3: 0 };
  var inAppendix = false;

  // 图编号计数器
  var figureCounters = {};        // 章节式：key = "章.节"
  var simpleFigureCounter = 0;    // 顺序式：全文递增
  var appendixFigureCounter = 0;

  // 表编号计数器
  var tableCounters = {};         // 章节式：key = "章.节"
  var simpleTableCounter = 0;     // 顺序式：全文递增
  var appendixTableCounter = 0;

  for (var i = 0; i < paras.length; i++) {
    var text = cleanText(paras[i]);
    if (!text) continue;

    // === 检测附录 ===
    var appMatch = text.match(/^附\s*录\s*([A-Z一二三四五六七八九十]+)[\s　]*(.*)$/i);
    if (appMatch) {
      inAppendix = true;
      appState.l1 = 0;
      appState.l2 = 0;
      appState.l3 = 0;
      appendixFigureCounter = 0;
      appendixTableCounter = 0;
      appState.letterIndex++;
      var letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
      appState.letter = letters.charAt((appState.letterIndex - 1) % 26);
      if (needHeading) {
        counts.headings++;
        var newText = '附录' + appState.letter + (appMatch[2] ? ' ' + appMatch[2] : '');
        console.log('[scan] 附录: ' + text + ' → ' + newText);
        if (text !== newText) plans.push({ oldText: text, newText: newText });
      }
      continue;
    }

    // === 附录内处理 ===
    if (inAppendix) {
      if (/^第[一二三四五六七八九十]+章/.test(text)) {
        inAppendix = false;
      } else {
        // 附录内标题
        if (needHeading) {
          var m1 = text.match(/^[A-Z]\.(\d+)\s+(.+)$/);
          if (m1 && isChineseTitle(m1[2])) {
            appState.l1++; appState.l2 = 0; appState.l3 = 0;
            counts.headings++;
            var newText = appState.letter + '.' + appState.l1 + ' ' + m1[2];
            if (text !== newText) plans.push({ oldText: text, newText: newText });
            continue;
          }
          var m2 = text.match(/^[A-Z]\.(\d+)\.(\d+)\s+(.+)$/);
          if (m2 && isChineseTitle(m2[3])) {
            if (appState.l1 === 0) appState.l1 = 1;
            appState.l2++; appState.l3 = 0;
            counts.headings++;
            var newText = appState.letter + '.' + appState.l1 + '.' + appState.l2 + ' ' + m2[3];
            if (text !== newText) plans.push({ oldText: text, newText: newText });
            continue;
          }
          var m3 = text.match(/^[A-Z]\.(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
          if (m3 && isChineseTitle(m3[4])) {
            if (appState.l1 === 0) appState.l1 = 1;
            if (appState.l2 === 0) appState.l2 = 1;
            appState.l3++;
            counts.headings++;
            var newText = appState.letter + '.' + appState.l1 + '.' + appState.l2 + '.' + appState.l3 + ' ' + m3[4];
            if (text !== newText) plans.push({ oldText: text, newText: newText });
            continue;
          }
        }

        // 附录内图编号：图A1、图A.1 或 图1
        if (needFigure) {
          var appFig = text.match(/^图\s*([A-Z])?\.?(\d+)\s+(.+)$/i);
          if (appFig) {
            appendixFigureCounter++;
            counts.figures++;
            var newText = '图' + appState.letter + appendixFigureCounter + ' ' + appFig[3];
            console.log('[scan] 附录图: ' + text + ' → ' + newText);
            if (text !== newText) plans.push({ oldText: text, newText: newText });
            continue;
          }
        }

        // 附录内表编号：表A1、表A.1 或 表1
        if (needTable) {
          var appTbl = text.match(/^表\s*([A-Z])?\.?(\d+)\s+(.+)$/i);
          if (appTbl) {
            appendixTableCounter++;
            counts.tables++;
            var newText = '表' + appState.letter + appendixTableCounter + ' ' + appTbl[3];
            console.log('[scan] 附录表: ' + text + ' → ' + newText);
            if (text !== newText) plans.push({ oldText: text, newText: newText });
            continue;
          }
        }
        continue;
      }
    }

    // === 正文标题（同时更新章节状态） ===
    // 一级：第X章
    var h1 = text.match(/^第([一二三四五六七八九十]+)章\s*(.*)$/);
    if (h1) {
      state.ch++;
      state.sec = 0; state.sub = 0; state.item = 0; state.subItem = 0;
      if (needHeading) {
        counts.headings++;
        var newText = '第' + (num2cn[state.ch] || state.ch) + '章 ' + h1[2];
        if (text !== newText) plans.push({ oldText: text, newText: newText });
      }
      continue;
    }

    // 一级：数字格式
    var h1n = text.match(/^(\d+)\s+([^\d\s].*)$/);
    if (h1n && isChineseTitle(h1n[2]) && !text.match(/^\d+\.\d/)) {
      state.ch++;
      state.sec = 0; state.sub = 0; state.item = 0; state.subItem = 0;
      if (needHeading) {
        counts.headings++;
        var newText = state.ch + ' ' + h1n[2];
        if (text !== newText) plans.push({ oldText: text, newText: newText });
      }
      continue;
    }

    // 二级：X.X
    var h2 = text.match(/^(\d+)\.(\d+)\s+(.+)$/);
    if (h2 && isChineseTitle(h2[3]) && text.indexOf('表') !== 0 && text.indexOf('图') !== 0) {
      if (state.ch === 0) state.ch = 1;
      state.sec++; state.sub = 0; state.item = 0; state.subItem = 0;
      if (needHeading) {
        counts.headings++;
        var newText = state.ch + '.' + state.sec + ' ' + h2[3];
        if (text !== newText) plans.push({ oldText: text, newText: newText });
      }
      continue;
    }

    // 三级：X.X.X
    var h3 = text.match(/^(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
    if (h3 && isChineseTitle(h3[4])) {
      if (state.ch === 0) state.ch = 1;
      if (state.sec === 0) state.sec = 1;
      state.sub++; state.item = 0; state.subItem = 0;
      if (needHeading) {
        counts.headings++;
        var newText = state.ch + '.' + state.sec + '.' + state.sub + ' ' + h3[4];
        if (text !== newText) plans.push({ oldText: text, newText: newText });
      }
      continue;
    }

    // 四级：X.X.X.X
    var h4 = text.match(/^(\d+)\.(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
    if (h4 && isChineseTitle(h4[5])) {
      if (state.ch === 0) state.ch = 1;
      if (state.sec === 0) state.sec = 1;
      if (state.sub === 0) state.sub = 1;
      state.item++; state.subItem = 0;
      if (needHeading) {
        counts.headings++;
        var newText = state.ch + '.' + state.sec + '.' + state.sub + '.' + state.item + ' ' + h4[5];
        if (text !== newText) plans.push({ oldText: text, newText: newText });
      }
      continue;
    }

    // 五级：X.X.X.X.X
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
        if (text !== newText) plans.push({ oldText: text, newText: newText });
      }
      continue;
    }

    // === 图编号 ===
    if (needFigure) {
      // 格式：图X.Y-Z 标题 或 图X-Y 标题 或 图X 标题
      var figMatch = text.match(/^图\s*(\d+)(?:\.(\d+))?(?:-(\d+))?\s+(.+)$/);
      if (figMatch) {
        counts.figures++;

        if (figFormat === 'simple') {
          // 顺序式：图1、图2、图3... 全文递增
          simpleFigureCounter++;
          var newText = '图' + simpleFigureCounter + ' ' + figMatch[4];
          console.log('[scan] 图(顺序式): ' + text + ' → ' + newText);
          if (text !== newText) plans.push({ oldText: text, newText: newText });
        } else {
          // 章节式：图X.Y-Z 格式
          if (state.ch === 0) state.ch = 1;
          var figCh = state.ch;
          var figSec = state.sec > 0 ? state.sec : 1;
          var figKey = figCh + '.' + figSec;

          figureCounters[figKey] = (figureCounters[figKey] || 0) + 1;
          var figNum = figureCounters[figKey];

          var newText = '图' + figCh + '.' + figSec + '-' + figNum + ' ' + figMatch[4];
          console.log('[scan] 图(章节式): ' + text + ' → ' + newText);
          if (text !== newText) plans.push({ oldText: text, newText: newText });
        }
        continue;
      }
    }

    // === 表编号 ===
    if (needTable) {
      // 格式：表X.Y-Z 标题 或 表X-Y 标题 或 表X 标题
      var tblMatch = text.match(/^表\s*(\d+)(?:\.(\d+))?(?:-(\d+))?\s+(.+)$/);
      if (tblMatch) {
        counts.tables++;

        if (tblFormat === 'simple') {
          // 顺序式：表1、表2、表3... 全文递增
          simpleTableCounter++;
          var newText = '表' + simpleTableCounter + ' ' + tblMatch[4];
          console.log('[scan] 表(顺序式): ' + text + ' → ' + newText);
          if (text !== newText) plans.push({ oldText: text, newText: newText });
        } else {
          // 章节式：表X.Y-Z 格式
          if (state.ch === 0) state.ch = 1;
          var tblCh = state.ch;
          var tblSec = state.sec > 0 ? state.sec : 1;
          var tblKey = tblCh + '.' + tblSec;

          tableCounters[tblKey] = (tableCounters[tblKey] || 0) + 1;
          var tblNum = tableCounters[tblKey];

          var newText = '表' + tblCh + '.' + tblSec + '-' + tblNum + ' ' + tblMatch[4];
          console.log('[scan] 表(章节式): ' + text + ' → ' + newText);
          if (text !== newText) plans.push({ oldText: text, newText: newText });
        }
        continue;
      }
    }
  }

  console.log('[scan] 规划完成：标题' + counts.headings + ' 图' + counts.figures + ' 表' + counts.tables + '，待修复' + plans.length + '处');

  // 从后往前替换
  var origTrack = doc.TrackRevisions;
  doc.TrackRevisions = true;
  var totalFixed = 0;
  var log = [];

  var replaceList = [];
  for (var p = 0; p < plans.length; p++) {
    try {
      var sr = doc.Range(0, doc.Content.End);
      sr.Find.ClearFormatting();
      sr.Find.Forward = true;
      sr.Find.Wrap = 0;
      if (sr.Find.Execute(plans[p].oldText, false, false, false, false, false, true, 1, false)) {
        replaceList.push({ start: sr.Start, end: sr.End, newText: plans[p].newText, oldText: plans[p].oldText });
      }
    } catch (e) {}
  }
  replaceList.sort(function(a, b) { return b.start - a.start; });
  for (var r = 0; r < replaceList.length; r++) {
    try {
      var rng = doc.Range(replaceList[r].start, replaceList[r].end);
      rng.Text = replaceList[r].newText;
      totalFixed++;
      log.push({ original: replaceList[r].oldText, suggested: replaceList[r].newText });
    } catch (e) {}
  }
  doc.TrackRevisions = origTrack;

  console.log('[scan] 完成，修复: ' + totalFixed);
  return { success: true, totalFixed: totalFixed, fixed: totalFixed, details: log, structure: counts, summary: { totalIssues: totalFixed } };
} catch (e) {
  return { success: false, error: String(e) };
}