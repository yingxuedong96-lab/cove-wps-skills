/**
 * scan-structure.js
 * 校对标题编号、图编号、表编号、公式编号
 * 支持自动编号检测与转换
 * scope: heading/figure/table/formula/numbering
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

  function clean(t) { return String(t || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim(); }
  function isChinese(s) { return s && /^[\u4e00-\u9fa5]/.test(s); }

  console.log('[scan] 开始, scope=' + scopeType);

  var state = { ch: 0, sec: 0, sub: 0, item: 0, subItem: 0 };
  var appState = { letter: '', letterIdx: 0, l1: 0, l2: 0, l3: 0 };
  var inApp = false;
  var figCtr = {}, figSim = 0, figApp = 0;
  var tblCtr = {}, tblSim = 0, tblApp = 0;
  var frmCtr = {}, frmSim = 0, frmApp = 0;

  var counts = { headings: 0, figures: 0, tables: 0, formulas: 0, autoConverted: 0 };
  var plans = [];
  var autoLog = [];

  // 开启修订
  var origTrack = doc.TrackRevisions;
  doc.TrackRevisions = true;

  var paras = doc.Paragraphs;
  var total = paras.Count;
  console.log('[scan] 段落数: ' + total);

  // 分批处理，每批100段
  var batchSize = 100;
  var batches = Math.ceil(total / batchSize);

  for (var batch = 0; batch < batches; batch++) {
    var start = batch * batchSize + 1;
    var end = Math.min((batch + 1) * batchSize, total);

    for (var i = start; i <= end; i++) {
      var para = paras.Item(i);
      var rng = para.Range;
      var text = clean(rng.Text);
      if (!text) continue;

      // 检测并转换自动编号（轻量级）
      try {
        var lf = rng.ListFormat;
        if (lf && lf.ListType !== 0) {
          var ls = lf.ListString;
          if (ls) {
            rng.InsertBefore(ls + ' ');
            lf.RemoveNumbers();
            counts.autoConverted++;
            autoLog.push(ls);
            text = clean(rng.Text);
          }
        }
      } catch (e) {}

      // === 处理逻辑（精简版）===
      // 附录
      var am = text.match(/^附\s*录\s*([A-Z一二三四五六七八九十]+)[\s　]*(.*)$/i);
      if (am) {
        inApp = true;
        appState.l1 = 0; appState.l2 = 0; appState.l3 = 0;
        figApp = 0; tblApp = 0; frmApp = 0;
        appState.letterIdx++;
        appState.letter = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.charAt((appState.letterIdx - 1) % 26);
        if (needHeading) {
          counts.headings++;
          var nt = '附录' + appState.letter + (am[2] ? ' ' + am[2] : '');
          if (text !== nt) plans.push({ para: para, old: text, new: nt });
        }
        continue;
      }

      // 附录内
      if (inApp && !/^第[一二三四五六七八九十]+章/.test(text)) {
        if (needHeading) {
          var m1 = text.match(/^[A-Z]\.(\d+)\s+(.+)$/);
          if (m1 && isChinese(m1[2])) {
            appState.l1++; counts.headings++;
            var nt = appState.letter + '.' + appState.l1 + ' ' + m1[2];
            if (text !== nt) plans.push({ para: para, old: text, new: nt });
            continue;
          }
        }
        if (needFigure && /^图\s*([A-Z])?\.?(\d+)\s+/i.test(text)) {
          figApp++; counts.figures++;
          var m = text.match(/^图\s*([A-Z])?\.?(\d+)\s+(.+)$/i);
          if (m) { var nt = '图' + appState.letter + figApp + ' ' + m[3]; if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
          continue;
        }
        if (needTable && /^表\s*([A-Z])?\.?(\d+)\s+/i.test(text)) {
          tblApp++; counts.tables++;
          var m = text.match(/^表\s*([A-Z])?\.?(\d+)\s+(.+)$/i);
          if (m) { var nt = '表' + appState.letter + tblApp + ' ' + m[3]; if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
          continue;
        }
        if (needFormula && /\(([A-Z])?\.?(\d+)\)$/.test(text)) {
          frmApp++; counts.formulas++;
          var nt = text.replace(/\(([A-Z])?\.?(\d+)\)$/, '(' + appState.letter + frmApp + ')');
          if (text !== nt) plans.push({ para: para, old: text, new: nt });
          continue;
        }
        continue;
      }
      inApp = false;

      // 一级：第X章
      var h1 = text.match(/^第([一二三四五六七八九十]+)章\s*(.*)$/);
      if (h1) {
        state.ch++; state.sec = 0; state.sub = 0; state.item = 0; state.subItem = 0;
        if (needHeading) { counts.headings++; var nt = '第' + (num2cn[state.ch] || state.ch) + '章 ' + h1[2]; if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
        continue;
      }

      // 一级：数字
      var h1n = text.match(/^(\d+)\s+([^\d\s].*)$/);
      if (h1n && isChinese(h1n[2]) && !/^\d+\.\d/.test(text)) {
        state.ch++; state.sec = 0; state.sub = 0; state.item = 0; state.subItem = 0;
        if (needHeading) { counts.headings++; var nt = state.ch + ' ' + h1n[2]; if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
        continue;
      }

      // 二级
      var h2 = text.match(/^(\d+)\.(\d+)\s+(.+)$/);
      if (h2 && isChinese(h2[3]) && text.indexOf('表') !== 0 && text.indexOf('图') !== 0) {
        if (state.ch === 0) state.ch = 1; state.sec++; state.sub = 0; state.item = 0; state.subItem = 0;
        if (needHeading) { counts.headings++; var nt = state.ch + '.' + state.sec + ' ' + h2[3]; if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
        continue;
      }

      // 三级
      var h3 = text.match(/^(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
      if (h3 && isChinese(h3[4])) {
        if (state.ch === 0) state.ch = 1; if (state.sec === 0) state.sec = 1;
        state.sub++; state.item = 0; state.subItem = 0;
        if (needHeading) { counts.headings++; var nt = state.ch + '.' + state.sec + '.' + state.sub + ' ' + h3[4]; if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
        continue;
      }

      // 四级
      var h4 = text.match(/^(\d+)\.(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
      if (h4 && isChinese(h4[5])) {
        if (state.ch === 0) state.ch = 1; if (state.sec === 0) state.sec = 1; if (state.sub === 0) state.sub = 1;
        state.item++; state.subItem = 0;
        if (needHeading) { counts.headings++; var nt = state.ch + '.' + state.sec + '.' + state.sub + '.' + state.item + ' ' + h4[5]; if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
        continue;
      }

      // 五级
      var h5 = text.match(/^(\d+)\.(\d+)\.(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
      if (h5 && isChinese(h5[6])) {
        if (state.ch === 0) state.ch = 1; if (state.sec === 0) state.sec = 1; if (state.sub === 0) state.sub = 1; if (state.item === 0) state.item = 1;
        state.subItem++;
        if (needHeading) { counts.headings++; var nt = state.ch + '.' + state.sec + '.' + state.sub + '.' + state.item + '.' + state.subItem + ' ' + h5[6]; if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
        continue;
      }

      // 图
      if (needFigure) {
        var fm = text.match(/^图\s*(\d+)(?:\.(\d+))?(?:-(\d+))?\s+(.+)$/);
        if (fm) {
          counts.figures++;
          if (figFormat === 'simple') { figSim++; var nt = '图' + figSim + ' ' + fm[4]; if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
          else { var fc = state.ch || 1, fs = state.sec || 1, fk = fc + '.' + fs; figCtr[fk] = (figCtr[fk] || 0) + 1; var nt = '图' + fc + '.' + fs + '-' + figCtr[fk] + ' ' + fm[4]; if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
          continue;
        }
      }

      // 表
      if (needTable) {
        var tm = text.match(/^表\s*(\d+)(?:\.(\d+))?(?:-(\d+))?\s+(.+)$/);
        if (tm) {
          counts.tables++;
          if (tblFormat === 'simple') { tblSim++; var nt = '表' + tblSim + ' ' + tm[4]; if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
          else { var tc = state.ch || 1, ts = state.sec || 1, tk = tc + '.' + ts; tblCtr[tk] = (tblCtr[tk] || 0) + 1; var nt = '表' + tc + '.' + ts + '-' + tblCtr[tk] + ' ' + tm[4]; if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
          continue;
        }
      }

      // 公式
      if (needFormula) {
        var pm = text.match(/\((\d+)(?:\.(\d+))?(?:-(\d+))?\)$/);
        if (pm) {
          counts.formulas++;
          if (frmFormat === 'simple') { frmSim++; var nt = text.replace(/\((\d+)(?:\.(\d+))?(?:-(\d+))?\)$/, '(' + frmSim + ')'); if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
          else { var fc = state.ch || 1, fs = state.sec || 1, fk = fc + '.' + fs; frmCtr[fk] = (frmCtr[fk] || 0) + 1; var nt = text.replace(/\((\d+)(?:\.(\d+))?(?:-(\d+))?\)$/, '(' + fc + '.' + fs + '-' + frmCtr[fk] + ')'); if (text !== nt) plans.push({ para: para, old: text, new: nt }); }
          continue;
        }
      }
    }
    // 每批处理后输出进度
    if (batch % 5 === 0 || batch === batches - 1) {
      console.log('[scan] 进度: ' + Math.min(end, total) + '/' + total);
    }
  }

  console.log('[scan] 规划完成，待修复' + plans.length + '处');

  // 执行修复
  var fixed = 0, log = [];
  for (var p = 0; p < plans.length; p++) {
    try {
      plans[p].para.Range.Text = plans[p].new;
      fixed++;
      log.push({ original: plans[p].old, suggested: plans[p].new });
    } catch (e) {}
  }

  doc.TrackRevisions = origTrack;

  var msg = '编号校对完成：修复 ' + fixed + ' 处';
  if (counts.autoConverted > 0) msg += '，转换自动编号 ' + counts.autoConverted + ' 处';

  return { success: true, totalFixed: fixed, fixed: fixed, details: log, structure: counts, autoNumberConversions: autoLog, summary: { totalIssues: fixed, autoNumberDetected: counts.autoConverted }, message: msg };
} catch (e) {
  return { success: false, error: String(e) };
}