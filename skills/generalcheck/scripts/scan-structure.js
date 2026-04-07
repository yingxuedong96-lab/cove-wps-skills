/**
 * scan-structure.js
 * 校对标题编号：正文（第X章、X.X、X.X.X、X.X.X.X）和附录（A、A.1、A.1.1）
 */
try {
  var doc = Application.ActiveDocument;
  if (!doc) return { success: false, error: '没有打开的文档' };

  var scopeType = typeof scope !== 'undefined' ? scope : 'heading';
  var needHeading = scopeType === 'numbering' || scopeType === 'heading';

  var cn2num = {
    '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
    '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
    '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15
  };
  var num2cn = {
    1: '一', 2: '二', 3: '三', 4: '四', 5: '五',
    6: '六', 7: '七', 8: '八', 9: '九', 10: '十',
    11: '十一', 12: '十二', 13: '十三', 14: '十四', 15: '十五'
  };

  function cleanText(text) {
    return String(text || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function isChineseTitle(str) {
    return str && /^[\u4e00-\u9fa5]/.test(str);
  }

  function pushPlan(plans, oldText, newText) {
    if (oldText && newText && oldText !== newText) {
      plans.push({ oldText: oldText, newText: newText });
    }
  }

  var docText = doc.Content && doc.Content.Text ? String(doc.Content.Text) : '';
  var paras = docText.split('\r');
  var totalParas = paras.length;
  console.log('[scan] 开始规划，总段落数: ' + totalParas);

  var plans = [];
  var counts = { headings: 0 };

  // 正文标题状态
  var expectedChapter = 0;
  var currentChapter = 1;
  var expectedSection = 0;
  var currentSection = 0;
  var expectedSubsection = 0;
  var currentSubsection = 0;
  var expectedItem = 0;
  var currentItem = 0;

  // 附录状态
  var inAppendix = false;
  var appendixIndex = 0;
  var appendixLevel1 = 0;
  var appendixLevel2 = 0;
  var appendixLevel3 = 0;

  for (var i = 0; i < totalParas; i++) {
    var text = cleanText(paras[i]);
    if (!text) continue;

    // 检测附录标题
    var appMatch = text.match(/^附\s*录\s*([A-Z])[\s　]*(.*)$/i);
    if (appMatch) {
      appendixIndex++;
      var newLetter = String.fromCharCode(64 + appendixIndex);
      inAppendix = true;
      appendixLevel1 = 0;
      appendixLevel2 = 0;
      appendixLevel3 = 0;
      counts.headings++;
      var newText = '附录 ' + newLetter + (appMatch[2] ? ' ' + appMatch[2] : '');
      console.log('[scan] 附录: ' + text + ' → ' + newText);
      pushPlan(plans, text, newText);
      continue;
    }

    // 在附录中
    if (inAppendix) {
      // 检测是否回到正文
      if (/^第[一二三四五六七八九十]+章/.test(text)) {
        inAppendix = false;
      }

      // 附录一级标题 A.1
      var appL1 = text.match(/^[A-Z]\.(\d+)\s+(.+)$/);
      if (appL1 && isChineseTitle(appL1[2])) {
        appendixLevel1++;
        appendixLevel2 = 0;
        appendixLevel3 = 0;
        counts.headings++;
        var newLetter = String.fromCharCode(64 + appendixIndex);
        var newText = newLetter + '.' + appendixLevel1 + ' ' + appL1[2];
        console.log('[scan] 附录一级: ' + text + ' → ' + newText);
        pushPlan(plans, text, newText);
        continue;
      }

      // 附录二级标题 A.1.1
      var appL2 = text.match(/^[A-Z]\.(\d+)\.(\d+)\s+(.+)$/);
      if (appL2 && isChineseTitle(appL2[4])) {
        if (appendixLevel1 === 0) appendixLevel1 = 1;
        appendixLevel2++;
        appendixLevel3 = 0;
        counts.headings++;
        var newLetter = String.fromCharCode(64 + appendixIndex);
        var newText = newLetter + '.' + appendixLevel1 + '.' + appendixLevel2 + ' ' + appL2[4];
        console.log('[scan] 附录二级: ' + text + ' → ' + newText);
        pushPlan(plans, text, newText);
        continue;
      }

      // 附录三级标题 A.1.1.1
      var appL3 = text.match(/^[A-Z]\.(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
      if (appL3 && isChineseTitle(appL3[5])) {
        if (appendixLevel1 === 0) appendixLevel1 = 1;
        if (appendixLevel2 === 0) appendixLevel2 = 1;
        appendixLevel3++;
        counts.headings++;
        var newLetter = String.fromCharCode(64 + appendixIndex);
        var newText = newLetter + '.' + appendixLevel1 + '.' + appendixLevel2 + '.' + appendixLevel3 + ' ' + appL3[5];
        console.log('[scan] 附录三级: ' + text + ' → ' + newText);
        pushPlan(plans, text, newText);
        continue;
      }

      continue; // 跳过附录中的其他段落
    }

    // 正文标题处理
    // 一级标题：第X章
    var m1 = text.match(/^第([一二三四五六七八九十]+)章\s*(.*)$/);
    if (m1) {
      expectedChapter++;
      currentChapter = expectedChapter;
      expectedSection = 0;
      expectedSubsection = 0;
      expectedItem = 0;
      currentSection = 0;
      currentSubsection = 0;
      currentItem = 0;
      counts.headings++;
      var newText = '第' + (num2cn[expectedChapter] || expectedChapter) + '章 ' + m1[2];
      console.log('[scan] 一级: ' + text + ' → ' + newText);
      pushPlan(plans, text, newText);
      continue;
    }

    // 一级标题：数字格式 如 "3 系统范围"
    var m1num = text.match(/^(\d+)\s+([^\d\s].*)$/);
    if (m1num && isChineseTitle(m1num[2])) {
      expectedChapter++;
      currentChapter = expectedChapter;
      expectedSection = 0;
      expectedSubsection = 0;
      expectedItem = 0;
      currentSection = 0;
      currentSubsection = 0;
      currentItem = 0;
      counts.headings++;
      var newText = expectedChapter + ' ' + m1num[2];
      console.log('[scan] 一级: ' + text + ' → ' + newText);
      pushPlan(plans, text, newText);
      continue;
    }

    // 二级标题：X.X
    var m2 = text.match(/^(\d+)\.(\d+)\s+(.+)$/);
    if (m2 && isChineseTitle(m2[3]) && text.indexOf('表') !== 0 && text.indexOf('图') !== 0) {
      if (expectedChapter === 0) { expectedChapter = 1; currentChapter = 1; }
      expectedSection++;
      currentSection = expectedSection;
      expectedSubsection = 0;
      expectedItem = 0;
      currentSubsection = 0;
      currentItem = 0;
      counts.headings++;
      var newText = currentChapter + '.' + expectedSection + ' ' + m2[3];
      console.log('[scan] 二级: ' + text + ' → ' + newText);
      pushPlan(plans, text, newText);
      continue;
    }

    // 三级标题：X.X.X
    var m3 = text.match(/^(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
    if (m3 && isChineseTitle(m3[4])) {
      if (expectedChapter === 0) { expectedChapter = 1; currentChapter = 1; }
      if (currentSection === 0) { currentSection = 1; expectedSection = 1; }
      expectedSubsection++;
      currentSubsection = expectedSubsection;
      expectedItem = 0;
      currentItem = 0;
      counts.headings++;
      var newText = currentChapter + '.' + currentSection + '.' + expectedSubsection + ' ' + m3[4];
      console.log('[scan] 三级: ' + text + ' → ' + newText);
      pushPlan(plans, text, newText);
      continue;
    }

    // 四级标题：X.X.X.X
    var m4 = text.match(/^(\d+)\.(\d+)\.(\d+)\.(\d+)\s+(.+)$/);
    if (m4 && isChineseTitle(m4[5])) {
      if (expectedChapter === 0) { expectedChapter = 1; currentChapter = 1; }
      if (currentSection === 0) { currentSection = 1; expectedSection = 1; }
      if (currentSubsection === 0) { currentSubsection = 1; expectedSubsection = 1; }
      expectedItem++;
      currentItem = expectedItem;
      counts.headings++;
      var newText = currentChapter + '.' + currentSection + '.' + currentSubsection + '.' + expectedItem + ' ' + m4[5];
      console.log('[scan] 四级: ' + text + ' → ' + newText);
      pushPlan(plans, text, newText);
      continue;
    }
  }

  console.log('[scan] 规划完成，标题' + counts.headings + '，待修复' + plans.length + '处');

  // 执行替换：从后往前替换，避免位置偏移
  var origTrack = doc.TrackRevisions;
  doc.TrackRevisions = true;

  var revisionLog = [];
  var totalFixed = 0;

  // 先收集所有位置
  var replaceList = [];
  for (var j = 0; j < plans.length; j++) {
    var plan = plans[j];
    try {
      var searchRange = doc.Range(0, doc.Content.End);
      searchRange.Find.ClearFormatting();
      searchRange.Find.Forward = true;
      searchRange.Find.Wrap = 0;
      searchRange.Find.MatchWildcards = false;
      var found = searchRange.Find.Execute(plan.oldText, false, false, false, false, false, true, 1, false);
      if (found) {
        replaceList.push({
          start: searchRange.Start,
          end: searchRange.End,
          oldText: plan.oldText,
          newText: plan.newText
        });
      }
    } catch (e) {}
  }

  // 按位置倒序排序，从后往前替换
  replaceList.sort(function(a, b) { return b.start - a.start; });

  for (var k = 0; k < replaceList.length; k++) {
    var item = replaceList[k];
    try {
      var range = doc.Range(item.start, item.end);
      range.Text = item.newText;
      totalFixed++;
      revisionLog.push({
        original: item.oldText,
        suggested: item.newText
      });
    } catch (e) {
      console.log('[scan] 替换失败: ' + item.oldText);
    }
  }

  doc.TrackRevisions = origTrack;

  console.log('[scan] 完成，修复: ' + totalFixed);
  return {
    success: true,
    totalFixed: totalFixed,
    fixed: totalFixed,
    details: revisionLog,
    structure: { headings: counts.headings },
    summary: { totalIssues: totalFixed }
  };
} catch (e) {
  console.warn('[scan] 错误: ' + e);
  return { success: false, error: String(e) };
}