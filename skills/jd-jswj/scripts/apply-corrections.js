// apply-corrections.js
// 以修订模式（Track Changes）将校对结果批量写入当前 WPS 活动文档
// 参数：corrections —— [{original, corrected, type, reason}, ...]

try {
  // ── 接收参数 ───────────────────────────────────────────────────────────────
  var correctionsParam = typeof corrections !== 'undefined' ? corrections : [];

  // JSON 三级容错：支持字符串或数组两种传入形式
  if (typeof correctionsParam === 'string') {
    var parsed = null;
    // 第一级：提取代码块
    var blockMatch = correctionsParam.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    var jsonStr = blockMatch ? blockMatch[1].trim() : correctionsParam.trim();
    // 第二级：直接解析
    try { parsed = JSON.parse(jsonStr); } catch (e1) {}
    // 第三级：修复尾逗号
    if (!parsed) {
      var fixed = jsonStr.replace(/,\s*}/g, '}').replace(/,\s*]/g, ']');
      try { parsed = JSON.parse(fixed); } catch (e2) {}
    }
    correctionsParam = parsed ? parsed : [];
  }

  if (!correctionsParam || correctionsParam.length === 0) {
    return { success: true, applied: 0, total: 0, notFound: [], message: '无需修改' };
  }

  // ── 获取文档 ───────────────────────────────────────────────────────────────
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, error: '无活动文档' };
  }

  // ── 开启修订模式 ───────────────────────────────────────────────────────────
  doc.TrackRevisions = true;

  var applied = 0;
  var notFound = [];

  // ── 从后往前应用，防止偏移漂移 ────────────────────────────────────────────
  for (var i = correctionsParam.length - 1; i >= 0; i--) {
    var item = correctionsParam[i];
    if (!item) { continue; }

    var orig = item.original ? String(item.original) : '';
    var corr = item.corrected !== undefined ? String(item.corrected) : '';

    if (orig.length === 0) { continue; }
    if (orig === corr) { continue; }

    try {
      var f = doc.Content.Find;
      f.ClearFormatting();
      f.Replacement.ClearFormatting();
      f.Text = orig;
      f.Replacement.Text = corr;
      f.Forward = true;
      f.Wrap = 1;  // wdFindContinue：全文搜索
      // Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards,
      //         MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format,
      //         ReplaceWith, Replace)
      // Replace=2 → wdReplaceAll
      var result = f.Execute(null, false, false, false, false, false, true, 1, false, null, 2);
      if (result) {
        applied++;
      } else {
        notFound.push(orig);
      }
    } catch (itemErr) {
      notFound.push(orig + ' [ERR: ' + String(itemErr) + ']');
    }
  }

  // ── 保存文档 ───────────────────────────────────────────────────────────────
  try { doc.Save(); } catch (saveErr) {}

  return {
    success: true,
    applied: applied,
    total: correctionsParam.length,
    notFound: notFound
  };

} catch (e) {
  return { success: false, error: String(e) };
}
