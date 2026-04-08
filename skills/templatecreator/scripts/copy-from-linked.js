/**
 * copy-from-linked.js - 从关联文档复制样式到当前文档
 * 版本: 26.0410.1011
 * 参数:
 *   - sourceDocPath: 关联文档的完整路径
 * 流程：
 *   1. 打开源文档提取样式
 *   2. 关闭源文档
 *   3. 应用样式到当前文档
 */
try {
  var VER = "26.0410.1011";
  console.log("[copy-from-linked] 版本: " + VER);

  var targetDoc = Application.ActiveDocument;
  if (!targetDoc) return { success: false, error: "没有打开的目标文档" };

  // 获取源文档路径
  var sourcePath = typeof sourceDocPath !== 'undefined' ? sourceDocPath : '';
  if (!sourcePath) {
    return { success: false, error: "缺少 sourceDocPath 参数，请提供关联文档的路径" };
  }

  console.log("[copy-from-linked] 源文档: " + sourcePath);
  console.log("[copy-from-linked] 目标文档: " + targetDoc.Name);

  // ============================================================
  // 步骤1: 打开源文档提取样式
  // ============================================================
  var sourceDoc;
  try {
    sourceDoc = Application.Documents.Open(sourcePath);
    console.log("[copy-from-linked] 成功打开源文档");
  } catch (e) {
    return { success: false, error: "无法打开源文档: " + String(e) };
  }

  // 自动检测文档类型
  var paperScore = 0, govScore = 0;
  var paras = sourceDoc.Paragraphs;
  var scanCount = Math.min(30, paras.Count);

  function clean(t) {
    return String(t || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  for (var i = 1; i <= scanCount; i++) {
    var text = clean(paras.Item(i).Range.Text);
    if (!text) continue;
    if (/^第[一二三四五六七八九十百零\d]+章/.test(text)) paperScore += 3;
    if (/^\d+\.\d+/.test(text)) paperScore += 2;
    if (/^附\s*录/.test(text)) paperScore += 2;
    if (/^摘\s*要|^Abstract/i.test(text)) paperScore += 2;
    if (/^图\s*\d+|^表\s*\d+/.test(text)) paperScore += 1;
    if (/^关于/.test(text)) govScore += 3;
    if (/通知$|决定$|意见$|办法$|规定$/.test(text)) govScore += 2;
    if (/^附\s*件/.test(text)) govScore += 2;
  }

  var isPaper = paperScore >= govScore;
  console.log("[copy-from-linked] 类型: " + (isPaper ? "论文报告" : "公文"));

  // 样式名称
  var STYLE_NAMES = {
    coverTitle: "封面标题", chapterTitle: "章标题",
    heading1: "一级标题", heading2: "二级标题", heading3: "三级标题",
    heading4: "四级标题", heading5: "五级标题",
    body: "正文", listItem: "列表项",
    figureCaption: "图名", tableCaption: "表名",
    appendixTitle: "附录标题", appendixSection: "附录节题",
    docTitle: "公文标题", attachment: "附件说明",
    reference: "参考文献", referenceTitle: "参考文献标题",
    abstract: "摘要", keyword: "关键词"
  };

  var lastAppendixTitle = false;

  function detectType(text, isPaper, fmt) {
    if (isPaper) {
      if (/^第[一二三四五六七八九十百零\d]+章/.test(text)) return "chapterTitle";
      if (/^\d+\.\d+\.\d+\.\d+\.\d+/.test(text)) return "heading5";
      if (/^\d+\.\d+\.\d+\.\d+/.test(text)) return "heading4";
      if (/^\d+\.\d+\.\d+[^.\d]/.test(text)) return "heading3";
      if (/^\d+\.\d+[^.\d]/.test(text)) return "heading2";
      if (/^\d+\s+[^\d\.\s]/.test(text)) return "heading1";
      if (/^附\s*录\s*[A-Z０-９０-９]/.test(text)) { lastAppendixTitle = true; return "appendixTitle"; }
      if (/^[A-Z]\.\d+/.test(text)) return "appendixSection";
      if (/^图\s*\d+/.test(text)) return "figureCaption";
      if (/^表\s*\d+/.test(text)) return "tableCaption";
      if (/^参考文[献獻]/.test(text)) return "referenceTitle";
      if (/^\[\d+\]/.test(text)) return "reference";
      if (/^目\s*录$|^目次$/.test(text)) return "tocTitle";
      if (/^摘\s*要|^Abstract/i.test(text)) return "abstract";
      if (/^关键词|^关键字|^Key\s*words/i.test(text)) return "keyword";
      if (/^[\(（]\d+[\)）]/.test(text)) return "listItem";
      if (/^[\(（][a-z][\)）]/.test(text)) return "listItem";
      if (/^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮]/.test(text)) return "listItem";
      if (lastAppendixTitle && fmt && fmt.bold && fmt.fontSize >= 14) { lastAppendixTitle = false; return "appendixTitle"; }
    } else {
      if (/^关于|通知$|决定$|意见$|办法$|规定$/.test(text)) return "docTitle";
      if (/^[一二三四五六七八九十]+、/.test(text)) return "heading1";
      if (/^[\(（][一二三四五六七八九十]+[\)）]/.test(text)) return "heading2";
      if (/^\d+\.\s/.test(text)) return "heading3";
      if (/^附\s*件/.test(text)) return "attachment";
      if (/^图\s*\d+/.test(text)) return "figureCaption";
      if (/^表\s*\d+/.test(text)) return "tableCaption";
    }
    if (fmt && fmt.firstLineIndent < 0) return "listItem";
    if (isPaper && !/^附\s*录/.test(text)) lastAppendixTitle = false;
    return "body";
  }

  function extractFormat(para) {
    var rng = para.Range, fmt = para.Format;
    return {
      fontCN: rng.Font.NameFarEast || rng.Font.Name || "",
      fontEN: rng.Font.NameAscii || "",
      fontSize: rng.Font.Size,
      bold: rng.Font.Bold ? true : false,
      italic: rng.Font.Italic ? true : false,
      alignment: fmt.Alignment,
      firstLineIndent: fmt.FirstLineIndent || 0,
      leftIndent: fmt.LeftIndent || 0,
      spaceBefore: fmt.SpaceBefore || 0,
      spaceAfter: fmt.SpaceAfter || 0,
      lineSpacing: fmt.LineSpacing || 0,
      lineSpacingRule: fmt.LineSpacingRule
    };
  }

  // 提取样式
  var styles = {};
  paras = sourceDoc.Paragraphs;

  for (var i = 1; i <= paras.Count; i++) {
    var para = paras.Item(i);
    var text = clean(para.Range.Text);
    if (!text) continue;

    var fmt = extractFormat(para);
    var type = detectType(text, isPaper, fmt);

    if (!type) {
      if (fmt.bold && fmt.fontSize >= 20) type = "coverTitle";
      else if (fmt.bold && fmt.fontSize >= 16) type = "heading1";
      else if (fmt.bold && fmt.fontSize >= 14) type = "heading2";
      else type = "body";
    }

    if (!styles[type]) {
      styles[type] = { id: type, name: STYLE_NAMES[type] || type, count: 0, format: fmt, samples: [] };
    }
    styles[type].count++;
    if (styles[type].samples.length < 3) {
      styles[type].samples.push(text.substring(0, 50));
    }
  }

  console.log("[copy-from-linked] 提取样式数: " + Object.keys(styles).length);

  // 关闭源文档
  try {
    sourceDoc.Close(false);
    console.log("[copy-from-linked] 已关闭源文档");
  } catch (e) {
    console.log("[copy-from-linked] 关闭源文档失败: " + String(e));
  }

  // ============================================================
  // 步骤2: 应用样式到目标文档
  // ============================================================
  function applyFormat(para, fmt) {
    try {
      var range = para.Range;
      var paraFormat = para.Format;

      if (fmt.fontCN) range.Font.NameFarEast = fmt.fontCN;
      if (fmt.fontEN) range.Font.NameAscii = fmt.fontEN;
      if (fmt.fontSize) range.Font.Size = fmt.fontSize;
      if (fmt.bold !== undefined) range.Font.Bold = fmt.bold;
      if (fmt.italic !== undefined) range.Font.Italic = fmt.italic;

      if (fmt.alignment !== undefined) paraFormat.Alignment = fmt.alignment;
      if (fmt.firstLineIndent) paraFormat.FirstLineIndent = fmt.firstLineIndent;
      if (fmt.leftIndent) paraFormat.LeftIndent = fmt.leftIndent;
      if (fmt.spaceBefore) paraFormat.SpaceBefore = fmt.spaceBefore;
      if (fmt.spaceAfter) paraFormat.SpaceAfter = fmt.spaceAfter;
      if (fmt.lineSpacing) paraFormat.LineSpacing = fmt.lineSpacing;
      if (fmt.lineSpacingRule !== undefined) paraFormat.LineSpacingRule = fmt.lineSpacingRule;
    } catch (e) {
      console.log("[copy-from-linked] 格式应用失败: " + String(e));
    }
  }

  var targetParas = targetDoc.Paragraphs;
  var appliedCounts = {};
  lastAppendixTitle = false;

  for (var t in styles) appliedCounts[t] = 0;

  for (var i = 1; i <= targetParas.Count; i++) {
    var para = targetParas.Item(i);
    var text = clean(para.Range.Text);
    if (!text) continue;

    var rng = para.Range;
    var currentFmt = {
      bold: rng.Font.Bold ? true : false,
      fontSize: rng.Font.Size,
      firstLineIndent: para.Format.FirstLineIndent || 0
    };

    var type = detectType(text, isPaper, currentFmt);

    if (styles[type] && styles[type].format) {
      applyFormat(para, styles[type].format);
      appliedCounts[type]++;
    }
  }

  // ============================================================
  // 统计并返回
  // ============================================================
  var totalApplied = 0;
  var detailLines = [];

  for (var t in appliedCounts) {
    if (appliedCounts[t] > 0) {
      totalApplied += appliedCounts[t];
      detailLines.push(styles[t].name + ": " + appliedCounts[t] + "处");
    }
  }

  var lines = [];
  lines.push("✅ 样式复制完成！");
  lines.push("📋 源文档：" + sourcePath.split('/').pop());
  lines.push("📄 目标文档：" + targetDoc.Name);
  lines.push("📊 共处理 " + totalApplied + " 个段落");
  if (detailLines.length > 0) {
    lines.push("\n应用详情：\n" + detailLines.join("\n"));
  }

  console.log("[copy-from-linked] 完成，共应用: " + totalApplied);

  return {
    success: true,
    message: lines.join("\n"),
    applied: appliedCounts,
    totalApplied: totalApplied
  };

} catch (e) {
  return { success: false, error: String(e) };
}