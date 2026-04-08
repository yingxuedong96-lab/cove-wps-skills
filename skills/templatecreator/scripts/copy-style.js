/**
 * copy-style.js - 智能样式复制
 * 版本: 26.0410.1013
 *
 * 模式：
 * - 有 sourceDocPath 参数 → 从关联文档复制样式到当前文档
 * - 无参数 → 从当前文档提取样式
 */
try {
  var VER = "26.0410.1013";
  console.log("[copy-style] 版本: " + VER);

  var targetDoc = Application.ActiveDocument;
  if (!targetDoc) return { success: false, error: "没有打开的文档" };

  var sourcePath = typeof sourceDocPath !== 'undefined' ? sourceDocPath : '';

  // ============================================================
  // 工具函数
  // ============================================================
  function clean(t) {
    return String(t || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  var STYLE_NAMES = {
    chapterTitle: "章标题", heading1: "一级标题", heading2: "二级标题",
    heading3: "三级标题", heading4: "四级标题", heading5: "五级标题",
    body: "正文", listItem: "列表项", figureCaption: "图名", tableCaption: "表名",
    appendixTitle: "附录标题", appendixSection: "附录节题",
    docTitle: "公文标题", attachment: "附件说明", coverTitle: "封面标题"
  };

  function detectType(text, isPaper, fmt) {
    if (isPaper) {
      if (/^第[一二三四五六七八九十百零\d]+章/.test(text)) return "chapterTitle";
      if (/^\d+\.\d+\.\d+\.\d+\.\d+/.test(text)) return "heading5";
      if (/^\d+\.\d+\.\d+\.\d+/.test(text)) return "heading4";
      if (/^\d+\.\d+\.\d+[^.\d]/.test(text)) return "heading3";
      if (/^\d+\.\d+[^.\d]/.test(text)) return "heading2";
      if (/^\d+\s+[^\d\.\s]/.test(text)) return "heading1";
      if (/^附\s*录\s*[A-Z０-９]/.test(text)) return "appendixTitle";
      if (/^[A-Z]\.\d+/.test(text)) return "appendixSection";
      if (/^图\s*\d+/.test(text)) return "figureCaption";
      if (/^表\s*\d+/.test(text)) return "tableCaption";
      if (/^[\(（]\d+[\)）]/.test(text)) return "listItem";
      if (/^[\(（][a-z][\)）]/.test(text)) return "listItem";
      if (/^[①②③④⑤⑥⑦⑧⑨⑩]/.test(text)) return "listItem";
    } else {
      if (/^关于|通知$|决定$|意见$|办法$|规定$/.test(text)) return "docTitle";
      if (/^[一二三四五六七八九十]+、/.test(text)) return "heading1";
      if (/^[\(（][一二三四五六七八九十]+[\)）]/.test(text)) return "heading2";
      if (/^\d+\.\s/.test(text)) return "heading3";
      if (/^图\s*\d+/.test(text)) return "figureCaption";
      if (/^表\s*\d+/.test(text)) return "tableCaption";
    }
    if (fmt && fmt.firstLineIndent < 0) return "listItem";
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

  function applyFormat(para, fmt) {
    try {
      var range = para.Range, paraFormat = para.Format;
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
    } catch (e) {}
  }

  // ============================================================
  // 提取样式函数
  // ============================================================
  function extractStyles(doc) {
    var paras = doc.Paragraphs;
    var paperScore = 0, govScore = 0;
    var scanCount = Math.min(30, paras.Count);

    for (var i = 1; i <= scanCount; i++) {
      var text = clean(paras.Item(i).Range.Text);
      if (!text) continue;
      if (/^第[一二三四五六七八九十百零\d]+章/.test(text)) paperScore += 3;
      if (/^\d+\.\d+/.test(text)) paperScore += 2;
      if (/^附\s*录/.test(text)) paperScore += 2;
      if (/^摘\s*要|^Abstract/i.test(text)) paperScore += 2;
      if (/^关于/.test(text)) govScore += 3;
      if (/通知$|决定$|意见$|办法$|规定$/.test(text)) govScore += 2;
    }

    var isPaper = paperScore >= govScore;
    var styles = { isPaper: isPaper };

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

    return styles;
  }

  // ============================================================
  // 应用样式函数
  // ============================================================
  function applyStyles(doc, styles) {
    var paras = doc.Paragraphs;
    var isPaper = styles.isPaper;
    var appliedCounts = {};

    for (var t in styles) {
      if (t !== 'isPaper') appliedCounts[t] = 0;
    }

    for (var i = 1; i <= paras.Count; i++) {
      var para = paras.Item(i);
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

    var totalApplied = 0;
    var detailLines = [];
    for (var t in appliedCounts) {
      if (appliedCounts[t] > 0) {
        totalApplied += appliedCounts[t];
        detailLines.push((styles[t] ? styles[t].name : STYLE_NAMES[t] || t) + ": " + appliedCounts[t] + "处");
      }
    }

    return { totalApplied: totalApplied, detail: detailLines.join("\n") };
  }

  // ============================================================
  // 主逻辑
  // ============================================================
  if (sourcePath) {
    // 模式1：从关联文档复制样式
    console.log("[copy-style] 模式：从关联文档复制");
    console.log("[copy-style] 源文档: " + sourcePath);

    var sourceDoc;
    try {
      sourceDoc = Application.Documents.Open(sourcePath);
      console.log("[copy-style] 成功打开源文档");
    } catch (e) {
      return { success: false, error: "无法打开源文档: " + String(e) };
    }

    var styles = extractStyles(sourceDoc);
    console.log("[copy-style] 提取样式数: " + (Object.keys(styles).length - 1));

    try { sourceDoc.Close(false); } catch (e) {}

    var result = applyStyles(targetDoc, styles);

    return {
      success: true,
      message: "✅ 样式复制完成！\n📋 源文档：" + sourcePath.split('/').pop() + "\n📄 目标文档：" + targetDoc.Name + "\n📊 共处理 " + result.totalApplied + " 个段落\n\n" + result.detail
    };

  } else {
    // 模式2：从当前文档提取样式
    console.log("[copy-style] 模式：提取样式");
    console.log("[copy-style] 文档: " + targetDoc.Name);

    var styles = extractStyles(targetDoc);
    var styleCount = Object.keys(styles).length - 1;
    console.log("[copy-style] 提取样式数: " + styleCount);

    var detailLines = [];
    for (var t in styles) {
      if (t !== 'isPaper' && styles[t].count > 0) {
        detailLines.push(styles[t].name + ": " + styles[t].count + "处");
      }
    }

    return {
      success: true,
      message: "✅ 样式提取完成！\n📄 文档：" + targetDoc.Name + "\n📑 类型：" + (styles.isPaper ? "论文报告" : "公文") + "\n📊 共 " + styleCount + " 种样式\n\n" + detailLines.slice(0, 8).join("\n")
    };
  }

} catch (e) {
  return { success: false, error: String(e) };
}