/**
 * copy-style.js - 样式复制（提取或应用）
 * 版本: 26.0410.1007
 *
 * 智能模式：
 * - 如果收到 templateJson 参数 → 应用模式
 * - 如果没有 templateJson 参数 → 提取模式
 *
 * 用户流程：
 * 1. A文档点击「样式复制」→ 提取并提示"样式已提取，请切换到目标文档"
 * 2. B文档点击「样式复制」→ 应用样式
 */
try {
  var VER = "26.0410.1007";
  console.log("[copy-style] 版本: " + VER);

  var DOC = Application.ActiveDocument;
  if (!DOC) return { success: false, error: "没有打开的文档" };

  // 检查是否有传入的模板数据
  var templateJson = typeof templateJson !== 'undefined' ? templateJson : null;

  if (templateJson) {
    // ========== 应用模式 ==========
    return applyTemplate(DOC, templateJson);
  } else {
    // ========== 提取模式 ==========
    return extractTemplate(DOC);
  }

  // ============================================================
  // 提取模板
  // ============================================================
  function extractTemplate(DOC) {
    console.log("[copy-style] 模式: 提取");

    // 自动检测文档类型
    var paperScore = 0, govScore = 0;
    var paras = DOC.Paragraphs;
    var scanCount = Math.min(30, paras.Count);

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
    var docType = isPaper ? "论文报告" : "公文";
    console.log("[copy-style] 类型: " + docType);

    // 样式名称
    var STYLE_NAMES = {
      coverTitle: "封面标题", chapterTitle: "章标题",
      heading1: "一级标题", heading2: "二级标题", heading3: "三级标题",
      heading4: "四级标题", heading5: "五级标题",
      body: "正文", listItem: "列表项",
      figureCaption: "图名", tableCaption: "表名",
      appendixTitle: "附录标题", appendixSection: "附录节题",
      docTitle: "公文标题", attachment: "附件说明"
    };

    var alignMap = { 0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐" };
    var lineRuleMap = { 0: "单倍行距", 4: "固定值", 5: "多倍行距" };

    // 上下文状态
    var lastAppendixTitle = false;

    // 检测样式
    function detectStyle(text, isPaper, fmt) {
      if (isPaper) {
        if (/^第[一二三四五六七八九十百零\d]+章/.test(text)) return "chapterTitle";
        if (/^\d+\.\d+\.\d+\.\d+\.\d+/.test(text)) return "heading5";
        if (/^\d+\.\d+\.\d+\.\d+/.test(text)) return "heading4";
        if (/^\d+\.\d+\.\d+[^.\d]/.test(text)) return "heading3";
        if (/^\d+\.\d+[^.\d]/.test(text)) return "heading2";
        if (/^\d+\s+[^\d\.\s]/.test(text)) return "heading1";
        if (/^附\s*录\s*[A-Z０-９]/.test(text)) { lastAppendixTitle = true; return "appendixTitle"; }
        if (/^[A-Z]\.\d+/.test(text)) return "appendixSection";
        if (/^图\s*\d+/.test(text)) return "figureCaption";
        if (/^表\s*\d+/.test(text)) return "tableCaption";
        if (/^参考文[献獻]/.test(text)) return "referenceTitle";
        if (/^\[\d+\]/.test(text)) return "reference";
        if (/^摘\s*要|^Abstract/i.test(text)) return "abstract";
        if (/^式中|^注\s*\d*/.test(text)) return "formulaNote";
        if (/^[\(（]\d+[\)）]/.test(text)) return "listItem";
        if (/^[\(（][a-z][\)）]/.test(text)) return "listItem";
        if (/^[①②③④⑤⑥⑦⑧⑨⑩]/.test(text)) return "listItem";
        // 上下文感知
        if (lastAppendixTitle && fmt && fmt.bold && fmt.fontSize >= 14) {
          lastAppendixTitle = false;
          return "appendixTitle";
        }
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
      return null;
    }

    // 提取格式
    function extractFormat(para) {
      var rng = para.Range, fmt = para.Format;
      return {
        fontCN: rng.Font.NameFarEast || rng.Font.Name || "",
        fontEN: rng.Font.NameAscii || "",
        fontSize: rng.Font.Size,
        bold: rng.Font.Bold ? true : false,
        alignment: fmt.Alignment,
        firstLineIndent: fmt.FirstLineIndent || 0,
        leftIndent: fmt.LeftIndent || 0,
        rightIndent: fmt.RightIndent || 0,
        spaceBefore: fmt.SpaceBefore || 0,
        spaceAfter: fmt.SpaceAfter || 0,
        lineSpacing: fmt.LineSpacing || 0,
        lineSpacingRule: fmt.LineSpacingRule
      };
    }

    // 主提取
    var styles = {};
    var formatIndex = {};

    for (var i = 1; i <= paras.Count; i++) {
      var para = paras.Item(i);
      var text = clean(para.Range.Text);
      if (!text) continue;

      var fmt = extractFormat(para);
      var type = detectStyle(text, isPaper, fmt);

      if (!type) {
        if (fmt.bold && fmt.fontSize >= 20) type = "coverTitle";
        else if (fmt.bold && fmt.fontSize >= 16) type = "heading1";
        else if (fmt.bold && fmt.fontSize >= 14) type = "heading2";
        else type = "body";
      }

      if (!styles[type]) {
        styles[type] = { id: type, name: STYLE_NAMES[type] || type, count: 0, format: null, samples: [] };
      }
      styles[type].count++;
      if (!styles[type].format) {
        styles[type].format = Object.assign({}, fmt, {
          alignmentName: alignMap[fmt.alignment] || "",
          lineSpacingRuleName: lineRuleMap[fmt.lineSpacingRule] || ""
        });
      }
      if (styles[type].samples.length < 3) {
        styles[type].samples.push(text.substring(0, 50));
      }
    }

    // 构建模板数据
    var templateData = {
      version: VER,
      docName: DOC.Name,
      docType: docType,
      extractTime: new Date().toISOString(),
      styles: styles
    };

    var jsonContent = JSON.stringify(templateData, null, 2);

    // 返回提取结果，包含 templateJson 供下次调用
    var lines = [];
    lines.push("✅ 样式提取完成！");
    lines.push("📄 源文档：" + DOC.Name);
    lines.push("📑 类型：" + docType);
    lines.push("📊 共 " + Object.keys(styles).length + " 种样式");
    lines.push("\n请切换到目标文档，再次点击「样式复制」即可应用。");

    return {
      success: true,
      message: lines.join("\n"),
      styleCount: Object.keys(styles).length,
      templateJson: jsonContent,
      mode: "extracted"
    };
  }

  // ============================================================
  // 应用模板
  // ============================================================
  function applyTemplate(DOC, templateJsonStr) {
    console.log("[copy-style] 模式: 应用");

    var template;
    try {
      template = typeof templateJsonStr === 'string' ? JSON.parse(templateJsonStr) : templateJsonStr;
    } catch (e) {
      return { success: false, error: "模板数据格式错误" };
    }

    if (!template || !template.styles) {
      return { success: false, error: "无效的模板数据" };
    }

    var styles = template.styles;
    var paras = DOC.Paragraphs;
    var appliedCounts = {};

    // 初始化计数
    for (var t in styles) {
      appliedCounts[t] = 0;
    }

    // 检测段落类型
    function detectType(text, styles) {
      if (/^第[一二三四五六七八九十百零\d]+章/.test(text)) return "chapterTitle";
      if (/^\d+\.\d+\.\d+\.\d+\.\d+/.test(text)) return "heading5";
      if (/^\d+\.\d+\.\d+\.\d+/.test(text)) return "heading4";
      if (/^\d+\.\d+\.\d+[^.\d]/.test(text)) return "heading3";
      if (/^\d+\.\d+[^.\d]/.test(text)) return "heading2";
      if (/^\d+\s+[^\d\.\s]/.test(text)) return "heading1";
      if (/^附\s*录\s*[A-Z]/.test(text)) return "appendixTitle";
      if (/^[A-Z]\.\d+/.test(text)) return "appendixSection";
      if (/^图\s*\d+/.test(text)) return "figureCaption";
      if (/^表\s*\d+/.test(text)) return "tableCaption";
      if (/^[\(（]\d+[\)）]/.test(text)) return "listItem";
      return "body";
    }

    // 应用格式
    function applyFormat(para, format) {
      try {
        var range = para.Range;
        var paraFormat = para.Format;

        if (format.fontCN) range.Font.NameFarEast = format.fontCN;
        if (format.fontEN) range.Font.NameAscii = format.fontEN;
        if (format.fontSize) range.Font.Size = format.fontSize;
        if (format.bold !== undefined) range.Font.Bold = format.bold;
        if (format.alignment !== undefined) paraFormat.Alignment = format.alignment;
        if (format.firstLineIndent) paraFormat.FirstLineIndent = format.firstLineIndent;
        if (format.leftIndent) paraFormat.LeftIndent = format.leftIndent;
        if (format.spaceBefore) paraFormat.SpaceBefore = format.spaceBefore;
        if (format.spaceAfter) paraFormat.SpaceAfter = format.spaceAfter;
        if (format.lineSpacing) paraFormat.LineSpacing = format.lineSpacing;
        if (format.lineSpacingRule !== undefined) paraFormat.LineSpacingRule = format.lineSpacingRule;
      } catch (e) {
        console.log("[copy-style] 应用格式失败: " + String(e));
      }
    }

    // 遍历段落应用样式
    for (var i = 1; i <= paras.Count; i++) {
      var para = paras.Item(i);
      var text = clean(para.Range.Text);
      if (!text) continue;

      var type = detectType(text, styles);
      if (styles[type] && styles[type].format) {
        applyFormat(para, styles[type].format);
        appliedCounts[type]++;
      }
    }

    // 统计
    var totalApplied = 0;
    var detailLines = [];
    for (var t in appliedCounts) {
      if (appliedCounts[t] > 0) {
        totalApplied += appliedCounts[t];
        detailLines.push(styles[t].name + ": " + appliedCounts[t] + "处");
      }
    }

    var lines = [];
    lines.push("✅ 样式应用完成！");
    lines.push("📄 目标文档：" + DOC.Name);
    lines.push("📊 共处理 " + totalApplied + " 个段落");
    if (detailLines.length > 0) {
      lines.push("\n应用详情：\n" + detailLines.join("\n"));
    }

    return {
      success: true,
      message: lines.join("\n"),
      applied: appliedCounts,
      totalApplied: totalApplied,
      mode: "applied"
    };
  }

  // ============================================================
  // 工具函数
  // ============================================================
  function clean(t) {
    return String(t || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

} catch (e) {
  return { success: false, error: String(e) };
}