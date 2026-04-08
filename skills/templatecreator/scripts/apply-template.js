/**
 * apply-template.js - 将样式模板应用到当前文档
 * 版本: 26.0410.1008
 * 参数:
 *   - templateJson: 直接传入模板JSON字符串（从artifact获取）
 */
try {
  var VER = "26.0410.1008";
  console.log("[apply] 版本: " + VER);

  var DOC = Application.ActiveDocument;
  if (!DOC) return { success: false, error: "没有打开的文档" };

  // 获取参数（直接作为脚本变量传入）
  var templateJsonStr = typeof templateJson !== 'undefined' ? templateJson : null;
  if (!templateJsonStr) {
    return { success: false, error: "缺少 templateJson 参数，请先执行「提取模板」" };
  }

  // 解析模板数据
  var template;
  try {
    template = typeof templateJsonStr === 'string' ? JSON.parse(templateJsonStr) : templateJsonStr;
  } catch (e) {
    return { success: false, error: "模板数据格式错误: " + String(e) };
  }

  if (!template || !template.styles) {
    return { success: false, error: "无效的模板数据" };
  }

  console.log("[apply] 模板来源: " + template.docName);
  console.log("[apply] 类型: " + template.docType);

  // ============================================================
  // 工具函数
  // ============================================================
  function clean(t) {
    return String(t || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  // ============================================================
  // 检测段落类型（与提取脚本保持一致）
  // ============================================================
  var lastAppendixTitle = false;

  function detectType(text, isPaper, fmt) {
    if (isPaper) {
      // 论文报告检测
      if (/^第[一二三四五六七八九十百零\d]+章/.test(text)) return "chapterTitle";
      if (/^\d+\.\d+\.\d+\.\d+\.\d+/.test(text)) return "heading5";
      if (/^\d+\.\d+\.\d+\.\d+/.test(text)) return "heading4";
      if (/^\d+\.\d+\.\d+[^.\d]/.test(text)) return "heading3";
      if (/^\d+\.\d+[^.\d]/.test(text)) return "heading2";
      if (/^\d+\s+[^\d\.\s]/.test(text)) return "heading1";
      if (/^附\s*录\s*[A-Z０-９０-９]/.test(text)) {
        lastAppendixTitle = true;
        return "appendixTitle";
      }
      if (/^[A-Z]\.\d+/.test(text)) return "appendixSection";
      if (/^图\s*\d+/.test(text)) return "figureCaption";
      if (/^表\s*\d+/.test(text)) return "tableCaption";
      if (/^\([\d.-]+\)$/.test(text)) return "formulaCaption";
      if (/^参考文[献獻]/.test(text)) return "referenceTitle";
      if (/^\[\d+\]/.test(text)) return "reference";
      if (/^目\s*录$|^目次$/.test(text)) return "tocTitle";
      if (/^摘\s*要|^Abstract/i.test(text)) return "abstract";
      if (/^关键词|^关键字|^Key\s*words/i.test(text)) return "keyword";
      if (/^式中/.test(text)) return "formulaNote";
      if (/^注\s*\d*|^注\s*：/.test(text)) return "note";
      if (/^\d{4}年\d{1,2}月\d{1,2}日$/.test(text)) return "coverDate";
      if (/^[\(（]\d+[\)）]/.test(text)) return "listItem";
      if (/^[\(（][a-z][\)）]/.test(text)) return "listItem";
      if (/^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮]/.test(text)) return "listItem";
      if (/^[a-z][\)）\.]/.test(text)) return "listItem";
      // 上下文感知
      if (lastAppendixTitle && fmt && fmt.bold && fmt.fontSize >= 14) {
        lastAppendixTitle = false;
        return "appendixTitle";
      }
    } else {
      // 公文检测
      if (/^关于|通知$|决定$|意见$|办法$|规定$|批复$|请示$|报告$/.test(text)) return "docTitle";
      if (/^[一二三四五六七八九十]+、/.test(text)) return "heading1";
      if (/^[\(（][一二三四五六七八九十]+[\)）]/.test(text)) return "heading2";
      if (/^\d+\.\s/.test(text)) return "heading3";
      if (/^附\s*件/.test(text)) return "attachment";
      if (/\d{4}年\d{1,2}月\d{1,2}日$/.test(text)) return "signDate";
      if (/^抄送/.test(text)) return "copySender";
      if (/^印发机关|^印发单位/.test(text)) return "issuerDept";
      if (/^印发日期/.test(text)) return "issueDate";
      if (/^图\s*\d+/.test(text)) return "figureCaption";
      if (/^表\s*\d+/.test(text)) return "tableCaption";
      if (/^[\(（]\d+[\)）]/.test(text)) return "listItem";
      if (/^[a-z][\)）\.]/.test(text)) return "listItem";
    }

    // 基于格式的推断
    if (fmt && fmt.firstLineIndent < 0) return "listItem";

    // 重置附录状态
    if (isPaper && !/^附\s*录/.test(text)) {
      lastAppendixTitle = false;
    }

    return "body";
  }

  // ============================================================
  // 应用格式到段落
  // ============================================================
  function applyFormat(para, fmt) {
    try {
      var range = para.Range;
      var paraFormat = para.Format;

      // 字体
      if (fmt.fontCN) range.Font.NameFarEast = fmt.fontCN;
      if (fmt.fontEN) range.Font.NameAscii = fmt.fontEN;
      if (fmt.fontSize) range.Font.Size = fmt.fontSize;
      if (fmt.bold !== undefined) range.Font.Bold = fmt.bold;
      if (fmt.italic !== undefined) range.Font.Italic = fmt.italic;
      if (fmt.underline !== undefined) range.Font.Underline = fmt.underline;

      // 段落格式
      if (fmt.alignment !== undefined) paraFormat.Alignment = fmt.alignment;
      if (fmt.firstLineIndent) paraFormat.FirstLineIndent = fmt.firstLineIndent;
      if (fmt.leftIndent) paraFormat.LeftIndent = fmt.leftIndent;
      if (fmt.rightIndent) paraFormat.RightIndent = fmt.rightIndent;

      // 间距
      if (fmt.spaceBefore) paraFormat.SpaceBefore = fmt.spaceBefore;
      if (fmt.spaceAfter) paraFormat.SpaceAfter = fmt.spaceAfter;
      if (fmt.lineSpacing) paraFormat.LineSpacing = fmt.lineSpacing;
      if (fmt.lineSpacingRule !== undefined) paraFormat.LineSpacingRule = fmt.lineSpacingRule;
    } catch (e) {
      console.log("[apply] 格式应用失败: " + String(e));
    }
  }

  // ============================================================
  // 主应用逻辑
  // ============================================================
  var styles = template.styles;
  var isPaper = template.docType === "论文报告";
  var paras = DOC.Paragraphs;
  var appliedCounts = {};

  // 初始化计数
  for (var t in styles) {
    appliedCounts[t] = 0;
  }

  // 统计总段落数
  var totalParas = 0;
  for (var i = 1; i <= paras.Count; i++) {
    var text = clean(paras.Item(i).Range.Text);
    if (text) totalParas++;
  }

  console.log("[apply] 文档总段落: " + totalParas);

  // 遍历段落应用样式
  for (var i = 1; i <= paras.Count; i++) {
    var para = paras.Item(i);
    var text = clean(para.Range.Text);
    if (!text) continue;

    // 获取当前格式（用于智能检测）
    var rng = para.Range;
    var currentFmt = {
      bold: rng.Font.Bold ? true : false,
      fontSize: rng.Font.Size,
      firstLineIndent: para.Format.FirstLineIndent || 0
    };

    // 检测类型
    var type = detectType(text, isPaper, currentFmt);

    // 查找对应的样式定义
    if (styles[type] && styles[type].format) {
      applyFormat(para, styles[type].format);
      appliedCounts[type]++;
    }
  }

  // ============================================================
  // 应用页面设置
  // ============================================================
  if (template.pageSetup) {
    var ps = DOC.PageSetup;
    try {
      // 边距转换：cm -> pt (1cm ≈ 28.35pt)
      if (template.pageSetup.topMargin) {
        var topCm = parseFloat(template.pageSetup.topMargin);
        if (topCm > 0) ps.TopMargin = topCm * 28.35;
      }
      if (template.pageSetup.bottomMargin) {
        var bottomCm = parseFloat(template.pageSetup.bottomMargin);
        if (bottomCm > 0) ps.BottomMargin = bottomCm * 28.35;
      }
      if (template.pageSetup.leftMargin) {
        var leftCm = parseFloat(template.pageSetup.leftMargin);
        if (leftCm > 0) ps.LeftMargin = leftCm * 28.35;
      }
      if (template.pageSetup.rightMargin) {
        var rightCm = parseFloat(template.pageSetup.rightMargin);
        if (rightCm > 0) ps.RightMargin = rightCm * 28.35;
      }
      console.log("[apply] 页面设置已应用");
    } catch (e) {
      console.log("[apply] 页面设置失败: " + String(e));
    }
  }

  // ============================================================
  // 统计并返回
  // ============================================================
  var totalApplied = 0;
  var detailLines = [];
  var styleNames = {
    chapterTitle: "章标题",
    heading1: "一级标题", heading2: "二级标题", heading3: "三级标题",
    heading4: "四级标题", heading5: "五级标题",
    body: "正文", listItem: "列表项",
    figureCaption: "图名", tableCaption: "表名",
    appendixTitle: "附录标题", appendixSection: "附录节题",
    docTitle: "公文标题", attachment: "附件说明",
    reference: "参考文献", referenceTitle: "参考文献标题",
    abstract: "摘要", keyword: "关键词"
  };

  for (var t in appliedCounts) {
    if (appliedCounts[t] > 0) {
      totalApplied += appliedCounts[t];
      var name = styles[t] ? (styles[t].name || styleNames[t] || t) : styleNames[t] || t;
      detailLines.push(name + ": " + appliedCounts[t] + "处");
    }
  }

  var lines = [];
  lines.push("✅ 样式模板应用完成！");
  lines.push("📄 目标文档：" + DOC.Name);
  lines.push("📋 模板来源：" + template.docName);
  lines.push("📊 共处理 " + totalApplied + " 个段落");
  if (detailLines.length > 0) {
    lines.push("\n应用详情：\n" + detailLines.join("\n"));
  }

  console.log("[apply] 完成，共应用: " + totalApplied);

  return {
    success: true,
    message: lines.join("\n"),
    applied: appliedCounts,
    totalApplied: totalApplied
  };

} catch (e) {
  return { success: false, error: String(e) };
}