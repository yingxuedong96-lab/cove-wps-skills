/**
 * extract-template.js - 让框架自动包装 success
 * 版本: 26.0408.1845
 */
(function() {
  var VER = "26.0408.1845";
  console.log("[extract] 版本: " + VER);

  var DOC = Application.ActiveDocument;
  if (!DOC) return JSON.stringify({ error: "没有打开的文档" });

  var docTypeParam = typeof docType !== 'undefined' ? docType : '';
  if (!docTypeParam) {
    return JSON.stringify({
      needUserInput: true,
      question: "请选择文档类型：",
      options: ["论文/技术报告", "公文"]
    });
  }

  var isPaper = docTypeParam === "论文/技术报告" || docTypeParam === "paper";
  console.log("[extract] 类型: " + docTypeParam + ", 检测到样式数: ");

  var paras = DOC.Paragraphs;
  var styles = {};

  function clean(t) {
    return String(t || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function detectPaper(text) {
    if (/^\d+\.\d+\.\d+\.\d+\.\d+/.test(text)) return "heading5";
    if (/^\d+\.\d+\.\d+\.\d+/.test(text)) return "heading4";
    if (/^\d+\.\d+\.\d+/.test(text)) return "heading3";
    if (/^\d+\.\d+[^.\d]/.test(text)) return "heading2";
    if (/^\d+[^.\d]/.test(text)) return "heading1";
    if (/^第[一二三四五六七八九十\d]+章/.test(text)) return "chapterTitle";
    if (/^图\s*\d+/.test(text)) return "figureCaption";
    if (/^表\s*\d+/.test(text)) return "tableCaption";
    if (/^附\s*录/.test(text)) return "appendixTitle";
    if (/^[A-Z]\.\d+/.test(text)) return "appendixSection";
    if (/^参考文献/.test(text)) return "referenceTitle";
    if (/^\[\d+\]/.test(text)) return "reference";
    return null;
  }

  var styleNames = {
    heading1: "一级标题", heading2: "二级标题", heading3: "三级标题",
    heading4: "四级标题", heading5: "五级标题", chapterTitle: "章标题",
    docTitle: "论文标题", body: "正文", figureCaption: "图名", tableCaption: "表名",
    appendixTitle: "附录标题", appendixSection: "附录节题",
    referenceTitle: "参考文献标题", reference: "参考文献条目"
  };

  for (var i = 1; i <= paras.Count; i++) {
    var para = paras.Item(i);
    var text = clean(para.Range.Text);
    if (!text) continue;

    var fmt = {
      fontCN: para.Range.Font.NameFarEast || para.Range.Font.Name,
      fontSize: para.Range.Font.Size,
      bold: para.Range.Font.Bold,
      align: para.Format.Alignment,
      firstLineIndent: para.Format.FirstLineIndent / 240,
      lineSpacing: para.Format.LineSpacing
    };

    var type = isPaper ? detectPaper(text) : null;
    if (!type && fmt.bold && fmt.fontSize >= 14) type = "heading1";
    if (!type) type = "body";

    if (!styles[type]) styles[type] = { name: styleNames[type] || type, count: 0, formats: [], samples: [] };
    styles[type].count++;
    styles[type].formats.push(fmt);
    if (styles[type].samples.length < 3) styles[type].samples.push(text.substring(0, 40));
  }

  console.log("[extract] 样式数: " + Object.keys(styles).length);

  var alignMap = { 0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐" };
  var lines = [];

  lines.push("✅ 样式模板提取完成！");
  lines.push("📄 源文档：" + DOC.Name);
  lines.push("📊 共 " + Object.keys(styles).length + " 种样式\n");

  var typeOrder = ["heading1", "heading2", "heading3", "heading4", "heading5", "body", "figureCaption", "tableCaption", "appendixTitle"];
  typeOrder.forEach(function(t) {
    if (!styles[t]) return;
    var s = styles[t], fmt = s.formats[0];
    lines.push("### " + s.name + "（" + s.count + "处）");
    var p = [];
    if (fmt.fontCN) p.push("字体: " + fmt.fontCN);
    if (fmt.fontSize) p.push("字号: " + fmt.fontSize + "pt");
    if (fmt.bold) p.push("加粗");
    if (fmt.align !== undefined) p.push("对齐: " + alignMap[fmt.align]);
    if (fmt.firstLineIndent) p.push("首行缩进: " + fmt.firstLineIndent.toFixed(1) + "字符");
    lines.push(p.join(" | ") + "\n");
  });

  lines.push("## 页面设置");
  lines.push("- 上边距: " + (DOC.PageSetup.TopMargin / 567).toFixed(2) + "cm");
  lines.push("- 下边距: " + (DOC.PageSetup.BottomMargin / 567).toFixed(2) + "cm");

  var stylesTable = Object.keys(styles).map(function(t) {
    var s = styles[t], fmt = s.formats[0];
    return { 样式名称: s.name, 出现次数: s.count + "处", 字体: fmt.fontCN || "-", 字号: fmt.fontSize ? fmt.fontSize + "pt" : "-", 加粗: fmt.bold ? "是" : "否" };
  });

  var template = {
    name: DOC.Name.replace(/\.(docx|doc)$/i, '') + '_模板',
    docType: isPaper ? "paper" : "official",
    styles: Object.keys(styles).map(function(t) { return { id: t, name: styles[t].name, count: styles[t].count, format: styles[t].formats[0] }; }),
    pageSetup: { topMargin: DOC.PageSetup.TopMargin / 567, bottomMargin: DOC.PageSetup.BottomMargin / 567, leftMargin: DOC.PageSetup.LeftMargin / 567, rightMargin: DOC.PageSetup.RightMargin / 567 }
  };

  // 返回不带 success 的 JSON，让框架自动包装成 {success: true, data: ...}
  return JSON.stringify({
    scriptVersion: VER,
    message: lines.join("\n"),
    stylesTable: stylesTable,
    templateJson: template,
    pageSetup: { paperSize: "A4", topMargin: (DOC.PageSetup.TopMargin / 567).toFixed(2) + "cm", bottomMargin: (DOC.PageSetup.BottomMargin / 567).toFixed(2) + "cm" }
  });
})();