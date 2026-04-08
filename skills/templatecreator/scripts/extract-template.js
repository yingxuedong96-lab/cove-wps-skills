/**
 * extract-template.js - 完整样式提取（符合样式元素规范表）
 * 版本: 26.0409.1000
 */
try {
  var VER = "26.0409.1000";
  console.log("[extract] 版本: " + VER);

  var DOC = Application.ActiveDocument;
  if (!DOC) return { success: false, error: "没有打开的文档" };

  var docTypeParam = typeof docType !== 'undefined' ? docType : '';
  if (!docTypeParam) {
    return {
      success: true,
      needUserInput: true,
      question: "请选择文档类型：",
      options: ["论文/技术报告", "公文"]
    };
  }

  var isPaper = docTypeParam === "论文/技术报告" || docTypeParam === "paper";
  console.log("[extract] 类型: " + docTypeParam);

  // === 元素定义（根据样式元素规范表）===
  var ELEMENTS = {
    // 论文报告元素
    paper: {
      chapterTitle: { name: "章标题", pattern: /^第[一二三四五六七八九十\d]+章/ },
      heading1: { name: "一级标题", pattern: /^\d+\s+[^\d\.]/ },
      heading2: { name: "二级标题", pattern: /^\d+\.\d+\s/ },
      heading3: { name: "三级标题", pattern: /^\d+\.\d+\.\d+\s/ },
      heading4: { name: "四级标题", pattern: /^\d+\.\d+\.\d+\.\d+\s/ },
      heading5: { name: "五级标题", pattern: /^[\(（]\d+[\)）]|^[\(（][a-z][\)）]|^[①②③④⑤⑥⑦⑧⑨⑩]/ },
      body: { name: "正文" },
      figureCaption: { name: "图名", pattern: /^图\s*\d+/ },
      tableCaption: { name: "表名", pattern: /^表\s*\d+/ },
      appendixTitle: { name: "附录标题", pattern: /^附\s*录\s*[A-Z]/ },
      appendixSection: { name: "附录节题", pattern: /^[A-Z]\.\d+\s/ },
      referenceTitle: { name: "参考文献标题", pattern: /^参考文[献獻]/ },
      reference: { name: "参考文献条目", pattern: /^\[\d+\]/ },
      tocTitle: { name: "目录标题", pattern: /^目\s*录$|^目次$/ },
      abstract: { name: "摘要", pattern: /^摘\s*要|^Abstract/i },
      keyword: { name: "关键词", pattern: /^关键词|^关键字|^Key\s*words/i },
      formulaNote: { name: "公式说明", pattern: /^式中|^注\s*\d*/ }
    },
    // 公文元素
    official: {
      docTitle: { name: "公文标题", pattern: /^关于|通知$|决定$|意见$|办法$|规定$|批复$|请示$|报告$/ },
      heading1: { name: "一级标题", pattern: /^[一二三四五六七八九十]+、/ },
      heading2: { name: "二级标题", pattern: /^[\(（][一二三四五六七八九十]+[\)）]/ },
      heading3: { name: "三级标题", pattern: /^\d+\.\s/ },
      body: { name: "正文" },
      attachment: { name: "附件说明", pattern: /^附\s*件/ },
      signature: { name: "发文机关署名" },
      signDate: { name: "成文日期", pattern: /\d{4}年\d{1,2}月\d{1,2}日$/ },
      copySender: { name: "抄送机关", pattern: /^抄送/ },
      figureCaption: { name: "图名", pattern: /^图\s*\d+/ },
      tableCaption: { name: "表名", pattern: /^表\s*\d+/ }
    }
  };

  // === 工具函数 ===
  function clean(t) {
    return String(t || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function ptToSize(pt) {
    var map = { 42: "初号", 36: "小初", 26: "一号", 24: "小一", 22: "二号", 18: "小二", 16: "三号", 15: "小三", 14: "四号", 12: "小四", 10.5: "五号", 9: "小五", 7.5: "六号", 6.5: "小六" };
    for (var k in map) if (Math.abs(pt - parseFloat(k)) < 0.5) return map[k];
    return pt ? pt.toFixed(1) + "pt" : "-";
  }

  function detectElement(text, elementType) {
    var elements = ELEMENTS[elementType];
    for (var id in elements) {
      var el = elements[id];
      if (el.pattern && el.pattern.test(text)) return id;
    }
    return null;
  }

  // === 提取段落格式 ===
  function extractFormat(para) {
    var rng = para.Range;
    var fmt = para.Format;

    return {
      // 字体相关
      fontCN: rng.Font.NameFarEast || rng.Font.Name || "",
      fontEN: rng.Font.NameAscii || rng.Font.Name || "",
      fontSize: rng.Font.Size,
      fontSizeName: ptToSize(rng.Font.Size),
      bold: rng.Font.Bold ? true : false,
      italic: rng.Font.Italic ? true : false,
      underline: rng.Font.Underline ? true : false,
      color: rng.Font.Color ? ("000000" + rng.Font.Color.toString(16)).slice(-6).toUpperCase() : "",
      // 段落格式
      alignment: fmt.Alignment,
      firstLineIndent: fmt.FirstLineIndent ? (fmt.FirstLineIndent / 240).toFixed(1) : 0,
      leftIndent: fmt.LeftIndent ? (fmt.LeftIndent / 240).toFixed(1) : 0,
      rightIndent: fmt.RightIndent ? (fmt.RightIndent / 240).toFixed(1) : 0,
      // 间距
      spaceBefore: fmt.SpaceBefore,
      spaceAfter: fmt.SpaceAfter,
      lineSpacing: fmt.LineSpacing,
      lineSpacingRule: fmt.LineSpacingRule
    };
  }

  // === 主提取逻辑 ===
  var paras = DOC.Paragraphs;
  var styles = {};
  var elementSet = isPaper ? ELEMENTS.paper : ELEMENTS.official;

  // 初始化所有已定义元素
  for (var id in elementSet) {
    styles[id] = {
      id: id,
      name: elementSet[id].name,
      count: 0,
      formats: [],
      samples: []
    };
  }

  for (var i = 1; i <= paras.Count; i++) {
    var para = paras.Item(i);
    var text = clean(para.Range.Text);
    if (!text) continue;

    var fmt = extractFormat(para);
    var type = detectElement(text, isPaper ? "paper" : "official");

    // 智能推断
    if (!type) {
      if (fmt.bold && fmt.fontSize >= 16) type = isPaper ? "heading1" : "docTitle";
      else if (fmt.bold && fmt.fontSize >= 14) type = "heading2";
      else type = "body";
    }

    if (!styles[type]) {
      styles[type] = { id: type, name: type, count: 0, formats: [], samples: [] };
    }
    styles[type].count++;
    styles[type].formats.push(fmt);
    if (styles[type].samples.length < 3) styles[type].samples.push(text.substring(0, 50));
  }

  console.log("[extract] 样式数: " + Object.keys(styles).length);

  // === 生成输出 ===
  var alignMap = { 0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐", 4: "分散对齐" };
  var lineRuleMap = { 0: "单倍行距", 1: "1.5倍行距", 2: "2倍行距", 3: "最小值", 4: "固定值", 5: "多倍行距" };

  var lines = [];
  lines.push("✅ 样式模板提取完成！");
  lines.push("📄 源文档：" + DOC.Name);
  lines.push("📑 类型：" + (isPaper ? "论文报告样式" : "公文样式"));
  lines.push("📊 共 " + Object.keys(styles).filter(function(t) { return styles[t].count > 0; }).length + " 种样式\n");

  // 只显示有内容的样式
  var sortedTypes = Object.keys(styles).filter(function(t) { return styles[t].count > 0; })
    .sort(function(a, b) { return styles[b].count - styles[a].count; });

  lines.push("## 提取的样式详情\n");

  sortedTypes.forEach(function(t) {
    var s = styles[t], fmt = s.formats[0];
    lines.push("### " + s.name + "（" + s.count + "处）");

    var parts = [];
    if (fmt.fontCN) parts.push("字体: " + fmt.fontCN);
    if (fmt.fontSizeName && fmt.fontSizeName !== "-") parts.push("字号: " + fmt.fontSizeName);
    else if (fmt.fontSize) parts.push("字号: " + fmt.fontSize.toFixed(1) + "pt");
    if (fmt.bold) parts.push("加粗");
    if (fmt.italic) parts.push("斜体");
    if (fmt.underline) parts.push("下划线");

    if (fmt.alignment !== undefined) parts.push("对齐: " + (alignMap[fmt.alignment] || "未知"));
    if (parseFloat(fmt.firstLineIndent) > 0) parts.push("首行缩进: " + fmt.firstLineIndent + "字符");
    if (parseFloat(fmt.leftIndent) > 0) parts.push("左缩进: " + fmt.leftIndent + "字符");
    if (parseFloat(fmt.rightIndent) > 0) parts.push("右缩进: " + fmt.rightIndent + "字符");
    if (fmt.spaceBefore > 0) parts.push("段前: " + (fmt.spaceBefore / 20).toFixed(1) + "pt");
    if (fmt.spaceAfter > 0) parts.push("段后: " + (fmt.spaceAfter / 20).toFixed(1) + "pt");
    if (fmt.lineSpacing > 0) parts.push("行距: " + (fmt.lineSpacing / 20).toFixed(1) + "pt");

    lines.push(parts.join(" | ") + "\n");
  });

  // 页面设置
  lines.push("## 页面设置");
  lines.push("纸张: " + (DOC.PageSetup.PaperSize === 1 ? "A4" : DOC.PageSetup.PaperSize));
  lines.push("上边距: " + (DOC.PageSetup.TopMargin / 567).toFixed(2) + "cm | 下边距: " + (DOC.PageSetup.BottomMargin / 567).toFixed(2) + "cm");
  lines.push("左边距: " + (DOC.PageSetup.LeftMargin / 567).toFixed(2) + "cm | 右边距: " + (DOC.PageSetup.RightMargin / 567).toFixed(2) + "cm");

  // 页眉页脚
  try {
    var header = DOC.Sections.Item(1).Headers.Item(1);
    var footer = DOC.Sections.Item(1).Footers.Item(1);
    if (header.Range.Text && clean(header.Range.Text)) {
      lines.push("\n页眉: " + clean(header.Range.Text).substring(0, 30));
    }
  } catch(e) {}

  lines.push("\n模板已保存，可在后续"应用模板"时使用。");

  return {
    success: true,
    message: lines.join("\n"),
    styleCount: sortedTypes.length
  };
} catch (e) {
  return { success: false, error: String(e) };
}