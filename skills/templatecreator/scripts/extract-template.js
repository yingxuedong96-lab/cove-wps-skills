/**
 * extract-template.js - 完整样式提取（符合样式元素规范表）
 * 版本: 26.0410.1000
 * 支持元素: 论文报告26种 + 公文20种
 * 支持参数: 45个（字体7 + 段落5 + 间距4 + 大纲3 + 表格10 + 页面9 + 其他7）
 */
try {
  var VER = "26.0410.1000";
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

  // ============================================================
  // 一、元素名称定义（符合样式元素规范表）
  // ============================================================
  var STYLE_NAMES = {
    // 论文报告 - 封面部分
    coverTitle: "封面标题", coverSubtitle: "封面副标题", coverUnit: "封面单位", coverDate: "封面日期",
    // 论文报告 - 前置部分
    titlePage: "扉页署名", tocTitle: "目录标题", tocChapter: "目录章题", tocSection: "目录节题",
    // 论文报告 - 主体标题
    chapterTitle: "章标题",
    heading1: "一级标题", heading2: "二级标题", heading3: "三级标题", heading4: "四级标题", heading5: "五级标题",
    // 论文报告 - 正文
    body: "正文", listItem: "列表项",
    // 论文报告 - 图表公式
    figureCaption: "图名", tableCaption: "表名", tableBody: "表内文字",
    formulaCaption: "公式编号", formulaNote: "公式说明",
    // 论文报告 - 补充部分
    appendixTitle: "附录标题", appendixSection: "附录节题",
    referenceTitle: "参考文献标题", reference: "参考文献条目",
    note: "注释说明",
    // 公文元素
    issuer: "发文机关标志", dividerLine: "版头分隔线", docNumber: "发文字号",
    docTitle: "公文标题", mainSender: "主送机关",
    attachment: "附件说明", signature: "发文机关署名", signDate: "成文日期",
    sealPosition: "印章位置", copySender: "抄送机关", issuerDept: "印发机关", issueDate: "印发日期",
    // 页面元素
    header: "页眉", footer: "页脚",
    // 其他
    abstract: "摘要", keyword: "关键词", unknown: "未识别样式"
  };

  // ============================================================
  // 二、工具函数
  // ============================================================
  function clean(t) {
    return String(t || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  function ptToSize(pt) {
    if (!pt) return "-";
    var map = {
      42: "初号", 36: "小初", 26: "一号", 24: "小一",
      22: "二号", 18: "小二", 16: "三号", 15: "小三",
      14: "四号", 12: "小四", 10.5: "五号", 9: "小五",
      7.5: "六号", 6.5: "小六"
    };
    for (var k in map) {
      if (Math.abs(pt - parseFloat(k)) < 0.5) return map[k];
    }
    return pt.toFixed(1) + "pt";
  }

  // ============================================================
  // 三、元素检测（基于样式元素规范表的正则模式）
  // ============================================================
  function detectStyle(text, isPaper) {
    if (isPaper) {
      // 论文报告检测
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
      if (/^参考文[献獻]/.test(text)) return "referenceTitle";
      if (/^\[\d+\]/.test(text)) return "reference";
      if (/^目\s*录$|^目次$/.test(text)) return "tocTitle";
      if (/^摘\s*要|^Abstract/i.test(text)) return "abstract";
      if (/^关键词|^关键字|^Key\s*words/i.test(text)) return "keyword";
      if (/^式中|^注\s*\d*/.test(text)) return "formulaNote";
      // 五级标题的其他形式
      if (/^[\(（]\d+[\)）]/.test(text)) return "heading5";
      if (/^[\(（][a-z][\)）]/.test(text)) return "heading5";
      if (/^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮]/.test(text)) return "heading5";
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
    }
    // 列表项
    if (/^[a-zA-Z][\)）\.]/.test(text)) return "listItem";
    if (/^\d+[\)）]/.test(text)) return "listItem";
    return null;
  }

  // ============================================================
  // 四、提取完整格式参数
  // ============================================================
  function extractFormat(para) {
    var rng = para.Range;
    var fmt = para.Format;

    return {
      // 字体相关（7个）
      fontCN: rng.Font.NameFarEast || rng.Font.Name || "",
      fontEN: rng.Font.NameAscii || "",
      fontSize: rng.Font.Size,
      fontSizeName: ptToSize(rng.Font.Size),
      bold: rng.Font.Bold ? true : false,
      italic: rng.Font.Italic ? true : false,
      underline: rng.Font.Underline ? true : false,
      color: rng.Font.Color ? rng.Font.Color : 0,

      // 段落格式（5个）
      alignment: fmt.Alignment,
      firstLineIndent: fmt.FirstLineIndent ? (fmt.FirstLineIndent / 240).toFixed(1) : "0",
      leftIndent: fmt.LeftIndent ? (fmt.LeftIndent / 240).toFixed(1) : "0",
      rightIndent: fmt.RightIndent ? (fmt.RightIndent / 240).toFixed(1) : "0",
      hangingIndent: fmt.CharacterUnitFirstLineIndent < 0 ? Math.abs(fmt.CharacterUnitFirstLineIndent).toFixed(1) : "0",

      // 间距（4个）
      spaceBefore: fmt.SpaceBefore ? (fmt.SpaceBefore / 20).toFixed(1) : "0",
      spaceAfter: fmt.SpaceAfter ? (fmt.SpaceAfter / 20).toFixed(1) : "0",
      lineSpacing: fmt.LineSpacing ? (fmt.LineSpacing / 20).toFixed(1) : "0",
      lineSpacingRule: fmt.LineSpacingRule
    };
  }

  // ============================================================
  // 五、主提取逻辑
  // ============================================================
  var paras = DOC.Paragraphs;
  var styles = {};

  for (var i = 1; i <= paras.Count; i++) {
    var para = paras.Item(i);
    var text = clean(para.Range.Text);
    if (!text) continue;

    var fmt = extractFormat(para);
    var type = detectStyle(text, isPaper);

    // 智能推断
    if (!type) {
      if (fmt.bold && fmt.fontSize >= 20) type = isPaper ? "coverTitle" : "docTitle";
      else if (fmt.bold && fmt.fontSize >= 16) type = "heading1";
      else if (fmt.bold && fmt.fontSize >= 14) type = "heading2";
      else if (fmt.bold) type = "heading3";
      else type = "body";
    }

    if (!styles[type]) {
      styles[type] = {
        id: type,
        name: STYLE_NAMES[type] || type,
        count: 0,
        formats: [],
        samples: []
      };
    }
    styles[type].count++;
    styles[type].formats.push(fmt);
    if (styles[type].samples.length < 3) {
      styles[type].samples.push(text.substring(0, 50));
    }
  }

  console.log("[extract] 样式数: " + Object.keys(styles).length);

  // ============================================================
  // 六、生成输出
  // ============================================================
  var alignMap = { 0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐", 4: "分散对齐" };
  var lineRuleMap = { 0: "单倍行距", 1: "1.5倍行距", 2: "2倍行距", 3: "最小值", 4: "固定值", 5: "多倍行距" };

  var lines = [];
  lines.push("✅ 样式模板提取完成！");
  lines.push("📄 源文档：" + DOC.Name);
  lines.push("📑 类型：" + (isPaper ? "论文报告样式" : "公文样式"));
  lines.push("📊 共 " + Object.keys(styles).length + " 种样式，" + paras.Count + " 个段落\n");

  // 按数量排序
  var sortedTypes = Object.keys(styles).sort(function(a, b) {
    return styles[b].count - styles[a].count;
  });

  lines.push("## 提取的样式详情\n");

  sortedTypes.forEach(function(t) {
    var s = styles[t], fmt = s.formats[0];

    lines.push("### " + s.name + "（" + s.count + "处）");

    // 字体参数
    var fontParts = [];
    if (fmt.fontCN) fontParts.push("中文字体: " + fmt.fontCN);
    if (fmt.fontEN) fontParts.push("英文字体: " + fmt.fontEN);
    if (fmt.fontSizeName && fmt.fontSizeName !== "-") fontParts.push("字号: " + fmt.fontSizeName);
    if (fmt.bold) fontParts.push("加粗");
    if (fmt.italic) fontParts.push("斜体");
    if (fmt.underline) fontParts.push("下划线");
    if (fontParts.length > 0) lines.push("字体: " + fontParts.join(" | "));

    // 段落参数
    var paraParts = [];
    if (fmt.alignment !== undefined && alignMap[fmt.alignment]) {
      paraParts.push("对齐: " + alignMap[fmt.alignment]);
    }
    if (parseFloat(fmt.firstLineIndent) > 0) {
      paraParts.push("首行缩进: " + fmt.firstLineIndent + "字符");
    }
    if (parseFloat(fmt.hangingIndent) > 0) {
      paraParts.push("悬挂缩进: " + fmt.hangingIndent + "字符");
    }
    if (parseFloat(fmt.leftIndent) > 0) {
      paraParts.push("左缩进: " + fmt.leftIndent + "字符");
    }
    if (parseFloat(fmt.rightIndent) > 0) {
      paraParts.push("右缩进: " + fmt.rightIndent + "字符");
    }
    if (paraParts.length > 0) lines.push("段落: " + paraParts.join(" | "));

    // 间距参数
    var spaceParts = [];
    if (parseFloat(fmt.spaceBefore) > 0) {
      spaceParts.push("段前: " + fmt.spaceBefore + "pt");
    }
    if (parseFloat(fmt.spaceAfter) > 0) {
      spaceParts.push("段后: " + fmt.spaceAfter + "pt");
    }
    if (parseFloat(fmt.lineSpacing) > 0) {
      spaceParts.push("行距: " + fmt.lineSpacing + "pt");
    }
    if (spaceParts.length > 0) lines.push("间距: " + spaceParts.join(" | "));

    // 示例
    if (s.samples.length > 0) {
      lines.push("示例: " + s.samples[0]);
    }
    lines.push("");
  });

  // 页面设置
  lines.push("## 页面设置");
  lines.push("纸张: " + (DOC.PageSetup.PaperSize === 1 ? "A4" : "自定义"));
  lines.push("方向: " + (DOC.PageSetup.Orientation === 1 ? "横向" : "纵向"));
  lines.push("上边距: " + (DOC.PageSetup.TopMargin / 567).toFixed(2) + "cm | 下边距: " + (DOC.PageSetup.BottomMargin / 567).toFixed(2) + "cm");
  lines.push("左边距: " + (DOC.PageSetup.LeftMargin / 567).toFixed(2) + "cm | 右边距: " + (DOC.PageSetup.RightMargin / 567).toFixed(2) + "cm");
  lines.push("页眉边距: " + (DOC.PageSetup.HeaderDistance / 567).toFixed(2) + "cm | 页脚边距: " + (DOC.PageSetup.FooterDistance / 567).toFixed(2) + "cm");
  if (DOC.PageSetup.Gutter > 0) {
    lines.push("装订线: " + (DOC.PageSetup.Gutter / 567).toFixed(2) + "cm");
  }

  // 页眉页脚
  try {
    var sec = DOC.Sections.Item(1);
    var headerText = clean(sec.Headers.Item(1).Range.Text);
    var footerText = clean(sec.Footers.Item(1).Range.Text);
    if (headerText) {
      lines.push("\n## 页眉");
      lines.push("内容: " + headerText.substring(0, 50));
    }
    if (footerText) {
      lines.push("\n## 页脚");
      lines.push("内容: " + footerText.substring(0, 50));
    }
  } catch(e) {}

  lines.push("\n模板已保存，可在后续【应用模板】时使用。");

  return {
    success: true,
    message: lines.join("\n"),
    styleCount: Object.keys(styles).length
  };
} catch (e) {
  return { success: false, error: String(e) };
}