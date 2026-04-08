/**
 * extract-template.js - 返回JSON字符串格式
 * 版本: 26.0408.1630
 */
(function() {
  "use strict";
  var VER = "26.0408.1630";
  console.log("[extract] 开始执行，版本: " + VER);

  var DOC = Application.ActiveDocument;
  if (!DOC) {
    return JSON.stringify({ success: false, error: "没有打开的文档" });
  }

  console.log("[extract] 文档: " + DOC.Name + ", 段落数: " + DOC.Paragraphs.Count);

  // 使用全局变量获取参数（与 generalcheck 一致）
  var docTypeParam = typeof docType !== 'undefined' ? docType : '';
  console.log("[extract] docType参数: " + docTypeParam);

  if (!docTypeParam) {
    return JSON.stringify({
      success: true,
      needUserInput: true,
      question: "请选择文档类型：",
      options: ["论文/技术报告", "公文"]
    });
  }

  var isPaper = docTypeParam === "论文/技术报告" || docTypeParam === "paper";
  console.log("[extract] 文档类型: " + docTypeParam + ", isPaper=" + isPaper);

  // 收集段落信息
  var paras = DOC.Paragraphs;
  var debugSamples = [];
  var styles = {};

  // 清理文本
  function clean(t) {
    return String(t || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  // 检测标题级别 - 论文
  function detectPaper(text) {
    if (/^\d+\.\d+\.\d+\.\d+\.\d+/.test(text)) return "heading5";
    if (/^\d+\.\d+\.\d+\.\d+/.test(text)) return "heading4";
    if (/^\d+\.\d+\.\d+/.test(text)) return "heading3";
    if (/^\d+\.\d+[^.\d]/.test(text)) return "heading2";
    if (/^\d+[^.\d]/.test(text)) return "heading1";
    if (/^第[一二三四五六七八九十\d]+章/.test(text)) return "chapterTitle";
    if (/^摘要|^Abstract/.test(text)) return "abstractTitle";
    if (/^关键词|^Keywords/.test(text)) return "keywords";
    if (/^图\s*\d+/.test(text)) return "figureCaption";
    if (/^表\s*\d+/.test(text)) return "tableCaption";
    if (/^附\s*录/.test(text)) return "appendixTitle";
    if (/^[A-Z]\.\d+/.test(text)) return "appendixSection";
    if (/^参考文献/.test(text)) return "referenceTitle";
    if (/^\[\d+\]/.test(text)) return "reference";
    return null;
  }

  // 样式名称映射
  var styleNames = {
    heading1: "一级标题", heading2: "二级标题", heading3: "三级标题",
    heading4: "四级标题", heading5: "五级标题", chapterTitle: "章标题",
    docTitle: "论文标题", abstractTitle: "摘要标题", keywords: "关键词",
    body: "正文", figureCaption: "图名", tableCaption: "表名",
    appendixTitle: "附录标题", appendixSection: "附录节题",
    referenceTitle: "参考文献标题", reference: "参考文献条目"
  };

  // 处理每个段落
  for (var i = 1; i <= paras.Count; i++) {
    var para = paras.Item(i);
    var rawText = para.Range.Text;
    var text = clean(rawText);

    if (!text) continue;

    // 获取格式
    var fmt = {
      fontCN: para.Range.Font.NameFarEast || para.Range.Font.Name,
      fontSize: para.Range.Font.Size,
      bold: para.Range.Font.Bold,
      align: para.Format.Alignment,
      firstLineIndent: para.Format.FirstLineIndent / 240,
      lineSpacing: para.Format.LineSpacing
    };

    // 检测类型
    var type = isPaper ? detectPaper(text) : null;

    // 格式后备检测
    if (!type && fmt.bold && fmt.fontSize >= 14) {
      type = "heading1";
    }
    if (!type) {
      type = "body";
    }

    // 记录前10个段落用于调试
    if (debugSamples.length < 10) {
      debugSamples.push({
        i: i,
        text: text.substring(0, 30),
        type: type,
        fs: fmt.fontSize
      });
    }

    // 累计
    if (!styles[type]) {
      styles[type] = { name: styleNames[type] || type, count: 0, formats: [], samples: [] };
    }
    styles[type].count++;
    styles[type].formats.push(fmt);
    if (styles[type].samples.length < 3) {
      styles[type].samples.push(text.substring(0, 40));
    }
  }

  console.log("[extract] 检测完成，样式数: " + Object.keys(styles).length);

  // 生成输出
  var alignMap = { 0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐" };

  var lines = [];
  lines.push("✅ 样式模板提取完成！");
  lines.push("");
  lines.push("📄 源文档：" + DOC.Name);
  lines.push("📑 类型：" + (isPaper ? "论文报告样式" : "公文样式"));
  lines.push("📊 共 " + Object.keys(styles).length + " 种样式，" + paras.Count + " 个段落");
  lines.push("");
  lines.push("## 提取的样式详情");
  lines.push("");

  // 输出样式（按优先级排序）
  var typeOrder = ["docTitle", "heading1", "heading2", "heading3", "heading4", "heading5",
                   "body", "figureCaption", "tableCaption", "appendixTitle", "appendixSection"];
  typeOrder.forEach(function(t) {
    if (!styles[t]) return;
    var s = styles[t];
    var fmt = s.formats[0];
    lines.push("### " + s.name + "（" + s.count + "处）");
    var p = [];
    if (fmt.fontCN) p.push("字体: " + fmt.fontCN);
    if (fmt.fontSize) p.push("字号: " + fmt.fontSize + "pt");
    if (fmt.bold) p.push("加粗");
    if (fmt.align !== undefined) p.push("对齐: " + alignMap[fmt.align]);
    if (fmt.firstLineIndent) p.push("首行缩进: " + fmt.firstLineIndent.toFixed(1) + "字符");
    if (fmt.lineSpacing) p.push("行距: " + fmt.lineSpacing.toFixed(1) + "pt");
    lines.push(p.join(" | "));
    lines.push("");
  });

  // 输出其他样式
  Object.keys(styles).forEach(function(t) {
    if (typeOrder.indexOf(t) >= 0) return;
    var s = styles[t];
    var fmt = s.formats[0];
    lines.push("### " + s.name + "（" + s.count + "处）");
    var p = [];
    if (fmt.fontCN) p.push("字体: " + fmt.fontCN);
    if (fmt.fontSize) p.push("字号: " + fmt.fontSize + "pt");
    if (fmt.bold) p.push("加粗");
    lines.push(p.join(" | "));
    lines.push("");
  });

  // 页面设置
  lines.push("## 页面设置");
  lines.push("- 纸张: A4");
  lines.push("- 上边距: " + (DOC.PageSetup.TopMargin / 567).toFixed(2) + "cm");
  lines.push("- 下边距: " + (DOC.PageSetup.BottomMargin / 567).toFixed(2) + "cm");
  lines.push("- 左边距: " + (DOC.PageSetup.LeftMargin / 567).toFixed(2) + "cm");
  lines.push("- 右边距: " + (DOC.PageSetup.RightMargin / 567).toFixed(2) + "cm");

  // 生成stylesTable
  var stylesTable = Object.keys(styles).map(function(t) {
    var s = styles[t];
    var fmt = s.formats[0];
    return {
      样式名称: s.name,
      出现次数: s.count + "处",
      字体: fmt.fontCN || "-",
      字号: fmt.fontSize ? fmt.fontSize + "pt" : "-",
      加粗: fmt.bold ? "是" : "否",
      对齐: alignMap[fmt.align] || "-"
    };
  });

  // 生成templateJson
  var template = {
    name: DOC.Name.replace(/\.(docx|doc)$/i, '') + '_模板',
    docType: isPaper ? "paper" : "official",
    styles: Object.keys(styles).map(function(t) {
      return {
        id: t,
        name: styles[t].name,
        count: styles[t].count,
        format: styles[t].formats[0]
      };
    }),
    pageSetup: {
      topMargin: DOC.PageSetup.TopMargin / 567,
      bottomMargin: DOC.PageSetup.BottomMargin / 567,
      leftMargin: DOC.PageSetup.LeftMargin / 567,
      rightMargin: DOC.PageSetup.RightMargin / 567
    }
  };

  return JSON.stringify({
    success: true,
    scriptVersion: VER,
    message: lines.join("\n"),
    stylesTable: stylesTable,
    templateJson: template,
    template: template,
    pageSetup: {
      paperSize: "A4",
      topMargin: (DOC.PageSetup.TopMargin / 567).toFixed(2) + "cm",
      bottomMargin: (DOC.PageSetup.BottomMargin / 567).toFixed(2) + "cm",
      leftMargin: (DOC.PageSetup.LeftMargin / 567).toFixed(2) + "cm",
      rightMargin: (DOC.PageSetup.RightMargin / 567).toFixed(2) + "cm"
    }
  }, null, 2);

})();