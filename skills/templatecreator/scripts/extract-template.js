/**
 * extract-template.js - 使用全局变量获取参数（与generalcheck一致）
 * 版本: 26.0408.1625
 */
try {
  var VER = "26.0408.1625";
  console.log("[extract] 开始执行，版本: " + VER);

  var DOC = Application.ActiveDocument;
  if (!DOC) {
    return { success: false, error: "没有打开的文档" };
  }

  console.log("[extract] 文档: " + DOC.Name + ", 段落数: " + DOC.Paragraphs.Count);

  // 使用全局变量获取参数（与 generalcheck 一致）
  var docTypeParam = typeof docType !== 'undefined' ? docType : '';
  console.log("[extract] docType参数: " + docTypeParam);

  if (!docTypeParam) {
    return {
      success: true,
      needUserInput: true,
      question: "请选择文档类型：",
      options: ["论文/技术报告", "公文"]
    };
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
      align: para.Format.Alignment
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
        text: text.substring(0, 35),
        type: type,
        fs: fmt.fontSize,
        bold: fmt.bold
      });
    }

    // 累计
    if (!styles[type]) {
      styles[type] = { name: type, count: 0, formats: [], samples: [] };
    }
    styles[type].count++;
    styles[type].formats.push(fmt);
    if (styles[type].samples.length < 2) {
      styles[type].samples.push(text.substring(0, 40));
    }
  }

  console.log("[extract] 检测完成，样式数: " + Object.keys(styles).length);

  // 生成输出
  var alignMap = { 0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐" };

  var lines = [];
  lines.push("✅ 样式模板提取完成！v" + VER);
  lines.push("");
  lines.push("=== 调试信息 ===");
  lines.push("docType参数: " + docTypeParam);
  debugSamples.forEach(function(d) {
    lines.push("[" + d.i + "] \"" + d.text + "\" → " + d.type + " (" + d.fs + "pt)");
  });
  lines.push("样式列表: " + Object.keys(styles).join(", "));
  lines.push("");

  lines.push("📄 源文档：" + DOC.Name);
  lines.push("📊 共 " + Object.keys(styles).length + " 种样式");
  lines.push("");

  // 输出样式
  var typeOrder = ["heading1", "heading2", "heading3", "heading4", "heading5", "body", "figureCaption", "tableCaption", "appendixTitle"];
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
    lines.push(p.join(" | "));
    if (s.samples && s.samples.length > 0) {
      lines.push("示例: \"" + s.samples[0] + "\"");
    }
    lines.push("");
  });

  // 页面设置
  lines.push("## 页面设置");
  lines.push("- 上边距: " + (DOC.PageSetup.TopMargin / 567).toFixed(2) + "cm");
  lines.push("- 下边距: " + (DOC.PageSetup.BottomMargin / 567).toFixed(2) + "cm");

  return {
    success: true,
    scriptVersion: VER,
    message: lines.join("\n"),
    stylesTable: Object.keys(styles).map(function(k) {
      return { 样式: styles[k].name, 次数: styles[k].count + "处" };
    }),
    debugSamples: debugSamples
  };

} catch (e) {
  console.log("[extract] 错误: " + e);
  return { success: false, error: String(e) };
}