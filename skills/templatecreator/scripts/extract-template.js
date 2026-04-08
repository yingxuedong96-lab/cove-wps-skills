/**
 * extract-template.js - 稳定版本
 * 版本: 26.0408.1600 - 使用字符串匹配，避免正则状态问题
 */

(function() {
  const SCRIPT_VERSION = "26.0408.1600";
  console.log("[extract-template] 脚本版本: " + SCRIPT_VERSION);

  const DOC = Application.ActiveDocument;
  if (!DOC) {
    return JSON.stringify({ success: false, error: "没有打开的文档" });
  }

  // 清理文本
  function cleanText(text) {
    return String(text || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  // 检测函数 - 使用新创建的正则，避免状态问题
  function detectTag(text, docType) {
    // 论文报告样式
    if (docType === "paper") {
      // 按从长到短检测
      if (/^\d+\.\d+\.\d+\.\d+\.\d+/.test(text)) return { id: "heading5", name: "五级标题" };
      if (/^\d+\.\d+\.\d+\.\d+/.test(text)) return { id: "heading4", name: "四级标题" };
      if (/^\d+\.\d+\.\d+/.test(text)) return { id: "heading3", name: "三级标题" };
      if (/^\d+\.\d+[^.\d]/.test(text)) return { id: "heading2", name: "二级标题" };
      if (/^\d+[^.\d]/.test(text)) return { id: "heading1", name: "一级标题" };
      if (/^第[一二三四五六七八九十\d]+章/.test(text)) return { id: "chapterTitle", name: "章标题" };
      if (/^摘要|^Abstract/.test(text)) return { id: "abstractTitle", name: "摘要标题" };
      if (/^关键词|^Keywords/.test(text)) return { id: "keywords", name: "关键词" };
      if (/^目\s*录$|^目次$/.test(text)) return { id: "tocTitle", name: "目录标题" };
      if (/^\s*[a-z]\)|^\s*\d+\)|^[①②③④⑤⑥⑦⑧⑨⑩]/.test(text)) return { id: "listItem", name: "列表项" };
      if (/^图\s*\d+/.test(text)) return { id: "figureCaption", name: "图名" };
      if (/^表\s*\d+/.test(text)) return { id: "tableCaption", name: "表名" };
      if (/^\([\d.\-–—]+\)$/.test(text)) return { id: "formulaCaption", name: "公式编号" };
      if (/^式中[：:]/.test(text)) return { id: "formulaNote", name: "公式说明" };
      if (/^附\s*录/.test(text)) return { id: "appendixTitle", name: "附录标题" };
      if (/^[A-Z]\.\d+/.test(text)) return { id: "appendixSection", name: "附录节题" };
      if (/^参考文献/.test(text)) return { id: "referenceTitle", name: "参考文献标题" };
      if (/^\[\d+\]/.test(text)) return { id: "reference", name: "参考文献条目" };
      if (/^注\s*\d*/.test(text)) return { id: "note", name: "注释说明" };
    }
    // 公文样式
    else if (docType === "official") {
      if (/^[一二三四五六七八九十]+、/.test(text)) return { id: "heading1", name: "一级标题" };
      if (/^\([一二三四五六七八九十]+\)/.test(text)) return { id: "heading2", name: "二级标题" };
      if (/^\d+\.\s/.test(text)) return { id: "heading3", name: "三级标题" };
      if (/^附件/.test(text)) return { id: "attachment", name: "附件说明" };
      if (/\d{4}年\d{1,2}月\d{1,2}日/.test(text)) return { id: "signDate", name: "成文日期" };
      if (/^抄送/.test(text)) return { id: "copySender", name: "抄送机关" };
    }
    return null;
  }

  const params = Application.Env?.ScriptParams || {};

  // 选择文档类型
  if (!params.docType) {
    return JSON.stringify({
      success: true,
      needUserInput: true,
      stage: "selectDocType",
      question: "请选择文档类型：",
      options: ["论文/技术报告", "公文"]
    }, null, 2);
  }

  const docType = params.docType === "论文/技术报告" ? "paper" : "official";
  const docTypeName = docType === "paper" ? "论文报告样式" : "公文样式";

  // 获取段落
  const paraCount = DOC.Paragraphs.Count;
  const results = {};
  const debugList = [];

  // 遍历段落
  for (let i = 1; i <= paraCount; i++) {
    const para = DOC.Paragraphs.Item(i);
    const text = cleanText(para.Range.Text);
    if (!text) continue;

    // 获取格式
    const range = para.Range;
    const fmt = para.Format;
    const formatInfo = {
      fontCN: range.Font.NameFarEast || range.Font.Name,
      fontSize: range.Font.Size,
      bold: range.Font.Bold,
      alignment: fmt.Alignment,
      firstLineIndent: fmt.FirstLineIndent / 240,
      lineSpacing: fmt.LineSpacing
    };

    // 检测标签
    let tag = detectTag(text, docType);

    // 格式特征检测（作为后备）
    if (!tag) {
      if (formatInfo.bold && formatInfo.fontSize >= 14) {
        tag = { id: "heading1", name: "一级标题" };
      } else {
        tag = { id: "body", name: "正文" };
      }
    }

    // 记录调试信息（前15个段落）
    if (debugList.length < 15) {
      debugList.push({
        idx: i,
        text: text.substring(0, 30),
        tagId: tag.id,
        fontSize: formatInfo.fontSize,
        bold: formatInfo.bold
      });
    }

    // 累计结果
    if (!results[tag.id]) {
      results[tag.id] = { name: tag.name, count: 0, formats: [], samples: [] };
    }
    results[tag.id].count++;
    results[tag.id].formats.push(formatInfo);
    if (results[tag.id].samples.length < 3) {
      results[tag.id].samples.push(text.substring(0, 50));
    }
  }

  // 生成输出
  const alignMap = { 0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐" };
  const lines = [];

  lines.push("✅ 样式模板提取完成！版本: " + SCRIPT_VERSION);
  lines.push("");
  lines.push("════════════════════════════════════════");
  lines.push("【调试信息】前15个段落");
  lines.push("════════════════════════════════════════");
  debugList.forEach(d => {
    lines.push(`[${d.idx}] "${d.text}" => ${d.tagId} (${d.fontSize}pt, bold=${d.bold})`);
  });
  lines.push("════════════════════════════════════════");
  lines.push("");

  lines.push(`📄 源文档：${DOC.Name}`);
  lines.push(`📑 类型：${docTypeName}`);
  lines.push(`📊 共 ${Object.keys(results).length} 种样式，${paraCount} 个段落`);
  lines.push("");
  lines.push("## 提取的样式详情");
  lines.push("");

  // 按类型排序输出
  const tagOrder = ["docTitle", "heading1", "heading2", "heading3", "heading4", "heading5",
                    "body", "figureCaption", "tableCaption", "appendixTitle", "appendixSection"];
  const sortedTags = Object.keys(results).sort((a, b) => {
    const ia = tagOrder.indexOf(a);
    const ib = tagOrder.indexOf(b);
    return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
  });

  sortedTags.forEach(id => {
    const data = results[id];
    if (!data.formats.length) return;
    // 取众数格式
    const fmt = data.formats[0];
    lines.push(`### ${data.name}（${data.count}处）`);
    const params = [];
    if (fmt.fontCN) params.push(`字体: ${fmt.fontCN}`);
    if (fmt.fontSize) params.push(`字号: ${fmt.fontSize}pt`);
    if (fmt.bold) params.push("加粗");
    if (fmt.alignment !== undefined) params.push(`对齐: ${alignMap[fmt.alignment]}`);
    if (fmt.firstLineIndent) params.push(`首行缩进: ${fmt.firstLineIndent.toFixed(1)}字符`);
    if (fmt.lineSpacing) params.push(`行距: ${fmt.lineSpacing.toFixed(1)}pt`);
    lines.push(params.join(" | "));
    if (data.samples.length > 0) {
      lines.push(`示例: "${data.samples[0]}"`);
    }
    lines.push("");
  });

  // 页面设置
  lines.push("## 页面设置");
  lines.push(`- 纸张: A4`);
  lines.push(`- 上边距: ${(DOC.PageSetup.TopMargin / 567).toFixed(2)}cm`);
  lines.push(`- 下边距: ${(DOC.PageSetup.BottomMargin / 567).toFixed(2)}cm`);
  lines.push(`- 左边距: ${(DOC.PageSetup.LeftMargin / 567).toFixed(2)}cm`);
  lines.push(`- 右边距: ${(DOC.PageSetup.RightMargin / 567).toFixed(2)}cm`);

  // 构建样式表格
  const stylesTable = sortedTags.map(id => {
    const data = results[id];
    const fmt = data.formats[0];
    return {
      样式名称: data.name,
      出现次数: data.count + "处",
      字体: fmt.fontCN || "-",
      字号: fmt.fontSize ? fmt.fontSize + "pt" : "-",
      加粗: fmt.bold ? "是" : "否",
      对齐: alignMap[fmt.alignment] || "-"
    };
  });

  // 构建模板JSON
  const template = {
    name: DOC.Name.replace(/\.(docx|doc)$/i, '') + '_模板',
    docType: docType,
    styles: sortedTags.map(id => ({
      id: id,
      name: results[id].name,
      count: results[id].count,
      format: results[id].formats[0]
    })),
    pageSetup: {
      topMargin: DOC.PageSetup.TopMargin / 567,
      bottomMargin: DOC.PageSetup.BottomMargin / 567,
      leftMargin: DOC.PageSetup.LeftMargin / 567,
      rightMargin: DOC.PageSetup.RightMargin / 567
    }
  };

  return JSON.stringify({
    success: true,
    scriptVersion: SCRIPT_VERSION,
    message: lines.join("\n"),
    stylesTable: stylesTable,
    templateJson: template,
    template: template,
    pageSetup: {
      paperSize: "A4",
      topMargin: `${(DOC.PageSetup.TopMargin / 567).toFixed(2)}cm`,
      bottomMargin: `${(DOC.PageSetup.BottomMargin / 567).toFixed(2)}cm`,
      leftMargin: `${(DOC.PageSetup.LeftMargin / 567).toFixed(2)}cm`,
      rightMargin: `${(DOC.PageSetup.RightMargin / 567).toFixed(2)}cm`
    },
    debugList: debugList
  }, null, 2);

})();