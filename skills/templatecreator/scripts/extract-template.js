/**
 * extract-template.js - 调试版本
 * 版本: 26.0408.1430
 */

(function() {
  const SCRIPT_VERSION = "26.0408.1430";
  console.log("[extract-template] 脚本版本: " + SCRIPT_VERSION);

  const DOC = Application.ActiveDocument;
  if (!DOC) {
    return JSON.stringify({ success: false, error: "没有打开的文档" });
  }

  const STYLE_SPEC = {
    paper: {
      name: "论文报告样式",
      tags: [
        { id: "heading5", name: "五级标题", detectPattern: "^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+" },
        { id: "heading4", name: "四级标题", detectPattern: "^\\d+\\.\\d+\\.\\d+\\.\\d+" },
        { id: "heading3", name: "三级标题", detectPattern: "^\\d+\\.\\d+\\.\\d+" },
        { id: "heading2", name: "二级标题", detectPattern: "^\\d+\\.\\d+[^\\.\\d]" },
        { id: "heading1", name: "一级标题", detectPattern: "^\\d+[^\\.\\d]" },
        { id: "chapterTitle", name: "章标题", detectPattern: "^第[一二三四五六七八九十\\d]+章" },
        { id: "docTitle", name: "论文标题", detectPattern: null },
        { id: "abstractTitle", name: "摘要标题", detectPattern: "^摘要|^Abstract" },
        { id: "keywords", name: "关键词", detectPattern: "^关键词|^Keywords" },
        { id: "tocTitle", name: "目录标题", detectPattern: "^目\\s*录$|^目次$" },
        { id: "body", name: "正文", detectPattern: "default" },
        { id: "figureCaption", name: "图名", detectPattern: "^图\\s*\\d+" },
        { id: "tableCaption", name: "表名", detectPattern: "^表\\s*\\d+" },
        { id: "appendixTitle", name: "附录标题", detectPattern: "^附\\s*录" },
        { id: "appendixSection", name: "附录节题", detectPattern: "^[A-Z]\\.\\d+" },
        { id: "referenceTitle", name: "参考文献标题", detectPattern: "^参考文献" },
        { id: "reference", name: "参考文献条目", detectPattern: "^\\[\\d+\\]" }
      ]
    },
    official: {
      name: "公文样式",
      tags: [
        { id: "heading1", name: "一级标题", detectPattern: "^[一二三四五六七八九十]+、" },
        { id: "heading2", name: "二级标题", detectPattern: "^\\([一二三四五六七八九十]+\\)" },
        { id: "heading3", name: "三级标题", detectPattern: "^\\d+\\.\\s" },
        { id: "body", name: "正文", detectPattern: "default" }
      ]
    }
  };

  function extractParaFormat(para) {
    const range = para.Range;
    const format = para.Format;
    return {
      fontCN: range.Font.NameFarEast || range.Font.Name,
      fontSize: range.Font.Size,
      bold: range.Font.Bold,
      alignment: format.Alignment,
      firstLineIndent: format.FirstLineIndent / 240,
      lineSpacing: format.LineSpacing
    };
  }

  const params = Application.Env?.ScriptParams || {};

  if (!params.docType) {
    return JSON.stringify({
      success: true,
      needUserInput: true,
      stage: "selectDocType",
      question: "请选择文档类型：",
      options: ["论文/技术报告", "公文"]
    }, null, 2);
  }

  const docType = params.docType === "论文/技术报告" || params.docType === "paper" ? "paper" : "official";
  const spec = STYLE_SPEC[docType];
  const paragraphs = DOC.Paragraphs;

  // 收集前20个段落的调试信息
  const debugParas = [];
  for (let i = 1; i <= Math.min(20, paragraphs.Count); i++) {
    const para = paragraphs.Item(i);
    const rawText = para.Range.Text;
    const text = rawText.trim();
    if (!text) continue;

    // 检测匹配
    let matchedTag = null;
    for (const tag of spec.tags) {
      if (tag.detectPattern && tag.detectPattern !== "default") {
        try {
          const regex = new RegExp(tag.detectPattern);
          if (regex.test(text)) {
            matchedTag = tag.name;
            break;
          }
        } catch (e) {}
      }
    }

    // 获取字符编码
    const charCodes = [];
    for (let j = 0; j < Math.min(8, text.length); j++) {
      charCodes.push(text.charCodeAt(j));
    }

    const fmt = extractParaFormat(para);
    debugParas.push({
      idx: i,
      text: text.substring(0, 25),
      codes: charCodes.join(","),
      fontSize: fmt.fontSize,
      bold: fmt.bold,
      match: matchedTag || "未匹配"
    });
  }

  // 扫描所有段落
  const results = { matched: {} };
  for (let i = 1; i <= paragraphs.Count; i++) {
    const para = paragraphs.Item(i);
    const text = para.Range.Text.trim();
    if (!text) continue;

    const fmt = extractParaFormat(para);

    // 检测
    let detection = null;
    for (const tag of spec.tags) {
      if (tag.detectPattern && tag.detectPattern !== "default") {
        try {
          if (new RegExp(tag.detectPattern).test(text)) {
            detection = tag.id;
            break;
          }
        } catch (e) {}
      }
    }

    if (!detection) {
      // 格式检测
      if (fmt.bold && fmt.fontSize >= 14) {
        detection = "heading1";
      } else if (fmt.fontSize >= 10 && fmt.fontSize <= 14 && !fmt.bold) {
        detection = "body";
      }
    }

    if (detection) {
      if (!results.matched[detection]) {
        results.matched[detection] = { formats: [], samples: [] };
      }
      results.matched[detection].formats.push(fmt);
      if (results.matched[detection].samples.length < 3) {
        results.matched[detection].samples.push(text.substring(0, 40));
      }
    }
  }

  // 生成模板
  const template = {
    name: DOC.Name.replace(/\.(docx|doc)$/i, '') + '_模板',
    docType: docType,
    styles: []
  };

  for (const tag of spec.tags) {
    const data = results.matched[tag.id];
    if (data && data.formats.length > 0) {
      template.styles.push({
        id: tag.id,
        name: tag.name,
        count: data.formats.length,
        samples: data.samples
      });
    }
  }

  // 把调试信息作为特殊样式加入
  template.styles.unshift({
    id: "_DEBUG_INFO_",
    name: "===调试信息===",
    count: debugParas.length,
    debugParas: debugParas
  });

  // 生成详细message
  const lines = [];
  lines.push("✅ 样式模板提取完成！版本: " + SCRIPT_VERSION);
  lines.push("");
  lines.push("📄 源文档：" + DOC.Name);
  lines.push("📊 共 " + template.styles.length + " 种样式");
  lines.push("");

  // 调试信息
  lines.push("══════════════════════════════════════");
  lines.push("🔍 前20个段落调试信息");
  lines.push("══════════════════════════════════════");
  debugParas.forEach(p => {
    lines.push(`[${p.idx}] "${p.text}"`);
    lines.push(`    字符码: ${p.codes}`);
    lines.push(`    字号: ${p.fontSize}pt, 加粗: ${p.bold}, 匹配: ${p.match}`);
  });
  lines.push("══════════════════════════════════════");
  lines.push("");

  // 样式信息
  lines.push("## 提取的样式详情");
  lines.push("");
  template.styles.forEach(s => {
    if (s.id !== "_DEBUG_INFO_") {
      lines.push(`### ${s.name}（${s.count}处）`);
      if (s.samples && s.samples.length > 0) {
        lines.push(`示例: "${s.samples[0]}"`);
      }
      lines.push("");
    }
  });

  return JSON.stringify({
    success: true,
    scriptVersion: SCRIPT_VERSION,
    message: lines.join("\n"),
    template: template,
    debugParas: debugParas  // 单独输出，确保能看到
  }, null, 2);

})();