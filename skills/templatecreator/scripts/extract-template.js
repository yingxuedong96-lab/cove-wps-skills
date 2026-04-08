/**
 * extract-template.js - 使用样式规范表进行结构化提取
 * 版本: 26.0408.1415 - 调试版本：在message中显示段落文本
 *
 * 流程：
 * 1. 用户选择文档类型（公文/论文）
 * 2. 扫描文档，将格式匹配到规范表中的标签
 * 3. 不确定时汇总，ask user确认
 * 4. 生成完整模板JSON
 */

(function() {
  const SCRIPT_VERSION = "26.0408.1415";
  console.log("[extract-template] 脚本版本: " + SCRIPT_VERSION);

  const DOC = Application.ActiveDocument;
  if (!DOC) {
    return JSON.stringify({ success: false, error: "没有打开的文档" });
  }

  // 样式规范表 - 完整版，与样式元素规范表.md保持一致
  const STYLE_SPEC = {
    paper: {
      name: "论文报告样式",
      tags: [
        // 标题类（按层级从深到浅排列）
        { id: "heading5", name: "五级标题", detectPattern: "^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+", detectHint: "如'1.1.1.1.1'" },
        { id: "heading4", name: "四级标题", detectPattern: "^\\d+\\.\\d+\\.\\d+\\.\\d+", detectHint: "如'1.1.1.1'" },
        { id: "heading3", name: "三级标题", detectPattern: "^\\d+\\.\\d+\\.\\d+", detectHint: "如'1.1.1'" },
        { id: "heading2", name: "二级标题", detectPattern: "^\\d+\\.\\d+[^\\.\\d]", detectHint: "如'1.1 背景'" },
        { id: "heading1", name: "一级标题", detectPattern: "^\\d+[^\\.\\d]", detectHint: "如'1 范围'" },
        { id: "chapterTitle", name: "章标题", detectPattern: "^第[一二三四五六七八九十\\d]+章", detectHint: "如'第一章'" },

        // 前置部分
        { id: "docTitle", name: "论文标题", detectPattern: null, detectHint: "文档首段大字居中" },
        { id: "abstractTitle", name: "摘要标题", detectPattern: "^摘要|^Abstract", detectHint: "'摘要'" },
        { id: "keywords", name: "关键词", detectPattern: "^关键词|^Keywords", detectHint: "'关键词'" },
        { id: "tocTitle", name: "目录标题", detectPattern: "^目\\s*录$|^目次$", detectHint: "'目录'" },

        // 正文类
        { id: "body", name: "正文", detectPattern: "default", detectHint: "默认类型" },
        { id: "listItem", name: "列表项", detectPattern: "^\\s*[a-z]\\)|^\\s*[a-z][\\.\\s]|^\\s*\\d+\\)|^[①②③④⑤⑥⑦⑧⑨⑩]|^\\(\\d+\\)|^\\([a-z]\\)", detectHint: "如'a)'、'1)'、'①'、'(1)'" },

        // 图表公式
        { id: "figureCaption", name: "图名", detectPattern: "^图\\s*\\d+", detectHint: "'图'开头" },
        { id: "tableCaption", name: "表名", detectPattern: "^表\\s*\\d+", detectHint: "'表'开头" },

        // 附录/参考文献
        { id: "appendixTitle", name: "附录标题", detectPattern: "^附\\s*录", detectHint: "'附录'" },
        { id: "appendixSection", name: "附录节题", detectPattern: "^[A-Z]\\.\\d+", detectHint: "如'A.1'" },
        { id: "referenceTitle", name: "参考文献标题", detectPattern: "^参考文献", detectHint: "'参考文献'" },
        { id: "reference", name: "参考文献条目", detectPattern: "^\\[\\d+\\]", detectHint: "如'[1]'" },

        // 注释
        { id: "note", name: "注释说明", detectPattern: "^注\\s*\\d*", detectHint: "'注'开头" }
      ]
    },
    official: {
      name: "公文样式",
      tags: [
        { id: "issuer", name: "发文机关标志", detectPattern: null, detectHint: "如'XX市人民政府文件'" },
        { id: "docNumber", name: "发文字号", detectPattern: "[\\d]{4}[\\d号]|〔[\\d]{4}〕[\\d号]", detectHint: "如'国发〔2024〕1号'" },
        { id: "docTitle", name: "公文标题", detectPattern: null, detectHint: "主标题居中大字" },
        { id: "heading1", name: "一级标题", detectPattern: "^[一二三四五六七八九十]+、", detectHint: "如'一、'" },
        { id: "heading2", name: "二级标题", detectPattern: "^\\([一二三四五六七八九十]+\\)", detectHint: "如'(一)'" },
        { id: "heading3", name: "三级标题", detectPattern: "^\\d+\\.\\s", detectHint: "如'1.'" },
        { id: "body", name: "正文", detectPattern: "default", detectHint: "默认类型" },
        { id: "attachment", name: "附件说明", detectPattern: "^附件", detectHint: "'附件'开头" },
        { id: "signature", name: "发文机关署名", detectPattern: null, detectHint: "落款单位" },
        { id: "signDate", name: "成文日期", detectPattern: "\\d{4}年\\d{1,2}月\\d{1,2}日", detectHint: "如'2024年1月1日'" },
        { id: "copySender", name: "抄送机关", detectPattern: "^抄送", detectHint: "'抄送'开头" }
      ]
    }
  };

  // 属性提取函数
  function extractParaFormat(para) {
    const range = para.Range;
    const format = para.Format;
    return {
      fontCN: range.Font.NameFarEast || range.Font.Name,
      fontEN: range.Font.NameAscii,
      fontSize: range.Font.Size,
      bold: range.Font.Bold,
      italic: range.Font.Italic,
      color: range.Font.Color,
      alignment: format.Alignment,
      firstLineIndent: format.FirstLineIndent / 240,
      leftIndent: format.LeftIndent / 240,
      hangingIndent: -format.FirstLineIndent / 240,
      lineSpacing: format.LineSpacing,
      lineSpacingRule: format.LineSpacingRule,
      spaceBefore: format.SpaceBefore,
      spaceAfter: format.SpaceAfter
    };
  }

  // 格式特征描述
  function formatSignature(fmt) {
    const alignMap = { 0: "左齐", 1: "居中", 2: "右齐", 3: "两端" };
    const parts = [];
    parts.push(fmt.fontSize ? `${fmt.fontSize}pt` : "");
    parts.push(fmt.fontCN || "");
    parts.push(fmt.bold ? "加粗" : "");
    parts.push(alignMap[fmt.alignment] || "");
    parts.push(fmt.firstLineIndent ? `缩进${fmt.firstLineIndent}字符` : "");
    return parts.filter(p => p).join("") || "未知格式";
  }

  // 按模式检测标签类型
  function detectByPattern(text, spec, debugList) {
    const matches = [];
    for (const tag of spec.tags) {
      if (tag.detectPattern && tag.detectPattern !== "default") {
        try {
          const regex = new RegExp(tag.detectPattern);
          if (regex.test(text)) {
            matches.push({ tagId: tag.id, tagName: tag.name, pattern: tag.detectPattern });
          }
        } catch (e) {
          debugList.push(`[正则错误] ${tag.id}: ${tag.detectPattern} - ${e}`);
        }
      }
    }

    if (matches.length > 0) {
      // 按模式长度降序排序
      matches.sort((a, b) => (b.pattern?.length || 0) - (a.pattern?.length || 0));
      return { tagId: matches[0].tagId, tagName: matches[0].tagName, method: "pattern", confidence: "high" };
    }
    return null;
  }

  // 按格式特征检测
  function detectByFormat(fmt, isFirstPara) {
    if (isFirstPara && fmt.fontSize >= 18 && fmt.alignment === 1) {
      return { tagId: "docTitle", tagName: "主标题", method: "format", confidence: "medium" };
    }
    if (fmt.bold && fmt.fontSize >= 14) {
      if (fmt.fontSize >= 16) return { tagId: "heading1", confidence: "low" };
      if (fmt.fontSize >= 15) return { tagId: "heading2", confidence: "low" };
      if (fmt.fontSize >= 14) return { tagId: "heading3", confidence: "low" };
    }
    return null;
  }

  // ========== 主逻辑 ==========

  const params = Application.Env?.ScriptParams || {};

  if (!params.docType) {
    return JSON.stringify({
      success: true,
      needUserInput: true,
      stage: "selectDocType",
      question: "请选择文档类型，以便使用对应的样式规范表：",
      options: ["论文/技术报告", "公文"],
      note: "选择后将继续提取样式"
    }, null, 2);
  }

  const docType = params.docType === "论文/技术报告" || params.docType === "paper" ? "paper" : "official";
  const spec = STYLE_SPEC[docType];
  const userMapping = params.confirmMapping || {};

  const paragraphs = DOC.Paragraphs;
  const results = {
    matched: {},
    uncertain: [],
    unmatched: []
  };

  let isFirstPara = true;
  const debugList = [];  // 调试信息列表
  const paraSamples = []; // 前15个段落的原始文本

  for (let i = 1; i <= paragraphs.Count; i++) {
    const para = paragraphs.Item(i);
    const rawText = para.Range.Text;
    const text = rawText.trim();

    // 记录前15个段落的原始信息
    if (i <= 15) {
      const charCodes = [];
      for (let j = 0; j < Math.min(text.length, 10); j++) {
        charCodes.push(text.charCodeAt(j));
      }
      paraSamples.push({
        index: i,
        text: text.substring(0, 30),
        charCodes: charCodes.join(","),
        length: text.length
      });
    }

    if (!text) continue;

    const fmt = extractParaFormat(para);
    const sig = formatSignature(fmt);

    // 模式匹配
    let detection = detectByPattern(text, spec, debugList);

    // 格式特征检测
    if (!detection) {
      detection = detectByFormat(fmt, isFirstPara);
    }

    // 用户映射
    if (!detection && userMapping[sig]) {
      detection = { tagId: userMapping[sig], method: "userConfirmed", confidence: "high" };
    }

    isFirstPara = false;

    // 记录结果
    if (detection) {
      if (!results.matched[detection.tagId]) {
        results.matched[detection.tagId] = { formats: [], samples: [], tagId: detection.tagId };
      }
      results.matched[detection.tagId].formats.push(fmt);
      if (results.matched[detection.tagId].samples.length < 3) {
        results.matched[detection.tagId].samples.push(text.substring(0, 50));
      }
      results.matched[detection.tagId].confidence = detection.confidence;
      results.matched[detection.tagId].method = detection.method;
    } else if (fmt.fontSize >= 10 && fmt.fontSize <= 14 && !fmt.bold) {
      if (!results.matched["body"]) {
        results.matched["body"] = { formats: [], samples: [] };
      }
      results.matched["body"].formats.push(fmt);
      if (results.matched["body"].samples.length < 3) {
        results.matched["body"].samples.push(text.substring(0, 50));
      }
    } else {
      results.unmatched.push({
        index: i,
        text: text.substring(0, 30),
        format: sig,
        fontSize: fmt.fontSize,
        bold: fmt.bold
      });
    }
  }

  // 处理未匹配格式
  const specialUnmatched = results.unmatched.filter(u => u.fontSize >= 14 || u.bold);
  if (specialUnmatched.length > 0 && !params.userConfirmedUnmatched) {
    const formatGroups = {};
    specialUnmatched.forEach(u => {
      const sig = formatSignature(u);
      if (!formatGroups[sig]) formatGroups[sig] = { count: 0, samples: [] };
      formatGroups[sig].count++;
      if (formatGroups[sig].samples.length < 3) formatGroups[sig].samples.push(u.text);
    });

    return JSON.stringify({
      success: true,
      needUserInput: true,
      stage: "confirmUnmatched",
      docType: docType,
      matchedCount: Object.keys(results.matched).length,
      matchedSummary: Object.entries(results.matched).map(([k, v]) => `${k}: ${v.formats.length}处`).join(", "),
      unmatchedFormats: Object.entries(formatGroups).map(([sig, data]) => ({
        formatSignature: sig,
        count: data.count,
        samples: data.samples
      })),
      availableTags: spec.tags.map(t => ({ id: t.id, name: t.name, hint: t.detectHint })),
      question: `检测到以下格式未能自动识别，请帮助确认：\n\n${Object.entries(formatGroups).map(([sig, data]) =>
        `- ${sig}（${data.count}处）：示例 "${data.samples[0]}"`
      ).join('\n')}`,
      note: "回复格式如：'22pt黑体加粗居中 是 主标题'"
    }, null, 2);
  }

  // 合并格式
  function mergeFormats(formatList) {
    if (!formatList.length) return null;
    const groups = {};
    formatList.forEach(f => {
      const key = `${f.fontCN}_${f.fontSize}_${f.bold}_${f.alignment}`;
      if (!groups[key]) groups[key] = [];
      groups[key].push(f);
    });
    let maxGroup = Object.values(groups).sort((a, b) => b.length - a.length)[0];
    return maxGroup ? maxGroup[0] : formatList[0];
  }

  // 生成模板
  const template = {
    name: DOC.Name.replace(/\.(docx|doc)$/i, '') + '_模板',
    version: '1.0',
    docType: docType,
    docTypeName: spec.name,
    extractedFrom: DOC.Name,
    extractedAt: new Date().toISOString().split('T')[0],
    styles: [],
    pageSetup: {
      topMargin: DOC.PageSetup.TopMargin / 567,
      bottomMargin: DOC.PageSetup.BottomMargin / 567,
      leftMargin: DOC.PageSetup.LeftMargin / 567,
      rightMargin: DOC.PageSetup.RightMargin / 567
    }
  };

  for (const tag of spec.tags) {
    const data = results.matched[tag.id];
    if (!data || !data.formats.length) continue;
    const mergedFmt = mergeFormats(data.formats);
    template.styles.push({
      id: tag.id,
      name: tag.name,
      count: data.formats.length,
      detect: { pattern: tag.detectPattern || null, hint: tag.detectHint },
      format: mergedFmt
    });
  }

  // 生成详细信息
  const detailLines = [];

  // === 调试信息（放在最前面）===
  detailLines.push("## 🔍 调试信息");
  detailLines.push("");
  detailLines.push(`脚本版本: ${SCRIPT_VERSION}`);
  detailLines.push(`文档类型: ${spec.name}`);
  detailLines.push(`段落总数: ${paragraphs.Count}`);
  detailLines.push("");
  detailLines.push("### 前15个段落原始文本");
  detailLines.push("");
  paraSamples.forEach(p => {
    detailLines.push(`[${p.index}] "${p.text}" (字符码: ${p.charCodes})`);
  });
  detailLines.push("");

  if (debugList.length > 0) {
    detailLines.push("### 正则匹配日志");
    debugList.forEach(d => detailLines.push(d));
    detailLines.push("");
  }

  // === 样式提取结果 ===
  detailLines.push("---");
  detailLines.push("");
  detailLines.push("✅ 样式模板提取完成！");
  detailLines.push("");
  detailLines.push(`📄 源文档：${DOC.Name}`);
  detailLines.push(`📑 类型：${spec.name}`);
  detailLines.push(`📊 共 ${template.styles.length} 种样式，${paragraphs.Count} 个段落`);
  detailLines.push("");
  detailLines.push("## 提取的样式详情");
  detailLines.push("");

  template.styles.forEach(s => {
    const fmt = s.format || {};
    const alignMap = { 0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐" };

    detailLines.push(`### ${s.name}（${s.count}处）`);
    const params = [];
    if (fmt.fontCN) params.push(`字体: ${fmt.fontCN}`);
    if (fmt.fontSize) params.push(`字号: ${fmt.fontSize}pt`);
    if (fmt.bold) params.push("加粗");
    if (fmt.italic) params.push("斜体");
    if (fmt.alignment !== undefined) params.push(`对齐: ${alignMap[fmt.alignment] || '未知'}`);
    if (fmt.firstLineIndent) params.push(`首行缩进: ${fmt.firstLineIndent.toFixed(1)}字符`);
    if (fmt.lineSpacing) params.push(`行距: ${fmt.lineSpacing.toFixed(1)}pt`);
    if (fmt.spaceBefore) params.push(`段前: ${fmt.spaceBefore.toFixed(1)}pt`);
    if (fmt.spaceAfter) params.push(`段后: ${fmt.spaceAfter.toFixed(1)}pt`);

    detailLines.push(params.join(" | "));
    detailLines.push("");
  });

  detailLines.push("## 页面设置");
  detailLines.push(`- 纸张: A4`);
  detailLines.push(`- 上边距: ${template.pageSetup.topMargin.toFixed(2)}cm | 下边距: ${template.pageSetup.bottomMargin.toFixed(2)}cm`);
  detailLines.push(`- 左边距: ${template.pageSetup.leftMargin.toFixed(2)}cm | 右边距: ${template.pageSetup.rightMargin.toFixed(2)}cm`);
  detailLines.push("");
  detailLines.push(`🔧 脚本版本: ${SCRIPT_VERSION}`);

  const userMessage = detailLines.join("\n");

  // 生成样式表格
  const stylesTable = template.styles.map(s => {
    const fmt = s.format || {};
    return {
      样式名称: s.name,
      出现次数: s.count + "处",
      字体: fmt.fontCN || "-",
      字号: fmt.fontSize ? fmt.fontSize + "pt" : "-",
      加粗: fmt.bold ? "是" : "否",
      对齐: { 0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐" }[fmt.alignment] || "-",
      首行缩进: fmt.firstLineIndent ? fmt.firstLineIndent.toFixed(1) + "字符" : "-"
    };
  });

  return JSON.stringify({
    success: true,
    scriptVersion: SCRIPT_VERSION,
    message: userMessage,
    stylesTable: stylesTable,
    templateJson: template,
    styleDetails: template.styles.map(s => ({ name: s.name, count: s.count, format: s.format })),
    template: template
  }, null, 2);

})();