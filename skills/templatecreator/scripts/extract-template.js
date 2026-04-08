/**
 * extract-template.js - 使用样式规范表进行结构化提取
 * 版本: 26.0408.1400 - 修复正则表达式匹配问题（去掉末尾空格要求）
 *
 * 流程：
 * 1. 用户选择文档类型（公文/论文）
 * 2. 扫描文档，将格式匹配到规范表中的标签
 * 3. 不确定时汇总，ask user确认
 * 4. 生成完整模板JSON
 */

(function() {
  const SCRIPT_VERSION = "26.0408.1400";
  console.log("[extract-template] 脚本版本: " + SCRIPT_VERSION);

  const DOC = Application.ActiveDocument;
  if (!DOC) {
    return JSON.stringify({ success: false, error: "没有打开的文档" });
  }

  // 样式规范表 - 完整版，与样式元素规范表.md保持一致
  // 注意：检测顺序很重要，长的模式要放在前面（如五级标题在四级标题前面）
  // 修复：去掉末尾空格要求，改为匹配编号后紧跟非数字或字符串结束
  const STYLE_SPEC = {
    paper: {
      name: "论文报告样式",
      tags: [
        // 标题类（按层级从深到浅排列，确保长编号先匹配）
        // 使用 [^\d\.] 确保编号后不是数字或点号，避免误匹配
        { id: "heading5", name: "五级标题", detectPattern: "^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+[\\s：:]", detectHint: "如'1.1.1.1.1'" },
        { id: "heading4", name: "四级标题", detectPattern: "^\\d+\\.\\d+\\.\\d+\\.\\d+[\\s：:]", detectHint: "如'1.1.1.1'" },
        { id: "heading3", name: "三级标题", detectPattern: "^\\d+\\.\\d+\\.\\d+[\\s：:]", detectHint: "如'1.1.1'" },
        { id: "heading2", name: "二级标题", detectPattern: "^\\d+\\.\\d+[\\s：:]", detectHint: "如'1.1 背景'" },
        { id: "heading1", name: "一级标题", detectPattern: "^\\d+[\\s：:]+[^\\d\\.]", detectHint: "如'1 范围'（数字后直接跟汉字）" },
        { id: "chapterTitle", name: "章标题", detectPattern: "^第[一二三四五六七八九十\\d]+章", detectHint: "如'第一章 范围'" },

        // 标题/前置部分
        { id: "docTitle", name: "论文标题", detectPattern: null, detectHint: "文档首段，字号最大居中" },
        { id: "abstractTitle", name: "摘要标题", detectPattern: "^摘要|^Abstract", detectHint: "'摘要'" },
        { id: "keywords", name: "关键词", detectPattern: "^关键词|^Keywords", detectHint: "'关键词'" },
        { id: "tocTitle", name: "目录标题", detectPattern: "^目\\s*录$|^目次$", detectHint: "'目录'" },

        // 正文类
        { id: "body", name: "正文", detectPattern: "default", detectHint: "默认类型" },
        { id: "listItem", name: "列表项", detectPattern: "^\\s*[a-z]\\)|^\\s*[a-z][\\.\\s]|^\\s*\\d+\\)|^[①②③④⑤⑥⑦⑧⑨⑩]", detectHint: "如'a)'、'a.'、'a)'、'1)'、'①'" },

        // 图表公式
        { id: "figureCaption", name: "图名", detectPattern: "^图\\s*\\d+", detectHint: "'图'开头" },
        { id: "tableCaption", name: "表名", detectPattern: "^表\\s*\\d+", detectHint: "'表'开头" },

        // 附录/参考文献
        { id: "appendixTitle", name: "附录标题", detectPattern: "^附\\s*录\\s*[A-Z]?", detectHint: "'附录'或'附录A'或'附 录 A'" },
        { id: "appendixSection", name: "附录节题", detectPattern: "^[A-Z]\\.(\\d+\\.)*\\d+[\\s：:]", detectHint: "如'A.1'或'A.1.1'" },
        { id: "referenceTitle", name: "参考文献标题", detectPattern: "^参考文献", detectHint: "'参考文献'" },
        { id: "reference", name: "参考文献条目", detectPattern: "^\\[\\d+\\]", detectHint: "如'[1]'" },

        // 注释
        { id: "note", name: "注释说明", detectPattern: "^注\\s*\\d*", detectHint: "'注'开头" }
      ]
    },
    official: {
      name: "公文样式",
      tags: [
        // 版头
        { id: "issuer", name: "发文机关标志", detectPattern: null, detectHint: "如'XX市人民政府文件'" },
        { id: "docNumber", name: "发文字号", detectPattern: "[\\d]{4}[\\d号]|〔[\\d]{4}〕[\\d号]", detectHint: "如'国发〔2024〕1号'" },

        // 标题
        { id: "docTitle", name: "公文标题", detectPattern: null, detectHint: "主标题居中大字" },
        { id: "heading1", name: "一级标题", detectPattern: "^[一二三四五六七八九十]+、", detectHint: "如'一、'" },
        { id: "heading2", name: "二级标题", detectPattern: "^\\([一二三四五六七八九十]+\\)", detectHint: "如'(一)'" },
        { id: "heading3", name: "三级标题", detectPattern: "^\\d+\\.\\s", detectHint: "如'1.'" },

        // 正文
        { id: "body", name: "正文", detectPattern: "default", detectHint: "默认类型" },

        // 结尾
        { id: "attachment", name: "附件说明", detectPattern: "^附件", detectHint: "'附件'开头" },
        { id: "signature", name: "发文机关署名", detectPattern: null, detectHint: "落款单位" },
        { id: "signDate", name: "成文日期", detectPattern: "\\d{4}年\\d{1,2}月\\d{1,2}日", detectHint: "如'2024年1月1日'" },

        // 版记
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

  // 格式特征描述（用于用户确认）
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

  // 按模式检测标签类型（精确匹配，优先匹配最长的模式）
  function detectByPattern(text, spec) {
    // 先检查是否匹配任何模式，记录所有匹配
    const matches = [];
    for (const tag of spec.tags) {
      if (tag.detectPattern && tag.detectPattern !== "default") {
        try {
          const regex = new RegExp(tag.detectPattern);
          if (regex.test(text)) {
            matches.push({ tagId: tag.id, tagName: tag.name, pattern: tag.detectPattern });
          }
        } catch (e) {
          console.log("[detectByPattern] 正则错误: " + tag.detectPattern + " - " + e);
        }
      }
    }

    // 调试：打印前几个段落的匹配结果
    if (text.length < 30 && matches.length > 0) {
      console.log("[detectByPattern] 文本: " + text + " => 匹配: " + matches.map(m => m.tagId).join(","));
    }

    // 如果有多个匹配，选择最具体的那个（优先选择模式长的）
    if (matches.length > 0) {
      // 按模式长度降序排序，选择最长的模式
      matches.sort((a, b) => (b.pattern?.length || 0) - (a.pattern?.length || 0));
      return { tagId: matches[0].tagId, tagName: matches[0].tagName, method: "pattern", confidence: "high" };
    }
    return null;
  }

  // 按格式特征检测（用于模式匹配失败的段落）
  function detectByFormat(fmt, isFirstPara) {
    // 首段大字号居中可能是标题
    if (isFirstPara && fmt.fontSize >= 18 && fmt.alignment === 1) {
      return { tagId: "docTitle", tagName: "主标题", method: "format", confidence: "medium", reason: "首段大字居中" };
    }
    // 大字号加粗可能是标题
    if (fmt.bold && fmt.fontSize >= 14) {
      if (fmt.fontSize >= 16) return { tagId: "heading1", confidence: "low", reason: "16pt加粗" };
      if (fmt.fontSize >= 15) return { tagId: "heading2", confidence: "low", reason: "15pt加粗" };
      if (fmt.fontSize >= 14) return { tagId: "heading3", confidence: "low", reason: "14pt加粗" };
    }
    return null;
  }

  // ========== 主逻辑 ==========

  const params = Application.Env?.ScriptParams || {};

  // 如果用户还没选择文档类型，先询问
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

  // 如果有用户确认的映射，直接应用
  const userMapping = params.confirmMapping || {};

  // 扫描文档
  const paragraphs = DOC.Paragraphs;
  const results = {
    matched: {},      // tagId -> { formats: [], samples: [] }
    uncertain: [],    // 未确定的格式
    unmatched: []     // 无法识别的段落
  };

  let isFirstPara = true;

  // 调试：记录处理过程
  let debugLog = [];

  for (let i = 1; i <= paragraphs.Count; i++) {
    const para = paragraphs.Item(i);
    const text = para.Range.Text.trim();
    if (!text) continue;

    // 调试：打印前15个段落的原始文本
    if (i <= 15) {
      console.log("[段落" + i + "] " + text.substring(0, 40));
    }

    const fmt = extractParaFormat(para);
    const sig = formatSignature(fmt);

    // 1. 先尝试模式匹配
    let detection = detectByPattern(text, spec);

    // 调试：记录前10个段落
    if (i <= 10) {
      debugLog.push({
        index: i,
        text: text.substring(0, 30),
        detection: detection ? detection.tagId : "null",
        fontSize: fmt.fontSize,
        bold: fmt.bold
      });
    }

    // 2. 模式匹配失败，尝试格式特征检测
    if (!detection) {
      detection = detectByFormat(fmt, isFirstPara);
    }

    // 3. 检查用户已确认的映射
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
      // 可能是正文
      if (!results.matched["body"]) {
        results.matched["body"] = { formats: [], samples: [] };
      }
      results.matched["body"].formats.push(fmt);
      if (results.matched["body"].samples.length < 3) {
        results.matched["body"].samples.push(text.substring(0, 50));
      }
    } else {
      // 无法识别
      results.unmatched.push({
        index: i,
        text: text.substring(0, 30),
        format: sig,
        fontSize: fmt.fontSize,
        fontCN: fmt.fontCN,
        bold: fmt.bold,
        alignment: fmt.alignment
      });
    }
  }

  // 检查是否有低置信度的匹配需要确认
  const needsConfirm = [];
  for (const [tagId, data] of Object.entries(results.matched)) {
    if (data.confidence === "low" && !userMapping[formatSignature(data.formats[0])]) {
      const tagInfo = spec.tags.find(t => t.id === tagId);
      needsConfirm.push({
        tagId: tagId,
        tagName: tagInfo?.name || tagId,
        formatSignature: formatSignature(data.formats[0]),
        count: data.formats.length,
        samples: data.samples,
        reason: data.method === "format" ? "格式特征检测" : "未知"
      });
    }
  }

  // 如果有未匹配的特殊格式，也询问用户
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
      question: `检测到以下格式未能自动识别，请帮助确认它们属于哪种标签类型：\n\n${Object.entries(formatGroups).map(([sig, data]) =>
        `- ${sig}（${data.count}处）：示例 "${data.samples[0]}"`
      ).join('\n')}\n\n如果不确定，可以说"帮我找一下类似的格式"或直接指定标签类型。`,
      note: "回复格式如：'22pt黑体加粗居中 是 主标题' 或 '帮我找一下'"
    }, null, 2);
  }

  // 合并格式（取众数）
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

  // 添加已匹配的样式
  for (const tag of spec.tags) {
    const data = results.matched[tag.id];
    if (!data || !data.formats.length) continue;

    const mergedFmt = mergeFormats(data.formats);
    const styleEntry = {
      id: tag.id,
      name: tag.name,
      count: data.formats.length,
      detect: {
        pattern: tag.detectPattern || null,
        hint: tag.detectHint
      },
      format: mergedFmt
    };
    template.styles.push(styleEntry);
  }

  // 保存模板
  const templateFileName = `模板_${docType}_${new Date().toISOString().split('T')[0].replace(/-/g, '')}.json`;
  const skillPath = Application.Env?.SkillPath || '';
  const fullTemplatePath = skillPath ? `${skillPath}/templates/${templateFileName}` : `templates/${templateFileName}`;

  // 生成详细参数描述
  function formatStyleDetail(style) {
    const fmt = style.format;
    const alignMap = { 0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐" };
    const lineRuleMap = { 0: "单倍行距", 1: "最小值", 4: "固定值" };

    const parts = [];
    if (fmt.fontCN) parts.push(`字体: ${fmt.fontCN}`);
    if (fmt.fontSize) parts.push(`字号: ${fmt.fontSize}pt`);
    if (fmt.bold) parts.push("加粗");
    if (fmt.italic) parts.push("斜体");
    if (fmt.alignment !== undefined) parts.push(`对齐: ${alignMap[fmt.alignment] || '未知'}`);
    if (fmt.firstLineIndent) parts.push(`首行缩进: ${fmt.firstLineIndent.toFixed(1)}字符`);
    if (fmt.leftIndent) parts.push(`左缩进: ${fmt.leftIndent.toFixed(1)}字符`);
    if (fmt.lineSpacing) parts.push(`行距: ${fmt.lineSpacing.toFixed(1)}pt`);
    if (fmt.lineSpacingRule !== undefined) parts.push(`行距规则: ${lineRuleMap[fmt.lineSpacingRule] || '自动'}`);
    if (fmt.spaceBefore) parts.push(`段前: ${fmt.spaceBefore.toFixed(1)}pt`);
    if (fmt.spaceAfter) parts.push(`段后: ${fmt.spaceAfter.toFixed(1)}pt`);

    return parts.join(" | ");
  }

  // 生成每种样式的详细信息
  const styleDetails = template.styles.map(s => ({
    name: s.name,
    count: s.count,
    params: formatStyleDetail(s),
    format: s.format  // 保留原始格式数据
  }));

  // 页面设置信息
  const pageSetupInfo = {
    paperSize: "A4",
    topMargin: `${template.pageSetup.topMargin.toFixed(2)}cm`,
    bottomMargin: `${template.pageSetup.bottomMargin.toFixed(2)}cm`,
    leftMargin: `${template.pageSetup.leftMargin.toFixed(2)}cm`,
    rightMargin: `${template.pageSetup.rightMargin.toFixed(2)}cm`
  };

  // 尝试保存模板文件
  const templateJsonString = JSON.stringify(template, null, 2);

  // 生成详细的样式表格（供UI展示）
  const stylesTable = template.styles.map(s => {
    const fmt = s.format || {};
    return {
      样式名称: s.name,
      出现次数: s.count + "处",
      字体: fmt.fontCN || "-",
      字号: fmt.fontSize ? fmt.fontSize + "pt" : "-",
      加粗: fmt.bold ? "是" : "否",
      对齐: { 0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐" }[fmt.alignment] || "-",
      首行缩进: fmt.firstLineIndent ? fmt.firstLineIndent.toFixed(1) + "字符" : "-",
      行距: fmt.lineSpacing ? fmt.lineSpacing.toFixed(1) + "pt" : "-"
    };
  });

  // 生成用户可直接阅读的详细信息
  const detailLines = [];
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
    const lineRuleMap = { 0: "单倍", 1: "最小值", 4: "固定值" };

    detailLines.push(`### ${s.name}（${s.count}处）`);
    const params = [];
    if (fmt.fontCN) params.push(`字体: ${fmt.fontCN}`);
    if (fmt.fontSize) params.push(`字号: ${fmt.fontSize}pt`);
    if (fmt.bold) params.push("加粗");
    if (fmt.italic) params.push("斜体");
    if (fmt.alignment !== undefined) params.push(`对齐: ${alignMap[fmt.alignment] || '未知'}`);
    if (fmt.firstLineIndent) params.push(`首行缩进: ${fmt.firstLineIndent.toFixed(1)}字符`);
    if (fmt.leftIndent) params.push(`左缩进: ${fmt.leftIndent.toFixed(1)}字符`);
    if (fmt.lineSpacing) params.push(`行距: ${fmt.lineSpacing.toFixed(1)}pt`);
    if (fmt.lineSpacingRule !== undefined) params.push(`行距规则: ${lineRuleMap[fmt.lineSpacingRule] || '自动'}`);
    if (fmt.spaceBefore) params.push(`段前: ${fmt.spaceBefore.toFixed(1)}pt`);
    if (fmt.spaceAfter) params.push(`段后: ${fmt.spaceAfter.toFixed(1)}pt`);

    detailLines.push(params.join(" | "));
    detailLines.push("");
  });

  detailLines.push("## 页面设置");
  detailLines.push(`- 纸张: A4`);
  detailLines.push(`- 上边距: ${pageSetupInfo.topMargin} | 下边距: ${pageSetupInfo.bottomMargin}`);
  detailLines.push(`- 左边距: ${pageSetupInfo.leftMargin} | 右边距: ${pageSetupInfo.rightMargin}`);
  detailLines.push("");
  detailLines.push(`📁 模板文件: ${templateFileName}`);
  detailLines.push("");
  detailLines.push(`🔧 脚本版本: ${SCRIPT_VERSION}`);

  const userMessage = detailLines.join("\n");

  return JSON.stringify({
    success: true,
    scriptVersion: SCRIPT_VERSION,

    // ===== 调试信息 =====
    debugLog: debugLog,

    // ===== 核心展示信息（必须用message字段，UI只展示这个）=====
    message: userMessage,

    // ===== 详细样式表格 =====
    stylesTable: stylesTable,

    // ===== 页面设置 =====
    pageSetup: pageSetupInfo,

    // ===== 完整模板JSON =====
    templateJson: template,
    templateFileName: templateFileName,

    // ===== 原始详细数据 =====
    styleDetails: styleDetails,

    // ===== 兼容旧格式 =====
    template: template
  }, null, 2);

})();