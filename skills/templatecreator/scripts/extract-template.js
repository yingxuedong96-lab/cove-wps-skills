/**
 * extract-template.js - 完整功能版本
 * 版本: 26.0408.1530 - 补充公式编号、公式说明、列表项检测，添加页眉页脚提取
 */

(function() {
  const SCRIPT_VERSION = "26.0408.1530";
  console.log("[extract-template] 脚本版本: " + SCRIPT_VERSION);

  const DOC = Application.ActiveDocument;
  if (!DOC) {
    return JSON.stringify({ success: false, error: "没有打开的文档" });
  }

  // 清理文本（移除WPS特殊字符）
  function cleanText(text) {
    return String(text || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  // 样式规范表 - 完整版，与样式元素规范表.md保持一致
  const STYLE_SPEC = {
    paper: {
      name: "论文报告样式",
      tags: [
        // 标题类（按层级从深到浅排列）
        { id: "heading5", name: "五级标题", pattern: /^\d+\.\d+\.\d+\.\d+\.\d+/ },
        { id: "heading4", name: "四级标题", pattern: /^\d+\.\d+\.\d+\.\d+/ },
        { id: "heading3", name: "三级标题", pattern: /^\d+\.\d+\.\d+/ },
        { id: "heading2", name: "二级标题", pattern: /^\d+\.\d+[^.\d]/ },
        { id: "heading1", name: "一级标题", pattern: /^\d+[^.\d\s]/ },
        { id: "chapterTitle", name: "章标题", pattern: /^第[一二三四五六七八九十\d]+章/ },

        // 前置部分
        { id: "docTitle", name: "论文标题", pattern: null },
        { id: "abstractTitle", name: "摘要标题", pattern: /^摘要|^Abstract/ },
        { id: "keywords", name: "关键词", pattern: /^关键词|^Keywords/ },
        { id: "tocTitle", name: "目录标题", pattern: /^目\s*录$|^目次$/ },

        // 正文类
        { id: "body", name: "正文", pattern: null },
        { id: "listItem", name: "列表项", pattern: /^\s*[a-z]\)|^\s*\d+\)|^[①②③④⑤⑥⑦⑧⑨⑩]|^\([a-z]\)|^\(\d+\)/ },

        // 图表公式
        { id: "figureCaption", name: "图名", pattern: /^图\s*\d+/ },
        { id: "tableCaption", name: "表名", pattern: /^表\s*\d+/ },
        { id: "formulaCaption", name: "公式编号", pattern: /^\([\d.\-–—]+\)$/ },
        { id: "formulaNote", name: "公式说明", pattern: /^式中[：:]/ },

        // 附录/参考文献
        { id: "appendixTitle", name: "附录标题", pattern: /^附\s*录/ },
        { id: "appendixSection", name: "附录节题", pattern: /^[A-Z]\.\d+/ },
        { id: "referenceTitle", name: "参考文献标题", pattern: /^参考文献/ },
        { id: "reference", name: "参考文献条目", pattern: /^\[\d+\]/ },

        // 注释
        { id: "note", name: "注释说明", pattern: /^注\s*\d*/ }
      ]
    },
    official: {
      name: "公文样式",
      tags: [
        // 版头
        { id: "issuer", name: "发文机关标志", pattern: null },
        { id: "docNumber", name: "发文字号", pattern: /[\d]{4}[\d号]|〔[\d]{4}〕[\d号]/ },

        // 标题
        { id: "docTitle", name: "公文标题", pattern: null },
        { id: "mainSender", name: "主送机关", pattern: null },
        { id: "heading1", name: "一级标题", pattern: /^[一二三四五六七八九十]+、/ },
        { id: "heading2", name: "二级标题", pattern: /^\([一二三四五六七八九十]+\)/ },
        { id: "heading3", name: "三级标题", pattern: /^\d+\.\s/ },

        // 正文
        { id: "body", name: "正文", pattern: null },
        { id: "listItem", name: "附件列表", pattern: /^\s*\d+\./ },

        // 结尾
        { id: "attachment", name: "附件说明", pattern: /^附件/ },
        { id: "signature", name: "发文机关署名", pattern: null },
        { id: "signDate", name: "成文日期", pattern: /\d{4}年\d{1,2}月\d{1,2}日/ },

        // 版记
        { id: "copySender", name: "抄送机关", pattern: /^抄送/ }
      ]
    }
  };

  const params = Application.Env?.ScriptParams || {};

  // 步骤1：选择文档类型
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

  // 使用 doc.Content.Text 获取全部文本
  const docText = DOC.Content && DOC.Content.Text ? String(DOC.Content.Text) : '';
  const textParas = docText.split('\r');

  console.log("[extract-template] 总段落数: " + textParas.length);

  // 获取段落的格式信息（需要通过 Paragraphs 对象）
  const paraCount = DOC.Paragraphs.Count;
  const results = {
    matched: {},
    unmatched: []
  };

  // 调试信息
  const debugInfo = [];

  // 遍历段落
  for (let i = 1; i <= paraCount; i++) {
    const para = DOC.Paragraphs.Item(i);
    const rawText = para.Range.Text;
    const text = cleanText(rawText);

    if (!text) continue;

    // 记录前20个段落的调试信息
    if (debugInfo.length < 20) {
      const codes = [];
      for (let j = 0; j < Math.min(8, text.length); j++) {
        codes.push(text.charCodeAt(j));
      }

      // 测试每个pattern
      const matchTests = [];
      for (const tag of spec.tags) {
        if (tag.pattern) {
          try {
            if (tag.pattern.test(text)) {
              matchTests.push(tag.name);
            }
          } catch (e) {}
        }
      }

      debugInfo.push({
        idx: i,
        text: text.substring(0, 25),
        codes: codes.join(","),
        matches: matchTests.length > 0 ? matchTests.join(",") : "无"
      });
    }

    // 提取格式
    const range = para.Range;
    const format = para.Format;
    const fmt = {
      fontCN: range.Font.NameFarEast || range.Font.Name,
      fontEN: range.Font.NameAscii,
      fontSize: range.Font.Size,
      bold: range.Font.Bold,
      italic: range.Font.Italic,
      color: range.Font.Color,
      alignment: format.Alignment,
      firstLineIndent: format.FirstLineIndent / 240,
      leftIndent: format.LeftIndent / 240,
      lineSpacing: format.LineSpacing,
      lineSpacingRule: format.LineSpacingRule,
      spaceBefore: format.SpaceBefore,
      spaceAfter: format.SpaceAfter
    };

    // 模式匹配
    let detection = null;
    for (const tag of spec.tags) {
      if (tag.pattern) {
        try {
          if (tag.pattern.test(text)) {
            detection = { tagId: tag.id, tagName: tag.name, confidence: "high" };
            break;
          }
        } catch (e) {}
      }
    }

    // 格式特征检测
    if (!detection) {
      if (fmt.bold && fmt.fontSize >= 14) {
        detection = { tagId: "heading1", tagName: "一级标题", confidence: "low" };
      } else if (fmt.fontSize >= 10 && fmt.fontSize <= 14) {
        detection = { tagId: "body", tagName: "正文", confidence: "medium" };
      }
    }

    // 记录结果
    if (detection) {
      if (!results.matched[detection.tagId]) {
        results.matched[detection.tagId] = { formats: [], samples: [], tagName: detection.tagName };
      }
      results.matched[detection.tagId].formats.push(fmt);
      if (results.matched[detection.tagId].samples.length < 3) {
        results.matched[detection.tagId].samples.push(text.substring(0, 50));
      }
    } else {
      results.unmatched.push({ text: text.substring(0, 30), fmt: fmt });
    }
  }

  // 合并格式（取众数）
  function mergeFormats(formatList) {
    if (!formatList || !formatList.length) return null;
    const groups = {};
    formatList.forEach(f => {
      const key = `${f.fontCN}_${f.fontSize}_${f.bold}_${f.alignment}`;
      if (!groups[key]) groups[key] = [];
      groups[key].push(f);
    });
    const sorted = Object.values(groups).sort((a, b) => b.length - a.length);
    return sorted[0] ? sorted[0][0] : formatList[0];
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

  // 添加匹配的样式
  for (const tag of spec.tags) {
    const data = results.matched[tag.id];
    if (data && data.formats.length > 0) {
      const mergedFmt = mergeFormats(data.formats);
      template.styles.push({
        id: tag.id,
        name: tag.name,
        count: data.formats.length,
        format: mergedFmt,
        samples: data.samples
      });
    }
  }

  // 提取页眉页脚
  const headerFooter = { header: null, footer: null };
  try {
    const sections = DOC.Sections;
    if (sections && sections.Count > 0) {
      const section = sections.Item(1);
      // 页眉
      const header = section.Headers.Item(1); // wdHeaderFooterPrimary = 1
      if (header && header.Range && header.Range.Text) {
        const headerText = cleanText(header.Range.Text);
        if (headerText) {
          const headerFont = header.Range.Font;
          headerFooter.header = {
            text: headerText.substring(0, 50),
            fontCN: headerFont.NameFarEast || headerFont.Name,
            fontSize: headerFont.Size,
            bold: headerFont.Bold,
            alignment: header.Range.ParagraphFormat.Alignment
          };
        }
      }
      // 页脚
      const footer = section.Footers.Item(1);
      if (footer && footer.Range && footer.Range.Text) {
        const footerText = cleanText(footer.Range.Text);
        if (footerText) {
          const footerFont = footer.Range.Font;
          headerFooter.footer = {
            text: footerText.substring(0, 50),
            fontCN: footerFont.NameFarEast || footerFont.Name,
            fontSize: footerFont.Size,
            bold: footerFont.Bold,
            alignment: footer.Range.ParagraphFormat.Alignment
          };
        }
      }
    }
  } catch (e) {
    console.log("[extract-template] 页眉页脚提取失败: " + e);
  }

  // 将页眉页脚添加到模板
  if (headerFooter.header) {
    template.styles.push({
      id: "header",
      name: "页眉",
      count: 1,
      format: headerFooter.header,
      samples: [headerFooter.header.text]
    });
  }
  if (headerFooter.footer) {
    template.styles.push({
      id: "footer",
      name: "页脚",
      count: 1,
      format: headerFooter.footer,
      samples: [headerFooter.footer.text]
    });
  }

  // 生成详细信息
  const alignMap = { 0: "左对齐", 1: "居中", 2: "右对齐", 3: "两端对齐" };
  const lineRuleMap = { 0: "单倍", 1: "最小值", 4: "固定值" };

  const lines = [];
  lines.push("✅ 样式模板提取完成！");
  lines.push("");
  lines.push("══════════════════════════════════════════════════");
  lines.push("【调试信息】前20个段落检测情况");
  lines.push("══════════════════════════════════════════════════");
  debugInfo.forEach(d => {
    lines.push(`[${d.idx}] "${d.text}"`);
    lines.push(`   字符码: ${d.codes} | 匹配: ${d.matches}`);
  });
  lines.push("══════════════════════════════════════════════════");
  lines.push("");

  lines.push(`📄 源文档：${DOC.Name}`);
  lines.push(`📑 类型：${spec.name}`);
  lines.push(`📊 共 ${template.styles.length} 种样式，${paraCount} 个段落`);
  lines.push("");
  lines.push("## 提取的样式详情");
  lines.push("");

  template.styles.forEach(s => {
    const fmt = s.format || {};
    lines.push(`### ${s.name}（${s.count}处）`);
    const params = [];
    if (fmt.fontCN) params.push(`字体: ${fmt.fontCN}`);
    if (fmt.fontSize) params.push(`字号: ${fmt.fontSize}pt`);
    if (fmt.bold) params.push("加粗");
    if (fmt.italic) params.push("斜体");
    if (fmt.alignment !== undefined) params.push(`对齐: ${alignMap[fmt.alignment] || '未知'}`);
    if (fmt.firstLineIndent) params.push(`首行缩进: ${fmt.firstLineIndent.toFixed(1)}字符`);
    if (fmt.lineSpacing) params.push(`行距: ${fmt.lineSpacing.toFixed(1)}pt`);
    lines.push(params.join(" | "));
    lines.push("");
  });

  lines.push("## 页面设置");
  lines.push(`- 纸张: A4`);
  lines.push(`- 上边距: ${template.pageSetup.topMargin.toFixed(2)}cm | 下边距: ${template.pageSetup.bottomMargin.toFixed(2)}cm`);
  lines.push(`- 左边距: ${template.pageSetup.leftMargin.toFixed(2)}cm | 右边距: ${template.pageSetup.rightMargin.toFixed(2)}`);
  lines.push("");

  // 页眉页脚信息
  if (headerFooter.header || headerFooter.footer) {
    lines.push("## 页眉页脚");
    if (headerFooter.header) {
      const h = headerFooter.header;
      lines.push(`### 页眉`);
      lines.push(`内容: "${h.text}"`);
      const hParams = [];
      if (h.fontCN) hParams.push(`字体: ${h.fontCN}`);
      if (h.fontSize) hParams.push(`字号: ${h.fontSize}pt`);
      if (h.bold) hParams.push("加粗");
      if (h.alignment !== undefined) hParams.push(`对齐: ${alignMap[h.alignment]}`);
      lines.push(hParams.join(" | "));
    }
    if (headerFooter.footer) {
      const f = headerFooter.footer;
      lines.push(`### 页脚`);
      lines.push(`内容: "${f.text}"`);
      const fParams = [];
      if (f.fontCN) fParams.push(`字体: ${f.fontCN}`);
      if (f.fontSize) fParams.push(`字号: ${f.fontSize}pt`);
      if (f.bold) fParams.push("加粗");
      if (f.alignment !== undefined) fParams.push(`对齐: ${alignMap[f.alignment]}`);
      lines.push(fParams.join(" | "));
    }
    lines.push("");
  }

  lines.push(`🔧 脚本版本: ${SCRIPT_VERSION}`);

  // 生成样式表格
  const stylesTable = template.styles.map(s => {
    const fmt = s.format || {};
    return {
      样式名称: s.name,
      出现次数: s.count + "处",
      字体: fmt.fontCN || "-",
      字号: fmt.fontSize ? fmt.fontSize + "pt" : "-",
      加粗: fmt.bold ? "是" : "否",
      对齐: alignMap[fmt.alignment] || "-",
      首行缩进: fmt.firstLineIndent ? fmt.firstLineIndent.toFixed(1) + "字符" : "-",
      行距: fmt.lineSpacing ? fmt.lineSpacing.toFixed(1) + "pt" : "-"
    };
  });

  return JSON.stringify({
    success: true,
    scriptVersion: SCRIPT_VERSION,
    message: lines.join("\n"),
    stylesTable: stylesTable,
    pageSetup: {
      paperSize: "A4",
      topMargin: `${template.pageSetup.topMargin.toFixed(2)}cm`,
      bottomMargin: `${template.pageSetup.bottomMargin.toFixed(2)}cm`,
      leftMargin: `${template.pageSetup.leftMargin.toFixed(2)}cm`,
      rightMargin: `${template.pageSetup.rightMargin.toFixed(2)}cm`
    },
    templateJson: template,
    template: template,
    debugInfo: debugInfo
  }, null, 2);

})();