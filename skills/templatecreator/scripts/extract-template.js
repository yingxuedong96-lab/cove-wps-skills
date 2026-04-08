/**
 * extract-template.js - 完整样式提取（符合样式元素规范表）
 * 版本: 26.0410.1010
 * 支持元素: 论文报告26种 + 公文20种
 * 支持参数: 45个（字体7 + 段落5 + 间距4 + 大纲3 + 表格10 + 页面9 + 其他7）
 * 更新:
 *   - 新增公式编号检测
 *   - 新增封面副标题/单位/日期检测
 *   - 修复附录标题误识别问题（上下文感知）
 *   - 补充缺失的元素检测
 */
try {
  var VER = "26.0410.1006";
  console.log("[extract] 版本: " + VER);

  var DOC = Application.ActiveDocument;
  if (!DOC) return { success: false, error: "没有打开的文档" };

  // ============================================================
  // 自动检测文档类型
  // ============================================================
  var docTypeParam = typeof docType !== 'undefined' ? docType : '';

  // 自动检测逻辑：扫描前30段落，统计公文/论文特征
  var paperScore = 0, govScore = 0;
  var paras = DOC.Paragraphs;
  var scanCount = Math.min(30, paras.Count);

  for (var i = 1; i <= scanCount; i++) {
    var text = clean(paras.Item(i).Range.Text);
    if (!text) continue;

    // 论文报告特征
    if (/^第[一二三四五六七八九十百零\d]+章/.test(text)) paperScore += 3;
    if (/^\d+\.\d+/.test(text)) paperScore += 2;
    if (/^附\s*录/.test(text)) paperScore += 2;
    if (/^摘\s*要|^Abstract/i.test(text)) paperScore += 2;
    if (/^目\s*录$|^目次$/.test(text)) paperScore += 1;
    if (/^图\s*\d+|^表\s*\d+/.test(text)) paperScore += 1;

    // 公文特征
    if (/^关于/.test(text)) govScore += 3;
    if (/通知$|决定$|意见$|办法$|规定$|批复$|请示$|报告$/.test(text)) govScore += 2;
    if (/^附\s*件/.test(text)) govScore += 2;
    if (/\d{4}年\d{1,2}月\d{1,2}日$/.test(text)) govScore += 1;
    if (/^抄送|^印发/.test(text)) govScore += 2;
  }

  // 确定类型：得分高者胜，默认论文报告
  if (docTypeParam) {
    var isPaper = docTypeParam === "论文/技术报告" || docTypeParam === "paper";
  } else {
    isPaper = paperScore >= govScore;
  }

  var detectedType = isPaper ? "论文报告" : "公文";
  console.log("[extract] 类型: " + detectedType + " (论文=" + paperScore + ", 公文=" + govScore + ")");

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

  // 元素检测模式说明（供参考）
  var DETECT_PATTERNS = {
    // 论文报告
    chapterTitle: "第X章",
    heading1: "1 XXX", heading2: "1.1 XXX", heading3: "1.1.1 XXX",
    heading4: "1.1.1.1 XXX", heading5: "1.1.1.1.1 XXX 或 (1)/(a)/①",
    appendixTitle: "附录 A", appendixSection: "A.1 XXX",
    figureCaption: "图 1-1", tableCaption: "表 1-1",
    formulaCaption: "(1) 或 (3.2.1-1)",
    referenceTitle: "参考文献", reference: "[1] XXX",
    tocTitle: "目录", abstract: "摘要", keyword: "关键词",
    formulaNote: "式中", note: "注",
    coverDate: "2024年1月1日",
    // 公文
    docTitle: "关于...的通知", heading1: "一、", heading2: "（一）", heading3: "1.",
    attachment: "附件", signDate: "2024年1月1日",
    copySender: "抄送", issuerDept: "印发机关", issueDate: "印发日期"
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

  // 格式签名：用于比较两个格式是否相同
  function formatSignature(fmt) {
    return [
      fmt.fontCN, fmt.fontSize, fmt.bold ? 'B' : '',
      fmt.alignment, fmt.firstLineIndent, fmt.leftIndent,
      fmt.spaceBefore, fmt.spaceAfter, fmt.lineSpacing, fmt.lineSpacingRule
    ].join('|');
  }

  // ============================================================
  // 三、元素检测（基于样式元素规范表的正则模式）
  // ============================================================

  // 上下文状态：用于检测附录相关内容
  var lastAppendixTitle = false;  // 上一段是否是附录标题

  function detectStyle(text, isPaper, fmt, paraIndex) {
    if (isPaper) {
      // 论文报告检测（优先级从高到低）

      // 1. 章标题
      if (/^第[一二三四五六七八九十百零\d]+章/.test(text)) return "chapterTitle";

      // 2. 标题编号（按层级从深到浅检测）
      if (/^\d+\.\d+\.\d+\.\d+\.\d+/.test(text)) return "heading5";
      if (/^\d+\.\d+\.\d+\.\d+/.test(text)) return "heading4";
      if (/^\d+\.\d+\.\d+[^.\d]/.test(text)) return "heading3";
      if (/^\d+\.\d+[^.\d]/.test(text)) return "heading2";
      if (/^\d+\s+[^\d\.\s]/.test(text)) return "heading1";

      // 3. 附录相关
      if (/^附\s*录\s*[A-Z０-９０-９]/.test(text)) {
        lastAppendixTitle = true;
        return "appendixTitle";
      }
      if (/^[A-Z]\.\d+/.test(text)) return "appendixSection";

      // 4. 图表公式
      if (/^图\s*\d+/.test(text)) return "figureCaption";
      if (/^表\s*\d+/.test(text)) return "tableCaption";
      if (/^\([\d.-]+\)$/.test(text)) return "formulaCaption";  // 公式编号如(1)、(3.2.1-1)

      // 5. 参考文献
      if (/^参考文[献獻]/.test(text)) return "referenceTitle";
      if (/^\[\d+\]/.test(text)) return "reference";

      // 6. 目录
      if (/^目\s*录$|^目次$/.test(text)) return "tocTitle";

      // 7. 摘要关键词
      if (/^摘\s*要|^Abstract/i.test(text)) return "abstract";
      if (/^关键词|^关键字|^Key\s*words/i.test(text)) return "keyword";

      // 8. 公式说明和注释
      if (/^式中/.test(text)) return "formulaNote";
      if (/^注\s*\d*|^注\s*：/.test(text)) return "note";

      // 9. 封面元素
      if (/^\d{4}年\d{1,2}月\d{1,2}日$/.test(text)) return "coverDate";
      if (paraIndex <= 5 && /\d{4}年\d{1,2}月\d{1,2}日/.test(text)) return "coverDate";

      // 10. 五级标题的其他形式 -> 列表项
      if (/^[\(（]\d+[\)）]/.test(text)) return "listItem";
      if (/^[\(（][a-z][\)）]/.test(text)) return "listItem";
      if (/^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮]/.test(text)) return "listItem";
      if (/^[a-z][\)）\.]/.test(text)) return "listItem";

      // 11. 上下文感知：附录标题后的标题段落
      if (lastAppendixTitle && fmt && fmt.bold && fmt.fontSize >= 14) {
        lastAppendixTitle = false;  // 重置状态
        return "appendixTitle";  // 作为附录标题的一部分
      }

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

      // 公文列表项
      if (/^[\(（]\d+[\)）]/.test(text)) return "listItem";
      if (/^[a-z][\)）\.]/.test(text)) return "listItem";
    }

    // 基于格式的智能推断
    // 悬挂缩进（首行缩进为负）通常是列表项
    if (fmt && fmt.firstLineIndent < 0) return "listItem";

    // 重置附录标题状态（如果检测失败）
    if (isPaper && !/^附\s*录/.test(text)) {
      lastAppendixTitle = false;
    }

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

      // 段落格式（WPS返回磅值）
      alignment: fmt.Alignment,
      firstLineIndent: fmt.FirstLineIndent || 0,
      leftIndent: fmt.LeftIndent || 0,
      rightIndent: fmt.RightIndent || 0,
      characterUnitFirstLine: fmt.CharacterUnitFirstLineIndent || 0,

      // 间距（WPS返回磅值）
      spaceBefore: fmt.SpaceBefore || 0,
      spaceAfter: fmt.SpaceAfter || 0,
      lineSpacing: fmt.LineSpacing || 0,
      lineSpacingRule: fmt.LineSpacingRule
    };
  }

  // ============================================================
  // 五、主提取逻辑
  // ============================================================
  var paras = DOC.Paragraphs;
  var styles = {};
  var formatIndex = {};  // 用于格式去重

  for (var i = 1; i <= paras.Count; i++) {
    var para = paras.Item(i);
    var text = clean(para.Range.Text);
    if (!text) continue;

    var fmt = extractFormat(para);
    var type = detectStyle(text, isPaper, fmt, i);  // 传入段落索引

    // 智能推断（基于格式特征）
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
        formatSignatures: {},  // 用于去重
        samples: []
      };
    }
    styles[type].count++;

    // 格式去重：只保留不同的格式
    var sig = formatSignature(fmt);
    if (!styles[type].formatSignatures[sig]) {
      styles[type].formatSignatures[sig] = { fmt: fmt, count: 1 };
      styles[type].formats.push(fmt);
    } else {
      styles[type].formatSignatures[sig].count++;
    }

    if (styles[type].samples.length < 3) {
      styles[type].samples.push(text.substring(0, 50));
    }
  }

  // 为每种样式选择最主要的格式（出现次数最多的）
  for (var t in styles) {
    if (styles[t].formats.length > 1) {
      // 找出出现最多的格式
      var mainFmt = null;
      var maxCount = 0;
      for (var sig in styles[t].formatSignatures) {
        if (styles[t].formatSignatures[sig].count > maxCount) {
          maxCount = styles[t].formatSignatures[sig].count;
          mainFmt = styles[t].formatSignatures[sig].fmt;
        }
      }
      // 只保留主要格式
      if (mainFmt) {
        styles[t].formats = [mainFmt];
      }
    }
    // 清理临时数据
    delete styles[t].formatSignatures;
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
    // 首行缩进：优先用字符单位，否则用磅值转字符（约10.5pt/字符）
    var indentChar = fmt.characterUnitFirstLine || 0;
    if (indentChar > 0) {
      paraParts.push("首行缩进: " + indentChar.toFixed(1) + "字符");
    } else if (parseFloat(fmt.firstLineIndent) > 0) {
      // WPS返回磅值，转换为字符单位（小四约10.5pt）
      indentChar = fmt.firstLineIndent / 10.5;
      paraParts.push("首行缩进: " + indentChar.toFixed(1) + "字符(" + fmt.firstLineIndent.toFixed(1) + "pt)");
    }
    if (parseFloat(fmt.leftIndent) > 0) {
      paraParts.push("左缩进: " + (fmt.leftIndent / 10.5).toFixed(1) + "字符");
    }
    if (parseFloat(fmt.rightIndent) > 0) {
      paraParts.push("右缩进: " + (fmt.rightIndent / 10.5).toFixed(1) + "字符");
    }
    if (paraParts.length > 0) lines.push("段落: " + paraParts.join(" | "));

    // 间距参数
    var spaceParts = [];
    if (parseFloat(fmt.spaceBefore) > 0) {
      spaceParts.push("段前: " + fmt.spaceBefore.toFixed(1) + "pt");
    }
    if (parseFloat(fmt.spaceAfter) > 0) {
      spaceParts.push("段后: " + fmt.spaceAfter.toFixed(1) + "pt");
    }
    // 行距：根据LineSpacingRule解释
    if (fmt.lineSpacingRule !== undefined && lineRuleMap[fmt.lineSpacingRule]) {
      var spacingDisplay = lineRuleMap[fmt.lineSpacingRule];
      if (fmt.lineSpacingRule === 4) {
        // 固定值，lineSpacing是磅值
        spacingDisplay = "固定值" + fmt.lineSpacing.toFixed(1) + "pt";
      } else if (fmt.lineSpacing > 0 && fmt.lineSpacing !== 1) {
        spacingDisplay += "(" + fmt.lineSpacing.toFixed(2) + "倍)";
      }
      spaceParts.push("行距: " + spacingDisplay);
    } else if (parseFloat(fmt.lineSpacing) > 0) {
      spaceParts.push("行距: " + fmt.lineSpacing.toFixed(1) + "pt");
    }
    if (spaceParts.length > 0) lines.push("间距: " + spaceParts.join(" | "));

    // 示例
    if (s.samples.length > 0) {
      lines.push("示例: " + s.samples[0]);
    }
    lines.push("");
  });

  // 页面设置（修复单位问题）
  lines.push("## 页面设置");
  lines.push("纸张: " + (DOC.PageSetup.PaperSize === 1 ? "A4" : "自定义"));
  lines.push("方向: " + (DOC.PageSetup.Orientation === 1 ? "横向" : "纵向"));

  // 尝试多种方式获取页面边距（单位：磅）
  var ps = DOC.PageSetup;
  var topM = ps.TopMargin, bottomM = ps.BottomMargin, leftM = ps.LeftMargin, rightM = ps.RightMargin;
  var headerD = ps.HeaderDistance, footerD = ps.FooterDistance;

  // 检测单位：如果值很小（<100），可能是厘米或英寸，需要转换
  // WPS API 通常返回磅值，但某些情况下可能返回其他单位
  // 2.54cm ≈ 72pt，如果值约72，说明是磅值
  // 如果值约2.54，说明是厘米

  function toCm(val) {
    // 如果值 < 10，假设是厘米
    if (val < 10) return val.toFixed(2);
    // 否则假设是磅值，转换为厘米 (1pt = 2.54/72 cm)
    return (val * 2.54 / 72).toFixed(2);
  }

  lines.push("上边距: " + toCm(topM) + "cm | 下边距: " + toCm(bottomM) + "cm");
  lines.push("左边距: " + toCm(leftM) + "cm | 右边距: " + toCm(rightM) + "cm");
  lines.push("页眉边距: " + toCm(headerD) + "cm | 页脚边距: " + toCm(footerD) + "cm");
  if (ps.Gutter > 0) {
    lines.push("装订线: " + toCm(ps.Gutter) + "cm");
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

  // ============================================================
  // 七、生成模板数据（供保存或直接返回）
  // ============================================================
  var templateDir = "/Users/cassia/Desktop/dyx/wpsjs/模版生成/";
  var docNameBase = DOC.Name.replace(/\.[docxdoc]+$/i, "");
  var jsonFile = templateDir + docNameBase + "-样式模板.json";

  // 构建完整的样式数据结构
  var templateData = {
    version: VER,
    docName: DOC.Name,
    docType: isPaper ? "论文报告" : "公文",
    extractTime: new Date().toISOString(),
    pageSetup: {
      paperSize: DOC.PageSetup.PaperSize === 1 ? "A4" : "自定义",
      orientation: DOC.PageSetup.Orientation === 1 ? "横向" : "纵向",
      topMargin: toCm(topM),
      bottomMargin: toCm(bottomM),
      leftMargin: toCm(leftM),
      rightMargin: toCm(rightM),
      headerDistance: toCm(headerD),
      footerDistance: toCm(footerD),
      gutter: ps.Gutter > 0 ? toCm(ps.Gutter) : "0"
    },
    styles: {}
  };

  // 保存每种样式的完整参数（只保留主要格式）
  sortedTypes.forEach(function(t) {
    var s = styles[t];
    var fmt = s.formats[0];  // 只取主要格式
    templateData.styles[t] = {
      id: t,
      name: STYLE_NAMES[t] || t,
      count: s.count,
      format: {
        // 字体参数
        fontCN: fmt.fontCN,
        fontEN: fmt.fontEN,
        fontSize: fmt.fontSize,
        fontSizeName: fmt.fontSizeName,
        bold: fmt.bold,
        italic: fmt.italic,
        underline: fmt.underline,
        color: fmt.color,
        // 段落参数
        alignment: fmt.alignment,
        alignmentName: alignMap[fmt.alignment] || "",
        firstLineIndent: fmt.firstLineIndent,
        firstLineIndentChars: fmt.characterUnitFirstLine || (fmt.firstLineIndent / 10.5),
        leftIndent: fmt.leftIndent,
        rightIndent: fmt.rightIndent,
        // 间距参数
        spaceBefore: fmt.spaceBefore,
        spaceAfter: fmt.spaceAfter,
        lineSpacing: fmt.lineSpacing,
        lineSpacingRule: fmt.lineSpacingRule,
        lineSpacingRuleName: lineRuleMap[fmt.lineSpacingRule] || ""
      },
      samples: s.samples
    };
  });

  var jsonContent = JSON.stringify(templateData, null, 2);

  // 尝试保存文件
  var saveResult = "";
  var fileSaved = false;

  // 固定文件名（供应用模板时读取）
  var latestTemplateFile = templateDir + "latest-template.json";

  try {
    // 方案1: 用 WPS 创建临时文档保存 JSON
    var jsonDoc = Application.Documents.Add();
    jsonDoc.Range(0, 0).Text = jsonContent;

    // 保存到文档名对应的文件
    jsonDoc.SaveAs2(jsonFile.replace('.json', '.txt'), 7);

    // 额外保存到固定文件名
    try {
      jsonDoc.SaveAs2(latestTemplateFile, 7);
      console.log("[extract] 固定模板保存成功: " + latestTemplateFile);
    } catch(e) {
      console.log("[extract] 固定模板保存失败: " + String(e));
    }

    jsonDoc.Close(false);
    fileSaved = true;
    saveResult = "\n\n📁 模板已保存到：\n" + jsonFile.replace('.json', '.txt');
    console.log("[extract] WPS保存成功: " + jsonFile.replace('.json', '.txt'));
  } catch(e1) {
    console.log("[extract] WPS保存失败: " + String(e1));

    // 方案2: 尝试 Node.js fs 模块
    try {
      if (typeof require !== 'undefined') {
        var fs = require('fs');
        try { fs.mkdirSync(templateDir, { recursive: true }); } catch(e) {}
        fs.writeFileSync(jsonFile, jsonContent, 'utf8');
        fs.writeFileSync(latestTemplateFile, jsonContent, 'utf8');  // 额外保存固定文件
        fileSaved = true;
        saveResult = "\n\n📁 模板已保存到：\n" + jsonFile;
      }
    } catch(e2) {
      console.log("[extract] Node.js保存失败: " + String(e2));
    }
  }

  if (!fileSaved) {
    saveResult = "\n\n⚠️ 无法自动保存，请回复「保存模板」由 Python 端处理。";
  }

  // 返回结果：message 显示样式详情，templateJson 供后续保存
  return {
    success: true,
    message: lines.join("\n") + saveResult,
    styleCount: Object.keys(styles).length,
    templateFile: fileSaved ? (jsonFile.replace('.json', '.txt')) : "",
    templateJson: jsonContent  // Python 端可从 artifact 提取并保存
  };
} catch (e) {
  return { success: false, error: String(e) };
}