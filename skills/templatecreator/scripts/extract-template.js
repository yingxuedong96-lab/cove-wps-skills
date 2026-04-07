/**
 * extract-template.js - 从当前文档提取样式模板
 *
 * 返回值：
 * - success: true/false
 * - template: 样式模板对象
 * - templatePath: 保存路径
 * - message: 提示信息
 * - needUserInput: true时需要LLM调用askUser确认
 * - uncertainFormats: 不确定的格式列表
 */

(function() {
  const DOC = Application.ActiveDocument;
  if (!DOC) {
    return JSON.stringify({ success: false, error: "没有打开的文档" });
  }

  // 格式特征检测配置
  const STYLE_TYPES = {
    docTitle: { name: "主标题", patterns: [], description: "文档首段或大字号居中" },
    heading1: { name: "一级标题", patterns: ["^\\d+\\s", "^\\d+\\s+.+"], description: "数字开头如'1 范围'" },
    heading2: { name: "二级标题", patterns: ["^\\d+\\.\\d+\\s"], description: "如'1.1 概述'" },
    heading3: { name: "三级标题", patterns: ["^\\d+\\.\\d+\\.\\d+\\s"], description: "如'1.1.1 设计'" },
    heading4: { name: "四级标题", patterns: ["^\\(\\d+\\)", "^\\d+\\)"], description: "如'(1)'" },
    heading5: { name: "五级标题", patterns: ["^[①②③④⑤⑥⑦⑧⑨⑩]"], description: "带圈数字" },
    figureCaption: { name: "图名", patterns: ["^图\\d+", "^图\\d+-\\d+"], description: "'图X'开头" },
    tableCaption: { name: "表名", patterns: ["^表\\d+", "^表\\d+-\\d+"], description: "'表X'开头" },
    appendixTitle: { name: "附录标题", patterns: ["^附录\\s*[A-Z]|^附录\\s*\\d"], description: "'附录'开头" },
    appendixSection: { name: "附录节题", patterns: ["^[A-Z]\\.[A-Z]\\s", "^附录.*\\.\\d+"], description: "附录编号" },
    listItem: { name: "列表项", patterns: ["^\\s*[-•·]", "^\\s*[a-z]\\)"], description: "列表符号开头" },
    body: { name: "正文", patterns: [], description: "默认段落" }
  };

  // 检测段落元素类型
  function detectElementType(para) {
    const text = para.Range.Text.trim();
    if (!text) return null;

    // 按优先级检查模式
    const typeOrder = ['docTitle', 'heading1', 'heading2', 'heading3', 'heading4', 'heading5',
                       'figureCaption', 'tableCaption', 'appendixTitle', 'appendixSection', 'listItem'];

    for (const type of typeOrder) {
      const config = STYLE_TYPES[type];
      for (const pattern of config.patterns) {
        if (new RegExp(pattern).test(text)) {
          return { type, confidence: 'high', pattern };
        }
      }
    }

    // 模式不匹配，使用格式特征检测
    const fontSize = para.Range.Font.Size;
    const fontName = para.Range.Font.Name;
    const bold = para.Range.Font.Bold;
    const alignment = para.Format.Alignment;

    // 大字号居中可能是主标题
    if (fontSize >= 22 && alignment === 1) { // 1 = wdAlignParagraphCenter
      return { type: 'docTitle', confidence: 'low', reason: '大字号居中' };
    }

    // 大字号加粗可能是标题
    if (fontSize >= 14 && bold) {
      if (fontSize >= 16) return { type: 'heading1', confidence: 'low', reason: '16pt加粗' };
      if (fontSize >= 15) return { type: 'heading2', confidence: 'low', reason: '15pt加粗' };
      if (fontSize >= 14) return { type: 'heading3', confidence: 'low', reason: '14pt加粗' };
    }

    // 默认为正文
    return { type: 'body', confidence: 'default' };
  }

  // 提取段落格式信息
  function extractFormat(para) {
    const range = para.Range;
    const format = para.Format;

    return {
      fontCN: range.Font.NameFarEast || range.Font.Name,
      fontEN: range.Font.NameAscii || "",
      fontSize: range.Font.Size,
      bold: range.Font.Bold,
      italic: range.Font.Italic,
      alignment: format.Alignment, // 0=左 1=中 2=右 3=两端
      firstLineIndent: format.FirstLineIndent / 240, // DXA转字符数
      leftIndent: format.LeftIndent / 240,
      lineSpacing: format.LineSpacing, // 行距值
      lineSpacingRule: format.LineSpacingRule, // 0=单倍 1=最小 4=固定
      spaceBefore: format.SpaceBefore,
      spaceAfter: format.SpaceAfter
    };
  }

  // 格式特征转可读描述
  function formatKey(format) {
    const alignMap = { 0: '左对齐', 1: '居中', 2: '右对齐', 3: '两端对齐' };
    return `${format.fontSize}pt${format.bold ? '加粗' : ''}${format.fontCN}${alignMap[format.alignment] || ''}`;
  }

  // 主提取逻辑
  const paragraphs = DOC.Paragraphs;
  const styleMap = {}; // 类型 -> 格式统计
  const uncertainFormats = [];
  let isFirstNonEmpty = true;

  for (let i = 1; i <= paragraphs.Count; i++) {
    const para = paragraphs.Item(i);
    const text = para.Range.Text.trim();
    if (!text) continue;

    const detection = detectElementType(para);
    if (!detection) continue;

    const format = extractFormat(para);
    const key = formatKey(format);

    // 首段特殊处理
    if (isFirstNonEmpty) {
      isFirstNonEmpty = false;
      // 首段大字号可能是主标题
      if (format.fontSize >= 18) {
        detection.type = 'docTitle';
        detection.confidence = 'high';
        detection.reason = '首段';
      }
    }

    // 低置信度检测，记录不确定格式
    if (detection.confidence === 'low') {
      const existing = uncertainFormats.find(f => f.formatKey === key);
      if (existing) {
        existing.count++;
        if (existing.samples.length < 3) existing.samples.push(text.substring(0, 30));
      } else {
        uncertainFormats.push({
          formatKey: key,
          suggestedType: detection.type,
          count: 1,
          samples: [text.substring(0, 30)],
          reason: detection.reason
        });
      }
    }

    // 收集格式数据
    if (!styleMap[detection.type]) {
      styleMap[detection.type] = { formats: [], patterns: [] };
    }
    styleMap[detection.type].formats.push(format);
    if (detection.pattern) {
      styleMap[detection.type].patterns.push(detection.pattern);
    }
  }

  // 合并同类格式（取众数或首个）
  function mergeFormats(formatList) {
    if (!formatList.length) return null;

    // 按字体字号分组
    const groups = {};
    formatList.forEach(f => {
      const key = `${f.fontCN}_${f.fontSize}_${f.bold}`;
      if (!groups[key]) groups[key] = [];
      groups[key].push(f);
    });

    // 取最大的组
    let maxGroup = null;
    let maxCount = 0;
    for (const [key, group] of Object.entries(groups)) {
      if (group.length > maxCount) {
        maxCount = group.length;
        maxGroup = group;
      }
    }

    // 合并属性
    const merged = { ...maxGroup[0] };
    return merged;
  }

  // 生成模板
  const template = {
    name: DOC.Name.replace('.docx', '').replace('.doc', '') + '_模板',
    version: '1.0',
    extractedFrom: DOC.Name,
    extractedAt: new Date().toISOString().split('T')[0],
    styles: [],
    pageSetup: {
      topMargin: DOC.PageSetup.TopMargin / 567, // DXA转cm
      bottomMargin: DOC.PageSetup.BottomMargin / 567,
      leftMargin: DOC.PageSetup.LeftMargin / 567,
      rightMargin: DOC.PageSetup.RightMargin / 567,
      paperSize: DOC.PageSetup.PaperSize
    }
  };

  // 类型排序
  const typeOrder = ['docTitle', 'heading1', 'heading2', 'heading3', 'heading4', 'heading5',
                     'body', 'figureCaption', 'tableCaption', 'listItem', 'appendixTitle', 'appendixSection'];

  for (const type of typeOrder) {
    const data = styleMap[type];
    if (!data || !data.formats.length) continue;

    const mergedFormat = mergeFormats(data.formats);
    const styleDef = STYLE_TYPES[type];

    const styleEntry = {
      type: type,
      name: styleDef.name,
      detect: {
        pattern: data.patterns.length ? data.patterns[0] : null,
        description: styleDef.description
      },
      format: mergedFormat
    };

    // 如果是正文，添加默认标记
    if (type === 'body') {
      styleEntry.detect.pattern = 'default';
    }

    template.styles.push(styleEntry);
  }

  // 检查是否需要用户确认
  if (uncertainFormats.length > 0 && uncertainFormats.some(f => f.suggestedType === 'docTitle' || f.suggestedType.startsWith('heading'))) {
    // 有不确定的标题类格式，需要确认
    const significantUncertain = uncertainFormats.filter(f =>
      f.suggestedType === 'docTitle' || f.suggestedType.startsWith('heading')
    );

    return JSON.stringify({
      success: true,
      needUserInput: true,
      uncertainFormats: significantUncertain,
      partialTemplate: template,
      question: `检测到以下格式可能对应不同的标题级别，请确认映射关系：\n${significantUncertain.map(f =>
        `- ${f.formatKey}（${f.count}处）：示例 "${f.samples[0]}"，建议为${STYLE_TYPES[f.suggestedType]?.name || f.suggestedType}`
      ).join('\n')}\n\n请选择或指定正确的映射。`
    }, null, 2);
  }

  // 无需确认，保存模板
  const templatesDir = 'templates';
  const templateFileName = `模板_${new Date().toISOString().split('T')[0].replace(/-/g, '')}.json`;

  // 使用WPS文件系统保存（如果可用）
  try {
    // 尝试写入到技能目录
    const skillPath = Application.Env?.SkillPath || '';
    if (skillPath) {
      // cove-wps提供的保存接口
      Application.SaveFile?.(`${skillPath}/${templatesDir}/${templateFileName}`, JSON.stringify(template, null, 2));
    }
  } catch (e) {
    // 保存失败，返回模板让用户手动保存
  }

  return JSON.stringify({
    success: true,
    template: template,
    message: `已提取${template.styles.length}种样式`,
    templatePath: `${templatesDir}/${templateFileName}`
  }, null, 2);

})();