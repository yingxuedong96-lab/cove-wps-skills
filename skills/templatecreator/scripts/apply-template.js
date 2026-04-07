/**
 * apply-template.js - 将样式模板应用到当前文档
 *
 * 参数：
 * - templateName: 模板名称（从templates目录查找）
 * - templatePath: 模板路径（直接指定）
 * - templateJson: 直接传入模板对象（用于提取后立即应用）
 * - confirmMapping: 用户确认的格式映射（从askUser返回后传入）
 *
 * 返回值：
 * - success: true/false
 * - applied: 各类型应用的段落数统计
 * - message: 提示信息
 */

(function() {
  const params = Application.Env?.ScriptParams || {};
  const DOC = Application.ActiveDocument;

  if (!DOC) {
    return JSON.stringify({ success: false, error: "没有打开的文档" });
  }

  // 加载模板
  let template = null;

  if (params.templateJson) {
    // 直接使用传入的模板
    template = params.templateJson;
  } else if (params.templatePath) {
    // 从路径加载
    try {
      const skillPath = Application.Env?.SkillPath || '';
      const content = Application.LoadFile?.(`${skillPath}/${params.templatePath}`);
      if (content) {
        template = JSON.parse(content);
      }
    } catch (e) {
      return JSON.stringify({ success: false, error: `无法加载模板: ${params.templatePath}` });
    }
  } else if (params.templateName) {
    // 按名称查找
    try {
      const skillPath = Application.Env?.SkillPath || '';
      const templatesDir = `${skillPath}/templates`;
      const files = Application.ListFiles?.(templatesDir) || [];

      for (const file of files) {
        if (file.endsWith('.json')) {
          const content = Application.LoadFile?.(`${templatesDir}/${file}`);
          const t = JSON.parse(content);
          if (t.name === params.templateName || file.includes(params.templateName)) {
            template = t;
            break;
          }
        }
      }

      if (!template) {
        return JSON.stringify({ success: false, error: `未找到模板: ${params.templateName}` });
      }
    } catch (e) {
      return JSON.stringify({ success: false, error: `查找模板失败: ${e.message}` });
    }
  } else {
    return JSON.stringify({ success: false, error: "请指定templateName、templatePath或templateJson" });
  }

  // 处理用户确认的映射
  if (params.confirmMapping && template.styles) {
    // 更新模板中的类型映射
    for (const [type, formatKey] of Object.entries(params.confirmMapping)) {
      // 找到对应格式的样式定义并更新类型
      const matchingStyle = template.styles.find(s =>
        formatKey.includes(`${s.format?.fontSize}pt`) &&
        formatKey.includes(s.format?.fontCN)
      );
      if (matchingStyle) {
        matchingStyle.type = type;
        matchingStyle.name = getTypeName(type);
      }
    }
  }

  function getTypeName(type) {
    const names = {
      docTitle: '主标题', heading1: '一级标题', heading2: '二级标题',
      heading3: '三级标题', heading4: '四级标题', heading5: '五级标题',
      body: '正文', figureCaption: '图名', tableCaption: '表名',
      listItem: '列表项', appendixTitle: '附录标题', appendixSection: '附录节题'
    };
    return names[type] || type;
  }

  // 应用格式到段落
  function applyFormat(para, format) {
    const range = para.Range;
    const paraFormat = para.Format;

    // 字体
    if (format.fontCN) range.Font.NameFarEast = format.fontCN;
    if (format.fontEN) range.Font.NameAscii = format.fontEN;
    if (format.fontSize) range.Font.Size = format.fontSize;
    if (format.bold !== undefined) range.Font.Bold = format.bold;
    if (format.italic !== undefined) range.Font.Italic = format.italic;

    // 段落格式
    if (format.alignment !== undefined) paraFormat.Alignment = format.alignment;
    if (format.firstLineIndent) paraFormat.FirstLineIndent = format.firstLineIndent * 240;
    if (format.leftIndent) paraFormat.LeftIndent = format.leftIndent * 240;
    if (format.lineSpacing) paraFormat.LineSpacing = format.lineSpacing;
    if (format.lineSpacingRule) paraFormat.LineSpacingRule = format.lineSpacingRule;
    if (format.spaceBefore) paraFormat.SpaceBefore = format.spaceBefore;
    if (format.spaceAfter) paraFormat.SpaceAfter = format.spaceAfter;
  }

  // 检测段落类型
  function detectType(para, styles) {
    const text = para.Range.Text.trim();
    if (!text) return null;

    // 按模式匹配
    for (const style of styles) {
      if (!style.detect?.pattern) continue;

      if (style.detect.pattern === 'default') {
        // 正文作为默认
        continue;
      }

      try {
        if (new RegExp(style.detect.pattern).test(text)) {
          return style.type;
        }
      } catch (e) {
        // 正则错误，跳过
      }
    }

    // 首段可能是主标题
    const paraIndex = getParaIndex(para);
    if (paraIndex === 1) {
      const docTitleStyle = styles.find(s => s.type === 'docTitle');
      if (docTitleStyle && para.Range.Font.Size >= 18) {
        return 'docTitle';
      }
    }

    // 默认正文
    return 'body';
  }

  function getParaIndex(para) {
    // 获取段落索引（从1开始）
    const paras = DOC.Paragraphs;
    for (let i = 1; i <= paras.Count; i++) {
      if (paras.Item(i) === para) return i;
    }
    return 0;
  }

  // 应用页面设置
  if (template.pageSetup) {
    const ps = DOC.PageSetup;
    if (template.pageSetup.topMargin) ps.TopMargin = template.pageSetup.topMargin * 567;
    if (template.pageSetup.bottomMargin) ps.BottomMargin = template.pageSetup.bottomMargin * 567;
    if (template.pageSetup.leftMargin) ps.LeftMargin = template.pageSetup.leftMargin * 567;
    if (template.pageSetup.rightMargin) ps.RightMargin = template.pageSetup.rightMargin * 567;
    if (template.pageSetup.paperSize) ps.PaperSize = template.pageSetup.paperSize;
  }

  // 统计应用结果
  const appliedCounts = {};
  template.styles.forEach(s => appliedCounts[s.type] = 0);

  // 遍历段落应用样式
  const paragraphs = DOC.Paragraphs;
  let firstParaProcessed = false;

  for (let i = 1; i <= paragraphs.Count; i++) {
    const para = paragraphs.Item(i);
    const text = para.Range.Text.trim();
    if (!text) continue;

    // 特殊处理首段
    if (!firstParaProcessed) {
      firstParaProcessed = true;
      const docTitleStyle = template.styles.find(s => s.type === 'docTitle');
      if (docTitleStyle && para.Range.Font.Size >= 18) {
        applyFormat(para, docTitleStyle.format);
        appliedCounts.docTitle++;
        continue;
      }
    }

    // 检测类型并应用
    const detectedType = detectType(para, template.styles);
    const matchedStyle = template.styles.find(s => s.type === detectedType);

    if (matchedStyle && matchedStyle.format) {
      applyFormat(para, matchedStyle.format);
      appliedCounts[detectedType]++;
    }
  }

  // 返回结果
  const totalApplied = Object.values(appliedCounts).reduce((a, b) => a + b, 0);

  return JSON.stringify({
    success: true,
    applied: appliedCounts,
    templateName: template.name,
    message: `已应用模板"${template.name}"，共处理${totalApplied}个段落`
  }, null, 2);

})();