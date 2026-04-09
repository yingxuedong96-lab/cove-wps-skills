/**
 * format-engine.js - Range批量处理版 + 规范文本校验 + 长文档优化
 *
 * 核心优化：
 * 1. 使用 Range 对象批量设置格式，而非逐段落设置
 * 2. 预排除表格和图片段落，避免重复处理
 * 3. 超长正文使用整体Range一次应用（参考zslong的useWideBodyApply）
 * 4. specText 参数校验，确保只处理用户规范中明确提到的类型
 */

try {
  var startTime = Date.now();
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, error: '无活动文档' };
  }

  var paraCount = doc.Paragraphs.Count || 0;
  console.log('[format] 开始处理，段落数: ' + paraCount);

  // 字号映射表（中文字号 → 磅值）
  var FONT_SIZE_MAP = {
    '初号': 42, '小初': 36, '一号': 26, '小一': 24,
    '二号': 22, '小二': 18, '三号': 16, '小三': 15,
    '四号': 14, '小四': 12, '五号': 10.5, '小五': 9,
    '六号': 7.5, '小六': 6.5, '七号': 5.5, '八号': 5
  };

  // 解析中文字号
  function parseFontSize(sizeText) {
    if (!sizeText) return 0;
    var s = sizeText.replace(/\s/g, '');
    if (FONT_SIZE_MAP[s]) return FONT_SIZE_MAP[s];
    // 尝试直接数字
    var num = parseFloat(s);
    if (num > 0) return num;
    return 0;
  }

  // 解析配置
  var configData = typeof config === 'string' ? JSON.parse(config) : config;
  var rules = configData.paragraphRules || {};
  var patterns = configData.numberingPatterns || {};
  var fontDefaults = configData.fontDefaults || { fontCN: '宋体', fontEN: 'Times New Roman' };

  // ========================================
  // 自动解析 specText 生成 paragraphRules（如果为空）
  // ========================================
  var specText = configData.specText || '';

  // 解析中文字号
  function parseFontSizeAuto(sizeText) {
    if (!sizeText) return 0;
    var s = sizeText.replace(/\s/g, '');
    if (FONT_SIZE_MAP[s]) return FONT_SIZE_MAP[s];
    var num = parseFloat(s);
    if (num > 0) return num;
    return 0;
  }

  // 从 specText 自动解析规则
  function autoParseRules(text) {
    var autoRules = {};
    var lowerText = text.toLowerCase();

    // 解析字体
    function parseFont(t) {
      if (t.indexOf('黑体') !== -1) return '黑体';
      if (t.indexOf('宋体') !== -1) return '宋体';
      if (t.indexOf('楷体') !== -1) return '楷体';
      if (t.indexOf('仿宋') !== -1) return '仿宋';
      return '宋体';
    }

    // 解析字号
    function parseSize(t) {
      var match = t.match(/(初号|小初|一号|小一|二号|小二|三号|小三|四号|小四|五号|小五|六号|小六)/);
      return match ? parseFontSizeAuto(match[1]) : 0;
    }

    // 解析对齐
    function parseAlign(t) {
      if (t.indexOf('居中') !== -1) return 1;
      if (t.indexOf('右对齐') !== -1 || t.indexOf('靠右') !== -1) return 2;
      if (t.indexOf('两端对齐') !== -1) return 3;
      if (t.indexOf('左对齐') !== -1 || t.indexOf('靠左') !== -1) return 0;
      return -1;
    }

    // 解析加粗
    function parseBold(t) {
      return t.indexOf('加粗') !== -1;
    }

    // 图名规则
    if (text.indexOf('图名') !== -1 || text.indexOf('图标题') !== -1) {
      var figMatch = text.match(/图名[^。；,，]*/);
      var figText = figMatch ? figMatch[0] : '';
      autoRules.figureCaption = {
        fontCN: parseFont(figText),
        fontSize: parseSize(figText) || 9,
        alignment: parseAlign(figText) >= 0 ? parseAlign(figText) : 1,
        bold: parseBold(figText)
      };
      console.log('[format] 自动解析图名规则: ' + JSON.stringify(autoRules.figureCaption));
    }

    // 表名规则
    if (text.indexOf('表名') !== -1 || text.indexOf('表标题') !== -1) {
      var tcMatch = text.match(/表名[^。；,，]*/);
      var tcText = tcMatch ? tcMatch[0] : '';
      autoRules.tableCaption = {
        fontCN: parseFont(tcText),
        fontSize: parseSize(tcText) || 9,
        alignment: parseAlign(tcText) >= 0 ? parseAlign(tcText) : 1,
        bold: parseBold(tcText)
      };
      console.log('[format] 自动解析表名规则: ' + JSON.stringify(autoRules.tableCaption));
    }

    // 表头规则
    if (text.indexOf('表头') !== -1) {
      var thMatch = text.match(/表头[^。；,，]*/);
      var thText = thMatch ? thMatch[0] : '';
      autoRules.tableHeader = {
        fontCN: parseFont(thText),
        fontSize: parseSize(thText) || 10.5,
        alignment: parseAlign(thText) >= 0 ? parseAlign(thText) : 1,
        bold: parseBold(thText) || true
      };
      console.log('[format] 自动解析表头规则: ' + JSON.stringify(autoRules.tableHeader));
    }

    // 表格内容规则
    if (text.indexOf('表格内容') !== -1) {
      var tcmMatch = text.match(/表格内容[^。；,，]*/);
      var tcmText = tcmMatch ? tcmMatch[0] : '';
      autoRules.tableContent = {
        fontCN: parseFont(tcmText),
        fontSize: parseSize(tcmText) || 10.5,
        alignment: parseAlign(tcmText) >= 0 ? parseAlign(tcmText) : 0,
        bold: parseBold(tcmText)
      };
      console.log('[format] 自动解析表格内容规则: ' + JSON.stringify(autoRules.tableContent));
    }

    // 正文规则
    if (text.indexOf('正文') !== -1) {
      var bodyMatch = text.match(/正文[^。；,，]*/);
      var bodyText = bodyMatch ? bodyMatch[0] : '';
      autoRules.body = {
        fontCN: parseFont(bodyText),
        fontSize: parseSize(bodyText) || 12,
        alignment: parseAlign(bodyText) >= 0 ? parseAlign(bodyText) : 3,
        bold: parseBold(bodyText)
      };
      // 首行缩进
      var indentMatch = bodyText.match(/首行缩进(\d)字符/);
      if (indentMatch) {
        autoRules.body.firstLineIndent = parseInt(indentMatch[1]) * 12;
      }
      console.log('[format] 自动解析正文规则: ' + JSON.stringify(autoRules.body));
    }

    return autoRules;
  }

  // 如果 paragraphRules 为空，自动解析
  if (Object.keys(rules).length === 0 && specText) {
    console.log('[format] paragraphRules 为空，自动解析 specText');
    rules = autoParseRules(specText);
    console.log('[format] 自动解析结果: ' + JSON.stringify(rules));
  }

  // ========================================
  // 规范文本校验（关键：只接受用户明确提到的类型）
  // ========================================
  var TYPE_KEYWORDS = {
    docTitle: ['主标题', '文档标题', '文章标题', '报告标题', '封面标题'],
    zhangTitle: ['章标题', '一级标题', '章节标题', '标题一', 'Heading 1', '章名', 'chapterTitle'],
    heading2: ['二级标题', '标题二', 'Heading 2', 'headingTwo'],
    heading3: ['三级标题', '标题三', 'Heading 3'],
    heading4: ['四级标题', '标题四', 'Heading 4'],
    heading5: ['五级标题', '标题五', 'Heading 5'],
    body: ['正文', '正文格式', '段落格式', '正文内容', '文本内容', 'content'],
    tableCaption: ['表名', '表格名', '表标题', '表号'],
    figureCaption: ['图名', '图片名', '图标题', '图号', '图片标题', '插图名', 'figure'],
    ref: ['参考文献', '引用文献', 'reference'],
    tableHeader: ['表头用', '表头字体', '表头格式'],  // 只匹配字体格式相关的
    tableContent: ['表格内容用', '表格内容字体', '单元格用', '表格正文用'],  // 更具体
    headerFooter: ['页眉页脚', '页眉用', '页脚用', '页码']
  };

  var specText = configData.specText || '';
  if (!specText) {
    console.log('[format] ⚠️ 错误：配置中缺少 specText 字段！');
    return {
      success: false,
      error: '配置缺少 specText 字段，无法校验规则类型。',
      hint: '示例：{"specText": "章标题用黑体三号居中加粗...", "paragraphRules": {...}}'
    };
  }

  console.log('[format] 规范文本校验启用，长度=' + specText.length);
  console.log('[format] specText内容: ' + specText.substring(0, 200) + (specText.length > 200 ? '...' : ''));

  // 从规范文本中提取用户提到的类型
  var mentionedTypes = [];
  for (var typeKey in TYPE_KEYWORDS) {
    var keywords = TYPE_KEYWORDS[typeKey];
    for (var k = 0; k < keywords.length; k++) {
      if (specText.indexOf(keywords[k]) !== -1) {
        mentionedTypes.push(typeKey);
        break;
      }
    }
  }

  // ⚠️ 类型名别名映射：Agent可能生成的错误类型名 → 正确类型名
  var ALIAS_TO_TYPE = {
    'reference': 'ref',
    'chapterTitle': 'zhangTitle',
    'headingTwo': 'heading2',
    'headingThree': 'heading3',
    'headingFour': 'heading4',
    'content': 'body',
    'figure': 'figureCaption',
    'table': 'tableCaption'
  };

  // 转换规则中的错误类型名
  var correctedRules = {};
  var originalTypes = Object.keys(rules);
  console.log('[format] Agent原始生成的类型: ' + originalTypes.join(', '));
  for (var key in rules) {
    var correctKey = ALIAS_TO_TYPE[key] || key;
    if (correctKey !== key) {
      console.log('[format] 类型名修正: ' + key + ' → ' + correctKey);
    }
    correctedRules[correctKey] = rules[key];
  }
  rules = correctedRules;

  console.log('[format] 规范文本中提到的类型: ' + mentionedTypes.join(', '));

  // 过滤规则：只保留规范中提到的类型
  var filteredRules = {};
  var removedTypes = [];
  for (var key in rules) {
    if (mentionedTypes.indexOf(key) !== -1) {
      filteredRules[key] = rules[key];
    } else {
      removedTypes.push(key);
    }
  }

  if (removedTypes.length > 0) {
    console.log('[format] ⚠️ 已移除未提及的规则类型: ' + removedTypes.join(', '));
    console.log('[format] ⚠️ 规范文本只允许: ' + mentionedTypes.join(', '));
  }

  // ⚠️ 检查是否有应该有但缺失的规则
  var missingTypes = [];
  for (var m = 0; m < mentionedTypes.length; m++) {
    if (!filteredRules[mentionedTypes[m]]) {
      missingTypes.push(mentionedTypes[m]);
    }
  }
  if (missingTypes.length > 0) {
    console.log('[format] ⚠️ 规范提到但配置缺失的类型: ' + missingTypes.join(', '));
  }

  rules = filteredRules;

  // 修正编号正则
  for (var key in patterns) {
    if (Array.isArray(patterns[key])) {
      for (var i = 0; i < patterns[key].length; i++) {
        if (typeof patterns[key][i] === 'string') {
          patterns[key][i] = patterns[key][i].replace(/第\s+\[/g, '第[');
        }
      }
    }
  }

  // 修正规则配置
  for (var key in rules) {
    var rule = rules[key];
    // 行距修正
    if (rule.lineSpacing !== undefined && rule.lineSpacingRule === undefined) {
      if (rule.lineSpacing === 1.5) { rule.lineSpacingRule = 1; delete rule.lineSpacing; }
      else if (rule.lineSpacing === 1) { rule.lineSpacingRule = 0; delete rule.lineSpacing; }
      else if (rule.lineSpacing === 2) { rule.lineSpacingRule = 2; delete rule.lineSpacing; }
    }
    // 首行缩进修正（正值：字符数→磅值）
    if (rule.firstLineIndent !== undefined && rule.firstLineIndent > 0 && rule.firstLineIndent < 10) {
      rule.firstLineIndent = rule.firstLineIndent * 12;
    }
    // 悬挂缩进修正（负值：字符数→磅值，如 -2 → -24）
    if (rule.firstLineIndent !== undefined && rule.firstLineIndent < 0 && rule.firstLineIndent > -10) {
      rule.firstLineIndent = rule.firstLineIndent * 12;
      console.log('[format] 悬挂缩进转换: ' + (rule.firstLineIndent / 12) + '字符 → ' + rule.firstLineIndent + '磅');
    }
    // ⚠️ 悬挂缩进必须同时设置 LeftIndent（后续行缩进量）
    if (rule.firstLineIndent !== undefined && rule.firstLineIndent < 0) {
      rule.leftIndent = -rule.firstLineIndent;  // 后续行缩进 = 首行负缩进的绝对值
      console.log('[format] 悬挂缩进: LeftIndent=' + rule.leftIndent + '磅, FirstLineIndent=' + rule.firstLineIndent + '磅');
    }
  }

  console.log('[format] 配置解析完成');

  // ========================================
  // 读取全文并分类
  // ========================================
  var docText = '';
  try { docText = doc.Content ? doc.Content.Text : ''; } catch (e) {}
  var logicalParas = docText.split('\r');
  var maxParaCount = Math.min(paraCount, logicalParas.length);
  var useUltraFastMode = paraCount > 6000;

  // ⚠️ 简化处理：不做表格段落排除
  // 表格有自己的样式，不会被正文整体 Range 设置影响
  // 避免遍历检查导致循环超限

  console.log('[format] 表格数: ' + (doc.Tables ? doc.Tables.Count : 0) + ', 段落数: ' + paraCount);

  // 默认正则 + 用户正则
  var allPatterns = {
    zhangTitle: ['^第[一二三四五六七八九十百零]+章', '^第\\s*[一二三四五六七八九十百零]+章', '^第\\d{1,3}章', '^\\d{1,3}\\s+[\\u4e00-\\u9fff]', '^[一二三四五六七八九十]+、'],
    heading2: ['^\\d+\\.\\d+\\s', '^[（(][一二三四五六七八九十]+[）)]'],
    heading3: ['^\\d+\\.\\d+\\.\\d+\\s', '^[（(]\\d+[）)]'],
    heading4: ['^\\d+\\.\\d+\\.\\d+\\.\\d+\\s', '^[（(][一二三四五六七八九十]+[）)][（(]\\d+[）)]'],
    heading5: ['^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+\\s'],
    tableCaption: ['^表\\s*\\d+', '^表\\s*[\\d\\.\\-]+'],
    figureCaption: ['^图\\s*\\d+', '^图\\s*[\\d\\.\\-]+'],
    ref: ['^参考文献', '^\\[\\d+\\]'],
    formula: []  // 公式通过特征检测，不用正则
  };

  // 公式编号正则（用于后续处理）
  var formulaNumRegex = /\(([A-Z]?\d+(?:\.\d+)*(?:[-—–]\d+)?)\)\s*$/;
  var mathSymbols = /[=±×÷∑ΣΔ√∫≈≠≤≥λρμτωφψβγαεδ]/;

  for (var t in allPatterns) {
    allPatterns[t] = (patterns[t] || []).concat(allPatterns[t]);
  }

  var compiled = {};
  for (var t in allPatterns) {
    compiled[t] = [];
    for (var i = 0; i < allPatterns[t].length; i++) {
      try { compiled[t].push(new RegExp(allPatterns[t][i])); } catch (e) {}
    }
  }

  function cleanText(t) {
    return String(t || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  var inRef = false;
  function classify(text) {
    text = cleanText(text);
    if (!text) return 'empty';

    for (var i = 0; i < compiled.ref.length; i++) {
      if (compiled.ref[i].test(text)) {
        if (/^参考文献/.test(text)) return 'zhangTitle';
        inRef = true;
        return 'ref';
      }
    }
    if (inRef) return 'ref';

    for (var i = 0; i < compiled.tableCaption.length; i++) {
      if (compiled.tableCaption[i].test(text)) return 'tableCaption';
    }
    for (var i = 0; i < compiled.figureCaption.length; i++) {
      if (compiled.figureCaption[i].test(text)) return 'figureCaption';
    }
    for (var i = 0; compiled.zhangTitle && i < compiled.zhangTitle.length; i++) {
      if (compiled.zhangTitle[i].test(text)) return 'zhangTitle';
    }
    // heading5 必须最先检测，因为 1.1.1.1.1 也匹配 1.1.1.1、1.1.1、1.1
    for (var i = 0; compiled.heading5 && i < compiled.heading5.length; i++) {
      if (compiled.heading5[i].test(text)) return 'heading5';
    }
    // heading4 必须在 heading2/heading3 之前检测
    for (var i = 0; compiled.heading4 && i < compiled.heading4.length; i++) {
      if (compiled.heading4[i].test(text)) return 'heading4';
    }
    for (var i = 0; compiled.heading3 && i < compiled.heading3.length; i++) {
      if (compiled.heading3[i].test(text)) return 'heading3';
    }
    for (var i = 0; compiled.heading2 && i < compiled.heading2.length; i++) {
      if (compiled.heading2[i].test(text)) return 'heading2';
    }

    return text.length <= 2 ? 'short' : 'body';
  }

  // 分类所有段落
  var classifications = [];
  var typeIndices = {};  // 按类型存储段落索引
  var excludedCount = 0;

  // 只初始化用户规则中定义的类型
  for (var t in rules) {
    if (rules[t]) typeIndices[t] = [];
  }

  for (var i = 0; i < maxParaCount; i++) {
    var paraIndex = i + 1;  // WPS 段落索引从1开始
    var type = classify(logicalParas[i]);
    classifications.push(type);

    // 只记录用户规则中定义的类型（不做表格排除）
    if (type !== 'empty' && type !== 'short' && typeIndices[type]) {
      typeIndices[type].push(paraIndex);
    }
  }

  // 处理文档标题（docTitle）：如果用户定义了 docTitle 规则，将第一个非空段落标记为 docTitle
  if (typeIndices['docTitle']) {
    var firstNonEmptyIndex = -1;
    var firstNonEmptyType = '';
    for (var i = 0; i < classifications.length; i++) {
      if (classifications[i] !== 'empty') {
        firstNonEmptyIndex = i + 1;  // WPS 段落索引从1开始
        firstNonEmptyType = classifications[i];
        break;
      }
    }
    // 如果第一个非空段落不是其他标题类型，标记为 docTitle
    if (firstNonEmptyIndex > 0 && firstNonEmptyType !== 'zhangTitle' && firstNonEmptyType !== 'heading2' && firstNonEmptyType !== 'heading3' && firstNonEmptyType !== 'heading4' && firstNonEmptyType !== 'tableCaption' && firstNonEmptyType !== 'figureCaption' && firstNonEmptyType !== 'ref') {
      // 从 body 中移除，添加到 docTitle
      if (typeIndices['body']) {
        var bodyIdx = typeIndices['body'].indexOf(firstNonEmptyIndex);
        if (bodyIdx !== -1) typeIndices['body'].splice(bodyIdx, 1);
      }
      if (typeIndices['docTitle'].indexOf(firstNonEmptyIndex) === -1) {
        typeIndices['docTitle'].push(firstNonEmptyIndex);
        classifications[firstNonEmptyIndex - 1] = 'docTitle';
        console.log('[format] 文档标题(docTitle): 段落 ' + firstNonEmptyIndex);
      }
    }
  }

  // 统计
  var counts = {};
  for (var i = 0; i < classifications.length; i++) {
    counts[classifications[i]] = (counts[classifications[i]] || 0) + 1;
  }

  // 显示分类结果和规则匹配情况
  console.log('[format] 分类结果: ' + JSON.stringify(counts));
  var definedTypes = Object.keys(rules).filter(function(t) { return rules[t]; });
  console.log('[format] 用户定义的规则类型: ' + definedTypes.join(', '));

  // 显示各规则的完整配置
  for (var t in rules) {
    if (rules[t]) {
      console.log('[format] ' + t + '规则: ' + JSON.stringify(rules[t]));
    }
  }

  // ========================================
  // 批量应用格式（Range优化）
  // ========================================
  var origTrack = false;
  try { origTrack = doc.TrackRevisions; doc.TrackRevisions = false; } catch (e) {}

  var applied = 0;
  var errors = 0;

  // 辅助函数：对 Range 应用格式规则
  function applyRuleToRange(range, rule) {
    if (!range) return false;
    try {
      if (range.Font) {
        var f = range.Font;
        if (rule.fontCN || fontDefaults.fontCN) f.NameFarEast = rule.fontCN || fontDefaults.fontCN;
        if (rule.fontEN || fontDefaults.fontEN) f.Name = rule.fontEN || fontDefaults.fontEN;
        if (rule.fontSize !== undefined) f.Size = rule.fontSize;
        if (rule.bold !== undefined) f.Bold = rule.bold ? -1 : 0;
      }
      if (range.ParagraphFormat) {
        var pf = range.ParagraphFormat;
        if (rule.alignment !== undefined) pf.Alignment = rule.alignment;
        if (rule.firstLineIndent !== undefined) pf.FirstLineIndent = rule.firstLineIndent;
        if (rule.leftIndent !== undefined) pf.LeftIndent = rule.leftIndent;  // 悬挂缩进需要
        if (rule.spaceBefore !== undefined) pf.SpaceBefore = rule.spaceBefore;
        if (rule.spaceAfter !== undefined) pf.SpaceAfter = rule.spaceAfter;
        if (rule.lineSpacingRule !== undefined) pf.LineSpacingRule = rule.lineSpacingRule;
        if (rule.lineSpacing !== undefined) pf.LineSpacing = rule.lineSpacing;
      }
      return true;
    } catch (e) { return false; }
  }

  try {
    // 正文批量处理：用连续 Range 一次设置
    var bodyIndices = typeIndices['body'] || [];
    if (bodyIndices.length > 0 && rules.body) {
      // ⚠️ 长文档优化：根据分段数选择策略
      var segments = [];
      var segStart = bodyIndices[0];
      var segEnd = bodyIndices[0];

      for (var i = 1; i < bodyIndices.length; i++) {
        if (bodyIndices[i] === segEnd + 1) {
          segEnd = bodyIndices[i];
        } else {
          segments.push({ start: segStart, end: segEnd });
          segStart = bodyIndices[i];
          segEnd = bodyIndices[i];
        }
      }
      segments.push({ start: segStart, end: segEnd });

      console.log('[format] 正文分段数: ' + segments.length + ', 总段落: ' + bodyIndices.length);

      // ⚠️ 长文档优化：直接用整体Range，不检查表格
      // 表格内容有自己的样式，正文格式设置不会影响表格
      var tryWideRange = segments.length > 30 || useUltraFastMode;

      if (tryWideRange && bodyIndices.length > 100) {
        try {
          var firstBodyPara = doc.Paragraphs.Item(bodyIndices[0]);
          var lastBodyPara = doc.Paragraphs.Item(bodyIndices[bodyIndices.length - 1]);
          if (firstBodyPara && firstBodyPara.Range && lastBodyPara && lastBodyPara.Range) {
            var wideRange = doc.Range(firstBodyPara.Range.Start, lastBodyPara.Range.End);
            if (applyRuleToRange(wideRange, rules.body)) {
              applied += bodyIndices.length;
              console.log('[format] 正文整体Range处理: ' + bodyIndices.length + '段');
              // 直接跳过后续分段处理
              bodyIndices = [];
            }
          }
        } catch (wideErr) {
          console.log('[format] 整体Range失败: ' + wideErr);
        }
      }

      // 分段处理（整体Range失败时才执行）
      if (bodyIndices.length > 0) {
        var processedCount = 0;
        for (var s = 0; s < segments.length; s++) {
          try {
            var startPara = doc.Paragraphs.Item(segments[s].start);
            var endPara = doc.Paragraphs.Item(segments[s].end);
            if (!startPara || !startPara.Range || !endPara || !endPara.Range) continue;
            var segRange = doc.Range(startPara.Range.Start, endPara.Range.End);
            if (applyRuleToRange(segRange, rules.body)) {
              processedCount += segments[s].end - segments[s].start + 1;
            }
          } catch (segErr) {}
        }
        applied += processedCount;
        console.log('[format] 正文分段处理: ' + processedCount + '段');
      }

      // ⚠️ 修复图片段落行距：固定行距会裁剪图片，改用单倍行距
      try {
        var inlineCount = doc.InlineShapes ? doc.InlineShapes.Count : 0;
        var fixedImagePara = 0;
        for (var imgIdx = 1; imgIdx <= inlineCount; imgIdx++) {
          try {
            var inlineShape = doc.InlineShapes.Item(imgIdx);
            if (inlineShape && inlineShape.Range && inlineShape.Range.ParagraphFormat) {
              // 图片段落改用单倍行距，避免图片被裁剪
              inlineShape.Range.ParagraphFormat.LineSpacingRule = 0;  // 单倍行距
              fixedImagePara++;
            }
          } catch (e) {}
        }
        if (fixedImagePara > 0) {
          console.log('[format] 图片段落行距修复: ' + fixedImagePara + '个');
        }
      } catch (e) {}
    }

    // 标题、图表等：逐个处理（数量少）
    // 大纲级别映射（用于导航窗格和目录生成）
    var outlineLevelMap = {
      'docTitle': 0,     // 文档标题（无大纲级别）
      'zhangTitle': 1,   // 一级大纲
      'heading2': 2,     // 二级大纲
      'heading3': 3,     // 三级大纲
      'heading4': 4,     // 四级大纲
      'heading5': 5      // 五级大纲
    };

    var singleTypes = ['docTitle', 'zhangTitle', 'heading2', 'heading3', 'heading4', 'heading5', 'tableCaption', 'figureCaption', 'ref'];
    for (var t = 0; t < singleTypes.length; t++) {
      var typeName = singleTypes[t];
      var indices = typeIndices[typeName] || [];
      var rule = rules[typeName];
      if (!rule || indices.length === 0) continue;

      // 获取该类型对应的大纲级别
      var outlineLevel = outlineLevelMap[typeName];

      for (var i = 0; i < indices.length; i++) {
        try {
          var para = doc.Paragraphs.Item(indices[i]);
          if (para && para.Range) {
            if (applyRuleToRange(para.Range, rule)) {
              applied++;
            }
            // 设置大纲级别（用于导航窗格和目录）
            if (outlineLevel && para.Range.ParagraphFormat) {
              try {
                para.Range.ParagraphFormat.OutlineLevel = outlineLevel;
              } catch (e) {}
            }
          }
        } catch (e) { errors++; }
      }
      console.log('[format] ' + typeName + ': ' + indices.length + '段' + (outlineLevel ? ' (大纲' + outlineLevel + ')' : ''));
    }

    // ========================================
    // 表格处理（tableHeader / tableContent）
    // ========================================
    var tableRules = rules.tableHeader || rules.tableContent;
    if (tableRules) {
      console.log('[format] 启用表格处理');
      try {
        var tableCount = doc.Tables ? doc.Tables.Count : 0;
        var tableApplied = 0;

        for (var tblIdx = 1; tblIdx <= tableCount; tblIdx++) {
          try {
            var table = doc.Tables.Item(tblIdx);
            if (!table || !table.Rows) continue;

            var rowCount = table.Rows.Count;

            // 处理表头（第一行整体Range）
            if (rules.tableHeader && rowCount > 0) {
              try {
                var headerRow = table.Rows.Item(1);
                if (headerRow && headerRow.Range) {
                  applyRuleToRange(headerRow.Range, rules.tableHeader);
                  tableApplied += headerRow.Cells ? headerRow.Cells.Count : 1;
                }
              } catch (e) {}
            }

            // 处理表格内容（其余行整体Range）
            if (rules.tableContent && rowCount > 1) {
              try {
                // 从第2行到最后一行创建Range
                var startRow = table.Rows.Item(2);
                var endRow = table.Rows.Item(rowCount);
                if (startRow && startRow.Range && endRow && endRow.Range) {
                  var contentRange = doc.Range(startRow.Range.Start, endRow.Range.End);
                  applyRuleToRange(contentRange, rules.tableContent);
                  tableApplied += rowCount - 1;
                }
              } catch (e) {
                // 如果整体Range失败，逐行处理
                for (var r = 2; r <= rowCount; r++) {
                  try {
                    var row = table.Rows.Item(r);
                    if (row && row.Range) {
                      applyRuleToRange(row.Range, rules.tableContent);
                      tableApplied++;
                    }
                  } catch (e2) {}
                }
              }
            }
          } catch (e) {}
        }
        applied += tableApplied;
        console.log('[format] 表格处理: ' + tableCount + '个表格, ' + tableApplied + '行');
      } catch (e) {
        console.log('[format] 表格处理失败: ' + e);
      }
    }

    // ========================================
    // 图片处理（居中/左对齐/右对齐）
    // ========================================
    var elementSettings = configData.elementSettings || {};
    var specTextLower = specText.toLowerCase();

    // 从 specText 自动识别图片对齐规则（支持多种写法）
    // ⚠️ 优先级：明确指定 > 隐含居中
    var imageAlign = -1;  // -1=未指定, 0=左, 1=居中, 2=右

    // 明确指定对齐方式
    if (specTextLower.indexOf('图片左对齐') !== -1 || specTextLower.indexOf('图片靠左') !== -1) {
      imageAlign = 0;
      console.log('[format] 图片对齐: 左对齐');
    } else if (specTextLower.indexOf('图片右对齐') !== -1 || specTextLower.indexOf('图片靠右') !== -1) {
      imageAlign = 2;
      console.log('[format] 图片对齐: 右对齐');
    } else if (specTextLower.indexOf('图片居中') !== -1 ||
               specTextLower.indexOf('图片段落居中') !== -1 ||
               (specTextLower.indexOf('图片') !== -1 && specTextLower.indexOf('居中') !== -1) ||
               elementSettings.imageCenter) {
      imageAlign = 1;
      console.log('[format] 图片对齐: 居中');
    }

    if (imageAlign >= 0) {
      elementSettings.imageAlign = imageAlign;
    }

    if (elementSettings.imageAlign >= 0 || elementSettings.imageCenter) {
      // 兼容旧配置
      var finalAlign = elementSettings.imageAlign || (elementSettings.imageCenter ? 1 : -1);
      if (finalAlign < 0) finalAlign = 1;  // 默认居中

      try {
        var inlineCount = doc.InlineShapes ? doc.InlineShapes.Count : 0;
        var imageAligned = 0;
        for (var imgIdx = 1; imgIdx <= inlineCount; imgIdx++) {
          try {
            var inlineShape = doc.InlineShapes.Item(imgIdx);
            if (inlineShape && inlineShape.Range && inlineShape.Range.ParagraphFormat) {
              inlineShape.Range.ParagraphFormat.Alignment = finalAlign;
              imageAligned++;
            }
          } catch (e) {}
        }
        if (imageAligned > 0) {
          applied += imageAligned;
          console.log('[format] 图片对齐(' + (finalAlign===0?'左':finalAlign===1?'居中':'右') + '): ' + imageAligned + '个');
        }
      } catch (e) {
        console.log('[format] 图片处理失败: ' + e);
      }
    }

    // ========================================
    // 表格处理（等宽、跨页重复表头）
    // ========================================
    // 从 specText 自动识别表格规则
    if (specTextLower.indexOf('表格等宽') !== -1 || specTextLower.indexOf('与页面等宽') !== -1 || elementSettings.tableFullWidth) {
      elementSettings.tableFullWidth = true;
      console.log('[format] 启用表格等宽');
    }
    if (specTextLower.indexOf('跨页重复') !== -1 || specTextLower.indexOf('重复表头') !== -1 || elementSettings.tableHeadingRepeat) {
      elementSettings.tableHeadingRepeat = true;
      console.log('[format] 启用跨页重复表头');
    }

    if (elementSettings.tableFullWidth || elementSettings.tableHeadingRepeat) {
      try {
        var tableCount = doc.Tables ? doc.Tables.Count : 0;
        var tablesProcessed = 0;

        for (var tblIdx = 1; tblIdx <= tableCount; tblIdx++) {
          try {
            var table = doc.Tables.Item(tblIdx);
            if (!table) continue;

            // 表格等宽：设置为页面宽度
            if (elementSettings.tableFullWidth) {
              try {
                // 获取页面宽度（从第一个节）
                var pageWidth = 595;  // 默认A4宽度约595磅（210mm）
                var leftMargin = 72;   // 默认左右边距各72磅（25.4mm）
                var rightMargin = 72;
                try {
                  var section = doc.Sections.Item(1);
                  if (section && section.PageSetup) {
                    pageWidth = section.PageSetup.PageWidth;
                    leftMargin = section.PageSetup.LeftMargin;
                    rightMargin = section.PageSetup.RightMargin;
                  }
                } catch (e) {}

                var usableWidth = pageWidth - leftMargin - rightMargin;

                // 尝试多种方式设置表格宽度
                try {
                  // 方法1: PreferredWidth
                  table.PreferredWidthType = 3;  // 3 = wdPreferredWidthPoints（磅值）
                  table.PreferredWidth = usableWidth;
                } catch (e1) {}

                try {
                  // 方法2: 直接设置 Columns 宽度
                  if (table.Columns) {
                    var colCount = table.Columns.Count;
                    var colWidth = usableWidth / colCount;
                    for (var c = 1; c <= colCount; c++) {
                      try {
                        table.Columns.Item(c).Width = colWidth;
                      } catch (e2) {}
                    }
                  }
                } catch (e3) {}

                try {
                  // 方法3: 禁止自动调整
                  table.AllowAutoFit = false;
                } catch (e4) {}

              } catch (e) {
                // console.log('[format] 表格宽度设置失败: ' + e);
              }
            }

            // 跨页重复表头：设置首行 HeadingFormat
            if (elementSettings.tableHeadingRepeat) {
              try {
                if (table.Rows && table.Rows.Count > 0) {
                  table.Rows.Item(1).HeadingFormat = true;
                }
              } catch (e) {}
            }

            tablesProcessed++;
          } catch (e) {}
        }

        applied += tablesProcessed;  // 计入已处理数量
        if (tablesProcessed > 0) {
          console.log('[format] 表格处理: ' + tablesProcessed + '个');
        }
      } catch (e) {
        console.log('[format] 表格处理失败: ' + e);
      }
    }

    // ========================================
    // 页眉页脚处理
    // ========================================
    var hfRule = rules.headerFooter;
    if (hfRule || specTextLower.indexOf('页眉') !== -1 || specTextLower.indexOf('页脚') !== -1 || specTextLower.indexOf('页码') !== -1) {
      console.log('[format] 启用页眉页脚处理');
      try {
        var sections = doc.Sections;
        var sectionCount = sections ? sections.Count : 0;
        var hfProcessed = 0;

        // 解析页眉页脚格式
        var hfFontSize = 9;  // 默认小五号
        var hfFontCN = '宋体';
        var hfFontEN = 'Times New Roman';
        var pageNumAlign = 1;  // 默认居中
        var pageNumStyle = 0;  // 默认阿拉伯数字
        var headerLine = false;  // 页眉线

        // 从 specText 解析
        var fontSizeMatch = specTextLower.match(/小五号|五号|小四号|四号/);
        if (fontSizeMatch) {
          hfFontSize = parseFontSize(fontSizeMatch[0]);
        }

        if (specTextLower.indexOf('页码居中') !== -1) pageNumAlign = 1;
        else if (specTextLower.indexOf('页码左') !== -1) pageNumAlign = 0;
        else if (specTextLower.indexOf('页码右') !== -1) pageNumAlign = 2;

        if (specTextLower.indexOf('阿拉伯') !== -1) pageNumStyle = 0;
        else if (specTextLower.indexOf('罗马') !== -1) pageNumStyle = 3;

        if (specTextLower.indexOf('页眉线') !== -1 || specTextLower.indexOf('单细线') !== -1) {
          headerLine = true;
        }

        // 从规则中获取字体
        if (hfRule) {
          if (hfRule.fontCN) hfFontCN = hfRule.fontCN;
          if (hfRule.fontEN) hfFontEN = hfRule.fontEN;
          if (hfRule.fontSize) hfFontSize = hfRule.fontSize;
        }

        console.log('[format] 页眉页脚设置: 字号=' + hfFontSize + ', 页码对齐=' + pageNumAlign);

        for (var si = 1; si <= sectionCount; si++) {
          try {
            var section = sections.Item(si);
            if (!section) continue;

            // 处理页眉
            try {
              var header = section.Headers.Item(1);  // wdHeaderFooterPrimary
              if (header && header.Range) {
                // 设置字体
                if (header.Range.Font) {
                  header.Range.Font.NameFarEast = hfFontCN;
                  header.Range.Font.Name = hfFontEN;
                  header.Range.Font.Size = hfFontSize;
                }
                // 设置页眉线（下划线）
                if (headerLine && header.Range.ParagraphFormat) {
                  try {
                    header.Range.ParagraphFormat.Borders.Item(-3).LineStyle = 1;  // wdBorderBottom = -3
                  } catch (e) {}
                }
                hfProcessed++;
              }
            } catch (e) {}

            // 处理页脚（页码）
            try {
              var footer = section.Footers.Item(1);  // wdHeaderFooterPrimary
              if (footer) {
                // 设置字体
                if (footer.Range && footer.Range.Font) {
                  footer.Range.Font.NameFarEast = hfFontCN;
                  footer.Range.Font.Name = hfFontEN;
                  footer.Range.Font.Size = hfFontSize;
                }

                // 添加/设置页码
                try {
                  var pageNumbers = footer.PageNumbers;
                  if (pageNumbers) {
                    pageNumbers.NumberStyle = pageNumStyle;
                    pageNumbers.Add(pageNumAlign);
                  }
                } catch (e) {}

                hfProcessed++;
              }
            } catch (e) {}
          } catch (e) {}
        }

        applied += hfProcessed;
        console.log('[format] 页眉页脚处理: ' + sectionCount + '节, ' + hfProcessed + '个');
      } catch (e) {
        console.log('[format] 页眉页脚处理失败: ' + e);
      }
    }

    // ========================================
    // 公式处理（居中、编号右对齐、字体字号）
    // ========================================
    if (specTextLower.indexOf('公式居中') !== -1 || specTextLower.indexOf('公式编号') !== -1 || specTextLower.indexOf('公式用') !== -1 || elementSettings.formulaLayout) {
      console.log('[format] 启用公式排版');
      try {
        var formulaCount = 0;
        var formulaWithNumberCount = 0;

        // 公式编号正则：(...) 格式，如 (1-1), (2.1-3), (A1)
        var formulaNumRegex = /\(([A-Z]?\d+(?:\.\d+)*(?:[-—–]\d+)?)\)\s*$/;

        // 解析公式字体字号设置
        var formulaFontCN = '';
        var formulaFontSize = 0;
        var formulaFontMatch = specTextLower.match(/公式用([宋黑楷仿][体])\s*(小?[一二三四五]号|初号)/);
        if (formulaFontMatch) {
          formulaFontCN = formulaFontMatch[1];
          formulaFontSize = parseFontSize(formulaFontMatch[2]);
          console.log('[format] 公式字体: ' + formulaFontCN + ', 字号: ' + formulaFontSize);
        }
        // 强数学符号（更严格）
        var strongMathSymbols = /[±×÷∑ΣΔ√∫≈≠≤≥]/;

        // 快速搜索：只找带公式编号的段落（更严格筛选）
        var formulaIndices = [];
        var fullText = doc.Content.Text || '';
        var lines = fullText.split('\r');

        for (var lineIdx = 0; lineIdx < lines.length; lineIdx++) {
          var lineText = cleanText(lines[lineIdx]);
          if (!lineText) continue;

          // 只检测有公式编号的段落（更严格）
          if (formulaNumRegex.test(lineText)) {
            formulaIndices.push(lineIdx + 1);  // 段落索引从1开始
          }
        }

        console.log('[format] 检测到公式编号段落: ' + formulaIndices.length + '个');

        // 页面尺寸（只获取一次）
        var pageWidth = 595, leftMargin = 72, rightMargin = 72;
        try {
          var ps = doc.PageSetup;
          if (ps) {
            pageWidth = ps.PageWidth || 595;
            leftMargin = ps.LeftMargin || 72;
            rightMargin = ps.RightMargin || 72;
          }
        } catch (e) {}
        var contentWidth = pageWidth - leftMargin - rightMargin;
        var centerPos = contentWidth / 2;
        var rightPos = contentWidth;

        // 只处理带编号的公式段落
        for (var fi = 0; fi < formulaIndices.length; fi++) {
          var pIdx = formulaIndices[fi];
          try {
            var para = doc.Paragraphs.Item(pIdx);
            if (!para || !para.Range) continue;

            var text = String(para.Range.Text || '').replace(/[\r\u0007]/g, '').trim();
            if (!text) continue;

            // 检查公式编号
            var numMatch = text.match(formulaNumRegex);
            if (!numMatch) continue;

            var formulaNumber = numMatch[0].trim();
            var formulaBody = text.substring(0, text.lastIndexOf(numMatch[0])).trim();

            formulaCount++;
            formulaWithNumberCount++;

            // 修改文本：\t公式\t编号
            var paraMark = /\r$/.test(para.Range.Text) ? '\r' : '';
            var newText = '\t' + formulaBody + '\t' + formulaNumber + paraMark;

            if (newText !== para.Range.Text) {
              para.Range.Text = newText;
            }

            // 设置段落格式
            try {
              if (para.Format) {
                para.Format.Alignment = 0;  // 左对齐

                // 设置制表位
                if (para.Format.TabStops) {
                  para.Format.TabStops.ClearAll();
                  para.Format.TabStops.Add(centerPos, 1, 0);  // 居中制表位
                  para.Format.TabStops.Add(rightPos, 2, 0);   // 右对齐制表位
                }
              }
            } catch (e) {}

            // 应用字体字号
            if (formulaFontCN || formulaFontSize) {
              try {
                var rng = para.Range;
                if (rng && rng.Font) {
                  if (formulaFontCN) {
                    rng.Font.Name = formulaFontCN;
                    rng.Font.NameFarEast = formulaFontCN;
                  }
                  if (formulaFontSize) {
                    rng.Font.Size = formulaFontSize;
                  }
                }
              } catch (e) {}
            }

          } catch (e) {}
        }

        console.log('[format] 公式处理: ' + formulaCount + '个');
        applied += formulaCount;

      } catch (e) {
        console.log('[format] 公式处理失败: ' + e);
      }
    }

    // ========================================
    // 页眉页脚处理
    // ========================================
    if (specTextLower.indexOf('页眉') !== -1 || specTextLower.indexOf('页脚') !== -1) {
      console.log('[format] 启用页眉页脚排版');
      try {
        var hfCount = 0;

        // 解析页眉页脚字体字号
        var hfFontCN = '宋体';
        var hfFontEN = 'Arial';
        var hfFontSize = 9;  // 小五号

        var hfFontMatch = specTextLower.match(/页眉页脚[字号]?[：:]?\s*小?[一二三四五]号/);
        if (hfFontMatch) {
          hfFontSize = parseFontSize(hfFontMatch[0].replace(/页眉页脚[字号]?[：:]?\s*/, ''));
        }

        // 解析字体
        var cnFontMatch = specTextLower.match(/中文字体[为是]\s*([宋黑楷仿][体])/);
        if (cnFontMatch) hfFontCN = cnFontMatch[1];

        var enFontMatch = specTextLower.match(/西文字体[为是]\s*([A-Za-z\s]+)/);
        if (enFontMatch) hfFontEN = enFontMatch[1].trim();

        console.log('[format] 页眉页脚字体: ' + hfFontCN + '/' + hfFontEN + ', 字号: ' + hfFontSize);

        // 处理所有节的页眉页脚
        var sections = doc.Sections;
        if (sections && sections.Count > 0) {
          for (var si = 1; si <= sections.Count; si++) {
            var sec = sections.Item(si);
            if (!sec) continue;

            // 页眉
            try {
              var header = sec.Headers.Item(1);  // wdHeaderFooterPrimary = 1
              if (header && header.Range) {
                var hfRange = header.Range;
                if (hfRange.Font) {
                  hfRange.Font.NameFarEast = hfFontCN;
                  hfRange.Font.Name = hfFontEN;
                  hfRange.Font.Size = hfFontSize;
                }

                // 页眉线（下边框）
                if (specTextLower.indexOf('页眉线') !== -1) {
                  try {
                    // 设置段落下边框
                    if (header.Range.ParagraphFormat && header.Range.ParagraphFormat.Borders) {
                      var borders = header.Range.ParagraphFormat.Borders;
                      var bottomBorder = borders.Item(-3);  // wdBorderBottom = -3
                      bottomBorder.LineStyle = 1;  // wdLineStyleSingle
                      bottomBorder.LineWidth = 0.5;
                    }
                  } catch (e) {}
                }
                hfCount++;
              }
            } catch (e) {}

            // 页脚
            try {
              var footer = sec.Footers.Item(1);
              if (footer && footer.Range) {
                var ffRange = footer.Range;
                if (ffRange.Font) {
                  ffRange.Font.NameFarEast = hfFontCN;
                  ffRange.Font.Name = hfFontEN;
                  ffRange.Font.Size = hfFontSize;
                }

                // 页脚线（上边框）
                if (specTextLower.indexOf('页脚线') !== -1) {
                  try {
                    if (footer.Range.ParagraphFormat && footer.Range.ParagraphFormat.Borders) {
                      var fborders = footer.Range.ParagraphFormat.Borders;
                      var topBorder = fborders.Item(-1);  // wdBorderTop = -1
                      // 双细线
                      topBorder.LineStyle = 6;  // wdLineStyleDouble
                      topBorder.LineWidth = 0.5;
                    }
                  } catch (e) {}
                }
                hfCount++;
              }
            } catch (e) {}
          }
        }

        console.log('[format] 页眉页脚处理: ' + hfCount + '个节');
        applied += hfCount;

      } catch (e) {
        console.log('[format] 页眉页脚处理失败: ' + e);
      }
    }

    // ========================================
    // 页面设置处理
    // ========================================
    if (specTextLower.indexOf('页边距') !== -1 || specTextLower.indexOf('纸张') !== -1 || specTextLower.indexOf('A4') !== -1 || specTextLower.indexOf('A3') !== -1) {
      console.log('[format] 启用页面设置');
      try {
        var pageSetupCount = 0;

        // cm转磅（精确公式：1cm = 72/2.54 磅）
        function cmToPoints(cm) {
          return cm * 72 / 2.54;
        }

        // 解析页边距（兼容多种格式）
        var topMargin = 0, bottomMargin = 0, leftMargin = 0, rightMargin = 0;

        // 格式1: 页边距上下2.5cm左右2.5cm（无空格无逗号）
        var marginMatch1 = specTextLower.match(/页边距\s*上下\s*([\d.]+)\s*(?:cm|厘米)?\s*左右\s*([\d.]+)\s*(?:cm|厘米)?/);
        if (marginMatch1) {
          topMargin = cmToPoints(parseFloat(marginMatch1[1]));
          bottomMargin = topMargin;
          leftMargin = cmToPoints(parseFloat(marginMatch1[2]));
          rightMargin = leftMargin;
          console.log('[format] 页边距解析(格式1): 上下=' + marginMatch1[1] + ', 左右=' + marginMatch1[2]);
        }

        // 格式2: 页边距上下 2.5cm，左右 2.5cm（有空格有逗号）
        if (!marginMatch1) {
          var marginMatch2 = specTextLower.match(/页边距[：:]?\s*上下\s*([\d.]+)\s*(?:cm|厘米)\s*[，,、]\s*左右\s*([\d.]+)\s*(?:cm|厘米)/);
          if (marginMatch2) {
            topMargin = cmToPoints(parseFloat(marginMatch2[1]));
            bottomMargin = topMargin;
            leftMargin = cmToPoints(parseFloat(marginMatch2[2]));
            rightMargin = leftMargin;
            console.log('[format] 页边距解析(格式2): 上下=' + marginMatch2[1] + ', 左右=' + marginMatch2[2]);
          }
        }

        // 格式3: 页边距2.5cm（四边相同）
        if (!marginMatch1 && !marginMatch2) {
          var marginMatch3 = specTextLower.match(/页边距\s*([\d.]+)\s*(?:cm|厘米)/);
          if (marginMatch3) {
            topMargin = cmToPoints(parseFloat(marginMatch3[1]));
            bottomMargin = topMargin;
            leftMargin = topMargin;
            rightMargin = topMargin;
            console.log('[format] 页边距解析(格式3): 四边=' + marginMatch3[1]);
          }
        }

        // 解析纸张大小
        var paperWidth = 0, paperHeight = 0;
        if (specTextLower.indexOf('a4') !== -1) {
          paperWidth = 595.35;   // A4宽度 210mm ≈ 595.35磅
          paperHeight = 841.95;  // A4高度 297mm ≈ 841.95磅
          console.log('[format] 纸张: A4');
        } else if (specTextLower.indexOf('a3') !== -1) {
          paperWidth = 841.95;   // A3宽度 297mm
          paperHeight = 1190.7;  // A3高度 420mm
          console.log('[format] 纸张: A3');
        }

        // 解析纸张方向
        var orientation = 0;  // 0=纵向, 1=横向
        if (specTextLower.indexOf('横向') !== -1) {
          orientation = 1;
          // 横向时交换宽高
          var temp = paperWidth;
          paperWidth = paperHeight;
          paperHeight = temp;
          console.log('[format] 方向: 横向');
        } else {
          console.log('[format] 方向: 纵向');
        }

        // 解析页眉页脚距边界
        var headerDist = 0, footerDist = 0;
        var hfDistMatch = specTextLower.match(/页眉页脚距边界\s*([\d.]+)\s*(?:cm|厘米)/);
        if (hfDistMatch) {
          headerDist = cmToPoints(parseFloat(hfDistMatch[1]));
          footerDist = headerDist;
          console.log('[format] 页眉页脚距边界: ' + hfDistMatch[1] + 'cm');
        }

        // 应用到所有节
        var sections = doc.Sections;
        if (sections && sections.Count > 0) {
          for (var si = 1; si <= sections.Count; si++) {
            var sec = sections.Item(si);
            if (!sec || !sec.PageSetup) continue;

            try {
              var ps = sec.PageSetup;
              if (topMargin > 0) ps.TopMargin = topMargin;
              if (bottomMargin > 0) ps.BottomMargin = bottomMargin;
              if (leftMargin > 0) ps.LeftMargin = leftMargin;
              if (rightMargin > 0) ps.RightMargin = rightMargin;
              if (paperWidth > 0) ps.PageWidth = paperWidth;
              if (paperHeight > 0) ps.PageHeight = paperHeight;
              ps.Orientation = orientation;
              if (headerDist > 0) ps.HeaderDistance = headerDist;
              if (footerDist > 0) ps.FooterDistance = footerDist;
              pageSetupCount++;
            } catch (e) {}
          }
        }

        console.log('[format] 页面设置: ' + pageSetupCount + '个节');
        applied += pageSetupCount;

      } catch (e) {
        console.log('[format] 页面设置失败: ' + e);
      }
    }

    // ========================================
    // 页码格式处理
    // ========================================
    if (specTextLower.indexOf('页码') !== -1) {
      console.log('[format] 启用页码设置');
      try {
        var pageNumCount = 0;

        // 页码格式常量
        var wdFieldPage = 33;        // PAGE域
        var wdAlignPageNumberCenter = 1;  // 居中
        var wdAlignPageNumberLeft = 0;    // 左对齐
        var wdAlignPageNumberRight = 2;   // 右对齐

        // 页码编号格式常量
        var wdNumberStyleArabic = 0;      // 阿拉伯数字 (1, 2, 3)
        var wdNumberStyleArabicLZ = 1;    // 带前导零阿拉伯 (01, 02, 03)
        var wdNumberStyleRomanLower = 2;  // 小写罗马 (i, ii, iii)
        var wdNumberStyleRomanUpper = 3;  // 大写罗马 (I, II, III)
        var wdNumberStyleChinese = 4;     // 中文数字 (一, 二, 三)

        // 解析页码位置
        var pageNumAlign = wdAlignPageNumberCenter;  // 默认居中
        if (specTextLower.indexOf('页码左对齐') !== -1 || specTextLower.indexOf('页码左侧') !== -1) {
          pageNumAlign = wdAlignPageNumberLeft;
        } else if (specTextLower.indexOf('页码右对齐') !== -1 || specTextLower.indexOf('页码右侧') !== -1) {
          pageNumAlign = wdAlignPageNumberRight;
        }

        // 解析页码格式
        var pageNumStyle = wdNumberStyleArabic;  // 默认阿拉伯
        if (specTextLower.indexOf('页码罗马数字') !== -1 || specTextLower.indexOf('页码大写罗马') !== -1) {
          pageNumStyle = wdNumberStyleRomanUpper;
        } else if (specTextLower.indexOf('页码小写罗马') !== -1) {
          pageNumStyle = wdNumberStyleRomanLower;
        } else if (specTextLower.indexOf('页码中文数字') !== -1 || specTextLower.indexOf('页码汉字') !== -1) {
          pageNumStyle = wdNumberStyleChinese;
        }

        // 解析起始页码
        var startPageNum = 0;
        var startMatch = specTextLower.match(/起始页码\s*(\d+)/);
        if (startMatch) {
          startPageNum = parseInt(startMatch[1]);
        }

        console.log('[format] 页码设置: 对齐=' + pageNumAlign + ', 格式=' + pageNumStyle + ', 起始=' + startPageNum);

        // 应用到所有节
        var sections = doc.Sections;
        if (sections && sections.Count > 0) {
          for (var si = 1; si <= sections.Count; si++) {
            var sec = sections.Item(si);
            if (!sec) continue;

            try {
              // 设置起始页码
              if (startPageNum > 0 && sec.PageSetup) {
                sec.PageSetup.StartingPageNumber = startPageNum;
              }

              // 设置页码格式和对齐位置
              // 注意：WPS JS API中页码通过PageNumbers对象设置
              var footers = sec.Footers;
              if (footers) {
                // 清除现有页码（主页脚）
                var primaryFooter = footers.Item(1);  // wdHeaderFooterPrimary
                if (primaryFooter) {
                  // 先清除页脚中的旧内容
                  try {
                    if (primaryFooter.Range) {
                      primaryFooter.Range.Delete();
                    }
                  } catch (e) {}

                  // 添加页码域
                  try {
                    var pn = primaryFooter.PageNumbers;
                    if (pn) {
                      pn.RestartNumberingAtSection = (startPageNum > 0);
                      pn.NumberStyle = pageNumStyle;
                      pn.Add(pageNumAlign);
                    }
                  } catch (e) {}

                  // 设置页码字体（如果指定了页眉页脚字体）
                  var hfFontSize = 9;
                  var hfFontCN = '宋体';
                  var hfFontEN = 'Arial';

                  var hfFontMatch = specTextLower.match(/页眉页脚[字号]?[：:]?\s*小?[一二三四五]号/);
                  if (hfFontMatch) {
                    hfFontSize = parseFontSize(hfFontMatch[0].replace(/页眉页脚[字号]?[：:]?\s*/, ''));
                  }
                  var cnFontMatch = specTextLower.match(/中文字体[为是]\s*([宋黑楷仿][体])/);
                  if (cnFontMatch) hfFontCN = cnFontMatch[1];
                  var enFontMatch = specTextLower.match(/西文字体[为是]\s*([A-Za-z\s]+)/);
                  if (enFontMatch) hfFontEN = enFontMatch[1].trim();

                  try {
                    if (primaryFooter.Range && primaryFooter.Range.Font) {
                      primaryFooter.Range.Font.NameFarEast = hfFontCN;
                      primaryFooter.Range.Font.Name = hfFontEN;
                      primaryFooter.Range.Font.Size = hfFontSize;
                    }
                  } catch (e) {}
                }
              }
              pageNumCount++;
            } catch (e) {}
          }
        }

        console.log('[format] 页码处理: ' + pageNumCount + '个节');
        applied += pageNumCount;

      } catch (e) {
        console.log('[format] 页码处理失败: ' + e);
      }
    }

  } finally {
    try { doc.TrackRevisions = origTrack; } catch (e) {}
  }

  var elapsed = Date.now() - startTime;

  // 只统计实际处理的类型
  var processedCounts = {};
  for (var t in typeIndices) {
    if (typeIndices[t] && typeIndices[t].length > 0) {
      processedCounts[t] = typeIndices[t].length;
    }
  }

  console.log('[format] 完成: applied=' + applied + ' errors=' + errors + ' time=' + elapsed + 'ms');

  return {
    success: true,
    applied: applied,
    errors: errors,
    elapsedMs: elapsed,
    processedTypes: processedCounts,
    total: paraCount
  };

} catch (e) {
  console.error('[format] 错误: ' + e);
  return { success: false, error: String(e) };
}