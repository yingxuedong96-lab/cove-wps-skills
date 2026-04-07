/**
 * validate-config.js - 配置校验与智能修正
 *
 * 用于修正不同LLM模型可能产生的解析差异，确保配置一致性
 */

// 标准对齐方式映射（防止LLM解析错误）
var ALIGNMENT_MAP = {
  '左对齐': 0, '靠左': 0, '左': 0,
  '居中': 1, '居中对齐': 1, '居中排列': 1,
  '右对齐': 2, '靠右': 2, '右': 2,
  '两端对齐': 3, '分散对齐': 3, '两端': 3, ' justify': 3
};

// 段落类型的默认语义规则（用于检测不合理配置）
var DEFAULT_SEMANTICS = {
  zhangTitle: { alignment: 1, bold: true },        // 章标题通常居中加粗
  heading2: { alignment: 0, bold: true },          // 二级标题通常左对齐加粗
  heading3: { alignment: 0 },                       // 三级标题通常左对齐
  heading4: { alignment: 0 },                       // 四级标题通常左对齐
  body: { alignment: 3, bold: false },             // 正文通常两端对齐不加粗
  tableCaption: { alignment: 1, bold: false },     // 表名通常居中不加粗
  figureCaption: { alignment: 1, bold: false },    // 图名通常居中不加粗
  reference: { alignment: 0, bold: false }         // 参考文献通常左对齐不加粗
};

// 智能修正函数
function validateAndFix(config) {
  if (!config) return config;

  console.log('[validate] 开始配置校验与修正...');

  var rules = config.paragraphRules || {};
  var patterns = config.numberingPatterns || {};
  var warnings = [];
  var fixes = [];

  // 0. 修正编号正则中的常见错误
  var patternFixes = {
    // "第 章" 中间多余空格
    '\\^第\\s+\\[': '^第[',
    // 其他常见错误
    '\\^第\\s+第': '^第第'
  };

  for (var key in patterns) {
    if (Array.isArray(patterns[key])) {
      for (var i = 0; i < patterns[key].length; i++) {
        var p = patterns[key][i];
        var origP = p;
        // 修正常见正则错误
        p = p.replace(/^第\s+\[/, '^第[');  // "第 [" -> "第["
        p = p.replace(/^第\s+第/, '^第');    // "第 第" -> "第"
        if (p !== origP) {
          patterns[key][i] = p;
          fixes.push('numberingPatterns.' + key + '[' + i + ']: 修正正则空格');
        }
      }
    }
  }

  // 1. 字段名规范化
  for (var key in rules) {
    var rule = rules[key];

    // lineSpacingType -> lineSpacingRule
    if (rule.lineSpacingType) {
      if (rule.lineSpacingType === 'fixed' || rule.lineSpacingType === 'Fixed') {
        if (rule.lineSpacingRule === undefined) {
          rule.lineSpacingRule = 4;
          fixes.push(key + '.lineSpacingType -> lineSpacingRule=4');
        }
      }
      delete rule.lineSpacingType; // 删除非标准字段
    }

    // lineSpacing 数值转 lineSpacingRule（LLM 常见错误）
    // 1.5 -> lineSpacingRule=1, 2 -> lineSpacingRule=2, 1 -> lineSpacingRule=0
    if (rule.lineSpacing !== undefined && rule.lineSpacingRule === undefined) {
      var ls = rule.lineSpacing;
      if (ls === 1 || ls === '1' || ls === 'single') {
        rule.lineSpacingRule = 0;  // 单倍行距
        delete rule.lineSpacing;
        fixes.push(key + '.lineSpacing=1 -> lineSpacingRule=0 (单倍行距)');
      } else if (ls === 1.5 || ls === '1.5') {
        rule.lineSpacingRule = 1;  // 1.5倍行距
        delete rule.lineSpacing;
        fixes.push(key + '.lineSpacing=1.5 -> lineSpacingRule=1 (1.5倍行距)');
      } else if (ls === 2 || ls === '2' || ls === 'double') {
        rule.lineSpacingRule = 2;  // 2倍行距
        delete rule.lineSpacing;
        fixes.push(key + '.lineSpacing=2 -> lineSpacingRule=2 (2倍行距)');
      } else if (typeof ls === 'number' && ls > 2 && ls < 100) {
        // 可能是倍数，如 1.25, 1.15 等
        // WPS 不支持任意倍数，需要转为固定值或忽略
        warnings.push(key + '.lineSpacing=' + ls + ' 不支持，请使用 lineSpacingRule (0=单倍,1=1.5倍,2=2倍) 或 lineSpacing+lineSpacingRule=4 (固定值)');
        delete rule.lineSpacing;
      }
      // 大于 100 的视为固定磅值，需要设置 lineSpacingRule=4
      if (typeof ls === 'number' && ls >= 100) {
        rule.lineSpacingRule = 4;
        fixes.push(key + '.lineSpacing=' + ls + ' -> lineSpacingRule=4 (固定行距)');
      }
    }

    // firstLineIndent 字符转磅值
    if (rule.firstLineIndent !== undefined && rule.firstLineIndent < 10 && rule.firstLineIndent > 0) {
      var oldVal = rule.firstLineIndent;
      rule.firstLineIndent = oldVal * 12;
      fixes.push(key + '.firstLineIndent: ' + oldVal + '字符 -> ' + rule.firstLineIndent + '磅');
    }
  }

  // 2. 语义合理性检查
  for (var key in DEFAULT_SEMANTICS) {
    if (rules[key]) {
      var rule = rules[key];
      var semantics = DEFAULT_SEMANTICS[key];

      // 检查对齐方式是否合理
      if (semantics.alignment !== undefined && rule.alignment !== undefined) {
        // 如果配置与默认语义不同，记录警告（不强制修正，因为可能是用户故意指定）
        if (rule.alignment !== semantics.alignment) {
          warnings.push(key + '.alignment=' + rule.alignment + ' (默认语义: ' + semantics.alignment + ')');
        }
      }

      // 检查加粗是否合理
      if (semantics.bold !== undefined && rule.bold !== undefined) {
        if (rule.bold !== semantics.bold) {
          warnings.push(key + '.bold=' + rule.bold + ' (默认语义: ' + semantics.bold + ')');
        }
      }
    }
  }

  // 3. 页面设置格式转换
  if (config.pageSetup) {
    var ps = config.pageSetup;
    // 转换为标准documentInfo格式
    if (!config.documentInfo) {
      config.documentInfo = { pageMargins: {} };
    }
    if (!config.documentInfo.pageMargins) {
      config.documentInfo.pageMargins = {};
    }

    // 厘米转磅
    var cm2pt = 28.35;
    if (ps.marginTop !== undefined) {
      config.documentInfo.pageMargins.top = ps.marginTop * cm2pt;
      fixes.push('pageSetup.marginTop -> documentInfo.pageMargins.top');
    }
    if (ps.marginBottom !== undefined) {
      config.documentInfo.pageMargins.bottom = ps.marginBottom * cm2pt;
      fixes.push('pageSetup.marginBottom -> documentInfo.pageMargins.bottom');
    }
    if (ps.marginLeft !== undefined) {
      config.documentInfo.pageMargins.left = ps.marginLeft * cm2pt;
      fixes.push('pageSetup.marginLeft -> documentInfo.pageMargins.left');
    }
    if (ps.marginRight !== undefined) {
      config.documentInfo.pageMargins.right = ps.marginRight * cm2pt;
      fixes.push('pageSetup.marginRight -> documentInfo.pageMargins.right');
    }
    if (ps.headerMargin !== undefined) {
      config.documentInfo.pageMargins.headerDistance = ps.headerMargin * cm2pt;
    }
    if (ps.footerMargin !== undefined) {
      config.documentInfo.pageMargins.footerDistance = ps.footerMargin * cm2pt;
    }
  }

  // 输出修正报告
  if (fixes.length > 0) {
    console.log('[validate] 自动修正 ' + fixes.length + ' 项:');
    for (var i = 0; i < fixes.length; i++) {
      console.log('[validate]   - ' + fixes[i]);
    }
  }
  if (warnings.length > 0) {
    console.log('[validate] 语义警告 ' + warnings.length + ' 项:');
    for (var i = 0; i < warnings.length; i++) {
      console.log('[validate]   - ' + warnings[i]);
    }
  }
  if (fixes.length === 0 && warnings.length === 0) {
    console.log('[validate] 配置校验通过，无需修正');
  }

  return config;
}

// 导出函数（WPS环境）
if (typeof config !== 'undefined') {
  var configData = typeof config === 'string' ? JSON.parse(config) : config;
  var validated = validateAndFix(configData);
  config = JSON.stringify(validated);
}