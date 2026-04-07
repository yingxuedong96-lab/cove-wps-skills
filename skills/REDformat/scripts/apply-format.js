/**
 * apply-format.js
 * 会议纪要红头文件排版。
 *
 * 结构：
 * - 红头：发文单位名 + "专题会议纪要"
 * - 发文字号
 * - 标题
 * - 正文（含一级/二级标题）
 * - 参会人员
 * - 版记：抄送、印发、共印
 *
 * 出参: { success: boolean, paragraphCount: number, appliedStyles: object }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, paragraphCount: 0, appliedStyles: {} };
  }

  // ========================================
  // 样式规范配置
  // ========================================
  var STYLE_RULES = {
    'hongtou': { name: '会议纪要-红头', fontCN: '方正小标宋简体', fontEN: 'Times New Roman', fontSize: 42, bold: false, alignment: 1, firstLineIndent: 0, fontColor: 255 },
    'huiyijiyao': { name: '会议纪要-标题字', fontCN: '方正小标宋简体', fontEN: 'Times New Roman', fontSize: 42, bold: false, alignment: 1, firstLineIndent: 0, fontColor: 255 },
    'fawenzihao': { name: '会议纪要-发文字号', fontCN: '仿宋_GB2312', fontEN: 'Times New Roman', fontSize: 16, bold: false, alignment: 1, firstLineIndent: 0 },
    'biaoti': { name: '会议纪要-标题', fontCN: '方正小标宋简体', fontEN: 'Times New Roman', fontSize: 22, bold: true, alignment: 1, firstLineIndent: 0 },
    'heading1': { name: '会议纪要-一级标题', fontCN: '黑体', fontEN: '黑体', fontSize: 16, bold: false, alignment: 0, firstLineIndent: 0 },
    'heading2': { name: '会议纪要-二级标题', fontCN: '楷体_GB2312', fontEN: 'Times New Roman', fontSize: 16, bold: false, alignment: 0, firstLineIndent: 0 },
    'body': { name: '会议纪要-正文', fontCN: '仿宋_GB2312', fontEN: 'Times New Roman', fontSize: 16, bold: false, alignment: 0, firstLineIndent: 32 },
    'canyuan': { name: '会议纪要-参会人员', fontCN: '仿宋_GB2312', fontEN: 'Times New Roman', fontSize: 16, bold: false, alignment: 0, firstLineIndent: 32 },
    'chaosong': { name: '会议纪要-抄送', fontCN: '仿宋_GB2312', fontEN: 'Times New Roman', fontSize: 14, bold: false, alignment: 0, firstLineIndent: 0 },
    'yinfa': { name: '会议纪要-印发', fontCN: '仿宋_GB2312', fontEN: 'Times New Roman', fontSize: 14, bold: false, alignment: 0, firstLineIndent: 0 },
    'gongyin': { name: '会议纪要-共印', fontCN: '仿宋_GB2312', fontEN: 'Times New Roman', fontSize: 14, bold: false, alignment: 0, firstLineIndent: 0 }
  };

  // ========================================
  // 识别规则
  // ========================================
  function isHuiyijiyao(text) { return text && /^专题会议纪要$/.test(text); }
  function isFawenzihao(text) { return text && /^[〔\[].*[〕\]]/.test(text); }
  function isHeading1(text) { return text && /^[一二三四五六七八九十]+、/.test(text); }
  function isHeading2(text) { return text && /^[（(][一二三四五六七八九十]+[）)]/.test(text); }
  function isCanyuan(text) { return text && /^参会人员[：:]/.test(text); }
  function isChaosong(text) { return text && /^抄送[：:]/.test(text); }
  function isYinfa(text) { return text && /印发$/.test(text); }
  function isGongyin(text) { return text && /^共印\d+份$/.test(text); }
  function isShuming(text) { return text && /政府|委员会|管委会|办公室$/.test(text) && text.length < 30; }

  // ========================================
  // 应用样式
  // ========================================
  function applyStyleRule(para, rule) {
    var range = para.Range;
    var fmt = para.Format;

    // 设置字体
    if (range && range.Font) {
      if (rule.fontCN) range.Font.NameFarEast = rule.fontCN;
      if (rule.fontEN) range.Font.Name = rule.fontEN;
      if (rule.fontSize) range.Font.Size = rule.fontSize;
      if (rule.bold) range.Font.Bold = true;
      if (rule.fontColor) range.Font.Color = rule.fontColor;
    }

    // 设置段落格式
    if (fmt) {
      if (rule.alignment !== undefined) {
        fmt.Alignment = rule.alignment;
      }
      if (rule.firstLineIndent !== undefined) {
        fmt.FirstLineIndent = rule.firstLineIndent;
      }
    }

    // 应用样式名
    if (rule.name) {
      try { para.Style = rule.name; } catch (e) {}
    }
  }

  // 插入横线
  function insertHorizontalLine(para, position, color) {
    try {
      var fmt = para.Format;
      if (fmt && fmt.Borders) {
        // WPS 边框索引: -1=top, -2=left, -3=bottom, -4=right
        var borderIndex = position === 'top' ? -1 : -3;
        fmt.Borders.Item(borderIndex).LineStyle = 1;
        fmt.Borders.Item(borderIndex).LineWidth = 12;
        fmt.Borders.Item(borderIndex).Color = color || 0;
      }
    } catch (e) {
      console.warn('[REDformat] 边框设置失败', e);
    }
  }

  // 创建版记表格
  function createBanjiTable(doc, chaosongIdx, allTexts, appliedStyles) {
    if (chaosongIdx < 0) return;

    try {
      // 找到抄送、印发、共印的内容
      var chaosongText = '';
      var yinfaText = '';
      var gongyinText = '';

      for (var i = chaosongIdx; i < allTexts.length; i++) {
        var t = allTexts[i];
        if (isChaosong(t)) {
          chaosongText = t;
        } else if (isYinfa(t)) {
          yinfaText = t;
        } else if (isGongyin(t)) {
          gongyinText = t;
        }
      }

      if (!chaosongText) return;

      // 解析印发机关和日期
      var yinfaOrg = '';
      var yinfaDate = '';
      if (yinfaText) {
        var match = yinfaText.match(/^(.+?)\s+(\d{4}年\d+月\d+日)印发$/);
        if (match) {
          yinfaOrg = match[1];
          yinfaDate = match[2];
        } else {
          yinfaOrg = yinfaText.replace(/印发$/, '').trim();
        }
      }

      // 先获取抄送段落的Range，然后再删除
      var chaosongPara = doc.Paragraphs.Item(chaosongIdx + 1);
      var insertRange = chaosongPara.Range;

      // 从后往前删除版记段落
      var parasToDelete = [];
      for (var i = chaosongIdx; i < allTexts.length; i++) {
        var t = allTexts[i];
        if (isChaosong(t) || isYinfa(t) || isGongyin(t)) {
          parasToDelete.push(i + 1);
        }
      }
      parasToDelete.sort(function(a, b) { return b - a; });
      for (var j = 0; j < parasToDelete.length; j++) {
        doc.Paragraphs.Item(parasToDelete[j]).Range.Delete();
      }

      // 计算表格宽度
      var tableWidth = 595.3 - 85 - 85;

      // 创建表格：4行2列
      var table = doc.Tables.Add(insertRange, 4, 2);
      table.PreferredWidth = tableWidth;

      // 设置列宽
      table.Columns.Item(1).Width = tableWidth * 0.6;
      table.Columns.Item(2).Width = tableWidth * 0.4;

      // ===== 第一步：填入所有内容 =====

      // 行1：空行
      table.Cell(1, 1).Range.Text = '';
      table.Cell(1, 2).Range.Text = '';

      // 行2：左侧抄送，右侧空
      table.Cell(2, 1).Range.Text = chaosongText;
      table.Cell(2, 1).Range.Font.NameFarEast = '仿宋';
      table.Cell(2, 1).Range.Font.Name = '仿宋';
      table.Cell(2, 1).Range.Font.Size = 14;
      table.Cell(2, 2).Range.Text = '';

      // 行3：印发机关（左）+ 日期印发（右）
      table.Cell(3, 1).Range.Text = yinfaOrg;
      table.Cell(3, 1).Range.Font.NameFarEast = '仿宋';
      table.Cell(3, 1).Range.Font.Name = '仿宋';
      table.Cell(3, 1).Range.Font.Size = 14;

      table.Cell(3, 2).Range.Text = yinfaDate + '印发';
      table.Cell(3, 2).Range.Font.NameFarEast = '仿宋';
      table.Cell(3, 2).Range.Font.Name = '仿宋';
      table.Cell(3, 2).Range.Font.Size = 14;
      table.Cell(3, 2).Range.ParagraphFormat.Alignment = 2;

      // 行4：左侧空，右侧共印（右对齐）
      table.Cell(4, 1).Range.Text = '';
      table.Cell(4, 2).Range.Text = gongyinText;
      table.Cell(4, 2).Range.Font.NameFarEast = '仿宋';
      table.Cell(4, 2).Range.Font.Name = '仿宋';
      table.Cell(4, 2).Range.Font.Size = 14;
      table.Cell(4, 2).Range.ParagraphFormat.Alignment = 2;

      // ===== 第二步：合并单元格 =====
      table.Cell(2, 1).Merge(table.Cell(2, 2));
      table.Cell(4, 1).Merge(table.Cell(4, 2));

      // ===== 第三步：设置边框 =====
      table.Borders.Enable = true;

      // 取消左边框、右边框、内部垂直边框
      table.Borders.Item(-2).LineStyle = 0;  // left
      table.Borders.Item(-4).LineStyle = 0;  // right
      table.Borders.Item(-6).LineStyle = 0;  // vertical

      // 设置水平边框样式
      table.Borders.Item(-5).LineStyle = 1;
      table.Borders.Item(-5).LineWidth = 4;
      table.Borders.Item(-5).Color = 0;

      // 第一行无边框
      table.Rows.Item(1).Borders.Item(-1).LineStyle = 0;

      // 最后一行无边框
      table.Rows.Item(4).Borders.Item(-3).LineStyle = 0;

      // ===== 第四步：设置单元格内容底部对齐 =====
      for (var row = 1; row <= 4; row++) {
        for (var col = 1; col <= 2; col++) {
          try {
            var cell = table.Cell(row, col);
            cell.VerticalAlignment = 3;  // wdCellAlignVerticalBottom
          } catch (e) {}
        }
      }

      // ===== 第五步：设置表格环绕定位到底部 =====
      try {
        var rows = table.Rows;

        // 开启环绕
        rows.WrapAroundText = true;

        // 水平居中（相对于页边距）
        rows.RelativeHorizontalPosition = 0;  // wdRelativeHorizontalPositionMargin
        rows.HorizontalPosition = -999995;    // wdTableCenter

        // 垂直底端（相对于页边距底部区域）
        rows.RelativeVerticalPosition = 5;    // wdRelativeVerticalPositionBottomMarginArea
        rows.VerticalPosition = -999997;      // wdTableBottom

        console.log('[REDformat] 表格环绕定位设置成功');
      } catch (e) {
        console.warn('[REDformat] 表格环绕定位设置失败', e);
      }

      appliedStyles['banji_table'] = 1;
      console.log('[REDformat] 版记表格创建成功');

    } catch (e) {
      console.warn('[REDformat] 创建版记表格失败', e);
    }
  }

  // ========================================
  // 页面格式
  // ========================================
  try {
    var ps = doc.PageSetup;
    ps.PageWidth = 595.3;
    ps.PageHeight = 841.9;
    ps.TopMargin = 113;
    ps.BottomMargin = 113;
    ps.LeftMargin = 85;
    ps.RightMargin = 85;
  } catch (e) {}

  // ========================================
  // 主流程
  // ========================================
  var appliedStyles = {};
  var totalParas = doc.Paragraphs.Count;

  // 收集所有段落
  var allTexts = [];
  for (var i = 1; i <= totalParas; i++) {
    var para = doc.Paragraphs.Item(i);
    var text = '';
    try { text = para.Range.Text; if (text) text = text.replace(/[\r\n]/g, '').trim(); } catch (e) {}
    allTexts.push(text);
  }

  // 找关键位置
  var titleIndex = -1;
  var chaosongIndex = -1;

  for (var i = 0; i < allTexts.length; i++) {
    var text = allTexts[i];
    // 标题：发文字号之后的段落，包含"关于"
    if (titleIndex === -1) {
      if (i > 0 && isFawenzihao(allTexts[i - 1])) {
        if (text && /关于|会议纪要$/.test(text)) {
          titleIndex = i;
        }
      }
    }
    if (isChaosong(text)) chaosongIndex = i;
  }

  console.log('[REDformat] title=' + titleIndex + ', chaosong=' + chaosongIndex);

  // 遍历并应用样式
  for (var i = 1; i <= totalParas; i++) {
    var para = doc.Paragraphs.Item(i);
    var text = allTexts[i - 1];
    if (!text) continue;

    var idx = i - 1;
    var elementType = 'body';

    // ===== 版头部分 =====
    // 第一段是红头（发文单位名）
    if (idx === 0) {
      elementType = 'hongtou';
    }
    else if (isHuiyijiyao(text)) {
      elementType = 'huiyijiyao';
    }
    else if (isFawenzihao(text)) {
      elementType = 'fawenzihao';
      insertHorizontalLine(para, 'bottom', 255);  // 发文字号下方红色横线
    }
    // ===== 版记部分 =====
    // 跳过版记段落，它们会被表格替换
    else if (isChaosong(text) || isYinfa(text) || isGongyin(text)) {
      continue;
    }
    else if (isCanyuan(text)) elementType = 'canyuan';
    // ===== 主体部分 =====
    else if (idx === titleIndex) elementType = 'biaoti';
    else if (isHeading1(text)) elementType = 'heading1';
    else if (isHeading2(text)) elementType = 'heading2';

    var rule = STYLE_RULES[elementType] || STYLE_RULES['body'];
    applyStyleRule(para, rule);
    appliedStyles[elementType] = (appliedStyles[elementType] || 0) + 1;
  }

  // ===== 创建版记表格 =====
  createBanjiTable(doc, chaosongIndex, allTexts, appliedStyles);

  return { success: true, paragraphCount: totalParas, appliedStyles: appliedStyles };

} catch (e) {
  console.warn('[apply-format]', e);
  return { success: false, paragraphCount: 0, appliedStyles: {}, error: String(e) };
}