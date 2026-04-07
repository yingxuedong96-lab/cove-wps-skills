/**
 * add-separators.js
 * 添加 GB/T 9704-2012 公文分隔线
 *
 * 支持四种公文格式：
 * - 通用公文：红色单线
 * - 信函格式：红色双线（上粗下细）
 * - 命令(令)格式：无红色分隔线
 * - 纪要格式：无红色分隔线
 *
 * 版记分隔线：黑色，首条/末条粗线(0.35mm)，中间细线(0.25mm)
 *
 * 出参: { success: boolean, formatType: string, separators: object, message: string }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, formatType: 'general', separators: {}, message: '未找到活动文档' };
  }

  // ========================================
  // 常量定义
  // ========================================
  var MM_TO_POINTS = 2.835;
  var CONTENT_WIDTH = 156 * MM_TO_POINTS;  // 版心宽度

  // 分隔线颜色（WPS 使用 BGR 格式）
  var RED = 0x0000FF;    // BGR 格式的红色
  var BLACK = 0x000000;  // 黑色

  // 线条粗细（sz值，单位是 1/8 磅）
  // 国标要求：粗线 0.35mm ≈ 1pt，细线 0.25mm ≈ 0.7pt
  // 实测：sz=12 ≈ 1.5pt ≈ 0.53mm（粗线），sz=8 ≈ 1pt ≈ 0.35mm（细线）
  var THICK_LINE = 12;  // sz=12 ≈ 1.5pt，用于首条/末条分隔线
  var THIN_LINE = 8;    // sz=8 ≈ 1pt，用于中间分隔线

  // 格式类型
  var FORMAT_TYPES = {
    GENERAL: 'general',
    LETTER: 'letter',
    ORDER: 'order',
    MINUTES: 'minutes'
  };

  // ========================================
  // 辅助函数
  // ========================================

  function getParaText(para) {
    try {
      var text = para.Range.Text;
      return text ? text.replace(/[\r\n]/g, '').trim() : '';
    } catch (e) {
      return '';
    }
  }

  function isDocNumber(text) {
    return /〔\d{4}〕\d+号/.test(text) || /\[\d{4}\]\d+号/.test(text);
  }

  function isChaosong(text) {
    return /^抄送[：:]/.test(text);
  }

  function isYinfa(text) {
    return /印发\s*$/.test(text);
  }

  // ========================================
  // 格式类型检测
  // ========================================
  function detectFormatType() {
    var totalParas = doc.Paragraphs.Count;

    for (var i = 1; i <= Math.min(10, totalParas); i++) {
      var para = doc.Paragraphs.Item(i);
      var text = getParaText(para);
      var alignment = para.Format.Alignment;

      if (!text) continue;

      // 命令(令)格式：发文机关标志含"命令"或"令"
      // 注意：不要求居中，原文档可能格式不规范
      if (/命令|令$/.test(text)) {
        return FORMAT_TYPES.ORDER;
      }

      // 纪要格式：发文机关标志含"纪要"
      // 注意：不要求居中，原文档可能格式不规范
      if (/纪要/.test(text)) {
        return FORMAT_TYPES.MINUTES;
      }

      // 信函格式检测（多种方式）
      // 方式1：发文字号右对齐
      if (/〔\d{4}〕\d+号/.test(text) || /\[\d{4}\]\d+号/.test(text)) {
        if (alignment === 2) {
          return FORMAT_TYPES.LETTER;
        }
      }

      // 方式2：发文机关标志不含"文件"且后续标题含"函"
      if (alignment === 1 && !/文件$/.test(text)) {
        if (/政府$|办公室$|委员会$|厅$|局$|委$/.test(text)) {
          for (var j = i + 1; j <= Math.min(20, totalParas); j++) {
            var nextText = getParaText(doc.Paragraphs.Item(j));
            if (nextText && /函$|复函$/.test(nextText)) {
              return FORMAT_TYPES.LETTER;
            }
          }
        }
      }

      // 方式3：标题含"函"且回溯发文机关标志不含"文件"
      if (text.length > 5 && text.length <= 50 && /函$|复函$/.test(text)) {
        for (var k = 1; k < i; k++) {
          var prevText = getParaText(doc.Paragraphs.Item(k));
          if (prevText && /政府$|办公室$|委员会$|厅$|局$|委$/.test(prevText) && !/文件$/.test(prevText)) {
            return FORMAT_TYPES.LETTER;
          }
        }
      }
    }

    return FORMAT_TYPES.GENERAL;
  }

  // ========================================
  // 添加段落分隔线方法
  // ========================================

  /**
   * 为段落添加下边框
   */
  function addBottomBorder(para, color, lineWidth) {
    try {
      var borders = para.Format.Borders;
      if (borders) {
        var bottomBorder = borders.Item(-3);  // wdBorderBottom = -3
        bottomBorder.LineStyle = 1;           // wdLineStyleSingle
        bottomBorder.LineWidth = lineWidth || THIN_LINE;
        bottomBorder.Color = color;
      }
    } catch (e) {
      console.warn('[add-separators] 添加下边框失败', e);
    }
  }

  /**
   * 为段落添加上边框
   */
  function addTopBorder(para, color, lineWidth) {
    try {
      var borders = para.Format.Borders;
      if (borders) {
        var topBorder = borders.Item(-1);  // wdBorderTop = -1
        topBorder.LineStyle = 1;           // wdLineStyleSingle
        topBorder.LineWidth = lineWidth || THIN_LINE;
        topBorder.Color = color;
      }
    } catch (e) {
      console.warn('[add-separators] 添加上边框失败', e);
    }
  }

  /**
   * 添加上红色双线（发文机关标志下）
   * 规范：上粗下细双线
   *
   * WPS LineStyle常量测试：
   * - 1 = Single (单线)
   * - 5 = Double (双线等宽)
   * - 尝试找到上粗下细的正确值
   */
  function addTopDoubleLine(para) {
    try {
      var borders = para.Format.Borders;
      if (borders) {
        var bottomBorder = borders.Item(-3);  // wdBorderBottom

        // 尝试多种LineStyle值
        // 根据WPS/Word API，双线样式需要测试
        // 5 = Double, 可能的值还有其他
        bottomBorder.LineStyle = 5;           // Double双线
        bottomBorder.LineWidth = THICK_LINE;
        bottomBorder.Color = RED;
      }
    } catch (e) {
      console.warn('[add-separators] 添加上红色双线失败', e);
    }
  }

  /**
   * 添加下红色双线（距下页边20mm）
   * 规范：上细下粗双线
   */
  function addBottomDoubleLine(para) {
    try {
      var borders = para.Format.Borders;
      if (borders) {
        var topBorder = borders.Item(-1);  // wdBorderTop
        topBorder.LineStyle = 5;           // Double双线
        topBorder.LineWidth = THICK_LINE;
        topBorder.Color = RED;
      }
    } catch (e) {
      console.warn('[add-separators] 添加下红色双线失败', e);
    }
  }

  /**
   * 添加双线分隔线（信函格式专用，兼容旧版本）
   */
  function addDoubleLine(para) {
    addTopDoubleLine(para);
  }

  // ========================================
  // 主流程
  // ========================================
  var formatType = detectFormatType();
  var totalParas = doc.Paragraphs.Count;
  var separators = {
    header: { added: false, position: -1, type: 'none' },
    footerFirst: { added: false, position: -1 },
    footerMiddle: { added: false, position: -1 },
    footerLast: { added: false, position: -1 }
  };

  // 查找关键位置
  var documentFlagIndex = -1;  // 发文机关标志位置
  var docNumberIndex = -1;
  var chaosongIndex = -1;
  var yinfaIndex = -1;
  var lastParaIndex = -1;

  for (var i = 1; i <= totalParas; i++) {
    var para = doc.Paragraphs.Item(i);
    var text = getParaText(para);

    if (!text) continue;

    lastParaIndex = i;

    // 发文机关标志：居中且含机关名
    var alignment = para.Format.Alignment;
    if (alignment === 1 && /政府$|办公室$|委员会$|厅$|局$|委$/.test(text)) {
      if (documentFlagIndex < 0) {
        documentFlagIndex = i;
      }
    }

    if (isDocNumber(text)) {
      docNumberIndex = i;
    }

    if (isChaosong(text)) {
      chaosongIndex = i;
    }

    if (isYinfa(text)) {
      yinfaIndex = i;
    }
  }

  // ========================================
  // 添加版头分隔线（根据格式类型）
  // ========================================
  if (formatType === FORMAT_TYPES.LETTER && documentFlagIndex > 0) {
    // 信函格式：使用Shape绘制红色双线（文武线）
    // 国标10.1：发文机关标志上边缘至上页边30mm，下4mm处印红色双线（上粗下细）
    // 注：WPS段落边框LineStyle=5实际显示为点划线，需改用Shape绘制
    try {
      var flagPara = doc.Paragraphs.Item(documentFlagIndex);
      var flagRange = flagPara.Range;

      // 双线参数
      var lineWidth = 170 * MM_TO_POINTS;  // 170mm长

      // 第一条红色双线位置计算：
      // 国标10.1：发文机关标志上边缘至上页边为30mm
      // 发文机关标志高度：2号字(22磅)约8mm + 行距约2mm ≈ 10mm
      // 发文机关标志下边缘 ≈ 30 + 10 = 40mm
      // 第一条双线位置 = 发文机关标志下边缘 + 4mm ≈ 44mm
      var topLineY = 44 * MM_TO_POINTS;  // 发文机关标志下4mm ≈ 44mm

      // 线条居中：A4宽度210mm居中点
      var centerX = (210 * MM_TO_POINTS) / 2;
      var lineStartX = centerX - (lineWidth / 2);

      // 使用Shape绘制双线（双线=两条紧挨着的平行线）
      var shapes = doc.Shapes;
      if (shapes) {
        // 国标"双线"：两条线紧挨着，间距约2-3pt形成清晰双线效果
        // 粗线1.5pt约0.53mm，细线0.75pt约0.26mm
        // 上粗下细：上线粗(1.5pt)，下线细(0.75pt)
        var gapPoints = 2.5;  // 约0.88mm间距，确保双线清晰可见

        // 上线（粗线）- 上粗下细双线的上面那条
        var line1 = shapes.AddLine(lineStartX, topLineY, lineStartX + lineWidth, topLineY, flagRange);
        line1.Line.ForeColor.RGB = RED;
        line1.Line.Weight = 1.5;  // 粗线约0.53mm
        line1.WrapFormat.Type = 3;  // wdWrapNone
        // 清除阴影（国标无阴影要求）
        try { line1.Shadow.Visible = false; } catch (e) {}
        // 设置相对于页面定位
        // WPS API常量：wdRelativeVerticalPositionPage = 1, wdRelativeHorizontalPositionPage = 1
        try {
          line1.RelativeVerticalPosition = 1;  // wdRelativeVerticalPositionPage = 1（相对于页面）
          line1.Top = topLineY;
          line1.RelativeHorizontalPosition = 1;  // wdRelativeHorizontalPositionPage = 1
          line1.Left = lineStartX;
        } catch (e) {}

        // 下线（细线）- 上粗下细双线的下面那条，在上线下方
        var line2 = shapes.AddLine(lineStartX, topLineY + gapPoints, lineStartX + lineWidth, topLineY + gapPoints, flagRange);
        line2.Line.ForeColor.RGB = RED;
        line2.Line.Weight = 0.75;  // 细线约0.26mm
        line2.WrapFormat.Type = 3;
        // 清除阴影（国标无阴影要求）
        try { line2.Shadow.Visible = false; } catch (e) {}
        // 设置相对于页面定位
        try {
          line2.RelativeVerticalPosition = 1;  // wdRelativeVerticalPositionPage = 1（相对于页面）
          line2.Top = topLineY + gapPoints;
          line2.RelativeHorizontalPosition = 1;  // wdRelativeHorizontalPositionPage = 1
          line2.Left = lineStartX;
        } catch (e) {}

        console.log('[add-separators] 信函第一条双线已添加（页面绝对定位Y=49mm）');
      }

      // 移除段落边框（如果有）
      try {
        var borders = flagPara.Format.Borders;
        if (borders) {
          borders.Item(-3).LineStyle = 0;  // 清除下边框
        }
      } catch (e) {}

      separators.header.added = true;
      separators.header.position = documentFlagIndex;
      separators.header.type = 'double-red-top-shape';

    } catch (e) {
      console.warn('[add-separators] 添加信函上线失败', e);
    }
  } else if (docNumberIndex > 0) {
    // 其他格式：红色单线在发文字号下方
    try {
      var headerPara = doc.Paragraphs.Item(docNumberIndex);

      switch (formatType) {
        case FORMAT_TYPES.GENERAL:
          // 通用公文：红色单线
          addBottomBorder(headerPara, RED, THICK_LINE);
          separators.header.type = 'single-red';
          break;

        case FORMAT_TYPES.ORDER:
        case FORMAT_TYPES.MINUTES:
          // 命令/纪要格式：无红色分隔线
          separators.header.type = 'none';
          break;

        default:
          addBottomBorder(headerPara, RED, THICK_LINE);
          separators.header.type = 'single-red';
      }

      separators.header.added = true;
      separators.header.position = docNumberIndex;

    } catch (e) {
      console.warn('[add-separators] 添加版头分隔线失败', e);
    }
  }

  // 注：版记分隔线由 apply-format.js 中的版记表格处理

  // ========================================
  // 信函格式：添加下红色双线（距下页边20mm）
  // 国标10.1：距下页边20mm处印红色双线（上细下粗），线长170mm居中
  // 实现方案：在版记处理时添加（apply-format.js），这里只设置页脚
  // ========================================
  if (formatType === FORMAT_TYPES.LETTER) {
    try {
      var sections = doc.Sections;
      if (sections && sections.Count > 0) {
        var section = sections.Item(1);

        // 设置首页不同（首页不显示页码）
        try {
          section.PageSetup.OddAndEvenPagesHeaderFooter = true;
          section.PageSetup.DifferentFirstPageHeaderFooter = true;
        } catch (e) {
          console.warn('[add-separators] 设置首页不同失败', e);
        }

        // 清空首页页脚
        try {
          var firstFooter = section.Footers.Item(2);  // wdHeaderFooterFirstPage
          if (firstFooter) {
            var footerRange = firstFooter.Range;
            footerRange.Delete();
            footerRange.InsertAfter(" ");
            footerRange.Delete();
            try {
              footerRange.ParagraphFormat.Alignment = 1;
            } catch (e) {}
          }
        } catch (e) {
          console.warn('[add-separators] 清空首页页脚失败', e);
        }

        // 第二条红色双线将在 apply-format.js 的版记处理中添加
        // 这里设置标记，让 apply-format.js 知道需要添加双线
        separators.footerFirst.added = false;  // 标记未添加，由 apply-format.js 处理
        separators.footerFirst.type = 'pending-for-apply-format';
        console.log('[add-separators] 第二条双线将由 apply-format.js 在版记处理时添加');
      }
    } catch (e) {
      console.warn('[add-separators] 信函格式设置失败', e);
    }
  }

  // ========================================
  // 返回结果
  // ========================================
  var formatNames = {
    'general': '通用公文格式',
    'letter': '信函格式',
    'order': '命令(令)格式',
    'minutes': '纪要格式'
  };

  return {
    success: true,
    formatType: formatType,
    formatName: formatNames[formatType] || '通用公文格式',
    separators: separators,
    message: '分隔线设置完成（' + (formatNames[formatType] || '通用公文格式') + '）'
  };

} catch (e) {
  console.warn('[add-separators]', e);
  return { success: false, formatType: 'general', separators: {}, message: String(e) };
}