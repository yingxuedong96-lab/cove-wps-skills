/**
 * apply-format.js
 * 应用 GB/T 9704-2012 国标公文排版格式
 *
 * 支持四种公文格式：通用公文、信函格式、命令(令)格式、纪要格式
 *
 * 出参: { success: boolean, formatType: string, paragraphCount: number, appliedStyles: object, detectedElements: object }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, formatType: 'general', paragraphCount: 0, appliedStyles: {} };
  }

  // ========================================
  // 格式类型定义
  // ========================================
  var FORMAT_TYPES = {
    GENERAL: 'general',
    LETTER: 'letter',
    ORDER: 'order',
    MINUTES: 'minutes'
  };

  // ========================================
  // 格式类型检测
  // ========================================
  function detectFormatType() {
    var totalParas = doc.Paragraphs.Count;
    var detectedFlags = {
      hasOrder: false,
      hasMinutes: false,
      hasLetter: false,
      hasGeneral: false,
      reasons: []
    };

    for (var i = 1; i <= Math.min(10, totalParas); i++) {
      var para = doc.Paragraphs.Item(i);
      var text = para.Range.Text ? para.Range.Text.replace(/[\r\n]/g, '').trim() : '';
      var alignment = para.Format.Alignment;

      if (!text) continue;

      // 命令(令)格式：发文机关标志含"命令"或"令"
      // 注意：优先检查居中的，但也接受非居中的（原文档可能格式不规范）
      if (/命令|令$/.test(text)) {
        detectedFlags.hasOrder = true;
        detectedFlags.reasons.push('命令格式：发文机关标志含命令/令');
        return FORMAT_TYPES.ORDER;
      }

      // 纪要格式：发文机关标志含"纪要"
      // 注意：优先检查居中的，但也接受非居中的（原文档可能格式不规范）
      if (/纪要/.test(text)) {
        detectedFlags.hasMinutes = true;
        detectedFlags.reasons.push('纪要格式：发文机关标志含纪要');
        return FORMAT_TYPES.MINUTES;
      }

      // 信函格式检测（多种方式）
      // 方式1：发文字号右对齐
      if (/〔\d{4}〕\d+号/.test(text) || /\[\d{4}\]\d+号/.test(text)) {
        if (alignment === 2) {
          detectedFlags.hasLetter = true;
          detectedFlags.reasons.push('信函格式：发文字号右对齐');
          return FORMAT_TYPES.LETTER;
        }
      }

      // 方式2：发文机关标志不含"文件"且标题含"函"
      if (alignment === 1 && !/文件$/.test(text)) {
        if (/政府$|办公室$|委员会$|厅$|局$|委$/.test(text)) {
          // 检查后续段落是否含"函"
          for (var j = i + 1; j <= Math.min(20, totalParas); j++) {
            var nextText = doc.Paragraphs.Item(j).Range.Text;
            if (nextText) {
              var cleanText = nextText.replace(/[\r\n]/g, '').trim();
              if (/函$|复函$/.test(cleanText)) {
                detectedFlags.hasLetter = true;
                detectedFlags.reasons.push('信函格式：标志不含文件+标题含函');
                return FORMAT_TYPES.LETTER;
              }
            }
          }
        }
      }

      // 方式3：标题含"函"且发文机关标志不含"文件"
      if (text.length > 5 && text.length <= 50 && /函$|复函$/.test(text)) {
        // 回溯检查发文机关标志
        for (var k = 1; k < i; k++) {
          var prevText = getParaText(doc.Paragraphs.Item(k));
          if (prevText && /政府$|办公室$|委员会$|厅$|局$|委$/.test(prevText) && !/文件$/.test(prevText)) {
            detectedFlags.hasLetter = true;
            detectedFlags.reasons.push('信函格式：标题含函+标志不含文件');
            return FORMAT_TYPES.LETTER;
          }
        }
      }

      // 通用公文：发文机关标志含"文件"
      if (alignment === 1 && /文件$/.test(text)) {
        detectedFlags.hasGeneral = true;
        detectedFlags.reasons.push('通用格式：发文机关标志含文件');
      }
    }

    // 默认返回通用公文
    if (!detectedFlags.reasons.length) {
      detectedFlags.reasons.push('默认：未识别到特殊格式特征');
    }
    return FORMAT_TYPES.GENERAL;
  }

  var formatType = detectFormatType();

  // ========================================
  // 样式规则定义
  // ========================================
  var STYLE_RULES = {
    // 版头
    'fenzhen': {
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',  // 份号用阿拉伯数字，西文字体用Times New Roman
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    'miji': {
      fontCN: '黑体',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    'jinji': {
      fontCN: '黑体',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    'documentFlag': {
      fontCN: '方正小标宋简体',
      fontSize: 22,  // 22磅=2号字
      bold: true,
      alignment: 1,
      firstLineIndent: 0,
      color: 0x0000FF  // 红色（BGR格式）
    },
    // 纪要标志（纪要格式专用）
    'jiyaoBiaozhi': {
      fontCN: '方正小标宋简体',
      fontSize: 22,  // 22磅，与发文机关标志相同
      bold: false,
      alignment: 1,  // 居中
      firstLineIndent: 0,
      color: 0x0000FF  // 红色（BGR格式）
    },
    // 命令标志（命令格式专用）
    // 国标10.2：发文机关全称+"命令"或"令"，居中，上边缘至版心20mm，红色小标宋体字
    'orderBiaozhi': {
      fontCN: '方正小标宋简体',
      fontSize: 22,  // 22磅=2号字
      bold: false,
      alignment: 1,  // 居中
      firstLineIndent: 0,
      color: 0x0000FF,  // 红色（BGR格式）
      spaceBefore: 57  // 20mm ≈ 57磅
    },
    // 令号（命令格式专用）
    // 国标10.2：标志下空二行居中编排令号
    'orderNumber': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,  // 3号字
      bold: false,
      alignment: 1,  // 居中
      firstLineIndent: 0,
      spaceBefore: 60  // 标志下空二行（每行约30pt）
    },
    // 纪要编号（纪要格式专用）
    'jiyaoNumber': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 1,  // 居中
      firstLineIndent: 0
    },
    'docNumber': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 1,
      firstLineIndent: 0
    },
    'qianfaren': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 2,
      firstLineIndent: 0,
      rightIndent: 10.5  // 10.5磅×20=210缇=1字符，居右空一字
    },

    // 主体
    'title': {
      fontCN: '小标宋',
      fontSize: 22,
      bold: false,  // 小标宋本身具有粗体效果，不需要额外加粗
      alignment: 1,
      firstLineIndent: 0,
      spaceBefore: 32  // 国标7.3.1：红色分隔线下空二行，每行约16pt
    },
    'zhushong': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0,
      spaceBefore: 16  // 国标7.3.2：标题下空一行
    },
    'heading1': {
      fontCN: '黑体',
      fontSize: 16,
      bold: false,
      alignment: 3,  // 两端对齐
      firstLineIndent: 21,  // 21磅×20=420缇=2字符，首行缩进
      lineSpacing: 28
    },
    'heading2': {
      fontCN: '楷体_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 3,  // 两端对齐
      firstLineIndent: 21,  // 21磅×20=420缇=2字符，首行缩进
      lineSpacing: 28
    },
    'heading3': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,  // 国标7.3.3：第三层用仿宋体，不加粗
      alignment: 3,
      firstLineIndent: 21,
      lineSpacing: 28
    },
    'heading4': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,  // 国标7.3.3：第四层用仿宋体，不加粗
      alignment: 3,
      firstLineIndent: 21,
      lineSpacing: 28
    },
    // 出席人员（纪要格式专用）
    // 国标10.3：左空二字编排"出席"二字，后标全角冒号，冒号后用仿宋
    'chuxi': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 0,  // 左对齐
      firstLineIndent: 0,
      leftIndent: 21  // 21磅×20=420缇=2字符，左空二字
    },
    // 请假人员（纪要格式专用）
    'qingjia': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0,
      leftIndent: 21
    },
    // 列席人员（纪要格式专用）
    'liexi': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0,
      leftIndent: 21
    },
    'body': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 3,
      firstLineIndent: 21,  // 21磅×20=420缇=2字符
      lineSpacing: 28
    },
    'fujian': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 21,  // 21磅×20=420缇=2字符，左空二字
      lineSpacing: 28
    },
    'fujianContinue': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 0,
      // 缩进计算（与附件名称首字对齐）：
      // 第一行："附件：1.城市..."（"附件："左空2字）
      //   "附件："=4字 + "1."=2字，附件名称首字"城"在第8字位置
      // 第二行："2.安全生产..."
      //   "2."=2字，要让"安"在第8字位置
      //   缩进 = 8 - 2 = 6字 = 1260缇 = 63磅
      firstLineIndent: 63  // 63磅×20=1260缇=6字符
    },
    'signature': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 2,
      firstLineIndent: 0
      // 右缩进在署名检测逻辑中根据日期长度计算
    },
    'date': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 2,
      firstLineIndent: 0,
      rightIndent: 21  // 21磅×20=420缇=2字符（日期右空二字，固定）
    },
    'fuzhu': {
      fontCN: '仿宋_GB2312',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 21
    },

    // 版记
    'chaosong': {
      fontCN: '仿宋_GB2312',
      fontSize: 14,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    'yinfa': {
      fontCN: '仿宋_GB2312',
      fontSize: 14,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },

    // 附件
    'attachmentMarker': {
      fontCN: '黑体',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    'attachmentTitle': {
      fontCN: '黑体',
      fontSize: 16,
      bold: false,
      alignment: 1,
      firstLineIndent: 0
    }
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

  function setParaTextPreserveMark(para, text) {
    try {
      var range = para.Range.Duplicate;
      if (range.End > range.Start) {
        range.End = range.End - 1;  // 保留段落标记，避免后续内容粘连到本段
      }
      range.Text = text || '';
      return true;
    } catch (e) {
      try {
        para.Range.Text = text || '';
        return true;
      } catch (e2) {
        return false;
      }
    }
  }

  function extractLeadingDate(text) {
    if (!text) return '';
    var m = text.match(/^\s*(\d{4}年\d{1,2}月\d{1,2}日)/);
    return m ? m[1] : '';
  }

  function normalizeChaosongText(text) {
    if (!text) return '';
    return text
      .replace(/^\s*\d{4}年\d{1,2}月\d{1,2}日[\s\u3000]*/, '')
      .replace(/^[\s\u3000]*抄送[：:]/, '')
      .replace(/^[\s\u3000]*送[：:]/, '')
      .replace(/^[\s\u3000]+/, '')
      .replace(/[\s\u3000]+$/, '')
      .replace(/[。]+$/, '');
  }

  function markBanjiShape(shape, kind) {
    if (!shape) return;
    var tag = 'GBT9704_BANJI_' + (kind || 'ITEM');
    try { shape.Name = tag; } catch (e) {}
    try { shape.AlternativeText = tag; } catch (e) {}
    try { shape.Title = tag; } catch (e) {}
  }

  function isOwnBanjiShape(shape) {
    if (!shape) return false;
    var candidates = [];
    try { candidates.push(shape.Name || ''); } catch (e) {}
    try { candidates.push(shape.AlternativeText || ''); } catch (e) {}
    try { candidates.push(shape.Title || ''); } catch (e) {}
    for (var i = 0; i < candidates.length; i++) {
      if (/^GBT9704_BANJI_/.test(candidates[i])) return true;
    }
    return false;
  }

  function cleanupBottomBanjiShapes(minTopPt) {
    try {
      var shapes = doc.Shapes;
      if (!shapes || shapes.Count <= 0) return;
      for (var i = shapes.Count; i >= 1; i--) {
        try {
          var shape = shapes.Item(i);
          if (!isOwnBanjiShape(shape)) continue;
          var topPos = 0;
          try { topPos = shape.Top || 0; } catch (eTop) {}
          if (topPos < minTopPt) continue;
          shape.Delete();
        } catch (eInner) {}
      }
    } catch (e) {
      console.warn('[apply-format] 清理底部版记Shape失败', e);
    }
  }

  function addFixedTextBox(anchorRange, x, y, width, height, text, alignment, setTabStop) {
    try {
      var shapes = doc.Shapes;
      if (!shapes) return null;
      var textBox = shapes.AddTextbox(1, x, y, width, height, anchorRange);
      textBox.TextFrame.MarginLeft = 10.5;
      textBox.TextFrame.MarginRight = 10.5;
      textBox.TextFrame.MarginTop = 0;
      textBox.TextFrame.MarginBottom = 0;
      textBox.TextFrame.WordWrap = true;
      textBox.Line.Visible = false;
      textBox.Fill.Visible = false;

      var tf = textBox.TextFrame;
      if (tf && tf.TextRange) {
        tf.TextRange.Text = text;
        tf.TextRange.Font.NameFarEast = '仿宋_GB2312';
        tf.TextRange.Font.Name = '仿宋_GB2312';
        tf.TextRange.Font.Size = 14;
        tf.TextRange.ParagraphFormat.Alignment = alignment;
        if (setTabStop) {
          try {
            tf.TextRange.ParagraphFormat.TabStops.Add(width - 21, 2, 0);
          } catch (eTab) {}
        }
      }

      try {
        textBox.RelativeVerticalPosition = 1;
        textBox.Top = y;
        textBox.RelativeHorizontalPosition = 1;
        textBox.Left = x;
      } catch (ePos) {}
      markBanjiShape(textBox, setTabStop ? 'TEXT_YINFA' : 'TEXT_CHAOSONG');
      return textBox;
    } catch (e) {
      console.warn('[apply-format] 添加版记文本框失败', e);
      return null;
    }
  }

  function addFixedLine(anchorRange, x, y, width, weight, color) {
    try {
      var shapes = doc.Shapes;
      if (!shapes) return null;
      var line = shapes.AddLine(x, y, x + width, y, anchorRange);
      line.Line.ForeColor.RGB = color;
      line.Line.Weight = weight;
      line.WrapFormat.Type = 3;
      try { line.Shadow.Visible = false; } catch (eShadow) {}
      try {
        line.RelativeVerticalPosition = 1;
        line.Top = y;
        line.RelativeHorizontalPosition = 1;
        line.Left = x;
      } catch (ePos) {}
      markBanjiShape(line, 'LINE');
      return line;
    } catch (e) {
      console.warn('[apply-format] 添加版记线条失败', e);
      return null;
    }
  }

  function createBottomBanjiLayout(anchorRange, chaosongText, yinfaContent) {
    var mm = 2.835;
    var x = 28 * mm;
    var width = 156 * mm;
    var rowTopY = chaosongText ? (247 * mm) : (254.5 * mm);
    var midY = 254.5 * mm;
    var bottomY = 262 * mm;
    var rowHeight = 6.2 * mm;
    var thickLineWeight = 1.0;   // 首末分隔线，保持单线但避免视觉过粗
    var thinLineWeight = 0.7;    // 中间分隔线，较首末线更细

    cleanupBottomBanjiShapes(680);

    if (chaosongText) {
      addFixedLine(anchorRange, x, rowTopY, width, thickLineWeight, 0);
      addFixedTextBox(anchorRange, x, rowTopY + (1.0 * mm), width, rowHeight, '抄送：' + chaosongText + '。', 0, false);
      addFixedLine(anchorRange, x, midY, width, thinLineWeight, 0);
      addFixedTextBox(anchorRange, x, midY + (1.0 * mm), width, rowHeight, yinfaContent, 0, true);
      addFixedLine(anchorRange, x, bottomY, width, thickLineWeight, 0);
    } else {
      addFixedLine(anchorRange, x, rowTopY, width, thickLineWeight, 0);
      addFixedTextBox(anchorRange, x, rowTopY + (1.0 * mm), width, rowHeight, yinfaContent, 0, true);
      addFixedLine(anchorRange, x, bottomY, width, thickLineWeight, 0);
    }
  }

  function getAlignment(para) {
    try {
      return para.Format.Alignment;
    } catch (e) {
      return -1;
    }
  }

  function isEmpty(text) {
    return !text || text.length === 0;
  }

  // ========================================
  // 识别规则函数
  // ========================================

  function isFenzhen(text, index) {
    if (index > 5) return false;
    return /^\d{6}$/.test(text);
  }

  function isMiji(text, index) {
    if (index > 5) return false;
    return /秘密|机密|绝密/.test(text);
  }

  function isJinji(text, index) {
    if (index > 5) return false;
    return /特急|加急/.test(text);
  }

  function isDocumentFlag(text, index, separatorIndex) {
    if (separatorIndex > 0 && index > separatorIndex) return false;
    if (/文件$/.test(text)) return true;
    if (/^.+政府$|^.+办公室$|^.+委员会$|^.+厅$|^.+局$|^.+委$/.test(text)) return true;
    return false;
  }

  function isDocNumber(text) {
    // 发文字号段落应该是独立的发文字号，不是嵌入在句子中的引用
    // 国标要求使用六角括号"〔〕"
    // 检查是否以发文字号开头，或者整个段落就是发文字号
    if (/^[^〕]+〔\d{4}〕\d+号$/.test(text)) return true;  // 如 "XX政函〔2024〕56号"
    if (/^[^〕]+〔\d{4}〕\d+号\s*$/.test(text)) return true;
    // 注意：不再接受方括号[]，国标要求使用六角括号〔〕
    // 排除包含在句子中的引用（如 "你委《...》（XX发改函〔2024〕56号）收悉"）
    // 如果段落包含句号、引号等句子标点，说明不是独立的发文字号
    if (/[。！？；：]/.test(text)) return false;
    // 如果段落以"你委"、"《"等开头，说明不是发文字号
    if (/^[你我院校公司《]/.test(text)) return false;
    return false;
  }

  function isQianfaren(text) {
    return /^签发人[：:]/.test(text);
  }

  function isZhushong(text, titleIndex, currentIndex) {
    if (!text) return false;

    // 长度判断：主送机关通常较短（机构名称列表）
    if (text.length > 100) return false;

    // 位置判断：主送机关应在标题后几个段落内
    if (titleIndex && currentIndex && currentIndex - titleIndex > 5) return false;

    // 必须以冒号结尾
    if (/[：:]$/.test(text)) {
      // 主送机关特征：包含机关关键词
      // 注意：机关列表可能包含逗号、顿号分隔，不应排除
      // 排除完整句子（含句号、感叹号、问号）
      if (/[。！？]/.test(text)) return false;

      // 包含机关关键词则识别为主送机关
      if (/政府|办公室|委员会|厅|局|委|公司|各|有关|部门|机构|单位|县|市|省|区/.test(text)) {
        return true;
      }
    }
    return false;
  }

  function isHeading1(text) {
    // 格式1: "一、"、"二、" 等
    if (/^[一二三四五六七八九十]+、/.test(text)) return true;
    // 格式2: "第一章"、"第二章" 等（命令、法规文件常用格式）
    if (/^第[一二三四五六七八九十]+章/.test(text)) return true;
    return false;
  }

  function isHeading2(text) {
    return /^[（(][一二三四五六七八九十]+[）)]/.test(text);
  }

  function isHeading3(text) {
    return /^\d+[\.\.]/.test(text);
  }

  function isHeading4(text) {
    return /^[（(]\d+[）)]/.test(text);
  }

  // ========================================
  // 命令(令)格式专用识别函数
  // ========================================

  // 命令标志：发文机关全称+"命令"或"令"
  function isOrderBiaozhi(text, index) {
    // 第一段或前3段中以"命令"或"令"结尾的段落
    if (index <= 3 && /命令$|令$/.test(text)) {
      return true;
    }
    return false;
  }

  // 令号：第X号 格式
  function isOrderNumber(text, index) {
    // 格式：第18号
    if (/^第\d+号$/.test(text)) {
      return true;
    }
    return false;
  }

  // ========================================
  // 纪要格式专用识别函数
  // ========================================

  // 纪要标志：前几段含"纪要"
  // 注意：不要求居中，因为原文档可能格式不规范
  function isJiyaoBiaozhi(text, index, alignment) {
    // 第一段或前3段中包含"纪要"结尾的段落
    if (index <= 3 && /纪要$/.test(text)) {
      return true;
    }
    return false;
  }

  // 纪要编号：〔YYYY〕第X号 格式
  // 注意：不要求居中，因为原文档可能格式不规范
  function isJiyaoNumber(text, alignment) {
    // 标准格式：〔2024〕第8号
    if (/^〔\d{4}〕第\d+号$/.test(text)) {
      return true;
    }
    // 变体格式：[2024]第8号（方括号）
    if (/^\[\d{4}\]第\d+号$/.test(text)) {
      return true;
    }
    // 变体格式：〔2024〕8号（省略"第"字）
    if (/^〔\d{4}〕\d+号$/.test(text)) {
      return true;
    }
    return false;
  }

  // 出席人员：以"出席："或"出席："开头
  function isChuxi(text) {
    return /^出席[：:]/.test(text);
  }

  // 请假人员：以"请假："或"请假："开头
  function isQingjia(text) {
    return /^请假[：:]/.test(text);
  }

  // 列席人员：以"列席："或"列席："开头
  function isLiexi(text) {
    return /^列席[：:]/.test(text);
  }

  // 纪要版记"送："
  function isJiyaoBanji(text) {
    return /^送[：:]/.test(text);
  }

  function isFujian(text) {
    return /^附件[：:]/.test(text);
  }

  // 附件说明续行：以数字编号开头，前一段是附件说明或附件说明续行
  function isFujianContinue(text, index) {
    // 匹配更多点号格式：半角点、全角点、顿号等
    if (!/^\d+[.．、．].+/.test(text)) return false;
    if (index <= 1) return false;
    // 检查前一段是否是附件说明或附件说明续行
    // 注意：可能前面插入了空行，需要跳过空行查找
    try {
      for (var j = index - 1; j >= 1; j--) {
        var prevText = getParaText(doc.Paragraphs.Item(j));
        if (prevText && prevText.length > 0) {
          // 找到非空段落，检查是否是附件说明或附件续行
          return /^附件[：:]/.test(prevText) || /^\d+[.．、．].+/.test(prevText);
        }
      }
    } catch (e) {}
    return false;
  }

  function isDate(text) {
    // 成文日期段落通常只包含日期，不含"印发"等版记内容
    // 长度通常在8-15字符之间
    if (!/\d{4}年\d{1,2}月\d{1,2}日/.test(text)) return false;
    if (/印发/.test(text)) return false;  // 排除印发行
    if (text.length > 20) return false;   // 排除过长的段落
    return true;
  }

  function isSignature(text, index, total, dateIndex) {
    if (!text) return false;
    // 命令格式署名：职务+姓名（如"市长 李四"、"省长 张三"）
    if (/^(市长|省长|县长|区长|州长|主任|局长|厅长|署长)[\s\u3000]+[\u4e00-\u9fa5]{2,4}$/.test(text)) {
      return true;
    }
    // 机关署名：以机关名称结尾
    if (!/政府$|办公室$|委员会$|厅$|局$|委$/.test(text)) return false;
    // 排除版记内容
    if (/抄送|印发|送[：:]/.test(text)) return false;
    // 长度和标点检查
    if (text.length > 30 || /[，。！？、；：]$/.test(text)) return false;
    return true;
  }

  function isFuzhu(text) {
    return /^\（[^）]+\）$/.test(text);
  }

  function isChaosong(text) {
    return /^\s*抄送[：:]/.test(text) || /^[\u3000]*抄送[：:]/.test(text);  // 允许前面有空格或全角空格
  }

  function isYinfa(text) {
    // 印发行：含"印发"且含日期
    return /印发/.test(text) && /\d{4}年/.test(text);
  }

  // 附件标识：如 "附件1"、"附件：1."
  function isAttachmentMarker(text) {
    return /^附件\d*$/.test(text) || /^附件[：:]\d*[\.\.]/.test(text);
  }

  // 附件标题：附件文档内的标题
  function isAttachmentTitle(text, prevIsMarker) {
    // 如果前一段是附件标识，当前段落可能是附件标题
    if (prevIsMarker && text.length > 0 && text.length <= 50) {
      return true;
    }
    return false;
  }

  // ========================================
  // 应用样式函数
  // ========================================
  function applyStyle(para, rule) {
    var range = para.Range;

    // 设置字体
    if (range && range.Font) {
      if (rule.fontCN) {
        range.Font.NameFarEast = rule.fontCN;
        range.Font.Name = rule.fontCN;  // 同时设置Name确保生效
        range.Font.NameAscii = rule.fontCN;  // 设置Ascii属性
        range.Font.NameOther = rule.fontCN;  // 设置Other属性
        // 强制重新设置FarEast确保生效（解决某些WPS版本eastAsia属性问题）
        range.Font.NameFarEast = rule.fontCN;
      }
      if (rule.fontEN) {
        range.Font.NameAscii = rule.fontEN;  // 西文字体
        range.Font.NameOther = rule.fontEN;  // 其他字符字体
      }
      if (rule.fontSize) range.Font.Size = rule.fontSize;
      if (rule.bold !== undefined) range.Font.Bold = rule.bold;
      if (rule.color !== undefined) range.Font.Color = rule.color;
    }

    // 设置段落格式 - 优先使用Range.ParagraphFormat
    var pFmt = null;
    try {
      pFmt = range.ParagraphFormat;
    } catch (e) {}

    if (!pFmt) {
      try {
        pFmt = para.Format;
      } catch (e) {}
    }

    if (pFmt) {
      if (rule.alignment !== undefined) pFmt.Alignment = rule.alignment;
      if (rule.firstLineIndent !== undefined) pFmt.FirstLineIndent = rule.firstLineIndent;
      if (rule.leftIndent !== undefined) pFmt.LeftIndent = rule.leftIndent;
      if (rule.rightIndent !== undefined) pFmt.RightIndent = rule.rightIndent;
      if (rule.spaceBefore !== undefined) pFmt.SpaceBefore = rule.spaceBefore;
      if (rule.spaceAfter !== undefined) pFmt.SpaceAfter = rule.spaceAfter;
      if (rule.lineSpacing !== undefined) {
        try {
          pFmt.LineSpacingRule = 4;  // wdLineSpaceExactly
        } catch (e) {}
        pFmt.LineSpacing = rule.lineSpacing;
      }
    }
  }

  // ========================================
  // 主流程
  // ========================================
  var appliedStyles = {};
  var detectedElements = {};
  var totalParas = doc.Paragraphs.Count;

  // 第一遍：识别关键位置
  var separatorIndex = -1;
  var titleIndex = -1;
  var dateIndex = -1;
  var signatureIndex = -1;  // 署名位置
  var fuzhuIndex = -1;      // 附注位置
  var attachmentStartIndex = -1;
  var chaosongIndex = -1;
  var yinfaIndex = -1;

  for (var i = 1; i <= totalParas; i++) {
    var para = doc.Paragraphs.Item(i);
    var text = getParaText(para);

    if (isEmpty(text)) continue;

    // 发文字号
    if (isDocNumber(text)) {
      separatorIndex = i;
      // 继续检查下一段是否是签发人，如果是，分隔线位置后移
      if (i < totalParas) {
        var nextText = getParaText(doc.Paragraphs.Item(i + 1));
        if (isQianfaren(nextText)) {
          separatorIndex = i + 1;  // 分隔线在签发人之后
        }
      }
    }

    // 纪要编号：纪要格式中作为分隔线位置
    var alignment = para.Format.Alignment;
    if (isJiyaoNumber(text, alignment)) {
      separatorIndex = i;
    }

    // 命令(令)格式：令号作为分隔线位置
    // 国标10.2：标志下空二行居中编排令号，令号下空二行编排正文
    if (formatType === FORMAT_TYPES.ORDER && isOrderNumber(text, i)) {
      separatorIndex = i;
      console.log('[apply-format] 命令格式：使用令号作为分隔线位置，索引=' + i);
    }

    // 纪要格式备选：如果没找到纪要编号，用纪要标志作为分隔线位置
    if (formatType === FORMAT_TYPES.MINUTES && separatorIndex < 0) {
      if (isJiyaoBiaozhi(text, i, alignment)) {
        separatorIndex = i;
        console.log('[apply-format] 纪要格式：使用纪要标志作为分隔线位置，索引=' + i);
      }
    }

    // 命令格式备选：如果没找到令号，用命令标志作为分隔线位置
    if (formatType === FORMAT_TYPES.ORDER && separatorIndex < 0) {
      if (isOrderBiaozhi(text, i)) {
        separatorIndex = i;
        console.log('[apply-format] 命令格式：使用命令标志作为分隔线位置，索引=' + i);
      }
    }

    // 标题：分隔线之后，排除签发人、主送机关等
    if (titleIndex < 0 && separatorIndex > 0 && i > separatorIndex) {
      // 排除签发人、主送机关、出席/请假/列席等
      if (!isQianfaren(text) && !isZhushong(text, titleIndex, i) && !isChuxi(text) && !isQingjia(text) && !isLiexi(text) && text.length > 0 && text.length <= 50) {
        titleIndex = i;
        console.log('[apply-format] 标题检测: 索引=' + i + ', 文本="' + text + '"');
      }
    }

    // 纪要格式特殊标题检测：含"纪要"且长度适中
    // 注意：不要求居中，因为原文档可能格式不规范
    if (formatType === FORMAT_TYPES.MINUTES && titleIndex < 0 && i > 1 && i <= 10 && text.length > 5 && text.length <= 50 && /纪要$/.test(text)) {
      // 排除版头标志（纪要标志通常在前3段）
      if (i > 2 || !isJiyaoBiaozhi(text, i, alignment)) {
        titleIndex = i;
        console.log('[apply-format] 纪要标题检测: 索引=' + i + ', 文本="' + text + '", 对齐=' + alignment);
      }
    }

    // 纪要格式：检测"会议纪要"模式的标题（如"XX市人民政府第8次常务会议纪要"）
    if (formatType === FORMAT_TYPES.MINUTES && titleIndex < 0 && i > 2 && i <= 10) {
      if (/政府.*会议纪要$/.test(text) || /办公室.*会议纪要$/.test(text)) {
        titleIndex = i;
        console.log('[apply-format] 纪要会议标题检测: 索引=' + i + ', 文本="' + text + '"');
      }
    }

    // 纪要格式备选标题检测：如果仍未找到标题，查找分隔线后的第一个非特殊段落
    if (formatType === FORMAT_TYPES.MINUTES && titleIndex < 0 && separatorIndex > 0 && i > separatorIndex && i <= separatorIndex + 5) {
      // 排除版头元素和出席/请假/列席
      if (!isQianfaren(text) && !isZhushong(text, titleIndex, i) && !isChuxi(text) && !isQingjia(text) && !isLiexi(text) &&
          !isJiyaoBiaozhi(text, i, para.Format.Alignment) && !isJiyaoNumber(text, para.Format.Alignment) &&
          text.length > 5 && !/^[〔\[]/.test(text)) {
        titleIndex = i;
        console.log('[apply-format] 纪要备选标题检测: 索引=' + i + ', 文本="' + text + '"');
      }
    }

    if (isDate(text) && dateIndex < 0) {
      dateIndex = i;
    }

    // 署名：日期前面2个段落内，非日期、非版记、非附注的短段落
    if (dateIndex > 0 && Math.abs(i - dateIndex) <= 2 && signatureIndex < 0) {
      if (!isDate(text) && !isChaosong(text) && !isYinfa(text) && !isFuzhu(text) &&
          text.length > 0 && text.length <= 30 && !/[，。！？、；：]$/.test(text)) {
        signatureIndex = i;
      }
    }

    // 附注（如"（此件公开发布）"）
    if (isFuzhu(text)) {
      fuzhuIndex = i;
    }

    // 附件区域开始：附注之后的"附件X"标识（在版记之前）
    if (fuzhuIndex > 0 && i > fuzhuIndex && /^附件\d*$/.test(text)) {
      if (attachmentStartIndex < 0) {
        attachmentStartIndex = i;
      }
    }

    if (isChaosong(text) || isJiyaoBanji(text)) {
      chaosongIndex = i;
    }

    if (isYinfa(text)) {
      yinfaIndex = i;
    }
  }

  // ========================================
  // 调试输出：第一遍检测结果
  // ========================================
  if (formatType === FORMAT_TYPES.MINUTES) {
    console.log('[apply-format] 纪要格式检测结果:');
    console.log('  - separatorIndex: ' + separatorIndex);
    console.log('  - titleIndex: ' + titleIndex);
    console.log('  - dateIndex: ' + dateIndex);
    console.log('  - chaosongIndex: ' + chaosongIndex);
    console.log('  - yinfaIndex: ' + yinfaIndex);
    // 输出前5段信息
    for (var di = 1; di <= Math.min(5, totalParas); di++) {
      var dt = getParaText(doc.Paragraphs.Item(di));
      var da = doc.Paragraphs.Item(di).Format.Alignment;
      console.log('  - 第' + di + '段: [' + da + '] ' + dt.substring(0, 30) + (dt.length > 30 ? '...' : ''));
    }
  }

  if (formatType === FORMAT_TYPES.ORDER) {
    console.log('[apply-format] 命令格式检测结果:');
    console.log('  - separatorIndex: ' + separatorIndex);
    console.log('  - titleIndex: ' + titleIndex);
    console.log('  - dateIndex: ' + dateIndex);
    // 输出前5段信息
    for (var di = 1; di <= Math.min(5, totalParas); di++) {
      var dt = getParaText(doc.Paragraphs.Item(di));
      var da = doc.Paragraphs.Item(di).Format.Alignment;
      console.log('  - 第' + di + '段: [' + da + '] ' + dt.substring(0, 30) + (dt.length > 30 ? '...' : ''));
    }
  }

  // 第二遍：应用样式
  var inAttachmentSection = false;
  var prevIsAttachmentMarker = false;

  for (var i = 1; i <= totalParas; i++) {
    var para = doc.Paragraphs.Item(i);
    var text = getParaText(para);

    if (isEmpty(text)) {
      prevIsAttachmentMarker = false;
      continue;
    }

    var type = 'body';

    // 判断是否进入附件区域（附注之后、版记之前）
    if (attachmentStartIndex > 0 && i >= attachmentStartIndex && (chaosongIndex < 0 || i < chaosongIndex)) {
      inAttachmentSection = true;
    }

    // 识别元素类型
    if (separatorIndex < 0 || i <= separatorIndex) {
      // 版头区域
      // 命令(令)格式：优先识别命令标志和令号
      if (formatType === FORMAT_TYPES.ORDER) {
        if (isOrderBiaozhi(text, i)) type = 'orderBiaozhi';
        else if (isOrderNumber(text, i)) type = 'orderNumber';
      }
      // 纪要格式：优先识别纪要标志和编号
      else if (formatType === FORMAT_TYPES.MINUTES) {
        var alignment = para.Format.Alignment;
        if (isJiyaoBiaozhi(text, i, alignment)) type = 'jiyaoBiaozhi';
        else if (isJiyaoNumber(text, alignment)) type = 'jiyaoNumber';
      }
      // 通用格式
      if (type === 'body') {
        if (isFenzhen(text, i)) type = 'fenzhen';
        else if (isMiji(text, i)) type = 'miji';
        else if (isJinji(text, i)) type = 'jinji';
        else if (isDocumentFlag(text, i, separatorIndex)) type = 'documentFlag';
        else if (isDocNumber(text)) type = 'docNumber';
        else if (isQianfaren(text)) type = 'qianfaren';
      }
    } else if (inAttachmentSection) {
      // 附件区域：只识别一级标题(一、二、)和二级标题（（一）（二）），编号段落按正文处理
      if (isAttachmentMarker(text)) {
        type = 'attachmentMarker';
        prevIsAttachmentMarker = true;
      } else if (isAttachmentTitle(text, prevIsAttachmentMarker)) {
        type = 'attachmentTitle';
        prevIsAttachmentMarker = false;
      } else if (isHeading1(text)) {
        type = 'heading1';
        prevIsAttachmentMarker = false;
      } else if (isHeading2(text)) {
        type = 'heading2';
        prevIsAttachmentMarker = false;
      } else {
        // 附件内的编号段落(1. 2. （1）（2）)按正文格式处理，不加粗
        type = 'body';
        prevIsAttachmentMarker = false;
      }
    } else if (chaosongIndex < 0 || i < chaosongIndex) {
      // 主体区域
      if (i === titleIndex) type = 'title';
      else if (isZhushong(text, titleIndex, i)) type = 'zhushong';
      else if (isHeading1(text)) type = 'heading1';
      else if (isHeading2(text)) type = 'heading2';
      // 纪要格式：出席/请假/列席人员
      else if (isChuxi(text)) type = 'chuxi';
      else if (isQingjia(text)) type = 'qingjia';
      else if (isLiexi(text)) type = 'liexi';
      // 附件说明续行：以数字编号开头，前一段是附件说明，按正文处理不加粗
      else if (isFujianContinue(text, i)) type = 'fujianContinue';
      else if (isHeading3(text)) type = 'heading3';
      else if (isHeading4(text)) type = 'heading4';
      else if (isFujian(text)) type = 'fujian';
      else if (isDate(text)) type = 'date';
      else if (isFuzhu(text)) type = 'fuzhu';
      else if (isSignature(text, i, totalParas, dateIndex)) type = 'signature';
    } else {
      // 版记区域
      if (isChaosong(text)) type = 'chaosong';
      else if (isJiyaoBanji(text)) type = 'chaosong';  // 纪要的"送："按抄送格式处理
      else if (isYinfa(text)) type = 'yinfa';
    }

    // ========================================
    // 命令(令)格式特殊处理 - 必须在样式应用之前执行
    // 国标10.2：发文机关标志居中，上边缘至版心20mm，红色小标宋体字
    // 标志下空二行居中编排令号，令号下空二行编排正文
    // ========================================
    if (formatType === FORMAT_TYPES.ORDER) {
      // 命令标志处理：发文机关全称+"命令"或"令"
      if (/命令$|令$/.test(text) && i <= 3) {
        type = 'orderBiaozhi';
      }
      // 令号处理：第X号
      else if (/^第\d+号$/.test(text)) {
        type = 'orderNumber';
      }
    }

    // 应用样式
    var rule = STYLE_RULES[type] || STYLE_RULES['body'];
    rule.type = type;
    applyStyle(para, rule);

    // 命令格式：强制设置样式（覆盖可能的不生效问题）
    if (formatType === FORMAT_TYPES.ORDER) {
      // 命令标志强制设置
      if (/命令$|令$/.test(text) && i <= 3) {
        try {
          // 先清除所有Run中的格式，确保不继承错误样式
          para.Range.Select();
          var sel = Application.Selection;
          if (sel) {
            // 清除所有格式
            try { sel.ClearFormatting(); } catch (e) {}
            // 设置正确的字体
            sel.Font.NameFarEast = '方正小标宋简体';
            sel.Font.Name = '方正小标宋简体';
            sel.Font.NameAscii = '方正小标宋简体';  // 同时设置Ascii
            sel.Font.NameOther = '方正小标宋简体';  // 设置Other属性
            sel.Font.Size = 22;  // 2号字
            sel.Font.Color = 0x0000FF;  // BGR格式的红色
            sel.ParagraphFormat.Alignment = 1;  // 居中
            sel.ParagraphFormat.FirstLineIndent = 0;
            if (i === 1) {
              sel.ParagraphFormat.SpaceBefore = 57;  // 20mm ≈ 57磅
            }
            sel.Collapse(0);
            try {
              var runs = para.Range.Words;
              for (var ri = 1; ri <= runs.Count; ri++) {
                runs.Item(ri).Font.NameFarEast = '方正小标宋简体';
                runs.Item(ri).Font.Name = '方正小标宋简体';
                runs.Item(ri).Font.NameAscii = '方正小标宋简体';
                runs.Item(ri).Font.NameOther = '方正小标宋简体';
                runs.Item(ri).Font.Size = 22;
                runs.Item(ri).Font.Color = 0x0000FF;
              }
            } catch (e) {}
            console.log('[apply-format] 命令标志样式已设置(Selection): 方正小标宋简体22号红色居中');
          } else {
            // 备选方案：直接使用Range API
            var range = para.Range;
            range.Font.NameFarEast = '方正小标宋简体';
            range.Font.Name = '方正小标宋简体';
            range.Font.NameAscii = '方正小标宋简体';
            range.Font.NameOther = '方正小标宋简体';
            range.Font.Size = 22;
            range.Font.Color = 0x0000FF;
            var pFmt = range.ParagraphFormat || para.Format;
            if (pFmt) {
              pFmt.Alignment = 1;
              pFmt.FirstLineIndent = 0;
              if (i === 1) {
                pFmt.SpaceBefore = 57;
              }
            }
            try {
              var runs2 = para.Range.Words;
              for (var rj = 1; rj <= runs2.Count; rj++) {
                runs2.Item(rj).Font.NameFarEast = '方正小标宋简体';
                runs2.Item(rj).Font.Name = '方正小标宋简体';
                runs2.Item(rj).Font.NameAscii = '方正小标宋简体';
                runs2.Item(rj).Font.NameOther = '方正小标宋简体';
                runs2.Item(rj).Font.Size = 22;
                runs2.Item(rj).Font.Color = 0x0000FF;
              }
            } catch (e) {}
            console.log('[apply-format] 命令标志样式已设置(Range): 方正小标宋简体22号红色居中');
          }
        } catch (e) {
          console.warn('[apply-format] 命令标志样式设置失败', e);
          // 最终备选
          try {
            para.Range.Font.NameFarEast = '方正小标宋简体';
            para.Range.Font.Name = '方正小标宋简体';
            para.Range.Font.Size = 22;
            para.Range.Font.Color = 0x0000FF;
          } catch (e2) {}
        }
      }
      // 令号强制设置：仿宋3号，标志下空二行
      else if (/^第\d+号$/.test(text)) {
        try {
          // 清除段落属性中的字体样式（避免继承命令标志的样式）
          try {
            var paraStyle = para.Style;
            if (paraStyle) {
              para.Style = null;  // 清除样式引用
            }
          } catch (styleErr) {}

          para.Range.Select();
          var sel = Application.Selection;
          if (sel) {
            // 清除所有格式
            try { sel.ClearFormatting(); } catch (e) {}
            sel.Font.NameFarEast = '仿宋_GB2312';
            sel.Font.Name = '仿宋_GB2312';
            sel.Font.NameAscii = '仿宋_GB2312';
            sel.Font.NameOther = '仿宋_GB2312';
            sel.Font.Size = 16;  // 3号字
            sel.ParagraphFormat.Alignment = 1;  // 居中
            sel.ParagraphFormat.FirstLineIndent = 0;
            sel.ParagraphFormat.SpaceBefore = 60;  // 标志下空二行
            sel.Collapse(0);
            console.log('[apply-format] 令号样式已设置(Selection): 仿宋3号居中');
          }
          // 双重保障：再用Range设置一次
          var range = para.Range;
          range.Font.NameFarEast = '仿宋_GB2312';
          range.Font.Name = '仿宋_GB2312';
          range.Font.NameAscii = '仿宋_GB2312';
          range.Font.NameOther = '仿宋_GB2312';
          range.Font.Size = 16;
          var pFmt = range.ParagraphFormat || para.Format;
          if (pFmt) {
            pFmt.Alignment = 1;
            pFmt.FirstLineIndent = 0;
            pFmt.SpaceBefore = 60;
          }
          console.log('[apply-format] 令号样式已设置(Range): 仿宋3号居中');
        } catch (e) {
          console.warn('[apply-format] 令号样式设置失败', e);
        }
      }
    }

    // 强制确保正文段落有缩进（针对WPS API可能不生效的情况）
    if (type === 'body' && rule.firstLineIndent) {
      try {
        var pFmt = para.Range.ParagraphFormat || para.Format;
        if (pFmt) {
          // 强制设置两次确保生效
          pFmt.FirstLineIndent = rule.firstLineIndent;
          pFmt.Alignment = rule.alignment;
        }
      } catch (e) {}
    }

    // 强制确保主送机关顶格（无缩进）
    if (type === 'zhushong') {
      try {
        var pFmt = para.Range.ParagraphFormat || para.Format;
        if (pFmt) {
          pFmt.FirstLineIndent = 0;
          pFmt.Alignment = 0;  // 左对齐顶格
        }
      } catch (e) {}
    }

    // 强制确保标题格式正确（小标宋2号、居中）
    if (type === 'title') {
      try {
        para.Range.Select();
        var sel = Application.Selection;
        if (sel) {
          sel.Font.NameFarEast = '小标宋';
          sel.Font.Name = '小标宋';
          sel.Font.NameAscii = '小标宋';  // 同时设置Ascii
          sel.Font.NameOther = '小标宋';  // 设置Other属性
          sel.Font.Size = 22;  // 2号字
          sel.ParagraphFormat.FirstLineIndent = 0;
          sel.ParagraphFormat.Alignment = 1;  // 居中
          // 纪要格式特殊处理：国标7.3.1 红色分隔线下空二行
          // 纪要格式无红色分隔线，但标题仍应在发文字号下空二行
          if (formatType === FORMAT_TYPES.MINUTES) {
            sel.ParagraphFormat.SpaceBefore = 60;  // 下空二行，每行约30pt
          }
          // 命令格式特殊处理：国标10.2 令号下空二行编排正文
          if (formatType === FORMAT_TYPES.ORDER) {
            sel.ParagraphFormat.SpaceBefore = 60;  // 下空二行，每行约30pt
          }
          sel.Collapse(0);
        } else {
          // 备选方案
          var range = para.Range;
          range.Font.NameFarEast = '小标宋';
          range.Font.Name = '小标宋';
          range.Font.NameAscii = '小标宋';
          range.Font.NameOther = '小标宋';
          range.Font.Size = 22;
          var pFmt = range.ParagraphFormat || para.Format;
          if (pFmt) {
            pFmt.FirstLineIndent = 0;
            pFmt.Alignment = 1;
            if (formatType === FORMAT_TYPES.MINUTES) {
              pFmt.SpaceBefore = 60;
            }
            if (formatType === FORMAT_TYPES.ORDER) {
              pFmt.SpaceBefore = 60;
            }
          }
        }
        console.log('[apply-format] 标题样式已设置: 小标宋2号居中');
      } catch (e) {
        console.warn('[apply-format] 标题样式设置失败', e);
      }
    }

    // ========================================
    // 命令格式特殊处理：签发人段落段前间距
    // 国标7.3.5.3：正文（或附件说明）下空二行右空四字加盖签发人签名章
    // ========================================
    if (formatType === FORMAT_TYPES.ORDER && type === 'signature') {
      try {
        var pFmt = para.Range.ParagraphFormat || para.Format;
        if (pFmt) {
          pFmt.SpaceBefore = 60;  // 正文下空二行，每行约30pt
          console.log('[apply-format] 命令格式签发人段前间距已设置: 60磅');
        }
      } catch (e) {
        console.warn('[apply-format] 命令格式签发人段前间距设置失败', e);
      }
    }

    // 附件说明续行：强制设置缩进并删除开头的空格
    if (type === 'fujianContinue') {
      try {
        var range = para.Range;
        var originalText = range.Text;
        if (originalText) {
          // 删除开头的空格
          var cleanedText = originalText.replace(/^[\s\u3000]+/, '');
          if (cleanedText !== originalText) {
            range.Text = cleanedText;
          }
        }
        // 强制设置缩进（63磅 = 1260缇 = 6字符）
        var pFmt = range.ParagraphFormat || para.Format;
        if (pFmt) {
          pFmt.FirstLineIndent = 63;
          pFmt.Alignment = 0;
        }
      } catch (e) {}
    }

    // ========================================
    // 纪要格式特殊处理：出席/请假/列席人员
    // 国标10.3："出席"用3号黑体字，冒号后用3号仿宋体字，左空二字
    // 国标10.3：在正文或附件说明下空一行编排
    // ========================================
    if (type === 'chuxi' || type === 'qingjia' || type === 'liexi') {
      try {
        var range = para.Range;
        var fullText = range.Text.replace(/[\r\n]/g, '').trim();

        // 匹配"出席："、"请假："、"列席："后的内容
        var labelMatch = fullText.match(/^(出席|请假|列席)[：:](.*)$/);
        if (labelMatch) {
          var labelText = labelMatch[1] + '：';  // "出席："
          var contentText = labelMatch[2];  // 人员名单

          console.log('[apply-format] 处理' + labelText + '内容: ' + contentText.substring(0, 20));

          // Step 1: 设置段落格式（左空二字，段前空一行）
          var pFmt = para.Range.ParagraphFormat || para.Format;
          if (pFmt) {
            pFmt.LeftIndent = 21;  // 21磅 ≈ 2字符宽度（修正：原31磅偏大）
            pFmt.FirstLineIndent = 0;
            pFmt.Alignment = 0;
            // 国标10.3：正文下空一行编排出席人员
            // 只有"出席"（第一个）添加段前间距
            if (labelMatch[1] === '出席') {
              pFmt.SpaceBefore = 30;  // 空一行，约30pt
            }
          }

          // Step 2: 使用Selection逐字符设置字体（更可靠）
          try {
            para.Range.Select();
            var sel = Application.Selection;
            if (sel) {
              // 先设置整个段落为仿宋
              sel.Font.NameFarEast = '仿宋_GB2312';
              sel.Font.Name = '仿宋_GB2312';
              sel.Font.Size = 16;

              // 再设置标签部分为黑体
              // 移动光标到段落开头
              sel.HomeKey(5);  // wdLine = 5
              // 选中标签部分
              for (var ci = 0; ci < labelText.length; ci++) {
                sel.MoveRight(1, 1, 1);  // wdCharacter=1, Count=1, Extend=1(选中)
              }
              // 设置标签为黑体
              sel.Font.NameFarEast = '黑体';
              sel.Font.Name = '黑体';
              sel.Font.Size = 16;

              // 取消选择
              sel.Collapse(0);
              console.log('[apply-format] ' + labelText + '处理完成（Selection方法）');
            }
          } catch (selErr) {
            console.warn('[apply-format] Selection方法失败，尝试Characters方法', selErr);
            // 备选方案：使用Characters逐字符设置
            try {
              var chars = para.Range.Characters;
              // 设置整个段落为仿宋
              range.Font.Name = '仿宋_GB2312';
              range.Font.Size = 16;
              // 设置标签部分为黑体
              for (var ci = 1; ci <= labelText.length && ci <= chars.Count; ci++) {
                chars.Item(ci).Font.Name = '黑体';
                chars.Item(ci).Font.NameFarEast = '黑体';
                chars.Item(ci).Font.Size = 16;
              }
              console.log('[apply-format] ' + labelText + '处理完成（Characters方法）');
            } catch (charErr) {
              console.warn('[apply-format] Characters方法也失败', charErr);
            }
          }
        }
      } catch (e) {
        console.warn('[apply-format] 出席/请假/列席处理失败', e);
      }
    }

    // 纪要标志：确保红色，设置位置（上边缘至版心上边缘35mm）
    if (type === 'jiyaoBiaozhi') {
      try {
        // 使用Selection API更可靠地设置字体
        para.Range.Select();
        var sel = Application.Selection;
        if (sel) {
          sel.Font.NameFarEast = '方正小标宋简体';
          sel.Font.Name = '方正小标宋简体';
          sel.Font.NameAscii = '方正小标宋简体';  // 同时设置Ascii
          sel.Font.NameOther = '方正小标宋简体';  // 设置Other属性
          sel.Font.Size = 22;  // 2号字
          sel.Font.Color = 0x0000FF;  // BGR格式的红色
          sel.ParagraphFormat.Alignment = 1;  // 居中
          sel.ParagraphFormat.FirstLineIndent = 0;
          // 国标10.3：上边缘至版心上边缘35mm
          // 版心上边缘 = 天头37mm，所以距页面顶部 = 37+35 = 72mm
          if (i === 1) {
            sel.ParagraphFormat.SpaceBefore = 99;  // 35mm ≈ 99磅
          }
          sel.Collapse(0);
        } else {
          // 备选方案
          var pFmt = para.Range.ParagraphFormat || para.Format;
          if (pFmt) {
            pFmt.Alignment = 1;
            pFmt.FirstLineIndent = 0;
            if (i === 1) {
              pFmt.SpaceBefore = 99;
            }
          }
          para.Range.Font.NameFarEast = '方正小标宋简体';
          para.Range.Font.Name = '方正小标宋简体';
          para.Range.Font.NameAscii = '方正小标宋简体';
          para.Range.Font.NameOther = '方正小标宋简体';
          para.Range.Font.Size = 22;
          para.Range.Font.Color = 0x0000FF;
        }
        console.log('[apply-format] 纪要标志样式已设置: 方正小标宋简体22号红色居中');
      } catch (e) {
        console.warn('[apply-format] 纪要标志样式设置失败', e);
      }
    }

    // 纪要编号：确保居中，仿宋3号，段前空二行
    // 国标7.2.5：发文字号编排在发文机关标志下空二行位置
    if (type === 'jiyaoNumber') {
      try {
        para.Range.Select();
        var sel = Application.Selection;
        if (sel) {
          sel.Font.NameFarEast = '仿宋_GB2312';
          sel.Font.Name = '仿宋_GB2312';
          sel.Font.NameAscii = '仿宋_GB2312';  // 同时设置Ascii
          sel.Font.NameOther = '仿宋_GB2312';  // 设置Other属性
          sel.Font.Size = 16;  // 3号字
          sel.ParagraphFormat.Alignment = 1;  // 居中
          sel.ParagraphFormat.FirstLineIndent = 0;
          sel.ParagraphFormat.SpaceBefore = 60;  // 下空二行，每行约30pt
          sel.Collapse(0);
        } else {
          var pFmt = para.Range.ParagraphFormat || para.Format;
          if (pFmt) {
            pFmt.Alignment = 1;
            pFmt.FirstLineIndent = 0;
            pFmt.SpaceBefore = 60;
          }
          para.Range.Font.NameFarEast = '仿宋_GB2312';
          para.Range.Font.Name = '仿宋_GB2312';
          para.Range.Font.NameAscii = '仿宋_GB2312';
          para.Range.Font.NameOther = '仿宋_GB2312';
          para.Range.Font.Size = 16;
        }
      } catch (e) {
        console.warn('[apply-format] 纪要编号样式设置失败', e);
      }
    }

    // ========================================
    // 正文与附件说明之间添加空行（国标7.3.4）
    // ========================================
    if (type === 'fujian' && i > 1) {
      try {
        // 检查前一段是否是正文（非空、非标题、非落款等）
        var prevPara = doc.Paragraphs.Item(i - 1);
        var prevText = getParaText(prevPara);
        if (prevText && prevText.length > 0) {
          // 判断前一段是否是正文类型（非标题、非主送机关、非空行）
          var prevIsBody = !isHeading1(prevText) && !isHeading2(prevText) &&
                          !isHeading3(prevText) && !isHeading4(prevText) &&
                          !isZhushong(prevText, titleIndex, i - 1) &&
                          !isDate(prevText) && !isSignature(prevText, i - 1, totalParas, dateIndex) &&
                          !isFuzhu(prevText) && !isFujian(prevText);

          if (prevIsBody) {
            // 在正文和附件说明之间插入空行
            // 方法：将附件说明段落拆分，插入空段落
            var currentRange = para.Range;
            currentRange.InsertParagraphBefore();
            // 新插入的空段落会成为当前索引的前一段
          }
        }
      } catch (e) {
        console.warn('[apply-format] 插入附件说明前空行失败', e);
      }
    }

    // ========================================
    // 信函格式特殊处理
    // ========================================
    if (formatType === FORMAT_TYPES.LETTER) {
      try {
        var pFmt2 = para.Range.ParagraphFormat || para.Format;
        if (!pFmt2) pFmt2 = para.Format;

        // 发文字号：顶格居版心右边缘，与红线的距离为3号汉字高度的7/8
        // 国标10.1：发文字号与第一条红色双线的距离为3号汉字高度的7/8（约14pt/4.8mm）
        // 计算：双线在44mm（发文机关标志下4mm），发文机关标志结束于约40mm
        // 发文字号位置 = 44 + 4.8 = 48.8mm
        // 段前间距 = 48.8 - 40 = 8.8mm ≈ 25pt
        if (isDocNumber(text) && pFmt2) {
          pFmt2.Alignment = 2;  // 右对齐
          pFmt2.RightIndent = 0;  // 不缩进，顶格
          // 段前间距：确保发文字号在双线下方约4.8mm处
          pFmt2.SpaceBefore = 25;  // 约8.8mm（确保与双线距离正确）
        }

        // 发文机关标志：红色小标宋体字（国标10.1）
        // 信函格式发文机关标志：居中排布，上边缘至上页边为30mm，推荐使用红色小标宋体字
        if (type === 'documentFlag' && i <= 3) {
          try {
            para.Range.Select();
            var sel = Application.Selection;
            if (sel) {
              sel.Font.NameFarEast = '方正小标宋简体';
              sel.Font.Name = '方正小标宋简体';
              sel.Font.NameAscii = '方正小标宋简体';
              sel.Font.NameOther = '方正小标宋简体';
              sel.Font.Size = 22;  // 2号字
              sel.Font.Color = 0x0000FF;  // BGR格式的红色
              sel.ParagraphFormat.Alignment = 1;  // 居中
              sel.ParagraphFormat.FirstLineIndent = 0;
              sel.Collapse(0);
              console.log('[apply-format] 信函发文机关标志样式已设置: 红色小标宋体22号居中');
            } else {
              // 备选方案：直接使用Range API
              var range = para.Range;
              range.Font.NameFarEast = '方正小标宋简体';
              range.Font.Name = '方正小标宋简体';
              range.Font.NameAscii = '方正小标宋简体';
              range.Font.NameOther = '方正小标宋简体';
              range.Font.Size = 22;
              range.Font.Color = 0x0000FF;
              var pFmtDoc = range.ParagraphFormat || para.Format;
              if (pFmtDoc) {
                pFmtDoc.Alignment = 1;
                pFmtDoc.FirstLineIndent = 0;
              }
            }
          } catch (e) {
            console.warn('[apply-format] 信函发文机关标志样式设置失败', e);
          }
        }

        // 标题：居中无缩进
        if (i === titleIndex && pFmt2) {
          pFmt2.FirstLineIndent = 0;
          pFmt2.Alignment = 1;  // 居中
        }

        // 主送机关：顶格
        if (type === 'zhushong' && pFmt2) {
          pFmt2.FirstLineIndent = 0;
          pFmt2.Alignment = 0;  // 左对齐顶格
        }
      } catch (e) {}
    }

    // 统计
    appliedStyles[type] = (appliedStyles[type] || 0) + 1;
    detectedElements[type] = (detectedElements[type] || 0) + 1;
  }

  // ========================================
  // 特殊处理：落款与成文日期字数比较调整
  // 国标7.3.5.2：署名右空二字，日期首字比署名首字右移二字
  // 如成文日期长于发文机关署名，应当使成文日期右空二字编排，并相应增加发文机关署名右空字数
  // ========================================
  try {
    // 从文档末尾向前搜索，找到最后的日期和署名段落
    var signatureParaIndex = -1;
    var dateParaIndex = -1;
    var signatureText = '';
    var dateText = '';
    var totalParasForSig = doc.Paragraphs.Count;

    // 从后向前搜索，找到最后的成文日期（排除印发行）
    for (var i = totalParasForSig; i >= 1; i--) {
      var t = getParaText(doc.Paragraphs.Item(i));
      // 跳过版记区域
      if (isChaosong(t) || isYinfa(t)) continue;
      if (isDate(t) && dateParaIndex < 0) {
        dateParaIndex = i;
        dateText = t;
        break;  // 找到最后一个日期就停止
      }
    }

    // 在日期前后2个段落内查找署名
    if (dateParaIndex > 0) {
      for (var i = Math.max(1, dateParaIndex - 2); i <= Math.min(totalParasForSig, dateParaIndex + 2); i++) {
        if (i === dateParaIndex) continue;  // 跳过日期段落本身
        var t = getParaText(doc.Paragraphs.Item(i));
        // 署名：非日期、非附注、非空、长度≤30、不以标点结尾
        if (!isDate(t) && !isFuzhu(t) && t.length > 0 && t.length <= 30 && !/[，。！？、；：]$/.test(t)) {
          // 优先选择日期前一个段落作为署名
          if (i === dateParaIndex - 1 || signatureParaIndex < 0) {
            signatureParaIndex = i;
            signatureText = t;
          }
        }
      }
    }

    // 备选：如果仍未找到署名，在文档末尾查找政府机构名称
    if (signatureParaIndex < 0) {
      for (var i = totalParasForSig; i >= Math.max(1, totalParasForSig - 10); i--) {
        var t = getParaText(doc.Paragraphs.Item(i));
        // 查找政府机构名称：含"政府"、"办公室"等且不以标点结尾
        if (t.length > 0 && t.length <= 30 && !isDate(t) && !isChaosong(t) && !isYinfa(t) &&
            /政府$|办公室$|委员会$|厅$|局$/.test(t) && !/[，。！？、；：]$/.test(t)) {
          signatureParaIndex = i;
          signatureText = t;
          console.log('[apply-format] 备选署名检测: 找到"' + t + '"在段落' + i);
          break;
        }
      }
    }

    // 字数比较处理
    // 国标7.3.5.2：不加盖印章时
    // - 署名右空二字编排
    // - 日期首字比署名首字右移二字
    // 推导：
    // - 日期右空 = 2字（固定）
    // - 署名右空 = 2字 + (日期长度 - 署名长度)
    if (signatureParaIndex > 0 && dateParaIndex > 0) {
      var sigLen = signatureText.length;
      var dateLen = dateText.length;

      var sigPara = doc.Paragraphs.Item(signatureParaIndex);
      var datePara = doc.Paragraphs.Item(dateParaIndex);

      // 计算右空值（单位：磅，1字符 = 10.5磅）
      // 国标7.3.5.2：日期首字比署名首字右移二字 → 日期右空固定4字
      var dateRightPt = 42;  // 日期右空固定4字 = 42磅
      var sigRightPt;  // 署名右空

      var lenDiff = dateLen - sigLen;  // 日期长度 - 署名长度
      // 正确公式：署名右空 = 2字 + lenDiff
      sigRightPt = 21 + 10.5 * lenDiff;  // 2字 + 差值
      sigRightPt = Math.max(21, sigRightPt);  // 最小2字

      // 输出调试信息
      console.log('[apply-format] 署名检测: 署名="' + signatureText + '"(长度' + sigLen + '), 日期="' + dateText + '"(长度' + dateLen + '), 署名右空=' + sigRightPt + '磅');

      // 使用Selection设置，更可靠
      try {
        sigPara.Range.Select();
        var sel = Application.Selection;
        if (sel && sel.ParagraphFormat) {
          sel.ParagraphFormat.RightIndent = sigRightPt;
          sel.ParagraphFormat.Alignment = 2;  // 右对齐
          console.log('[apply-format] 使用Selection设置署名右缩进: ' + sigRightPt + '磅');
        }
      } catch (e) {
        // 备选方案
        var pFmt = sigPara.Range.ParagraphFormat || sigPara.Format;
        if (pFmt) {
          pFmt.RightIndent = sigRightPt;
          pFmt.Alignment = 2;
          console.log('[apply-format] 署名右缩进已设置: ' + sigRightPt + '磅');
        }
      }

      try {
        datePara.Range.Select();
        var sel2 = Application.Selection;
        if (sel2 && sel2.ParagraphFormat) {
          sel2.ParagraphFormat.RightIndent = dateRightPt;
          sel2.ParagraphFormat.Alignment = 2;
        }
      } catch (e) {
        var pFmt2 = datePara.Range.ParagraphFormat || datePara.Format;
        if (pFmt2) {
          pFmt2.RightIndent = dateRightPt;
          pFmt2.Alignment = 2;
        }
      }
    } else {
      console.warn('[apply-format] 署名检测失败: signatureParaIndex=' + signatureParaIndex + ', dateParaIndex=' + dateParaIndex);
    }
  } catch (e) {
    console.warn('[apply-format] 落款字数调整失败', e);
  }

  // ========================================
  // 特殊处理：上行文签发人与发文字号同行
  // 国标7.2.5：上行文的发文字号居左空一字编排，与最后一个签发人姓名处在同一行
  // 国标7.2.6：签发人居右空一字编排
  // ========================================
  var qianfarenMerged = false;  // 标记是否已合并签发人

  try {
    var qianfarenParaIndex = -1;
    var qianfarenText = '';
    var docNumberParaIndex = -1;
    var docNumberText = '';

    for (var i = 1; i <= Math.min(15, totalParas); i++) {
      var t = getParaText(doc.Paragraphs.Item(i));
      if (isQianfaren(t)) {
        qianfarenParaIndex = i;
        qianfarenText = t;
      }
      if (isDocNumber(t)) {
        docNumberParaIndex = i;
        docNumberText = t;
      }
    }

    // 如果有签发人（上行文），需要将发文字号和签发人放在同一行
    if (qianfarenParaIndex > 0 && docNumberParaIndex > 0) {
      var docNumPara = doc.Paragraphs.Item(docNumberParaIndex);
      var qianfarenPara = doc.Paragraphs.Item(qianfarenParaIndex);

      // 提取签发人姓名
      var nameMatch = qianfarenText.match(/^签发人[：:](.+)$/);
      var qianfarenName = nameMatch ? nameMatch[1].trim() : '';

      // 先删除签发人段落
      try {
        qianfarenPara.Range.Delete();
      } catch (e) {}

      // 清空发文字号段落
      var numRange = docNumPara.Range;
      numRange.Delete();

      // 分别插入发文字号和签发人（用制表符分隔）
      // 发文字号部分 - 仿宋
      numRange.InsertAfter(docNumberText);
      numRange.Font.NameFarEast = '仿宋_GB2312';
      numRange.Font.Name = '仿宋_GB2312';
      numRange.Font.Size = 16;

      // 插入制表符
      numRange.Collapse(0);  // 移到末尾
      numRange.InsertAfter('\t');

      // 签发人标签部分 - 仿宋
      numRange.Collapse(0);
      numRange.InsertAfter('签发人：');
      numRange.Font.NameFarEast = '仿宋_GB2312';
      numRange.Font.Name = '仿宋_GB2312';
      numRange.Font.Size = 16;

      // 签发人姓名部分 - 楷体
      numRange.Collapse(0);
      numRange.InsertAfter(qianfarenName);
      numRange.Font.NameFarEast = '楷体_GB2312';
      numRange.Font.Name = '楷体_GB2312';
      numRange.Font.Size = 16;

      // 设置段落格式：左空一字
      var pFmt = docNumPara.Range.ParagraphFormat;
      if (pFmt) {
        pFmt.Alignment = 3;  // 两端对齐（配合制表符实现左右分布）
        pFmt.FirstLineIndent = 0;
        pFmt.LeftIndent = 10.5;  // 左空一字（10.5磅=210缇=1字符）
        pFmt.RightIndent = 10.5;  // 右空一字（签发人居右空一字）
      }

      // 设置制表符：右对齐到版心右边缘（减去右空一字）
      // 版心宽度 = 156mm = 442磅
      // 右对齐制表位 = 442 - 10.5(右空一字) = 431.5磅
      // ========================================
      // 关键修复：直接添加制表位，不使用RemoveAll（WPS不支持）
      try {
        docNumPara.Range.Select();
        var sel = Application.Selection;
        if (sel && sel.ParagraphFormat) {
          var selTabs = sel.ParagraphFormat.TabStops;
          if (selTabs) {
            // 直接添加右对齐制表位（WPS会自动处理重复）
            selTabs.Add(431.5, 2, 0);  // position, alignment, leader
            console.log('[apply-format] 使用Selection设置制表位成功');
          }
        }
      } catch (e) {
        console.warn('[apply-format] Selection制表位设置失败', e);
        // 备选方案：尝试直接用段落制表位
        try {
          var tabStops = docNumPara.Range.ParagraphFormat.TabStops;
          if (tabStops) {
            tabStops.Add(431.5, 2, 0);
            console.log('[apply-format] 使用段落制表位成功');
          }
        } catch (e2) {
          console.warn('[apply-format] 所有制表位设置方法失败', e2);
        }
      }

      qianfarenMerged = true;
      appliedStyles['qianfaren_combined'] = 1;
    }
  } catch (e) {
    console.warn('[apply-format] 签发人同行处理失败', e);
  }

  // ========================================
  // 特殊处理：签发人姓名用楷体（仅处理未被合并的情况）
  // ========================================
  if (!qianfarenMerged) {
    try {
      for (var i = 1; i <= doc.Paragraphs.Count; i++) {
        var para = doc.Paragraphs.Item(i);
        var text = getParaText(para);

        // 只处理独立的签发人段落（整段只有签发人内容）
        if (/^签发人[：:]/.test(text) && !/\t/.test(text)) {
          var match = text.match(/^签发人[：:](.+)$/);
          if (match && match[1]) {
            var range = para.Range;
            // "签发人："是4个字符（签、发、人、：），姓名从第5个字符开始
            // 使用SetRange选择姓名部分（0-based，所以第5个字符是索引4）
            try {
              var nameRange = para.Range;
              nameRange.SetRange(4, range.Characters.Count);  // 从索引4到末尾
              nameRange.Font.NameFarEast = '楷体_GB2312';
              nameRange.Font.Name = '楷体_GB2312';
            } catch (e) {
              // 后备方案：逐字符设置
              var chars = range.Characters;
              for (var c = 5; c <= chars.Count; c++) {
                try {
                  chars.Item(c).Font.NameFarEast = '楷体_GB2312';
                  chars.Item(c).Font.Name = '楷体_GB2312';
                } catch (e2) {}
              }
            }
          }
        }
      }
    } catch (e) {
      console.warn('[apply-format] 签发人姓名字体设置失败', e);
    }
  }

  // ========================================
  // 特殊处理：创建版记表格
  // 参考REDformat的简洁实现
  // 确保署名→日期→版记的正确顺序
  // ========================================
  try {
    // 信函格式：清理文档顶部区域可能存在的旧红色线条Shape
    // 注意：只删除顶部区域（发文机关标志下方）的线条，保留底部区域
    // 这些Shape可能是之前实现遗留的
    if (formatType === FORMAT_TYPES.LETTER) {
      try {
        var shapes = doc.Shapes;
        if (shapes && shapes.Count > 0) {
          // 从后往前删除，避免索引问题
          for (var si = shapes.Count; si >= 1; si--) {
            try {
              var shape = shapes.Item(si);
              // 检查是否是红色线条（Line类型、红色）
              if (shape.Type === 9) {  // msoLine = 9
                var lineColor = shape.Line.ForeColor.RGB;
                if (lineColor === 0x0000FF) {  // 红色（BGR格式）
                  // 只删除顶部区域的线条（Y坐标小于100mm，即发文机关标志附近）
                  // 顶部区域：Y < 100mm ≈ 283.5磅
                  try {
                    var topPos = shape.Top;  // 获取线条的顶部位置（磅）
                    if (topPos < 283.5) {  // 只删除顶部区域
                      shape.Delete();
                    }
                  } catch (posErr) {
                    // 如果无法获取位置，保守起见不删除
                  }
                }
              }
            } catch (e) {}
          }
        }
      } catch (e) {
        console.warn('[apply-format] 清理多余Shape失败', e);
      }
    }

    // 重新获取最新段落数（前面可能有插入/删除操作）
    var currentTotalParas = doc.Paragraphs.Count;

    // ========================================
    // 纪要格式特殊处理：不执行版记删除
    // 纪要格式的"送："不删除原内容，只添加格式
    // ========================================
    var skipBanjiDeletion = false;
    if (formatType === FORMAT_TYPES.MINUTES) {
      // 纪要格式：只收集信息，不删除内容
      skipBanjiDeletion = true;
      console.log('[apply-format] 纪要格式：跳过版记删除');
    }

    if ((chaosongIndex > 0 || yinfaIndex > 0) && !skipBanjiDeletion) {
      // 收集版记内容（使用预先保存的索引）
      var chaosongText = '';
      var yinfaOrg = '';
      var yinfaDate = '';
      var yinfaFullText = '';  // 保存原始完整文本作为后备
      var dateTextFound = '';
      var signatureTextFound = '';
      var jiyaoBiaozhiText = '';  // 纪要标志文本（用于提取印发机关）

      for (var i = 1; i <= currentTotalParas; i++) {
        var t = getParaText(doc.Paragraphs.Item(i));

        // 收集纪要标志文本（用于提取印发机关名）
        var alignment = doc.Paragraphs.Item(i).Format.Alignment;
        if (isJiyaoBiaozhi(t, i, alignment)) {
          jiyaoBiaozhiText = t;
        }

        if (/^\s*\d{4}年\d{1,2}月\d{1,2}日[\s\u3000]*(抄送|送)[：:]/.test(t)) {
          if (!dateTextFound) {
            dateTextFound = extractLeadingDate(t);
          }
          chaosongText = normalizeChaosongText(t);
        } else if (isChaosong(t) || isJiyaoBanji(t)) {
          // 提取抄送内容：去除"抄送："或"送："前缀、前后空格、结尾句号
          // 最终只保留机关名称列表，填充时再添加标准格式
          chaosongText = normalizeChaosongText(t);
        } else if (isYinfa(t)) {
          yinfaFullText = t;  // 先保存原始文本
          // 改进的解析逻辑：兼容多种格式
          // 格式1: "XX办公室  2024年3月15日印发"（空格分隔）
          // 格式2: "XX办公室                    2024年3月15日印发"（多空格）
          // 格式3: "XX市人民政府办公室 2024年3月15日印发"

          // 先提取日期
          var dateMatch = t.match(/(\d{4}年\d{1,2}月\d{1,2}日)/);
          if (dateMatch) {
            yinfaDate = dateMatch[1];
          }

          // 再提取机关名（日期之前的内容）
          var orgMatch = t.match(/^(.+?)(?=\s*\d{4}年)/);
          if (orgMatch) {
            yinfaOrg = orgMatch[1].trim();
          }

          // 如果无法解析机关名，使用后备方案
          if (!yinfaOrg) {
            yinfaOrg = t.replace(/\s*\d{4}年\d{1,2}月\d{1,2}日.*$/, '').trim();
          }
        } else if (isDate(t) && !dateTextFound) {
          dateTextFound = t;
        }
        // 收集署名文本
        if (signatureIndex > 0 && i === signatureIndex) {
          signatureTextFound = t;
        }
      }

      // ========================================
      // 纪要格式特殊处理：如果未检测到印发机关，从纪要标志提取
      // ========================================
      if (formatType === FORMAT_TYPES.MINUTES) {
        if (!yinfaOrg && jiyaoBiaozhiText) {
          // 从"XX市人民政府常务会议纪要"提取"XX市人民政府办公室"
          // 改进正则：处理"第X次"等序号
          var orgMatch = jiyaoBiaozhiText.match(/^(.+?)(第\d+次)?(常务会议|会议)?纪要$/);
          if (orgMatch) {
            yinfaOrg = orgMatch[1] + '办公室';
            console.log('[apply-format] 从纪要标志提取印发机关: ' + yinfaOrg);
          }
        }
        // 如果未检测到印发日期，使用当前日期或成文日期
        if (!yinfaDate && dateTextFound) {
          yinfaDate = dateTextFound;
        }
      }

      // 使用预先保存的署名索引
      var signatureIdx = signatureIndex;

      // 获取版记起始位置
      var banjiStartIdx = chaosongIndex > 0 ? chaosongIndex : yinfaIndex;

      // ========================================
      // 关键修复：在删除前先保存署名段落内容
      // ========================================
      var signatureTextSaved = '';
      if (signatureIdx > 0 && signatureIdx <= currentTotalParas) {
        try {
          signatureTextSaved = getParaText(doc.Paragraphs.Item(signatureIdx));
        } catch (e) {}
      }

      // 从后往前删除版记段落和日期（保留署名）
      var deleteStartIdx = banjiStartIdx;

      // 安全删除：每次循环重新获取段落数
      while (doc.Paragraphs.Count >= deleteStartIdx) {
        var currentCount = doc.Paragraphs.Count;
        var deletedSomething = false;

        for (var i = currentCount; i >= deleteStartIdx; i--) {
          // 保留署名段落
          if (signatureIdx > 0 && i === signatureIdx) continue;
          try {
            doc.Paragraphs.Item(i).Range.Delete();
            deletedSomething = true;
            break;  // 每次只删除一个，然后重新获取count
          } catch (e) {}
        }

        // 如果没有删除任何内容，退出循环（避免无限循环）
        if (!deletedSomething) break;
      }

      // 注意：不在这里删除署名后的日期，因为signatureIdx可能已失效
      // 删除操作应该在重新定位署名后再进行

      // ========================================
      // 删除后重新定位：找到文档末尾的署名段落
      // ========================================
      var newTotalParas = doc.Paragraphs.Count;
      var actualSignatureIdx = -1;

      // 在文档末尾查找署名段落
      for (var i = newTotalParas; i >= Math.max(1, newTotalParas - 5); i--) {
        var t = getParaText(doc.Paragraphs.Item(i));
        if (t && (isSignature(t, i, newTotalParas, -1) || t === signatureTextSaved)) {
          actualSignatureIdx = i;
          break;
        }
      }

      // 如果没找到署名，使用文档最后一个段落
      if (actualSignatureIdx < 0) {
        actualSignatureIdx = newTotalParas;
      }

      // ========================================
      // 现在用正确的actualSignatureIdx删除署名后的日期
      // ========================================
      if (actualSignatureIdx > 0 && actualSignatureIdx < doc.Paragraphs.Count) {
        try {
          var nextParaIdx = actualSignatureIdx + 1;
          var nextParaText = getParaText(doc.Paragraphs.Item(nextParaIdx));
          if (isDate(nextParaText)) {
            doc.Paragraphs.Item(nextParaIdx).Range.Delete();
          }
        } catch (e) {}
      }

      // ========================================
      // Step 1: 设置署名格式
      // 注意：署名右空在Step 2根据日期长度计算
      // ========================================
      if (actualSignatureIdx > 0 && actualSignatureIdx <= doc.Paragraphs.Count) {
        try {
          var sigPara = doc.Paragraphs.Item(actualSignatureIdx);
          var sigRange = sigPara.Range;
          sigRange.Font.NameFarEast = '仿宋_GB2312';
          sigRange.Font.Name = '仿宋_GB2312';
          sigRange.Font.Size = 16;
          try {
            var pFmt = sigRange.ParagraphFormat;
            if (pFmt) {
              pFmt.Alignment = 2;
              // 右空在Step 2计算后设置
              pFmt.FirstLineIndent = 0;
            }
          } catch (e) {}
        } catch (e) {}
      }

      // ========================================
      // Step 2: 在署名后插入日期（仅当署名后没有日期时）
      // 国标7.3.5.2：日期首字比署名首字右移二字
      // 如成文日期长于发文机关署名，应当使成文日期右空二字编排，并相应增加发文机关署名右空字数
      // ========================================
      // 先检查署名后是否已有日期
      var hasDateAfterSignature = false;
      if (actualSignatureIdx > 0 && actualSignatureIdx < doc.Paragraphs.Count) {
        try {
          var nextText = getParaText(doc.Paragraphs.Item(actualSignatureIdx + 1));
          if (isDate(nextText) || !!extractLeadingDate(nextText)) {
            hasDateAfterSignature = true;
          }
        } catch (e) {}
      }

      // 获取署名文本和长度
      var sigTextForCalc = '';
      var sigLenForCalc = 0;
      if (actualSignatureIdx > 0) {
        sigTextForCalc = getParaText(doc.Paragraphs.Item(actualSignatureIdx));
        sigLenForCalc = sigTextForCalc.length;
      }
      var dateLenForCalc = dateTextFound ? dateTextFound.length : 0;

      // 调试输出
      console.log('[apply-format] 版记部分署名计算: dateTextFound="' + dateTextFound + '"(长度' + dateLenForCalc + '), sigTextForCalc="' + sigTextForCalc + '"(长度' + sigLenForCalc + ')');

      // 计算署名和日期的右空
      // 国标7.3.5.2：不加盖印章时
      // - 署名右空二字编排
      // - 日期首字比署名首字右移二字 → 日期右空固定4字
      // 推导：
      // - 日期右空 = 4字（固定）
      // - 署名右空 = 2字 + (日期长度 - 署名长度)
      // 单位：1字符 = 10.5磅
      var dateRightIndentPt = 42;  // 日期右空固定4字 = 42磅
      var sigRightIndentPt;  // 署名右空

      // 如果dateLenForCalc为0，使用默认值9（日期通常约9字）
      if (dateLenForCalc === 0) {
        dateLenForCalc = 9;  // 默认日期长度
        console.warn('[apply-format] dateTextFound为空，使用默认长度9');
      }

      var lenDiff = dateLenForCalc - sigLenForCalc;  // 日期长度 - 署名长度
      // 正确公式：署名右空 = 2字 + lenDiff
      sigRightIndentPt = 21 + 10.5 * lenDiff;  // 2字 + 差值
      sigRightIndentPt = Math.max(21, sigRightIndentPt);  // 最小2字

      console.log('[apply-format] 版记部分署名右空: lenDiff=' + lenDiff + ', sigRightIndentPt=' + sigRightIndentPt + '磅');

      // 更新署名右空
      if (actualSignatureIdx > 0) {
        try {
          var sigParaToUpdate = doc.Paragraphs.Item(actualSignatureIdx);
          // 使用Selection设置，更可靠
          sigParaToUpdate.Range.Select();
          var sel = Application.Selection;
          if (sel && sel.ParagraphFormat) {
            sel.ParagraphFormat.RightIndent = sigRightIndentPt;
            sel.ParagraphFormat.Alignment = 2;
            console.log('[apply-format] 版记部分署名右缩进已设置(Selection): ' + sigRightIndentPt + '磅');
          } else {
            // 备选方案
            var pFmt = sigParaToUpdate.Range.ParagraphFormat || sigParaToUpdate.Format;
            if (pFmt) {
              pFmt.RightIndent = sigRightIndentPt;
              pFmt.Alignment = 2;
              console.log('[apply-format] 版记部分署名右缩进已设置: ' + sigRightIndentPt + '磅');
            }
          }
        } catch (e) {
          console.warn('[apply-format] 版记部分署名右缩进设置失败', e);
        }
      }

      // 只有当没有日期时才插入
      if (actualSignatureIdx > 0 && dateTextFound && !hasDateAfterSignature) {
        try {
          var sigPara2 = doc.Paragraphs.Item(actualSignatureIdx);
          var sigRange2 = sigPara2.Range;
          sigRange2.Collapse(0);
          sigRange2.InsertParagraphAfter();

          var datePara = doc.Paragraphs.Item(actualSignatureIdx + 1);
          var dateRange = datePara.Range;
          setParaTextPreserveMark(datePara, dateTextFound);
          dateRange.Font.NameFarEast = '仿宋_GB2312';
          dateRange.Font.Name = '仿宋_GB2312';
          dateRange.Font.Size = 16;
          try {
            var pFmt = dateRange.ParagraphFormat;
            if (pFmt) {
              pFmt.Alignment = 2;
              pFmt.RightIndent = dateRightIndentPt;
              pFmt.FirstLineIndent = 0;
            }
          } catch (e) {}

        } catch (e) {
          console.warn('[apply-format] 插入日期失败', e);
        }
      } else if (hasDateAfterSignature) {
        // 如果已有日期，直接更新该日期段落格式，不改动署名索引
        try {
          var existingDatePara = doc.Paragraphs.Item(actualSignatureIdx + 1);
          var existingDateText = getParaText(existingDatePara);
          var normalizedDateText = extractLeadingDate(existingDateText) || dateTextFound || existingDateText;
          if (normalizedDateText && existingDateText !== normalizedDateText) {
            setParaTextPreserveMark(existingDatePara, normalizedDateText);
          }
          var existingDateRange = existingDatePara.Range;
          existingDateRange.Font.NameFarEast = '仿宋_GB2312';
          existingDateRange.Font.Name = '仿宋_GB2312';
          existingDateRange.Font.Size = 16;
          try {
            var pFmt = existingDateRange.ParagraphFormat;
            if (pFmt) {
              pFmt.Alignment = 2;
              pFmt.RightIndent = dateRightIndentPt;
              pFmt.FirstLineIndent = 0;
            }
          } catch (e) {}
        } catch (e) {}
      }

      // ========================================
      // Step 3: 在文档末尾插入版记
      // ========================================
      var finalTotalParas = doc.Paragraphs.Count;

      // 在最后一个段落后插入版记
      try {
        var lastPara = doc.Paragraphs.Item(finalTotalParas);
        var lastRange = lastPara.Range;
        lastRange.Collapse(0);  // 折叠到末尾
        lastRange.InsertParagraphAfter();

        var banjiPara = doc.Paragraphs.Item(finalTotalParas + 1);
        var banjiRange = banjiPara.Range;

      // ========================================
      // 信函格式：只保留抄送机关，不加印发机关和分隔线
      // 国标10.1："版记不加印发机关和印发日期、分隔线，位于公文最后一面版心内最下方"
      // 第二条红色双线距下页边20mm，版记在其上方
      // ========================================
      if (formatType === FORMAT_TYPES.LETTER) {
        // ========================================
        // 信函版记实现方案：
        // 1. 先删除所有版记相关段落（抄送、印发等）
        // 2. 使用文本框（Shape）在页面底部固定位置放置版记内容
        // 国标10.1："版记不加印发机关和印发日期、分隔线，位于公文最后一面版心内最下方"
        // ========================================

        // Step 1: 删除所有版记段落（从文档末尾开始删除抄送/印发段落）
        try {
          var letterTotalParas = doc.Paragraphs.Count;
          for (var li = letterTotalParas; li >= 1; li--) {
            var lt = getParaText(doc.Paragraphs.Item(li));
            if (isChaosong(lt) || isYinfa(lt) || /^\s*抄送/.test(lt) || /印发/.test(lt)) {
              try {
                doc.Paragraphs.Item(li).Range.Delete();
                console.log('[apply-format] 信函格式：已删除版记段落 [' + lt.substring(0, 20) + ']');
              } catch (e) {}
            }
          }
        } catch (e) {
          console.warn('[apply-format] 信函格式删除版记段落失败', e);
        }

        // Step 2: 删除新插入的空段落内容
        try {
          banjiRange.Text = '';
        } catch (e) {}

        // Step 3: 添加版记文本框（页面底部固定位置）
        // 国标10.1：第二条红色双线距下页边20mm = 277mm
        // 国标要求：文字与双线的距离为3号汉字高度的7/8 ≈ 4.9mm
        // 抄送内容可能需要2行（版心156mm，4号字约5mm/字，一行约30字）
        // 文字行高度（2行）：约14mm
        // 正确计算：文本框底部 = 277 - 4.9 ≈ 272mm，文本框顶部 = 272 - 14 ≈ 258mm
        // 保守设置：文本框顶部255mm，高度20mm，底部275mm
        try {
          var shapes = doc.Shapes;
          if (shapes && chaosongText) {
            // 版记文本框位置（修正：往上移动，避免与双线重叠）
            var banjiY = 255 * 2.835;  // 距页面顶部255mm（修正：原265mm导致文字溢出）
            var banjiX = 28 * 2.835;   // 左页边距
            var banjiWidth = 156 * 2.835;  // 版心宽度
            var banjiHeight = 20 * 2.835;  // 高度20mm，可容纳2-3行文字

            // 添加文本框（锚定在文档末尾段落）
            var textBox = shapes.AddTextbox(1, banjiX, banjiY, banjiWidth, banjiHeight, banjiRange);
            textBox.TextFrame.MarginLeft = 10.5;  // 左空一字
            textBox.TextFrame.MarginRight = 10.5; // 右空一字
            textBox.TextFrame.MarginTop = 0;
            textBox.TextFrame.MarginBottom = 0;
            textBox.TextFrame.WordWrap = true;

            // 设置文本框格式：无边框、无填充
            textBox.Line.Visible = false;
            textBox.Fill.Visible = false;

            // 设置文本内容（chaosongText已不含"抄送："前缀和结尾句号）
            var tf = textBox.TextFrame;
            if (tf && tf.TextRange) {
              tf.TextRange.Text = '抄送：' + chaosongText + '。';
              tf.TextRange.Font.NameFarEast = '仿宋_GB2312';
              tf.TextRange.Font.Name = '仿宋_GB2312';
              tf.TextRange.Font.Size = 14;  // 4号字
              tf.TextRange.ParagraphFormat.Alignment = 0;  // 左对齐
            }

            // 设置相对于页面定位（关键：确保文本框固定在页面位置）
            try {
              textBox.RelativeVerticalPosition = 1;  // wdRelativeVerticalPositionPage = 1
              textBox.Top = banjiY;
              textBox.RelativeHorizontalPosition = 1;  // wdRelativeHorizontalPositionPage = 1
              textBox.Left = banjiX;
            } catch (e) {}

            console.log('[apply-format] 信函版记文本框已添加（页面定位Y=265mm）');
          } else {
            console.warn('[apply-format] 信函格式：shapes或chaosongText为空，无法添加文本框');
          }
        } catch (e) {
          console.warn('[apply-format] 添加版记文本框失败', e);
        }

        // ========================================
        // Step 2: 添加第二条红色双线（使用页面绝对定位）
        // 国标10.1：距下页边20mm处，上细下粗，线长170mm
        // ========================================
        try {
          var lineWidth = 170 * 2.835;  // 170mm转换为磅
          var gapPoints = 2.5;  // 双线间距
          var centerX = (210 * 2.835) / 2;  // A4宽度居中
          var lineStartX = centerX - (lineWidth / 2);

          // Y坐标：距下页边20mm = 距页面顶部277mm
          var bottomDoubleLineY = 277 * 2.835;

          var shapes = doc.Shapes;
          if (shapes) {
            // 上线（细线）- 上细下粗双线的上面那条
            var line1 = shapes.AddLine(lineStartX, bottomDoubleLineY, lineStartX + lineWidth, bottomDoubleLineY, banjiRange);
            line1.Line.ForeColor.RGB = 0x0000FF;  // 红色（BGR）
            line1.Line.Weight = 0.75;  // 细线
            line1.WrapFormat.Type = 3;
            // 清除阴影（国标无阴影要求）
            try { line1.Shadow.Visible = false; } catch (e) {}
            try {
              line1.RelativeVerticalPosition = 1;
              line1.Top = bottomDoubleLineY;
              line1.RelativeHorizontalPosition = 1;
              line1.Left = lineStartX;
            } catch (e) {}

            // 下线（粗线）
            var line2 = shapes.AddLine(lineStartX, bottomDoubleLineY + gapPoints, lineStartX + lineWidth, bottomDoubleLineY + gapPoints, banjiRange);
            line2.Line.ForeColor.RGB = 0x0000FF;
            line2.Line.Weight = 1.5;  // 粗线
            line2.WrapFormat.Type = 3;
            // 清除阴影（国标无阴影要求）
            try { line2.Shadow.Visible = false; } catch (e) {}
            try {
              line2.RelativeVerticalPosition = 1;
              line2.Top = bottomDoubleLineY + gapPoints;
              line2.RelativeHorizontalPosition = 1;
              line2.Left = lineStartX;
            } catch (e) {}

            console.log('[apply-format] 信函第二条双线已添加（Y=277mm，无阴影）');
          }
        } catch (e) {
          console.warn('[apply-format] 添加第二条双线失败', e);
        }

        appliedStyles['banji_letter'] = 1;
      } else {
        // ========================================
        // 通用公文：版记固定在最后一页版心底部
        // ========================================
        try {
          setParaTextPreserveMark(banjiPara, '');
        } catch (e) {}

        var yinfaContent = '';
        if (chaosongText) {
          if (yinfaOrg && yinfaDate) {
            yinfaContent = yinfaOrg + '\t' + yinfaDate + '印发';
          } else if (yinfaFullText) {
            yinfaContent = yinfaFullText.replace(/^[\s\u3000]+/, '').replace(/[\s\u3000]+$/, '').replace(/\s+/g, '\t').replace(/印发/, '') + '印发';
          } else if (yinfaOrg) {
            yinfaContent = yinfaOrg;  // 有机关无日期
          } else if (yinfaDate) {
            yinfaContent = '\t' + yinfaDate + '印发';  // 有日期无机关
          } else {
            yinfaContent = '';
          }
        }

        console.log('[apply-format] 印发内容: ' + yinfaContent);
        createBottomBanjiLayout(banjiRange, chaosongText, yinfaContent);

        // 清理表格前残留的独立抄送段落，避免出现“普通段落抄送 + 表格版记”重复
        if (chaosongText) {
          try {
            for (var cleanupIdx = doc.Paragraphs.Count; cleanupIdx >= Math.max(actualSignatureIdx + 2, 1); cleanupIdx--) {
              var cleanupText = getParaText(doc.Paragraphs.Item(cleanupIdx));
              if (/^\s*抄送[：:]/.test(cleanupText) && normalizeChaosongText(cleanupText) === chaosongText) {
                doc.Paragraphs.Item(cleanupIdx).Range.Delete();
              }
            }
          } catch (e) {
            console.warn('[apply-format] 清理重复抄送段落失败', e);
          }
        }

        appliedStyles['banji_bottom_fixed'] = 1;
      } // end if-else (信函/通用公文)
      } catch (e) {
        console.warn('[apply-format] 创建版记失败', e);
      }

      // ========================================
      // 清理重复日期：确保文档中只有一个成文日期
      // 纪要格式跳过此步骤，因为会议介绍中可能含日期
      // ========================================
      if (formatType !== FORMAT_TYPES.MINUTES) {
        try {
          var dateParasFound = [];
          for (var i = 1; i <= doc.Paragraphs.Count; i++) {
            var t = getParaText(doc.Paragraphs.Item(i));
            if (isDate(t) && !isYinfa(t)) {  // 是日期但不是印发行
              dateParasFound.push({ index: i, text: t });
            }
          }
          // 如果有多个日期，只保留最后一个（署名后的那个）
          if (dateParasFound.length > 1) {
            for (var j = 0; j < dateParasFound.length - 1; j++) {
              try {
                // 从前往后删除，索引会变化，需要调整
                var idxToDelete = dateParasFound[j].index - j;
                doc.Paragraphs.Item(idxToDelete).Range.Delete();
              } catch (e) {}
            }
          }
        } catch (e) {
          console.warn('[apply-format] 清理重复日期失败', e);
        }
      }

    } // end if chaosongIndex or yinfaIndex

    // ========================================
    // 纪要格式特殊处理：创建版记表格（含分隔线、印发信息）
    // 国标7.4：版记应有分隔线、印发机关和印发日期
    // 国标7.4.2：抄送机关用4号仿宋，左右各空一字
    // ========================================
    if (skipBanjiDeletion && formatType === FORMAT_TYPES.MINUTES) {
      try {
        console.log('[apply-format] 纪要格式：创建底部定位版记');

        // 收集版记信息
        var jiyaoChaosongText = '';
        var jiyaoYinfaOrg = '';
        var jiyaoYinfaDate = '';

        for (var i = 1; i <= doc.Paragraphs.Count; i++) {
          var t = getParaText(doc.Paragraphs.Item(i));
          if (isJiyaoBanji(t)) {
            // "送："转换为"抄送："，移除结尾句号（后面会统一添加）
            jiyaoChaosongText = t.replace(/^送[：:]/, '').replace(/^[\s\u3000]+/, '').replace(/[\s\u3000]+$/, '').replace(/[。]+$/, '');
          }
          // 从纪要标志提取印发机关
          var alignment = doc.Paragraphs.Item(i).Format.Alignment;
          if (isJiyaoBiaozhi(t, i, alignment)) {
            // 从"XX市人民政府常务会议纪要"提取"XX市人民政府办公室"
            // 改进正则：处理"第X次"等序号，如"XX市人民政府第8次常务会议纪要"
            var orgMatch = t.match(/^(.+?)(第\d+次)?(常务会议|会议)?纪要$/);
            if (orgMatch) {
              jiyaoYinfaOrg = orgMatch[1] + '办公室';
            }
          }
          // 提取成文日期作为印发日期
          if (isDate(t) && !jiyaoYinfaDate) {
            jiyaoYinfaDate = t;
          }
        }

        // 在文档末尾创建版记锚点
        var jiyaoFinalParas = doc.Paragraphs.Count;
        var lastPara = doc.Paragraphs.Item(jiyaoFinalParas);
        var lastRange = lastPara.Range;
        lastRange.Collapse(0);
        lastRange.InsertParagraphAfter();
        var jiyaoBanjiPara = doc.Paragraphs.Item(jiyaoFinalParas + 1);
        var jiyaoBanjiRange = jiyaoBanjiPara.Range;
        try {
          setParaTextPreserveMark(jiyaoBanjiPara, '');
        } catch (e) {}

        // 印发行
        var jiyaoYinfaContent = '';
        if (jiyaoYinfaOrg && jiyaoYinfaDate) {
          jiyaoYinfaContent = jiyaoYinfaOrg + '\t' + jiyaoYinfaDate + '印发';
        } else if (jiyaoYinfaOrg) {
          jiyaoYinfaContent = jiyaoYinfaOrg;
        } else if (jiyaoYinfaDate) {
          jiyaoYinfaContent = '\t' + jiyaoYinfaDate + '印发';
        } else {
          jiyaoYinfaContent = '';
        }

        createBottomBanjiLayout(jiyaoBanjiRange, jiyaoChaosongText, jiyaoYinfaContent);

        // 删除原有的"送："段落
        for (var i = doc.Paragraphs.Count; i >= 1; i--) {
          var t = getParaText(doc.Paragraphs.Item(i));
          if (isJiyaoBanji(t)) {
            try {
              doc.Paragraphs.Item(i).Range.Delete();
            } catch(e) {}
          }
        }

        appliedStyles['jiyao_banji_bottom_fixed'] = 1;
        console.log('[apply-format] 纪要格式底部版记已创建');
      } catch (e) {
        console.warn('[apply-format] 纪要格式版记处理失败', e);
      }
    }
  } catch (e) {
    console.warn('[apply-format] 版记处理失败', e);
  }

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
    paragraphCount: totalParas,
    appliedStyles: appliedStyles,
    detectedElements: detectedElements,
    debug: {
      titleIndex: titleIndex,
      separatorIndex: separatorIndex,
      dateIndex: dateIndex,
      chaosongIndex: chaosongIndex,
      yinfaIndex: yinfaIndex
    }
  };

} catch (e) {
  console.warn('[apply-format]', e);
  return { success: false, formatType: 'general', paragraphCount: 0, appliedStyles: {}, error: String(e) };
}
