/**
 * setup-special-elements.js
 * 处理特定公文格式的特殊要素
 *
 * 支持的特殊要素：
 * - 命令(令)格式：令号、签发人签名章
 * - 纪要格式：出席/请假/列席名单
 *
 * 出参: { success: boolean, formatType: string, elements: object, message: string }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, formatType: 'general', elements: {}, message: '未找到活动文档' };
  }

  // ========================================
  // 常量定义
  // ========================================
  var MM_TO_POINTS = 2.835;

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

      if (alignment === 1 && /命令|令$/.test(text)) {
        return FORMAT_TYPES.ORDER;
      }

      if (alignment === 1 && /纪要/.test(text)) {
        return FORMAT_TYPES.MINUTES;
      }

      if (/〔\d{4}〕\d+号/.test(text) || /\[\d{4}\]\d+号/.test(text)) {
        if (alignment === 2) {
          return FORMAT_TYPES.LETTER;
        }
      }

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
    }

    return FORMAT_TYPES.GENERAL;
  }

  var formatType = detectFormatType();
  var elements = {
    orderNumber: { found: false, index: -1 },
    attendeeList: { found: false, indices: [] },
    signature: { found: false, index: -1 }
  };

  // ========================================
  // 命令(令)格式特殊处理
  // ========================================
  if (formatType === FORMAT_TYPES.ORDER) {
    var totalParas = doc.Paragraphs.Count;

    for (var i = 1; i <= totalParas; i++) {
      var para = doc.Paragraphs.Item(i);
      var text = getParaText(para);

      // 识别令号（第X号）
      if (/^第\d+号$/.test(text)) {
        elements.orderNumber.found = true;
        elements.orderNumber.index = i;

        // 设置令号格式：3号仿宋居中
        try {
          var range = para.Range;
          range.Font.NameFarEast = "仿宋_GB2312";
          range.Font.Name = "仿宋_GB2312";
          range.Font.Size = 16;  // 3号
          para.Format.Alignment = 1;  // 居中
        } catch (e) {
          console.warn('[setup-special-elements] 设置令号格式失败', e);
        }
      }

      // 识别签发人签名章位置（命令末尾的署名）
      // 通常在日期之前
      if (/^\d{4}年\d{1,2}月\d{1,2}日$/.test(text)) {
        // 检查前一段是否是署名
        if (i > 1) {
          var prevText = getParaText(doc.Paragraphs.Item(i - 1));
          if (prevText && prevText.length <= 30 && !/[，。！？、；：]$/.test(prevText)) {
            elements.signature.found = true;
            elements.signature.index = i - 1;

            // 设置署名格式
            try {
              var sigPara = doc.Paragraphs.Item(i - 1);
              sigPara.Format.Alignment = 2;  // 右对齐
              sigPara.Format.RightIndent = 42;  // 右空二字
            } catch (e) {}
          }
        }
      }
    }
  }

  // ========================================
  // 纪要格式特殊处理
  // ========================================
  if (formatType === FORMAT_TYPES.MINUTES) {
    var totalParas = doc.Paragraphs.Count;

    for (var i = 1; i <= totalParas; i++) {
      var para = doc.Paragraphs.Item(i);
      var text = getParaText(para);

      // 识别出席/请假/列席名单
      if (/^出席[：:]/.test(text) || /^请假[：:]/.test(text) || /^列席[：:]/.test(text)) {
        elements.attendeeList.found = true;
        elements.attendeeList.indices.push(i);

        // 设置名单格式：仿宋体，左对齐
        try {
          var range = para.Range;
          range.Font.NameFarEast = "仿宋_GB2312";
          range.Font.Size = 16;  // 3号
          para.Format.Alignment = 0;  // 左对齐
          para.Format.FirstLineIndent = 0;  // 无首行缩进
        } catch (e) {
          console.warn('[setup-special-elements] 设置出席名单格式失败', e);
        }
      }
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

  var message = '特殊要素处理完成';

  if (formatType === FORMAT_TYPES.ORDER) {
    if (elements.orderNumber.found) {
      message += '（令号已设置）';
    }
    if (elements.signature.found) {
      message += '（签发人署名已设置）';
    }
  }

  if (formatType === FORMAT_TYPES.MINUTES) {
    if (elements.attendeeList.found) {
      message += '（出席名单已设置，共' + elements.attendeeList.indices.length + '处）';
    }
  }

  return {
    success: true,
    formatType: formatType,
    formatName: formatNames[formatType] || '通用公文格式',
    elements: elements,
    message: message
  };

} catch (e) {
  console.warn('[setup-special-elements]', e);
  return { success: false, formatType: 'general', elements: {}, message: String(e) };
}
