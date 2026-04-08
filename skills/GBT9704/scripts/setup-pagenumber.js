/**
 * setup-pagenumber.js
 * 设置 GB/T 9704-2012 公文页码格式
 *
 * 支持四种公文格式：
 * - 通用公文：首页显示页码
 * - 信函格式：首页不显示页码
 * - 命令(令)格式：首页显示页码
 * - 纪要格式：首页显示页码
 *
 * 页码规则：
 * - 字体：4号半角宋体阿拉伯数字
 * - 格式：一字线 + 页码 + 一字线（如 "— 1 —"）
 * - 位置：版心下边缘之下 7mm
 * - 单页码：居右空一字
 * - 双页码：居左空一字
 *
 * 出参: { success: boolean, formatType: string, message: string }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, formatType: 'general', message: '未找到活动文档' };
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
  // 格式类型检测
  // ========================================
  function detectFormatType() {
    var totalParas = doc.Paragraphs.Count;

    for (var i = 1; i <= Math.min(10, totalParas); i++) {
      var para = doc.Paragraphs.Item(i);
      var text = para.Range.Text ? para.Range.Text.replace(/[\r\n]/g, '').trim() : '';
      var alignment = para.Format.Alignment;

      if (!text) continue;

      // 命令(令)格式
      if (alignment === 1 && /命令|令$/.test(text)) {
        return FORMAT_TYPES.ORDER;
      }

      // 纪要格式
      if (alignment === 1 && /纪要/.test(text)) {
        return FORMAT_TYPES.MINUTES;
      }

      // 信函格式：发文字号右对齐（国标要求使用六角括号〔〕）
      if (/〔\d{4}〕\d+号/.test(text)) {
        if (alignment === 2) {
          return FORMAT_TYPES.LETTER;
        }
      }

      // 信函格式：发文机关标志不含"文件"
      if (alignment === 1 && !/文件$/.test(text)) {
        if (/政府$|办公室$|委员会$|厅$|局$|委$/.test(text)) {
          for (var j = i + 1; j <= Math.min(20, totalParas); j++) {
            var nextText = doc.Paragraphs.Item(j).Range.Text;
            if (nextText && /函$|复函$/.test(nextText.replace(/[\r\n]/g, '').trim())) {
              return FORMAT_TYPES.LETTER;
            }
          }
        }
      }
    }

    return FORMAT_TYPES.GENERAL;
  }

  var formatType = detectFormatType();

  // ========================================
  // 页码设置
  // ========================================

  // 判断是否需要首页页码
  // 通用公文、命令、纪要：首页显示页码
  // 信函格式：首页不显示页码
  var showFirstPageNumber = (formatType !== FORMAT_TYPES.LETTER);

  // 信函格式特殊处理：
  // 国标10.1：首页不显示页码
  // 仅通过“首页不同”控制首页页脚，后续页仍按常规页码规则设置

  // 1. 设置奇偶页不同
  try {
    doc.PageSetup.OddAndEvenPagesHeaderFooter = true;
  } catch (e) {
    console.warn('[setup-pagenumber] 设置奇偶页不同失败', e);
  }

  // 2. 设置首页不同
  // 重要：需要先关闭再根据需要开启，否则文档可能有残留设置
  try {
    doc.PageSetup.DifferentFirstPageHeaderFooter = !showFirstPageNumber;
  } catch (e) {
    console.warn('[setup-pagenumber] 设置首页不同失败', e);
  }

  // 3. 设置页脚距离
  try {
    doc.PageSetup.FooterDistance = 28 * MM_TO_POINTS;
  } catch (e) {
    console.warn('[setup-pagenumber] 设置页脚距离失败', e);
  }

  // ========================================
  // 获取节
  // ========================================
  var section = null;
  try {
    section = doc.Sections.Item(1);
  } catch (e) {
    console.warn('[setup-pagenumber] 获取节失败', e);
    return { success: false, formatType: formatType, message: '无法获取文档节' };
  }

  // ========================================
  // 设置页脚内容
  // ========================================

  // 设置奇数页页脚（居右空一字）
  try {
    var oddFooter = section.Footers.Item(1);  // wdHeaderFooterPrimary
    if (oddFooter) {
      var oddRange = oddFooter.Range;
      oddRange.Delete();

      oddRange.Select();
      var sel = Application.Selection;

      sel.TypeText("— ");
      sel.Fields.Add(sel.Range, 33, "", true);  // wdFieldPage = 33
      sel.TypeText(" —");

      oddFooter.Range.Font.NameFarEast = "宋体";
      oddFooter.Range.Font.Name = "Times New Roman";
      oddFooter.Range.Font.Size = 14;
      oddFooter.Range.ParagraphFormat.Alignment = 2;  // 右对齐
      oddFooter.Range.ParagraphFormat.RightIndent = 10.5;  // 右缩进一字（210缇）
      oddFooter.Range.ParagraphFormat.LeftIndent = 0;
      oddFooter.Range.ParagraphFormat.FirstLineIndent = 0;
    }
  } catch (e) {
    console.warn('[setup-pagenumber] 设置奇数页页脚失败', e);
  }

  // 设置偶数页页脚（居左空一字）
  try {
    var evenFooter = section.Footers.Item(3);  // wdHeaderFooterEvenPages
    if (evenFooter) {
      var evenRange = evenFooter.Range;
      evenRange.Delete();

      evenRange.Select();
      var sel = Application.Selection;

      sel.TypeText("— ");
      sel.Fields.Add(sel.Range, 33, "", true);
      sel.TypeText(" —");

      evenFooter.Range.Font.NameFarEast = "宋体";
      evenFooter.Range.Font.Name = "Times New Roman";
      evenFooter.Range.Font.Size = 14;
      evenFooter.Range.ParagraphFormat.Alignment = 0;  // 左对齐
      evenFooter.Range.ParagraphFormat.LeftIndent = 10.5;  // 左缩进一字（210缇）
      evenFooter.Range.ParagraphFormat.RightIndent = 0;
      evenFooter.Range.ParagraphFormat.FirstLineIndent = 0;
    }
  } catch (e) {
    console.warn('[setup-pagenumber] 设置偶数页页脚失败', e);
  }

  // 设置首页页脚（仅信函格式需要清空）
  // 重要：需要显式设置内容才能触发WPS创建首页页脚引用
  if (!showFirstPageNumber) {
    try {
      // 确保属性已设置
      section.PageSetup.DifferentFirstPageHeaderFooter = true;

      var firstFooter = section.Footers.Item(2);  // wdHeaderFooterFirstPage
      if (firstFooter) {
        var footerRange = firstFooter.Range;

        // 先清空内容
        footerRange.Delete();

        // 插入一个空格并立即删除，触发WPS创建页脚文件
        footerRange.InsertAfter(" ");
        footerRange.Delete();

        // 显式设置为空段落
        footerRange.Text = "";

        // 设置段落格式
        try {
          footerRange.ParagraphFormat.Alignment = 1;  // 居中
          footerRange.ParagraphFormat.FirstLineIndent = 0;
        } catch (e) {}

        console.log('[setup-pagenumber] 首页页脚已创建并清空');
      }
    } catch (e) {
      console.warn('[setup-pagenumber] 设置首页页脚失败', e);

      // 备选方案：重置属性并重新设置
      try {
        section.PageSetup.DifferentFirstPageHeaderFooter = false;
        section.PageSetup.DifferentFirstPageHeaderFooter = true;

        var firstFooter2 = section.Footers.Item(2);
        if (firstFooter2) {
          firstFooter2.Range.Delete();
        }
      } catch (e2) {
        console.warn('[setup-pagenumber] 备选方案也失败', e2);
      }
    }
  }

  // ========================================
  // 退出页眉页脚编辑视图，返回正文
  // ========================================
  try {
    doc.Range(0, 0).Select();
  } catch (e) {
    console.warn('[setup-pagenumber] 退出页眉页脚视图失败', e);
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

  var firstPageNote = showFirstPageNumber ? '首页显示页码' : '首页无页码';

  return {
    success: true,
    formatType: formatType,
    formatName: formatNames[formatType] || '通用公文格式',
    message: '页码设置完成（' + (formatNames[formatType] || '通用公文格式') + '，' + firstPageNote + '）'
  };

} catch (e) {
  console.warn('[setup-pagenumber]', e);
  return { success: false, formatType: 'general', message: String(e) };
}
