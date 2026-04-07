/**
 * setup-styles.js
 * 在文档中创建会议纪要红头文件排版所需的样式。
 *
 * 出参: { success: boolean, message: string }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, message: '未找到活动文档' };
  }

  // 样式定义（会议纪要红头文件格式）
  var styleDefs = [
    {
      name: '会议纪要-红头',
      fontCN: '方正小标宋简体',
      fontEN: 'Times New Roman',
      fontSize: 42,
      bold: false,
      alignment: 1,
      firstLineIndent: 0,
      fontColor: 255
    },
    {
      name: '会议纪要-标题字',
      fontCN: '方正小标宋简体',
      fontEN: 'Times New Roman',
      fontSize: 42,
      bold: false,
      alignment: 1,
      firstLineIndent: 0,
      fontColor: 255
    },
    {
      name: '会议纪要-发文字号',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 1,
      firstLineIndent: 0
    },
    {
      name: '会议纪要-标题',
      fontCN: '方正小标宋简体',
      fontEN: 'Times New Roman',
      fontSize: 22,
      bold: true,
      alignment: 1,
      firstLineIndent: 0
    },
    {
      name: '会议纪要-一级标题',
      fontCN: '黑体',
      fontEN: '黑体',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: '会议纪要-二级标题',
      fontCN: '楷体_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: '会议纪要-正文',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 32
    },
    {
      name: '会议纪要-参会人员',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 32
    },
    {
      name: '会议纪要-抄送',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 14,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: '会议纪要-印发',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 14,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: '会议纪要-共印',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 14,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    }
  ];

  var created = 0;
  var existing = 0;

  for (var i = 0; i < styleDefs.length; i++) {
    var def = styleDefs[i];
    var style = null;

    try {
      style = doc.Styles.Item(def.name);
      existing++;
    } catch (e) {
      try {
        style = doc.Styles.Add(def.name, 0);
        created++;
      } catch (e2) {
        continue;
      }
    }

    if (style) {
      try {
        style.Font.NameFarEast = def.fontCN;
        style.Font.Name = def.fontEN;
        style.Font.Size = def.fontSize;
        style.Font.Bold = def.bold;
        if (def.fontColor) {
          style.Font.Color = def.fontColor;
        }
        style.ParagraphFormat.Alignment = def.alignment;
        style.ParagraphFormat.FirstLineIndent = def.firstLineIndent;
      } catch (e3) {}
    }
  }

  return {
    success: true,
    message: '会议纪要样式准备完成，新建 ' + created + ' 个，已存在 ' + existing + ' 个'
  };

} catch (e) {
  console.warn('[setup-styles]', e);
  return { success: false, message: String(e) };
}