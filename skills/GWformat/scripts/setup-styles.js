/**
 * setup-styles.js
 * 在文档中创建XX集团公文排版所需的样式。
 *
 * 出参: { success: boolean, message: string }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, message: '未找到活动文档' };
  }

  // 样式定义（按XX集团公文排版格式）
  var styleDefs = [
    {
      name: '集团1标题',
      fontCN: '方正小标宋简体',
      fontEN: 'Times New Roman',
      fontSize: 22,
      bold: true,
      alignment: 1,
      firstLineIndent: 0,
      spaceAfter: 14
    },
    {
      name: '集团主送单位抬头',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: '集团2级标题黑体',
      fontCN: '黑体',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: '集团3级段落重点',
      fontCN: '楷体_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 10
    },
    {
      name: '集团4级数字编号',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 10
    },
    {
      name: '集团正文文本缩进',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 32
    },
    {
      name: '集团结尾语',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 32
    },
    {
      name: '集团附件',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: '集团落款',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 2,
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
        style.ParagraphFormat.Alignment = def.alignment;
        style.ParagraphFormat.FirstLineIndent = def.firstLineIndent;
        if (def.spaceAfter) {
          style.ParagraphFormat.SpaceAfter = def.spaceAfter;
        }
      } catch (e3) {}
    }
  }

  return {
    success: true,
    message: '样式准备完成，新建 ' + created + ' 个，已存在 ' + existing + ' 个'
  };

} catch (e) {
  console.warn('[setup-styles]', e);
  return { success: false, message: String(e) };
}