/**
 * setup-styles.js
 * 创建 GB/T 9704-2012 国标公文所需的所有样式
 *
 * 出参: { success: boolean, created: number, existing: number, message: string }
 */

try {
  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, created: 0, existing: 0, message: '未找到活动文档' };
  }

  // ========================================
  // 样式定义（21个公文要素）
  // ========================================
  var styleDefs = [
    // ===== 版头 =====
    {
      name: 'GBT9704-份号',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-密级',
      fontCN: '黑体',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-紧急程度',
      fontCN: '黑体',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-发文机关标志',
      fontCN: '方正小标宋简体',
      fontEN: 'Times New Roman',
      fontSize: 36,  // 可根据实际情况调整
      bold: true,
      alignment: 1,  // 居中
      firstLineIndent: 0,
      color: 0xFF0000  // 红色
    },
    {
      name: 'GBT9704-发文字号',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 1,  // 居中
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-签发人',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 2,  // 右对齐
      firstLineIndent: 0
    },

    // ===== 主体 =====
    {
      name: 'GBT9704-标题',
      fontCN: '方正小标宋简体',
      fontEN: 'Times New Roman',
      fontSize: 22,  // 2号
      bold: true,
      alignment: 1,  // 居中
      firstLineIndent: 0,
      spaceBefore: 0,
      spaceAfter: 0
    },
    {
      name: 'GBT9704-主送机关',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,  // 左对齐
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-一级标题',
      fontCN: '黑体',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-二级标题',
      fontCN: '楷体_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-三级标题',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: true,  // 加粗
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-四级标题',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: true,  // 加粗
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-正文',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 3,  // 两端对齐
      firstLineIndent: 32,  // 首行缩进2字符
      lineSpacing: 28  // 行距
    },
    {
      name: 'GBT9704-附件说明',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-落款',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 2,  // 右对齐
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-成文日期',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 2,  // 右对齐
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-附注',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,
      firstLineIndent: 32
    },

    // ===== 版记 =====
    {
      name: 'GBT9704-抄送',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 14,  // 4号
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-印发',
      fontCN: '仿宋_GB2312',
      fontEN: 'Times New Roman',
      fontSize: 14,  // 4号
      bold: false,
      alignment: 0,
      firstLineIndent: 0
    },

    // ===== 页码 =====
    {
      name: 'GBT9704-页码',
      fontCN: '宋体',
      fontEN: 'Times New Roman',
      fontSize: 14,  // 4号
      bold: false,
      alignment: 1  // 居中
    },

    // ===== 其他 =====
    {
      name: 'GBT9704-附件标题',
      fontCN: '方正小标宋简体',
      fontEN: 'Times New Roman',
      fontSize: 22,
      bold: true,
      alignment: 1,  // 居中
      firstLineIndent: 0
    },
    {
      name: 'GBT9704-附件标识',
      fontCN: '黑体',
      fontEN: 'Times New Roman',
      fontSize: 16,
      bold: false,
      alignment: 0,  // 左对齐，顶格
      firstLineIndent: 0
    }
  ];

  // ========================================
  // 创建或更新样式
  // ========================================
  var created = 0;
  var existing = 0;
  var updated = 0;

  for (var i = 0; i < styleDefs.length; i++) {
    var def = styleDefs[i];
    var style = null;

    // 尝试获取现有样式
    try {
      style = doc.Styles.Item(def.name);
      existing++;
    } catch (e) {
      // 样式不存在，创建新样式
      try {
        style = doc.Styles.Add(def.name, 0);  // 0 = 段落样式
        created++;
      } catch (e2) {
        // 创建失败，跳过
        console.warn('[setup-styles] 创建样式失败:', def.name, e2);
        continue;
      }
    }

    // 设置样式属性
    if (style) {
      try {
        // 字体设置
        style.Font.NameFarEast = def.fontCN;
        style.Font.Name = def.fontEN;
        style.Font.Size = def.fontSize;
        style.Font.Bold = def.bold ? true : false;
        if (def.color !== undefined) {
          style.Font.Color = def.color;
        }

        // 段落格式设置
        style.ParagraphFormat.Alignment = def.alignment;
        style.ParagraphFormat.FirstLineIndent = def.firstLineIndent || 0;

        // 行距设置
        if (def.lineSpacing !== undefined) {
          style.ParagraphFormat.LineSpacingRule = 4;  // wdLineSpaceExactly
          style.ParagraphFormat.LineSpacing = def.lineSpacing;
        }

        // 段前段后间距
        if (def.spaceBefore !== undefined) {
          style.ParagraphFormat.SpaceBefore = def.spaceBefore;
        }
        if (def.spaceAfter !== undefined) {
          style.ParagraphFormat.SpaceAfter = def.spaceAfter;
        }

      } catch (e3) {
        console.warn('[setup-styles] 设置样式属性失败:', def.name, e3);
      }
    }
  }

  return {
    success: true,
    created: created,
    existing: existing,
    message: '样式准备完成，新建 ' + created + ' 个，已存在 ' + existing + ' 个'
  };

} catch (e) {
  console.warn('[setup-styles]', e);
  return { success: false, created: 0, existing: 0, message: String(e) };
}