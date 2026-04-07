/**
 * generate-report.js
 * 研究报告生成脚本
 *
 * 入参:
 *   - topic (string): 报告主题（action=generate 时必需）
 *   - action (string): "generate" | "write"
 *   - content (string): 生成的报告内容（action=write 时必需）
 *
 * 出参:
 *   - action=generate: { success: boolean, content?: string, error?: string }
 *   - action=write: { success: boolean, error?: string }
 */

// ============ 格式配置 ============
var FONT_NAME = '微软雅黑';
var BODY_SIZE = 14;  // 四号字 = 14pt
var LINE_SPACING = 28;  // 固定值28磅
var FIRST_LINE_INDENT = 21;  // 首行缩进2字符（约21磅）
// ================================

try {
  var action = typeof action !== 'undefined' ? action : 'generate';

  if (action === 'generate') {
    // ============ 生成报告内容 ============
    var reportTopic = typeof topic !== 'undefined' ? topic : '';

    if (!reportTopic) {
      return { success: false, error: '报告主题不能为空' };
    }

    // TODO: 调用服务端 API 生成内容
    // 临时返回 Markdown 格式的示例内容，等待服务端对接
    var generatedContent = generateSampleReport(reportTopic);

    return { success: true, content: generatedContent };

  } else if (action === 'write') {
    // ============ 写入文档 ============
    var reportContent = typeof content !== 'undefined' ? content : '';

    // 类型检查：确保 reportContent 是字符串
    if (typeof reportContent !== 'string') {
      return { success: false, error: '报告内容必须是字符串，实际类型：' + typeof reportContent };
    }

    if (!reportContent) {
      return { success: false, error: '报告内容不能为空' };
    }

    var doc = Application.ActiveDocument;
    if (!doc) {
      return { success: false, error: '未找到活动文档' };
    }

    // 清空当前文档
    doc.Content.Text = '';

    // 解析 Markdown 内容并写入
    var lines = reportContent.split('\n');
    var isFirstParagraph = true;

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];

      // 跳过空行
      if (!line || line.trim() === '') {
        continue;
      }

      // 标题处理（# 标题）
      if (line.indexOf('#') === 0) {
        var headingLevel = 0;
        var j = 0;
        while (j < line.length && line.charAt(j) === '#') {
          headingLevel++;
          j++;
        }
        var headingText = line.substring(j).trim();

        // 插入标题
        var newPara = doc.Content.Paragraphs.Add();
        if (newPara && newPara.Range) {
          newPara.Range.Text = headingText;
          // 设置标题样式
          if (headingLevel >= 1 && headingLevel <= 9) {
            try {
              newPara.Style = doc.Styles.Item('Heading ' + headingLevel);
            } catch (e) {
              // 样式设置失败，使用默认样式
            }
          }
          // 应用字体格式：微软雅黑、加粗
          applyFont(newPara.Range, true, getHeadingSize(headingLevel));
        }
        isFirstParagraph = false;
        continue;
      }

      // 表格处理（| 开头）
      if (line.indexOf('|') === 0) {
        // 检测是否为表格分隔行（包含 --- 或 :-- ）
        if (line.indexOf('---') !== -1 || line.indexOf(':--') !== -1 || line.indexOf(':-') !== -1) {
          continue; // 跳过分隔行
        }

        // 收集表格行
        var tableRows = [];
        while (i < lines.length && lines[i].indexOf('|') === 0) {
          if (lines[i].indexOf('---') === -1 && lines[i].indexOf(':--') === -1 && lines[i].indexOf(':-') === -1) {
            tableRows.push(lines[i]);
          }
          i++;
        }
        i--; // 回退一行，因为 for 循环会 i++

        // 解析并插入表格
        if (tableRows.length > 1) {
          insertTable(doc, tableRows);
        }
        isFirstParagraph = false;
        continue;
      }

      // 普通段落（首行缩进不加粗）
      if (isFirstParagraph) {
        var firstRange = doc.Content.Paragraphs.Item(1).Range;
        firstRange.Text = line;
        applyFont(firstRange, false, BODY_SIZE, true); // 第四个参数表示首行缩进
        isFirstParagraph = false;
      } else {
        var para = doc.Content.Paragraphs.Add();
        if (para && para.Range) {
          para.Range.Text = line;
          applyFont(para.Range, false, BODY_SIZE, true);
        }
      }
    }

    return { success: true };

  } else {
    return { success: false, error: '未知操作：' + action };
  }

} catch (e) {
  console.warn('[generate-report]', e);
  return { success: false, error: String(e) };
}

// ============ 辅助函数 ============

/**
 * 应用字体格式和段落格式
 * @param {Object} range 文档范围对象
 * @param {boolean} bold 是否加粗
 * @param {number} size 字号（磅值）
 * @param {boolean} indent 是否首行缩进
 */
function applyFont(range, bold, size, indent) {
  try {
    if (!range) {
      return;
    }
    // 字体格式
    if (range.Font) {
      range.Font.Name = FONT_NAME;
      range.Font.Bold = bold ? 1 : 0;
      range.Font.Size = size;
    }
    // 段落格式：固定行距28磅
    if (range.ParagraphFormat) {
      range.ParagraphFormat.LineSpacing = LINE_SPACING;
      range.ParagraphFormat.LineSpacingRule = 3; // wdLineSpacingFixed = 3
      // 首行缩进2字符
      if (indent) {
        range.ParagraphFormat.FirstLineIndent = FIRST_LINE_INDENT;
      }
    }
  } catch (e) {
    console.warn('[applyFont]', e);
  }
}

/**
 * 根据标题级别获取字号
 * @param {number} level 标题级别 (1-6)
 * @returns {number} 字号（磅值）
 */
function getHeadingSize(level) {
  // 中文文档字号对照：一号=26, 二号=22, 三号=16, 四号=14, 五号=11, 小五号=9
  // 标题通常比正文大，这里用：
  // 一级标题: 22pt (二号)
  // 二级标题: 18pt
  // 三级标题: 16pt (三号)
  // 四级及以下: 14pt (四号)
  var sizes = {
    1: 22,
    2: 18,
    3: 16,
    4: 14,
    5: 14,
    6: 14
  };
  return sizes[level] || BODY_SIZE;
}

/**
 * 生成示例报告内容（临时实现，待服务端对接）
 * @param {string} topic 报告主题
 * @returns {string} Markdown 格式的报告内容
 */
function generateSampleReport(topic) {
  var report = '';
  report += '# ' + topic + '\n\n';
  report += '## 一、执行摘要\n\n';
  report += '本报告对[' + topic + ']进行深入分析，旨在全面梳理该领域的发展脉络，探讨其发展现状、主要趋势及面临的核心挑战。报告基于行业公开数据、权威研究机构报告及专家访谈等多种信息源，力求客观、准确地呈现[' + topic + ']的全貌，为相关决策提供参考依据。\n\n';
  report += '通过系统研究，我们发现[' + topic + ']正处于快速发展的关键阶段，技术创新活跃，应用场景持续拓展，市场规模稳步扩大。同时，行业发展也面临核心技术突破、人才培养、政策监管等多方面挑战，需要产业链各方协同应对。\n\n';
  report += '## 二、背景介绍\n\n';
  report += '### 2.1 研究背景\n\n';
  report += '随着经济的快速发展和科技的不断进步，[' + topic + ']已成为当今社会中备受关注的重要领域。近年来，国家陆续出台了一系列扶持政策，为该领域的发展创造了良好的政策环境。同时，资本市场对该领域的关注度持续提升，大量资金涌入，进一步推动了行业的蓬勃发展。\n\n';
  report += '### 2.2 研究意义\n\n';
  report += '深入研究[' + topic + ']具有重要的理论价值和实践意义。从理论层面看，有助于丰富和完善相关理论体系，为学术研究提供新的视角和案例。从实践层面看，通过对行业发展现状和趋势的深入分析，可以为政府制定产业政策、企业制定发展战略、投资机构决策等提供参考。\n\n';
  report += '### 2.3 研究方法\n\n';
  report += '本报告采用文献研究法、案例分析法和比较研究法等多种研究方法。通过广泛收集和梳理国内外相关文献资料，系统梳理[' + topic + ']的理论基础和发展历程；通过典型案例分析，深入探讨行业发展的内在规律；通过横向比较，借鉴国际先进经验和做法。\n\n';
  report += '## 三、市场分析\n\n';
  report += '### 3.1 市场规模\n\n';
  report += '| 年份 | 市场规模 | 同比增长 | 市场集中度 |\n';
  report += '|------|----------|----------|------------|\n';
  report += '| 2021 | 80亿 | 12% | 65% |\n';
  report += '| 2022 | 95亿 | 19% | 68% |\n';
  report += '| 2023 | 120亿 | 26% | 70% |\n';
  report += '| 2024 | 155亿 | 29% | 72% |\n';
  report += '| 2025 | 200亿 | 29% | 75% |\n\n';
  report += '根据上表数据可以看出，过去五年间[' + topic + ']市场规模呈现出持续快速增长态势，年均复合增长率保持在20%以上。预计未来三至五年，随着技术成熟度提高和应用场景的进一步拓展，市场规模有望继续保持高速增长。\n\n';
  report += '### 3.2 竞争格局\n\n';
  report += '当前[' + topic + ']市场呈现出第一梯队与第二梯队并存的竞争格局。以大型龙头企业为代表的第一梯队占据了市场的主导地位，拥有较强的技术研发能力和品牌影响力；以创新型中小企业为代表的第二梯队专注于细分领域的技术创新和应用探索；此外，部分大型企业通过跨界布局进入该领域，进一步加剧了市场竞争。\n\n';
  report += '### 3.3 产业链分析\n\n';
  report += '[' + topic + ']产业链上游主要包括核心技术和关键零部件的研发与制造，中游为产品集成和系统解决方案的提供，下游则是应用场景的拓展和服务的落地。整体来看，产业链各环节联系紧密，协同发展趋势明显。\n\n';
  report += '## 四、发展趋势\n\n';
  report += '### 4.1 技术创新趋势\n\n';
  report += '技术创新是推动[' + topic + ']发展的核心动力。当前，技术创新主要体现在以下几个方向：一是核心算法的持续优化和突破，二是软硬件协同设计的深入推进，三是新材料和新工艺的广泛应用。未来，随着人工智能、大数据、云计算等新一代信息技术的深度融合，技术创新将进入新一轮加速期。\n\n';
  report += '### 4.2 应用场景拓展\n\n';
  report += '[' + topic + ']的应用场景正在不断丰富和拓展。从传统的优势领域向新兴领域延伸，从单一场景向多场景融合发展。例如，在智能制造、智慧城市、数字政府、医疗健康、教育培训等领域都有着广阔的应用前景。随着技术的成熟和成本的下降，应用场景将进一步向大众消费市场渗透。\n\n';
  report += '### 4.3 商业模式创新\n\n';
  report += '商业模式创新是[' + topic + ']发展的重要特征。从单一的产品销售向综合解决方案提供商转型，从一次性交付向持续性服务收费模式演进。平台化、生态化成为行业领先企业的重要战略选择，通过构建开放共享的平台生态，实现产业链上下游的协同发展。\n\n';
  report += '### 4.4 国际化发展\n\n';
  report += '在全球化背景下，[' + topic + ']的国际化发展是大势所趋。一方面，国内企业积极拓展国际市场；另一方面，国际企业也加大了对中国市场的投入。跨境合作、技术交流、人才流动等形式的国际化合作日益频繁，为行业发展注入了新的活力。\n\n';
  report += '## 五、面临挑战\n\n';
  report += '### 5.1 核心技术突破难度大\n\n';
  report += '虽然[' + topic + ']在多个领域取得了重要进展，但在核心技术层面仍面临诸多挑战。部分关键零部件和核心算法仍依赖进口，自主创新能力有待进一步提升。此外，从基础研究到产业化的转化周期较长，资金投入大，企业面临较大的研发压力。\n\n';
  report += '### 5.2 人才竞争激烈\n\n';
  report += '[' + topic + ']属于知识密集型产业，对专业人才的需求旺盛。目前，行业内高端人才相对短缺，尤其是既懂技术又懂市场的复合型人才更是稀缺。人才竞争日趋激烈，人力成本不断上升，成为制约企业发展的重要因素。\n\n';
  report += '### 5.3 政策监管不确定性\n\n';
  report += '随着[' + topic + ']的快速发展，政策监管也在不断完善和调整。数据安全、隐私保护、伦理规范等领域的监管要求日益严格，对企业的合规运营提出了更高要求。政策的不确定性增加了企业的经营风险，需要企业密切关注政策动向，及时调整发展策略。\n\n';
  report += '### 5.4 市场竞争加剧\n\n';
  report += '市场的快速成长吸引了越来越多的参与者，市场竞争日趋激烈。价格战、同质化竞争等问题时有发生，部分领域出现了产能过剩的迹象。企业需要通过差异化竞争、提升产品和服务质量来实现可持续发展。\n\n';
  report += '## 六、结论与建议\n\n';
  report += '### 6.1 主要结论\n\n';
  report += '通过对[' + topic + ']的深入研究，本报告得出以下主要结论：一是市场需求持续旺盛，发展前景广阔；二是技术创新是核心驱动力，需持续加大研发投入；三是产业链协同发展趋势明显，生态构建成为竞争关键；四是挑战与机遇并存，需要多方协同应对。\n\n';
  report += '### 6.2 发展建议\n\n';
  report += '针对[' + topic + ']的发展现状和未来趋势，本报告提出以下建议：一是加强核心技术研发，提升自主创新能力；二是注重人才培养和引进，打造高素质的专业团队；三是积极拓展应用场景，推动技术与市场的深度融合；四是加强产业链上下游合作，构建开放共赢的生态系统；五是密切关注政策动向，确保企业合规稳健运营。\n\n';
  report += '### 6.3 展望未来\n\n';
  report += '展望未来，[' + topic + ']有望继续保持快速发展态势。随着技术的不断成熟、应用场景的持续拓展以及政策环境的日益完善，该领域将迎来更大的发展机遇。我们建议相关各方抓住历史机遇，加强合作，共同推动[' + topic + ']的高质量发展。\n\n';
  return report;
}

/**
 * 插入表格到文档
 * @param {Object} doc WPS 文档对象
 * @param {string[]} rows 表格行数据（Markdown 格式）
 */
function insertTable(doc, rows) {
  if (!rows || rows.length < 2) {
    return;
  }

  // 解析列数（第一行的 | 分隔数量）
  var firstRow = rows[0];
  var cells = firstRow.split('|');
  // 去掉首尾空单元格
  if (cells.length > 0 && cells[0].trim() === '') {
    cells.shift();
  }
  if (cells.length > 0 && cells[cells.length - 1].trim() === '') {
    cells.pop();
  }
  var colCount = cells.length;

  if (colCount === 0) {
    return;
  }

  // 创建表格
  try {
    var range = doc.Content;
    range.Collapse(0); // wdCollapseEnd

    var table = doc.Tables.Add(range, rows.length, colCount);
    if (!table) {
      return;
    }

    // 设置表格边框为实线
    try {
      var tbl = table;
      if (tbl && tbl.Borders) {
        // 遍历所有8种边框类型，设置实线
        for (var b = 1; b <= 8; b++) {
          try {
            tbl.Borders.Item(b).LineStyle = 1;  // wdLineStyleSingle = 1
          } catch (e2) {}
        }
        // 额外：清除单元格的斜线（对角线）
        try {
          tbl.Cell(1, 1).DiagonalDirection = 0;  // wdDiagonalNone = 0
        } catch (e2) {}
        // 遍历所有单元格清除斜线
        try {
          for (var ri = 1; ri <= rows.length; ri++) {
            for (var ci = 1; ci <= colCount; ci++) {
              try {
                tbl.Cell(ri, ci).DiagonalDirection = 0;
              } catch (e3) {}
            }
          }
        } catch (e2) {}
      }
    } catch (e) {
      console.warn('[insertTable] border error:', e);
    }

    // 填充数据
    for (var r = 0; r < rows.length; r++) {
      var rowCells = rows[r].split('|');
      // 清理首尾空单元格
      if (rowCells.length > 0 && rowCells[0].trim() === '') {
        rowCells.shift();
      }
      if (rowCells.length > 0 && rowCells[rowCells.length - 1].trim() === '') {
        rowCells.pop();
      }

      for (var c = 0; c < rowCells.length && c < colCount; c++) {
        try {
          var rawText = rowCells[c] ? rowCells[c] : '';
          // 彻底清理：移除所有 | 字符，首尾空白，替换连续空白
          var cellText = rawText.replace(/\|/g, '').replace(/^\s+|\s+$/g, '').replace(/\s+/g, ' ');
          if (cellText === '-') {
            cellText = ''; // 分隔行不显示
          }
          var cellRange = table.Cell(r + 1, c + 1).Range;
          cellRange.Text = cellText;
          // 应用字体格式
          applyFont(cellRange, false, BODY_SIZE, false);
        } catch (e) {
          console.warn('[insertTable] cell error:', e);
        }
      }
    }
  } catch (e) {
    console.warn('[insertTable]', e);
  }
}