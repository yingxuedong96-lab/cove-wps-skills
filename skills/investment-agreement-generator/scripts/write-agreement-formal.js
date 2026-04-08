/**
 * write-agreement-formal.js
 * 从零构建正式版增资协议文档，无需模板文件。
 * 格式参照标准法律文档规范。
 *
 * 入参: {
 *   parties: {
 *     companyName: string,
 *     founderName: string,
 *     esopPlatform: string,
 *     investorA: string,
 *     investorB: string,
 *     resourcePlatform: string,
 *     signDate: string,
 *     companyCode: string,
 *     founderIdType: string,
 *     founderIdNo: string,
 *     founderNationality: string,
 *     esopCode: string,
 *     investorACode: string,
 *     investorAType: string,
 *     investorBCode: string,
 *     investorBType: string,
 *     resourceCode: string,
 *     resourceCapital: string,
 *     establishDate: string,
 *     registeredCapital: string,
 *     mainBusiness: string,
 *     totalInvestment: string,
 *     newCapital: string
 *   },
 *   capitalTable: array,
 *   clauses: string
 * }
 * 出参: { success: boolean, message: string, docName: string }
 */

try {
  var parties = typeof parties !== 'undefined' ? parties : {};
  var capitalTable = typeof capitalTable !== 'undefined' ? capitalTable : [];
  var clauses = typeof clauses !== 'undefined' ? clauses : '';
  
  // 创建新文档
  var newDoc = Application.Documents.Add();
  if (!newDoc) {
    return { success: false, message: '无法创建新文档' };
  }
  
  var selection = Application.Selection;
  
  // ========== 格式定义 ==========
  
  // 封面格式：微软雅黑、四号(14pt)、居中
  function setCoverFormat(bold) {
    selection.Font.Name = '微软雅黑';
    selection.Font.Size = 14;
    selection.Font.Bold = bold ? true : false;
    selection.ParagraphFormat.Alignment = 1;
    selection.ParagraphFormat.LineSpacingRule = 0;
    selection.ParagraphFormat.SpaceBefore = 0;
    selection.ParagraphFormat.SpaceAfter = 0;
    selection.ParagraphFormat.FirstLineIndent = 0;
  }
  
  // 正文格式
  function setBodyFormat() {
    selection.Font.Name = '微软雅黑';
    selection.Font.Size = 9;
    selection.Font.Bold = false;
    selection.ParagraphFormat.Alignment = 0;
    selection.ParagraphFormat.LineSpacingRule = 4;
    selection.ParagraphFormat.LineSpacing = 25;
    selection.ParagraphFormat.FirstLineIndent = newDoc.Application.MillimetersToPoints(7);
    selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
  }
  
  // 章节标题格式：居中、加粗，标题前后不空行
  function setClauseTitleFormat() {
    selection.Font.Name = '微软雅黑';
    selection.Font.Size = 9;
    selection.Font.Bold = true;
    selection.ParagraphFormat.Alignment = 1; // wdAlignParagraphCenter
    selection.ParagraphFormat.LineSpacingRule = 4;
    selection.ParagraphFormat.LineSpacing = 25;
    selection.ParagraphFormat.SpaceBefore = 0;
    selection.ParagraphFormat.SpaceAfter = 0;
    selection.ParagraphFormat.FirstLineIndent = 0;
  }
  
  // "鉴于"标题格式
  function setRecitalTitleFormat() {
    selection.Font.Name = '微软雅黑';
    selection.Font.Size = 9;
    selection.Font.Bold = true;
    selection.ParagraphFormat.Alignment = 0;
    selection.ParagraphFormat.LineSpacingRule = 4;
    selection.ParagraphFormat.LineSpacing = 25;
    selection.ParagraphFormat.FirstLineIndent = newDoc.Application.MillimetersToPoints(7);
    selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
  }
  
  // 附件标题格式：居中、加粗
  function setAttachmentTitleFormat() {
    selection.Font.Name = '微软雅黑';
    selection.Font.Size = 9;
    selection.Font.Bold = true;
    selection.ParagraphFormat.Alignment = 1;
    selection.ParagraphFormat.LineSpacingRule = 4;
    selection.ParagraphFormat.LineSpacing = 25;
    selection.ParagraphFormat.FirstLineIndent = 0;
  }
  
  // 签署区格式
  function setSignFormat() {
    selection.Font.Name = '微软雅黑';
    selection.Font.Size = 9;
    selection.Font.Bold = false;
    selection.ParagraphFormat.Alignment = 0;
    selection.ParagraphFormat.LineSpacingRule = 4;
    selection.ParagraphFormat.LineSpacing = 25;
    selection.ParagraphFormat.FirstLineIndent = 0;
  }
  
  // 辅助函数
  function insertPara(text, formatFunc, isBold) {
    if (formatFunc === setCoverFormat) {
      setCoverFormat(isBold);
    } else {
      formatFunc();
    }
    selection.TypeText(text);
    selection.TypeParagraph();
  }
  
  function insertEmptyLine() {
    selection.TypeParagraph();
  }
  
  function insertPageBreak() {
    selection.InsertBreak(7);
  }
  
  // ========== 第一页：封面 ==========
  
  // 先空2行
  insertEmptyLine();
  insertEmptyLine();
  
  // 第一行：标的公司名称
  insertPara(parties.companyName || '【标的公司名称】', setCoverFormat, true);
  
  // 第二行：创始股东
  insertPara(parties.founderName || '【标的公司创始股东、控股股东、实际控制人名称】', setCoverFormat, true);
  
  // 第三行：ESOP持股平台
  insertPara(parties.esopPlatform || '【ESOP持股平台】', setCoverFormat, true);
  
  // 空一行
  insertEmptyLine();
  
  // "与"
  insertPara('与', setCoverFormat, false);
  
  // 空一行
  insertEmptyLine();
  
  // 投资人A
  insertPara(parties.investorA || '【投资人A投资实体名称】', setCoverFormat, true);
  
  // 投资人B（如有）
  if (parties.investorB) {
    insertPara(parties.investorB, setCoverFormat, true);
  }
  
  // 资源方持股平台
  insertPara(parties.resourcePlatform || '【资源方持股平台】', setCoverFormat, true);
  
  // 空一行
  insertEmptyLine();
  
  // "关于"
  insertPara('关于', setCoverFormat, false);
  
  // 空一行
  insertEmptyLine();
  
  // 标的公司名称
  insertPara(parties.companyName || '【标的公司名称】', setCoverFormat, true);
  
  // 空一行
  insertEmptyLine();
  
  // "之"
  insertPara('之', setCoverFormat, false);
  
  // 空一行
  insertEmptyLine();
  
  // "增资协议"
  insertPara('增资协议', setCoverFormat, true);
  
  // 空一行
  insertEmptyLine();
  
  // 日期
  insertPara(parties.signDate || '【XXXX】年【XX】月【XX】日', setCoverFormat, false);
  
  // 分页
  insertPageBreak();
  
  // ========== 第二页：协议前言 ==========
  
  insertPara('本《增资协议》（"本协议"）由以下各方于' + (parties.signDate || '【】年【】月【】日') + '签署：', setBodyFormat);
  
  // 第1条：公司
  var companyCode = parties.companyCode || '【】';
  insertPara('1. ' + (parties.companyName || '【标的公司名称】') + '，一家依据中华人民共和国（"中国"，仅就本协议而言，不包括中国香港特别行政区、中国澳门特别行政区和中国台湾省）法律有效设立并合法存续的有限责任公司，其统一社会信用代码为' + companyCode + '（"公司"）；', setBodyFormat);
  
  // 第2条：创始人
  var founderIdType = parties.founderIdType || '【身份证】';
  var founderIdNo = parties.founderIdNo || '【】';
  var founderNationality = parties.founderNationality || '【】';
  insertPara('2. ' + (parties.founderName || '【标的公司创始股东、控股股东、实际控制人名称】') + '，一位' + founderNationality + '国公民，其' + founderIdType + '号码为' + founderIdNo + '（"创始人"）；', setBodyFormat);
  
  // 第3条：ESOP
  var esopCode = parties.esopCode || '【】';
  insertPara('3. ' + (parties.esopPlatform || '【ESOP持股平台】') + '，一家依据中国法律有效设立并合法存续的有限合伙企业，其统一社会信用代码为' + esopCode + '（"持股平台"，持股平台与创始人合称"创始股东"）；', setBodyFormat);
  
  // 第4条：投资人A
  var investorACode = parties.investorACode || '【】';
  var investorAType = parties.investorAType || '【责任公司/合伙企业】';
  insertPara('4. ' + (parties.investorA || '【投资人A投资实体名称】') + '，一家依据中国法律有效设立并合法存续的有限' + investorAType + '，其统一社会信用代码为' + investorACode + '（"投资人A"）；', setBodyFormat);
  
  // 第5条：投资人B
  var investorBCode = parties.investorBCode || '【】';
  var investorBType = parties.investorBType || '【责任公司/合伙企业】';
  var investorBName = parties.investorB || '【投资人B】';
  insertPara('5. ' + (parties.investorB || '【投资人B投资实体名称】') + '，一家依据中国法律有效设立并合法存续的有限' + investorBType + '，其统一社会信用代码为' + investorBCode + '（"' + investorBName + '"）；', setBodyFormat);
  
  // 第6条：资源方
  var resourceCode = parties.resourceCode || '【】';
  var resourceCapital = parties.resourceCapital || '【】';
  var resourcePlatformName = parties.resourcePlatform || '【资源方持股平台】';
  var investorBShort = parties.investorB || '【投资人B】';
  insertPara('6. ' + resourcePlatformName + '，一家依据中国法律有效设立并合法存续的有限合伙企业，其统一社会信用代码为' + resourceCode + '（"' + resourcePlatformName + '"，' + resourcePlatformName + '（仅就其持有的公司注册资本人民币' + resourceCapital + '元而言）与投资人A及' + investorBShort + '合称为"投资者"，' + resourcePlatformName + '仅就其持有的公司注册资本人民币' + resourceCapital + '元而言称为"其他普通股股东"）。', setBodyFormat);
  
  insertPara('在本协议中每一方以下单独称"一方"、"该方"，合称"各方"，互称"一方"、"其他方"。', setBodyFormat);
  
  // ========== 鉴于条款 ==========
  
  insertPara('鉴于：', setRecitalTitleFormat);
  
  // 鉴于第1条 - mainBusiness映射投资款用途
  var establishDate = parties.establishDate || 'XXXX年XX月XX日';
  var registeredCapital = parties.registeredCapital || '【】';
  var mainBusiness = parties.mainBusiness || '【】';
  insertPara('1. 公司是一家依照《中华人民共和国公司法》及相关法律法规设立的有限责任公司，成立于' + establishDate + '，注册资本为人民币' + registeredCapital + '元。集团公司（定义见下文）主要从事' + mainBusiness + '（"主营业务"，为免疑义，如公司根据股东协议（定义见下文）规定的程序调整了集团公司主要从事的业务，则主营业务涵盖的内容应被视为自动更新以反映前述被调整后的业务）。', setBodyFormat);
  
  // 鉴于第2条 + 只有表头的表格
  insertPara('2. 于本协议签署之日，公司的股权结构如下：', setBodyFormat);
  
  // 插入只有表头的表格（1行3列）
  var table = newDoc.Tables.Add(selection.Range, 1, 3);
  table.Borders.Enable = true;
  
  // 表头
  table.Cell(1, 1).Range.Text = '股东名称/姓名';
  table.Cell(1, 2).Range.Text = '出资额（人民币元）';
  table.Cell(1, 3).Range.Text = '股权比例';
  
  for (var c = 1; c <= 3; c++) {
    var headerCell = table.Cell(1, c).Range;
    headerCell.Font.Name = '微软雅黑';
    headerCell.Font.Size = 9;
    headerCell.Font.Bold = true;
    headerCell.ParagraphFormat.Alignment = 1;
  }
  
  // 移动光标到表格后
  selection.MoveDown(5, 1);
  
  // 鉴于第3条
  var totalInvestment = parties.totalInvestment || '【】';
  var newCapital = parties.newCapital || '【】';
  insertPara('3. 各投资者拟按照本协议的条款和条件分别且不连带地对公司进行增资，合计以人民币' + totalInvestment + '元的价格认购公司新增的注册资本人民币' + newCapital + '元。', setBodyFormat);
  
  // 鉴于第4条
  insertPara('4. 公司和创始股东拟按照本协议的条款和条件接受投资者对公司的增资。', setBodyFormat);
  
  insertPara('为此，协议各方本着平等互利的原则，经友好协商，依据《中华人民共和国公司法》《中华人民共和国民法典》及中国其他有关法律和法规，就投资者向公司增资事宜达成以下协议。', setBodyFormat);
  
  // 鉴于条款后空1行，直接开始正文条款（不换页）
  insertEmptyLine();
  
  // ========== 正文条款（15条） ==========
  
  // 条款正文格式：左对齐、首行缩进2字符、不加粗
  function setClauseBodyFormat() {
    selection.Font.Name = '微软雅黑';
    selection.Font.Size = 9;
    selection.Font.Bold = false;
    selection.ParagraphFormat.Alignment = 0;
    selection.ParagraphFormat.LineSpacingRule = 4;
    selection.ParagraphFormat.LineSpacing = 25;
    selection.ParagraphFormat.FirstLineIndent = newDoc.Application.MillimetersToPoints(7);
    selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
    selection.ParagraphFormat.SpaceBefore = 0;
    selection.ParagraphFormat.SpaceAfter = 0;
  }
  
  if (clauses) {
    var clauseLines = clauses.split('\n');
    for (var k = 0; k < clauseLines.length; k++) {
      var line = clauseLines[k];
      var trimmedLine = line.trim();
      
      // 跳过空行
      if (!trimmedLine) {
        continue;
      }
      
      // 判断是否是条款标题（如"第一条 定义"、"第十五条 其他"）
      var isTitle = false;
      if (trimmedLine.indexOf('第') === 0 && trimmedLine.indexOf('条') > 0) {
        var tiaoPos = trimmedLine.indexOf('条');
        var afterTiao = trimmedLine.substring(tiaoPos + 1);
        // 标题格式："第X条 标题名"或"第X条"（标题后跟空格或结束）
        if (afterTiao.length === 0 || afterTiao.charAt(0) === ' ' || afterTiao.charAt(0) === '　') {
          isTitle = true;
        } else if (trimmedLine.length <= 20) {
          // 短行且以"第"开头包含"条"，也视为标题
          isTitle = true;
        }
      }
      
      if (isTitle) {
        // 标题：居中、加粗
        insertPara(trimmedLine, setClauseTitleFormat);
      } else {
        // 正文：左对齐、首行缩进2字符、不加粗
        insertPara(trimmedLine, setClauseBodyFormat);
      }
    }
  }
  
  // 正文条款结束后：空2行，写"以下无正文"，再换页
  insertEmptyLine();
  insertEmptyLine();
  insertPara('[以下无正文]', setBodyFormat);
  insertPageBreak();
  
  // ========== 签署页 ==========
  
  var signParties = [
    parties.investorA || '【投资人A投资实体名称】',
    parties.investorB || null,
    parties.companyName || '【标的公司名称】',
    parties.esopPlatform || null,
    parties.resourcePlatform || '【资源方持股平台】'
  ].filter(function(p) { return p; });
  
  for (var m = 0; m < signParties.length; m++) {
    insertPara('[本页无正文，为《增资协议》签署页]', setBodyFormat);
    insertPara('此证，本协议的每一方已促使其正式授权的代表于文首所载的日期签订本协议，以昭信守。', setBodyFormat);
    insertEmptyLine();
    insertEmptyLine();
    
    selection.Font.Name = '微软雅黑';
    selection.Font.Size = 9;
    selection.Font.Bold = true;
    selection.ParagraphFormat.FirstLineIndent = 0;
    selection.TypeText(signParties[m] + '：');
    selection.TypeParagraph();
    
    insertPara('（盖章）', setSignFormat);
    insertEmptyLine();
    insertPara('签署：___________________________________', setSignFormat);
    insertEmptyLine();
    insertPara('姓名：【】', setSignFormat);
    insertPara('职位：【】', setSignFormat);
    
    // 签署页之间分页（最后一个签署页不需要分页）
    if (m < signParties.length - 1) {
      insertPageBreak();
    }
  }
  
  // ========== 附件页（每个附件独占一页）==========
  
  // 附件一
  insertPageBreak();
  insertPara('附件一', setAttachmentTitleFormat);
  insertPara('股东协议格式', setAttachmentTitleFormat);
  
  // 附件二
  insertPageBreak();
  insertPara('附件二', setAttachmentTitleFormat);
  insertPara('公司章程格式', setAttachmentTitleFormat);
  
  // 附件三
  insertPageBreak();
  insertPara('附件三', setAttachmentTitleFormat);
  insertPara('核心员工名单', setAttachmentTitleFormat);
  
  // 附件四
  insertPageBreak();
  insertPara('附件四', setAttachmentTitleFormat);
  insertPara('披露清单', setAttachmentTitleFormat);
  
  var docName = newDoc.Name || '增资协议-正式版.docx';
  
  return { success: true, message: '增资协议正式版已生成', docName: docName };
  
} catch (e) {
  console.warn('[write-agreement-formal]', e);
  return { success: false, message: String(e) };
}
