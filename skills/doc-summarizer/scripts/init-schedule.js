/**
 * init-schedule.js
 * 初始化调度器任务，按功能类型分片。
 *
 * 入参: isSelection (boolean), funcType (string)
 * 出参: ScheduleInitResult
 */

try {
  var isSelection = typeof isSelection !== 'undefined' ? isSelection : false;
  var funcType = typeof funcType !== 'undefined' ? funcType : 'extract';

  var doc = Application.ActiveDocument;
  if (!doc || !doc.Paragraphs) {
    return {
      tasks: [],
      promptTemplate: '',
      system: '',
      concurrency: 1,
      continueOnError: true,
      errorBudget: 0.5
    };
  }

  // 获取文档段落
  var tasks = [];
  var totalParagraphs = doc.Paragraphs.Count;

  // 按段落分片，每10段为一个chunk
  var chunkSize = 10;
  var chunkCount = Math.ceil(totalParagraphs / chunkSize);

  for (var i = 0; i < chunkCount; i++) {
    var startPara = i * chunkSize + 1;
    var endPara = Math.min((i + 1) * chunkSize, totalParagraphs);
    var chunkText = '';

    for (var j = startPara; j <= endPara; j++) {
      var para = doc.Paragraphs.Item(j);
      if (para && para.Range && para.Range.Text) {
        chunkText = chunkText + para.Range.Text;
      }
    }

    tasks.push({
      chunkIndex: i + 1,
      totalChunks: chunkCount,
      funcType: funcType,
      text: chunkText
    });
  }

  // 根据功能类型设置不同的 system prompt
  var systemPrompt = '';
  if (funcType === 'extract') {
    systemPrompt = '你是内容提炼专家。对提供的文档片段进行关键信息提炼，生成简短摘要。\n\n## 执行规范\n- 禁止调用任何工具\n- 直接分析消息中提供的文本\n- 保留核心观点和要点，语言简明\n- 保持客观中立，不添加主观评论\n\n## 输出格式\n{"type":"extract","summary":"片段总结摘要","keyPoints":["要点1","要点2"],"tags":["标签1","标签2"]}\n\n无内容时输出：{"type":"extract","summary":"","keyPoints":[],"tags":[]}';
  } else if (funcType === 'minutes') {
    systemPrompt = '你是会议纪要专家。对提供的文档片段进行分析，提取会议关键信息。\n\n## 执行规范\n- 禁止调用任何工具\n- 直接分析消息中提供的文本\n- 提取会议主题、讨论要点、决策事项、待办事项、责任人及截止日期\n\n## 输出格式\n{"type":"minutes","meetingTitle":"会议主题","discussionPoints":["讨论要点1","讨论要点2"],"decisions":["决策1","决策2"],"actionItems":[{"task":"待办事项","owner":"责任人","deadline":"截止日期"}]}\n\n无内容时输出：{"type":"minutes","meetingTitle":"","discussionPoints":[],"decisions":[],"actionItems":[]}';
  } else if (funcType === 'weekly') {
    systemPrompt = '你是周报助手专家。对提供的文档片段进行分析，提取工作周报信息。\n\n## 执行规范\n- 禁止调用任何工具\n- 直接分析消息中提供的文本\n- 提取本周完成工作、本周亮点、下周计划、需协调事项\n- 如原文未提及下周计划，不要强行生成\n\n## 输出格式\n{"type":"weekly","completed":["完成工作1","完成工作2"],"highlights":["亮点1","亮点2"],"nextWeekPlans":["下周计划1","下周计划2"],"coordinationItems":["需协调事项1","需协调事项2"]}\n\n无内容时输出：{"type":"weekly","completed":[],"highlights":[],"nextWeekPlans":[],"coordinationItems":[]}';
  }

  return {
    tasks: tasks,
    promptTemplate: '文档片段（第 {{chunkIndex}} 段 / 共 {{totalChunks}} 段，功能类型：{{funcType}}）：\n{{text}}',
    system: systemPrompt,
    concurrency: 5,
    continueOnError: true,
    errorBudget: 0.5,
    aggregatorScript: 'scripts/aggregate-results.js'
  };
} catch (e) {
  console.warn('[init-schedule]', e);
  return {
    tasks: [],
    promptTemplate: '',
    system: '',
    concurrency: 1,
    continueOnError: true,
    errorBudget: 0.5
  };
}
