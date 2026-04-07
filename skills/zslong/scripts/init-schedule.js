/**
 * init-schedule.js
 * 调度器初始化脚本，用于大文档分片处理。
 *
 * ⚠️ 编号类scope不能分片处理，直接返回空任务列表！
 *
 * 入参: fixMode (string) - 修复模式
 *       scope (string) - 校对范围
 * 出参: ScheduleInitResult
 */

try {
  // 参数接收
  var mode = typeof fixMode !== 'undefined' ? fixMode : 'standard';
  var scopeType = typeof scope !== 'undefined' ? scope : 'full';

  // ⚠️ 编号类scope不能分片，直接返回空任务
  var numberingScopes = ['numbering', 'heading', 'table', 'figure', 'formula', 'numbering-proof', 'full_proofread', 'value', 'table_content', 'font', 'figure_table_layout', 'figure_layout', 'table_layout', 'figure_caption', 'table_caption', 'figure_center', 'formula_layout', 'header_footer', 'page_setup'];
  // 检查 scope 或 taskType 参数
  var checkScope = scopeType || (typeof taskType !== 'undefined' ? taskType : '');
  for (var i = 0; i < numberingScopes.length; i++) {
    if (checkScope === numberingScopes[i] || checkScope.indexOf('numbering') === 0) {
      console.log('[init-schedule] 编号类scope不能分片处理: ' + checkScope);
      return {
        tasks: [],
        promptTemplate: '',
        system: '编号校对需要全局上下文，不能分片处理。',
        concurrency: 1,
        continueOnError: false,
        errorBudget: 0,
        taskCallbackScript: '',
        aggregatorScript: '',
        error: '编号类不支持调度器'
      };
    }
  }

  var doc = Application.ActiveDocument;
  if (!doc) {
    return {
      tasks: [],
      promptTemplate: '',
      system: '',
      concurrency: 5,
      continueOnError: true,
      errorBudget: 0.5,
      taskCallbackScript: '',
      aggregatorScript: ''
    };
  }

  // 获取文档内容
  var content = '';
  try {
    if (scopeType === 'selection') {
      var sel = Application.Selection;
      if (sel && sel.Range) {
        content = sel.Range.Text || '';
      }
    }
    if (!content) {
      content = doc.Content && doc.Content.Text ? doc.Content.Text : '';
    }
  } catch (e) {
    content = '';
  }

  // 分片处理
  var CHUNK_SIZE = 3000; // 每片约3000字符
  var chunks = [];
  var lines = content.split(/\r\n|\r|\n/);
  var currentChunk = '';
  var chunkIndex = 1;

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (currentChunk.length + line.length > CHUNK_SIZE && currentChunk.length > 0) {
      chunks.push({
        chunkIndex: chunkIndex,
        totalChunks: 0, // 后面更新
        text: currentChunk.trim()
      });
      currentChunk = line;
      chunkIndex++;
    } else {
      currentChunk += (currentChunk ? '\n' : '') + line;
    }
  }

  // 添加最后一块
  if (currentChunk.trim()) {
    chunks.push({
      chunkIndex: chunkIndex,
      totalChunks: 0,
      text: currentChunk.trim()
    });
  }

  // 更新 totalChunks
  var totalChunks = chunks.length;
  for (var j = 0; j < chunks.length; j++) {
    chunks[j].totalChunks = totalChunks;
  }

  // 返回调度器配置
  return {
    tasks: chunks,

    promptTemplate: '文档片段（第 {{chunkIndex}} 段 / 共 {{totalChunks}} 段）：\n\n{{text}}\n\n请检查上述文本中的格式和内容问题，输出JSON数组格式的问题列表。',

    system: '你是一名技术报告校对员。按照Q/BIDR-G-JS00-101-002-2017规范，检查文档片段中的格式和内容问题。\n\n## 检查规则\n- V-006: 温度偏差格式（如 20℃±2℃ 应为 20±2℃）\n- V-008: 数值范围应使用波浪号（如 10-15 应为 10～15）\n- M-001/002: 图表名称末尾不应有标点\n- M-005: 中文内容应使用中文括号\n- T-006: 表格中不应使用"同上"或"同左"\n\n## 执行规范\n- 禁止调用任何工具\n- 直接分析消息中提供的文本\n\n## 输出格式\n[{"index":段号,"rule":"规则ID","original":"原文","suggested":"建议","reason":"原因","autoFix":是否可自动修复}]\n\n无问题时输出：[]',

    concurrency: 5,
    continueOnError: true,
    errorBudget: 0.5,

    taskCallbackScript: 'scripts/apply-comments.js',

    aggregatorScript: 'scripts/aggregate-issues.js'
  };

} catch (e) {
  console.warn('[init-schedule]', e);
  return {
    tasks: [],
    promptTemplate: '',
    system: '',
    concurrency: 5,
    continueOnError: true,
    errorBudget: 0.5,
    taskCallbackScript: '',
    aggregatorScript: '',
    error: String(e)
  };
}
