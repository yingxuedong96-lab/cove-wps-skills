/**
 * list-templates.js - 列出可用的样式模板
 *
 * 返回值：
 * - success: true/false
 * - templates: 模板列表 [{name, path, extractedFrom, stylesCount, extractedAt}]
 * - message: 提示信息
 */

(function() {
  const skillPath = Application.Env?.SkillPath || '';
  const templatesDir = `${skillPath}/templates`;

  const templates = [];

  try {
    // 尝试列出文件
    const files = Application.ListFiles?.(templatesDir) || [];

    for (const file of files) {
      if (!file.endsWith('.json')) continue;

      try {
        const content = Application.LoadFile?.(`${templatesDir}/${file}`);
        if (content) {
          const template = JSON.parse(content);
          templates.push({
            name: template.name || file.replace('.json', ''),
            path: `templates/${file}`,
            extractedFrom: template.extractedFrom || '未知',
            stylesCount: template.styles?.length || 0,
            extractedAt: template.extractedAt || '未知'
          });
        }
      } catch (e) {
        // 解析失败，跳过
      }
    }
  } catch (e) {
    // 目录不存在或无法访问
    return JSON.stringify({
      success: true,
      templates: [],
      message: "模板目录不存在或为空，请先提取模板"
    }, null, 2);
  }

  if (templates.length === 0) {
    return JSON.stringify({
      success: true,
      templates: [],
      message: "暂无保存的模板，请先从文档提取样式模板"
    }, null, 2);
  }

  // 按提取日期排序（新的在前）
  templates.sort((a, b) => {
    if (a.extractedAt > b.extractedAt) return -1;
    if (a.extractedAt < b.extractedAt) return 1;
    return 0;
  });

  return JSON.stringify({
    success: true,
    templates: templates,
    message: `找到${templates.length}个模板`
  }, null, 2);

})();