/**
 * semantic-check.js
 * 接收语义检查结果并添加批注
 *
 * 入参:
 *   issues (array) - 语义检查发现的问题列表
 *     [{ rule, name, location, original, suggestion, message }]
 *   或者直接传入 result (array) - 兼容不同调用方式
 *
 * 出参: { commented, issues, summary }
 */

try {
  console.log('[semantic-check] 脚本开始执行');

  // 兼容多种参数名称
  var issueList = [];
  if (typeof issues !== 'undefined') {
    issueList = issues;
    console.log('[semantic-check] 收到 issues 参数');
  } else if (typeof result !== 'undefined') {
    issueList = result;
    console.log('[semantic-check] 收到 result 参数');
  } else {
    console.log('[semantic-check] 未收到 issues 或 result 参数');
  }

  // 如果是字符串，尝试解析为 JSON
  if (typeof issueList === 'string') {
    try {
      issueList = JSON.parse(issueList);
      console.log('[semantic-check] 字符串参数已解析为 JSON');
    } catch (e) {
      console.log('[semantic-check] JSON 解析失败: ' + e);
      issueList = [];
    }
  }

  if (!Array.isArray(issueList) || issueList.length === 0) {
    console.log('[semantic-check] 没有发现问题');
    return {
      commented: 0,
      issues: [],
      summary: '✅ 语义检查完成，未发现问题。'
    };
  }

  console.log('[semantic-check] 收到 ' + issueList.length + ' 个问题');

  var doc = Application.ActiveDocument;
  if (!doc) {
    console.log('[semantic-check] 无活动文档');
    return { error: '无活动文档', commented: 0, issues: [], summary: '❌ 无活动文档' };
  }

  var commented = 0;
  var byRule = {}; // 按规则统计

  for (var i = 0; i < issueList.length; i++) {
    var issue = issueList[i];
    var rule = issue.rule || 'L-000';
    var name = issue.name || '语义问题';
    var location = issue.location || '';
    var original = issue.original || '';
    var suggestion = issue.suggestion || '';
    var message = issue.message || '';

    // 统计
    byRule[rule] = (byRule[rule] || 0) + 1;

    // 构建批注内容
    var commentText = '[' + rule + '] ' + name;
    if (original) {
      commentText += '\n原文：' + original;
    }
    if (suggestion) {
      commentText += '\n建议：' + suggestion;
    }
    if (message) {
      commentText += '\n说明：' + message;
    }

    // 优先通过原文内容定位（更可靠）
    var foundByContent = false;
    if (original && original.length >= 2) {
      try {
        // 使用 Find 在文档中搜索原文
        var findRange = doc.Content;
        findRange.Find.ClearFormatting();
        findRange.Find.Text = original;
        findRange.Find.Forward = true;
        findRange.Find.Wrap = 1; // wdFindContinue
        findRange.MatchCase = false;
        findRange.MatchWholeWord = false;

        if (findRange.Find.Execute()) {
          // 找到了，在找到的位置添加批注
          doc.Comments.Add(findRange, commentText);
          commented++;
          foundByContent = true;
          console.log('[semantic-check] 通过内容搜索定位成功: ' + rule);
          continue;
        }
      } catch (e) {
        console.log('[semantic-check] 内容搜索失败: ' + e);
      }
    }

    // 如果内容搜索失败，尝试通过段落编号定位
    if (!foundByContent) {
      try {
        // 解析位置信息
        var paraIndex = 0;
        var paraMatch1 = location.match(/第\s*(\d+)\s*段/);
        var paraMatch2 = location.match(/段落\s*(\d+)/);
        var paraMatch3 = location.match(/(\d+)/);

        if (paraMatch1) {
          paraIndex = parseInt(paraMatch1[1]);
        } else if (paraMatch2) {
          paraIndex = parseInt(paraMatch2[1]);
        } else if (paraMatch3) {
          paraIndex = parseInt(paraMatch3[1]);
        }

        console.log('[semantic-check] 解析位置 "' + location + '" → 段落 ' + paraIndex);

        if (paraIndex > 0 && paraIndex <= doc.Paragraphs.Count) {
          var para = doc.Paragraphs.Item(paraIndex);
          var range = para.Range;
          doc.Comments.Add(range, commentText);
          commented++;
          console.log('[semantic-check] 在段落 ' + paraIndex + ' 添加批注: ' + rule);
        } else {
          console.log('[semantic-check] 无法定位段落: ' + location);
        }
      } catch (e) {
        console.log('[semantic-check] 添加批注失败: ' + e);
      }
    }
  }

  console.log('[semantic-check] 完成，共添加 ' + commented + ' 条批注');

  // 生成用户友好的摘要
  var summaryLines = [];
  summaryLines.push('✅ 语义检查完成，发现 ' + issueList.length + ' 个问题，已添加 ' + commented + ' 条批注\n');

  // 按规则统计
  summaryLines.push('**问题统计**：');
  for (var r in byRule) {
    if (byRule.hasOwnProperty(r)) {
      var ruleName = getRuleName(r);
      summaryLines.push('- ' + r + ' ' + ruleName + ': ' + byRule[r] + ' 处');
    }
  }

  // 列出主要问题（最多5条）
  summaryLines.push('\n**主要问题**：');
  var showCount = Math.min(issueList.length, 5);
  for (var j = 0; j < showCount; j++) {
    var iss = issueList[j];
    var rName = iss.name || '问题';
    var loc = iss.location || '';
    var orig = iss.original || '';
    if (orig.length > 50) {
      orig = orig.substring(0, 50) + '...';
    }
    summaryLines.push((j + 1) + '. 【' + (iss.rule || '') + ' ' + rName + '】' + loc);
    if (orig) {
      summaryLines.push('   原文：' + orig);
    }
  }

  if (issueList.length > 5) {
    summaryLines.push('\n...共 ' + issueList.length + ' 个问题，详见文档批注');
  }

  return {
    commented: commented,
    summary: summaryLines.join('\n')
  };

} catch (e) {
  console.warn('[semantic-check]', e);
  return { error: String(e), commented: 0, issues: [], summary: '❌ 执行出错：' + String(e) };
}

// 规则名称映射
function getRuleName(rule) {
  var names = {
    'L-001': '术语一致性',
    'L-002': '模糊引用',
    'L-003': '图表引用缺失',
    'L-004': '繁体字检测'
  };
  return names[rule] || '其他问题';
}