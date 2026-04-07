/**
 * aggregate-results.js
 * 汇总调度器子任务的结果，根据功能类型生成完整内容。
 *
 * 入参: results (array) - 所有子任务的返回结果
 * 出参: { success: boolean, content: string, title: string }
 */

try {
  var results = typeof results !== 'undefined' ? results : [];
  if (!results || results.length === 0) {
    return { success: false, content: '', title: '' };
  }

  // 获取功能类型（从第一个结果中获取）
  var funcType = 'extract';
  if (results.length > 0 && results[0] && results[0].type) {
    funcType = results[0].type;
  }

  var finalContent = '';
  var finalTitle = '';

  // 清理函数：移除 # 和 * 等符号
  function cleanText(text) {
    if (!text) return '';
    return text.replace(/[#*]/g, '').trim();
  }

  if (funcType === 'extract') {
    // 内容提炼聚合
    finalTitle = '内容提炼';

    var allSummaries = [];
    var allKeyPoints = [];
    var allTags = [];

    for (var i = 0; i < results.length; i++) {
      var result = results[i];
      if (result && result.summary) {
        allSummaries.push(cleanText(result.summary));
      }
      if (result && result.keyPoints && result.keyPoints.length > 0) {
        for (var j = 0; j < result.keyPoints.length; j++) {
          allKeyPoints.push(cleanText(result.keyPoints[j]));
        }
      }
      if (result && result.tags && result.tags.length > 0) {
        for (var k = 0; k < result.tags.length; k++) {
          var tag = cleanText(result.tags[k]);
          if (allTags.indexOf(tag) === -1) {
            allTags.push(tag);
          }
        }
      }
    }

    // 生成内容提炼
    if (allSummaries.length > 0) {
      for (var m = 0; m < allSummaries.length; m++) {
        if (allSummaries[m]) {
          finalContent = finalContent + allSummaries[m] + '\n';
        }
      }
      finalContent = finalContent + '\n';
    }

    if (allKeyPoints.length > 0) {
      finalContent = finalContent + '关键要点：\n';
      var maxPoints = allKeyPoints.length > 10 ? 10 : allKeyPoints.length;
      for (var n = 0; n < maxPoints; n++) {
        finalContent = finalContent + (n + 1) + '. ' + allKeyPoints[n] + '\n';
      }
    }

  } else if (funcType === 'minutes') {
    // 会议纪要聚合
    finalTitle = '会议纪要';

    var meetingTitles = [];
    var allDiscussionPoints = [];
    var allDecisions = [];
    var allActionItems = [];

    for (var p = 0; p < results.length; p++) {
      var r1 = results[p];
      if (r1 && r1.meetingTitle) {
        meetingTitles.push(cleanText(r1.meetingTitle));
      }
      if (r1 && r1.discussionPoints && r1.discussionPoints.length > 0) {
        for (var q = 0; q < r1.discussionPoints.length; q++) {
          allDiscussionPoints.push(cleanText(r1.discussionPoints[q]));
        }
      }
      if (r1 && r1.decisions && r1.decisions.length > 0) {
        for (var s = 0; s < r1.decisions.length; s++) {
          allDecisions.push(cleanText(r1.decisions[s]));
        }
      }
      if (r1 && r1.actionItems && r1.actionItems.length > 0) {
        for (var t = 0; t < r1.actionItems.length; t++) {
          allActionItems.push(r1.actionItems[t]);
        }
      }
    }

    // 生成会议纪要
    if (meetingTitles.length > 0) {
      finalContent = finalContent + '会议主题：' + meetingTitles[0] + '\n\n';
    }

    if (allDiscussionPoints.length > 0) {
      finalContent = finalContent + '讨论要点：\n';
      for (var u = 0; u < allDiscussionPoints.length; u++) {
        finalContent = finalContent + (u + 1) + '. ' + allDiscussionPoints[u] + '\n';
      }
      finalContent = finalContent + '\n';
    }

    if (allDecisions.length > 0) {
      finalContent = finalContent + '决策事项：\n';
      for (var v = 0; v < allDecisions.length; v++) {
        finalContent = finalContent + (v + 1) + '. ' + allDecisions[v] + '\n';
      }
      finalContent = finalContent + '\n';
    }

    if (allActionItems.length > 0) {
      finalContent = finalContent + '待办事项：\n';
      for (var w = 0; w < allActionItems.length; w++) {
        var item = allActionItems[w];
        var taskText = cleanText(item.task || '');
        var ownerText = cleanText(item.owner || '');
        var deadlineText = cleanText(item.deadline || '');
        if (taskText) {
          var actionText = '- ' + taskText;
          if (ownerText || deadlineText) {
            actionText = actionText + ' (' + ownerText + (ownerText && deadlineText ? '，' : '') + deadlineText + ')';
          }
          finalContent = finalContent + actionText + '\n';
        }
      }
    }

  } else if (funcType === 'weekly') {
    // 周报助手聚合
    finalTitle = '工作周报';

    var allCompleted = [];
    var allHighlights = [];
    var allNextWeekPlans = [];
    var allCoordinationItems = [];

    for (var x = 0; x < results.length; x++) {
      var r2 = results[x];
      if (r2 && r2.completed && r2.completed.length > 0) {
        for (var y = 0; y < r2.completed.length; y++) {
          allCompleted.push(cleanText(r2.completed[y]));
        }
      }
      if (r2 && r2.highlights && r2.highlights.length > 0) {
        for (var z = 0; z < r2.highlights.length; z++) {
          allHighlights.push(cleanText(r2.highlights[z]));
        }
      }
      if (r2 && r2.nextWeekPlans && r2.nextWeekPlans.length > 0) {
        for (var aa = 0; aa < r2.nextWeekPlans.length; aa++) {
          allNextWeekPlans.push(cleanText(r2.nextWeekPlans[aa]));
        }
      }
      if (r2 && r2.coordinationItems && r2.coordinationItems.length > 0) {
        for (var ab = 0; ab < r2.coordinationItems.length; ab++) {
          allCoordinationItems.push(cleanText(r2.coordinationItems[ab]));
        }
      }
    }

    // 生成工作周报
    if (allCompleted.length > 0) {
      finalContent = finalContent + '一、本周完成工作\n';
      for (var ac = 0; ac < allCompleted.length; ac++) {
        finalContent = finalContent + (ac + 1) + '. ' + allCompleted[ac] + '\n';
      }
      finalContent = finalContent + '\n';
    }

    if (allHighlights.length > 0) {
      finalContent = finalContent + '二、本周亮点\n';
      for (var ad = 0; ad < allHighlights.length; ad++) {
        finalContent = finalContent + (ad + 1) + '. ' + allHighlights[ad] + '\n';
      }
      finalContent = finalContent + '\n';
    }

    if (allNextWeekPlans.length > 0) {
      finalContent = finalContent + '三、下周计划\n';
      for (var ae = 0; ae < allNextWeekPlans.length; ae++) {
        finalContent = finalContent + (ae + 1) + '. ' + allNextWeekPlans[ae] + '\n';
      }
      finalContent = finalContent + '\n';
    }

    if (allCoordinationItems.length > 0) {
      finalContent = finalContent + '四、需协调事项\n';
      for (var af = 0; af < allCoordinationItems.length; af++) {
        finalContent = finalContent + (af + 1) + '. ' + allCoordinationItems[af] + '\n';
      }
    }
  }

  return {
    success: true,
    content: finalContent,
    title: finalTitle
  };
} catch (e) {
  console.warn('[aggregate-results]', e);
  return { success: false, content: '', title: '' };
}
