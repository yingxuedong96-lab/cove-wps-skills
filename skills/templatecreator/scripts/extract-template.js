/**
 * extract-template.js - 使用 Content.Text 获取段落（与其他脚本一致）
 * 版本: 26.0408.1445
 */

(function() {
  const SCRIPT_VERSION = "26.0408.1445";
  console.log("[extract-template] 脚本版本: " + SCRIPT_VERSION);

  const DOC = Application.ActiveDocument;
  if (!DOC) {
    return JSON.stringify({ success: false, error: "没有打开的文档" });
  }

  // 清理文本（移除WPS特殊字符）
  function cleanText(text) {
    return String(text || '').replace(/\u0007/g, '').replace(/[\r\n]/g, '').trim();
  }

  const STYLE_SPEC = {
    paper: {
      name: "论文报告样式",
      tags: [
        { id: "heading5", name: "五级标题", pattern: /^\\d+\\.\\d+\\.\\d+\\.\\d+\\.\\d+/ },
        { id: "heading4", name: "四级标题", pattern: /^\\d+\\.\\d+\\.\\d+\\.\\d+/ },
        { id: "heading3", name: "三级标题", pattern: /^\\d+\\.\\d+\\.\\d+/ },
        { id: "heading2", name: "二级标题", pattern: /^\\d+\\.\\d+[^\\.\\d]/ },
        { id: "heading1", name: "一级标题", pattern: /^\\d+[^\\.\\d]/ },
        { id: "chapterTitle", name: "章标题", pattern: /^第[一二三四五六七八九十\\d]+章/ },
        { id: "abstractTitle", name: "摘要标题", pattern: /^摘要|^Abstract/ },
        { id: "keywords", name: "关键词", pattern: /^关键词|^Keywords/ },
        { id: "tocTitle", name: "目录标题", pattern: /^目\\s*录$|^目次$/ },
        { id: "figureCaption", name: "图名", pattern: /^图\\s*\\d+/ },
        { id: "tableCaption", name: "表名", pattern: /^表\\s*\\d+/ },
        { id: "appendixTitle", name: "附录标题", pattern: /^附\\s*录/ },
        { id: "appendixSection", name: "附录节题", pattern: /^[A-Z]\\.\\d+/ },
        { id: "referenceTitle", name: "参考文献标题", pattern: /^参考文献/ },
        { id: "reference", name: "参考文献条目", pattern: /^\\[\\d+\\]/ },
        { id: "body", name: "正文", pattern: null }
      ]
    },
    official: {
      name: "公文样式",
      tags: [
        { id: "heading1", name: "一级标题", pattern: /^[一二三四五六七八九十]+、/ },
        { id: "heading2", name: "二级标题", pattern: /^\\([一二三四五六七八九十]+\\)/ },
        { id: "heading3", name: "三级标题", pattern: /^\\d+\\.\\s/ },
        { id: "body", name: "正文", pattern: null }
      ]
    }
  };

  const params = Application.Env?.ScriptParams || {};

  if (!params.docType) {
    return JSON.stringify({
      success: true,
      needUserInput: true,
      stage: "selectDocType",
      question: "请选择文档类型：",
      options: ["论文/技术报告", "公文"]
    }, null, 2);
  }

  const docType = params.docType === "论文/技术报告" || params.docType === "paper" ? "paper" : "official";
  const spec = STYLE_SPEC[docType];

  // 使用 doc.Content.Text 获取全部文本，然后分割
  const docText = DOC.Content && DOC.Content.Text ? String(DOC.Content.Text) : '';
  const paras = docText.split('\r');

  console.log("[extract-template] 总段落数: " + paras.length);

  // 收集前20个段落的调试信息
  const debugParas = [];
  const results = { matched: {} };

  for (let i = 0; i < paras.length; i++) {
    const rawText = paras[i];
    const text = cleanText(rawText);
    if (!text) continue;

    // 记录前20个段落
    if (debugParas.length < 20) {
      const firstChars = [];
      for (let j = 0; j < Math.min(8, text.length); j++) {
        firstChars.push(text.charCodeAt(j));
      }

      // 测试匹配
      let matchResult = "未匹配";
      for (const tag of spec.tags) {
        if (tag.pattern) {
          try {
            if (tag.pattern.test(text)) {
              matchResult = tag.name;
              break;
            }
          } catch (e) {}
        }
      }

      debugParas.push({
        idx: i + 1,
        text: text.substring(0, 30),
        codes: firstChars.join(","),
        match: matchResult
      });
    }

    // 检测标签
    let detection = null;
    for (const tag of spec.tags) {
      if (tag.pattern) {
        try {
          if (tag.pattern.test(text)) {
            detection = tag.id;
            break;
          }
        } catch (e) {}
      }
    }

    // 默认为正文
    if (!detection) {
      detection = "body";
    }

    if (!results.matched[detection]) {
      results.matched[detection] = { count: 0, samples: [] };
    }
    results.matched[detection].count++;
    if (results.matched[detection].samples.length < 3) {
      results.matched[detection].samples.push(text.substring(0, 50));
    }
  }

  // 生成样式列表
  const styles = [];
  for (const tag of spec.tags) {
    const data = results.matched[tag.id];
    if (data && data.count > 0) {
      styles.push({
        name: tag.name,
        count: data.count,
        samples: data.samples
      });
    }
  }

  // 生成输出
  const lines = [];
  lines.push("✅ 样式模板提取完成！版本: " + SCRIPT_VERSION);
  lines.push("");
  lines.push("══════════════════════════════════════════════════");
  lines.push("【调试信息】前20个段落检测情况");
  lines.push("══════════════════════════════════════════════════");
  debugParas.forEach(p => {
    lines.push(`[${p.idx}] "${p.text}"`);
    lines.push(`   字符码: ${p.codes} | 匹配: ${p.match}`);
  });
  lines.push("══════════════════════════════════════════════════");
  lines.push("");
  lines.push("📄 源文档：" + DOC.Name);
  lines.push("📊 共 " + styles.length + " 种样式，" + paras.length + " 个段落");
  lines.push("");

  lines.push("## 提取的样式详情");
  lines.push("");
  styles.forEach(s => {
    lines.push("### " + s.name + "（" + s.count + "处）");
    if (s.samples.length > 0) {
      lines.push("示例: \"" + s.samples[0] + "\"");
    }
    lines.push("");
  });

  return JSON.stringify({
    success: true,
    scriptVersion: SCRIPT_VERSION,
    message: lines.join("\n"),
    styles: styles,
    debugParas: debugParas
  }, null, 2);

})();