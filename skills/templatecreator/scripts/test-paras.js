/**
 * 最小测试脚本 - 输出前10个段落的原始信息
 */
(function() {
  const DOC = Application.ActiveDocument;
  if (!DOC) return JSON.stringify({ error: "no doc" });

  const paras = DOC.Paragraphs;
  const result = [];

  for (let i = 1; i <= Math.min(10, paras.Count); i++) {
    const para = paras.Item(i);
    const rawText = para.Range.Text;
    const text = String(rawText || '').replace(/\u0007/g, '').trim();

    result.push({
      idx: i,
      rawLen: rawText ? rawText.length : 0,
      cleanLen: text.length,
      text: text.substring(0, 40),
      firstChars: text.substring(0, 5).split('').map(c => c.charCodeAt(0)).join(',')
    });
  }

  return JSON.stringify({
    totalParas: paras.Count,
    samples: result
  }, null, 2);
})();