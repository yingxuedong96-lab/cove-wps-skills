/**
 * save-template.js - 从最近的 artifact 中提取模板并保存到指定目录
 * 版本: 26.0410.1003
 * 用法: 提取模板后说「保存模板」
 */
try {
  var VER = "26.0410.1003";
  console.log("[save-template] 版本: " + VER);

  var templateDir = "/Users/cassia/Desktop/dyx/wpsjs/模版生成/";

  // 找到最近的包含 templateJson 的 artifact
  var artifactDir = "/Users/cassia/Library/Containers/com.kingsoft.wpsoffice.mac/Data/cove-wps/artifacts/";

  // 此脚本需要 Python 端配合执行
  // 返回提示信息让 Python 处理

  return {
    success: true,
    message: "请执行以下 Python 命令保存模板：\n\npython3 save_template_from_artifact.py",
    needPython: true,
    templateDir: templateDir
  };

} catch (e) {
  return { success: false, error: String(e) };
}