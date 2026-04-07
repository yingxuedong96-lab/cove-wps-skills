/**
 * apply-page-setup.js
 * 应用页面设置（页边距、纸张大小等）
 *
 * 入参: config (JSON配置字符串或对象)
 * 出参: { success, applied, error }
 */

try {
  var configData = typeof config === 'string' ? JSON.parse(config) : config;
  var docInfo = configData.documentInfo || {};

  var doc = Application.ActiveDocument;
  if (!doc) {
    return { success: false, error: '无活动文档', applied: false };
  }

  var applied = false;
  var changes = [];

  // 页边距设置
  if (docInfo.pageMargins) {
    var margins = docInfo.pageMargins;
    var sections = doc.Sections;

    for (var s = 1; s <= sections.Count; s++) {
      var ps = sections.Item(s).PageSetup;

      if (margins.top !== undefined) {
        try {
          ps.TopMargin = margins.top;
          changes.push('上边距=' + margins.top);
        } catch (e) {}
      }
      if (margins.bottom !== undefined) {
        try {
          ps.BottomMargin = margins.bottom;
          changes.push('下边距=' + margins.bottom);
        } catch (e) {}
      }
      if (margins.left !== undefined) {
        try {
          ps.LeftMargin = margins.left;
          changes.push('左边距=' + margins.left);
        } catch (e) {}
      }
      if (margins.right !== undefined) {
        try {
          ps.RightMargin = margins.right;
          changes.push('右边距=' + margins.right);
        } catch (e) {}
      }
      if (margins.headerDistance !== undefined) {
        try {
          ps.HeaderDistance = margins.headerDistance;
          changes.push('页眉距边界=' + margins.headerDistance);
        } catch (e) {}
      }
      if (margins.footerDistance !== undefined) {
        try {
          ps.FooterDistance = margins.footerDistance;
          changes.push('页脚距边界=' + margins.footerDistance);
        } catch (e) {}
      }
    }
    applied = true;
  }

  // 纸张大小设置
  if (docInfo.pageSize) {
    var size = docInfo.pageSize;
    var sections = doc.Sections;

    for (var s = 1; s <= sections.Count; s++) {
      var ps = sections.Item(s).PageSetup;

      if (size.width !== undefined) {
        try {
          ps.PageWidth = size.width;
          changes.push('纸张宽度=' + size.width);
        } catch (e) {}
      }
      if (size.height !== undefined) {
        try {
          ps.PageHeight = size.height;
          changes.push('纸张高度=' + size.height);
        } catch (e) {}
      }
    }
    applied = true;
  }

  // 方向设置
  if (docInfo.orientation !== undefined) {
    var sections = doc.Sections;
    for (var s = 1; s <= sections.Count; s++) {
      try {
        sections.Item(s).PageSetup.Orientation = docInfo.orientation;
        changes.push('纸张方向=' + (docInfo.orientation === 0 ? '纵向' : '横向'));
      } catch (e) {}
    }
    applied = true;
  }

  console.log('[page-setup] 应用成功: ' + changes.join(', '));

  return {
    success: true,
    applied: applied,
    changes: changes
  };

} catch (e) {
  console.error('[page-setup] 错误: ' + e);
  return {
    success: false,
    error: String(e),
    applied: false
  };
}