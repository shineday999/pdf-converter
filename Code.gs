/**
 * Code.gs
 * 主控：Web App 入口 + HTML include 載入器
 */

/**
 * Web App 入口
 */
function doGet() {
  return HtmlService.createTemplateFromFile('HTML')
    .evaluate()
    .setTitle('Sabre PDF Converter')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * HTML 模板 include 載入器
 * 用法：<?!= include('StyleTheme') ?>
 * @param {string} filename
 * @returns {string}
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 取得目前登入使用者 Email（給前端預填收件人用）
 * @returns {string}
 */
function getCurrentUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    return '';
  }
}
