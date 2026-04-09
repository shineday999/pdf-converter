/**
 * Config.gs
 * 讀取 GAS Script Properties（機敏資料不寫死在程式碼）
 *
 * Script Properties 需在 GAS 專案設定中手動新增：
 *   - OUTPUT_FOLDER_ID  : Google Drive 輸出資料夾 ID（可選）
 *   - DEFAULT_RECIPIENT : 預設收件人 Email（可選）
 *   - LOG_SHEET_ID      : 記錄用 Google Sheet ID（必填）
 */

/**
 * 讀取單一 Script Property
 * @param {string} key
 * @returns {string|null}
 */
function getConfig(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * 讀取所有 Script Properties（除錯用，勿在生產環境暴露）
 * @returns {Object}
 */
function getAllConfig() {
  return PropertiesService.getScriptProperties().getProperties();
}
