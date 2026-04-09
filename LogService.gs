/**
 * LogService.gs
 * 靜默記錄使用紀錄至 Google Sheets
 * 透過 try-catch 確保記錄失敗不影響主流程
 */

var LOG_SHEET_NAME = 'ConversionLog';
var LOG_HEADERS = [
  'Timestamp', 'User Email', 'File Count', 'File Names',
  'Output Formats', 'DPI Setting', 'Page Range', 'Total Pages',
  'Output Size (KB)', 'Recipient Email', 'Status', 'Error Message', 'Duration (s)'
];

/**
 * 取得或建立記錄 Sheet
 * @returns {Sheet|null}
 */
function _getOrCreateLogSheet() {
  var sheetId = getConfig('LOG_SHEET_ID');
  if (!sheetId) return null;

  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(LOG_HEADERS);
    sheet.getRange(1, 1, 1, LOG_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#FF6B00')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 記錄一筆轉換紀錄（靜默執行）
 * @param {Object} data
 * @param {string[]}  data.fileNames
 * @param {string[]}  data.formats
 * @param {number}    data.dpi
 * @param {string}    data.pageRange
 * @param {number}    data.totalPages
 * @param {number}    data.outputSizeKb
 * @param {string}    data.recipientEmail
 * @param {string}    data.status         'success' | 'error'
 * @param {string}    [data.errorMessage]
 * @param {number}    data.durationSeconds
 */
function logConversion(data) {
  try {
    var sheet = _getOrCreateLogSheet();
    if (!sheet) return;

    var userEmail = '';
    try { userEmail = Session.getActiveUser().getEmail(); } catch(e) {}

    var ts = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      ts,
      userEmail,
      data.fileNames ? data.fileNames.length : 0,
      (data.fileNames || []).join(', '),
      (data.formats || []).join(', '),
      data.dpi || '',
      data.pageRange || 'all',
      data.totalPages || 0,
      data.outputSizeKb || 0,
      data.recipientEmail || '',
      data.status || '',
      data.errorMessage || '',
      data.durationSeconds || 0
    ]);
  } catch (e) {
    // 靜默失敗 — 不影響任何主流程
  }
}
