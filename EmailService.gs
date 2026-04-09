/**
 * EmailService.gs
 * 使用 GmailApp 寄送轉換結果 Email
 */

var EMAIL_SUBJECT_PREFIX = '📄 Sabre PDF Converter — ';
var SMALL_FILE_LIMIT_BYTES = 20 * 1024 * 1024; // 20MB 以下直接附件

/**
 * 寄送轉換完成通知 Email
 * @param {string} recipientEmail       收件人
 * @param {Array<{fileId, webViewLink, fileName, size}>} fileInfoArray  已上傳至 Drive 的檔案資訊
 * @param {string} sourcePdfName        來源 PDF 名稱
 * @param {string} folderId             Drive 資料夾 ID（提供總資料夾連結）
 * @returns {{success: boolean, message: string}}
 */
function sendConversionResult(recipientEmail, fileInfoArray, sourcePdfName, folderId) {
  try {
    var folderUrl = 'https://drive.google.com/drive/folders/' + folderId;
    var subject = EMAIL_SUBJECT_PREFIX + sourcePdfName + ' 轉換完成';
    var body = buildEmailBody(fileInfoArray, sourcePdfName, folderUrl);

    // 小檔（< 20MB）直接附件，大檔只給 Drive 連結
    var attachments = [];
    fileInfoArray.forEach(function(info) {
      if (info.size < SMALL_FILE_LIMIT_BYTES) {
        try {
          var file = DriveApp.getFileById(info.fileId);
          attachments.push(file.getBlob());
        } catch (e) {
          // 附件失敗不中斷主流程
        }
      }
    });

    GmailApp.sendEmail(recipientEmail, subject, '', {
      htmlBody: body,
      attachments: attachments,
      name: 'Sabre PDF Converter'
    });

    return { success: true, message: 'Email 已寄送至 ' + recipientEmail };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * 組裝 HTML Email 內容
 * @param {Array} fileInfoArray
 * @param {string} sourcePdfName
 * @param {string} folderUrl
 * @returns {string} HTML
 */
function buildEmailBody(fileInfoArray, sourcePdfName, folderUrl) {
  var rows = fileInfoArray.map(function(f) {
    var sizeStr = f.size > 1024 * 1024
      ? (f.size / 1024 / 1024).toFixed(2) + ' MB'
      : (f.size / 1024).toFixed(1) + ' KB';
    return '<tr>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #E5E7EB;">' +
        '<a href="' + f.webViewLink + '" style="color:#FF6B00;text-decoration:none;font-weight:600;">' + f.fileName + '</a>' +
      '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #E5E7EB;color:#6B7280;font-size:13px;">' + sizeStr + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #E5E7EB;">' +
        '<a href="' + f.webViewLink + '" style="color:#FF6B00;font-size:13px;">開啟</a>' +
      '</td>' +
    '</tr>';
  }).join('');

  return '<!DOCTYPE html><html><body style="font-family:Inter,sans-serif;background:#F8F9FC;margin:0;padding:20px;">' +
    '<div style="max-width:600px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.06);">' +
      '<div style="background:#FF6B00;padding:24px 32px;">' +
        '<h1 style="color:#fff;margin:0;font-size:20px;font-weight:700;">🧭 Sabre PDF Converter</h1>' +
        '<p style="color:rgba(255,255,255,0.85);margin:4px 0 0;font-size:14px;">轉換完成通知</p>' +
      '</div>' +
      '<div style="padding:28px 32px;">' +
        '<p style="color:#1A1F36;margin:0 0 20px;">您的 PDF 檔案 <strong>' + sourcePdfName + '</strong> 已完成轉換，共 ' + fileInfoArray.length + ' 個輸出檔案：</p>' +
        '<table style="width:100%;border-collapse:collapse;border:1px solid #E5E7EB;border-radius:8px;overflow:hidden;">' +
          '<thead>' +
            '<tr style="background:#F8F9FC;">' +
              '<th style="padding:10px 12px;text-align:left;font-size:13px;color:#6B7280;font-weight:600;">檔案名稱</th>' +
              '<th style="padding:10px 12px;text-align:left;font-size:13px;color:#6B7280;font-weight:600;">大小</th>' +
              '<th style="padding:10px 12px;text-align:left;font-size:13px;color:#6B7280;font-weight:600;">操作</th>' +
            '</tr>' +
          '</thead>' +
          '<tbody>' + rows + '</tbody>' +
        '</table>' +
        '<div style="margin-top:20px;text-align:center;">' +
          '<a href="' + folderUrl + '" style="display:inline-block;background:#FF6B00;color:#fff;padding:12px 28px;border-radius:8px;text-decoration:none;font-weight:600;font-size:14px;">📁 開啟 Drive 資料夾</a>' +
        '</div>' +
      '</div>' +
      '<div style="padding:16px 32px;background:#F8F9FC;border-top:1px solid #E5E7EB;text-align:center;">' +
        '<p style="margin:0;font-size:12px;color:#6B7280;">Developed by Anderson &nbsp;|&nbsp; Sabre PDF Converter v1.0.0 &nbsp;|&nbsp; © 2026 All Rights Reserved</p>' +
      '</div>' +
    '</div>' +
  '</body></html>';
}
