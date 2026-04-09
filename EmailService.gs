/**
 * EmailService.gs
 * 使用 GmailApp 寄送轉換結果 Email
 */

var EMAIL_SUBJECT_PREFIX = '[PDF Converter] ';
var SMALL_FILE_LIMIT_BYTES = 20 * 1024 * 1024; // 20MB 以下直接附件

/**
 * 寄送轉換完成通知 Email
 */
function sendConversionResult(recipientEmail, fileInfoArray, sourcePdfName, folderId) {
  try {
    var folderUrl = 'https://drive.google.com/drive/folders/' + folderId;
    var subject = EMAIL_SUBJECT_PREFIX + sourcePdfName + ' - Conversion Complete';
    var body = buildEmailBody(fileInfoArray, sourcePdfName, folderUrl);

    var attachments = [];
    fileInfoArray.forEach(function(info) {
      if (info.size < SMALL_FILE_LIMIT_BYTES) {
        try {
          var file = DriveApp.getFileById(info.fileId);
          attachments.push(file.getBlob());
        } catch (e) { /* 附件失敗不中斷主流程 */ }
      }
    });

    GmailApp.sendEmail(recipientEmail, subject, buildPlainText(fileInfoArray, sourcePdfName), {
      htmlBody: body,
      attachments: attachments,
      name: 'PDF Converter'
    });

    return { success: true, message: 'Email 已寄送至 ' + recipientEmail };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * 純文字備用版本（供不支援 HTML 的 email client）
 */
function buildPlainText(fileInfoArray, sourcePdfName) {
  var lines = [
    'PDF Converter - Conversion Complete',
    '=====================================',
    '',
    'Source File: ' + sourcePdfName,
    'Output Files: ' + fileInfoArray.length,
    '',
  ];
  fileInfoArray.forEach(function(f, i) {
    var sizeStr = f.size > 1024 * 1024
      ? (f.size / 1024 / 1024).toFixed(2) + ' MB'
      : (f.size / 1024).toFixed(1) + ' KB';
    lines.push((i + 1) + '. ' + f.fileName + ' (' + sizeStr + ')');
  });
  lines.push('', '-----', 'Developed by Anderson | PDF Converter v1.0.0 | (c) 2026');
  return lines.join('\n');
}

/**
 * 專業 HTML Email 模板（無 Emoji，純 inline CSS，廣泛相容）
 */
function buildEmailBody(fileInfoArray, sourcePdfName, folderUrl) {
  var ts = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');

  /* ── 檔案列表 rows ── */
  var rows = fileInfoArray.map(function(f, i) {
    var sizeStr = f.size > 1024 * 1024
      ? (f.size / 1024 / 1024).toFixed(2) + ' MB'
      : (f.size / 1024).toFixed(1) + ' KB';
    var bg = i % 2 === 0 ? '#FFFFFF' : '#FAFAFA';
    return '<tr style="background:' + bg + ';">' +
      '<td style="padding:11px 16px;font-size:13px;color:#1A1F36;border-bottom:1px solid #EEEEEE;word-break:break-all;">' +
        f.fileName +
      '</td>' +
      '<td style="padding:11px 16px;font-size:13px;color:#6B7280;border-bottom:1px solid #EEEEEE;white-space:nowrap;">' +
        sizeStr +
      '</td>' +
      '<td style="padding:11px 16px;font-size:13px;border-bottom:1px solid #EEEEEE;white-space:nowrap;">' +
        '<span style="display:inline-block;background:#ECFDF5;color:#059669;' +
        'padding:2px 10px;border-radius:20px;font-size:12px;font-weight:600;letter-spacing:0.3px;">Completed</span>' +
      '</td>' +
    '</tr>';
  }).join('');

  return '<!DOCTYPE html>' +
  '<html lang="zh-TW">' +
  '<head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>PDF Converter - Conversion Complete</title>' +
  '</head>' +
  '<body style="margin:0;padding:0;background:#F0F2F5;font-family:Arial,Helvetica,sans-serif;">' +

  '<!-- Wrapper -->' +
  '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#F0F2F5;padding:32px 16px;">' +
  '<tr><td align="center">' +

  '<!-- Card -->' +
  '<table width="600" cellpadding="0" cellspacing="0" border="0" ' +
    'style="max-width:600px;width:100%;background:#FFFFFF;border-radius:8px;' +
    'overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">' +

    '<!-- Header -->' +
    '<tr>' +
      '<td style="background:#FF6B00;padding:28px 36px;">' +
        '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr>' +
          '<td>' +
            '<div style="font-size:11px;color:rgba(255,255,255,0.7);letter-spacing:2px;' +
              'text-transform:uppercase;margin-bottom:6px;">PDF Converter</div>' +
            '<div style="font-size:22px;font-weight:700;color:#FFFFFF;line-height:1.2;">' +
              'Conversion Complete' +
            '</div>' +
          '</td>' +
          '<td align="right" valign="top">' +
            '<div style="background:rgba(255,255,255,0.2);border-radius:4px;' +
              'padding:4px 10px;font-size:12px;color:#FFFFFF;white-space:nowrap;">' +
              ts +
            '</div>' +
          '</td>' +
        '</tr></table>' +
      '</td>' +
    '</tr>' +

    '<!-- Summary Bar -->' +
    '<tr>' +
      '<td style="background:#FFF8F3;border-bottom:1px solid #FFE0C8;padding:14px 36px;">' +
        '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr>' +
          '<td style="font-size:13px;color:#6B7280;">Source File</td>' +
          '<td style="font-size:13px;font-weight:600;color:#1A1F36;text-align:right;">' +
            sourcePdfName +
          '</td>' +
        '</tr><tr>' +
          '<td style="font-size:13px;color:#6B7280;padding-top:4px;">Output Files</td>' +
          '<td style="font-size:13px;font-weight:600;color:#FF6B00;text-align:right;padding-top:4px;">' +
            fileInfoArray.length + ' file' + (fileInfoArray.length > 1 ? 's' : '') +
          '</td>' +
        '</tr></table>' +
      '</td>' +
    '</tr>' +

    '<!-- Body -->' +
    '<tr>' +
      '<td style="padding:28px 36px;">' +

        '<p style="margin:0 0 16px;font-size:14px;color:#374151;line-height:1.6;">' +
          'The following files have been converted and are attached to this email.' +
          ' You may also access them via Google Drive.' +
        '</p>' +

        '<!-- File Table -->' +
        '<table width="100%" cellpadding="0" cellspacing="0" border="0" ' +
          'style="border:1px solid #E5E7EB;border-radius:6px;overflow:hidden;margin-bottom:24px;">' +
          '<thead>' +
            '<tr style="background:#F9FAFB;">' +
              '<th style="padding:10px 16px;text-align:left;font-size:11px;font-weight:700;' +
                'color:#9CA3AF;text-transform:uppercase;letter-spacing:0.8px;' +
                'border-bottom:1px solid #E5E7EB;">File Name</th>' +
              '<th style="padding:10px 16px;text-align:left;font-size:11px;font-weight:700;' +
                'color:#9CA3AF;text-transform:uppercase;letter-spacing:0.8px;' +
                'border-bottom:1px solid #E5E7EB;">Size</th>' +
              '<th style="padding:10px 16px;text-align:left;font-size:11px;font-weight:700;' +
                'color:#9CA3AF;text-transform:uppercase;letter-spacing:0.8px;' +
                'border-bottom:1px solid #E5E7EB;">Status</th>' +
            '</tr>' +
          '</thead>' +
          '<tbody>' + rows + '</tbody>' +
        '</table>' +

        '<!-- CTA Button -->' +
        '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr>' +
          '<td align="center">' +
            '<a href="' + folderUrl + '" ' +
              'style="display:inline-block;background:#FF6B00;color:#FFFFFF;' +
              'padding:12px 32px;border-radius:6px;text-decoration:none;' +
              'font-size:14px;font-weight:600;letter-spacing:0.3px;">' +
              'Open Google Drive Folder' +
            '</a>' +
          '</td>' +
        '</tr></table>' +

      '</td>' +
    '</tr>' +

    '<!-- Footer -->' +
    '<tr>' +
      '<td style="background:#F9FAFB;border-top:1px solid #E5E7EB;padding:16px 36px;text-align:center;">' +
        '<p style="margin:0;font-size:11px;color:#9CA3AF;line-height:1.6;">' +
          'Developed by Anderson &nbsp;&nbsp;|&nbsp;&nbsp; PDF Converter v1.0.0' +
          '&nbsp;&nbsp;|&nbsp;&nbsp; &copy; 2026 All Rights Reserved' +
        '</p>' +
        '<p style="margin:6px 0 0;font-size:11px;color:#D1D5DB;">' +
          'This is an automated notification. Please do not reply to this email.' +
        '</p>' +
      '</td>' +
    '</tr>' +

  '</table>' +
  '<!-- /Card -->' +

  '</td></tr></table>' +
  '<!-- /Wrapper -->' +

  '</body></html>';
}
