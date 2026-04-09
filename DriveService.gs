/**
 * DriveService.gs
 * Google Drive API v3：建立輸出資料夾、存檔、取得分享連結
 */

var ROOT_FOLDER_NAME = 'Sabre PDF Converter';

/**
 * 取得或建立根資料夾
 * @returns {Folder}
 */
function getOrCreateRootFolder() {
  var folders = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(ROOT_FOLDER_NAME);
}

/**
 * 依來源 PDF 名稱建立本次轉換輸出子資料夾
 * @param {string} baseName  來源 PDF 檔名（不含副檔名）
 * @returns {string}  子資料夾 ID
 */
function createOutputFolder(baseName) {
  var root = getOrCreateRootFolder();
  var ts = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss');
  var folderName = baseName + '_' + ts;
  var folder = root.createFolder(folderName);
  // 設定資料夾為「連結有效即可存取」
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return folder.getId();
}

/**
 * 將 Base64 資料存入指定資料夾
 * @param {string} base64Data   Base64 編碼的檔案內容
 * @param {string} fileName     目標檔名（含副檔名）
 * @param {string} mimeType     MIME type
 * @param {string} folderId     目標資料夾 ID
 * @returns {{fileId: string, webViewLink: string, fileName: string, size: number}}
 */
function saveFileToDrive(base64Data, fileName, mimeType, folderId) {
  var bytes = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(bytes, mimeType, fileName);

  var folder = DriveApp.getFolderById(folderId);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    fileId: file.getId(),
    webViewLink: file.getUrl(),
    fileName: fileName,
    size: file.getSize()
  };
}

/**
 * 批次存入多個檔案（供前端逐批呼叫）
 * @param {Array<{base64: string, name: string, mime: string}>} files
 * @param {string} folderId
 * @returns {Array<{fileId, webViewLink, fileName, size}>}
 */
function saveFilesToDrive(files, folderId) {
  return files.map(function(f) {
    return saveFileToDrive(f.base64, f.name, f.mime, folderId);
  });
}
