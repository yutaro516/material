// ==========================================================
// ✅ スプレッドシート情報
// ==========================================================
const SPREADSHEET_ID = '1ASjhxPSwbz3TRSzBqyjL1BJzpkDbAJZ-918dSPvVxxg';
const SHEET_NAME = 'フォームの回答1';

const DATE_COLUMN_INDEX = 0;
const ITEM_COLUMN_INDEX = 1;
const LOCATION_COLUMN_INDEX = 2;
const PHOTO_COLUMN_INDEX = 3;


// ==========================================================
// ✅ doGet（入口）
// ==========================================================
function doGet(e) {

  let place = "";
  if (e && e.parameter && e.parameter.location) {
    place = e.parameter.location.trim();
  }

  const html = HtmlService.createTemplateFromFile('Index');
  html.place = place;

  return html.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ==========================================================
// ✅ 落とし物データ取得（Base64画像付き）
// ==========================================================
function getLostItemsByLocation(place) {

  if (!place) return [];

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const values = sheet.getDataRange().getValues();

    const result = [];

    for (let i = 1; i < values.length; i++) {

      const row = values[i];
      const rowPlace = row[LOCATION_COLUMN_INDEX];

      if (!rowPlace) continue;

      if (rowPlace.toString().trim() === place) {

        const timestamp =
          row[DATE_COLUMN_INDEX] instanceof Date
            ? row[DATE_COLUMN_INDEX].toLocaleString()
            : row[DATE_COLUMN_INDEX];

        const photoUrl = row[PHOTO_COLUMN_INDEX];
        let base64 = "";

        if (photoUrl) {
          base64 = convertDriveImageToBase64(photoUrl);
        }

        result.push({
          timestamp: timestamp,
          item: row[ITEM_COLUMN_INDEX],
          base64img: base64
        });
      }
    }

    return result;

  } catch (err) {
    Logger.log(err);
    return [];
  }
}


// ==========================================================
// ✅ Drive URL（open?id= にも対応）→ fileId 抽出
// ==========================================================
function extractFileId(url) {
  if (!url) return "";

  const match =
    url.match(/\/d\/([a-zA-Z0-9_-]+)/) ||
    url.match(/id=([a-zA-Z0-9_-]+)/);

  return match ? match[1] : "";
}


// ==========================================================
// ✅ Google Drive の画像URL → Base64へ変換
// ==========================================================
function convertDriveImageToBase64(url) {

  try {
    const fileId = extractFileId(url);
    if (!fileId) return "";

    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();

    const contentType = blob.getContentType();
    const base64Data = Utilities.base64Encode(blob.getBytes());

    return `data:${contentType};base64,${base64Data}`;

  } catch (err) {
    Logger.log("Base64変換失敗: " + err);
    return "";
  }
}


// ==========================================================
// ✅ HTML include
// ==========================================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
