const SPREADSHEET_ID = '1ASjhxPSwbz3TRSzBqyjL1BJzpkDbAJZ-918dSPvVxxg';
const SHEET_NAME = 'フォームの回答1';
const DATE_COLUMN_INDEX = 0;
const ITEM_COLUMN_INDEX = 1;
const LOCATION_COLUMN_INDEX = 2;
const PHOTO_COLUMN_INDEX = 3;

function doGet(e) {
  let place = "";
  if (e && e.parameter && e.parameter.location) place = e.parameter.location.trim();
  const html = HtmlService.createTemplateFromFile('Index');
  html.place = place;
  return html.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

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
        const timestamp = row[DATE_COLUMN_INDEX] instanceof Date
          ? row[DATE_COLUMN_INDEX].toLocaleString()
          : row[DATE_COLUMN_INDEX];

        // 写真をBase64化
        const photoUrl = convertToBase64(row[PHOTO_COLUMN_INDEX]);

        result.push({
          timestamp: timestamp,
          item: row[ITEM_COLUMN_INDEX],
          photo: photoUrl
        });
      }
    }
    return result;
  } catch (err) {
    Logger.log(err);
    return [];
  }
}

// Drive URL → Base64 変換
function convertToBase64(url) {
  if (!url) return "";
  let fileId = extractFileId(url);
  if (!fileId) return "";
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
  } catch(e) {
    Logger.log(e);
    return "";
  }
}

function extractFileId(url) {
  if (!url) return "";
  const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
  return match ? match[1] : "";
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
