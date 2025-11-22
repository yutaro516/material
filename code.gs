// ==========================================================
// ====== 設定 ======
const SPREADSHEET_ID = '1ASjhxPSwbz3TRSzBqyjL1BJzpkDbAJZ-918dSPvVxxg';
const MAIN_SHEET_NAME = 'フォームの回答1';       // 落とし物登録フォーム
const COLLECTED_SHEET_NAME = 'フォームの回答 2'; // 回収済みフォーム
const DATE_COLUMN_INDEX = 0;
const ITEM_COLUMN_INDEX = 1;
const LOCATION_COLUMN_INDEX = 2;
const PHOTO_COLUMN_INDEX = 3;
const FEATURE_COLUMN_INDEX = 4; // 特徴列がある場合

// ==========================================================
// ====== ホーム画面表示 ======
function doGet(e) {
  let place = "";
  if (e && e.parameter && e.parameter.location) place = e.parameter.location.trim();
  const html = HtmlService.createTemplateFromFile('Index');
  html.place = place;
  return html.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================================
// ====== 落とし物取得 ======
function getLostItemsByLocation(place) {
  if (!place) return [];
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);

    removeCollectedItems(); // 回収済みを削除してから取得

    const values = mainSheet.getDataRange().getValues();
    const result = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowPlace = row[LOCATION_COLUMN_INDEX];
      if (!rowPlace) continue;
      if (rowPlace.toString().trim() === place) {
        const timestamp = row[DATE_COLUMN_INDEX] instanceof Date
          ? row[DATE_COLUMN_INDEX].toLocaleString()
          : row[DATE_COLUMN_INDEX];

        const photoUrl = convertToBase64(row[PHOTO_COLUMN_INDEX]);

        result.push({
          timestamp: timestamp,
          item: row[ITEM_COLUMN_INDEX],
          feature: row[FEATURE_COLUMN_INDEX],
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

// ==========================================================
// ====== 回収済みの落とし物を削除 ======
function removeCollectedItems() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  const collectedSheet = ss.getSheetByName(COLLECTED_SHEET_NAME);

  const mainData = mainSheet.getDataRange().getValues();
  const collectedData = collectedSheet.getDataRange().getValues();

  for (let i = mainData.length - 1; i >= 1; i--) {
    const mainItem = mainData[i][ITEM_COLUMN_INDEX];
    const mainPlace = mainData[i][LOCATION_COLUMN_INDEX];
    const mainFeature = FEATURE_COLUMN_INDEX !== undefined ? mainData[i][FEATURE_COLUMN_INDEX] : null;

    for (let j = 1; j < collectedData.length; j++) {
      const collectedItem = collectedData[j][ITEM_COLUMN_INDEX];
      const collectedPlace = collectedData[j][LOCATION_COLUMN_INDEX];
      const collectedFeature = FEATURE_COLUMN_INDEX !== undefined ? collectedData[j][FEATURE_COLUMN_INDEX] : null;

      // 落とし物・場所が一致、特徴があれば特徴も一致または無視
      if (mainItem && mainPlace &&
          mainItem === collectedItem &&
          mainPlace === collectedPlace &&
          (!mainFeature || !collectedFeature || mainFeature === collectedFeature)) {
        mainSheet.deleteRow(i + 1);
        break;
      }
    }
  }
}

// ==========================================================
// ====== Drive URL → Base64変換 ======
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

// ==========================================================
// ====== HTML読み込み ======
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================================
// ====== 回収済みフォーム送信時トリガー ======
function onFormSubmit(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);

    const collectedItem = e.namedValues["落とし物"];
    const collectedPlace = e.namedValues["場所"];
    const collectedFeature = e.namedValues["特徴"];

    if (!collectedItem || !collectedPlace) return;

    const values = mainSheet.getDataRange().getValues();

    for (let i = values.length - 1; i >= 1; i--) {
      const rowItem = values[i][ITEM_COLUMN_INDEX];
      const rowPlace = values[i][LOCATION_COLUMN_INDEX];
      const rowFeature = FEATURE_COLUMN_INDEX !== undefined ? values[i][FEATURE_COLUMN_INDEX] : null;

      if (rowItem === collectedItem[0] &&
          rowPlace === collectedPlace[0] &&
          (!rowFeature || rowFeature === collectedFeature[0])) {
        mainSheet.deleteRow(i + 1);
        Logger.log("削除しました: " + rowItem);
        break;
      }
    }
  } catch(err) {
    Logger.log("回収処理エラー: " + err);
  }
}
