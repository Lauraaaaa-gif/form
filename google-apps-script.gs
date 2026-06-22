const SPREADSHEET_ID = "1hMl7sQocX9yHjbcGpoSkoxzqAP4knoqRYm-TLhyVtLo";
const SHEET_NAME = "報名資料";
const HEADERS = [
  "收據編號",
  "時間",
  "FB名稱",
  "遊戲名稱",
  "UID",
  "商品名稱",
  "數量",
  "狀態",
  "排單日期",
  "確認時間",
  "管理備註",
  "來源"
];

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  if (params.uid) {
    return output_(findOrderByUid_(params.uid), params.callback);
  }

  return ContentService
    .createTextOutput("報名表後端已啟用")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const sheet = getSheet_();
  const data = e.parameter || {};

  sheet.appendRow([
    data.orderNo || makeOrderNo_(),
    new Date(),
    data.fbName || "",
    data.gameName || "",
    data.uid || "",
    data.productName || "",
    Number(data.quantity || 0),
    data.status || "待確認",
    data.scheduledDate || "",
    data.confirmedAt || "",
    data.adminNote || "",
    data.source || "第五排單"
  ]);

  return output_({ ok: true });
}

function findOrderByUid_(uid) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || [];
  const uidIndex = headers.indexOf("UID");

  if (uidIndex < 0) {
    return { ok: false, message: "查無 UID 欄位。" };
  }

  const target = String(uid).trim();
  for (let rowIndex = values.length - 1; rowIndex >= 1; rowIndex -= 1) {
    const row = values[rowIndex];
    if (String(row[uidIndex]).trim() === target) {
      return {
        ok: true,
        orderNo: getCell_(headers, row, "收據編號"),
        createdAt: formatDate_(getCell_(headers, row, "時間")),
        fbName: getCell_(headers, row, "FB名稱"),
        gameName: getCell_(headers, row, "遊戲名稱"),
        uid: getCell_(headers, row, "UID"),
        productName: getCell_(headers, row, "商品名稱"),
        quantity: getCell_(headers, row, "數量"),
        status: getCell_(headers, row, "狀態") || "待確認",
        scheduledDate: formatDate_(getCell_(headers, row, "排單日期")),
        confirmedAt: formatDate_(getCell_(headers, row, "確認時間")),
        adminNote: getCell_(headers, row, "管理備註")
      };
    }
  }

  return { ok: false, message: "查無此 UID 的訂單。" };
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  ensureHeaders_(sheet);
  return sheet;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    formatSheet_(sheet);
    return;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  if (headers[0] === "時間") {
    sheet.insertColumnBefore(1);
    sheet.getRange(1, 1).setValue("收據編號");
  }

  headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  HEADERS.forEach(header => {
    if (!headers.includes(header)) {
      const sourceColumn = headers.indexOf("來源") + 1;
      if (sourceColumn > 0) {
        sheet.insertColumnBefore(sourceColumn);
        sheet.getRange(1, sourceColumn).setValue(header);
      } else {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      }
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
  });

  const orderedHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (orderedHeaders.join("|") !== HEADERS.join("|")) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }

  formatSheet_(sheet);
}

function formatSheet_(sheet) {
  sheet.getRange("B:B").setNumberFormat("yyyy-mm-dd hh:mm:ss");
  sheet.getRange("G:G").setNumberFormat("0");
  sheet.getRange("I:J").setNumberFormat("yyyy-mm-dd");
  sheet.setFrozenRows(1);

  const statusRange = sheet.getRange("H2:H");
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["待確認", "已確認", "資料有誤", "已取消", "已完成"], true)
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(rule);
}

function getCell_(headers, row, header) {
  const index = headers.indexOf(header);
  return index >= 0 ? row[index] : "";
}

function formatDate_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, "Asia/Taipei", "yyyy-MM-dd HH:mm");
  }
  return String(value);
}

function output_(data, callback) {
  const json = JSON.stringify(data);

  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${json});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function makeOrderNo_() {
  const date = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd");
  const random = Math.random().toString(36).slice(2, 7).toUpperCase();
  return `FORM-${date}-${random}`;
}
