/**
 * 📒 รายรับ/รายจ่าย💸 — Backend (Google Apps Script)
 *
 * วิธีติดตั้ง:
 * 1. สร้าง Google Sheet ใหม่
 * 2. Extensions > Apps Script > ลบโค้ดเดิม วางไฟล์นี้แทน
 * 3. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. คัดลอก Web app URL ไปใส่ในหน้า ⚙️ ของแอพ
 */

const SHEET_NAME = "Transactions";
const HEADERS = ["id", "type", "amount", "title", "categoryId", "walletId", "createdAt"];
const META_SHEET_NAME = "Meta";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getMetaSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(META_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(META_SHEET_NAME);
    sheet.appendRow(["key", "value"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getMeta_(key) {
  const sheet = getMetaSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === key) return values[i][1];
  }
  return null;
}

function setMeta_(key, value) {
  const sheet = getMetaSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function rowsToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).map((r, idx) => {
    const obj = { _row: idx + 2 };
    headers.forEach((h, i) => (obj[h] = r[i]));
    return obj;
  });
}

function doGet(e) {
  const sheet = getSheet_();
  let rows = rowsToObjects_(sheet).filter((r) => r.id !== "" && r.id !== undefined);
  const month = e.parameter.month; // "YYYY-MM" optional filter
  if (month) {
    rows = rows.filter((r) => String(r.createdAt).slice(0, 7) === month);
  }
  rows.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  rows.forEach((r) => delete r._row);
  const investmentValue = getMeta_("investment_value");
  return jsonOut_({ ok: true, transactions: rows, investmentValue: investmentValue === null ? null : Number(investmentValue) });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const sheet = getSheet_();

  if (body.action === "add") {
    const id = `${Date.now()}`;
    sheet.appendRow([
      id,
      body.type,
      Number(body.amount) || 0,
      body.title || "",
      body.categoryId || "other",
      body.walletId || "cash",
      body.createdAt || new Date().toISOString(),
    ]);
    return jsonOut_({ ok: true, id });
  }

  if (body.action === "delete") {
    const rows = rowsToObjects_(sheet);
    const target = rows.find((r) => String(r.id) === String(body.id));
    if (target) sheet.deleteRow(target._row);
    return jsonOut_({ ok: true });
  }

  if (body.action === "update") {
    const rows = rowsToObjects_(sheet);
    const target = rows.find((r) => String(r.id) === String(body.id));
    if (target) {
      const rowNum = target._row;
      const newRow = HEADERS.map((h) => (body[h] !== undefined ? body[h] : target[h]));
      sheet.getRange(rowNum, 1, 1, HEADERS.length).setValues([newRow]);
    }
    return jsonOut_({ ok: true });
  }

  if (body.action === "transfer") {
    const ts = body.createdAt || new Date().toISOString();
    const fromLabel = body.fromLabel || body.fromWallet;
    const toLabel = body.toLabel || body.toWallet;
    const amount = Number(body.amount) || 0;
    const idOut = `${Date.now()}-out`;
    const idIn = `${Date.now()}-in`;
    sheet.appendRow([idOut, "expense", amount, `โอนไป ${toLabel}`, "transfer", body.fromWallet, ts]);
    sheet.appendRow([idIn, "income", amount, `รับโอนจาก ${fromLabel}`, "transfer", body.toWallet, ts]);
    return jsonOut_({ ok: true });
  }

  if (body.action === "setMeta") {
    setMeta_(body.key, body.value);
    return jsonOut_({ ok: true });
  }

  return jsonOut_({ ok: false, error: "unknown action" });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
