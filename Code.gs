const SHEET_NAME = "บันทึกรอบบิล";
const DAILY_DEDUCTION = 25;
const RATE_PER_BILL = 15;
var currentCallback_ = "";

function doGet(e) {
  try {
    currentCallback_ = String((e && e.parameter && e.parameter.callback) || "");
    if (currentCallback_ && !/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(currentCallback_)) {
      currentCallback_ = "";
    }
    const action = (e.parameter.action || "list").toLowerCase();
    const sheet = getSheet_();
    if (action === "list") {
      const values = sheet.getDataRange().getValues();
      const records = values.slice(1).filter(row => row[0]).map(row => ({
        date: formatDate_(row[0]),
        totalBills: Number(row[1]) || 0,
        deduction: Number(row[2]) || DAILY_DEDUCTION,
        netBills: Number(row[3]) || 0,
        grossMoney: Number(row[4]) || 0,
        netMoney: Number(row[5]) || 0
      }));
      return json_({success:true, records:records});
    }
    if (action === "upsert") {
      const date = String(e.parameter.date || "");
      const total = Number(e.parameter.totalBills);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !Number.isInteger(total) || total < 0) {
        return json_({success:false, error:"วันที่หรือจำนวนบิลไม่ถูกต้อง"});
      }
      const net = Math.max(0, total - DAILY_DEDUCTION);
      const record = [date, total, DAILY_DEDUCTION, net, total * RATE_PER_BILL, net * RATE_PER_BILL];
      const values = sheet.getDataRange().getValues();
      let foundRow = -1;
      for (let i = 1; i < values.length; i++) {
        if (formatDate_(values[i][0]) === date) { foundRow = i + 1; break; }
      }
      if (foundRow > 0) sheet.getRange(foundRow, 1, 1, record.length).setValues([record]);
      else sheet.appendRow(record);
      return json_({success:true});
    }
    if (action === "delete") {
      const date = String(e.parameter.date || "");
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json_({success:false, error:"วันที่ไม่ถูกต้อง"});
      const values = sheet.getDataRange().getValues();
      for (let i = values.length - 1; i >= 1; i--) {
        if (formatDate_(values[i][0]) === date) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return json_({success:true});
    }
    return json_({success:false, error:"ไม่รู้จักคำสั่งนี้"});
  } catch (err) {
    return json_({success:false, error:String(err)});
  }
}

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) throw new Error("ยังไม่ได้ตั้งค่า SPREADSHEET_ID ใน Script Properties");
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["วันที่", "บิลทั้งหมด", "หักบิล", "บิลสุทธิ", "เงินก่อนหัก (บาท)", "เงินหลังหัก (บาท)"]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
  }
  return sheet;
}

function formatDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(value).slice(0, 10);
}

function json_(obj) {
  // รองรับ JSONP สำหรับเว็บที่เรียก Apps Script จากอีกโดเมน
  const callback = (typeof currentCallback_ !== "undefined") ? currentCallback_ : "";
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + JSON.stringify(obj) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// เรียกฟังก์ชันนี้จาก Apps Script editor หนึ่งครั้ง เพื่อใส่ ID ของ Google Sheets
function setupSpreadsheet() {
  const id = "ใส่_SPREADSHEET_ID_ตรงนี้";
  if (id === "ใส่_SPREADSHEET_ID_ตรงนี้") {
    throw new Error("แก้ค่า id ใน setupSpreadsheet() ให้เป็น ID ของ Google Sheets ก่อน");
  }
  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", id);
  getSheet_();
}
