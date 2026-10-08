/**
 * Salud 2 Go — Apps Script V12
 * Backend único para Leads, Diagnóstico, Cotizaciones y Funnel.
 * Usa los nombres de columnas de la fila 1, no posiciones fijas.
 */
const SHEET_NAMES = {
  lead: "Leads",
  diagnostic: "Diagnóstico",
  quote: "Cotizaciones",
  funnel: "Funnel"
};

function doGet() {
  return jsonResponse_({success:true, service:"Salud 2 Go", version:"V12", timestamp:new Date().toISOString()});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({success:false, error:"empty_body"});
    }

    const payload = JSON.parse(e.postData.contents);
    const type = String(payload.type || "").trim().toLowerCase();
    const sheetName = SHEET_NAMES[type];

    if (!sheetName) {
      return jsonResponse_({success:false, error:"unsupported_type", received_type:type});
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return jsonResponse_({success:false, error:"sheet_not_found", sheet:sheetName});
    }

    appendByHeader_(sheet, payload);
    return jsonResponse_({success:true, type:type, sheet:sheetName});

  } catch (err) {
    console.error(err);
    return jsonResponse_({success:false, error:String(err && err.message ? err.message : err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function appendByHeader_(sheet, payload) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) throw new Error("La hoja " + sheet.getName() + " no tiene encabezados.");

  const headers = sheet.getRange(1,1,1,lastColumn).getValues()[0].map(h => String(h || "").trim());
  if (!headers.some(Boolean)) throw new Error("La fila 1 de " + sheet.getName() + " está vacía.");

  const data = Object.assign({}, payload);
  if (!data.timestamp) data.timestamp = new Date();

  const row = headers.map(header => {
    if (!header || !Object.prototype.hasOwnProperty.call(data, header)) return "";
    return normalizeCellValue_(data[header]);
  });

  sheet.appendRow(row);
}

function normalizeCellValue_(value) {
  if (value === null || typeof value === "undefined") return "";
  if (Array.isArray(value)) return value.map(v => typeof v === "object" ? JSON.stringify(v) : String(v)).join(", ");
  if (typeof value === "object" && !(value instanceof Date)) return JSON.stringify(value);
  return value;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
