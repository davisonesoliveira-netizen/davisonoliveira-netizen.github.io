/**
 * Registro de visitas do portfólio -> Google Sheets
 *
 * Como usar:
 * 1. Crie uma planilha no Google Sheets (ex.: "Visitas do portfólio").
 * 2. Menu Extensões > Apps Script. Apague o conteúdo e cole este arquivo.
 * 3. Implantar > Nova implantação > Tipo: App da Web.
 *    - Executar como: Eu
 *    - Quem tem acesso: Qualquer pessoa
 * 4. Copie a URL terminada em /exec e cole em ENDPOINT no index.html.
 */

const SHEET_NAME = "Visitas";

const HEADERS = [
  "received_at", "event", "visitor_id", "session_id", "returning",
  "page", "referrer", "utm_source", "utm_medium", "utm_campaign",
  "language", "timezone", "screen", "viewport", "device", "browser", "os",
  "dark_mode", "target", "duration_s", "max_scroll_pct"
];

const MAX_LEN = 200;
const ALLOWED_EVENTS = ["visit", "click", "leave"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const data = JSON.parse(e.postData.contents);
    if (ALLOWED_EVENTS.indexOf(data.event) === -1) {
      return text_("ignored");
    }

    const sheet = getSheet_();
    const row = HEADERS.map(function (h) {
      if (h === "received_at") return new Date();
      return clean_(data[h]);
    });
    sheet.appendRow(row);
    return text_("ok");
  } catch (err) {
    return text_("error");
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return text_("ok");
}

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

// Limita o tamanho e impede injeção de fórmulas (=, +, -, @) na planilha.
function clean_(value) {
  if (value === undefined || value === null) return "";
  let s = String(value).slice(0, MAX_LEN);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function text_(msg) {
  return ContentService.createTextOutput(msg);
}

/**
 * Opcional: apaga registros com mais de 12 meses.
 * Para automatizar: Acionadores (ícone de relógio) > Adicionar acionador >
 * função purgeOldRows > Baseado em tempo > Mensal.
 */
function purgeOldRows() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return;
  const limit = new Date();
  limit.setMonth(limit.getMonth() - 12);
  const dates = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = dates.length - 1; i >= 0; i--) {
    if (dates[i][0] instanceof Date && dates[i][0] < limit) {
      sheet.deleteRow(i + 2);
    }
  }
}
