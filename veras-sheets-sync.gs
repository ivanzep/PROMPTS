// Veras Prompt Manager — Google Sheets sync
// Deploy as: Web App | Execute as: Me | Access: Anyone
//
// Sheet columns (in order):
//   id · title · project · category · preset · lighting · rating · favorite
//   tags (pipe-separated) · geoOverride · matOverride · promptStrength
//   prompt · negative · created (YYYY-MM-DD)

const SHEET_NAME = 'Prompts';
const HEADERS = ['id','title','project','category','preset','lighting','rating','favorite','tags',
                 'geoOverride','matOverride','promptStrength','prompt','negative','created'];

// ── GET — return all prompts as JSON ────────────────────────────────────────
function doGet(e) {
  try {
    const sheet = getSheet();
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return jsonResp({ prompts: [] });
    const headers = rows[0].map(String);
    const prompts = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i] != null ? row[i] : '');
      obj.rating         = Number(obj.rating)         || 0;
      obj.favorite       = obj.favorite === 'yes'     || obj.favorite === true;
      obj.tags           = obj.tags ? String(obj.tags).split('|').filter(Boolean) : [];
      obj.geoOverride    = Number(obj.geoOverride)    || 20;
      obj.matOverride    = Number(obj.matOverride)    || 80;
      obj.promptStrength = Number(obj.promptStrength) || 60;
      obj.created        = obj.created ? new Date(obj.created).getTime() : Date.now();
      obj.category       = String(obj.category || 'other');
      return obj;
    }).filter(p => p.id && p.prompt);
    return jsonResp({ prompts });
  } catch (err) {
    return jsonResp({ error: err.message });
  }
}

// ── POST — replace all rows with posted prompts ─────────────────────────────
function doPost(e) {
  try {
    const data    = JSON.parse(e.postData.contents);
    const prompts = data.prompts || [];
    const sheet   = getSheet();
    sheet.clearContents();
    sheet.appendRow(HEADERS);
    prompts.forEach(p => sheet.appendRow(buildRow(p)));
    return jsonResp({ ok: true, count: prompts.length });
  } catch (err) {
    return jsonResp({ error: err.message });
  }
}

// ── Helpers ──────────────────────────────────────────────────────────────────
function buildRow(p) {
  return HEADERS.map(h => {
    switch (h) {
      case 'tags':          return (p.tags || []).join('|');
      case 'favorite':      return p.favorite ? 'yes' : 'no';
      case 'created':       return new Date(p.created || Date.now()).toISOString().slice(0, 10);
      case 'rating':        return Number(p.rating)         || 0;
      case 'geoOverride':   return Number(p.geoOverride)    ?? 20;
      case 'matOverride':   return Number(p.matOverride)    ?? 80;
      case 'promptStrength':return Number(p.promptStrength) ?? 60;
      default:              return p[h] != null ? p[h] : '';
    }
  });
}

function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
