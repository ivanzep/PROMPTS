// Veras Prompt Manager — Google Sheets sync
// Deploy as: Web App | Execute as: Me | Access: Anyone

const SHEET_NAME = 'Prompts';
const HEADERS = ['id','title','project','category','preset','lighting','rating','favorite','tags',
                 'geoOverride','matOverride','promptStrength','prompt','negative','created'];

function doGet(e) {
  try {
    const sheet = getSheet();
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return jsonResp({ prompts: [] });
    const headers = rows[0];
    const prompts = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      obj.rating = Number(obj.rating) || 0;
      obj.favorite = obj.favorite === 'yes' || obj.favorite === true;
      obj.tags = obj.tags ? String(obj.tags).split('|').filter(Boolean) : [];
      obj.geoOverride = Number(obj.geoOverride) || 20;
      obj.matOverride = Number(obj.matOverride) || 80;
      obj.promptStrength = Number(obj.promptStrength) || 60;
      obj.created = obj.created ? new Date(obj.created).getTime() : Date.now();
      return obj;
    });
    return jsonResp({ prompts });
  } catch(err) { return jsonResp({ error: err.message }); }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const prompts = data.prompts || [];
    const sheet = getSheet();
    sheet.clearContents();
    sheet.appendRow(HEADERS);
    prompts.forEach(p => {
      sheet.appendRow(HEADERS.map(h => {
        if (h === 'tags') return (p.tags||[]).join('|');
        if (h === 'favorite') return p.favorite ? 'yes' : 'no';
        if (h === 'created') return new Date(p.created || Date.now()).toISOString().slice(0,10);
        return p[h] ?? '';
      }));
    });
    return jsonResp({ ok: true, count: prompts.length });
  } catch(err) { return jsonResp({ error: err.message }); }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
