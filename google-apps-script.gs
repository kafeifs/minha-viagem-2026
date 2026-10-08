const SPREADSHEET_ID = '1C_GoUCYpQ8ICn12swSboxNi8Z8UwNmuzgZRzvWhshkg';
const DAYOFF_SHEET = 'Cronograma Day-off';
const PARKS_SHEET = 'Cronograma Parques';

function doGet(e) {
  const rawCallback = (e && e.parameter && e.parameter.callback) || '';
  const callback = /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(rawCallback) ? rawCallback : '';
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const daySheet = ss.getSheetByName(DAYOFF_SHEET);
    const parkSheet = ss.getSheetByName(PARKS_SHEET);
    if (!daySheet || !parkSheet) throw new Error('Uma das abas configuradas não foi encontrada.');

    const payload = {
      ok: true,
      updatedAt: new Date().toISOString(),
      dayRows: daySheet.getDataRange().getDisplayValues(),
      parkRows: parkSheet.getDataRange().getDisplayValues()
    };

    const json = JSON.stringify(payload).replace(/</g, '\\u003c');
    const body = callback ? callback + '(' + json + ')' : json;

    return ContentService
      .createTextOutput(body)
      .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  } catch (err) {
    const payload = { ok: false, error: String(err && err.message ? err.message : err) };
    const json = JSON.stringify(payload).replace(/</g, '\\u003c');
    const body = callback ? callback + '(' + json + ')' : json;
    return ContentService
      .createTextOutput(body)
      .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  }
}
