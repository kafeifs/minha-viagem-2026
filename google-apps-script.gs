const SPREADSHEET_ID = '1C_GoUCYpQ8ICn12swSboxNi8Z8UwNmuzgZRzvWhshkg';
const DAYOFF_SHEET = 'Cronograma Day-off';
const PARKS_SHEET = 'Cronograma Parques';

function doGet(e) {
  const callback = (e && e.parameter && e.parameter.callback) || '';
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const daySheet = ss.getSheetByName(DAYOFF_SHEET);
    const parkSheet = ss.getSheetByName(PARKS_SHEET);
    if (!daySheet || !parkSheet) throw new Error('Uma das abas configuradas não foi encontrada.');

    const payload = {
      ok: true,
      updatedAt: new Date().toISOString(),
      dayRows: daySheet.getDataRange().getDisplayValues(),
      parkRows: parkSheet.getDataRange().getDisplayValues(),
      comments: [],
      commentsWarning: null
    };

    const body = callback
      ? callback + '(' + JSON.stringify(payload).replace(/</g, '\u003c') + ')'
      : JSON.stringify(payload);

    return ContentService
      .createTextOutput(body)
      .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  } catch (err) {
    const payload = { ok:false, error:String(err && err.message ? err.message : err) };
    const body = callback ? callback + '(' + JSON.stringify(payload) + ')' : JSON.stringify(payload);
    return ContentService
      .createTextOutput(body)
      .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  }
}
