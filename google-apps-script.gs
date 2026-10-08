const SPREADSHEET_ID = '1C_GoUCYpQ8ICn12swSboxNi8Z8UwNmuzgZRzvWhshkg';
const DAYOFF_SHEET = 'Cronograma Day-off';
const PARKS_SHEET = 'Cronograma Parques';

function getCellComments_() {
  const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(SPREADSHEET_ID)
    + '?commentsViewMode=COMMENTS_VIEW_MODE_INCLUDED'
    + '&fields=comments(commentId,anchorId,headPost(content),replies(content),status),sheets(properties(sheetId,title),commentAnchors(anchorId,range))';
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error('Não foi possível ler os comentários do Google Sheets (' + status + ').');
  }
  const data = JSON.parse(response.getContentText());
  const sheetNames = {};
  const anchors = {};
  (data.sheets || []).forEach(sh => {
    const sid = sh.properties && sh.properties.sheetId;
    const title = sh.properties && sh.properties.title;
    sheetNames[sid] = title;
    (sh.commentAnchors || []).forEach(a => {
      const r = a.range || {};
      if (r.startRowIndex == null || r.startColumnIndex == null) return;
      anchors[a.anchorId] = {
        sheet: title,
        row: r.startRowIndex,
        col: r.startColumnIndex
      };
    });
  });

  const comments = [];
  (data.comments || []).forEach(thread => {
    const a = anchors[thread.anchorId];
    if (!a) return;
    const head = thread.headPost && thread.headPost.content ? String(thread.headPost.content).trim() : '';
    const replies = (thread.replies || []).map(r => r && r.content ? String(r.content).trim() : '').filter(Boolean);
    const content = [head].concat(replies).filter(Boolean).join('\n');
    if (!content) return;
    comments.push({
      sheet: a.sheet,
      row: a.row,
      col: a.col,
      content,
      status: thread.status || 'OPEN'
    });
  });
  return comments;
}

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
      comments: getCellComments_()
    };

    const body = callback
      ? callback + '(' + JSON.stringify(payload).replace(/</g, '\\u003c') + ')'
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
