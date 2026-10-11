// Google Apps Script — database sheet untuk Survei Event Jalan
// 1) Buat Google Sheet baru > Extensions > Apps Script > tempel kode ini
// 2) Ganti TOKEN, lalu Deploy > New deployment > Web app
//    (Execute as: Me, Who has access: Anyone) > salin URL /exec ke menu Data di aplikasi
const TOKEN = 'ganti-token-anda';
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents), ss = SpreadsheetApp.getActive();
    if (p.token !== TOKEN) return out({ ok: false, err: 'Token salah' });
    for (const name in p.tables) {
      const t = p.tables[name];
      const sh = ss.getSheetByName(name) || ss.insertSheet(name);
      if (sh.getLastRow() === 0) { sh.appendRow(t.header); sh.setFrozenRows(1); }
      const last = sh.getLastRow();
      const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
      const add = [];
      t.rows.forEach(r => {
        const i = ids.indexOf(String(r[0]));
        if (i >= 0) sh.getRange(i + 2, 1, 1, r.length).setValues([r]); else add.push(r);
      });
      if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, add[0].length).setValues(add);
    }
    return out({ ok: true });
  } catch (err) { return out({ ok: false, err: String(err) }); }
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
