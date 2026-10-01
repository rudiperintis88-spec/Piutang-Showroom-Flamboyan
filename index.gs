/**
 * SINKRON CATATAN & TUGAS PENAGIHAN — Google Apps Script
 * Cara pasang: lihat langkah di bawah / di jawaban Claude.
 *  1) Ganti SYNC_KEY dengan kata rahasia Anda (harus SAMA dengan SYNC_KEY di file dashboard).
 *  2) Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 *  3) Salin URL "/exec" ke SYNC_URL di file dashboard.
 */
var SYNC_KEY = 'GANTI_DENGAN_KUNCI_RAHASIA';
var SHEET_NAME = 'Catatan';

/* ===== Aturan penggabungan catatan/tugas (SAMA persis di dashboard dan di Apps Script) ===== */
function ppMergeRec(a, b) {
  a = a || {}; b = b || {};
  const out = { nama: a.nama || b.nama || '', jumlah: a.jumlah || b.jumlah || 0, leasing: a.leasing || b.leasing || '', jto: a.jto || b.jto || '' };
  out.delAt = [a.delAt, b.delAt].filter(Boolean).sort().pop() || null;             // hapus-riwayat: yang terbaru menang
  const pick = (b.lu || '') > (a.lu || '') ? b : a;                                  // status lunas: perubahan terbaru menang
  out.lunasAt = pick.lunasAt || null; out.lu = pick.lu || null;
  const notes = {};                                                                  // catatan: gabungan semua perangkat (hanya bertambah)
  [a, b].forEach(r => (r.notes || []).forEach(n => { const id = n.id || ('L' + (n.t || '') + '|' + n.x); notes[id] = Object.assign({}, n, { id: id }); }));
  out.notes = Object.keys(notes).map(k => notes[k]).filter(n => !out.delAt || (n.t && n.t > out.delAt))
    .sort((p, q) => String(p.t || '').localeCompare(String(q.t || '')));
  const tasks = {};                                                                  // tugas: versi dengan waktu ubah terbaru menang
  [a, b].forEach(r => (r.tasks || []).forEach(t => { const p = tasks[t.id]; if (!p || (t.u || t.t || '') > (p.u || p.t || '')) tasks[t.id] = t; }));
  out.tasks = Object.keys(tasks).map(k => tasks[k]).filter(t => !out.delAt || (t.t && t.t > out.delAt));
  return out;
}

/* ===== Server: simpan satu baris per faktur di tab "Catatan" ===== */
function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(['divisi', 'faktur', 'data (JSON)', 'diubah']); }
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return out_({ ok: true, info: 'Sinkron catatan penagihan aktif' }); }
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.key !== SYNC_KEY) return out_({ ok: false, error: 'Kunci salah' });
    var div = String(req.divisi || '');
    if (!div) return out_({ ok: false, error: 'Divisi kosong' });
    var sh = sheet_(), last = sh.getLastRow();
    var data = last > 1 ? sh.getRange(2, 1, last - 1, 3).getValues() : [];
    var rowOf = {}, db = {};
    data.forEach(function (r, i) { if (r[0] === div) { rowOf[r[1]] = i + 2; try { db[r[1]] = JSON.parse(r[2]); } catch (x) {} } });
    var inc = req.db || {};
    Object.keys(inc).forEach(function (f) {
      var m = ppMergeRec(db[f], inc[f]), json = JSON.stringify(m);
      db[f] = m;
      if (json.length > 45000) return;                       // batas isi satu sel Google Sheets
      if (rowOf[f]) sh.getRange(rowOf[f], 3, 1, 2).setValues([[json, new Date()]]);
      else { sh.appendRow([div, f, json, new Date()]); rowOf[f] = sh.getLastRow(); }
    });
    return out_({ ok: true, db: db });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
