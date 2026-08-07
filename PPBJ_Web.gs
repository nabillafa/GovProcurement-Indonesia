/**
 * Web app PPBJ — antarmuka satu pintu untuk spreadsheet, dokumen, dan PDF.
 * File ini dipasang pada proyek Apps Script yang sama dengan PPBJ_Automation.gs.
 */

const WEB_LETTER_TYPES = [
  'Penetapan HPS', 'Permintaan Pengadaan', 'Surat Pesanan', 'Pemeriksaan Pekerjaan',
  'BAST', 'Kuitansi', 'Invoice', 'Pemeriksaan Administrasi', 'Pembayaran'
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sistem Administrasi Pengadaan — Example Public Organization')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function bukaAplikasiWeb() {
  const url = ScriptApp.getService().getUrl();
  const body = url
    ? '<p style="font:14px Arial">Aplikasi web siap dibuka.</p><p><a href="' + url + '" target="_blank" style="background:#075ca8;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none;font:600 14px Arial">Buka aplikasi PPBJ</a></p>'
    : '<p style="font:14px Arial">Web app belum di-deploy. Pilih <b>Deploy &gt; New deployment &gt; Web app</b> di Apps Script.</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(body).setWidth(420).setHeight(170), 'Aplikasi Web PPBJ');
}

function webBootstrap(year) {
  ensureWebSchema_();
  const user = webRequireUser_();
  const activeYear = Number(year || new Date().getFullYear());
  return webSafe_({
    year: activeYear,
    years: webYears_(),
    dashboard: webDashboard_(activeYear),
    packages: webPackageList_(activeYear),
    vendors: webTableDisplay_(CFG.SHEETS.DB_VENDOR),
    teams: webTableDisplay_(CFG.SHEETS.DB_TEAM),
    officers: webTableDisplay_(CFG.SHEETS.DB_OFFICER),
    user: user,
    users: (user.role === 'Owner' || user.role === 'Admin') ? webListUsers_() : [],
    calendar: (user.role === 'Owner' || user.role === 'Admin') ? webListCalendar_() : [],
    taxOptions: webTaxOptions_(),
    revision: sh_(CFG.SHEETS.UPDATE).getRange('B5').getDisplayValue() || 'Belum ditetapkan',
    webAppUrl: ScriptApp.getService().getUrl() || ''
  });
}

function webYears_() {
  const s = sh_(CFG.SHEETS.DB_PACKAGE);
  const values = s.getLastRow() < 2 ? [] : s.getRange(2, 2, s.getLastRow() - 1, 1).getValues().flat();
  const years = values.map(Number).filter(Boolean);
  years.push(new Date().getFullYear());
  return Array.from(new Set(years)).sort((a, b) => b - a);
}

function webTableDisplay_(sheetName) {
  const s = sh_(sheetName);
  if (s.getLastRow() < 2) return [];
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getDisplayValues()[0];
  return s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).getDisplayValues()
    .filter(r => r.some(Boolean))
    .map(r => {
      const o = {};
      headers.forEach((h, i) => { if (h) o[h] = r[i]; });
      return o;
    });
}

function webDashboard_(year) {
  const packages = webPackageList_(year);
  const pok = sh_(CFG.SHEETS.DB_POK);
  const pokRows = pok.getLastRow() < 2 ? [] : pok.getRange(2, 1, pok.getLastRow() - 1, 22).getValues().filter(r => r[0] && r[2] === 'Ya');
  const statuses = {Draft: 0, 'FP Dibuat': 0, Kontrak: 0, Selesai: 0};
  packages.forEach(p => { statuses[p.status] = (statuses[p.status] || 0) + 1; });
  const totalContract = packages.reduce((a, p) => a + Number(p.contract || 0), 0);
  const totalBudget = pokRows.reduce((a, r) => a + Number(r[17] || 0), 0);
  const totalCommitted = pokRows.reduce((a, r) => a + Number(r[18] || 0), 0);
  const monthValues = Array(12).fill(0);
  packages.forEach(p => {
    if (p.poDate) {
      const m = Number(String(p.poDate).slice(5, 7));
      if (m >= 1 && m <= 12) monthValues[m - 1] += Number(p.contract || 0);
    }
  });
  const allocationIds = new Set();
  const allocationSheet = sh_(CFG.SHEETS.ALLOCATION);
  if (allocationSheet.getLastRow() >= 5) {
    allocationSheet.getRange(5, 1, allocationSheet.getLastRow() - 4, 1).getDisplayValues().flat().filter(Boolean).forEach(v => allocationIds.add(v));
  }
  const issues = packages.filter(p =>
    (p.status !== 'Draft' && !allocationIds.has(p.id)) ||
    (p.status !== 'Draft' && Number(p.fp || 0) <= 0) ||
    Number(p.contract || 0) > Number(p.hps || 0)
  );
  return {
    totalPackages: packages.length,
    totalContract: totalContract,
    totalBudget: totalBudget,
    totalCommitted: totalCommitted,
    remainingBudget: totalBudget - totalCommitted,
    statuses: statuses,
    issueCount: issues.length,
    issues: issues.slice(0, 8),
    months: monthValues,
    vendors: webShare_(packages.map(p => p.vendor)),
    teams: webShare_(packages.map(p => p.team)),
    recent: packages.slice(0, 8)
  };
}

function webShare_(values) {
  const counts = {};
  values.filter(Boolean).forEach(v => { counts[v] = (counts[v] || 0) + 1; });
  const total = Object.values(counts).reduce((a, b) => a + b, 0) || 1;
  return Object.keys(counts).map(k => ({label: k, count: counts[k], share: counts[k] / total}))
    .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label)).slice(0, 6);
}

function webPackageList_(year) {
  const user = webRequireUser_();
  const canSeeAll = user.role === 'Owner' || user.role === 'Admin';
  const owners = webPackageOwnerMap_();
  const s = sh_(CFG.SHEETS.DB_PACKAGE);
  if (s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 31).getValues().filter(r =>
    r[0] && Number(r[1]) === Number(year) && (canSeeAll || String(owners[String(r[0])] || '').toLowerCase() === user.email)
  ).map(r => ({
    id: String(r[0]), year: Number(r[1]), seq: Number(r[2]), name: String(r[3] || ''), team: String(r[4] || ''),
    teamLead: String(r[5] || ''), teamNip: String(r[6] || ''), vendor: String(r[7] || ''), status: String(r[8] || ''),
    fp: Number(r[9] || 0), contract: Number(r[10] || 0), hps: Number(r[11] || 0), offer: Number(r[12] || 0),
    requestDate: webDate_(r[13]), poDate: webDate_(r[14]), bastDate: webDate_(r[15]), mak: String(r[16] || ''),
    procurementOfficer: String(r[22] || ''), ppk: String(r[24] || ''), note: String(r[28] || ''), updatedAt: webDateTime_(r[30]), ownerEmail: owners[String(r[0])] || ''
  })).sort((a, b) => b.seq - a.seq);
}

function webGetPackage(id) {
  webAssertPackageAccess_(id);
  const p = findObject_(CFG.SHEETS.DB_PACKAGE, 'ID Paket', id);
  if (!p) throw new Error('Paket tidak ditemukan: ' + id);
  const detailsSheet = sh_(CFG.SHEETS.DB_DETAIL);
  const details = detailsSheet.getLastRow() < 2 ? [] : detailsSheet.getRange(2, 1, detailsSheet.getLastRow() - 1, 15).getValues()
    .filter(r => String(r[0]) === String(id)).map(r => ({
      no: Number(r[1]), description: String(r[2] || ''), specification: String(r[3] || ''), quantity: Number(r[4] || 0), unit: String(r[5] || ''),
      netUnit: Number(r[6] || 0), netTotal: Number(r[7] || 0), hpsUnit: Number(r[8] || 0), hpsTotal: Number(r[9] || 0),
      offerUnit: Number(r[10] || 0), offerTotal: Number(r[11] || 0), contractUnit: Number(r[12] || 0), contractTotal: Number(r[13] || 0), pokId: String(r[14] || '')
    }));
  const allocationsSheet = sh_(CFG.SHEETS.ALLOCATION);
  const pokMak = webPokMakMap_();
  const allocations = allocationsSheet.getLastRow() < 5 ? [] : allocationsSheet.getRange(5, 1, allocationsSheet.getLastRow() - 4, 10).getValues()
    .filter(r => String(r[0]) === String(id)).map(r => ({pokId: String(r[1] || ''), code: String(r[2] || ''), mak: String(pokMak[String(r[1])] || ''), description: String(r[3] || ''), fp: Number(r[4] || 0), contract: Number(r[5] || 0), commitment: Number(r[7] || 0), remaining: Number(r[8] || 0), note: String(r[9] || '')}));
  const letterSheet = sh_(CFG.SHEETS.DB_LETTER);
  const existingLetters = letterSheet.getLastRow() < 2 ? [] : letterSheet.getRange(2, 1, letterSheet.getLastRow() - 1, 6).getValues().filter(r => String(r[0]) === String(id));
  const letters = WEB_LETTER_TYPES.map(type => {
    const row = existingLetters.find(r => String(r[1]) === type) || [];
    return {type: type, date: webDate_(row[2]), number: String(row[3] || ''), pic: String(row[4] || ''), note: String(row[5] || '')};
  });
  const ext = getPackageExt_(id);
  return webSafe_({
    package: {
      id: p['ID Paket'], year: Number(p['Tahun']), seq: Number(p['Nomor Urut']), name: p['Nama Pengadaan'], team: p['Tim'], teamLead: p['Ketua Tim'], teamNip: String(p['NIP Ketua Tim']),
      vendor: p['Penyedia'], status: p['Status'], fp: Number(p['Nilai FP'] || 0), contract: Number(p['Nilai Kontrak'] || 0), hps: Number(p['Nilai HPS'] || 0), offer: Number(p['Nilai Penawaran'] || 0),
      requestDate: webDate_(p['Tanggal Permintaan']), poDate: webDate_(p['Tanggal PO']), bastDate: webDate_(p['Tanggal BAST']), hpsDate: ext.hpsDate, arrivalDate: ext.arrivalDate,
      mak: p['MAK Utama'], procurementOfficer: p['Pejabat Pengadaan'], ppk: p['PPK'], note: p['Catatan'], ownerEmail: ext.ownerEmail
    }, details: details, allocations: allocations, letters: letters, costs: ext.costs, archives: webArchives_(id)
  });
}

function webNextPackageId(year) {
  webRequireUser_();
  const y = Number(year || new Date().getFullYear());
  const s = sh_(CFG.SHEETS.DB_PACKAGE);
  const rows = s.getLastRow() < 2 ? [] : s.getRange(2, 1, s.getLastRow() - 1, 3).getValues();
  const seq = rows.filter(r => Number(r[1]) === y).reduce((m, r) => Math.max(m, Number(r[2] || 0)), 0) + 1;
  return {id: 'PKG-' + y + '-' + String(seq).padStart(3, '0'), year: y, seq: seq};
}

function webSavePackage(payload) {
  const user = webRequireUser_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const raw = payload && payload.package ? payload.package : {};
    const team = findObject_(CFG.SHEETS.DB_TEAM, 'Tim', raw.team) || {};
    const x = {
      id: String(raw.id || '').trim(), year: Number(raw.year), seq: Number(raw.seq), name: String(raw.name || '').trim(),
      team: String(raw.team || '').trim(), teamLead: String(raw.teamLead || team['Ketua Tim'] || '').trim(), teamNip: String(raw.teamNip || team['NIP'] || '').trim(),
      vendor: String(raw.vendor || '').trim(), status: String(raw.status || 'Draft').trim(), requestDate: webParseDate_(raw.requestDate), poDate: webParseDate_(raw.poDate), bastDate: webParseDate_(raw.bastDate),
      hps: webNumber_(raw.hps), offer: webNumber_(raw.offer), contract: webNumber_(raw.contract), fp: webNumber_(raw.fp), mak: String(raw.mak || '').trim(),
      procurementOfficer: String(raw.procurementOfficer || '').trim(), ppk: String(raw.ppk || '').trim(), note: String(raw.note || '').trim()
    };
    const db = sh_(CFG.SHEETS.DB_PACKAGE);
    const row = findRow_(db, 1, x.id);
    if (row) webAssertPackageAccess_(x.id);
    validateInput_(x, Boolean(row));
    const old = row ? rowObject_(db, row) : null;
    const targetRow = row || firstBlankRow_(db, 1, 2);
    db.getRange(targetRow, 1, 1, 31).setValues([packageRowValues_(x, old)]);
    webSaveDetails_(x.id, payload.details || []);
    webSaveAllocations_(x.id, x.status, payload.allocations || []);
    const savedLetters = webSaveLetters_(x.id, payload.letters || []);
    webSavePackageExt_(x.id, row ? null : user.email, raw.hpsDate, raw.arrivalDate, payload.costs || {});
    SpreadsheetApp.flush();
    segarkanDashboard();
    return {ok: true, id: x.id, letters: savedLetters, message: row ? 'Paket berhasil diperbarui.' : 'Paket baru berhasil disimpan.'};
  } finally {
    lock.releaseLock();
  }
}

function webSaveDetails_(id, details) {
  const s = sh_(CFG.SHEETS.DB_DETAIL);
  webClearMatchingRows_(s, 1, id, 2, 15);
  details.filter(d => d && d.description).forEach((d, i) => {
    const q = webNumber_(d.quantity);
    const net = webNumber_(d.netUnit), hps = webNumber_(d.hpsUnit), offer = webNumber_(d.offerUnit), contract = webNumber_(d.contractUnit);
    s.getRange(firstBlankRow_(s, 1, 2), 1, 1, 15).setValues([[
      id, Number(d.no || i + 1), String(d.description || ''), String(d.specification || ''), q, String(d.unit || ''),
      net, q * net, hps, q * hps, offer, q * offer, contract, q * contract, String(d.pokId || '')
    ]]);
  });
}

function webSaveAllocations_(id, status, allocations) {
  const s = sh_(CFG.SHEETS.ALLOCATION);
  webClearMatchingRows_(s, 1, id, 5, 10);
  allocations.filter(a => a && a.pokId).forEach(a => {
    const row = firstBlankRow_(s, 1, 5);
    s.getRange(row, 1, 1, 10).setValues([[id, String(a.pokId), String(a.code || ''), String(a.description || ''), webNumber_(a.fp), webNumber_(a.contract), status, '', '', String(a.note || '')]]);
    s.getRange(row, 8).setFormula('=IF(A' + row + '="","",IF(G' + row + '="Draft",0,IF(G' + row + '="FP Dibuat",E' + row + ',IF(OR(G' + row + '="Kontrak",G' + row + '="Selesai"),F' + row + ',0))))');
    s.getRange(row, 9).setFormula('=IF(B' + row + '="","",IFERROR(INDEX(\'DB_POK\'!$T$2:$T$1000,MATCH(B' + row + ',\'DB_POK\'!$A$2:$A$1000,0)),""))');
  });
}

function webSaveLetters_(id, letters) {
  const s = sh_(CFG.SHEETS.DB_LETTER);
  webClearMatchingRows_(s, 1, id, 2, 6);
  const reserved = {};
  letters.filter(l => l && (l.number || l.date)).forEach(l => {
    const date = webParseDate_(l.date);
    let number = String(l.number || '').trim();
    if (!number && validDate_(date) && String(l.type) !== 'Invoice') number = webNextLetterNumber_(date, id, reserved);
    l.number = number;
    s.getRange(firstBlankRow_(s, 1, 2), 1, 1, 6).setValues([[id, String(l.type || ''), date, number, String(l.pic || ''), String(l.note || '')]]);
  });
  return WEB_LETTER_TYPES.map(type => {
    const found = letters.find(l => l.type === type) || {};
    return {type: type, date: found.date || '', number: found.number || '', pic: found.pic || '', note: found.note || ''};
  });
}

function webClearMatchingRows_(sheet, column, value, startRow, width) {
  if (sheet.getLastRow() < startRow) return;
  const vals = sheet.getRange(startRow, column, sheet.getLastRow() - startRow + 1, 1).getDisplayValues().flat();
  vals.forEach((v, i) => { if (String(v) === String(value)) sheet.getRange(startRow + i, 1, 1, width).clearContent(); });
}

function webSearchPok(query, limit) {
  const q = String(query || '').toLowerCase().trim();
  if (q.length < 2) return [];
  const tokens = q.split(/\s+/).filter(Boolean);
  const s = sh_(CFG.SHEETS.DB_POK);
  if (s.getLastRow() < 2) return [];
  const rows = s.getRange(2, 1, s.getLastRow() - 1, 22).getValues();
  const matches = rows.filter(r => r[0] && String(r[2]) === 'Ya' && tokens.every(t => String(r[21] || (r[11] + ' ' + r[12] + ' ' + r[13])).toLowerCase().includes(t)))
    .slice(0, Math.min(Number(limit || 30), 50)).map(r => ({
      id: String(r[0]), revision: String(r[1]), code: String(r[11]), mak: String(r[12]), description: String(r[13]), volume: Number(r[14] || 0), unit: String(r[15]), unitPrice: Number(r[16] || 0), budget: Number(r[17] || 0), commitment: Number(r[18] || 0), remaining: Number(r[19] || 0), status: String(r[20] || '')
    }));
  return webSafe_(matches);
}

function webCreateDocuments(id) {
  const user = webAssertPackageAccess_(id);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ctl = sh_(CFG.SHEETS.DOCS);
    const templateId = extractId_(ctl.getRange('B5').getDisplayValue());
    const folderId = extractId_(ctl.getRange('B6').getDisplayValue());
    const version = Number(ctl.getRange('B7').getValue() || 1);
    if (!findObject_(CFG.SHEETS.DB_PACKAGE, 'ID Paket', id)) throw new Error('Paket tidak ditemukan: ' + id);
    if (!templateId || !folderId) throw new Error('Template atau folder Shared Drive belum diatur pada sheet Dokumen.');
    const data = collectMergeData_(id);
    const folder = DriveApp.getFolderById(folderId);
    const title = id + '_Dokumen_Pengadaan_v' + version;
    const copy = DriveApp.getFileById(templateId).makeCopy(title, folder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();
    Object.keys(data.map).forEach(k => body.replaceText('\\{\\{' + escapeRegex_(k) + '\\}\\}', String(data.map[k] == null ? '' : data.map[k])));
    insertDetailTable_(body, data.details);
    doc.saveAndClose();
    const pdf = folder.createFile(copy.getAs(MimeType.PDF).setName(title + '.pdf'));
    try { copy.addEditor(user.email); } catch (e) {}
    try { pdf.addViewer(user.email); } catch (e) {}
    const archive = sh_(CFG.SHEETS.ARCHIVE_DOCS);
    const row = firstBlankRow_(archive, 1, 5);
    archive.getRange(row, 1, 1, 10).setValues([['ARS-' + id.replace('PKG-', '') + '-v' + version, id, version, 'Dokumen Pengadaan Gabungan', new Date(), user.email, copy.getUrl(), pdf.getUrl(), 'Selesai', 'Dibuat dari web app ke Shared Drive kantor']]);
    ctl.getRange('B4').setValue(id);
    ctl.getRange('B7').setValue(version + 1);
    ctl.getRange('B8').setValue('Selesai: ' + title);
    SpreadsheetApp.flush();
    return {ok: true, title: title, version: version, docsUrl: copy.getUrl(), pdfUrl: pdf.getUrl(), archives: webArchives_(id)};
  } finally {
    lock.releaseLock();
  }
}

function webArchives_(id) {
  const s = sh_(CFG.SHEETS.ARCHIVE_DOCS);
  if (s.getLastRow() < 5) return [];
  const range = s.getRange(5, 1, s.getLastRow() - 4, 10);
  const values = range.getValues();
  const richText = range.getRichTextValues();
  const formulas = range.getFormulas();
  return values.map((r, i) => ({
    archiveId: String(r[0] || ''),
    packageId: String(r[1] || ''),
    version: Number(r[2] || 0),
    type: String(r[3] || ''),
    createdAt: webDateTime_(r[4]),
    author: String(r[5] || ''),
    // getValues() hanya mengembalikan label seperti "Buka PDF" pada sel
    // hyperlink. Ambil target URL dari rich text/formula agar tombol web valid.
    docsUrl: webLinkUrl_(r[6], richText[i][6], formulas[i][6]),
    pdfUrl: webLinkUrl_(r[7], richText[i][7], formulas[i][7]),
    status: String(r[8] || '')
  })).filter(r => !id || r.packageId === String(id)).reverse();
}

function webLinkUrl_(value, richText, formula) {
  let url = '';
  if (richText && typeof richText.getLinkUrl === 'function') {
    url = richText.getLinkUrl() || '';
  }
  if (!url && richText && typeof richText.getRuns === 'function') {
    const linkedRun = richText.getRuns().find(run =>
      run && typeof run.getLinkUrl === 'function' && run.getLinkUrl()
    );
    if (linkedRun) url = linkedRun.getLinkUrl();
  }
  if (!url && formula) {
    const match = String(formula).match(/^=HYPERLINK\(\s*"((?:[^"]|"")+)"/i);
    if (match) url = match[1].replace(/""/g, '"');
  }
  if (!url && /^https:\/\//i.test(String(value || '').trim())) {
    url = String(value).trim();
  }
  return /^https:\/\/(?:docs|drive)\.google\.com\//i.test(url) ? url : '';
}

function ensureWebSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  [CFG.SHEETS.DB_USER, CFG.SHEETS.DB_PACKAGE_EXT, CFG.SHEETS.DB_CALENDAR, CFG.SHEETS.IMPORT_LOG].forEach(name => {
    if (!ss.getSheetByName(name)) throw new Error('Pembaruan database belum dijalankan. Owner harus menjalankan upgradeSystemV2() satu kali dari Apps Script.');
  });
}

function webCurrentEmail_() {
  const key = webAuthKey_();
  const email = String(CacheService.getScriptCache().get('AUTH_' + key) || '').trim().toLowerCase();
  if (!email) throw new Error('AUTH_REQUIRED|Silakan masuk dengan Gmail yang telah didaftarkan.');
  return email;
}

function webAuthKey_() {
  const key = String(Session.getTemporaryActiveUserKey() || '');
  if (!key) throw new Error('Sesi Google tidak tersedia. Pastikan deployment hanya dapat diakses pengguna yang masuk dengan Google Account.');
  return key;
}

function webStartLogin(email) {
  ensureWebSchema_();
  email = String(email || '').trim().toLowerCase();
  if (!/^[^\s@]+@(?:gmail\.com|googlemail\.com)$/i.test(email)) throw new Error('Gunakan alamat Gmail yang valid.');
  const row = findObject_(CFG.SHEETS.DB_USER, 'Email', email);
  if (!row || String(row['Aktif']).toLowerCase() !== 'ya') throw new Error('Email belum terdaftar atau sedang nonaktif.');
  const cache = CacheService.getScriptCache();
  const key = webAuthKey_();
  if (cache.get('OTP_WAIT_' + key)) throw new Error('Tunggu satu menit sebelum meminta kode baru.');
  const code = String(Math.floor(100000 + Math.random() * 900000));
  cache.put('OTP_' + key, email + '|' + code, 600);
  cache.put('OTP_WAIT_' + key, '1', 60);
  MailApp.sendEmail({
    to: email,
    subject: 'Kode masuk Sistem Pengadaan Example Public Organization',
    htmlBody: '<p>Kode masuk Anda:</p><p style="font:700 28px Arial;letter-spacing:5px">' + code + '</p><p>Berlaku 10 menit. Jangan berikan kode ini kepada orang lain.</p>',
    name: 'Sistem Pengadaan Example Public Organization'
  });
  return {ok:true, message:'Kode 6 digit telah dikirim ke ' + email + '.'};
}

function webVerifyLogin(email, code) {
  email = String(email || '').trim().toLowerCase();
  code = String(code || '').replace(/\D/g, '');
  const cache = CacheService.getScriptCache();
  const key = webAuthKey_();
  const saved = String(cache.get('OTP_' + key) || '');
  if (!saved || saved !== email + '|' + code) throw new Error('Kode salah atau sudah kedaluwarsa.');
  cache.remove('OTP_' + key);
  cache.put('AUTH_' + key, email, 21600);
  return {ok:true, message:'Berhasil masuk sebagai ' + email + '.'};
}

function webLogout() {
  CacheService.getScriptCache().remove('AUTH_' + webAuthKey_());
  return {ok:true};
}

function webRequireUser_() {
  ensureWebSchema_();
  const email = webCurrentEmail_();
  const row = findObject_(CFG.SHEETS.DB_USER, 'Email', email);
  if (!row || String(row['Aktif']).toLowerCase() !== 'ya') {
    throw new Error('Akun ' + email + ' belum terdaftar atau sedang nonaktif. Hubungi Owner aplikasi.');
  }
  return {
    email: email,
    name: String(row['Nama'] || email),
    role: String(row['Peran'] || 'User'),
    canUpdateBudget: String(row['Boleh Update Anggaran']).toLowerCase() === 'ya'
  };
}

function webListUsers_() {
  return webTableDisplay_(CFG.SHEETS.DB_USER).map(r => ({
    email: String(r.Email || '').toLowerCase(), name: r.Nama || '', role: r.Peran || 'User',
    active: String(r.Aktif).toLowerCase() === 'ya', canUpdateBudget: String(r['Boleh Update Anggaran']).toLowerCase() === 'ya'
  }));
}

function webListUsers() {
  const actor = webRequireUser_();
  if (actor.role !== 'Owner' && actor.role !== 'Admin') throw new Error('Menu pengguna hanya untuk Owner/Admin.');
  return webSafe_(webListUsers_());
}

function webSaveUser(payload) {
  const actor = webRequireUser_();
  if (actor.role !== 'Owner' && actor.role !== 'Admin') throw new Error('Hanya Owner/Admin yang dapat mengelola pengguna.');
  const email = String(payload && payload.email || '').trim().toLowerCase();
  const name = String(payload && payload.name || '').trim();
  let role = String(payload && payload.role || 'User');
  if (!/^[^\s@]+@(?:gmail\.com|googlemail\.com)$/i.test(email)) throw new Error('Gunakan alamat Gmail yang valid.');
  const s = sh_(CFG.SHEETS.DB_USER);
  const existingRow = findRow_(s, 1, email);
  const existing = existingRow ? rowObject_(s, existingRow) : null;
  if (existing && String(existing.Peran) === 'Owner' && actor.email !== email) throw new Error('Owner tidak dapat diubah oleh akun lain.');
  if (actor.role === 'Admin') {
    if (existing && String(existing.Peran) !== 'User') throw new Error('Admin hanya dapat mengelola akun User.');
    role = 'User';
  }
  if (!['Owner','Admin','User'].includes(role)) role = 'User';
  if (role === 'Owner' && email !== 'owner@example.com') throw new Error('Owner tetap Example Owner.');
  const target = existingRow || firstBlankRow_(s, 1, 2);
  s.getRange(target, 1, 1, 7).setValues([[
    email, name || email, role, payload.active === false ? 'Tidak' : 'Ya', payload.canUpdateBudget ? 'Ya' : 'Tidak',
    existing && existing['Ditambahkan Pada'] ? existing['Ditambahkan Pada'] : new Date(), actor.email
  ]]);
  if (payload.active !== false) webShareResourcesWithUser_(email);
  return {ok: true, users: webListUsers_(), message: 'Pengguna ' + email + ' berhasil disimpan.'};
}

function webShareResourcesWithUser_(email) {
  const ctl = sh_(CFG.SHEETS.DOCS);
  const templateId = extractId_(ctl.getRange('B5').getDisplayValue());
  const folderId = extractId_(ctl.getRange('B6').getDisplayValue());
  try { if (templateId) DriveApp.getFileById(templateId).addViewer(email); } catch (e) {}
  try { if (folderId) DriveApp.getFolderById(folderId).addEditor(email); } catch (e) {}
}

function webPackageOwnerMap_() {
  const s = sh_(CFG.SHEETS.DB_PACKAGE_EXT);
  if (s.getLastRow() < 2) return {};
  const out = {};
  s.getRange(2, 1, s.getLastRow() - 1, 2).getDisplayValues().forEach(r => { if (r[0]) out[String(r[0])] = String(r[1] || '').toLowerCase(); });
  return out;
}

function getPackageExt_(id) {
  ensureWebSchema_();
  const row = findObject_(CFG.SHEETS.DB_PACKAGE_EXT, 'ID Paket', id) || {};
  let costs = {};
  try { costs = JSON.parse(String(row['Biaya JSON'] || '{}')); } catch (e) { costs = {}; }
  return {
    ownerEmail: String(row['Email Pemilik'] || '').toLowerCase(),
    hpsDate: webDate_(row['Tanggal Penetapan HPS']),
    arrivalDate: webDate_(row['Tanggal Barang Sampai']),
    costs: Object.assign({shipping:0, administration:0, stampEnabled:true, stampAmount:10000, taxes:[]}, costs)
  };
}

function webSavePackageExt_(id, ownerEmail, hpsDate, arrivalDate, costs) {
  const s = sh_(CFG.SHEETS.DB_PACKAGE_EXT);
  const row = findRow_(s, 1, id);
  const old = row ? rowObject_(s, row) : {};
  const owner = String(ownerEmail || old['Email Pemilik'] || webCurrentEmail_()).toLowerCase();
  const safeCosts = {
    shipping: Math.round(webNumber_(costs.shipping)), administration: Math.round(webNumber_(costs.administration)),
    stampEnabled: costs.stampEnabled !== false, stampAmount: Math.round(webNumber_(costs.stampAmount || 10000)),
    taxes: Array.isArray(costs.taxes) ? costs.taxes.slice(0, 12).map(t => ({key:String(t.key||''),label:String(t.label||'').slice(0,80),rate:Math.max(0,Math.min(100,Number(t.rate||0))),kind:t.kind==='addition'?'addition':'withholding',checked:Boolean(t.checked)})) : []
  };
  s.getRange(row || firstBlankRow_(s, 1, 2), 1, 1, 6).setValues([[
    id, owner, webParseDate_(hpsDate), webParseDate_(arrivalDate), JSON.stringify(safeCosts), new Date()
  ]]);
}

function webAssertPackageAccess_(id) {
  const user = webRequireUser_();
  if (user.role === 'Owner' || user.role === 'Admin') return user;
  const owner = webPackageOwnerMap_()[String(id)] || '';
  if (owner !== user.email) throw new Error('Paket ini bukan milik akun ' + user.email + '.');
  return user;
}

function webPokMakMap_() {
  const s = sh_(CFG.SHEETS.DB_POK);
  if (s.getLastRow() < 2) return {};
  const out = {};
  s.getRange(2, 1, s.getLastRow() - 1, 13).getDisplayValues().forEach(r => { if (r[0]) out[String(r[0])] = String(r[12] || ''); });
  return out;
}

function webTaxOptions_() {
  return [
    {key:'printing-2.5', label:'PPh percetakan (tarif kerja 2,5% — verifikasi bendahara)', rate:2.5, kind:'withholding', recommended:true},
    {key:'pph22-1.5', label:'PPh 22 pembelian barang 1,5%', rate:1.5, kind:'withholding'},
    {key:'pph23-2', label:'PPh 23 jasa 2%', rate:2, kind:'withholding'},
    {key:'ppn-nonluxury', label:'PPN nonmewah efektif 11%', rate:11, kind:'addition'},
    {key:'ppn-luxury', label:'PPN mewah 12%', rate:12, kind:'addition'}
  ];
}

function webGenerateLetterSchedule(id, hpsIso, arrivalIso) {
  if (findObject_(CFG.SHEETS.DB_PACKAGE, 'ID Paket', id)) webAssertPackageAccess_(id); else webRequireUser_();
  if (!hpsIso || !arrivalIso) throw new Error('Isi Tanggal Penetapan HPS dan Tanggal Barang Sampai.');
  const hps = rollToWorkday_(webParseDate_(hpsIso));
  const arrival = webParseDate_(arrivalIso);
  const request = addCalendarThenRoll_(hps, 1);
  const order = addCalendarThenRoll_(request, 1);
  const inspection = addCalendarThenRoll_(arrival, 1);
  const bast = addCalendarThenRoll_(inspection, 1);
  const receipt = addCalendarThenRoll_(arrival, 3);
  const administration = addCalendarThenRoll_(receipt, 7);
  const payment = addCalendarThenRoll_(administration, 7);
  const dates = {
    'Penetapan HPS':hps, 'Permintaan Pengadaan':request, 'Surat Pesanan':order,
    'Pemeriksaan Pekerjaan':inspection, 'BAST':bast, 'Kuitansi':receipt, 'Invoice':receipt,
    'Pemeriksaan Administrasi':administration, 'Pembayaran':payment
  };
  const reserved = {};
  const letters = WEB_LETTER_TYPES.map(type => ({
    type:type, date:webDate_(dates[type]), number:type === 'Invoice' ? '' : webNextLetterNumber_(dates[type], id, reserved), pic:'', note:''
  }));
  return webSafe_({letters:letters, requestDate:webDate_(request), poDate:webDate_(order), bastDate:webDate_(bast), hpsDate:webDate_(hps), arrivalDate:webDate_(arrival)});
}

function addCalendarThenRoll_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + Number(days || 0));
  return rollToWorkday_(d);
}

function rollToWorkday_(date) {
  const d = new Date(date);
  const holidays = webActiveHolidaySet_();
  for (let i = 0; i < 370; i++) {
    const iso = Utilities.formatDate(d, CFG.TZ, 'yyyy-MM-dd');
    if (d.getDay() !== 0 && d.getDay() !== 6 && !holidays[iso]) return d;
    d.setDate(d.getDate() + 1);
  }
  throw new Error('Kalender kerja tidak dapat dihitung.');
}

function webActiveHolidaySet_() {
  const s = sh_(CFG.SHEETS.DB_CALENDAR);
  const out = {};
  if (s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 4).getValues().forEach(r => {
    if (validDate_(r[0]) && String(r[3]).toLowerCase() === 'ya') out[Utilities.formatDate(new Date(r[0]), CFG.TZ, 'yyyy-MM-dd')] = true;
  });
  return out;
}

function webNextLetterNumber_(date, packageId, reserved) {
  const iso = Utilities.formatDate(new Date(date), CFG.TZ, 'yyyy-MM-dd');
  let max = Number(reserved[iso] || 0);
  const s = sh_(CFG.SHEETS.DB_LETTER);
  if (s.getLastRow() >= 2) {
    s.getRange(2, 1, s.getLastRow() - 1, 4).getValues().forEach(r => {
      if (String(r[0]) === String(packageId) || !validDate_(r[2]) || Utilities.formatDate(new Date(r[2]), CFG.TZ, 'yyyy-MM-dd') !== iso) return;
      const m = String(r[3] || '').match(/^B\.\d{2}-\d{2}\.(\d+)/i);
      if (m) max = Math.max(max, Number(m[1]));
    });
  }
  const seq = max + 1;
  reserved[iso] = seq;
  const d = new Date(date);
  const dd = Utilities.formatDate(d, CFG.TZ, 'dd-MM');
  const mm = Utilities.formatDate(d, CFG.TZ, 'MM');
  const yyyy = Utilities.formatDate(d, CFG.TZ, 'yyyy');
  return 'B.' + dd + '.' + String(seq).padStart(2, '0') + '/UNIT/CLASS/' + mm + '/' + yyyy;
}

function webListCalendar_() {
  const s = sh_(CFG.SHEETS.DB_CALENDAR);
  if (s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 7).getValues().filter(r => r[0]).map(r => ({
    date:webDate_(r[0]), label:String(r[1]||''), type:String(r[2]||''), active:String(r[3]).toLowerCase()==='ya', source:String(r[4]||'')
  })).sort((a,b)=>a.date.localeCompare(b.date));
}

function webSaveCalendar(payload) {
  const actor = webRequireUser_();
  if (actor.role !== 'Owner' && actor.role !== 'Admin') throw new Error('Kalender hanya dapat dikelola Owner/Admin.');
  const date = webParseDate_(payload.date);
  if (!validDate_(date)) throw new Error('Tanggal kalender tidak valid.');
  const s = sh_(CFG.SHEETS.DB_CALENDAR);
  let row = 0;
  if (s.getLastRow() >= 2) {
    const iso = webDate_(date);
    s.getRange(2,1,s.getLastRow()-1,1).getValues().some((r,i) => {
      if (webDate_(r[0]) === iso) { row = i + 2; return true; }
      return false;
    });
  }
  s.getRange(row || firstBlankRow_(s,1,2),1,1,7).setValues([[
    date, String(payload.label||'Hari nonkerja'), String(payload.type||'Khusus'), payload.active===false?'Tidak':'Ya',
    String(payload.source||'Input Owner/Admin'), actor.email, new Date()
  ]]);
  return {ok:true, calendar:webListCalendar_(), message:'Kalender kerja diperbarui.'};
}

function webAssertCanUpdateBudget_() {
  const user = webRequireUser_();
  if (!user.canUpdateBudget) throw new Error('Akun ini tidak memiliki izin Update Anggaran.');
  return user;
}

function webUploadBudgetExcel(payload) {
  const user = webAssertCanUpdateBudget_();
  const name = String(payload && payload.name || '').replace(/[^a-zA-Z0-9._ -]/g,'_');
  if (!/\.xlsx?$/i.test(name)) throw new Error('Unggah file Excel .xlsx atau .xls.');
  const bytes = Utilities.base64Decode(String(payload.base64 || ''));
  if (!bytes.length || bytes.length > 20 * 1024 * 1024) throw new Error('File kosong atau melebihi batas 20 MB.');
  const mime = /\.xls$/i.test(name) ? 'application/vnd.ms-excel' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const blob = Utilities.newBlob(bytes, mime, name);
  const folderId = extractId_(sh_(CFG.SHEETS.DOCS).getRange('B6').getDisplayValue());
  if (!folderId) throw new Error('Folder arsip belum diatur pada sheet Dokumen.');
  const folder = DriveApp.getFolderById(folderId);
  const archive = folder.createFile(blob).setDescription('Arsip impor POK oleh ' + user.email);
  let converted;
  if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.create === 'function') {
    converted = Drive.Files.create({name:'TMP_IMPORT_' + Date.now() + '_' + name, mimeType:MimeType.GOOGLE_SHEETS, parents:[folderId]}, blob, {fields:'id,name'});
  } else if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.insert === 'function') {
    converted = Drive.Files.insert({title:'TMP_IMPORT_' + Date.now() + '_' + name, mimeType:MimeType.GOOGLE_SHEETS, parents:[{id:folderId}]}, blob, {convert:true});
  } else {
    throw new Error('Layanan Drive API belum aktif. Aktifkan Advanced Google Service “Drive API” pada proyek Apps Script.');
  }
  const source = SpreadsheetApp.openById(converted.id);
  return webSafe_({token:converted.id, name:name, archiveUrl:archive.getUrl(), sheets:source.getSheets().map(s=>s.getName())});
}

function webPreviewBudgetImport(token, sheetName, archiveUrl, fileName) {
  webAssertCanUpdateBudget_();
  const source = SpreadsheetApp.openById(String(token));
  const sheet = source.getSheetByName(String(sheetName));
  if (!sheet) throw new Error('Sheet sumber tidak ditemukan.');
  const parsed = parsePokSheet_(sheet, String(sheetName));
  if (!parsed.length) throw new Error('Tidak ada rincian POK yang terbaca pada sheet terpilih.');
  const current = sh_(CFG.SHEETS.DB_POK).getRange(2,1,Math.max(sh_(CFG.SHEETS.DB_POK).getLastRow()-1,1),22).getValues().filter(r=>r[0]);
  const oldMap = new Map(current.map(r=>[String(r[0]),r]));
  const newMap = new Map(parsed.map(r=>[String(r[0]),r]));
  let added=0, changed=0, removed=0;
  newMap.forEach((r,k)=>{ if(!oldMap.has(k)) added++; else if(Number(oldMap.get(k)[17]||0)!==Number(r[17]||0)) changed++; });
  oldMap.forEach((r,k)=>{ if(!newMap.has(k)) removed++; });
  return webSafe_({token:String(token),sheetName:String(sheetName),archiveUrl:String(archiveUrl||''),fileName:String(fileName||''),revision:String(sheetName),rows:parsed.length,total:parsed.reduce((a,r)=>a+Number(r[17]||0),0),added:added,changed:changed,removed:removed});
}

function webApplyBudgetImport(preview) {
  const user = webAssertCanUpdateBudget_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const source = SpreadsheetApp.openById(String(preview.token));
    const sheet = source.getSheetByName(String(preview.sheetName));
    if (!sheet) throw new Error('Sheet sumber tidak ditemukan.');
    const revision = String(preview.revision || preview.sheetName);
    const fresh = parsePokSheet_(sheet, revision);
    activateWebPokRevision_(fresh, revision, String(preview.archiveUrl||''), user.email);
    const log = sh_(CFG.SHEETS.IMPORT_LOG);
    log.getRange(firstBlankRow_(log,1,2),1,1,9).setValues([[
      new Date(), user.email, String(preview.fileName||''), String(preview.sheetName||''), revision, fresh.length,
      fresh.reduce((a,r)=>a+Number(r[17]||0),0), String(preview.archiveUrl||''), 'Aktif'
    ]]);
    return {ok:true,revision:revision,message:'Revisi ' + revision + ' berhasil diaktifkan. File Excel asli tetap diarsipkan.'};
  } finally { lock.releaseLock(); }
}

function activateWebPokRevision_(fresh, revision, sourceUrl, email) {
  if (!fresh.length) throw new Error('Data POK kosong.');
  const db = sh_(CFG.SHEETS.DB_POK);
  const old = db.getRange(2,1,Math.max(db.getLastRow()-1,1),22).getValues().filter(r=>r[0]);
  const archive = sh_(CFG.SHEETS.ARCHIVE_POK);
  if (old.length) archive.getRange(firstBlankRow_(archive,1,2),1,old.length,23).setValues(old.map(r=>r.concat([new Date()])));
  if (db.getLastRow()>1) db.getRange(2,1,db.getLastRow()-1,22).clearContent();
  db.getRange(2,1,fresh.length,18).setValues(fresh);
  fresh.forEach((r,i)=>{
    const row=i+2;
    db.getRange(row,19).setFormula('=SUMIF(\'Alokasi_Anggaran\'!$B$5:$B$500,A' + row + ',\'Alokasi_Anggaran\'!$H$5:$H$500)');
    db.getRange(row,20).setFormula('=R' + row + '-S' + row);
    db.getRange(row,21).setFormula('=IF(T' + row + '<0,"Tidak Cukup",IF(T' + row + '=0,"Habis","Tersedia"))');
    db.getRange(row,22).setValue(searchTerms_(r[11],r[12],r[13]));
  });
  rebuildDbSearch_();
  const hist=sh_(CFG.SHEETS.HISTORY_POK);
  if(hist.getLastRow()>1) hist.getRange(2,2,hist.getLastRow()-1,1).setValue('Tidak');
  hist.getRange(firstBlankRow_(hist,1,2),1,1,8).setValues([[
    revision,'Ya',new Date(),sourceUrl,fresh.length,fresh.filter(r=>Number(r[17])>0).length,fresh.reduce((a,r)=>a+Number(r[17]||0),0),'Impor web oleh '+email
  ]]);
  sh_(CFG.SHEETS.UPDATE).getRange('B5').setValue(revision);
  sh_(CFG.SHEETS.UPDATE).getRange('B9').setValue('Revisi aktif: '+revision);
  SpreadsheetApp.flush();
}

function webGetArchives(id) { webAssertPackageAccess_(id); return webSafe_(webArchives_(id)); }
function webNumber_(v) { return Number(String(v == null ? '' : v).replace(/[^0-9.-]/g, '')) || 0; }
function webParseDate_(v) { return v ? new Date(String(v).slice(0, 10) + 'T00:00:00+07:00') : ''; }
function webDate_(v) { return validDate_(v) ? Utilities.formatDate(new Date(v), CFG.TZ, 'yyyy-MM-dd') : ''; }
function webDateTime_(v) {
  if (!validDate_(v)) return '';
  const d = new Date(v);
  return formatDate_(d) + ' ' + Utilities.formatDate(d, CFG.TZ, 'HH:mm');
}
function webSafe_(value) { return JSON.parse(JSON.stringify(value)); }

