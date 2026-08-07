/**
 * ProcureFlow — antarmuka satu pintu untuk spreadsheet, dokumen, dan PDF.
 * File ini dipasang pada proyek Apps Script yang sama dengan PPBJ_Automation.gs.
 */

const WEB_LETTER_TYPES = [
  'Penetapan HPS', 'Permintaan Pengadaan', 'Surat Pesanan', 'Pemeriksaan Pekerjaan',
  'BAST', 'Kuitansi', 'Pemeriksaan Administrasi', 'Pembayaran'
];

function doGet() {
  requireAuthorizedUser_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ProcureFlow — Sistem Administrasi Pengadaan')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Menolak akses jika allowlist belum dikonfigurasi atau pengguna tidak cocok.
 * Deploy sebagai "User accessing the web app" agar email pengguna tersedia.
 */
function requireAuthorizedUser_() {
  const props = PropertiesService.getScriptProperties();
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const allowedDomain = String(props.getProperty('ALLOWED_DOMAIN') || '').trim().toLowerCase().replace(/^@/, '');
  const allowedEmails = String(props.getProperty('ALLOWED_EMAILS') || '')
    .split(',').map(v => v.trim().toLowerCase()).filter(Boolean);
  if (!email) throw new Error('Identitas pengguna tidak tersedia. Deploy web app sebagai pengguna yang mengakses.');
  if (!allowedDomain && !allowedEmails.length) throw new Error('Allowlist belum diatur. Isi ALLOWED_DOMAIN atau ALLOWED_EMAILS pada Script Properties.');
  const domainMatch = allowedDomain && email.endsWith('@' + allowedDomain);
  if (!domainMatch && !allowedEmails.includes(email)) throw new Error('Akun ini tidak memiliki akses ke aplikasi.');
  return email;
}

function bukaAplikasiWeb() {
  const url = ScriptApp.getService().getUrl();
  const body = url
    ? '<p style="font:14px Arial">Aplikasi web siap dibuka.</p><p><a href="' + url + '" target="_blank" style="background:#075ca8;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none;font:600 14px Arial">Buka aplikasi PPBJ</a></p>'
    : '<p style="font:14px Arial">Web app belum di-deploy. Pilih <b>Deploy &gt; New deployment &gt; Web app</b> di Apps Script.</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(body).setWidth(420).setHeight(170), 'Aplikasi Web PPBJ');
}

function webBootstrap(year) {
  requireAuthorizedUser_();
  const activeYear = Number(year || new Date().getFullYear());
  return webSafe_({
    year: activeYear,
    years: webYears_(),
    dashboard: webDashboard_(activeYear),
    packages: webPackageList_(activeYear),
    vendors: webTableDisplay_(CFG.SHEETS.DB_VENDOR),
    teams: webTableDisplay_(CFG.SHEETS.DB_TEAM),
    officers: webTableDisplay_(CFG.SHEETS.DB_OFFICER),
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
  const s = sh_(CFG.SHEETS.DB_PACKAGE);
  if (s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 31).getValues().filter(r => r[0] && Number(r[1]) === Number(year)).map(r => ({
    id: String(r[0]), year: Number(r[1]), seq: Number(r[2]), name: String(r[3] || ''), team: String(r[4] || ''),
    teamLead: String(r[5] || ''), teamNip: String(r[6] || ''), vendor: String(r[7] || ''), status: String(r[8] || ''),
    fp: Number(r[9] || 0), contract: Number(r[10] || 0), hps: Number(r[11] || 0), offer: Number(r[12] || 0),
    requestDate: webDate_(r[13]), poDate: webDate_(r[14]), bastDate: webDate_(r[15]), mak: String(r[16] || ''),
    procurementOfficer: String(r[22] || ''), ppk: String(r[24] || ''), note: String(r[28] || ''), updatedAt: webDateTime_(r[30])
  })).sort((a, b) => b.seq - a.seq);
}

function webGetPackage(id) {
  requireAuthorizedUser_();
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
  const allocations = allocationsSheet.getLastRow() < 5 ? [] : allocationsSheet.getRange(5, 1, allocationsSheet.getLastRow() - 4, 10).getValues()
    .filter(r => String(r[0]) === String(id)).map(r => ({pokId: String(r[1] || ''), code: String(r[2] || ''), description: String(r[3] || ''), fp: Number(r[4] || 0), contract: Number(r[5] || 0), commitment: Number(r[7] || 0), remaining: Number(r[8] || 0), note: String(r[9] || '')}));
  const letterSheet = sh_(CFG.SHEETS.DB_LETTER);
  const existingLetters = letterSheet.getLastRow() < 2 ? [] : letterSheet.getRange(2, 1, letterSheet.getLastRow() - 1, 6).getValues().filter(r => String(r[0]) === String(id));
  const letters = WEB_LETTER_TYPES.map(type => {
    const row = existingLetters.find(r => String(r[1]) === type) || [];
    return {type: type, date: webDate_(row[2]), number: String(row[3] || ''), pic: String(row[4] || ''), note: String(row[5] || '')};
  });
  return webSafe_({
    package: {
      id: p['ID Paket'], year: Number(p['Tahun']), seq: Number(p['Nomor Urut']), name: p['Nama Pengadaan'], team: p['Tim'], teamLead: p['Ketua Tim'], teamNip: String(p['NIP Ketua Tim']),
      vendor: p['Penyedia'], status: p['Status'], fp: Number(p['Nilai FP'] || 0), contract: Number(p['Nilai Kontrak'] || 0), hps: Number(p['Nilai HPS'] || 0), offer: Number(p['Nilai Penawaran'] || 0),
      requestDate: webDate_(p['Tanggal Permintaan']), poDate: webDate_(p['Tanggal PO']), bastDate: webDate_(p['Tanggal BAST']), mak: p['MAK Utama'], procurementOfficer: p['Pejabat Pengadaan'], ppk: p['PPK'], note: p['Catatan']
    }, details: details, allocations: allocations, letters: letters, archives: webArchives_(id)
  });
}

function webNextPackageId(year) {
  requireAuthorizedUser_();
  const y = Number(year || new Date().getFullYear());
  const list = webPackageList_(y);
  const seq = list.reduce((m, p) => Math.max(m, Number(p.seq || 0)), 0) + 1;
  return {id: 'PKG-' + y + '-' + String(seq).padStart(3, '0'), year: y, seq: seq};
}

function webSavePackage(payload) {
  requireAuthorizedUser_();
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
    validateInput_(x, Boolean(row));
    const old = row ? rowObject_(db, row) : null;
    const targetRow = row || firstBlankRow_(db, 1, 2);
    db.getRange(targetRow, 1, 1, 31).setValues([packageRowValues_(x, old)]);
    webSaveDetails_(x.id, payload.details || []);
    webSaveAllocations_(x.id, x.status, payload.allocations || []);
    webSaveLetters_(x.id, payload.letters || []);
    SpreadsheetApp.flush();
    segarkanDashboard();
    return {ok: true, id: x.id, message: row ? 'Paket berhasil diperbarui.' : 'Paket baru berhasil disimpan.'};
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
  letters.filter(l => l && (l.number || l.date)).forEach(l => {
    s.getRange(firstBlankRow_(s, 1, 2), 1, 1, 6).setValues([[id, String(l.type || ''), webParseDate_(l.date), String(l.number || ''), String(l.pic || ''), String(l.note || '')]]);
  });
}

function webClearMatchingRows_(sheet, column, value, startRow, width) {
  if (sheet.getLastRow() < startRow) return;
  const vals = sheet.getRange(startRow, column, sheet.getLastRow() - startRow + 1, 1).getDisplayValues().flat();
  vals.forEach((v, i) => { if (String(v) === String(value)) sheet.getRange(startRow + i, 1, 1, width).clearContent(); });
}

function webSearchPok(query, limit) {
  requireAuthorizedUser_();
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
  requireAuthorizedUser_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ctl = sh_(CFG.SHEETS.DOCS);
    const templateId = extractId_(ctl.getRange('B5').getDisplayValue());
    const folderId = extractId_(ctl.getRange('B6').getDisplayValue());
    const version = Number(ctl.getRange('B7').getValue() || 1);
    if (!findObject_(CFG.SHEETS.DB_PACKAGE, 'ID Paket', id)) throw new Error('Paket tidak ditemukan: ' + id);
    if (!templateId || !folderId) throw new Error('Template atau folder arsip belum diatur pada sheet Dokumen.');
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
    const archive = sh_(CFG.SHEETS.ARCHIVE_DOCS);
    const row = firstBlankRow_(archive, 1, 5);
    archive.getRange(row, 1, 1, 10).setValues([['ARS-' + id.replace('PKG-', '') + '-v' + version, id, version, 'Dokumen Pengadaan Gabungan', new Date(), Session.getActiveUser().getEmail() || 'Pengguna', copy.getUrl(), pdf.getUrl(), 'Selesai', 'Dibuat dari web app']]);
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
    // Hyperlink cells expose their label through getValues(); retrieve the
    // actual target from rich text/formula before returning it to the client.
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

function webGetArchives(id) { requireAuthorizedUser_(); return webSafe_(webArchives_(id)); }
function webNumber_(v) { return Number(String(v == null ? '' : v).replace(/[^0-9.-]/g, '')) || 0; }
function webParseDate_(v) { return v ? new Date(String(v).slice(0, 10) + 'T00:00:00+07:00') : ''; }
function webDate_(v) { return validDate_(v) ? Utilities.formatDate(new Date(v), CFG.TZ, 'yyyy-MM-dd') : ''; }
function webDateTime_(v) { return validDate_(v) ? Utilities.formatDate(new Date(v), CFG.TZ, 'd MMM yyyy HH:mm') : ''; }
function webSafe_(value) { return JSON.parse(JSON.stringify(value)); }
