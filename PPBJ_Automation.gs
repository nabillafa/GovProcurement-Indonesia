/**
 * ProcureFlow — sistem administrasi pengadaan berbasis Google Workspace.
 * Versi portofolio: tidak memuat data, identitas, atau ID aset organisasi.
 * Dipasang satu kali melalui Extensions > Apps Script.
 */

const CFG = Object.freeze({
  TZ: 'Asia/Jakarta',
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '',
  ORGANIZATION_NAME: PropertiesService.getScriptProperties().getProperty('ORGANIZATION_NAME') || 'Nama Organisasi',
  ORGANIZATION_ADDRESS: PropertiesService.getScriptProperties().getProperty('ORGANIZATION_ADDRESS') || 'Alamat Organisasi',
  DIPA_REFERENCE: PropertiesService.getScriptProperties().getProperty('DIPA_REFERENCE') || 'Referensi anggaran organisasi',
  SHEETS: {
    HOME: 'Beranda', INPUT: 'Input_Paket', DETAIL: 'Detail_Barang', CALC: 'Kalkulator_Harga',
    SEARCH: 'Cari_POK', ALLOCATION: 'Alokasi_Anggaran', UPDATE: 'Update_POK', DOCS: 'Dokumen',
    ARCHIVE_DOCS: 'Arsip_Dokumen', DB_PACKAGE: 'DB_Paket', DB_DETAIL: 'DB_Detail',
    DB_LETTER: 'DB_Surat', DB_POK: 'DB_POK', ARCHIVE_POK: 'Arsip_POK', HISTORY_POK: 'Riwayat_POK',
    DB_VENDOR: 'DB_Vendor', DB_TEAM: 'DB_Tim', DB_OFFICER: 'DB_Pejabat', DB_SEARCH: 'DB_Search',
    CALC_DASH: 'Calc_Dashboard', TMP_POK: 'TMP_POK'
  }
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PPBJ')
    .addItem('Buka Aplikasi Web', 'bukaAplikasiWeb')
    .addSeparator()
    .addItem('Muat Paket ke Form', 'muatPaketAktif')
    .addItem('Simpan Paket Baru', 'simpanPaketBaru')
    .addItem('Perbarui Paket Aktif', 'perbaruiPaketAktif')
    .addSeparator()
    .addItem('Pilih Rincian POK Aktif', 'pilihRincianPokAktif')
    .addItem('Tambah Alokasi Anggaran', 'tambahAlokasiAnggaran')
    .addItem('Alokasikan Harga Kontrak', 'alokasikanHargaKontrak')
    .addSeparator()
    .addItem('Preview Revisi POK', 'previewRevisiPok')
    .addItem('Aktifkan Revisi POK', 'aktifkanRevisiPok')
    .addSeparator()
    .addItem('Preview Dokumen', 'previewDokumen')
    .addItem('Buat Google Docs & PDF', 'buatGoogleDocsDanPdf')
    .addSeparator()
    .addItem('Segarkan Dashboard', 'segarkanDashboard')
    .addToUi();
}

function setupSystem() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CFG.TZ);
  segarkanDashboard();
  onOpen();
  ss.toast('Sistem aktif. Menu PPBJ sudah dipasang.', 'PPBJ', 6);
}

function sh_(name) {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active && !CFG.SPREADSHEET_ID) {
    throw new Error('SPREADSHEET_ID belum diatur pada Script Properties.');
  }
  const book = active || SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  const sheet = book.getSheetByName(name);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + name);
  return sheet;
}

function headers_(sheet, row) {
  row = row || 1;
  const lastCol = sheet.getLastColumn();
  const vals = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
  const out = {};
  vals.forEach((v, i) => { if (v) out[v.trim()] = i + 1; });
  return out;
}

function firstBlankRow_(sheet, column, startRow) {
  column = column || 1;
  startRow = startRow || 2;
  const max = Math.max(sheet.getLastRow(), startRow);
  const values = sheet.getRange(startRow, column, max - startRow + 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) if (!values[i][0]) return startRow + i;
  return max + 1;
}

function findRow_(sheet, column, value, startRow) {
  startRow = startRow || 2;
  const last = Math.max(sheet.getLastRow(), startRow);
  const vals = sheet.getRange(startRow, column, last - startRow + 1, 1).getDisplayValues().flat();
  const key = String(value).trim();
  const idx = vals.findIndex(v => String(v).trim() === key);
  return idx < 0 ? 0 : startRow + idx;
}

function rowObject_(sheet, row, headerRow) {
  headerRow = headerRow || 1;
  const map = headers_(sheet, headerRow);
  const vals = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const out = {};
  Object.keys(map).forEach(k => out[k] = vals[map[k] - 1]);
  return out;
}

function findObject_(sheetName, keyHeader, keyValue) {
  const sheet = sh_(sheetName);
  const h = headers_(sheet);
  const col = h[keyHeader];
  if (!col) throw new Error('Kolom ' + keyHeader + ' tidak ditemukan di ' + sheetName);
  const row = findRow_(sheet, col, keyValue);
  return row ? rowObject_(sheet, row) : null;
}

function getInput_() {
  const s = sh_(CFG.SHEETS.INPUT);
  const v = s.getRange('D5:D24').getValues().flat();
  return {
    id: String(v[0] || '').trim(), year: Number(v[1]), seq: Number(v[2]), name: String(v[3] || '').trim(),
    team: String(v[4] || '').trim(), teamLead: String(v[5] || '').trim(), teamNip: String(v[6] || '').trim(),
    vendor: String(v[7] || '').trim(), status: String(v[8] || '').trim(), requestDate: v[9], poDate: v[10],
    bastDate: v[11], hps: Number(v[12] || 0), offer: Number(v[13] || 0), contract: Number(v[14] || 0),
    fp: Number(v[15] || 0), mak: String(v[16] || '').trim(), procurementOfficer: String(v[17] || '').trim(),
    ppk: String(v[18] || '').trim(), note: String(v[19] || '').trim()
  };
}

function validateInput_(x, allowExisting) {
  const errors = [];
  if (!x.id || !/^PKG-\d{4}-\d{3,}$/.test(x.id)) errors.push('ID Paket tidak valid.');
  if (!x.name) errors.push('Nama pengadaan belum diisi.');
  if (!x.team || !x.teamLead || !x.teamNip) errors.push('Tim/Ketua Tim belum terhubung.');
  if (!x.vendor) errors.push('Penyedia belum dipilih.');
  if (!x.status) errors.push('Status belum dipilih.');
  if (x.contract > x.hps) errors.push('Nilai kontrak melebihi HPS.');
  if (x.status !== 'Draft' && x.fp < x.contract) errors.push('Nilai FP lebih kecil daripada kontrak.');
  if (x.bastDate && x.poDate && new Date(x.bastDate) < new Date(x.poDate)) errors.push('Tanggal BAST mendahului PO.');
  const db = sh_(CFG.SHEETS.DB_PACKAGE);
  const exists = findRow_(db, 1, x.id);
  if (!allowExisting && exists) errors.push('ID Paket sudah ada. Gunakan Perbarui Paket Aktif.');
  if (allowExisting && !exists) errors.push('ID Paket belum ada. Gunakan Simpan Paket Baru.');
  if (errors.length) throw new Error(errors.join('\n'));
}

function parseMak_(mak) {
  const m = String(mak || '').match(/^(\d{3}\.\d{2}\.[A-Z]{2})\.(\d{4})\.([A-Z]{3}\.\d{3})\.(\d{3})\.([A-Z])\.(\d{6})$/i);
  if (!m) return {program: '', activity: '', output: '', component: '', account: ''};
  return {program: m[1], activity: m[2], output: m[3].replace(/^\d{4}\./, ''), component: m[4] + '.' + m[5], account: m[6]};
}

function officerNip_(role, name) {
  const s = sh_(CFG.SHEETS.DB_OFFICER);
  const rows = s.getRange(2, 1, Math.max(s.getLastRow() - 1, 1), 4).getDisplayValues();
  const match = rows.find(r => r[0] === role && r[1] === name);
  return match ? match[2] : '';
}

function packageRowValues_(x, existing) {
  const mak = parseMak_(x.mak);
  const pjphp = existing && existing['PjPHP'] ? existing['PjPHP'] : x.ppk;
  const pjphpNip = existing && existing['NIP PjPHP'] ? existing['NIP PjPHP'] : officerNip_('PjPHP', pjphp) || officerNip_('PPK', x.ppk);
  return [
    x.id, x.year, x.seq, x.name, x.team, x.teamLead, String(x.teamNip), x.vendor, x.status,
    x.fp, x.contract, x.hps, x.offer, x.requestDate || '', x.poDate || '', x.bastDate || '', x.mak,
    mak.program, mak.activity, mak.output, mak.component, mak.account, x.procurementOfficer,
    officerNip_('Pejabat Pengadaan', x.procurementOfficer), x.ppk, officerNip_('PPK', x.ppk),
    pjphp, pjphpNip, x.note, existing && existing['Dibuat Pada'] ? existing['Dibuat Pada'] : new Date(), new Date()
  ];
}

function simpanPaketBaru() {
  const x = getInput_();
  validateInput_(x, false);
  const db = sh_(CFG.SHEETS.DB_PACKAGE);
  const row = firstBlankRow_(db, 1, 2);
  db.getRange(row, 1, 1, 31).setValues([packageRowValues_(x, null)]);
  simpanDetailAktif_(x.id, false);
  segarkanDashboard();
  SpreadsheetApp.getActive().toast(x.id + ' berhasil disimpan tanpa menimpa paket lama.', 'PPBJ', 6);
}

function perbaruiPaketAktif() {
  const x = getInput_();
  validateInput_(x, true);
  const db = sh_(CFG.SHEETS.DB_PACKAGE);
  const row = findRow_(db, 1, x.id);
  const old = rowObject_(db, row);
  db.getRange(row, 1, 1, 31).setValues([packageRowValues_(x, old)]);
  simpanDetailAktif_(x.id, true);
  sinkronkanStatusAlokasi_(x.id, x.status);
  segarkanDashboard();
  SpreadsheetApp.getActive().toast(x.id + ' berhasil diperbarui.', 'PPBJ', 5);
}

function simpanDetailAktif_(packageId, replaceExisting) {
  const src = sh_(CFG.SHEETS.DETAIL);
  const rows = src.getRange('A7:N26').getValues().filter(r => r[0] !== '' && r[1] !== '');
  const db = sh_(CFG.SHEETS.DB_DETAIL);
  if (replaceExisting) {
    const all = db.getRange(2, 1, Math.max(db.getLastRow() - 1, 1), 15).getValues();
    all.forEach((r, i) => { if (String(r[0]) === packageId) db.getRange(i + 2, 1, 1, 15).clearContent(); });
  }
  rows.forEach(r => {
    const out = [packageId, Number(r[0]), r[1], r[2], Number(r[3]), r[4], Number(r[5]), Number(r[6]),
      Number(r[7]), Number(r[8]), Number(r[9]), Number(r[10]), Number(r[11]), Number(r[12]), r[13]];
    db.getRange(firstBlankRow_(db, 1, 2), 1, 1, 15).setValues([out]);
  });
}

function muatPaketAktif() {
  const input = sh_(CFG.SHEETS.INPUT);
  const id = String(input.getRange('D5').getDisplayValue() || sh_(CFG.SHEETS.DOCS).getRange('B4').getDisplayValue()).trim();
  if (!id) throw new Error('Pilih ID paket terlebih dahulu.');
  const p = findObject_(CFG.SHEETS.DB_PACKAGE, 'ID Paket', id);
  if (!p) throw new Error('Paket tidak ditemukan: ' + id);
  input.getRange('D5:D24').setValues([
    [p['ID Paket']], [p['Tahun']], [p['Nomor Urut']], [p['Nama Pengadaan']], [p['Tim']], [p['Ketua Tim']],
    [p['NIP Ketua Tim']], [p['Penyedia']], [p['Status']], [p['Tanggal Permintaan']], [p['Tanggal PO']],
    [p['Tanggal BAST']], [p['Nilai HPS']], [p['Nilai Penawaran']], [p['Nilai Kontrak']], [p['Nilai FP']],
    [p['MAK Utama']], [p['Pejabat Pengadaan']], [p['PPK']], [p['Catatan']]
  ]);
  const detail = sh_(CFG.SHEETS.DETAIL);
  detail.getRange('B4').setValue(id);
  detail.getRange('A7:N26').clearContent();
  const dbd = sh_(CFG.SHEETS.DB_DETAIL);
  const rows = dbd.getRange(2, 1, Math.max(dbd.getLastRow() - 1, 1), 15).getValues().filter(r => String(r[0]) === id);
  const visible = rows.slice(0, 20).map(r => [r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12], r[13], r[14]]);
  if (visible.length) detail.getRange(7, 1, visible.length, 14).setValues(visible);
  sh_(CFG.SHEETS.CALC).getRange('B4').setValue(id);
  sh_(CFG.SHEETS.DOCS).getRange('B4').setValue(id);
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast(id + ' dimuat ke form.', 'PPBJ', 4);
}

function sinkronkanStatusAlokasi_(packageId, status) {
  const s = sh_(CFG.SHEETS.ALLOCATION);
  const last = Math.max(s.getLastRow(), 5);
  const vals = s.getRange(5, 1, last - 4, 10).getValues();
  vals.forEach((r, i) => { if (String(r[0]) === packageId) s.getRange(i + 5, 7).setValue(status); });
}

function pilihRincianPokAktif() {
  const s = SpreadsheetApp.getActiveSheet();
  if (s.getName() !== CFG.SHEETS.SEARCH || s.getActiveRange().getRow() < 8) throw new Error('Klik salah satu baris hasil pada sheet Cari_POK.');
  const id = String(s.getRange(s.getActiveRange().getRow(), 1).getDisplayValue()).trim();
  if (!id || !id.startsWith('POK|')) throw new Error('Baris tersebut bukan hasil POK yang valid.');
  s.getRange('B5').setValue(id);
  SpreadsheetApp.getActive().toast('Rincian POK terpilih.', 'PPBJ', 3);
}

function tambahAlokasiAnggaran() {
  const ui = SpreadsheetApp.getUi();
  const packageId = sh_(CFG.SHEETS.INPUT).getRange('D5').getDisplayValue().trim();
  const pokId = sh_(CFG.SHEETS.SEARCH).getRange('B5').getDisplayValue().trim();
  if (!packageId || !pokId) throw new Error('ID Paket atau ID Rincian POK belum dipilih.');
  const p = findObject_(CFG.SHEETS.DB_PACKAGE, 'ID Paket', packageId) || {Status: sh_(CFG.SHEETS.INPUT).getRange('D13').getDisplayValue()};
  const pok = findObject_(CFG.SHEETS.DB_POK, 'ID_Rincian', pokId);
  if (!pok) throw new Error('Rincian POK tidak ditemukan.');
  const fpPrompt = ui.prompt('Nilai FP', 'Masukkan nilai FP untuk rincian ini (tanpa titik):', ui.ButtonSet.OK_CANCEL);
  if (fpPrompt.getSelectedButton() !== ui.Button.OK) return;
  const contractPrompt = ui.prompt('Nilai Kontrak', 'Masukkan nilai kontrak untuk rincian ini (tanpa titik):', ui.ButtonSet.OK_CANCEL);
  if (contractPrompt.getSelectedButton() !== ui.Button.OK) return;
  const fp = Number(String(fpPrompt.getResponseText()).replace(/\D/g, ''));
  const contract = Number(String(contractPrompt.getResponseText()).replace(/\D/g, ''));
  if (!Number.isFinite(fp) || !Number.isFinite(contract)) throw new Error('Nilai tidak valid.');
  const s = sh_(CFG.SHEETS.ALLOCATION);
  const row = firstBlankRow_(s, 1, 5);
  s.getRange(row, 1, 1, 10).setValues([[packageId, pokId, pok['Kode POK'], pok['Uraian Rincian'], fp, contract, p['Status'] || 'Draft', '', '', '']]);
  s.getRange(row, 8).setFormula(`=IF(A${row}="","",IF(G${row}="Draft",0,IF(G${row}="FP Dibuat",E${row},IF(OR(G${row}="Kontrak",G${row}="Selesai"),F${row},0))))`);
  s.getRange(row, 9).setFormula(`=IF(B${row}="","",IFERROR(INDEX('DB_POK'!$T$2:$T$1000,MATCH(B${row},'DB_POK'!$A$2:$A$1000,0)),""))`);
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast('Alokasi ditambahkan.', 'PPBJ', 3);
}

function alokasikanHargaKontrak() {
  const detail = sh_(CFG.SHEETS.DETAIL);
  const target = Number(sh_(CFG.SHEETS.CALC).getRange('D27').getValue());
  if (!(target > 0)) throw new Error('Nilai kontrak pada Kalkulator_Harga belum valid.');
  const rows = detail.getRange('A7:N26').getValues();
  const items = [];
  rows.forEach((r, i) => {
    const q = Number(r[3]);
    const net = Number(r[5]);
    if (r[0] !== '' && q > 0 && net >= 0) items.push({row: i + 7, qty: q, net: net});
  });
  if (!items.length) throw new Error('Belum ada rincian barang.');
  const netTotal = items.reduce((s, x) => s + x.qty * x.net, 0);
  if (!(netTotal > 0)) throw new Error('Total harga bersih harus lebih dari nol.');
  const factor = target / netTotal;
  items.forEach(x => x.gross = Math.round(x.net * factor));
  let current = items.reduce((s, x) => s + x.qty * x.gross, 0);
  let delta = Math.round(target - current);
  if (delta !== 0) {
    const result = coinAdjust_(items.map(x => Math.round(x.qty)), Math.abs(delta));
    if (result) {
      result.forEach((count, i) => items[i].gross += (delta > 0 ? count : -count));
      delta = 0;
    }
  }
  items.forEach(x => detail.getRange(x.row, 12).setValue(x.gross));
  SpreadsheetApp.flush();
  const finalTotal = Number(detail.getRange('M28').getValue());
  if (finalTotal !== target) {
    SpreadsheetApp.getUi().alert('Alokasi selesai dengan selisih Rp' + formatNumber_(target - finalTotal) + '. Tambahkan baris “Selisih Pembulatan” volume 1 agar total tepat.');
  } else {
    SpreadsheetApp.getActive().toast('Harga kontrak berhasil dialokasikan tepat Rp' + formatNumber_(target) + '.', 'PPBJ', 5);
  }
}

function coinAdjust_(coins, target) {
  if (target === 0) return Array(coins.length).fill(0);
  if (target > 100000) return null;
  const prev = Array(target + 1).fill(-1);
  const used = Array(target + 1).fill(-1);
  prev[0] = 0;
  for (let v = 1; v <= target; v++) {
    for (let i = 0; i < coins.length; i++) {
      const c = coins[i];
      if (c <= v && prev[v - c] !== -1) { prev[v] = v - c; used[v] = i; break; }
    }
  }
  if (prev[target] === -1) return null;
  const counts = Array(coins.length).fill(0);
  for (let v = target; v > 0; v = prev[v]) counts[used[v]]++;
  return counts;
}

function segarkanDashboard() {
  const year = Number(sh_(CFG.SHEETS.HOME).getRange('B4').getValue());
  const db = sh_(CFG.SHEETS.DB_PACKAGE);
  const rows = db.getRange(2, 1, Math.max(db.getLastRow() - 1, 1), 31).getValues().filter(r => r[0] && Number(r[1]) === year);
  const vendor = countShare_(rows.map(r => r[7]));
  const team = countShare_(rows.map(r => r[4]));
  const calc = sh_(CFG.SHEETS.CALC_DASH);
  calc.getRange('K2:L6').clearContent();
  calc.getRange('N2:O6').clearContent();
  if (vendor.length) calc.getRange(2, 11, Math.min(5, vendor.length), 2).setValues(vendor.slice(0, 5));
  if (team.length) calc.getRange(2, 14, Math.min(5, team.length), 2).setValues(team.slice(0, 5));
  SpreadsheetApp.flush();
}

function countShare_(values) {
  const map = {};
  values.filter(Boolean).forEach(v => map[v] = (map[v] || 0) + 1);
  const total = Object.values(map).reduce((a, b) => a + b, 0) || 1;
  return Object.keys(map).map(k => [k, map[k] / total]).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
}

function previewRevisiPok() {
  const ctl = sh_(CFG.SHEETS.UPDATE);
  const url = ctl.getRange('B4').getDisplayValue().trim();
  if (!url) throw new Error('Tautan matriks terbaru belum diisi.');
  const source = SpreadsheetApp.openByUrl(url);
  const candidates = source.getSheets().map(s => {
    const m = s.getName().match(/^POK[- ]?(\d+)$/i);
    return m ? {sheet: s, num: Number(m[1])} : null;
  }).filter(Boolean).sort((a, b) => b.num - a.num);
  if (!candidates.length) throw new Error('Tidak ditemukan sheet bernama POK-nomor.');
  const newest = candidates[0];
  const parsed = parsePokSheet_(newest.sheet, newest.sheet.getName());
  const tmpName = CFG.SHEETS.TMP_POK;
  const ss = SpreadsheetApp.getActive();
  let tmp = ss.getSheetByName(tmpName);
  if (!tmp) tmp = ss.insertSheet(tmpName);
  tmp.clear();
  tmp.getRange(1, 1, parsed.length + 1, 18).setValues([['ID_Rincian','Revisi','Aktif','Baris Sumber','Program','Kegiatan','KRO','RO','Komponen','Subkomponen','Akun','Kode POK','MAK Lengkap','Uraian Rincian','Volume','Satuan','Harga Satuan','Pagu Aktif'], ...parsed]);
  tmp.hideSheet();

  const currentSheet = sh_(CFG.SHEETS.DB_POK);
  const current = currentSheet.getRange(2, 1, Math.max(currentSheet.getLastRow() - 1, 1), 22).getValues().filter(r => r[0]);
  const oldMap = new Map(current.map(r => [String(r[0]), r]));
  const newMap = new Map(parsed.map(r => [String(r[0]), r]));
  const changes = [];
  parsed.forEach(r => {
    const old = oldMap.get(String(r[0]));
    if (!old) changes.push(['Baru', r[0], r[11], r[13], 0, r[17], r[17], '']);
    else if (Number(old[17]) !== Number(r[17])) {
      const d = Number(r[17]) - Number(old[17]);
      changes.push([d > 0 ? 'Naik' : 'Turun', r[0], r[11], r[13], old[17], r[17], d, '']);
    }
  });
  current.forEach(r => { if (!newMap.has(String(r[0]))) changes.push(['Dihapus', r[0], r[11], r[13], r[17], 0, -Number(r[17]), '']); });
  ctl.getRange('A13:H1000').clearContent();
  if (changes.length) ctl.getRange(13, 1, changes.length, 8).setValues(changes);
  ctl.getRange('B6').setValue(newest.sheet.getName());
  ctl.getRange('B9').setValue(changes.length + ' perubahan siap ditinjau');
  SpreadsheetApp.getActive().toast('Preview ' + newest.sheet.getName() + ' selesai: ' + changes.length + ' perubahan.', 'PPBJ', 6);
}

function parsePokSheet_(sheet, revision) {
  const values = sheet.getDataRange().getValues();
  const c = {program:'', activity:'', kro:'', ro:'', component:'', sub:'', account:''};
  const raw = [];
  for (let i = 4; i < values.length; i++) {
    const row = values[i];
    const code = String(row[0] == null ? '' : row[0]).trim();
    const dash = String(row[2] == null ? '' : row[2]).trim();
    const desc = String(row[3] == null ? '' : row[3]).trim();
    if (/^\d{3}\.\d{2}\.[A-Z]{2}$/i.test(code)) c.program = code;
    else if (/^\d{4}$/.test(code)) c.activity = code;
    else if (/^\d{4}\.[A-Z]{3}$/i.test(code)) c.kro = code;
    else if (/^\d{4}\.[A-Z]{3}\.\d{3}$/i.test(code)) c.ro = code;
    else if (/^\d{3}$/.test(code)) c.component = code;
    else if (/^[A-Z]$/i.test(code)) c.sub = code;
    else if (/^\d{6}$/.test(code)) c.account = code;
    if (dash === '-' && desc) {
      const pok = [c.ro, c.component, c.sub, c.account].filter(Boolean).join('.');
      const mak = [c.program, c.ro, c.component, c.sub, c.account].filter(Boolean).join('.');
      raw.push({row:i+1, program:c.program, activity:c.activity, kro:c.kro, ro:c.ro, component:c.component, sub:c.sub, account:c.account, pok:pok, mak:mak, desc:desc, volume:Number(row[8]||0), unit:String(row[9]||''), unitPrice:Number(row[10]||0), budget:Number(row[11]||0)});
    }
  }
  const counts = {};
  return raw.map(x => {
    const base = x.pok + '|' + slug_(x.desc);
    counts[base] = (counts[base] || 0) + 1;
    const id = 'POK|' + base + '|' + counts[base];
    return [id, revision, 'Ya', x.row, x.program, x.activity, x.kro, x.ro, x.component, x.sub, x.account, x.pok, x.mak, x.desc, x.volume, x.unit, x.unitPrice, x.budget];
  });
}

function aktifkanRevisiPok() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Aktifkan revisi POK?', 'Revisi aktif lama akan dipindahkan ke Arsip_POK dan tidak dihapus.', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActive();
  const tmp = ss.getSheetByName(CFG.SHEETS.TMP_POK);
  if (!tmp || tmp.getLastRow() < 2) throw new Error('Jalankan Preview Revisi POK terlebih dahulu.');
  const fresh = tmp.getRange(2, 1, tmp.getLastRow() - 1, 18).getValues().filter(r => r[0]);
  const db = sh_(CFG.SHEETS.DB_POK);
  const old = db.getRange(2, 1, Math.max(db.getLastRow() - 1, 1), 22).getValues().filter(r => r[0]);
  const archive = sh_(CFG.SHEETS.ARCHIVE_POK);
  if (old.length) {
    const start = firstBlankRow_(archive, 1, 2);
    archive.getRange(start, 1, old.length, 23).setValues(old.map(r => [...r, new Date()]));
  }
  if (db.getLastRow() > 1) db.getRange(2, 1, db.getLastRow() - 1, 22).clearContent();
  if (fresh.length) db.getRange(2, 1, fresh.length, 18).setValues(fresh);
  fresh.forEach((r, i) => {
    const row = i + 2;
    db.getRange(row, 19).setFormula(`=SUMIF('Alokasi_Anggaran'!$B$2:$B$500,A${row},'Alokasi_Anggaran'!$H$2:$H$500)`);
    db.getRange(row, 20).setFormula(`=R${row}-S${row}`);
    db.getRange(row, 21).setFormula(`=IF(T${row}<0,"Tidak Cukup",IF(T${row}=0,"Habis","Tersedia"))`);
    db.getRange(row, 22).setValue(searchTerms_(r[11], r[12], r[13]));
  });
  rebuildDbSearch_();
  const hist = sh_(CFG.SHEETS.HISTORY_POK);
  if (hist.getLastRow() > 1) hist.getRange(2, 2, hist.getLastRow() - 1, 1).setValue('Tidak');
  const revision = String(fresh[0][1]);
  const total = fresh.reduce((s, r) => s + Number(r[17] || 0), 0);
  const hrow = firstBlankRow_(hist, 1, 2);
  hist.getRange(hrow, 1, 1, 8).setValues([[revision, 'Ya', new Date(), sh_(CFG.SHEETS.UPDATE).getRange('B4').getDisplayValue(), fresh.length, fresh.filter(r => Number(r[17]) > 0).length, total, sh_(CFG.SHEETS.UPDATE).getRange('B8').getDisplayValue()]]);
  sh_(CFG.SHEETS.UPDATE).getRange('B5').setValue(revision);
  sh_(CFG.SHEETS.UPDATE).getRange('B9').setValue('Revisi aktif: ' + revision);
  tmp.clear();
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast(revision + ' aktif. Revisi lama tersimpan di Arsip_POK.', 'PPBJ', 7);
}

function rebuildDbSearch_() {
  const db = sh_(CFG.SHEETS.DB_POK);
  const out = sh_(CFG.SHEETS.DB_SEARCH);
  const rows = db.getRange(2, 1, Math.max(db.getLastRow() - 1, 1), 22).getValues().filter(r => r[0]);
  if (out.getLastRow() > 1) out.getRange(2, 1, out.getLastRow() - 1, 8).clearContent();
  const values = rows.map(r => [r[0], r[11], r[13], r[17], r[18], r[19], r[20], r[21]]);
  if (values.length) out.getRange(2, 1, values.length, 8).setValues(values);
}

function previewDokumen() {
  const id = sh_(CFG.SHEETS.DOCS).getRange('B4').getDisplayValue().trim();
  const p = findObject_(CFG.SHEETS.DB_PACKAGE, 'ID Paket', id);
  if (!p) throw new Error('Paket tidak ditemukan: ' + id);
  sh_(CFG.SHEETS.CALC).getRange('B4').setValue(id);
  SpreadsheetApp.flush();
  sh_(CFG.SHEETS.DOCS).activate();
  SpreadsheetApp.getActive().toast('Preview ' + id + ' diperbarui. Periksa panel kanan.', 'PPBJ', 5);
}

function buatGoogleDocsDanPdf() {
  const ctl = sh_(CFG.SHEETS.DOCS);
  const id = ctl.getRange('B4').getDisplayValue().trim();
  const templateId = extractId_(ctl.getRange('B5').getDisplayValue());
  const folderId = extractId_(ctl.getRange('B6').getDisplayValue());
  const version = Number(ctl.getRange('B7').getValue() || 1);
  if (!id) throw new Error('ID Paket belum dipilih.');
  if (!templateId || /ISI_/.test(ctl.getRange('B5').getDisplayValue())) throw new Error('Template Google Docs ID belum diisi.');
  if (!folderId || /ISI_/.test(ctl.getRange('B6').getDisplayValue())) throw new Error('Folder Arsip Drive ID belum diisi.');
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
  archive.getRange(row, 1, 1, 10).setValues([[
    'ARS-' + id.replace('PKG-', '') + '-v' + version, id, version, 'Dokumen Pengadaan Gabungan', new Date(),
    Session.getActiveUser().getEmail() || 'Pengguna', copy.getUrl(), pdf.getUrl(), 'Selesai', 'Dibuat otomatis'
  ]]);
  ctl.getRange('B7').setValue(version + 1);
  ctl.getRange('B8').setValue('Selesai: ' + title);
  SpreadsheetApp.getUi().alert('Dokumen selesai', 'Google Docs:\n' + copy.getUrl() + '\n\nPDF:\n' + pdf.getUrl(), SpreadsheetApp.getUi().ButtonSet.OK);
}

function collectMergeData_(id) {
  const p = findObject_(CFG.SHEETS.DB_PACKAGE, 'ID Paket', id);
  if (!p) throw new Error('Paket tidak ditemukan.');
  const vendor = findObject_(CFG.SHEETS.DB_VENDOR, 'Nama Usaha', p['Penyedia']) || {};
  const team = findObject_(CFG.SHEETS.DB_TEAM, 'Tim', p['Tim']) || {};
  const letters = {};
  const sl = sh_(CFG.SHEETS.DB_LETTER);
  sl.getRange(2, 1, Math.max(sl.getLastRow() - 1, 1), 6).getValues().forEach(r => { if (String(r[0]) === id) letters[String(r[1])] = {date:r[2], number:r[3]}; });
  const detailSheet = sh_(CFG.SHEETS.DB_DETAIL);
  const details = detailSheet.getRange(2, 1, Math.max(detailSheet.getLastRow() - 1, 1), 15).getValues().filter(r => String(r[0]) === id);
  const hps = Number(p['Nilai HPS'] || 0), nego = Number(p['Nilai Kontrak'] || 0);
  const po = p['Tanggal PO'], bast = p['Tanggal BAST'];
  const map = {
    ID_Paket: id, NO: p['Nomor Urut'], Nama_Kegiatan_Pengadaan: p['Nama Pengadaan'],
    NAMA_KEGIATAN_PENGADAAN: p['Nama Pengadaan'], Jenis_Pekerjaan: 'Barang/Jasa Lainnya',
    TANGGAL_PEMBUATAN_DOKUMEN: formatDate_(p['Tanggal Permintaan']), Nomor_Form: String(p['Nomor Urut']).padStart(3, '0'),
    Tanggal_Form: formatDate_(p['Tanggal Permintaan']), Kepala_Unit_Eselon_II: 'Kepala ' + CFG.ORGANIZATION_NAME,
    Sub_DirektoratBagian: p['Tim'], Tahun_Anggaran: p['Tahun'], DIPA: CFG.DIPA_REFERENCE,
    Program: '(' + p['Program'] + ') Program Organisasi',
    Kode_Kegiatan: '(' + p['Kegiatan'] + ') Kegiatan Organisasi', Output: '(' + p['Output'] + ') Output Kegiatan',
    Komponen: '(' + p['Komponen'] + ') Komponen', Akun: '(' + p['Akun'] + ') Belanja', Mata_Anggaran: p['MAK Utama'],
    Nilai_Anggaran_yang_Akan_Digunakan_Rp: formatNumber_(p['Nilai FP']), Nilai_Pagu_Anggaran_Rp: formatNumber_(p['Nilai FP']),
    Nama_PPK: p['PPK'], NIP_PPK: p['NIP PPK'], Label_PPK: 'Pejabat Pembuat Komitmen', Nama_Pejabat_Pengadaan: p['Pejabat Pengadaan'],
    Label_Pejabat_Pengadaan: 'Pejabat Pengadaan', NIP_Pejabat: p['NIP Pejabat'], Alamat: CFG.ORGANIZATION_ADDRESS,
    Nilai_HPS: formatCurrency_(hps), Nilai_HPS_Terbilang: titleCase_(terbilang_(hps)) + ' Rupiah', Nilai_Penawaran: formatNumber_(p['Nilai Penawaran']),
    Nilai_Nego: formatCurrency_(nego), Nilai_Terbilang_Nego: titleCase_(terbilang_(nego)) + ' Rupiah',
    Nama_Usaha: p['Penyedia'], Nama_Penyedia: vendor['Nama Penanggung Jawab'] || p['Penyedia'], Alamat_Penyedia: vendor['Alamat'] || '',
    Kota_Penyedia: vendor['Kota'] || '', Label_Pimpinan_Penyedia: vendor['Jabatan'] || 'Pimpinan', Nama_Pimpinan_Penyedia: vendor['Nama Penanggung Jawab'] || '',
    Nama_Bank: vendor['Bank'] || '', Nomor_Rekening: vendor['No. Rekening'] || '', Atas_Nama_Rekening: vendor['a.n. Rekening'] || '',
    SM_penandatangan_BAPP: team['Ketua Tim'] || p['Ketua Tim'], Jabatan_SM_Penandatangan_BAPP: team['Jabatan'] || '', NIP_SM_Penandatangan_BAPP: team['NIP'] || p['NIP Ketua Tim'],
    Nama_PjPHP: p['PjPHP'], NIP_PjPHP: p['NIP PjPHP'], Label_PjPHP: 'Pejabat Pemeriksa/Penerima Hasil Pekerjaan',
    Jangka_waktu_pelaksanaan_pekerjaan_dalam: dayDiff_(po, bast), Tanggal_Batas_Akhir_Pekerjaan: formatDate_(bast),
    Nomor_Penetapan_HPS: letterNo_(letters, 'Penetapan HPS'), Tanggal_Penetapan_HPS: formatDate_(letterDate_(letters, 'Penetapan HPS')),
    Nomor_Surat_Permintaan_Pengadaan: letterNo_(letters, 'Permintaan Pengadaan'), Tanggal_Surat_Permintaan_Pengadaan: formatDate_(letterDate_(letters, 'Permintaan Pengadaan')),
    Nomor_Surat_Pesanan: letterNo_(letters, 'Surat Pesanan'), Tanggal_Surat_Pesanan: formatDate_(letterDate_(letters, 'Surat Pesanan')),
    Nomor_BAPPBAKP: letterNo_(letters, 'Pemeriksaan Pekerjaan'), Tanggal_BAPPBAKP: formatDate_(letterDate_(letters, 'Pemeriksaan Pekerjaan')),
    Nomor_BAST: letterNo_(letters, 'BAST'), Tanggal_BAST: formatDate_(letterDate_(letters, 'BAST')),
    Nomor_BA_Pemeriksaan_Administrasi: letterNo_(letters, 'Pemeriksaan Administrasi'), Tanggal_BA_Pemeriksaan_Administrasi: formatDate_(letterDate_(letters, 'Pemeriksaan Administrasi')),
    Nomor_BAP: letterNo_(letters, 'Pembayaran'), Tanggal_BAP: formatDate_(letterDate_(letters, 'Pembayaran')),
    Nomor_Kuitansi: letterNo_(letters, 'Kuitansi'), Tanggal_Kuitansi: formatDate_(letterDate_(letters, 'Kuitansi')),
    Tanggal_Invoice: formatDate_(bast), Nomor_Invoice: '', Nomor_Surat_Penawaran: '', Tanggal_Surat_Penawaran: formatDate_(po), Nomor_BA_Nego: '',
    Tanggal_Penetapan_Spesifikasi_Teknis: formatDate_(po), Tgl_SK_PPK: '', No_SK_PPK: '', Nomor_SK_PjPHP: ''
  };
  addDateFields_(map, 'HPS', letterDate_(letters, 'Penetapan HPS'));
  addDateFields_(map, 'BA_Nego', po);
  addDateFields_(map, 'BAPPBAKP', letterDate_(letters, 'Pemeriksaan Pekerjaan'));
  addDateFields_(map, 'BAST', letterDate_(letters, 'BAST'));
  addDateFields_(map, 'BA_Pemeriksaan_Administrasi', letterDate_(letters, 'Pemeriksaan Administrasi'));
  addDateFields_(map, 'BAP', letterDate_(letters, 'Pembayaran'));
  map.Tanggal_BAPPBAKP_Terbilang = map.Tanggal_Terbilang_BAPPBAKP || '';
  map.Bulan_BAPP = map.Bulan_BAPPBAKP || '';
  map.Bulan_BA_Pemeriksaan_Administrasi = map.Bulan_Pemeriksaan_Administrasi || '';
  map.Tanggal_Terbilang_BA_Pemeriksaan_Adminis = map.Tanggal_Terbilang_BA_Pemeriksaan_Administrasi || '';
  map.Tanggal_BAP_Terbilang = map.Tanggal_Terbilang_BAP || '';
  map.Tanggal_Kuitansi_Terbilang = dateWords_(letterDate_(letters, 'Kuitansi'));
  return {map: map, details: details};
}

function addDateFields_(map, prefix, date) {
  const d = validDate_(date) ? new Date(date) : null;
  const dd = d ? d.getDate() : '';
  const month = d ? monthName_(d.getMonth()) : '';
  const year = d ? d.getFullYear() : '';
  map['Tanggal_Terbilang_' + prefix] = dd === '' ? '' : terbilang_(dd);
  map['Bulan_Terbilang_' + prefix] = month;
  map['Tahun_Terbilang_' + prefix] = year === '' ? '' : terbilang_(year);
  map['Gabungan_Terbilang_' + prefix] = dateWords_(d);
  map['Hari_' + prefix] = d ? dayName_(d.getDay()) : '';
  map['Bulan_' + prefix.replace('BA_', '')] = month;
}

function insertDetailTable_(body, rows) {
  const found = body.findText('\\{\\{TABEL_DETAIL\\}\\}');
  if (!found) return;
  const text = found.getElement().asText();
  const paragraph = text.getParent().asParagraph();
  const previous = paragraph.getPreviousSibling();
  let table;

  // Template terbaru memiliki tabel header native dengan baris pertama dipin.
  // Memakai tabel tersebut membuat header berulang otomatis di halaman lanjutan.
  if (previous && previous.getType() === DocumentApp.ElementType.TABLE) {
    table = previous.asTable();
    while (table.getNumRows() > 1) table.removeRow(table.getNumRows() - 1);
  } else {
    const index = body.getChildIndex(paragraph);
    table = body.insertTable(index, [['No', 'Uraian/Spesifikasi', 'Volume', 'Satuan', 'Harga Satuan', 'Jumlah']]);
  }

  // Remove the placeholder as an element. DocumentApp rejects setText('')
  // because it attempts to insert an empty text element.
  body.removeChild(paragraph);

  rows.forEach(r => {
    const row = table.appendTableRow();
    [docText_(r[1]), docText_(r[2]) + (r[3] ? '\n' + docText_(r[3]) : ''), docText_(r[4]), docText_(r[5]),
      formatCurrency_(r[12]), formatCurrency_(r[13])].forEach(v => appendSafeTableCell_(row, v));
  });
  const totalRow = table.appendTableRow();
  ['TOTAL', '', '', '', '', formatCurrency_(rows.reduce((s, r) => s + Number(r[13] || 0), 0))]
    .forEach(v => appendSafeTableCell_(totalRow, v));
  // Merge columns 1–5 from right to left. This works whether the runtime keeps
  // covered-cell indexes or removes merged cells from the row immediately.
  mergeLeadingTableCells_(totalRow, 5);

  for (let i = 0; i < table.getNumRows(); i++) {
    const row = table.getRow(i);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      const cellText = cell.editAsText();
      // Appended rows can inherit the header style, so explicitly keep only
      // the header bold and reset every detail/TOTAL cell to regular weight.
      cellText.setFontFamily('Arial').setFontSize(8).setBold(i === 0);

      // Vertically center every cell. Center the headers, utility columns,
      // and TOTAL label; keep descriptions left-aligned and money right-aligned.
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      const isHeader = i === 0;
      const isTotal = i === table.getNumRows() - 1;
      const horizontal = isHeader || (isTotal && c === 0) || (!isTotal && (c === 0 || c === 2 || c === 3))
        ? DocumentApp.HorizontalAlignment.CENTER
        : (!isTotal && c === 1 ? DocumentApp.HorizontalAlignment.LEFT : DocumentApp.HorizontalAlignment.RIGHT);
      setCellHorizontalAlignment_(cell, horizontal);
    }
  }
}

function mergeLeadingTableCells_(row, count) {
  if (count < 2) return row.getCell(0);
  for (let c = count - 1; c >= 1; c--) row.getCell(c).merge();
  const merged = row.getCell(0);
  if (typeof merged.getColSpan === 'function') {
    const span = merged.getColSpan();
    if (span > 0 && span !== count) throw new Error('The TOTAL row could not be merged across ' + count + ' columns.');
  }
  return merged;
}

function setCellHorizontalAlignment_(cell, alignment) {
  for (let i = 0; i < cell.getNumChildren(); i++) {
    const child = cell.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) child.asParagraph().setAlignment(alignment);
    if (child.getType() === DocumentApp.ElementType.LIST_ITEM) child.asListItem().setAlignment(alignment);
  }
}

function docText_(value) {
  return value == null ? '' : String(value);
}

function appendSafeTableCell_(row, value) {
  const text = docText_(value);
  return text === '' ? row.appendTableCell() : row.appendTableCell(text);
}

function letterNo_(letters, key) { return letters[key] ? letters[key].number || '' : ''; }
function letterDate_(letters, key) { return letters[key] ? letters[key].date || '' : ''; }
function dayDiff_(a, b) { return validDate_(a) && validDate_(b) ? Math.max(1, Math.round((new Date(b) - new Date(a)) / 86400000) + 1) + ' hari kalender' : ''; }
function extractId_(s) { const m = String(s || '').match(/[-\w]{25,}/); return m ? m[0] : ''; }
function escapeRegex_(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function slug_(s) { return String(s || '').toLowerCase().normalize('NFKD').replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 90); }
function searchTerms_() {
  const base = Array.prototype.slice.call(arguments).join(' ').toLowerCase();
  return base.trim();
}
function validDate_(d) { return d && !isNaN(new Date(d).getTime()); }
function formatDate_(d) { return validDate_(d) ? Utilities.formatDate(new Date(d), CFG.TZ, 'd MMMM yyyy') : ''; }
function formatNumber_(n) { return String(Math.round(Number(n || 0))).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function formatCurrency_(n) { return 'Rp' + formatNumber_(n) + ',-'; }
function dayName_(i) { return ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'][i] || ''; }
function monthName_(i) { return ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'][i] || ''; }
function dateWords_(d) { return validDate_(d) ? terbilang_(new Date(d).getDate()) + ' ' + monthName_(new Date(d).getMonth()).toLowerCase() + ' ' + terbilang_(new Date(d).getFullYear()) : ''; }
function titleCase_(s) { return String(s || '').toLowerCase().replace(/(^|\s)\S/g, x => x.toUpperCase()); }

function terbilang_(value) {
  const n = Math.floor(Math.abs(Number(value || 0)));
  const words = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan', 'sepuluh', 'sebelas'];
  function say(x) {
    if (x < 12) return words[x];
    if (x < 20) return say(x - 10) + ' belas';
    if (x < 100) return say(Math.floor(x / 10)) + ' puluh' + (x % 10 ? ' ' + say(x % 10) : '');
    if (x < 200) return 'seratus' + (x > 100 ? ' ' + say(x - 100) : '');
    if (x < 1000) return say(Math.floor(x / 100)) + ' ratus' + (x % 100 ? ' ' + say(x % 100) : '');
    if (x < 2000) return 'seribu' + (x > 1000 ? ' ' + say(x - 1000) : '');
    if (x < 1000000) return say(Math.floor(x / 1000)) + ' ribu' + (x % 1000 ? ' ' + say(x % 1000) : '');
    if (x < 1000000000) return say(Math.floor(x / 1000000)) + ' juta' + (x % 1000000 ? ' ' + say(x % 1000000) : '');
    if (x < 1000000000000) return say(Math.floor(x / 1000000000)) + ' miliar' + (x % 1000000000 ? ' ' + say(x % 1000000000) : '');
    return say(Math.floor(x / 1000000000000)) + ' triliun' + (x % 1000000000000 ? ' ' + say(x % 1000000000000) : '');
  }
  return (Number(value) < 0 ? 'minus ' : '') + (say(n) || 'nol');
}
