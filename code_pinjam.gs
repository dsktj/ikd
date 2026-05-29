/**
 * ============================================================
 *  INVENTORY KENDARAAN DINAS — FORM PEMINJAMAN
 *  Direktorat Sarana dan Keselamatan Transportasi Jalan
 *  code_pinjam.gs  |  Google Apps Script
 * ============================================================
 *
 *  SETUP:
 *  1. Buat Google Spreadsheet baru, salin Spreadsheet ID-nya ke SPREADSHEET_ID.
 *  2. Di Apps Script Project ini, buat file HTML baru bernama "pinjam"
 *     (File → New → HTML file), lalu paste isi pinjam.html ke dalamnya.
 *  3. Deploy → New Deployment → Web App
 *       Execute as : Me
 *       Who has access : Anyone   (atau Anyone within ... jika internal)
 *  4. Buka URL Web App → form langsung tampil.
 *
 *  CATATAN:
 *  - Tidak perlu GAS_URL di HTML karena pakai google.script.run (bukan fetch).
 *  - Sheet "Kendaraan" harus ada dengan header: Plat | Merk | Tipe | Model | Jenis
 * ============================================================
 */

// ── Konfigurasi ──────────────────────────────────────────────
// SPREADSHEET_ID diambil dari Script Properties (diset via setupSystem() di Code.gs).
// Gunakan getConfig() dari Code.gs jika tersedia, fallback ke Script Properties langsung.
function getSpreadsheetId() {
  try {
    // getConfig() didefinisikan di Code.gs — akan throw jika key tidak ada
    return getConfig('SPREADSHEET_ID');
  } catch(e) {
    // Fallback: baca langsung dari Script Properties
    var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!id) throw new Error('SPREADSHEET_ID belum diset. Jalankan setupSystem() dari GAS Editor.');
    return id;
  }
}

const SHEET_PEMINJAMAN = 'Peminjaman';
const SHEET_KENDARAAN  = 'Kendaraan';

const HEADERS_PEMINJAMAN = [
  'No', 'Timestamp', 'Plat Nomor', 'Merk', 'Tipe', 'Model', 'Jenis',
  'Tgl. Keluar', 'Jam Keluar', 'Tgl. Masuk', 'Jam Masuk',
  'Nama Pengguna', 'NIP Pengguna', 'No. HP Pengguna', 
  'Keperluan', 'Lokasi Tujuan', 'Kendala', 'Status',
];

// ── Routing via doGet() di Code.gs — entry point terpusat ──
// page=pinjam → pinjam.html (form publik)
// page=login  → Index.html (login)
// page=dashboard → Dashboard.html (admin)
// ── Fungsi dipanggil dari HTML via google.script.run ─────────

/**
 * Mengambil daftar kendaraan dari sheet Kendaraan.
 * Jika Status="Tersedia" → aktif, selain itu → disabled.
 * Dipanggil: google.script.run.withSuccessHandler(fn).getKendaraan()
 * @returns {Array<{plat,merk,tipe,model,jenis,dipinjam}>}
 */
function getKendaraan() {
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheetK = ss.getSheetByName(SHEET_KENDARAAN);
    if (!sheetK) { Logger.log('Sheet Kendaraan tidak ditemukan'); return []; }

    const rows = sheetK.getDataRange().getValues();
    if (rows.length < 2) return [];

    const headers = rows[0].map(function(h){ return String(h).trim().toLowerCase(); });
    function findCol(names, fallback) {
      var idx = -1;
      names.forEach(function(n){ if (idx < 0) idx = headers.indexOf(n); });
      return idx >= 0 ? idx : fallback;
    }
    const iP = findCol(['plat nomor','plat_nomor','plat'], 0);
    const iM = findCol(['merk'], 1);
    const iT = findCol(['tipe'], 2);
    const iMd = findCol(['model'], 3);
    const iJ = findCol(['jenis'], 4);
    const iS = findCol(['status'], -1);

    const result = [];
    for (var i = 1; i < rows.length; i++) {
      const r = rows[i];
      const plat = String(r[iP] || '').trim();
      if (!plat) continue;
      const status = iS >= 0 ? String(r[iS] || '').trim().toLowerCase() : '';
      result.push({
        plat:     plat,
        merk:     String(r[iM] || '').trim(),
        tipe:     String(r[iT] || '').trim(),
        model:    String(r[iMd] || '').trim(),
        jenis:    String(r[iJ] || '').trim(),
        dipinjam: status !== 'tersedia'
      });
    }
    return result;

  } catch(e) {
    Logger.log('getKendaraan ERROR: ' + e.message);
    return [];
  }
}

/**
 * Mengambil logo dari Google Drive sebagai base64 data URI.
 * Dipanggil: google.script.run.withSuccessHandler(fn).getLogos()
 * Fungsi ini juga ada di logo.gs — disalin ke sini agar scope
 * DriveApp ter-authorize bersama SpreadsheetApp saat deploy.
 * @returns {{ kemenhub?: string, hubdat?: string, sktj?: string }}
 */
function getLogos(token) {
  var LOGO_FILES = { kemenhub: 'kemenhub.png', hubdat: 'hubdat.png', sktj: 'sktj.png' };
  var result = {};
  try {
    var folderId = PropertiesService.getScriptProperties().getProperty('LOGO_FOLDER_ID') || '1K3NjyzSmvTzNnjSZjDlp2Y1Xz2rMH997';
    var folder = DriveApp.getFolderById(folderId);
    for (var key in LOGO_FILES) {
      try {
        var found = folder.getFilesByName(LOGO_FILES[key]);
        if (!found.hasNext()) continue;
        var file = found.next();
        var blob = file.getBlob();
        result[key] = 'data:' + blob.getContentType() + ';base64,'
                    + Utilities.base64Encode(blob.getBytes());
      } catch(ef) {
        Logger.log('[getLogos] Gagal baca ' + LOGO_FILES[key] + ': ' + ef.message);
      }
    }
  } catch(e) {
    Logger.log('[getLogos] Gagal buka folder: ' + e.message);
  }
  return result;
}

/**
 * Menyimpan data peminjaman ke sheet Peminjaman.
 * Dipanggil: google.script.run.withSuccessHandler(fn).simpanPeminjaman(data)
 * @param {Object} data
 * @returns {{ status: 'ok'|'error', noUrut: number, message?: string }}
 */
function ensureHeadersPeminjaman() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  let sheet = ss.getSheetByName(SHEET_PEMINJAMAN);
  if (!sheet) return;
  const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(h){ return String(h).trim().toLowerCase(); });
  var addCols = [];
  HEADERS_PEMINJAMAN.forEach(function(h){
    var key = h.trim().toLowerCase();
    if (existing.indexOf(key) < 0) addCols.push(h);
  });
  if (!addCols.length) return;
  var oldLen = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, oldLen).getValues()[0];
  headerRow = headerRow.concat(addCols);
  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  formatHeaderRow(sheet);
  sheet.autoResizeColumns(1, headerRow.length);
}

function simpanPeminjaman(data) {
  try {
    const ss    = SpreadsheetApp.openById(getSpreadsheetId());
    ensureHeadersPeminjaman();
    const sheet = getOrCreateSheet(ss, SHEET_PEMINJAMAN, HEADERS_PEMINJAMAN);

    // Baca header aktual dari sheet baris 1
    const totalCol  = sheet.getLastColumn();
    const headerRow = sheet.getRange(1, 1, 1, totalCol).getValues()[0];

    // Bangun headerMap: nama_lower → index (0-based)
    const headerMap = {};
    headerRow.forEach(function(h, i){
      var key = String(h).trim().toLowerCase();
      if (key) headerMap[key] = i;
    });

    // Tentukan lebar row = kolom 'status' + 1 (bukan lastCol yg bisa termasuk kolom kosong berformat)
    var statusIdx = headerMap['status'];
    if (statusIdx === undefined) statusIdx = totalCol - 1;
    var rowWidth = statusIdx + 1;

    function set(row, name, val) {
      var i = headerMap[name.toLowerCase()];
      if (i !== undefined && i < rowWidth) row[i] = sanitizeCell(val);
    }

    const lastRow = sheet.getLastRow();
    const noUrut  = Math.max(1, lastRow);
    const ts      = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    const row     = new Array(rowWidth).fill('');

    set(row, 'no',               noUrut);
    set(row, 'timestamp',        ts);
    set(row, 'nama lengkap',     data.namaLengkap    || '');
    set(row, 'nip',              data.nip            || '');
    set(row, 'jabatan',          data.jabatan        || '');
    set(row, 'nomor hp',         normalizePhone(data.nomorHp));
    set(row, 'plat nomor',       data.platNomor      || '');
    set(row, 'merk',             data.merk           || '');
    set(row, 'tipe',             data.tipe           || '');
    set(row, 'model',            data.model          || '');
    set(row, 'jenis',            data.jenis          || '');
    set(row, 'tgl. keluar',      data.tglPinjam      || '');
    set(row, 'jam keluar',       data.wktPinjam      || '');
    set(row, 'tgl. masuk',       data.tglKembali     || '');
    set(row, 'jam masuk',        data.wktKembali     || '');
    set(row, 'keperluan',        data.keperluan      || '');
    set(row, 'lokasi tujuan',    data.lokasiTujuan   || '');
    set(row, 'kendala',          data.kendala        || '');
    set(row, 'nama pengguna',    data.namaPengemudi  || '');
    set(row, 'nip pengguna',     data.nipPengemudi   || '');
    set(row, 'no. hp pengguna',  normalizePhone(data.telpPengemudi));
    set(row, 'jumlah penumpang', data.jumlahPenumpang || 0);
    set(row, 'daftar penumpang', data.penumpang      || '');
    set(row, 'telp penumpang',   normalizePhonesInList(data.telpPenumpang));
    set(row, 'status',           'Pengajuan');

    sheet.appendRow(row);
    const newRow = sheet.getLastRow();
    _formatRow(sheet, newRow);

    // Force format teks pada kolom nomor telepon
    ['nomor hp','no. hp pengguna','telp penumpang'].forEach(function(name){
      var i = headerMap[name];
      if (i === undefined || i >= rowWidth) return;
      var val = String(row[i] || '').trim();
      if (!val) return;
      sheet.getRange(newRow, i + 1).setNumberFormat('@STRING@').setValue(val);
    });

    // Update Status di sheet Kendaraan → "Permohonan"
    try {
      const sK = ss.getSheetByName(SHEET_KENDARAAN);
      if (sK) {
        const rK = sK.getDataRange().getValues();
        const hK = rK[0].map(function(h){ return String(h).trim().toLowerCase(); });
        var iPlatK = -1, iStatK = -1;
        for (var kc = 0; kc < hK.length; kc++) {
          if (hK[kc] === 'plat nomor' || hK[kc] === 'plat_nomor' || hK[kc] === 'plat') iPlatK = kc;
          if (hK[kc] === 'status') iStatK = kc;
        }
        if (iPlatK >= 0 && iStatK >= 0) {
          for (var kr = 1; kr < rK.length; kr++) {
            if (String(rK[kr][iPlatK] || '').trim().toUpperCase() === String(data.platNomor || '').trim().toUpperCase()) {
              sK.getRange(kr + 1, iStatK + 1).setValue('Permohonan');
              break;
            }
          }
        }
      }
    } catch(e) { Logger.log('simpanPeminjaman - update Kendaraan error: ' + e.message); }

    Logger.log('simpanPeminjaman OK — noUrut: ' + noUrut + ', rowWidth: ' + rowWidth);
    return { status: 'ok', noUrut: noUrut };
  } catch (err) {
    Logger.log('simpanPeminjaman ERROR: ' + err.message);
    return { status: 'error', message: err.message };
  }
}

// ── Helper Internal ───────────────────────────────────────────

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    _formatHeader(sheet, headers.length);
  }
  return sheet;
}

// ── Normalisasi Nomor Telepon ─────────────────────────────────
/**
 * Pastikan nomor HP/telepon disimpan sebagai teks di Sheets.
 * Awalan 0 tidak diubah — disimpan apa adanya.
 */
function normalizePhone(raw) {
  if (!raw) return '';
  return String(raw).trim(); // kembalikan apa adanya sebagai string
}

/**
 * Normalisasi daftar nomor telepon yang dipisahkan ' | '
 */
function normalizePhonesInList(raw) {
  if (!raw) return '';
  return raw.split(' | ').map(function(p){ return normalizePhone(p.trim()); }).join(' | ');
}

function _formatHeader(sheet, n) {
  sheet.getRange(1, 1, 1, n)
    .setBackground('#1565C0').setFontColor('#fff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, n, 120);
  sheet.setColumnWidth(3, 200);  // Nama Lengkap
  sheet.setColumnWidth(22, 240); // Daftar Penumpang

  // Format kolom nomor HP & telepon sebagai teks agar angka 0 di depan tidak hilang
  var colHP    = HEADERS_PEMINJAMAN.indexOf('Nomor HP') + 1;
  var colTelp  = HEADERS_PEMINJAMAN.indexOf('Telp Pengemudi') + 1;
  var colTelpP = HEADERS_PEMINJAMAN.indexOf('Telp Penumpang') + 1;
  [colHP, colTelp, colTelpP].forEach(function(col){
    if(col > 0) sheet.getRange(2, col, 1000, 1).setNumberFormat('@');
  });
}

function _formatRow(sheet, rowNum) {
  // Cari kolom Status secara aktual dari header sheet (bukan asumsi panjang HEADERS)
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var statusCol = 0;
  headerRow.forEach(function(h, i){
    if (String(h).trim().toLowerCase() === 'status') statusCol = i + 1;
  });
  const n  = statusCol > 0 ? statusCol : sheet.getLastColumn();
  const bg = rowNum % 2 === 0 ? '#E3F2FD' : '#FFFFFF';
  sheet.getRange(rowNum, 1, 1, n)
    .setBackground(bg)
    .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(rowNum, n)
    .setBackground('#FFF9C4').setFontWeight('bold').setHorizontalAlignment('center');
}

// ── Notifikasi Email (Opsional) ───────────────────────────────
function _kirimNotifikasi(data, noUrut) {
  const ADMIN_EMAIL = 'admin@kemenhub.go.id'; // ganti dengan email admin
  const subject = '[Peminjaman #' + noUrut + '] ' + data.platNomor + ' - ' + data.keperluan;
  const body = [
    'No. Urut    : ' + noUrut,
    'Timestamp   : ' + new Date().toLocaleString('id-ID'),
    '',
    'PEMOHON',
    'Nama        : ' + data.namaLengkap,
    'NIP         : ' + data.nip,
    'Jabatan     : ' + data.jabatan,
    'No. HP      : ' + data.nomorHp,
    '',
    'KENDARAAN   : ' + data.merk + ' ' + data.tipe + ' ' + data.model + ' - ' + data.platNomor,
    'Pinjam      : ' + data.tglPinjam + ' ' + data.wktPinjam,
    'Kembali     : ' + data.tglKembali + ' ' + data.wktKembali,
    '',
    'Keperluan   : ' + data.keperluan,
    'Lokasi      : ' + data.lokasiTujuan,
    '',
    'Pengemudi   : ' + data.namaPengemudi + ' (' + data.telpPengemudi + ')',
    'Penumpang   : ' + data.jumlahPenumpang + ' orang - ' + (data.penumpang || '-'),
  ].join('\n');

  try { MailApp.sendEmail(ADMIN_EMAIL, subject, body); }
  catch(e) { Logger.log('Email error: ' + e.message); }
}

// ── Utilitas Tambahan ─────────────────────────────────────────

/** Update status peminjaman - panggil manual dari editor jika perlu */
function updateStatus(noUrut, statusBaru) {
  const ss    = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = ss.getSheetByName(SHEET_PEMINJAMAN);
  if (!sheet) return false;

  const rows      = sheet.getDataRange().getValues();
  const statusCol = HEADERS_PEMINJAMAN.indexOf('Status') + 1;
  const colorMap  = {
    'Disetujui' : '#C8E6C9',
    'Ditolak'   : '#FFCDD2',
    'Selesai'   : '#B2EBF2',
    'Pengajuan' : '#FFF9C4',
  };

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == noUrut) {
      const cell = sheet.getRange(i + 1, statusCol);
      cell.setValue(statusBaru).setBackground(colorMap[statusBaru] || '#fff');
      return true;
    }
  }
  return false;
}

/**
 * Jalankan otomatis setiap hari via Time-based Trigger.
 * Cek setiap baris di sheet Peminjaman:
 *  - H-2 atau H-1 sebelum Tgl. Masuk & status bukan Selesai → orange peringatan
 *  - Melewati Tgl. Masuk & status bukan Selesai            → merah peringatan
 *
 * Cara pasang trigger:
 *   Jalankan setupTrigger() sekali dari GAS Editor.
 */
function updateStatusOtomatis() {
  try {
    const ss    = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName(SHEET_PEMINJAMAN);
    if (!sheet || sheet.getLastRow() < 2) return;

    const rows    = sheet.getDataRange().getValues();
    const heads   = rows[0].map(function(h){ return String(h).trim().toLowerCase(); });

    // Cari kolom Tgl. Masuk (nama baru dan lama)
    function hIdx(names) {
      var i = -1;
      names.forEach(function(n){ if(i<0) i = heads.indexOf(n); });
      return i;
    }
    const iTglMasuk = hIdx(['tgl. masuk','tgl. kembali','tanggal masuk','tanggal kembali']);
    const iStatus   = hIdx(['status']);
    if (iTglMasuk < 0 || iStatus < 0) return;

    const today     = new Date(); today.setHours(0,0,0,0);
    const statusCol = iStatus + 1;

    const SKIP_STATUS = ['selesai','belum perpanjang masa peminjaman'];

    for (var i = 1; i < rows.length; i++) {
      var rawTgl  = rows[i][iTglMasuk];
      var status  = String(rows[i][iStatus] || '').trim().toLowerCase();

      if (!rawTgl || status === 'selesai') continue;

      // Parse tanggal — support "dd mmmm yyyy" dan Date object
      var tglMasuk;
      if (rawTgl instanceof Date) {
        tglMasuk = new Date(rawTgl); tglMasuk.setHours(0,0,0,0);
      } else {
        var str = String(rawTgl).trim();
        tglMasuk = parseIndonesianDate(str);
        if (!tglMasuk) continue;
      }

      var diffMs   = tglMasuk - today;
      var diffDays = Math.round(diffMs / 86400000); // positif = belum lewat
      var cell     = sheet.getRange(i+1, statusCol);

      if (diffDays < 0) {
        // Sudah melewati Tgl. Masuk
        cell.setValue('Belum perpanjang masa peminjaman')
            .setBackground('#FFCDD2')
            .setFontColor('#B71C1C')
            .setFontWeight('bold');
      } else if (diffDays <= 2) {
        // H-2 atau H-1: hampir habis
        cell.setValue('Batas waktu peminjaman 30 hari sudah akan habis!')
            .setBackground('#FFE0B2')
            .setFontColor('#E65100')
            .setFontWeight('bold');
      }
      // Jika masih jauh, biarkan status tidak berubah
    }
    Logger.log('updateStatusOtomatis selesai.');
  } catch(e) {
    Logger.log('updateStatusOtomatis ERROR: ' + e.message);
  }
}

/** Parse tanggal format "dd mmmm yyyy" (Indonesia) → Date */
function parseIndonesianDate(str) {
  var BULAN = ['januari','februari','maret','april','mei','juni',
               'juli','agustus','september','oktober','november','desember'];
  var parts = str.toLowerCase().trim().split(/\s+/);
  if (parts.length < 3) return null;
  var d = parseInt(parts[0]);
  var m = BULAN.indexOf(parts[1]);
  var y = parseInt(parts[2]);
  if (isNaN(d) || m < 0 || isNaN(y)) return null;
  var dt = new Date(y, m, d); dt.setHours(0,0,0,0);
  return dt;
}

/**
 * Pasang Time-based Trigger harian untuk updateStatusOtomatis.
 * Jalankan sekali dari GAS Editor.
 */
function setupTrigger() {
  // Hapus trigger lama agar tidak dobel
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'updateStatusOtomatis') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('updateStatusOtomatis')
    .timeBased()
    .everyDays(1)
    .atHour(7) // Cek setiap hari jam 07.00 WIB
    .create();
  Logger.log('Trigger harian updateStatusOtomatis berhasil dipasang.');
}
