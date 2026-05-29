// ============================================================
//  INVENTORY KENDARAAN DINAS - Google Apps Script
//  Code.gs - Backend & Router
// ============================================================

// --- Spreadsheet Formula Protection ---
function sanitizeCell(val) {
  if (typeof val !== 'string' || !val) return val;
  // Cegah spreadsheet formula injection: = + - @
  if (/^[=+\-@]/.test(val.trim())) return "'" + val;
  return val;
}

// --- KONFIGURASI via Script Properties ---
function getConfig(key) {
  var val = PropertiesService.getScriptProperties().getProperty(key);
  if (val === null) throw new Error('Konfigurasi "' + key + '" belum diset. Jalankan setupSystem() sekali dari GAS Editor.');
  return val;
}

/**
 * Jalankan fungsi ini SEKALI dari GAS Editor untuk menyimpan konfigurasi
 * ke Script Properties (aman, tidak terlihat di source code).
 */
function setupSystem() {
  var props = PropertiesService.getScriptProperties();
  // Hanya set nilai DEFAULT jika belum ada (tidak overwrite nilai yang sudah diset)
  var defaults = {};
  defaults['SPREADSHEET_ID']    = '1FAnk26sO1B_vjMy2v_WUp74nViXeHyRgj5suhyiHQ1U';
  defaults['SHEET_USERS']       = 'Users';
  defaults['SHEET_KENDARAAN']   = 'Kendaraan';
  defaults['SHEET_PEMINJAMAN']  = 'Peminjaman';
  defaults['SHEET_LOG']         = 'Log';
  defaults['SHEET_MOBIL_KELUAR'] = 'MobilKeluar';
  defaults['SHEET_MOBIL_MASUK'] = 'MobilMasuk';
  defaults['SHEET_SESSIONS']    = 'Sessions';
  defaults['APP_NAME']          = 'Inventory Kendaraan Dinas';
  defaults['VERSION']           = '1.0.0';
  defaults['SESSION_TTL_HOURS'] = '8';
  defaults['LOGO_FOLDER_ID']    = '1K3NjyzSmvTzNnjSZjDlp2Y1Xz2rMH997';
  defaults['ADMIN_USERNAME']    = 'wiranto';
  defaults['ADMIN_NAME']        = 'Super Administrator';
  defaults['ADMIN_ROLE']        = 'Super Administrator';
  defaults['ADMIN_PASSWORD']    = hashPassword('super@dmin123');
  var count = 0;
  for (var key in defaults) {
    if (props.getProperty(key) === null) {
      props.setProperty(key, defaults[key]);
      count++;
    }
  }
  return { success: true, message: count + ' konfigurasi baru disimpan. (' + Object.keys(defaults).length + ' total)' };
}

// --- UPDATE TOKEN untuk auto-refresh ---
function getUpdateToken_() {
  var t = PropertiesService.getScriptProperties().getProperty('UPDATE_TOKEN');
  return t ? parseInt(t, 10) : 0;
}
function bumpUpdateToken_() {
  var next = getUpdateToken_() + 1;
  PropertiesService.getScriptProperties().setProperty('UPDATE_TOKEN', String(next));
}
function getUpdateTokenPublic(token) {
  try { requireAuth(token); return { success: true, token: getUpdateToken_() }; }
  catch(e) { return { success: false }; }
}

/**
 * Helper: validasi token dan return session object.
 * Gunakan di awal setiap fungsi yang membutuhkan autentikasi.
 */
function requireAuth(token) {
  var session = validateSession(token);
  if (!session.valid) throw new Error('Session tidak valid.');
  return session;
}

// ============================================================
//  ENTRY POINT - Web App
// ============================================================

function doGet(e) {
  const page  = (e && e.parameter && e.parameter.page)  || 'login';

  // Dashboard: token validasi dilakukan di frontend (sessionStorage)
  if (page === 'dashboard') {
    return HtmlService.createTemplateFromFile('Dashboard').evaluate()
      .setTitle(getConfig('APP_NAME') + ' — Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=5.0, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }

  // Form peminjaman publik (tanpa login)
  if (page === 'pinjam') {
    return HtmlService.createTemplateFromFile('pinjam').evaluate()
      .setTitle('Form Peminjaman Kendaraan Dinas')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=5.0, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }

  // Default: halaman login
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle(getConfig('APP_NAME'))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=5.0, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * Kembalikan URL deployment Web App ini (dipanggil dari frontend)
 */
function getScriptUrl(token) {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    return null;
  }
}

// ============================================================
//  AUTENTIKASI
// ============================================================

/**
 * Login user - dipanggil dari frontend
 * @param {string} username
 * @param {string} password
 * @returns {Object} result
 */
/**
 * Hash password menggunakan SHA-256, return hex string 64 karakter.
 */
function hashPassword(pw) {
  try {
    return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8)
      .map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'))
      .join('');
  } catch(e) {
    return pw; // fallback plain jika gagal
  }
}

// ── Super Administrator (disimpan di Script Properties, tidak hardcoded) ──
// Akses melalui getConfig()

// ============================================================
//  SESSION MANAGEMENT — Sheet-based (tidak pakai Script Properties)
// ============================================================

function getSessionSheetName() {
  try { return getConfig('SHEET_SESSIONS'); } catch(e) { return 'Sessions'; }
}

function getSessionTtlHours() {
  return parseInt(getConfig('SESSION_TTL_HOURS')) || 8;
}

function getSessionSheet() {
  const ss = getSpreadsheet();
  var sheetName = getSessionSheetName();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1,1,1,6).setValues([['Token','Username','Nama','Role','LoginTime','Expires']]);
    sh.hideSheet(); // sembunyikan dari pengguna biasa
  }
  return sh;
}

function saveSession(token, data) {
  const sh = getSessionSheet();
  sh.appendRow([token, data.username, data.nama, data.role, data.loginTime, data.expires]);
}

function getSession(token) {
  const sh = getSessionSheet();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(token).trim()) {
      return { token: rows[i][0], username: rows[i][1], nama: rows[i][2],
               role: rows[i][3], loginTime: rows[i][4], expires: rows[i][5] };
    }
  }
  return null;
}

function deleteSession(token) {
  const sh = getSessionSheet();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(token).trim()) {
      sh.deleteRow(i + 1);
      return;
    }
  }
}

function cleanExpiredSessions() {
  try {
    const sh = getSessionSheet();
    const rows = sh.getDataRange().getValues();
    const now = new Date();
    // Hapus dari bawah supaya index tidak bergeser
    for (let i = rows.length - 1; i >= 1; i--) {
      const expires = new Date(rows[i][5]);
      if (now > expires) sh.deleteRow(i + 1);
    }
  } catch(e) {}
}

function loginUser(username, password) {
  try {
    if (!username || !password) {
      return { success: false, message: 'Username dan password tidak boleh kosong.' };
    }

    // Bersihkan session kadaluarsa
    try { cleanExpiredSessions(); } catch(e) {}

    // Rate limiting: cek gagal login berulang
    var rateLimitKey = 'LOGIN_FAIL_' + username.toLowerCase().trim();
    var props = PropertiesService.getScriptProperties();
    var failData = props.getProperty(rateLimitKey);
    var failCount = 0, failTime = 0;
    if (failData) {
      var parts = failData.split('|');
      failCount = parseInt(parts[0], 10) || 0;
      failTime  = parseInt(parts[1], 10) || 0;
      // Reset jika sudah lewat 15 menit
      if (Date.now() - failTime > 15 * 60 * 1000) { failCount = 0; failTime = 0; }
    }
    if (failCount >= 5) {
      return { success: false, message: 'Akun diblokir sementara karena terlalu banyak percobaan gagal. Coba lagi 15 menit.' };
    }

    // Helper: catat gagal login
    function recordLoginFail() {
      props.setProperty(rateLimitKey, (failCount + 1) + '|' + Date.now());
    }

    // 1. Cek Super Administrator dari Script Properties
    var adminUser = getConfig('ADMIN_USERNAME');
    var adminPass = getConfig('ADMIN_PASSWORD');
    var adminName = getConfig('ADMIN_NAME');
    var adminRole = getConfig('ADMIN_ROLE');
    if (username.toLowerCase().trim() === adminUser &&
        (password === adminPass || hashPassword(password) === adminPass)) {
      // Reset fail counter on success
      props.deleteProperty(rateLimitKey);
      const token = generateToken();
      const now   = new Date();
      saveSession(token, {
        username: adminUser, nama: adminName,
        role: adminRole,
        loginTime: now.toISOString(),
        expires: new Date(now.getTime() + parseInt(getConfig('SESSION_TTL_HOURS')||8)*60*60*1000).toISOString()
      });
      writeLog(adminUser, 'LOGIN', 'Super Administrator login');
      return { success: true, token, nama: adminName, role: adminRole,
               username: adminUser,
               message: 'Login berhasil! Selamat datang, ' + adminName };
    }

    // 2. Cek di sheet Users
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_USERS'));
    if (!sheet) { initializeSheets(); return loginUser(username, password); }

    const data = sheet.getDataRange().getValues();
    const hdrs = data[0].map(h => String(h).trim());
    const unIdx = hdrs.indexOf('Username');
    const pwIdx = hdrs.indexOf('Password Hash');
    const nmIdx = hdrs.indexOf('Nama Lengkap');
    const rlIdx = hdrs.indexOf('Role');
    const stIdx = hdrs.indexOf('Status');
    if (unIdx === -1 || pwIdx === -1) {
      return { success: false, message: 'Konfigurasi sheet Users tidak valid.' };
    }
    for (let i = 1; i < data.length; i++) {
      const row     = data[i];
      const uname   = String(row[unIdx]).toLowerCase().trim();
      const pwd     = String(row[pwIdx]).trim();
      const status  = String(row[stIdx] || 'aktif').toLowerCase();
      const roleVal = String(row[rlIdx] || '').trim();

      if (uname === username.toLowerCase().trim()) {
        if (status !== 'aktif') {
          return { success: false, message: 'Akun Anda tidak aktif. Hubungi administrator.' };
        }
        if (pwd === password || pwd === hashPassword(password)) {
          // Reset fail counter on success
          props.deleteProperty(rateLimitKey);
          const token = generateToken();
          const nama  = nmIdx !== -1 ? String(row[nmIdx]) : '';
          const now   = new Date();
          saveSession(token, {
            username: uname, nama, role: roleVal,
            loginTime: now.toISOString(),
            expires: new Date(now.getTime() + getSessionTtlHours()*60*60*1000).toISOString()
          });
          const llCol = hdrs.indexOf('Last Login');
          if (llCol !== -1) sheet.getRange(i+1, llCol+1).setValue(new Date());
          writeLog(uname, 'LOGIN', 'Berhasil login');
          return { success: true, token, nama, role: roleVal,
                   username: uname,
                   message: 'Login berhasil! Selamat datang, ' + nama };
        } else {
          recordLoginFail();
          writeLog(username, 'LOGIN_GAGAL', 'Password salah');
          return { success: false, message: 'Password salah. Silakan coba lagi.' };
        }
      }
    }
    recordLoginFail();
    writeLog(username, 'LOGIN_GAGAL', 'Username tidak ditemukan');
    return { success: false, message: 'Username tidak ditemukan.' };

  } catch (err) {
    console.error('loginUser error:', err);
    return { success: false, message: 'Terjadi kesalahan sistem: ' + err.message };
  }
}

/**
 * Logout user
 */
function logoutUser(token) {
  try {
    const sess = getSession(token);
    if (sess) {
      writeLog(sess.username, 'LOGOUT', 'User logout');
      deleteSession(token);
    }
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Validasi token session
 */
function validateSession(token) {
  try {
    if (!token) return { valid: false };
    const sess = getSession(token);
    if (!sess) return { valid: false };

    const now     = new Date();
    const expires = new Date(sess.expires);
    if (now > expires) {
      deleteSession(token);
      return { valid: false, message: 'Session telah berakhir. Silakan login kembali.' };
    }
    return { valid: true, ...sess };
  } catch (err) {
    return { valid: false };
  }
}

/**
 * Ambil info user dari token (dipanggil frontend)
 */
function getSessionInfo(token) {
  return validateSession(token);
}

// ============================================================
//  INISIALISASI SPREADSHEET
// ============================================================

function initializeSheets() {
  const ss = getSpreadsheet();

  // --- Sheet Users ---
  let usersSheet = ss.getSheetByName(getConfig('SHEET_USERS'));
  if (!usersSheet) {
    usersSheet = ss.insertSheet(getConfig('SHEET_USERS'));
    usersSheet.appendRow(['ID', 'Tgl. Dibuat', 'Username', 'Password Hash', 'Role', 'Nama Lengkap', 'Jabatan', 'Last Login', 'Status']);
    // Tidak ada akun default — semua pengguna dibuat melalui menu Pengguna oleh Super Administrator
    formatHeaderRow(usersSheet);
    applyZebraRows(usersSheet);
  }

  // --- Sheet Kendaraan ---
  setupKendaraanSheet();

  // --- Sheet Peminjaman ---
  let peminjamanSheet = ss.getSheetByName(getConfig('SHEET_PEMINJAMAN'));
  if (!peminjamanSheet) {
    peminjamanSheet = ss.insertSheet(getConfig('SHEET_PEMINJAMAN'));
    peminjamanSheet.getRange(1, 1, 1, HEADERS_PEMINJAMAN.length).setValues([HEADERS_PEMINJAMAN]);
    formatHeaderRow(peminjamanSheet);
    applyZebraRows(peminjamanSheet);
  }

  // --- Sheet Log ---
  let logSheet = ss.getSheetByName(getConfig('SHEET_LOG'));
  if (!logSheet) {
    logSheet = ss.insertSheet(getConfig('SHEET_LOG'));
    logSheet.appendRow(['Timestamp', 'User', 'Aksi', 'Detail', 'IP']);
    formatHeaderRow(logSheet);
    applyZebraRows(logSheet);
  }
}

// ============================================================
//  HELPER FUNCTIONS
// ============================================================

function getSpreadsheet() {
  // Standalone web app — buka spreadsheet via ID
  return SpreadsheetApp.openById(getConfig('SPREADSHEET_ID'));
}

function generateToken() {
  return Utilities.getUuid().replace(/-/g, '') + Date.now().toString(36);
}

function formatHeaderRow(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setBackground('#1a73e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/**
 * Terapkan zebra striping (baris ganjil/genap beda warna) ke seluruh sheet.
 * Menggunakan SpreadsheetApp banding — satu kali call, berlaku otomatis untuk baris baru.
 */
function applyZebraRows(sheet) {
  try {
    // Hapus banding lama jika ada
    const bandings = sheet.getBandings();
    bandings.forEach(b => b.remove());
    // Terapkan banding baru: seluruh data range
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const range   = sheet.getRange(1, 1, sheet.getMaxRows(), lastCol);
    range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    // Override header dengan warna biru SKTJ
    formatHeaderRow(sheet);
  } catch(e) {
    console.warn('applyZebraRows error:', e.message);
  }
}

/**
 * Terapkan zebra ke semua sheet yang ada di spreadsheet.
 * Dipanggil sekali via Apps Script editor: Menu > Run > applyZebraAllSheets
 */
function applyZebraAllSheets() {
  try {
    const ss     = getSpreadsheet();
    const sheets = ss.getSheets();
    sheets.forEach(sheet => {
      const name = sheet.getName();
      // Skip sheet log (biarkan apa adanya)
      if (name.toLowerCase() === 'log') return;
      applyZebraRows(sheet);
      console.log('Zebra applied: ' + name);
    });
    return { success: true, message: 'Zebra striping berhasil diterapkan ke ' + sheets.length + ' sheet.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function writeLog(user, action, detail) {
  try {
    const ss = getSpreadsheet();
    const logSheet = ss.getSheetByName(getConfig('SHEET_LOG'));
    if (logSheet) {
      logSheet.appendRow([new Date(), user, action, detail, '']);
    }
  } catch (e) {
    console.error('writeLog error:', e);
  }
}

// ============================================================
//  FUNGSI KENDARAAN (akan dikembangkan)
// ============================================================

/**
 * Setup sheet Kendaraan: buat jika belum ada, set header baku, isi data contoh.
 * Jalankan fungsi ini SEKALI dari Apps Script Editor (Run > setupKendaraanSheet)
 * setelah deploy, atau akan dipanggil otomatis oleh initializeSheets.
 */
function setupKendaraanSheet() {
  const ss = getSpreadsheet();
  const HEADERS = ['ID_Kendaraan', 'Merk', 'Tipe', 'Jenis', 'Model', 'Tahun', 'Plat Nomor', 'Status', 'Pajak'];

  let sheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));

  if (!sheet) {
    sheet = ss.insertSheet(getConfig('SHEET_KENDARAAN'));
  } else {
    sheet.clearContents();
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  formatHeaderRow(sheet);
  applyZebraRows(sheet);
  sheet.autoResizeColumns(1, HEADERS.length);

  return { success: true, message: 'Sheet Kendaraan berhasil disiapkan.' };
}

/**
 * Ambil daftar kendaraan TANPA validasi token.
 * Digunakan langsung dari Dashboard tanpa perlu login token.
 */
/**
 * Fungsi debug — panggil dari browser console:
 *   google.script.run.withSuccessHandler(console.log).pingPublic()
 * Harus return { ok: true, sheetNames: [...] }
 */
function pingPublic(token) {
  try {
    var session = requireAuth(token);
    const ss     = getSpreadsheet();
    const sheets = ss.getSheets().map(function(s) { return s.getName(); });
    return { ok: true, spreadsheetName: ss.getName(), sheetNames: sheets };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

/**
 * Ambil daftar user TANPA token.
 */
function getListUsersPublic(token) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_USERS'));
    if (!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1)
      .filter(row => row.some(cell => String(cell).trim() !== ''))
      .filter(row => {
        // Super Administrator tidak pernah tampil
        const roleIdx = headers.indexOf('Role');
        const unIdx   = headers.indexOf('Username');
        const roleVal = roleIdx !== -1 ? String(row[roleIdx]).toLowerCase() : '';
        const unVal   = unIdx   !== -1 ? String(row[unIdx]).toLowerCase()  : '';
        return roleVal !== 'super administrator' && unVal !== getConfig('ADMIN_USERNAME').toLowerCase();
      })
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          if (h === 'Password' || h === 'Password Hash') {
            // Fix 2: tampilkan full SHA-256 hash (64 hex chars) — tidak bisa dibalikkan
            const raw = row[i] ? String(row[i]) : '';
            // Jika sudah berupa hex 64 chars (SHA-256), tampilkan apa adanya
            // Jika masih plain text lama, hash dulu untuk display
            obj['Password Hash'] = raw && /^[0-9a-f]{64}$/i.test(raw) ? raw : (raw ? hashPassword(raw) : '—');
          } else {
            obj[h] = (row[i] instanceof Date)
              ? Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'dd/MM/yyyy')
              : (row[i] !== null && row[i] !== undefined ? row[i] : '');
          }
        });
        return obj;
      });
    return { success: true, data: rows };
  } catch (err) {
    return { success: false, message: err.message };
  }
}
function addUserPublic(token, payload) {
  try {
    var session = requireAuth(token);
    if (!payload.nama)     return { success: false, message: 'Nama wajib diisi.' };
    if (!payload.username) return { success: false, message: 'Username wajib diisi.' };
    if (!payload.password) return { success: false, message: 'Password wajib diisi.' };
    if (!payload.role)     return { success: false, message: 'Role wajib dipilih.' };
    // Blokir pembuatan Super Administrator via form
    if (payload.role.toLowerCase() === 'super administrator') {
      return { success: false, message: 'Role Super Administrator tidak dapat dibuat via form.' };
    }
    // Blokir username sama dengan Super Administrator
    if (payload.username.toLowerCase().trim() === getConfig('ADMIN_USERNAME').toLowerCase()) {
      return { success: false, message: 'Username tidak tersedia.' };
    }

    const blockedRoles = ['superadministrator','super administrator'];
    if (blockedRoles.includes(payload.role.toLowerCase())) {
      return { success: false, message: 'Role Super Administrator tidak dapat dibuat melalui antarmuka ini.' };
    }
    if (payload.username.toLowerCase().trim() === getConfig('ADMIN_USERNAME').toLowerCase()) {
      return { success: false, message: 'Username tersebut tidak tersedia.' };
    }

    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_USERS'));
    if (!sheet) return { success: false, message: 'Sheet Users tidak ditemukan.' };

    const data  = sheet.getDataRange().getValues();
    const hdrs  = data[0].map(h => String(h).trim());
    const unCol = hdrs.indexOf('Username') !== -1 ? hdrs.indexOf('Username') : hdrs.indexOf('username');
    if (unCol !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][unCol]).trim().toLowerCase() === payload.username.toLowerCase()) {
          return { success: false, message: 'Username sudah digunakan.' };
        }
      }
    }

    const id  = 'USR_' + String(sheet.getLastRow()).padStart(4, '0');
    const now = new Date();
    // Fix 2: hash password SHA-256 sebelum disimpan
    const pwHash = hashPassword(payload.password);
    const newRow = hdrs.map(h => {
      switch (h) {
        case 'ID':             return id;
        case 'Nama Pengguna':
        case 'Nama Lengkap':   return payload.nama;
        case 'Jabatan':        return payload.jabatan || '';
        case 'Username':
        case 'username':       return payload.username;
        case 'Password':
        case 'Password Hash':  return pwHash;
        case 'Role':           return payload.role;
        case 'Status':         return 'aktif';
        case 'Last Login':     return '';
        case 'Tgl. Dibuat':
        case 'Dibuat':         return Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        default:               return '';
      }
    });
    sheet.appendRow(newRow);
    writeLog('dashboard', 'ADD_USER', 'Tambah user: ' + payload.username);
    bumpUpdateToken_(); return { success: true, id, message: 'Pengguna berhasil ditambahkan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Update user TANPA token.
 */
/**
 * Pastikan sheet Users memiliki semua kolom yang diperlukan.
 * Dijalankan saat ada operasi write ke Users.
 */
function ensureUsersSheetHeaders() {
  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_USERS'));
    if (!sheet || sheet.getLastRow() < 1) return;
    const hdrs    = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    // Migrasi: rename "Password" → "Password Hash"
    const pwIdx = hdrs.indexOf('Password');
    if (pwIdx !== -1 && hdrs.indexOf('Password Hash') === -1) {
      sheet.getRange(1, pwIdx + 1).setValue('Password Hash');
      hdrs[pwIdx] = 'Password Hash';
      console.log('Renamed column "Password" → "Password Hash"');
    }
    // Migrasi: rename "Nama Pengguna" → "Nama Lengkap"
    const npIdx = hdrs.indexOf('Nama Pengguna');
    if (npIdx !== -1 && hdrs.indexOf('Nama Lengkap') === -1) {
      sheet.getRange(1, npIdx + 1).setValue('Nama Lengkap');
      hdrs[npIdx] = 'Nama Lengkap';
      console.log('Renamed column "Nama Pengguna" → "Nama Lengkap"');
    }
    // Migrasi: rename "Dibuat" → "Tgl. Dibuat"
    const dbIdx = hdrs.indexOf('Dibuat');
    if (dbIdx !== -1 && hdrs.indexOf('Tgl. Dibuat') === -1) {
      sheet.getRange(1, dbIdx + 1).setValue('Tgl. Dibuat');
      hdrs[dbIdx] = 'Tgl. Dibuat';
      console.log('Renamed column "Dibuat" → "Tgl. Dibuat"');
    }
    const REQUIRED = ['ID', 'Tgl. Dibuat', 'Username', 'Password Hash', 'Role', 'Nama Lengkap', 'Jabatan', 'Last Login', 'Status'];
    REQUIRED.forEach(col => {
      if (hdrs.indexOf(col) === -1) {
        const nextCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextCol).setValue(col);
        console.log('Added missing column: ' + col);
      }
    });
  } catch(e) { console.error('ensureUsersSheetHeaders:', e); }
}

function updateUserPublic(token, id, payload) {
  try {
    var session = requireAuth(token);
    ensureUsersSheetHeaders();
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_USERS'));
    if (!sheet) return { success: false, message: 'Sheet Users tidak ditemukan.' };

    const data    = sheet.getDataRange().getValues();
    const hdrs    = data[0].map(h => String(h).trim());
    const idCol   = hdrs.indexOf('ID');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]).trim() === String(id).trim()) {
        const rowNum = i + 1;
        const updates = {
          'Nama Lengkap':  payload.nama     ?? '',
          'Jabatan':       payload.jabatan  ?? '',
          'Username':      payload.username ?? '',
          'username':      payload.username ?? '',
          'Role':          payload.role     ?? ''
        };
        if (payload.password) {
          const pwHash = hashPassword(payload.password);
          updates['Password'] = pwHash;
          updates['Password Hash'] = pwHash;
        }
        for (const [header, value] of Object.entries(updates)) {
          const c = hdrs.indexOf(header);
          // Fix: update even if value is empty string (allow clearing jabatan etc)
          if (c !== -1 && value !== undefined && value !== null) {
            sheet.getRange(rowNum, c + 1).setValue(value);
          }
        }
        writeLog('dashboard', 'UPDATE_USER', 'Update user ID: ' + id);
        bumpUpdateToken_(); return { success: true, message: 'Pengguna berhasil diperbarui.' };
      }
    }
    return { success: false, message: 'Pengguna tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Hapus user TANPA token.
 */
function deleteUserPublic(token, id) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_USERS'));
    if (!sheet) return { success: false, message: 'Sheet Users tidak ditemukan.' };

    const data  = sheet.getDataRange().getValues();
    const hdrs  = data[0].map(h => String(h).trim());
    const idCol = hdrs.indexOf('ID');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        writeLog('dashboard', 'DELETE_USER', 'Hapus user ID: ' + id);
        bumpUpdateToken_(); return { success: true };
      }
    }
    return { success: false, message: 'Pengguna tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getListKendaraanPublic(token) {
  try {
    var session = requireAuth(token);
    const ss  = getSpreadsheet();
    let sheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));

    // Coba nama alternatif jika tidak ketemu
    if (!sheet) {
      const names = ['kendaraan','Kendaraan','KENDARAAN','Data Kendaraan'];
      for (const n of names) {
        sheet = ss.getSheetByName(n);
        if (sheet) break;
      }
    }
    if (!sheet) return { success: false, message: 'Sheet "' + getConfig('SHEET_KENDARAAN') + '" tidak ditemukan. Sheet yang ada: ' + ss.getSheets().map(s=>s.getName()).join(', ') };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [], message: 'Sheet kosong atau hanya ada header' };

    // Trim semua header agar key bersih
    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1)
      .filter(row => row.some(cell => String(cell).trim() !== ''))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          obj[h] = (row[i] instanceof Date)
            ? Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'dd/MM/yyyy')
            : row[i];
        });
        return obj;
      });
    return { success: true, data: rows, headers: headers };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Tambah kendaraan TANPA token (langsung dari Dashboard).
 */
function addKendaraanPublic(token, payload) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
    if (!sheet) return { success: false, message: 'Sheet Kendaraan tidak ditemukan.' };

    const lastRow = sheet.getLastRow();
    const seq = String(lastRow).padStart(3, '0');
    const id  = 'IKD_' + seq;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const fieldMap = {
      'ID_Kendaraan': id, 'ID': id,
      'Merk':         payload.merk     || '',
      'Tipe':         payload.tipe     || '',
      'Jenis':        payload.jenis    || '',
      'Model':        payload.model    || '',
      'Tahun':        payload.tahun    || '',
      'Plat_Nomor':   payload.noPolisi || '',
      'Plat Nomor':   payload.noPolisi || '',
      'No. Polisi':   payload.noPolisi || '',
      'Status':       payload.status   || 'Tersedia',
      'Pajak':        payload.pajak    || '',
    };
    const newRow = headers.map(h => fieldMap[h] !== undefined ? fieldMap[h] : '');
    sheet.appendRow(newRow);
    writeLog('dashboard', 'ADD_KENDARAAN', 'Tambah kendaraan: ' + payload.noPolisi);
    bumpUpdateToken_(); return { success: true, id, message: 'Kendaraan berhasil ditambahkan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Update kendaraan TANPA token.
 */
function updateKendaraanPublic(token, id, payload) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
    if (!sheet) return { success: false, message: 'Sheet Kendaraan tidak ditemukan.' };

    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const col = name => headers.indexOf(name);

    // Cari baris: coba via ID dulu, fallback via Plat Nomor
    const idColIdx  = headers.indexOf('ID_Kendaraan') !== -1 ? headers.indexOf('ID_Kendaraan') : headers.indexOf('ID');
    const platCols  = ['Plat Nomor','Plat_Nomor','No. Polisi'].map(h => headers.indexOf(h)).filter(i => i !== -1);

    let targetRow = -1;
    if (id) {
      for (let i = 1; i < data.length; i++) {
        if (idColIdx !== -1 && String(data[i][idColIdx]).trim() === String(id).trim()) {
          targetRow = i + 1; break;
        }
      }
    }
    // Fallback: cari via plat nomor jika ID tidak match
    if (targetRow === -1 && payload.noPolisi && platCols.length) {
      for (let i = 1; i < data.length; i++) {
        for (const pc of platCols) {
          if (String(data[i][pc]).trim().toLowerCase() === payload.noPolisi.trim().toLowerCase()) {
            targetRow = i + 1; break;
          }
        }
        if (targetRow !== -1) break;
      }
    }

    if (targetRow === -1) return { success: false, message: 'Kendaraan tidak ditemukan di spreadsheet.' };

    const updates = {
      'Plat_Nomor': payload.noPolisi, 'Plat Nomor': payload.noPolisi, 'No. Polisi': payload.noPolisi,
      'Merk': payload.merk, 'Tipe': payload.tipe, 'Jenis': payload.jenis,
      'Model': payload.model, 'Tahun': payload.tahun, 'Pajak': payload.pajak
    };
    for (const [header, value] of Object.entries(updates)) {
      const c = col(header);
      if (c !== -1 && value !== undefined) sheet.getRange(targetRow, c + 1).setValue(value);
    }
    writeLog('dashboard', 'UPDATE_KENDARAAN', 'Update kendaraan: ' + payload.noPolisi);
    bumpUpdateToken_(); return { success: true, message: 'Kendaraan berhasil diperbarui.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Hapus kendaraan TANPA token.
 */
function deleteKendaraanPublic(token, id) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
    if (!sheet) return { success: false, message: 'Sheet Kendaraan tidak ditemukan.' };

    const data     = sheet.getDataRange().getValues();
    const hdrs     = data[0];
    const idColIdx = hdrs.indexOf('ID_Kendaraan') !== -1 ? hdrs.indexOf('ID_Kendaraan') : 0;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIdx]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        writeLog('dashboard', 'DELETE_KENDARAAN', 'Hapus kendaraan ID: ' + id);
        bumpUpdateToken_(); return { success: true, message: 'Kendaraan berhasil dihapus.' };
      }
    }
    return { success: false, message: 'Kendaraan tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Simpan mobil keluar TANPA token.
 */
function saveMobilKeluarPublic(token, payload) {
  try {
    var session = requireAuth(token);
    if (!payload.platNomor)    return { success: false, message: 'Plat nomor tidak boleh kosong.' };
    if (!payload.tujuan)       return { success: false, message: 'Tujuan tidak boleh kosong.' };
    if (!payload.namaPengguna) return { success: false, message: 'Nama pengguna tidak boleh kosong.' };

    const ss  = getSpreadsheet();
    let sheet = ss.getSheetByName(getConfig('SHEET_MOBIL_KELUAR'));
    if (!sheet) sheet = setupMobilKeluarSheet();

    const lastRow = sheet.getLastRow();
    const id      = 'KLR_' + String(lastRow).padStart(4, '0');

    let kondisi = {};
    try { kondisi = JSON.parse(payload.kondisi || '{}'); } catch(e) {}

    // Key kondisi tanpa prefix "Kondisi"
    const kondisiKeys = [
      'Ban Depan', 'Ban Belakang', 'Body Depan',
      'Body Samping Kanan', 'Body Samping Kiri', 'Body Belakang',
      'Ban Cadangan', 'Perlengkapan (Dongkrak, Kunci, Segitiga)', 'AC', 'Lampu', 'Lain-lain'
    ];

    // Struktur: ID, Tanggal, Waktu Input, Plat Nomor, Jam Keluar, KM, BBM,
    //           Tujuan, Nama Pengguna, Nomor HP, ACC Pimpinan, [kondisi x11 * 2], Catatan, Dicatat Keluar Oleh
    const now = new Date();
    const _fmtTgl = (t) => {
      if (!t) return Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      if (String(t).indexOf('-') !== -1 && String(t).length === 10) { const p = String(t).split('-'); return p[2]+'/'+p[1]+'/'+p[0]; }
      return String(t);
    };
    const row = [
      id,
      _fmtTgl(payload.tanggal),
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
      sanitizeCell(payload.platNomor    || ''),
      sanitizeCell(payload.jamKeluar    || ''),
      sanitizeCell(payload.kmKeluar     || ''),
      sanitizeCell(payload.bbm          || ''),
      sanitizeCell(payload.tujuan       || ''),
      sanitizeCell(payload.namaPengguna || ''),
      sanitizeCell(payload.nomorHp      || ''),
      sanitizeCell(getNamaLengkap(payload.dicatatOleh) || ''),  // Disetujui Oleh = Nama Lengkap user yang login
    ];

    kondisiKeys.forEach(function(key) {
      const item = kondisi[key] || {};
      row.push(sanitizeCell(item.status     || 'Baik'));
      row.push(sanitizeCell(item.keterangan || ''));
    });

    row.push(sanitizeCell(payload.catatan      || ''));
    row.push(sanitizeCell(getNamaLengkap(payload.dicatatOleh) || payload.dicatatOleh || 'dashboard'));

    sheet.appendRow(row);

    // Update status kendaraan → Keluar
    try {
      const kndSheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
      if (kndSheet) {
        const kndData = kndSheet.getDataRange().getValues();
        const hdrs    = kndData[0].map(h => String(h).trim());
        const platCol = (() => {
          for (const c of ['Plat Nomor','Plat_Nomor','No. Polisi']) { const i = hdrs.indexOf(c); if (i !== -1) return i; }
          return -1;
        })();
        const statusCol = hdrs.indexOf('Status');
        if (platCol !== -1 && statusCol !== -1) {
          for (let i = 1; i < kndData.length; i++) {
            if (String(kndData[i][platCol]).trim() === String(payload.platNomor).trim()) {
              kndSheet.getRange(i + 1, statusCol + 1).setValue('Mobil Keluar');
              break;
            }
          }
        }
      }
    } catch(e) { console.error('Update status error:', e); }

    writeLog('dashboard', 'CATAT_KELUAR', 'Catat keluar: ' + payload.platNomor + ' → ' + payload.tujuan);
    bumpUpdateToken_(); return { success: true, id, message: 'Data mobil keluar berhasil disimpan.' };
  } catch (err) {
    return { success: false, message: 'Terjadi kesalahan: ' + err.message };
  }
}

function updateMobilKeluarPublic(token, platNomor, payload) {
  try {
    var session = requireAuth(token);
    if (!platNomor) return { success: false, message: 'Plat nomor tidak boleh kosong.' };

    const ss    = getSpreadsheet();
    let sheet   = ss.getSheetByName(getConfig('SHEET_MOBIL_KELUAR'));
    if (!sheet) return { success: false, message: 'Sheet MobilKeluar tidak ditemukan.' };

    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());

    const platCol  = (() => { for (const c of ['Plat Nomor','Plat_Nomor']) { const i = headers.indexOf(c); if (i !== -1) return i; } return -1; })();
    const tglCol   = headers.indexOf('Tanggal');
    const jamCol   = headers.indexOf('Jam Keluar');
    const kmCol    = headers.indexOf('KM Keluar');
    const bbmCol   = headers.indexOf('BBM (%)') !== -1 ? headers.indexOf('BBM (%)') : headers.indexOf('BBM (L)');
    const accCol   = headers.indexOf('Disetujui Oleh') !== -1 ? headers.indexOf('Disetujui Oleh')
                   : headers.indexOf('Pemberi Akses Keluar') !== -1 ? headers.indexOf('Pemberi Akses Keluar')
                   : headers.indexOf('ACC Pimpinan');
    const catCol   = headers.indexOf('Catatan');
    const wktCol   = headers.indexOf('Waktu Input');
    const dicatatCol = headers.indexOf('Dicatat Keluar Oleh') !== -1
                     ? headers.indexOf('Dicatat Keluar Oleh')
                     : headers.indexOf('Dicatat Oleh');
    const namaCol  = headers.indexOf('Nama Pengguna');
    const hpCol    = headers.indexOf('No. HP Pengguna') !== -1 ? headers.indexOf('No. HP Pengguna') : headers.indexOf('Nomor HP');

    if (platCol === -1) return { success: false, message: 'Kolom Plat Nomor tidak ditemukan.' };

    // Cari baris yang paling baru untuk plat ini (scan dari bawah)
    let targetRow = -1;
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][platCol]).trim().toUpperCase() === String(platNomor).trim().toUpperCase()) {
        targetRow = i; break;
      }
    }
    if (targetRow === -1) return { success: false, message: 'Data keluar untuk ' + platNomor + ' tidak ditemukan.' };

    const rowNum = targetRow + 1;
    const now    = new Date();
    const tz     = Session.getScriptTimeZone();

    // Update kolom operasional
    const setVal = (col, val) => { if (col !== -1) sheet.getRange(rowNum, col+1).setValue(val); };

    // Konversi tanggal dari yyyy-mm-dd ke dd/MM/yyyy
    let tglFormatted = payload.tanggal || Utilities.formatDate(now, tz, 'dd/MM/yyyy');
    if (tglFormatted && tglFormatted.indexOf('-') !== -1 && tglFormatted.length === 10) {
      const parts = tglFormatted.split('-');
      tglFormatted = parts[2] + '/' + parts[1] + '/' + parts[0];
    }

    setVal(tglCol,     tglFormatted);
    setVal(jamCol,     payload.jamKeluar || '');
    setVal(kmCol,      payload.kmKeluar  || '');
    setVal(bbmCol,     payload.bbm       || '');
    setVal(catCol,     payload.catatan   || '');
    setVal(wktCol,     Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm:ss'));
    // Dicatat Keluar Oleh = Nama Lengkap operator yang mengisi (bukan approver)
    setVal(dicatatCol, payload.dicatatOleh || 'dashboard');
    // Update nama & HP jika tersedia dari payload
    if (payload.namaPengguna) setVal(namaCol, payload.namaPengguna);
    if (payload.nomorHp)      setVal(hpCol,   payload.nomorHp);

    // Update kondisi jika ada
    let kondisi = {};
    try { kondisi = JSON.parse(payload.kondisi || '{}'); } catch(e) {}
    const kondisiKeys = [
      'Ban Depan','Ban Belakang','Body Depan',
      'Body Samping Kanan','Body Samping Kiri','Body Belakang',
      'Ban Cadangan','Perlengkapan','AC','Lampu','Lain-lain'
    ];
    kondisiKeys.forEach(function(key) {
      const statusCol = headers.indexOf(key);
      const ketCol    = headers.indexOf('Ket. ' + key);
      const item      = kondisi[key] || {};
      if (statusCol !== -1) sheet.getRange(rowNum, statusCol+1).setValue(item.status     || '');
      if (ketCol    !== -1) sheet.getRange(rowNum, ketCol+1).setValue(item.keterangan || '');
    });

    writeLog('dashboard', 'UPDATE_KELUAR', 'Update keluar: ' + platNomor);

    // Cek apakah ini untuk Servis
    var isServis = false;
    try {
      const pmjSheet = ss.getSheetByName('Peminjaman');
      if (pmjSheet) {
        const pmjData = pmjSheet.getDataRange().getValues();
        const pmjHdr  = pmjData[0].map(h => String(h).trim().toLowerCase());
        const pPlatC  = (function() { for (const c of ['plat nomor','plat_nomor','no. polisi']) { const i = pmjHdr.indexOf(c); if (i !== -1) return i; } return -1; })();
        const pKepC   = pmjHdr.indexOf('keperluan');
        const pKenC   = pmjHdr.indexOf('kendala');
        if (pPlatC !== -1 && pKepC !== -1) {
          for (var pi = pmjData.length - 1; pi >= 1; pi--) {
            if (String(pmjData[pi][pPlatC]).trim().toUpperCase() === String(platNomor).trim().toUpperCase()) {
              var kep = String(pmjData[pi][pKepC] || '').trim().toLowerCase();
              if (kep.indexOf('servis') !== -1) { isServis = true; }
              break;
            }
          }
        }
      }
    } catch(e) { console.error('Servis lookup error:', e); }

    if (isServis) {
      // Servis: set status Kendaraan → "Servis", buat entry MobilServis
      try {
        // Update Kendaraan status ke Servis
        try {
          const kndSheet2 = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
          if (kndSheet2) {
            const kd2    = kndSheet2.getDataRange().getValues();
            const kh2    = kd2[0].map(h => String(h).trim());
            const kpCol2 = (() => { for (const c of ['Plat Nomor','Plat_Nomor','No. Polisi']) { const i = kh2.indexOf(c); if (i !== -1) return i; } return -1; })();
            const ksCol2 = kh2.indexOf('Status');
            if (kpCol2 !== -1 && ksCol2 !== -1) {
              for (let j2 = 1; j2 < kd2.length; j2++) {
                if (String(kd2[j2][kpCol2]).trim().toUpperCase() === String(platNomor).trim().toUpperCase()) {
                  kndSheet2.getRange(j2+1, ksCol2+1).setValue('Servis');
                  break;
                }
              }
            }
          }
        } catch(kErr2) { console.error('Servis: gagal update status kendaraan:', kErr2.message); }

        var kendalaVal = pKenC !== -1 ? String(pmjData[pi][pKenC] || '') : '';
        var srvPayload = {
          platNomor:        platNomor,
          tanggal:          tglFormatted,
          tujuan:           payload.tujuan || '',
          kendala:          kendalaVal,
          namaPengguna:     payload.namaPengguna || '',
          nomorHp:          payload.nomorHp || '',
          dicatatKeluarOleh: payload.dicatatOleh || 'dashboard',
          status:           'Belum Servis',
          updateStatus:     false
        };
        saveMobilServisPublicInternal(ss, srvPayload);
      } catch(e) { console.error('Gagal buat entry MobilServis:', e); }
    } else {
      // Normal: update status kendaraan → "Mobil Keluar"
      try {
        const kndSheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
        if (kndSheet) {
          const kd    = kndSheet.getDataRange().getValues();
          const kh    = kd[0].map(h => String(h).trim());
          const kpCol = (() => { for (const c of ['Plat Nomor','Plat_Nomor','No. Polisi']) { const i = kh.indexOf(c); if (i !== -1) return i; } return -1; })();
          const ksCol = kh.indexOf('Status');
          if (kpCol !== -1 && ksCol !== -1) {
            for (let j = 1; j < kd.length; j++) {
              if (String(kd[j][kpCol]).trim().toUpperCase() === String(platNomor).trim().toUpperCase()) {
                kndSheet.getRange(j+1, ksCol+1).setValue('Mobil Keluar');
                break;
              }
            }
          }
        }
      } catch(kErr) { console.error('updateMobilKeluar: gagal update status kendaraan:', kErr.message); }
    }

    bumpUpdateToken_(); return { success: true, message: 'Data berhasil diperbarui.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

/**
 * Ambil jumlah user TANPA token (untuk stat card).
 */
function getCountUsersPublic(token) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_USERS'));
    if (!sheet) return { success: true, count: 0 };
    const lastRow = sheet.getLastRow();
    return { success: true, count: Math.max(0, lastRow - 1) };
  } catch (err) {
    return { success: true, count: 0 };
  }
}

function getListKendaraan(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: 'Session tidak valid.' };

  try {
    const ss    = getSpreadsheet();
    let sheet   = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));

    // Jika sheet belum ada, buat otomatis
    if (!sheet) {
      setupKendaraanSheet();
      sheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };

    const headers = data[0];

    // Baca baris data, skip baris kosong
    const rows = data.slice(1)
      .filter(row => row.some(cell => String(cell).trim() !== ''))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[String(h).trim()] = row[i]; });
        return obj;
      });

    return { success: true, data: rows };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Tambah kendaraan baru ke sheet Kendaraan.
 * Payload dari Dashboard: { noPolisi, merk, tipe, jenis, model, tahun, status }
 */
function addKendaraan(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: 'Session tidak valid.' };
  if (!['admin', 'operator'].includes(session.role)) {
    return { success: false, message: 'Akses ditolak.' };
  }

  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
    if (!sheet) return { success: false, message: 'Sheet Kendaraan tidak ditemukan.' };

    // Buat ID unik dengan prefix IKD_
    const lastRow = sheet.getLastRow();
    const seq = String(lastRow).padStart(3, '0');
    const id = 'IKD_' + seq;

    // Baca header aktual sheet untuk menentukan urutan kolom
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Map payload ke header (mendukung header lama maupun baru)
    const fieldMap = {
      'ID_Kendaraan': id,
      'ID':           id,
      'Merk':         payload.merk     || '',
      'Tipe':         payload.tipe     || '',
      'Jenis':        payload.jenis    || '',
      'Model':        payload.model    || '',
      'Tahun':        payload.tahun    || '',
      'Plat_Nomor':   payload.noPolisi || '',
      'Plat Nomor':   payload.noPolisi || '',
      'No. Polisi':   payload.noPolisi || '',
      'Status':       payload.status   || 'Tersedia',
      'Pajak':        payload.pajak    || '',
    };

    const newRow = headers.map(h => fieldMap[h] !== undefined ? fieldMap[h] : '');
    sheet.appendRow(newRow);

    writeLog(session.username, 'ADD_KENDARAAN', 'Tambah kendaraan: ' + payload.noPolisi);
    return { success: true, id, message: 'Kendaraan berhasil ditambahkan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Update data kendaraan berdasarkan ID.
 * Payload dari Dashboard: { noPolisi, merk, tipe, jenis, model, tahun, status }
 */
function updateKendaraan(token, id, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: 'Session tidak valid.' };
  if (!['admin', 'operator'].includes(session.role)) {
    return { success: false, message: 'Akses ditolak.' };
  }

  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
    if (!sheet) return { success: false, message: 'Sheet Kendaraan tidak ditemukan.' };

    const data    = sheet.getDataRange().getValues();
    const headers = data[0];

    // Cari kolom berdasarkan header
    const col = (name) => headers.indexOf(name);

    // Cari kolom ID (mendukung ID_Kendaraan atau ID)
    const idColIdx = headers.indexOf('ID_Kendaraan') !== -1 ? headers.indexOf('ID_Kendaraan') : headers.indexOf('ID');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIdx]).trim() === String(id).trim()) {
        const rowNum = i + 1; // 1-based untuk sheet.getRange

        // Mendukung header baru (Plat_Nomor) maupun lama (No. Polisi)
        const updates = {
          'Plat_Nomor': payload.noPolisi,
          'Plat Nomor': payload.noPolisi,
          'No. Polisi': payload.noPolisi,
          'Merk':       payload.merk,
          'Tipe':       payload.tipe,
          'Jenis':      payload.jenis,
          'Model':      payload.model,
          'Tahun':      payload.tahun,
          'Status':     payload.status
        };

        for (const [header, value] of Object.entries(updates)) {
          const c = col(header);
          if (c !== -1 && value !== undefined) {
            sheet.getRange(rowNum, c + 1).setValue(value);
          }
        }

        writeLog(session.username, 'UPDATE_KENDARAAN', 'Update kendaraan ID: ' + id + ', No. Polisi: ' + payload.noPolisi);
        return { success: true, message: 'Kendaraan berhasil diperbarui.' };
      }
    }

    return { success: false, message: 'Kendaraan dengan ID ' + id + ' tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Hapus kendaraan berdasarkan ID.
 */
function deleteKendaraan(token, id) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: 'Session tidak valid.' };
  if (session.role !== 'admin') {
    return { success: false, message: 'Akses ditolak. Hanya admin yang bisa menghapus kendaraan.' };
  }

  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
    if (!sheet) return { success: false, message: 'Sheet Kendaraan tidak ditemukan.' };

    const data    = sheet.getDataRange().getValues();
    const hdrs    = data[0];
    const idColIdx = hdrs.indexOf('ID_Kendaraan') !== -1 ? hdrs.indexOf('ID_Kendaraan') : 0;
    const platColIdx = (() => {
      const candidates = ['Plat Nomor', 'Plat_Nomor', 'No. Polisi', 'No Polisi', 'Plat'];
      for (const c of candidates) { const i = hdrs.indexOf(c); if (i !== -1) return i; }
      return 1;
    })();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIdx]).trim() === String(id).trim()) {
        const noPolisi = data[i][platColIdx] || id;
        sheet.deleteRow(i + 1);
        writeLog(session.username, 'DELETE_KENDARAAN', 'Hapus kendaraan ID: ' + id + ', No. Polisi: ' + noPolisi);
        return { success: true, message: 'Kendaraan berhasil dihapus.' };
      }
    }

    return { success: false, message: 'Kendaraan dengan ID ' + id + ' tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Ambil daftar semua user (untuk stat card "Jumlah User").
 * Hanya admin yang boleh melihat data lengkap; role lain hanya dapat jumlah.
 */
function getListUsers(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: 'Session tidak valid.' };

  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_USERS'));
    if (!sheet) return { success: true, data: [] };

    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows    = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        // Sembunyikan kolom password untuk semua role
        if (String(h).toLowerCase() !== 'password') {
          obj[h] = row[i];
        }
      });
      return obj;
    });

    return { success: true, data: rows };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ============================================================
//  FUNGSI MOBIL KELUAR
// ============================================================

/**
 * Setup sheet MobilKeluar: buat jika belum ada.
 */
function setupMobilKeluarSheet() {
  const ss = getSpreadsheet();
  // Struktur:
  // A=ID, B=Tanggal, C=Waktu Input, D=Plat Nomor, E=Jam Keluar, F=KM Keluar, G=BBM (L),
  // H=Tujuan, I=Nama Pengguna, J=Nomor HP, K=ACC Pimpinan,
  // L..AF = kondisi + keterangan, AG=Catatan, AH=Dicatat Keluar Oleh
  const HEADERS = [
    'ID', 'Tanggal', 'Waktu Input', 'Plat Nomor', 'Jam Keluar', 'KM Keluar', 'BBM (%)',
    'Tujuan', 'Nama Pengguna', 'No. HP Pengguna', 'Disetujui Oleh',
    'Ban Depan', 'Ket. Ban Depan',
    'Ban Belakang', 'Ket. Ban Belakang',
    'Body Depan', 'Ket. Body Depan',
    'Body Samping Kanan', 'Ket. Body Samping Kanan',
    'Body Samping Kiri', 'Ket. Body Samping Kiri',
    'Body Belakang', 'Ket. Body Belakang',
    'Ban Cadangan', 'Ket. Ban Cadangan',
    'Perlengkapan', 'Ket. Perlengkapan',
    'AC', 'Ket. AC',
    'Lampu', 'Ket. Lampu',
    'Lain-lain', 'Ket. Lain-lain',
    'Catatan', 'Dicatat Keluar Oleh'
  ];

  let sheet = ss.getSheetByName(getConfig('SHEET_MOBIL_KELUAR'));
  if (!sheet) {
    sheet = ss.insertSheet(getConfig('SHEET_MOBIL_KELUAR'));
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    formatHeaderRow(sheet);
    applyZebraRows(sheet);
    sheet.autoResizeColumns(1, HEADERS.length);
  }
  return sheet;
}

/**
 * Simpan data mobil keluar baru.
 * Dipanggil dari frontend dengan payload lengkap.
 */
function saveMobilKeluar(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: 'Session tidak valid.' };
  if (!['admin', 'operator'].includes(session.role)) {
    return { success: false, message: 'Akses ditolak.' };
  }

  try {
    if (!payload.platNomor) return { success: false, message: 'Plat nomor tidak boleh kosong.' };
    if (!payload.tujuan)    return { success: false, message: 'Tujuan tidak boleh kosong.' };
    if (!payload.namaPengguna) return { success: false, message: 'Nama pengguna tidak boleh kosong.' };

    const ss    = getSpreadsheet();
    let sheet   = ss.getSheetByName(getConfig('SHEET_MOBIL_KELUAR'));
    if (!sheet) sheet = setupMobilKeluarSheet();

    // Generate ID unik
    const lastRow = sheet.getLastRow();
    const id = 'KLR_' + String(lastRow).padStart(4, '0');

    // Parse kondisi dari JSON string
    let kondisi = {};
    try { kondisi = JSON.parse(payload.kondisi || '{}'); } catch(e) {}

    const kondisiKeys = [
      'Ban Depan', 'Ban Belakang', 'Body Depan',
      'Body Samping Kanan', 'Body Samping Kiri', 'Body Belakang',
      'Ban Cadangan', 'Perlengkapan (Dongkrak, Kunci, Segitiga)', 'AC', 'Lampu', 'Lain-lain'
    ];

    // Bangun baris data sesuai urutan header
    const row = [
      id,
      payload.tanggal    || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'), // Waktu Input
      payload.platNomor  || '',
      payload.jamKeluar  || '',
      payload.kmKeluar   || '',
      payload.bbm        || '',
      payload.tujuan     || '',
      payload.namaPengguna || '',
      payload.nomorHp    || '',
      payload.jabatan    || '',
      payload.accPimpinan || '',
    ];

    // Tambah kolom kondisi (status + keterangan per item)
    kondisiKeys.forEach(function(key) {
      const item = kondisi[key] || {};
      row.push(item.status      || 'Baik');
      row.push(item.keterangan  || '');
    });

    row.push(payload.catatan   || '');
    row.push(session.nama      || session.username);
    row.push(new Date());

    sheet.appendRow(row);

    // Update status kendaraan jadi "Keluar" di sheet Kendaraan
    try {
      const kndSheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
      if (kndSheet) {
        const kndData = kndSheet.getDataRange().getValues();
        const hdrs    = kndData[0];
        const platCol = (() => {
          const candidates = ['Plat Nomor', 'Plat_Nomor', 'No. Polisi'];
          for (const c of candidates) { const i = hdrs.indexOf(c); if (i !== -1) return i; }
          return -1;
        })();
        const statusCol = hdrs.indexOf('Status');
        if (platCol !== -1 && statusCol !== -1) {
          for (let i = 1; i < kndData.length; i++) {
            if (String(kndData[i][platCol]).trim() === String(payload.platNomor).trim()) {
              kndSheet.getRange(i + 1, statusCol + 1).setValue('Mobil Keluar');
              break;
            }
          }
        }
      }
    } catch(e) {
      console.error('Update status kendaraan error:', e);
    }

    writeLog(session.username, 'CATAT_KELUAR', 'Catat keluar: ' + payload.platNomor + ' → ' + payload.tujuan);
    return { success: true, id, message: 'Data mobil keluar berhasil disimpan.' };

  } catch (err) {
    console.error('saveMobilKeluar error:', err);
    return { success: false, message: 'Terjadi kesalahan: ' + err.message };
  }
}

/**
 * Ambil daftar semua data mobil keluar.
 */
function getListMobilKeluar(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: 'Session tidak valid.' };

  try {
    const ss  = getSpreadsheet();
    let sheet = ss.getSheetByName(getConfig('SHEET_MOBIL_KELUAR'));
    if (!sheet) return { success: true, data: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };

    const headers = data[0];
    const rows = data.slice(1)
      .filter(row => row.some(cell => String(cell).trim() !== ''))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[String(h).trim()] = row[i]; });
        return obj;
      });

    return { success: true, data: rows };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Hapus data mobil keluar berdasarkan ID.
 */
function deleteMobilKeluar(token, id) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: 'Session tidak valid.' };
  if (!['admin', 'operator'].includes(session.role)) {
    return { success: false, message: 'Akses ditolak.' };
  }

  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_MOBIL_KELUAR'));
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan.' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        writeLog(session.username, 'DELETE_KELUAR', 'Hapus data keluar ID: ' + id);
        bumpUpdateToken_(); return { success: true, message: 'Data berhasil dihapus.' };
      }
    }
    return { success: false, message: 'Data dengan ID ' + id + ' tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ============================================================
//  TEMPLATE INCLUDE — untuk menyisipkan file HTML lain
//  Contoh pemakaian di file HTML:
//      <?!= include('MobileFrame') ?>
// ============================================================

/**
 * Ambil daftar mobil keluar TANPA token.
 */
function getListMobilKeluarPublic(token) {
  try {
    var session = requireAuth(token);
    const ss  = getSpreadsheet();
    let sheet = ss.getSheetByName(getConfig('SHEET_MOBIL_KELUAR'));
    if (!sheet) return { success: true, data: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };

    const headers = data[0].map(h => String(h).trim());

    // ── Lookup 1: Plat → Nomor HP Pengemudi dari sheet Peminjaman ──
    const hpByPlat = {};
    try {
      const pmjSheet = ss.getSheetByName('Peminjaman');
      if (pmjSheet) {
        const pmjData = pmjSheet.getDataRange().getValues();
        const pmjHdr  = pmjData[0].map(h => String(h).trim());
        const pPlatCol = (() => { for (const c of ['Plat Nomor','Plat_Nomor','No. Polisi']) { const i = pmjHdr.indexOf(c); if (i !== -1) return i; } return -1; })();
        const pHpCol   = (() => { for (const c of ['No. HP Pengguna','No. HP Pengemudi','Telp Pengemudi','HP Pengemudi','Nomor HP Pengemudi']) { const i = pmjHdr.indexOf(c); if (i !== -1) return i; } return -1; })();
        if (pPlatCol !== -1 && pHpCol !== -1) {
          for (let i = 1; i < pmjData.length; i++) {
            const p = String(pmjData[i][pPlatCol]).trim().toUpperCase();
            const h = String(pmjData[i][pHpCol]).trim();
            if (p && h) hpByPlat[p] = h;
          }
        }
      }
    } catch(e) { console.error('HP lookup error:', e); }

    // ── Lookup 2: Username/Nama → Nama Lengkap dari sheet Users ──
    const namaByUsername = {}; // username (lowercase) → Nama Pengguna
    const namaByNama     = {}; // nama (lowercase) → Nama Pengguna (self-lookup)
    try {
      const usersSheet = ss.getSheetByName(getConfig('SHEET_USERS'));
      if (usersSheet) {
        const ud   = usersSheet.getDataRange().getValues();
        const uh   = ud[0].map(h => String(h).trim());
        const uNamaCol = (() => { for (const c of ['Nama Pengguna','Nama Lengkap','Nama']) { const i = uh.indexOf(c); if (i !== -1) return i; } return -1; })();
        const uUserCol = uh.indexOf('Username') !== -1 ? uh.indexOf('Username') : uh.indexOf('username');
        for (let i = 1; i < ud.length; i++) {
          const un   = uUserCol !== -1 ? String(ud[i][uUserCol]).trim().toLowerCase() : '';
          const nm   = uNamaCol !== -1 ? String(ud[i][uNamaCol]).trim() : '';
          if (un && nm) namaByUsername[un] = nm;
          if (nm)       namaByNama[nm.toLowerCase()] = nm;
        }
        // Tambahkan Super Admin dari Script Properties
        namaByUsername[getConfig('ADMIN_USERNAME')] = getConfig('ADMIN_NAME');
      }
    } catch(e) { console.error('Users lookup error:', e); }

    // ── Lookup 3: Plat → Lokasi Tujuan dari sheet Peminjaman ──
    const tujuanByPlat = {};
    try {
      const pmjSheet2 = ss.getSheetByName('Peminjaman');
      if (pmjSheet2) {
        const pd  = pmjSheet2.getDataRange().getValues();
        const ph  = pd[0].map(h => String(h).trim());
        const ppCol = (() => { for (const c of ['Plat Nomor','Plat_Nomor']) { const i = ph.indexOf(c); if (i !== -1) return i; } return -1; })();
        const ptCol = ph.indexOf('Lokasi Tujuan');
        if (ppCol !== -1 && ptCol !== -1) {
          for (let i = 1; i < pd.length; i++) {
            const p = String(pd[i][ppCol]).trim().toUpperCase();
            const t = String(pd[i][ptCol]).trim();
            if (p && t) tujuanByPlat[p] = t;
          }
        }
      }
    } catch(e) { console.error('Tujuan lookup error:', e); }

    const tujColIdx  = headers.indexOf('Tujuan');
    const platColIdx = (() => { for (const c of ['Plat Nomor','Plat_Nomor']) { const i = headers.indexOf(c); if (i !== -1) return i; } return -1; })();
    const hpColIdx   = headers.indexOf('No. HP Pengguna') !== -1 ? headers.indexOf('No. HP Pengguna') : headers.indexOf('Nomor HP');
    const accColIdx  = headers.indexOf('Disetujui Oleh') !== -1 ? headers.indexOf('Disetujui Oleh')
                     : headers.indexOf('Pemberi Akses Keluar') !== -1 ? headers.indexOf('Pemberi Akses Keluar')
                     : headers.indexOf('ACC Pimpinan');

    const rows = data.slice(1)
      .filter(row => row.some(cell => String(cell).trim() !== ''))
      .map(row => {
        const obj = {};
        const bbmIdx2 = (() => { for (const c of ['BBM (%)','BBM (L)','BBM']) { const i = headers.indexOf(c); if (i !== -1) return i; } return -1; })();
        const tz2 = Session.getScriptTimeZone();
        headers.forEach((h, i) => {
          const v = row[i]; const lh = h.toLowerCase();
          if (v instanceof Date) {
            const yr = v.getFullYear();
            if (lh.indexOf('jam') !== -1) obj[h] = Utilities.formatDate(v, tz2, 'HH:mm');
            else if (h === 'Tanggal' || lh.indexOf('tanggal') !== -1 || lh.indexOf('tgl') !== -1) obj[h] = yr > 1900 ? Utilities.formatDate(v, tz2, 'dd/MM/yyyy') : '';
            else obj[h] = yr > 1900 ? Utilities.formatDate(v, tz2, 'dd/MM/yyyy HH:mm') : '';
          } else if (i === bbmIdx2 && typeof v === 'number') {
            obj[h] = (v > 0 && v <= 1) ? Math.round(v * 100) + '%' : (v > 1 ? Math.round(v) + '%' : String(v));
          } else { obj[h] = (v !== null && v !== undefined) ? String(v) : ''; }
        });

        const plat = platColIdx !== -1 ? String(row[platColIdx]).trim().toUpperCase() : '';
        const existingHp = hpColIdx !== -1 ? String(row[hpColIdx]).trim() : (String(obj['Nomor HP']||'').trim());
        if (!existingHp && plat && hpByPlat[plat]) {
          obj['No. HP Pengguna'] = hpByPlat[plat];
          obj['Nomor HP']        = hpByPlat[plat]; // backward compat
        }

        // Enrich Tujuan dari Peminjaman jika kolom kosong
        const existingTuj = tujColIdx !== -1 ? String(row[tujColIdx]).trim() : '';
        if (!existingTuj && plat && tujuanByPlat[plat]) {
          obj['Tujuan'] = tujuanByPlat[plat];
        }

        // Enrich Disetujui Oleh → Nama Lengkap dari Users
        const accKey = headers[accColIdx] || 'Disetujui Oleh';
        if (accColIdx !== -1) {
          const rawAcc = String(row[accColIdx]).trim();
          if (rawAcc) {
            const byUser = namaByUsername[rawAcc.toLowerCase()];
            const byNama = namaByNama[rawAcc.toLowerCase()];
            obj[accKey]            = byUser || byNama || rawAcc;
            obj['Disetujui Oleh']  = byUser || byNama || rawAcc; // always set standard key
          }
        }

        return obj;
      });

    rows.reverse();
    return { success: true, data: rows };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Hapus catatan mobil keluar TANPA token.
 */
function deleteMobilKeluarPublic(token, id) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(getConfig('SHEET_MOBIL_KELUAR'));
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan.' };

    const data     = sheet.getDataRange().getValues();
    const headers  = data[0].map(h => String(h).trim());
    const idColIdx = headers.indexOf('ID') !== -1 ? headers.indexOf('ID') : 0;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIdx]).trim() === String(id).trim()) {
        // Kembalikan status kendaraan ke Tersedia
        try {
          const platCol = headers.indexOf('Plat Nomor') !== -1 ? headers.indexOf('Plat Nomor') : headers.indexOf('Plat_Nomor');
          if (platCol !== -1) {
            const plat = data[i][platCol];
            const kndSheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
            if (kndSheet && plat) {
              const kndData = kndSheet.getDataRange().getValues();
              const kHdrs   = kndData[0].map(h => String(h).trim());
              const kPlatCol   = kHdrs.indexOf('Plat Nomor') !== -1 ? kHdrs.indexOf('Plat Nomor') : kHdrs.indexOf('Plat_Nomor');
              const kStatusCol = kHdrs.indexOf('Status');
              for (let j = 1; j < kndData.length; j++) {
                if (String(kndData[j][kPlatCol]).trim() === String(plat).trim()) {
                  kndSheet.getRange(j+1, kStatusCol+1).setValue('Tersedia');
                  break;
                }
              }
            }
          }
        } catch(e) {}
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Catatan berhasil dihapus.' };
      }
    }
    return { success: false, message: 'Data tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function setupMobilServisSheet() {
  const ss = getSpreadsheet();
  const HEADERS = [
    'ID', 'Keluar ID', 'Tanggal Masuk', 'Waktu Input', 'Plat Nomor', 'Jam Masuk', 'KM Masuk', 'BBM (%)',
    'Lokasi Tujuan', 'Kendala', 'Nama Pengguna', 'NIP Pengguna', 'No. HP Pengguna', 'Dicatat Keluar Oleh', 'Dicatat Masuk Oleh',
    'Ban Depan', 'Ket. Ban Depan',
    'Ban Belakang', 'Ket. Ban Belakang',
    'Body Depan', 'Ket. Body Depan',
    'Body Samping Kanan', 'Ket. Body Samping Kanan',
    'Body Samping Kiri', 'Ket. Body Samping Kiri',
    'Body Belakang', 'Ket. Body Belakang',
    'Ban Cadangan', 'Ket. Ban Cadangan',
    'Perlengkapan', 'Ket. Perlengkapan',
    'AC', 'Ket. AC',
    'Lampu', 'Ket. Lampu',
    'Lain-lain',
    'Catatan', 'Status'
  ];
  let sheet = ss.getSheetByName('MobilServis');
  if (!sheet) {
    sheet = ss.insertSheet('MobilServis');
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    formatHeaderRow(sheet);
    applyZebraRows(sheet);
    sheet.autoResizeColumns(1, HEADERS.length);
  } else {
    // Add any missing columns to existing sheet
    var existingHdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues()[0].map(h => String(h).trim());
    var missing = HEADERS.filter(function(h) { return existingHdrs.indexOf(h) === -1; });
    if (missing.length) {
      var col = existingHdrs.length + 1;
      missing.forEach(function(h) {
        sheet.getRange(1, col).setValue(h);
        col++;
      });
    }
  }
  return sheet;
}

function saveMobilServisPublic(token, payload) {
  try {
    var session = requireAuth(token);
    if (!payload.platNomor) return { success: false, message: 'Plat nomor tidak boleh kosong.' };

    const ss  = getSpreadsheet();
    let sheet = ss.getSheetByName('MobilServis');
    if (!sheet) sheet = setupMobilServisSheet();

    const totalCol  = sheet.getLastColumn();
    const headerRow = sheet.getRange(1, 1, 1, totalCol).getValues()[0];
    const colMap = {};
    headerRow.forEach(function(h, i) { colMap[String(h).trim()] = i; });

    var rowWidth = totalCol;
    var row = new Array(rowWidth).fill('');

    function set(h, v) {
      var idx = colMap[h];
      if (idx !== undefined) row[idx] = sanitizeCell(v);
    }

    const lastRow = sheet.getLastRow();
    const id      = 'SRV_' + String(lastRow).padStart(4, '0');
    const now     = new Date();
    const tz      = Session.getScriptTimeZone();

    set('ID', id);
    set('Tanggal', payload.tanggal || payload.tglServis || Utilities.formatDate(now, tz, 'dd/MM/yyyy'));
    set('Waktu Input', Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm:ss'));
    set('Plat Nomor', payload.platNomor || '');
    set('KM', payload.km || payload.kmMasuk || '');
    set('BBM (%)', payload.bbm || '');
    set('Tujuan', payload.tujuan || '');
    set('Nama Pengguna', payload.namaPengguna || '');
    set('No. HP Pengguna', payload.nomorHp || payload.hpPengguna || '');
    set('Dicatat Keluar Oleh', payload.dicatatKeluarOleh || '');
    set('Dicatat Masuk Oleh', payload.dicatatOleh || payload.dicatatMasukOleh || session.nama || 'dashboard');
    set('Status', payload.status || 'Selesai');

    sheet.appendRow(row);

    // Update status kendaraan → Servis (jika dari Catat Masuk servis) atau biarkan
    if (payload.updateStatus !== false) {
      try {
        const kndSheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
        if (kndSheet) {
          const kd      = kndSheet.getDataRange().getValues();
          const kh      = kd[0].map(h => String(h).trim().toLowerCase());
          const kPlat   = (function() { for (const c of ['plat nomor','plat_nomor','no. polisi']) { const i = kh.indexOf(c); if (i !== -1) return i; } return -1; })();
          const kStatus = kh.indexOf('status');
          if (kPlat !== -1 && kStatus !== -1) {
            for (let j = 1; j < kd.length; j++) {
              if (String(kd[j][kPlat]).trim().toUpperCase() === String(payload.platNomor).trim().toUpperCase()) {
                kndSheet.getRange(j+1, kStatus+1).setValue(payload.statusKendaraan || 'Servis');
                break;
              }
            }
          }
        }
      } catch(e) { console.error('Update status Servis error:', e); }
    }

    writeLog(session.username,'SERVIS','Servis: '+payload.platNomor);
    bumpUpdateToken_(); return { success: true, id, message: 'Data servis berhasil disimpan.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}
function getListMobilServisPublic(token) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('MobilServis');
    if (!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1)
      .filter(row => row.some(cell => String(cell).trim() !== ''))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          obj[h] = (row[i] instanceof Date)
            ? Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'dd/MM/yyyy')
            : row[i];
        });
        return obj;
      });
    rows.reverse();
    return { success: true, data: rows };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function deleteMobilServisPublic(token, id) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('MobilServis');
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan.' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        writeLog('dashboard', 'DELETE_SERVIS', 'Hapus servis ID: ' + id);
        bumpUpdateToken_(); return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function setupMobilMasukSheet() {
  const ss = getSpreadsheet();
  const HEADERS = [
    'ID', 'Keluar ID', 'Tanggal Masuk', 'Waktu Input', 'Plat Nomor', 'Jam Masuk', 'KM Masuk', 'BBM (%)',
    'Lokasi Tujuan', 'Nama Pengguna', 'NIP Pengguna', 'No. HP Pengguna', 'Dicatat Keluar Oleh', 'Dicatat Masuk Oleh',
    'Ban Depan', 'Ket. Ban Depan',
    'Ban Belakang', 'Ket. Ban Belakang',
    'Body Depan', 'Ket. Body Depan',
    'Body Samping Kanan', 'Ket. Body Samping Kanan',
    'Body Samping Kiri', 'Ket. Body Samping Kiri',
    'Body Belakang', 'Ket. Body Belakang',
    'Ban Cadangan', 'Ket. Ban Cadangan',
    'Perlengkapan', 'Ket. Perlengkapan',
    'AC', 'Ket. AC',
    'Lampu', 'Ket. Lampu',
    'Lain-lain',
    'Catatan'
  ];
  let sheet = ss.getSheetByName('MobilMasuk');
  if (!sheet) {
    sheet = ss.insertSheet('MobilMasuk');
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    formatHeaderRow(sheet);
    applyZebraRows(sheet);
    sheet.autoResizeColumns(1, HEADERS.length);
  }
  return sheet;
}

/**
 * Simpan data mobil masuk DAN update status kendaraan → Tersedia.
 * Fix 3: case close — baris keluar tidak diubah, status kendaraan cukup
 * dikembalikan ke Tersedia. Kalau keluar lagi, tercatat di baris baru.
 */
function saveMobilMasukPublic(token, payload) {
  try {
    var session = requireAuth(token);
    if (!payload.platNomor)    return { success: false, message: 'Plat nomor tidak boleh kosong.' };
    if (!payload.namaPengguna) return { success: false, message: 'Nama pengguna tidak boleh kosong.' };

    const ss  = getSpreadsheet();

    // ── Normal flow: simpan ke MobilMasuk ──
    let sheet = ss.getSheetByName('MobilMasuk');
    if (!sheet) sheet = setupMobilMasukSheet();

    const lastRow = sheet.getLastRow();
    const id      = 'MSK_' + String(lastRow).padStart(4, '0');

    let kondisi = {};
    try { kondisi = JSON.parse(payload.kondisi || '{}'); } catch(e) {}

    const kondisiKeys = [
      'Ban Depan','Ban Belakang','Body Depan',
      'Body Samping Kanan','Body Samping Kiri','Body Belakang',
      'Ban Cadangan','Perlengkapan','AC','Lampu','Lain-lain'
    ];

    const now = new Date();

    // Baca header aktual dari sheet, bukan hardcoded
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues()[0]
      .map(h => String(h).trim());

    // ── Pastikan kolom 'Keluar ID' ada ──────────────────────────
    if (headerRow.indexOf('Keluar ID') === -1) {
      const nCol = sheet.getLastColumn() || headerRow.length;
      sheet.getRange(1, nCol + 1).setValue('Keluar ID');
      headerRow.push('Keluar ID');
    }

    const colMap = {};
    headerRow.forEach((h, i) => { colMap[h] = i; });

    // Buat row array sejumlah header, isi default ''
    const row = new Array(headerRow.length).fill('');

    const set = (headerName, value) => {
      const idx = colMap[headerName];
      if (idx !== undefined) row[idx] = sanitizeCell(value);
    };

    set('ID', id);
    set('Keluar ID', payload.keluarId || '');
    set('Tanggal Masuk', payload.tanggal || Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy'));
    set('Tanggal', payload.tanggal || Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy'));
    set('Waktu Input', Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
    set('Plat Nomor', payload.platNomor || '');
    set('Jam Masuk', payload.jamMasuk || '');
    set('KM Masuk', payload.kmMasuk || '');
    set('BBM (%)', payload.bbm || '');
    set('Lokasi Tujuan', payload.tujuan || '');
    set('Tujuan', payload.tujuan || '');
    set('Nama Pengguna', payload.namaPengguna || '');
    set('NIP Pengguna', '');
    set('No. HP Pengguna', payload.nomorHp || '');
    set('Dicatat Keluar Oleh', payload.dicatatKeluarOleh || '');
    set('Dicatat Masuk Oleh', getNamaLengkap(payload.dicatatOleh) || payload.dicatatOleh || 'dashboard');
    set('Catatan', payload.catatan || '');

    kondisiKeys.forEach(function(key) {
      const item = kondisi[key] || {};
      set(key, item.status || 'Baik');
      set('Ket. ' + key, item.keterangan || '');
    });

    sheet.appendRow(row);

    // Fix 3: kembalikan status kendaraan → Tersedia
    try {
      const kndSheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
      if (kndSheet) {
        const kndData = kndSheet.getDataRange().getValues();
        const hdrs    = kndData[0].map(h => String(h).trim());
        const platCol = (() => {
          for (const c of ['Plat Nomor','Plat_Nomor','No. Polisi']) { const i = hdrs.indexOf(c); if (i !== -1) return i; }
          return -1;
        })();
        const statusCol = hdrs.indexOf('Status');
        if (platCol !== -1 && statusCol !== -1) {
          for (let i = 1; i < kndData.length; i++) {
            if (String(kndData[i][platCol]).trim() === String(payload.platNomor).trim()) {
              kndSheet.getRange(i + 1, statusCol + 1).setValue('Tersedia');
              break;
            }
          }
        }
      }
    } catch(e) { console.error('Update status kendaraan error:', e); }

    writeLog('dashboard', 'CATAT_MASUK', 'Catat masuk: ' + payload.platNomor);
    bumpUpdateToken_(); return { success: true, id, message: 'Data mobil masuk berhasil disimpan.' };
  } catch(err) {
    return { success: false, message: 'Terjadi kesalahan: ' + err.message };
  }
}

/**
 * Internal version of saveMobilServisPublic — dipanggil dari saveMobilMasukPublic
 * tanpa validasi token (sudah divalidasi di caller).
 */
function saveMobilServisPublicInternal(ss, payload) {
  try {
    // Helper untuk update baris yang sudah ada (jika keluarId diberikan)
    if (payload.keluarId) {
      let sheet = ss.getSheetByName('MobilServis');
      if (sheet) {
        const data    = sheet.getDataRange().getValues();
        const headers = data[0].map(h => String(h).trim());
        const kidCol  = headers.indexOf('Keluar ID');
        if (kidCol !== -1) {
          for (var ri = 1; ri < data.length; ri++) {
            if (String(data[ri][kidCol]).trim() === String(payload.keluarId).trim()) {
              const rowNum = ri + 1;
              const colMap = {};
              headers.forEach(function(h, i) { colMap[h] = i; });
              function upd(h, v) { var idx = colMap[h]; if (idx !== undefined) sheet.getRange(rowNum, idx+1).setValue(sanitizeCell(v)); }
              upd('Jam Masuk', payload.jamMasuk || '');
              upd('KM Masuk', payload.km || payload.kmMasuk || '');
              upd('BBM (%)', payload.bbm || '');
              upd('Dicatat Masuk Oleh', payload.dicatatOleh || payload.dicatatMasukOleh || '');
              upd('Status', payload.status || 'Selesai');
              return { success: true, id: data[ri][0], message: 'Data servis berhasil diperbarui.' };
            }
          }
        }
      }
    }

    let sheet = ss.getSheetByName('MobilServis');
    if (!sheet) sheet = setupMobilServisSheet();

    const totalCol  = sheet.getLastColumn();
    const headerRow = sheet.getRange(1, 1, 1, totalCol).getValues()[0];
    const colMap = {};
    headerRow.forEach(function(h, i) { colMap[String(h).trim()] = i; });

    var rowWidth = totalCol;
    var row = new Array(rowWidth).fill('');

    function set(h, v) {
      var idx = colMap[h];
      if (idx !== undefined) row[idx] = sanitizeCell(v);
    }

    const lastRow = sheet.getLastRow();
    const id      = 'SRV_' + String(lastRow).padStart(4, '0');
    const now     = new Date();
    const tz      = Session.getScriptTimeZone();

    set('ID', id);
    set('Keluar ID', payload.keluarId || '');
    set('Tanggal Masuk', payload.tanggal || payload.tglServis || Utilities.formatDate(now, tz, 'dd/MM/yyyy'));
    set('Waktu Input', Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm:ss'));
    set('Plat Nomor', payload.platNomor || '');
    set('Jam Masuk', payload.jamMasuk || '');
    set('KM Masuk', payload.km || payload.kmMasuk || '');
    set('BBM (%)', payload.bbm || '');
    set('Lokasi Tujuan', payload.tujuan || '');
    set('Kendala', payload.kendala || '');
    set('Nama Pengguna', payload.namaPengguna || '');
    set('No. HP Pengguna', payload.nomorHp || payload.hpPengguna || '');
    set('Dicatat Keluar Oleh', payload.dicatatKeluarOleh || '');
    set('Dicatat Masuk Oleh', payload.dicatatOleh || payload.dicatatMasukOleh || '');
    set('Status', payload.status || 'Selesai');

    sheet.appendRow(row);

    // Update status kendaraan → Servis
    try {
      const kndSheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
      if (kndSheet) {
        const kd      = kndSheet.getDataRange().getValues();
        const kh      = kd[0].map(h => String(h).trim().toLowerCase());
        const kPlat   = (function() { for (const c of ['plat nomor','plat_nomor','no. polisi']) { const i = kh.indexOf(c); if (i !== -1) return i; } return -1; })();
        const kStatus = kh.indexOf('status');
        if (kPlat !== -1 && kStatus !== -1) {
          for (let j = 1; j < kd.length; j++) {
            if (String(kd[j][kPlat]).trim().toUpperCase() === String(payload.platNomor).trim().toUpperCase()) {
              kndSheet.getRange(j+1, kStatus+1).setValue(payload.statusKendaraan || 'Servis');
              break;
            }
          }
        }
      }
    } catch(e) { console.error('Update status Servis error:', e); }

    return { success: true, id, message: 'Data servis berhasil disimpan.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function updateMobilServisSelesaiPublic(token, platNomor, km, bbm) {
  try {
    var session = requireAuth(token);
    if (!platNomor) return { success: false, message: 'Plat nomor tidak boleh kosong.' };
    if (!km) return { success: false, message: 'KM wajib diisi.' };
    if (!bbm) return { success: false, message: 'BBM wajib diisi.' };

    const ss    = getSpreadsheet();
    let sheet   = ss.getSheetByName('MobilServis');
    if (!sheet) return { success: false, message: 'Sheet MobilServis tidak ditemukan.' };

    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const platCol = (() => { for (const c of ['Plat Nomor','Plat_Nomor']) { const i = headers.indexOf(c); if (i !== -1) return i; } return -1; })();
    const stCol   = headers.indexOf('Status');
    const kmCol   = headers.indexOf('KM Masuk');
    const bbmCol  = headers.indexOf('BBM (%)') !== -1 ? headers.indexOf('BBM (%)') : headers.indexOf('BBM');
    const jamCol  = headers.indexOf('Waktu Input');
    const dicCol  = headers.indexOf('Dicatat Masuk Oleh');

    if (platCol === -1) return { success: false, message: 'Kolom Plat Nomor tidak ditemukan.' };

    // Cari baris untuk plat ini (paling baru dengan status "Belum Servis")
    let targetRow = -1;
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][platCol]).trim().toUpperCase() === String(platNomor).trim().toUpperCase()) {
        var st = String(data[i][stCol] || '').trim().toLowerCase();
        if (st.indexOf('belum servis') !== -1 || st.indexOf('belum') !== -1) {
          targetRow = i; break;
        }
      }
    }
    if (targetRow === -1) {
      // Fallback: ambil baris terakhir untuk plat ini
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][platCol]).trim().toUpperCase() === String(platNomor).trim().toUpperCase()) {
          targetRow = i; break;
        }
      }
    }
    if (targetRow === -1) return { success: false, message: 'Data servis untuk ' + platNomor + ' tidak ditemukan.' };

    const rowNum = targetRow + 1;
    const tz     = Session.getScriptTimeZone();
    const now    = new Date();

    // Update KM, BBM, Dicatat Masuk Oleh, status → Telah Servis
    if (kmCol  !== -1) sheet.getRange(rowNum, kmCol+1).setValue(sanitizeCell(km));
    if (bbmCol !== -1) sheet.getRange(rowNum, bbmCol+1).setValue(sanitizeCell(bbm + '%'));
    if (dicCol !== -1) sheet.getRange(rowNum, dicCol+1).setValue(session.nama || 'dashboard');
    if (stCol  !== -1) sheet.getRange(rowNum, stCol+1).setValue('Telah Servis');
    if (jamCol !== -1) sheet.getRange(rowNum, jamCol+1).setValue(Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm:ss'));

    // Update status kendaraan → Tersedia
    try {
      const kndSheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
      if (kndSheet) {
        const kd    = kndSheet.getDataRange().getValues();
        const kh    = kd[0].map(h => String(h).trim());
        const kpCol = (() => { for (const c of ['Plat Nomor','Plat_Nomor','No. Polisi']) { const i = kh.indexOf(c); if (i !== -1) return i; } return -1; })();
        const ksCol = kh.indexOf('Status');
        if (kpCol !== -1 && ksCol !== -1) {
          for (let j = 1; j < kd.length; j++) {
            if (String(kd[j][kpCol]).trim().toUpperCase() === String(platNomor).trim().toUpperCase()) {
              kndSheet.getRange(j+1, ksCol+1).setValue('Tersedia');
              break;
            }
          }
        }
      }
    } catch(kErr) { console.error('updateMobilServis: gagal update status kendaraan:', kErr.message); }

    writeLog('dashboard', 'SELESAI_SERVIS', 'Servis selesai: ' + platNomor);
    bumpUpdateToken_(); return { success: true, message: 'Servis selesai, status kendaraan menjadi Tersedia.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function getListMobilMasukPublic(token) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('MobilMasuk');
    if (!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1)
      .filter(row => row.some(cell => String(cell).trim() !== ''))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          obj[h] = (row[i] instanceof Date)
            ? Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
            : row[i];
        });
        return obj;
      });
    rows.reverse();
    return { success: true, data: rows };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function deleteMobilMasukPublic(token, id) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('MobilMasuk');
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan.' };
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        writeLog('dashboard', 'DELETE_MASUK', 'Hapus masuk ID: ' + id);
        bumpUpdateToken_(); return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

/**
 * Simpan permohonan peminjaman TANPA token.
 */
function savePeminjamanPublic(token, payload) {
  try {
    var session = requireAuth(token);
    if (!payload.platNomor)    return { success: false, message: 'Plat nomor tidak boleh kosong.' };
    if (!payload.namaPemohon)  return { success: false, message: 'Nama pemohon tidak boleh kosong.' };
    if (!payload.namaPengemudi) return { success: false, message: 'Nama pengguna tidak boleh kosong.' };

    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Peminjaman');
    if (!sheet) {
      sheet = ss.insertSheet('Peminjaman');
      sheet.getRange(1,1,1,HEADERS_PEMINJAMAN.length).setValues([HEADERS_PEMINJAMAN]);
      formatHeaderRow(sheet);
      applyZebraRows(sheet);
      sheet.autoResizeColumns(1, HEADERS_PEMINJAMAN.length);
    }

    const totalCol  = sheet.getLastColumn();
    const headerRow = sheet.getRange(1, 1, 1, totalCol).getValues()[0];
    const headerMap = {};
    headerRow.forEach(function(h, i){ var k = String(h).trim().toLowerCase(); if (k) headerMap[k] = i; });

    var statusIdx = headerMap['status'];
    if (statusIdx === undefined) statusIdx = totalCol - 1;
    var rowWidth = statusIdx + 1;

    function set(row, name, val) {
      var i = headerMap[name.toLowerCase()];
      if (i !== undefined && i < rowWidth) row[i] = sanitizeCell(val);
    }

    const now = new Date();
    const noUrut = Math.max(1, sheet.getLastRow());
    const ts = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    const row = new Array(rowWidth).fill('');

    set(row, 'no',               noUrut);
    set(row, 'timestamp',        ts);
    set(row, 'plat nomor',       payload.platNomor      || '');
    set(row, 'merk',             payload.merkKendaraan  || '');
    set(row, 'tipe',             payload.tipeKendaraan  || '');
    set(row, 'model',            payload.modelKendaraan || '');
    set(row, 'jenis',            payload.jenisKendaraan || '');
    set(row, 'tgl. keluar',      payload.tglPinjam      || '');
    set(row, 'jam keluar',       payload.jamPinjam      || '');
    set(row, 'tgl. masuk',       payload.tglKembali     || '');
    set(row, 'jam masuk',        payload.jamKembali     || '');
    set(row, 'nama pengguna',    payload.namaPengemudi  || '');
    set(row, 'nip pengguna',     '');
    set(row, 'no. hp pengguna',  payload.telpPengemudi  || '');
    set(row, 'keperluan',        payload.keperluan      || '');
    set(row, 'lokasi tujuan',    payload.lokasiTujuan   || '');
    set(row, 'status',           'Pengajuan');

    sheet.appendRow(row);
    const newRow = sheet.getLastRow();
    try { _formatRow(sheet, newRow); } catch(e) {}

    // Update status Kendaraan → Permohonan
    try {
      const kndSheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
      if (kndSheet) {
        const kndData = kndSheet.getDataRange().getValues();
        const hdrs    = kndData[0].map(h => String(h).trim());
        const platCol = (() => { for (const c of ['Plat Nomor','Plat_Nomor','No. Polisi']) { const i = hdrs.indexOf(c); if (i !== -1) return i; } return -1; })();
        const statusCol = hdrs.indexOf('Status');
        if (platCol !== -1 && statusCol !== -1) {
          for (let i = 1; i < kndData.length; i++) {
            if (String(kndData[i][platCol]).trim() === String(payload.platNomor).trim()) {
              kndSheet.getRange(i + 1, statusCol + 1).setValue('Permohonan');
              break;
            }
          }
        }
      }
    } catch(e) { console.error('Update status Permohonan error:', e); }

    writeLog(session.username,'PEMINJAMAN','Peminjaman: '+payload.platNomor+' oleh '+payload.namaPemohon);
    try { formatPeminjamanStatus(); } catch(e) {}
    bumpUpdateToken_(); return { success: true, id: noUrut, message: 'Permohonan berhasil disimpan.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

/**
 * Ambil daftar permohonan peminjaman TANPA token.
 * Filter: hanya status "Menunggu Persetujuan".
 */
/**
 * DEBUG: Kembalikan header dan baris pertama sheet Peminjaman
 * untuk mengetahui nama kolom yang sebenarnya.
 */
function getListPeminjamanPublic(token) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('Peminjaman');
    if (!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    const headers = data[0].map(h => String(h).trim());
    const tz = Session.getScriptTimeZone();

    const formatTime = v => {
      if (v instanceof Date) return Utilities.formatDate(v, tz, 'HH:mm');
      if (typeof v === 'number' && v > 0 && v < 1) {
        const mins = Math.round(v * 24 * 60);
        return String(Math.floor(mins/60)).padStart(2,'0') + ':' + String(mins%60).padStart(2,'0');
      }
      return (v !== null && v !== undefined) ? String(v) : '';
    };
    const formatDate = v => {
      if (v instanceof Date && v.getFullYear() > 1900) return Utilities.formatDate(v, tz, 'dd/MM/yyyy');
      return (v !== null && v !== undefined) ? String(v) : '';
    };

    const rows = data.slice(1)
      .filter(row => row.some(c => String(c).trim() !== ''))
      .map((row, i) => {
        const obj = { _rowIndex: i + 2 };
        headers.forEach((h, j) => {
          const v  = row[j];
          const lh = h.toLowerCase();
          if (lh.includes('jam') && !lh.includes('input')) {
            obj[h] = formatTime(v);
          } else if (lh.includes('tgl') || lh.includes('tanggal')) {
            obj[h] = formatDate(v);
          } else if (v instanceof Date) {
            obj[h] = v.getFullYear() > 1900 ? Utilities.formatDate(v, tz, 'dd/MM/yyyy HH:mm') : '';
          } else {
            obj[h] = (v !== null && v !== undefined) ? String(v) : '';
          }
        });
        return obj;
      })
      .filter(row => {
        const st = String(row['Status']||'').toLowerCase();
        return st.indexOf('pengajuan') !== -1 || st.indexOf('menunggu') !== -1;
      });
    rows.reverse();
    return { success: true, data: rows };
  } catch(err) {
    return { success: false, message: err.message };
  }
}


/**
 * Ambil Nama Lengkap dari sheet Users berdasarkan username atau nama.
 * Digunakan untuk mengisi kolom "Disetujui Oleh".
 */
function getNamaLengkap(usernameOrNama) {
  if (!usernameOrNama) return usernameOrNama || '';
  // Super Admin hardcoded
  if (String(usernameOrNama).toLowerCase().trim() === getConfig('ADMIN_USERNAME')) return getConfig('ADMIN_NAME');
  try {
    const ss = getSpreadsheet();
    const sh = ss.getSheetByName(getConfig('SHEET_USERS'));
    if (!sh) return usernameOrNama;
    const d = sh.getDataRange().getValues();
    const h = d[0].map(x => String(x).trim());
    const uCol = h.indexOf('Username') !== -1 ? h.indexOf('Username') : h.indexOf('username');
    const nCol = (() => { for (const c of ['Nama Pengguna','Nama Lengkap','Nama']) { const i = h.indexOf(c); if (i !== -1) return i; } return -1; })();
    if (uCol === -1 || nCol === -1) return usernameOrNama;
    // Cari berdasarkan username (exact match)
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][uCol]).trim().toLowerCase() === String(usernameOrNama).trim().toLowerCase()) {
        return String(d[i][nCol]).trim() || usernameOrNama;
      }
    }
    // Cari berdasarkan Nama Lengkap (fallback)
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][nCol]).trim().toLowerCase() === String(usernameOrNama).trim().toLowerCase()) {
        return String(d[i][nCol]).trim() || usernameOrNama;
      }
    }
  } catch(e) {}
  return usernameOrNama;
}
/**
 * Warnai kolom Status di sheet Peminjaman.
 * Kuning = Pengajuan / Menunggu, Hijau = Disetujui, Merah = Dibatalkan.
 * Jalankan manual dari GAS editor, atau dipanggil setelah save/update.
 */
function formatPeminjamanStatus() {
  try {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('Peminjaman');
    if (!sheet) return;
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const stCol   = headers.indexOf('Status');
    if (stCol === -1) return;
    for (let i = 1; i < data.length; i++) {
      const st  = String(data[i][stCol]).trim().toLowerCase();
      const cell = sheet.getRange(i+1, stCol+1);
      if (st === 'pengajuan' || st.indexOf('menunggu') !== -1) {
        cell.setBackground('#fef08a'); cell.setFontColor('#854d0e');
      } else if (st === 'disetujui' || st.indexOf('disetujui') !== -1) {
        cell.setBackground('#bbf7d0'); cell.setFontColor('#14532d');
      } else if (st === 'dibatalkan') {
        cell.setBackground('#fecaca'); cell.setFontColor('#7f1d1d');
      } else {
        cell.setBackground(null); cell.setFontColor(null);
      }
    }
  } catch(e) { console.error('formatPeminjamanStatus:', e); }
}
/**
 * Setujui permohonan peminjaman berdasarkan Plat Nomor.
 */
function setujuiPeminjamanPublic(token, platNomor, disetujuiOleh) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('Peminjaman');
    if (!sheet) return { success: false, message: 'Sheet Peminjaman tidak ditemukan.' };

    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());

    // Kolom-kolom yang dibutuhkan
    const platCol   = (() => { for (const c of ['Plat Nomor','Plat_Nomor','No. Polisi','No Polisi','Plat']) { const i = headers.indexOf(c); if (i !== -1) return i; } return -1; })();
    const stCol     = headers.indexOf('Status');
    const dstCol    = headers.indexOf('Disetujui Oleh');
    const nmPCol    = headers.indexOf('Nama Pengemudi');
    const nmPmCol   = headers.indexOf('Nama Pemohon');
    const tglPCol   = (() => { for (const c of ['Tgl. Keluar','Tgl Keluar','Tgl. Pinjam','Tgl Pinjam']) { const i = headers.indexOf(c); if (i !== -1) return i; } return -1; })();
    const jamPCol   = (() => { for (const c of ['Jam Keluar','Jam Pinjam']) { const i = headers.indexOf(c); if (i !== -1) return i; } return -1; })();
    const tglKCol   = (() => { for (const c of ['Tgl. Masuk','Tgl Masuk','Tgl. Kembali','Tgl Kembali']) { const i = headers.indexOf(c); if (i !== -1) return i; } return -1; })();
    const jamKCol   = (() => { for (const c of ['Jam Masuk','Jam Kembali']) { const i = headers.indexOf(c); if (i !== -1) return i; } return -1; })();
    const lokasiCol  = headers.indexOf('Lokasi Tujuan');
    const jabatanCol = headers.indexOf('Jabatan Pemohon');
    const nmDrvCol   = (() => { for (const c of ['Nama Pengguna','Nama Pengemudi','nama_pengemudi']) { const idx = headers.indexOf(c); if (idx !== -1) return idx; } return -1; })();
    const hpDrvCol   = (() => { for (const c of ['No. HP Pengguna','No. HP Pengemudi','Telp Pengemudi','HP Pengemudi','Nomor HP Pengemudi']) { const idx = headers.indexOf(c); if (idx !== -1) return idx; } return -1; })();

    if (platCol === -1) return { success: false, message: 'Kolom Plat Nomor tidak ditemukan di sheet Peminjaman.' };

    // ── Cari baris permohonan ──
    // Cari: Plat Nomor cocok + status "Pengajuan"/"Menunggu"
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const rowPlat   = String(data[i][platCol]).trim().toUpperCase();
      const rowStatus = stCol !== -1 ? String(data[i][stCol]).trim().toLowerCase() : '';
      if (rowPlat === String(platNomor).trim().toUpperCase()
          && (rowStatus.indexOf('pengajuan') !== -1 || rowStatus.indexOf('menunggu') !== -1)) {
        targetRow = i; break;
      }
    }
    // Fallback: ambil baris terbaru untuk plat ini
    if (targetRow === -1) {
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][platCol]).trim().toUpperCase() === String(platNomor).trim().toUpperCase()) {
          targetRow = i; break;
        }
      }
    }

    if (targetRow === -1) {
      return { success: false, message: 'Permohonan untuk kendaraan ' + platNomor + ' tidak ditemukan.' };
    }

    const i      = targetRow;
    const rowNum = i + 1;
    const plat   = String(data[i][platCol]).trim();

    // ── Deteksi Servis dari baris yang ditemukan ──
    const kepCol  = headers.indexOf('Keperluan');
    var isServis  = false;
    if (kepCol !== -1) {
      var keperluan = String(data[i][kepCol] || '').trim().toLowerCase();
      if (keperluan.indexOf('servis') !== -1) isServis = true;
    }

    // Update status permohonan
    if (stCol  !== -1) sheet.getRange(rowNum, stCol+1).setValue('Disetujui');
    if (dstCol !== -1) sheet.getRange(rowNum, dstCol+1).setValue(disetujuiOleh || 'administrator');

    // Tambahkan ke MobilKeluar
    try {
      let mkSheet = ss.getSheetByName(getConfig('SHEET_MOBIL_KELUAR'));
      if (!mkSheet) mkSheet = setupMobilKeluarSheet();
      const mkHeaders = mkSheet.getRange(1,1,1,mkSheet.getLastColumn()).getValues()[0].map(h=>String(h).trim());
      const mkLastRow = mkSheet.getLastRow();
      const mkId      = 'KLR_' + String(mkLastRow).padStart(4,'0');
      const mkRow = mkHeaders.map(h => {
        switch(h) {
          case 'ID':             return mkId;
          case 'Tanggal':        return tglPCol !== -1 ? (() => {
                                   const v = data[i][tglPCol];
                                   if (v instanceof Date && v.getFullYear() > 1900)
                                     return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
                                   return String(v || '');
                                 })() : '';
          case 'Waktu Input':    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
          case 'Plat Nomor':     return plat;
          case 'Jam Keluar':     return jamPCol !== -1 ? (() => {
                                   const v = data[i][jamPCol];
                                   if (v instanceof Date)
                                     return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
                                   if (typeof v === 'number' && v > 0 && v < 1) {
                                     const m = Math.round(v*24*60);
                                     return String(Math.floor(m/60)).padStart(2,'0')+':'+String(m%60).padStart(2,'0');
                                   }
                                   return String(v || '');
                                 })() : '';
          case 'KM Keluar':      return '';
          case 'BBM (%)':
          case 'BBM (L)':        return '';
          case 'Tujuan':         return lokasiCol !== -1 ? String(data[i][lokasiCol]) : '';
          case 'Nama Pengguna':  return nmDrvCol !== -1 ? String(data[i][nmDrvCol]) : (nmPmCol !== -1 ? String(data[i][nmPmCol]) : '');
          case 'No. HP Pengguna':
          case 'Nomor HP':       return hpDrvCol !== -1 ? String(data[i][hpDrvCol]) : '';
          case 'Disetujui Oleh':
          case 'Pemberi Akses Keluar':
          case 'ACC Pimpinan':   return getNamaLengkap(disetujuiOleh) || disetujuiOleh || 'administrator';
          case 'Catatan':        return 'Dari permohonan ' + plat;
          case 'Dicatat Keluar Oleh':   return '';  // diisi operator saat Update Data
          default:               return '';
        }
      });
      mkSheet.appendRow(mkRow);
    } catch(mkErr) {
      console.error('setujuiPeminjaman: gagal tambah MobilKeluar:', mkErr.message);
    }

    // Update status kendaraan → Keluar (atau Servis jika keperluan servis)
    try {
      const kndSheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
      if (kndSheet) {
        const kd      = kndSheet.getDataRange().getValues();
        const kh      = kd[0].map(h => String(h).trim());
        const kPlat   = (() => { for (const c of ['Plat Nomor','Plat_Nomor','No. Polisi','No Polisi']) { const idx = kh.indexOf(c); if (idx !== -1) return idx; } return -1; })();
        const kStatus = kh.indexOf('Status');
        if (kPlat !== -1 && kStatus !== -1) {
          for (let j = 1; j < kd.length; j++) {
            if (String(kd[j][kPlat]).trim().toUpperCase() === plat.toUpperCase()) {
              kndSheet.getRange(j+1, kStatus+1).setValue(isServis ? 'Servis' : 'Mobil Keluar');
              break;
            }
          }
        }
      }
    } catch(kErr) {
      console.error('setujuiPeminjaman: gagal update status kendaraan:', kErr.message);
    }

    writeLog('dashboard', 'SETUJUI_PEMINJAMAN', 'Setujui plat: ' + plat + ' oleh ' + (disetujuiOleh || 'administrator'));
    try { formatPeminjamanStatus(); } catch(e) {}
    bumpUpdateToken_(); return {
      success: true,
      keluarPayload: {
        platNomor:     plat,
        namaPengguna:  nmPmCol !== -1 ? String(data[i][nmPmCol]) : '',
        namaPengemudi: nmPCol  !== -1 ? String(data[i][nmPCol])  : '',
        tglPinjam:     tglPCol !== -1 ? String(data[i][tglPCol]) : '',
        jamKeluar:     jamPCol !== -1 ? String(data[i][jamPCol]) : '',
        tglKembali:    tglKCol !== -1 ? String(data[i][tglKCol]) : '',
        jamKembali:    jamKCol !== -1 ? String(data[i][jamKCol]) : '',
      }
    };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

/**
 * Batalkan permohonan: ubah status → "Dibatalkan".
 */
function batalkanPeminjamanPublic(token, platNomor) {
  try {
    var session = requireAuth(token);
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('Peminjaman');
    if (!sheet) return { success: false, message: 'Sheet Peminjaman tidak ditemukan.' };

    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());

    const platCol = (() => { for (const c of ['Plat Nomor','Plat_Nomor','No. Polisi','No Polisi','Plat']) { const i = headers.indexOf(c); if (i !== -1) return i; } return -1; })();
    const stCol   = headers.indexOf('Status');

    if (platCol === -1) return { success: false, message: 'Kolom Plat Nomor tidak ditemukan di sheet Peminjaman.' };

    // Cari baris: cocokkan Plat Nomor + status "Menunggu"
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const rowPlat   = String(data[i][platCol]).trim().toUpperCase();
      const rowStatus = stCol !== -1 ? String(data[i][stCol]).trim().toLowerCase() : '';
      if (rowPlat === String(platNomor).trim().toUpperCase()
          && (rowStatus.indexOf('pengajuan') !== -1 || rowStatus.indexOf('menunggu') !== -1)) {
        targetRow = i; break;
      }
    }
    // Fallback: cocokkan Plat Nomor saja (ambil baris terbaru)
    if (targetRow === -1) {
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][platCol]).trim().toUpperCase() === String(platNomor).trim().toUpperCase()) {
          targetRow = i; break;
        }
      }
    }

    if (targetRow === -1) {
      return { success: false, message: 'Permohonan untuk kendaraan ' + platNomor + ' tidak ditemukan.' };
    }

    const i      = targetRow;
    const rowNum = i + 1;
    const plat   = String(data[i][platCol]).trim();

    if (stCol !== -1) sheet.getRange(rowNum, stCol+1).setValue('Dibatalkan');

    // Kembalikan status kendaraan → Tersedia
    try {
      const kndSheet = ss.getSheetByName(getConfig('SHEET_KENDARAAN'));
      if (kndSheet) {
        const kd      = kndSheet.getDataRange().getValues();
        const kh      = kd[0].map(h => String(h).trim());
        const kPlat   = (() => { for (const c of ['Plat Nomor','Plat_Nomor','No. Polisi','No Polisi']) { const idx = kh.indexOf(c); if (idx !== -1) return idx; } return -1; })();
        const kStatus = kh.indexOf('Status');
        if (kPlat !== -1 && kStatus !== -1) {
          for (let j = 1; j < kd.length; j++) {
            if (String(kd[j][kPlat]).trim().toUpperCase() === plat.toUpperCase()) {
              kndSheet.getRange(j+1, kStatus+1).setValue('Tersedia');
              break;
            }
          }
        }
      }
    } catch(kErr) {
      console.error('batalkanPeminjaman: gagal update status kendaraan:', kErr.message);
    }

    writeLog('dashboard','BATALKAN_PEMINJAMAN','Batalkan plat: ' + plat);
    try { formatPeminjamanStatus(); } catch(e) {}
    bumpUpdateToken_(); return { success: true };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}