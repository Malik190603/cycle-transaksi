/**
 * ExternalDb.gs
 * v8.25.0 -- Master (data user/role) & Config (Config_Sistem, Config_Akses_Setting) pindah
 * ke SPREADSHEET TERPISAH dari spreadsheet operasional (Data Count/Riwayat/Log dst), supaya
 * spreadsheet operasional yang tumbuh terus tiap hari tidak bikin config ikut berat, dan
 * supaya config bisa dikontrol rapi dari 1 tempat.
 *
 * ID kedua spreadsheet baru disimpan di Script Properties (MASTER_DB_ID, CONFIG_DB_ID) --
 * BUKAN hardcode -- supaya kalau suatu saat perlu pindah lagi, tinggal ganti Script
 * Property, tidak perlu edit kode.
 *
 * FALLBACK: kalau Master DB / Config DB gagal diakses (ID belum di-set, spreadsheet
 * kehapus, izin bermasalah, dsb), sistem OTOMATIS balik pakai sheet "Master"/"Config_Sistem"/
 * "Config_Akses_Setting" LOKAL di spreadsheet operasional (sheet lama TETAP disimpan sebagai
 * cadangan, sengaja TIDAK dihapus saat migrasi) -- supaya orang tetap bisa login & app tetap
 * jalan walau Master DB sedang bermasalah. Setiap kali fallback ini kepakai, dicatat ke sheet
 * "Log_Sistem" di spreadsheet operasional supaya admin tahu ada yang perlu diperbaiki.
 */

const SCRIPT_PROP_MASTER_DB_ID = 'MASTER_DB_ID';
const SCRIPT_PROP_CONFIG_DB_ID = 'CONFIG_DB_ID';
const LOG_SISTEM_SHEET_NAME = 'Log_Sistem';
const LOG_SISTEM_HEADERS = ['Waktu', 'Sumber', 'Pesan'];

// ---------- Log_Sistem (audit trail fallback/warning) ----------

function catatLogSistem_(sumber, pesan) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(LOG_SISTEM_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_SISTEM_SHEET_NAME);
      sheet.appendRow(LOG_SISTEM_HEADERS);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), sumber, pesan]);
  } catch (e) {
    // Log_Sistem gagal ditulis bukan alasan buat gagalkan operasi utama -- diabaikan.
  }
}

// ---------- Master DB ----------

/**
 * Coba buka Master DB (spreadsheet terpisah). Return spreadsheet object kalau berhasil,
 * atau null kalau MASTER_DB_ID belum di-set / gagal dibuka (dicatat ke Log_Sistem sekali
 * per kegagalan, bukan berulang -- lihat cache 'masterDbDown' di bawah).
 */
function tryOpenMasterDb_() {
  const id = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_MASTER_DB_ID);
  if (!id) return null;

  // Kalau baru saja gagal dibuka (60 detik terakhir), jangan coba lagi -- supaya 1 kegagalan
  // (mis. timeout) tidak bikin SETIAP pemanggil getMasterSheet_() dalam eksekusi lain ikut
  // menunggu lama, cukup langsung fallback.
  const cache = CacheService.getScriptCache();
  if (cache.get('masterDbDown')) return null;

  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    try { cache.put('masterDbDown', '1', 60); } catch (e2) { /* abaikan */ }
    catatLogSistem_('Master DB', 'Gagal membuka Master DB (' + e.message + '). Fallback ke Master sheet lokal.');
    return null;
  }
}

/**
 * Sheet "Master" yang SEBENARNYA dipakai getMasterSheet_() (SheetHelpers.gs). Coba Master
 * DB dulu, fallback ke sheet Master lokal di spreadsheet operasional kalau gagal.
 */
function getMasterSheetWithFallback_() {
  const masterDb = tryOpenMasterDb_();
  if (masterDb) {
    const sheet = masterDb.getSheetByName(MASTER_SHEET_NAME);
    if (sheet) return sheet;
    catatLogSistem_('Master DB', 'Master DB terbuka tapi sheet "Master" tidak ditemukan di dalamnya. Fallback ke Master sheet lokal.');
  }
  return getLegacyMasterSheetLocal_();
}

function getLegacyMasterSheetLocal_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(MASTER_SHEET_NAME);
}

// ---------- Config DB ----------

function tryOpenConfigDb_() {
  const id = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_CONFIG_DB_ID);
  if (!id) return null;

  const cache = CacheService.getScriptCache();
  if (cache.get('configDbDown')) return null;

  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    try { cache.put('configDbDown', '1', 60); } catch (e2) { /* abaikan */ }
    catatLogSistem_('Config DB', 'Gagal membuka Config DB (' + e.message + '). Fallback ke sheet config lokal.');
    return null;
  }
}

/**
 * Return spreadsheet target untuk sheet-sheet Config_Sistem / Config_Akses_Setting --
 * Config DB kalau berhasil dibuka, atau spreadsheet operasional (lokal) kalau fallback.
 * Dipakai oleh OverflowConfig.gs supaya 2 sheet itu ikut pindah otomatis begitu
 * CONFIG_DB_ID di-set, TANPA ubah kode di OverflowConfig.gs sama sekali.
 */
function getConfigTargetSpreadsheet_() {
  const configDb = tryOpenConfigDb_();
  return configDb || SpreadsheetApp.getActiveSpreadsheet();
}

// ---------- Log_Perubahan_Config (audit trail SETIAP edit config oleh admin -- beda dari
// Log_Sistem yang khusus warning/fallback otomatis) ----------

const LOG_PERUBAHAN_CONFIG_SHEET_NAME = 'Log_Perubahan_Config';
const LOG_PERUBAHAN_CONFIG_HEADERS = ['Waktu', 'NIK', 'Sumber', 'Detail'];

function getLogPerubahanConfigSheet_() {
  const ss = getConfigTargetSpreadsheet_();
  let sheet = ss.getSheetByName(LOG_PERUBAHAN_CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_PERUBAHAN_CONFIG_SHEET_NAME);
    sheet.appendRow(LOG_PERUBAHAN_CONFIG_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Dipanggil dari SETIAP fungsi yang mengubah config (Level Assignment, User Management,
 * Akses Setting) -- supaya ada jejak siapa ubah apa kapan. Gagal nulis log TIDAK
 * menggagalkan operasi utama (sama seperti catatLogSistem_).
 */
function catatLogPerubahanConfig_(nik, sumber, detail) {
  try {
    getLogPerubahanConfigSheet_().appendRow([new Date(), nik || '(tidak diketahui)', sumber, detail]);
  } catch (e) {
    // diabaikan -- log gagal bukan alasan gagalkan operasi utama
  }
}

/**
 * Bikin Master DB & Config DB (spreadsheet baru terpisah), pindahkan (copy) data yang
 * SUDAH ADA sekarang, dan simpan ID-nya ke Script Properties. Sheet lokal yang lama
 * SENGAJA TIDAK dihapus -- tetap jadi cadangan fallback kalau Master DB/Config DB
 * bermasalah nanti. Aman dijalankan berkali-kali (idempotent): kalau Script Property
 * sudah ke-set, tidak bikin spreadsheet baru lagi, cuma dicek ulang datanya lengkap.
 */
function setupExternalDbMasterDanConfig() {
  const props = PropertiesService.getScriptProperties();
  const hasil = [];

  // ---- Master DB ----
  let masterDbId = props.getProperty(SCRIPT_PROP_MASTER_DB_ID);
  let masterDb;
  if (masterDbId) {
    try { masterDb = SpreadsheetApp.openById(masterDbId); } catch (e) { masterDb = null; }
  }
  if (!masterDb) {
    masterDb = SpreadsheetApp.create('Master DB - Cycle Count DC');
    props.setProperty(SCRIPT_PROP_MASTER_DB_ID, masterDb.getId());
    hasil.push('Master DB baru dibuat: ' + masterDb.getUrl());
  } else {
    hasil.push('Master DB sudah ada, dipakai lagi: ' + masterDb.getUrl());
  }

  let masterSheetBaru = masterDb.getSheetByName(MASTER_SHEET_NAME);
  const masterLokal = getLegacyMasterSheetLocal_();
  if (!masterSheetBaru && masterLokal) {
    masterSheetBaru = masterLokal.copyTo(masterDb);
    masterSheetBaru.setName(MASTER_SHEET_NAME);
    // copyTo selalu menaruh sheet hasil copy di posisi terakhir & sheet default "Sheet1"
    // bawaan SpreadsheetApp.create() masih nempel di Master DB -- dibuang supaya bersih.
    const sheet1 = masterDb.getSheetByName('Sheet1');
    if (sheet1 && masterDb.getSheets().length > 1) masterDb.deleteSheet(sheet1);
    hasil.push('Data sheet "Master" (' + (masterLokal.getLastRow() - 1) + ' baris) berhasil dipindah ke Master DB.');
  } else if (masterSheetBaru) {
    hasil.push('Sheet "Master" sudah ada di Master DB, tidak ditimpa (data existing di Master DB dipertahankan).');
  } else {
    hasil.push('PERINGATAN: sheet "Master" lokal tidak ditemukan, Master DB dibuat kosong.');
  }

  // ---- Config DB ----
  let configDbId = props.getProperty(SCRIPT_PROP_CONFIG_DB_ID);
  let configDb;
  if (configDbId) {
    try { configDb = SpreadsheetApp.openById(configDbId); } catch (e) { configDb = null; }
  }
  if (!configDb) {
    configDb = SpreadsheetApp.create('Config DB - Cycle Count DC');
    props.setProperty(SCRIPT_PROP_CONFIG_DB_ID, configDb.getId());
    hasil.push('Config DB baru dibuat: ' + configDb.getUrl());
  } else {
    hasil.push('Config DB sudah ada, dipakai lagi: ' + configDb.getUrl());
  }

  [CONFIG_SISTEM_SHEET_NAME, CONFIG_AKSES_SETTING_SHEET_NAME].forEach(function (namaSheet) {
    let target = configDb.getSheetByName(namaSheet);
    const lokal = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet);
    if (!target && lokal) {
      target = lokal.copyTo(configDb);
      target.setName(namaSheet);
      hasil.push('Sheet "' + namaSheet + '" berhasil dipindah ke Config DB.');
    } else if (target) {
      hasil.push('Sheet "' + namaSheet + '" sudah ada di Config DB, tidak ditimpa.');
    }
  });
  const configSheet1 = configDb.getSheetByName('Sheet1');
  if (configSheet1 && configDb.getSheets().length > 1) configDb.deleteSheet(configSheet1);

  // Pastikan sheet Config_Sistem & Config_Akses_Setting benar-benar terisi header+default
  // (kalau tadi belum ke-copy sama sekali karena sheet lokalnya juga belum pernah dibuat).
  getConfigSistemSheet_();
  getConfigAksesSettingSheet_();

  const pesan = hasil.join('\n');
  try {
    SpreadsheetApp.getUi().alert('Setup Master DB & Config DB selesai:\n\n' + pesan);
  } catch (e) {
    Logger.log(pesan);
  }
}

// ---------- Master_User (Fase Migrasi Centralized User + Facility) ----------

const MASTER_USER_SHEET_NAME = 'Master_User';
const MASTER_USER_HEADERS = ['Username', 'Role', 'Status', 'ID_Facility'];

/**
 * Mengambil sheet Master_User dari Master DB (fallback ke lokal).
 */
function getMasterUserSheetWithFallback_() {
  const masterDb = tryOpenMasterDb_();
  let sheet = null;
  if (masterDb) {
    sheet = masterDb.getSheetByName(MASTER_USER_SHEET_NAME);
  }
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_USER_SHEET_NAME);
  }
  return sheet;
}

/**
 * Fungsi sekali jalan untuk memigrasikan data dari sheet Master (kolom C/D/E)
 * dan sheet User_Facility_Assignment ke sheet Master_User baru.
 */
function migrasiMasterKeSheetBaru() {
  const masterSheet = getMasterSheetWithFallback_();
  const assignSheet = getUserFacilityAssignmentSheet_();
  
  const masterDb = tryOpenMasterDb_() || SpreadsheetApp.getActiveSpreadsheet();
  let newUserSheet = masterDb.getSheetByName(MASTER_USER_SHEET_NAME);
  
  if (!newUserSheet) {
    newUserSheet = masterDb.insertSheet(MASTER_USER_SHEET_NAME);
    newUserSheet.appendRow(MASTER_USER_HEADERS);
    newUserSheet.setFrozenRows(1);
  }
  
  // Baca existing assignment
  const assignments = {};
  if (assignSheet && assignSheet.getLastRow() >= 2) {
    const data = assignSheet.getRange(2, 1, assignSheet.getLastRow() - 1, 2).getValues();
    data.forEach(function(r) {
      const uname = String(r[0] || '').trim().toLowerCase();
      if (uname) assignments[uname] = String(r[1] || '').trim();
    });
  }
  
  // Baca master lama dan gabungkan
  if (masterSheet && masterSheet.getLastRow() >= 2) {
    const data = masterSheet.getRange(2, 3, masterSheet.getLastRow() - 1, 3).getValues(); // C, D, E
    const rows = [];
    data.forEach(function(r) {
      const uname = String(r[0] || '').trim();
      const role = String(r[1] || '').trim();
      const status = String(r[2] || '').trim();
      if (uname) {
        const facId = assignments[uname.toLowerCase()] || '';
        rows.push([uname, role, status || 'Aktif', facId]);
      }
    });
    
    if (rows.length > 0) {
       const existingLastRow = newUserSheet.getLastRow();
       if (existingLastRow >= 2) {
           newUserSheet.getRange(2, 1, existingLastRow - 1, 4).clearContent();
       }
       newUserSheet.getRange(2, 1, rows.length, 4).setValues(rows);
    }
  }
  return "Migrasi berhasil dilakukan ke sheet Master_User. Silakan refresh web app.";
}
