/**
 * SheetHelpers.gs
 * Helper akses/pembuatan sheet (Data Count, Riwayat, Master) & pencarian
 * baris
 * (Hasil pemecahan dari Code.gs -- semua file .gs berbagi global scope yang
 * sama di Apps Script,
 * jadi urutan file tidak masalah, fungsi di file lain tetap bisa saling
 * panggil.)
 *
 * v8.26.0 (Facility Management):
 * Semua fungsi akses sheet operasional SEKARANG menerima parameter opsional
 * 'username'.
 * Kalau username disediakan: sheet diambil dari spreadsheet facility user
 * tersebut
 * (via getOperasionalSpreadsheet_() di FacilityManagement.gs).
 * Kalau username TIDAK disediakan: fallback ke SpreadsheetApp.
 * getActiveSpreadsheet()
 * (kompatibilitas penuh dengan kode lama sebelum ada Facility Management).
 */

// ---------- Helper sheet ----------

/**
 * Sheet "Data Count".
 * @param {string} [username] - Opsional: username untuk menentukan facility
 * spreadsheet
 */
function getSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, HEADERS);
  }
  return sheet;
}

/**
 * Sheet "Riwayat".
 * @param {string} [username] - Opsional: username untuk menentukan facility
 * spreadsheet
 */
function getRiwayatSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RIWAYAT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RIWAYAT_SHEET_NAME);
    sheet.appendRow(RIWAYAT_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, RIWAYAT_HEADERS);
  }
  return sheet;
}

/**
 * Sheet "Riwayat_Archive".
 * @param {string} [username] - Opsional: username untuk menentukan facility
 * spreadsheet
 */
function getRiwayatArchiveSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RIWAYAT_ARCHIVE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RIWAYAT_ARCHIVE_SHEET_NAME);
    sheet.appendRow(RIWAYAT_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, RIWAYAT_HEADERS);
  }
  return sheet;
}

/**
 * Migrasi ringan: kalau HEADERS/RIWAYAT_HEADERS di kode ini bertambah kolom
 * baru (mis. saat
 * menambah AddWho_Transaksi), sheet yang SUDAH ADA datanya (dibuat sebelum
 * kolom ini ditambah)
 * tidak otomatis punya kolom itu -- hanya sheet yang BARU dibuat lewat
 * insertSheet() yang dapat
 * appendRow(HEADERS) penuh. Fungsi ini menambahkan HANYA header yang masih
 * kurang di ujung kanan,
 * TIDAK menyentuh/menggeser kolom yang sudah ada, supaya data lama tetap
 * aman.
 */
var _ensuredHeadersCache_ = {};

function ensureHeaderColumns_(sheet, expectedHeaders) {
  try {
    const sheetId = sheet.getSheetId();
    if (_ensuredHeadersCache_[sheetId]) return;
    const lastCol = sheet.getLastColumn();
    if (lastCol < expectedHeaders.length) {
      const missingHeaders = expectedHeaders.slice(lastCol);
      sheet.getRange(1, lastCol + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    }
    _ensuredHeadersCache_[sheetId] = true;
  } catch (e) { /* abaikan */ }
}

/**
 * Sheet "Log_Akses".
 * @param {string} [username] - Opsional: username untuk menentukan facility
 * spreadsheet
 */
function getLogAksesSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_AKSES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_AKSES_SHEET_NAME);
    sheet.appendRow(LOG_AKSES_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, LOG_AKSES_HEADERS);
  }
  return sheet;
}

/**
 * Sheet "Log_Anomali".
 * @param {string} [username] - Opsional: username untuk menentukan facility
 * spreadsheet
 */
function getLogAnomaliSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_ANOMALI_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_ANOMALI_SHEET_NAME);
    sheet.appendRow(LOG_ANOMALI_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, LOG_ANOMALI_HEADERS);
  }
  return sheet;
}

/**
 * Sheet "Log_Unassigned".
 * @param {string} [username] - Opsional: username untuk menentukan facility
 * spreadsheet
 */
function getLogUnassignedSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_UNASSIGNED_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_UNASSIGNED_SHEET_NAME);
    sheet.appendRow(LOG_UNASSIGNED_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, LOG_UNASSIGNED_HEADERS);
  }
  return sheet;
}

/**
 * Sheet "Log_Bukti_Investigasi".
 * @param {string} [username] - Opsional: username untuk menentukan facility
 * spreadsheet
 */
function getLogBuktiSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_BUKTI_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_BUKTI_SHEET_NAME);
    sheet.appendRow(LOG_BUKTI_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, LOG_BUKTI_HEADERS);
  }
  return sheet;
}

/**
 * Sheet "Ringkasan_Harian".
 * @param {string} [username] - Opsional: username untuk menentukan facility
 * spreadsheet
 */
function getRingkasanHarianSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RINGKASAN_HARIAN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RINGKASAN_HARIAN_SHEET_NAME);
    sheet.appendRow(RINGKASAN_HARIAN_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, RINGKASAN_HARIAN_HEADERS);
  }
  return sheet;
}

/**
 * Sheet "Antrian_Aktif".
 * @param {string} [username] - Opsional: username untuk menentukan facility
 * spreadsheet
 */
function getAntrianAktifSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ANTRIAN_AKTIF_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ANTRIAN_AKTIF_SHEET_NAME);
    sheet.appendRow(ANTRIAN_AKTIF_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, ANTRIAN_AKTIF_HEADERS);
  }
  return sheet;
}

/**
 * Sheet "Lokasi_Aktif" - daftar lokasi aktif per-facility.
 * v8.26.0: sebelumnya lokasi aktif dibaca dari master_db sheet Master kolom
 * A (GLOBAL).
 * Sekarang lokasi aktif PER-FACILITY, disimpan di sheet "Lokasi_Aktif" di
 * masing-masing
 * spreadsheet operasional facility.
 *
 * @param {string} [username] - Opsional: username untuk menentukan facility
 * spreadsheet
 */
function getLokasiAktifSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOKASI_AKTIF_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOKASI_AKTIF_SHEET_NAME);
    sheet.appendRow(LOKASI_AKTIF_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, LOKASI_AKTIF_HEADERS);
  }
  return sheet;
}

/**
 * Sheet "Queue_Ringkasan_Delta" (v8.29.2 - lihat catatan di Config.gs &
 * QueueDelta.gs).
 * @param {string} [username] - Opsional: username untuk menentukan facility
 * spreadsheet
 */
function getQueueRingkasanSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(QUEUE_RINGKASAN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(QUEUE_RINGKASAN_SHEET_NAME);
    sheet.appendRow(QUEUE_RINGKASAN_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, QUEUE_RINGKASAN_HEADERS);
  }
  return sheet;
}

/**
 * Sheet "Queue_Antrian_Counter" (v8.29.2 - lihat catatan di Config.gs &
 * QueueDelta.gs).
 * @param {string} [username] - Opsional: username untuk menentukan facility
 * spreadsheet
 */
function getQueueCounterSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(QUEUE_COUNTER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(QUEUE_COUNTER_SHEET_NAME);
    sheet.appendRow(QUEUE_COUNTER_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, QUEUE_COUNTER_HEADERS);
  }
  return sheet;
}

/**
 * Sheet "Queue_Submit_Cycle" (v8.30.0 - lihat catatan di Config.gs &
 * SubmitQueue.gs).
 * SINGAJA TIDAK menerima parameter username / TIDAK pernah panggil
 * getOperasionalSpreadsheet_
 * -- sheet ini hidup di spreadsheet BOUND milik script ini sendiri
 * (getActiveSpreadsheet()),
 * bukan openById(), itulah kenapa submitCount() jadi cepat: tidak perlu buka
 * spreadsheet
 * facility sama sekali di jalur ini.
 */
function getSubmitQueueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SUBMIT_QUEUE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SUBMIT_QUEUE_SHEET_NAME);
    sheet.appendRow(SUBMIT_QUEUE_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, SUBMIT_QUEUE_HEADERS);
  }
  return sheet;
}

/**
 * Sheet "Queue_Validasi" (v8.31.0). Sama seperti getSubmitQueueSheet_ --
 * bound spreadsheet,
 * bukan facility, supaya submitValidasi() tidak perlu openById().
 */
function getValidasiQueueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(VALIDASI_QUEUE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(VALIDASI_QUEUE_SHEET_NAME);
    sheet.appendRow(VALIDASI_QUEUE_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, VALIDASI_QUEUE_HEADERS);
  }
  return sheet;
}

/**
 * Sheet "Queue_Task_Investigasi" (v8.31.0). Sama seperti
 * getSubmitQueueSheet_ -- bound
 * spreadsheet, bukan facility, supaya updateTaskStatus() tidak perlu openById
 * ().
 */
function getTaskQueueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TASK_QUEUE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TASK_QUEUE_SHEET_NAME);
    sheet.appendRow(TASK_QUEUE_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, TASK_QUEUE_HEADERS);
  }
  return sheet;
}

function getMasterSheet_() {
  // v8.25.0: Master sekarang di spreadsheet terpisah (Master DB) dengan
  // fallback otomatis
  // ke sheet lokal kalau Master DB bermasalah -- lihat ExternalDb.gs.
  // v8.26.0: Master TETAP di master_db (terpusat), TIDAK dipindah
  // per-facility.
  return getMasterSheetWithFallback_();
}

function fixHeaderForSheet_(sheet, headers) {
  const currentHeaderRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const isMatch = currentHeaderRow.every(function (val, idx) { return String(val).trim() === headers[idx]; });
  if (isMatch) {
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return;
  }
  const firstCell = String(sheet.getRange(1, 1).getValue()).trim();
  if (headers.indexOf(firstCell) === -1 && firstCell !== '') {
    sheet.insertRowBefore(1);
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function riwayatRowCachePut_(id, rowIndex) {
  // v8.20.1: TTL dinaikkan ke maksimum yang diizinkan CacheService (21600
  // detik = 6 jam),
  // dari sebelumnya cuma 1800 detik (30 menit). Dikombinasikan dengan
  // "refresh on access" di
  // findRiwayatRow_ (di-put ULANG tiap kali ketemu, baik lewat cache hit
  // maupun scan) dan
  // caching langsung saat baris dibuat (lihat appendRiwayatCycle_ di
  // CycleCount.gs) -- selama
  // sebuah baris masih aktif dikerjakan (disentuh minimal 1x per 6 jam),
  // pencariannya SELALU
  // cache hit O(1), tidak pernah perlu scan kolom ID penuh lagi --
  // performanya jadi TIDAK
  // tergantung berapa banyak baris histori lama yang sudah menumpuk di sheet
  // Riwayat.
  try { CacheService.getScriptCache().put('riw_' + id, String(rowIndex), 21600); } catch (e) { /* abaikan */ }
}

function riwayatRowCacheGet_(id) {
  try {
    const v = CacheService.getScriptCache().get('riw_' + id);
    return v ? Number(v) : -1;
  } catch (e) { return -1; }
}

function findRiwayatRow_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const cachedRow = riwayatRowCacheGet_(id);
  if (cachedRow >= 2 && cachedRow <= lastRow) {
    if (String(sheet.getRange(cachedRow, 1).getValue()) === String(id)) {
      riwayatRowCachePut_(id, cachedRow); // refresh TTL -- selama masih aktif diakses, tidak akan pernah basi
      return cachedRow;
    }
  }

  // fallback: scan kolom ID (cuma kejadian kalau cache belum ada/sudah basi
  // -- lihat catatan di riwayatRowCachePut_)
  const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]) === String(id)) {
      const rowIndex = i + 2;
      riwayatRowCachePut_(id, rowIndex);
      return rowIndex;
    }
  }
  return -1;
}
