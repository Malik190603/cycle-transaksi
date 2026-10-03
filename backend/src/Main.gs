/**
 * Main.gs
 * Entry point WebApp (doGet), trigger (onOpen/onEdit), dan pengecekan role (requireRole_).
 * (Hasil pemecahan dari Code.gs -- semua file .gs berbagi global scope yang sama di Apps Script,
 *  jadi urutan file tidak masalah, fungsi di file lain tetap bisa saling panggil.)
 */

function doGet(e) {
  // Tampilan sekarang ada di aplikasi Android (APK); Web App ini hanya melayani doPost.
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;padding:24px;line-height:1.5">' +
    '<h3>Cycle Transaksi \u2014 server aktif (' + APP_VERSION + ')</h3>' +
    '<p>Halaman ini adalah server aplikasi Android Cycle Transaksi. Alamat halaman ini (yang berakhiran <b>/exec</b>) ' +
    'ditanam di aplikasi oleh pengembang; pengguna cukup masuk dengan NIK.</p></div>')
    .setTitle('Cycle Transaksi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Dipanggil dari scriptlet <?!= include('NamaFile'); ?> di Index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAppVersion() {
  return APP_VERSION;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cycle Count')
    .addItem('Setup Awal (buat semua sheet, facility & admin pertama)', 'setupAwal')
    .addItem('Setup Setting Level Rack (Master)', 'setupMasterLevelSettings')
    .addItem('Setup Mapping PIC Kategori Selisih (Master)', 'setupMasterAssignmentMapping')
    .addItem('Setup Config Overflow (Config_Sistem & Config_Akses_Setting)', 'setupOverflowConfigSheets')
    .addItem('Setup Master DB & Config DB (Pindah ke Spreadsheet Terpisah)', 'setupExternalDbMasterDanConfig')
    .addItem('Setup Facility Awal (Seed RDC Makassar dari sistem lama)', 'seedFacilityAwalDariSistemLama')
    .addItem('Perbaiki Header (Data Count & Riwayat)', 'perbaikiHeader')
    .addItem('Cek & Perbaiki Waktu_Input Kosong', 'cekPerbaikiWaktuInputKosong')
    .addItem('Cek & Jalankan Archive Riwayat Lama', 'cekArchiveRiwayatLama')
    .addItem('Rebuild Ringkasan Harian (kalau Dashboard terasa tidak sinkron)', 'rebuildRingkasanHarianMenu')
    .addItem('Rebuild Antrian Aktif (kalau Home terasa tidak sinkron)', 'rebuildAntrianAktifMenu')
    .addItem('Reset Semua Hasil Count Hari Ini', 'resetCount')
    .addToUi();
}

function onEdit(e) {
  try {
    if (e && e.range && e.range.getSheet().getName() === MASTER_SHEET_NAME) {
      clearMasterCache_();
    }
    if (e && e.range && e.range.getSheet().getName() === CONFIG_SISTEM_SHEET_NAME) {
      clearConfigSistemCache_();
    }
    if (e && e.range && e.range.getSheet().getName() === CONFIG_AKSES_SETTING_SHEET_NAME) {
      clearNikAksesSettingCache_();
    }
  } catch (err) {}
}

// ---------- Diagnostik manual (jalankan dari Apps Script editor, cek hasilnya di
// View -> Logs atau Execution log) ----------

function debugCekAksesConfig() {
  const nik = SETUP_ADMIN_NIK_;
  Logger.log('--- Cek Script Properties ---');
  Logger.log('MASTER_DB_ID: ' + PropertiesService.getScriptProperties().getProperty('MASTER_DB_ID'));
  Logger.log('CONFIG_DB_ID: ' + PropertiesService.getScriptProperties().getProperty('CONFIG_DB_ID'));

  Logger.log('--- Cek Master DB sheet Master ---');
  const masterSheet = getMasterSheet_();
  Logger.log('getMasterSheet_() ketemu: ' + !!masterSheet + (masterSheet ? (' (di spreadsheet: ' + masterSheet.getParent().getUrl() + ')') : ''));

  Logger.log('--- Cek sheet Config_Akses_Setting ---');
  const aksesSheet = getConfigAksesSettingSheet_();
  Logger.log('getConfigAksesSettingSheet_() di spreadsheet: ' + aksesSheet.getParent().getUrl());
  Logger.log('Isi baris 2 dst: ' + JSON.stringify(aksesSheet.getRange(2, 1, Math.max(aksesSheet.getLastRow() - 1, 0), 3).getValues()));

  Logger.log('--- Cek isNikPunyaAksesSetting_("' + nik + '") ---');
  Logger.log('Hasil: ' + isNikPunyaAksesSetting_(nik));

  Logger.log('--- Cek getUserRole("' + nik + '") lengkap ---');
  Logger.log(JSON.stringify(getUserRole(nik)));
}

function clearMasterCache_() {
  resetMemEksekusi_();
  const kunci = [
    'activeLocations',
    'masterUsernames',
    'assignableUsers',
    'levelSettings',
    'kategoriMapping',
    'inventoryUsers'
  ];
  // Daftar petugas & validator di-cache 5 menit PER FACILITY. Tanpa ikut dibuang di sini, user
  // yang baru ditambah (atau perannya baru diganti) belum muncul di Upload Data sampai 5 menit.
  try {
    const fs = getFacilitySheet_();
    const last = fs.getLastRow();
    if (last >= 2) {
      fs.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
        const id = String(r[0] || '').trim();
        if (id) { kunci.push('assignableUsers_fac_' + id); kunci.push('inventoryUsers_fac_' + id); }
      });
    }
  } catch (e) { /* sheet Facility belum ada: tidak ada cache per facility */ }
  CacheService.getScriptCache().removeAll(kunci);
  // v8.26.0: Bersihkan juga cache facility management
  clearFacilityCache_();
}

// =========================================================================
// SECURITY
// =========================================================================

function requireRole_(username, allowedRoles) {
  const info = getUserRole(username);

  if (!info) {
    throw new Error('User tidak dikenali. Silakan login ulang.');
  }

  if (allowedRoles && allowedRoles.indexOf(info.role) === -1) {
    throw new Error('Akses ditolak untuk role "' + info.role + '".');
  }

  return info;
}

// =========================================================================
// API ROUTER (GitHub Pages)
// =========================================================================
// CATATAN PERBAIKAN: sebelumnya doPost() cuma punya 1 case ("login"), semua
// action lain (24 fungsi backend yang sebenarnya dipakai frontend lewat
// google.script.run di HtmlService) balik "Unknown action". Daftar di bawah
// disusun dengan menelusuri SEMUA pemanggilan google.script.run.xxx(...) di
// html/*.html (Body1/JsCore/JsAdmin/JsCycleValidasi1/JsDashboard/
// JsTaskInvestigasi) supaya name action & urutan parameter PERSIS sama
// dengan fungsi .gs yang sudah ada -- tidak ada logic baru, murni routing.
//
// Kontrak response: setiap fungsi yang SUDAH mengembalikan object berisi
// field "success" (mis. submitCount, submitValidasi, updateTaskStatus,
// importRawData, closePlusMinusPair) diteruskan APA ADANYA. Fungsi lain
// (murni fetch data, tidak ada field "success") dibungkus jadi
// { success: true, data: <hasil> } supaya kontrak response tetap konsisten
// untuk konsumen REST API (client GitHub Pages).
//
// PENTING: request body JSON harus punya field yang namanya SAMA dengan
// nama parameter di bawah (lihat masing-masing case). Kalau frontend GitHub
// Pages yang sedang/sudah dibangun pakai nama field lain, sesuaikan
// pemetaan req.xxx di bawah ini (BUKAN nama fungsi .gs-nya).

function apiResult_(raw) {
  if (raw && typeof raw === 'object' && !Array.isArray(raw) && Object.prototype.hasOwnProperty.call(raw, 'success')) {
    return raw;
  }
  return { success: true, data: raw };
}

const API_ACTIONS = {
  // ---------- Login & versi ----------
  login: function (req) {
    const info = getUserRole(req.username);
    if (info) info.aksesSettingOverflow = isNikPunyaAksesSetting_(req.username);
    return info;
  },
  getAppVersion: function () { return getAppVersion(); },

  // ---------- Home ----------
  getHomeSummary: function (req) { return getHomeSummary(req.username); },
  getPlusMinusSummary: function (req) { return getPlusMinusSummary(req.username); },
  getMyPendingCount: function (req) { return getMyPendingCount(req.username); },
  getPendingValidasiCount: function (req) { return getPendingValidasiCount(req.username); },

  // ---------- Cycle Transaksi ----------
  getMyPendingTasks: function (req) { return getMyPendingTasks(req.username); },
  submitCount: function (req) { return submitCount(req.no, req.username, req.qty); },

  // ---------- Validasi ----------
  getPendingValidasi: function (req) { return getPendingValidasi(req.username); },
  submitValidasi: function (req) { return submitValidasi(req.id, req.username, req.qty); },

  // ---------- Admin: Upload & Import ----------
  getEquipmentReadyDefaults: function (req) { return getEquipmentReadyDefaults(req.username); },
  getAssignableUsers: function (req) { return getAssignableUsers(req.username); },
  importRawData: function (req) {
    return importRawData(req.transaksiRawRows, req.stockRows, req.selectedUsernames, req.username, req.reachTruckReady, req.tanggaReady, req.jobId);
  },
  getImportProgress: function (req) { return getImportProgress(req.jobId); },

  // ---------- Dashboard ----------
  getDashboardData: function (req) {
    return getDashboardData(req.periodType, req.periodValue, req.trendCount, req.username);
  },
  getErrorAnalysisDetail: function (req) {
    return getErrorAnalysisDetail(req.alasan, req.periodType, req.periodValue, req.username);
  },
  getUserDashboardDetail: function (req) {
    return getUserDashboardDetail(req.nama, req.periodType, req.periodValue, req.username);
  },
  getPendingBacklog: function (req) { return getPendingBacklog(req.username); },
  getBacklogDetailByDate: function (req) { return getBacklogDetailByDate(req.tanggal, req.username); },
  getProblemItemsDetail: function (req) { return getProblemItemsDetail(req.periodType, req.periodValue, req.username); },

  // ---------- Analytics: Accuracy Root Cause (v8.28.0) ----------
  getAnalyticsRootCauseData: function (req) {
    return getAnalyticsRootCauseData(req.dateFrom, req.dateTo, req.username);
  },
  getAnalyticsRootCauseDetail: function (req) {
    return getAnalyticsRootCauseDetail(req.rootCauseLabel, req.dateFrom, req.dateTo, req.username);
  },

  // ---------- Task Investigasi / Verifikasi ----------
  getInvestigasiFormData: function (req) { return getInvestigasiFormData(req.username); },
  getOpenTasks: function (req) { return getOpenTasks(req.username); },
  updateTaskStatus: function (req) {
    return updateTaskStatus(req.id, req.username, req.newStatus, req.catatan, req.kategori, req.picUsername, req.buktiRows);
  },
  getTaskLog: function (req) { return getTaskLog(req.limit, req.username, req.dateFrom, req.dateTo); },

  // ---------- Penyelesaian Plus Minus ----------
  getPlusMinusCandidates: function (req) { return getPlusMinusCandidates(req.username); },
  closePlusMinusPair: function (req) {
    return closePlusMinusPair(req.minusTaskId, req.plusTaskId, req.username, req.catatan, req.picUsername, req.buktiRows);
  },

  // ---------- Setting Overflow (v8.24.0) ----------
  // Semua 3 action ini requireRole_-nya di dalam masing-masing fungsi (OverflowConfig.gs)
  // via isNikPunyaAksesSetting_ -- BUKAN requireRole_(['admin']) biasa, karena akses
  // Setting ini berbasis daftar NIK terpisah (Config_Akses_Setting), bukan role Master.
  getSettingOverflow: function (req) { return getSettingOverflow(req.username); },
  saveSettingOverflow: function (req) {
    return saveSettingOverflow(req.username, req.overflowEnabled, req.minSisaReachTruck, req.maksOrang);
  },
  setOverflowForceOffToday: function (req) { return setOverflowForceOffToday(req.username); },

  // ---------- Kelola User & Akses Setting (v8.25.0) ----------
  // v8.25.0.1: 7 action ini sebelumnya KELEWATAN saat pertama kali menu Kelola User &
  // Akses Setting dibuat -- akibatnya kalau dipanggil lewat doPost/REST (bukan
  // google.script.run langsung dari HtmlService web app), selalu balik "Unknown action".
  getDaftarUserMaster: function (req) { return getDaftarUserMaster(req.username); },
  tambahUserMaster: function (req) { return tambahUserMaster(req.username, req.usernameBaru, req.roleBaru); },
  updateRoleUserMaster: function (req) {
    return updateRoleUserMaster(req.username, req.rowIndex, req.usernameKonfirmasi, req.roleBaru);
  },
  setStatusUserMaster: function (req) {
    return setStatusUserMaster(req.username, req.rowIndex, req.usernameKonfirmasi, req.statusBaru);
  },
  getDaftarAksesSetting: function (req) { return getDaftarAksesSetting(req.username); },
  tambahAksesSetting: function (req) { return tambahAksesSetting(req.username, req.nikBaru, req.namaBaru); },
  hapusAksesSetting: function (req) { return hapusAksesSetting(req.username, req.rowIndex, req.nikKonfirmasi); },

  // ---------- Pembagian Tugas Matrix (v8.26.0 Tahap 1) ----------
  getLevelAssignmentConfig: function (req) { return getLevelAssignmentConfig(req.username); },
  saveLevelAssignmentMatrix: function (req) {
    return saveLevelAssignmentMatrix(req.username, req.matrixBaru, req.maksGrupAlatBaru);
  },
  setModeAssignment: function (req) { return setModeAssignment(req.username, req.mode); },
  getLogPerubahanConfig: function (req) { return getLogPerubahanConfig(req.username); },

  // ---------- Facility Management (v8.26.0 Tahap 2) ----------
  getDaftarFacility: function (req) { return getDaftarFacility(req.username); },
  tambahFacility: function (req) {
    return tambahFacility(req.username, req.namaFacility, req.kodeFacility, req.namaSpreadsheet, req.daftarLokasi, req.jobId);
  },
  // v8.29.0: daftarkan facility dari spreadsheet yang SUDAH ADA (data lama, mis. Makassar) --
  // beda dari tambahFacility() di atas yang selalu bikin spreadsheet kosong baru.
  daftarkanFacilityExisting: function (req) {
    return daftarkanFacilityExisting(req.username, req.namaFacility, req.kodeFacility, req.spreadsheetIdAtauUrl, req.daftarLokasi, req.jobId);
  },
  updateNamaFacility: function (req) {
    return updateNamaFacility(req.username, req.facilityId, req.namaBaru, req.kodeBaru);
  },
  setStatusFacility: function (req) {
    return setStatusFacility(req.username, req.facilityId, req.statusBaru);
  },
  assignUserKeFacility: function (req) {
    return assignUserKeFacility(req.username, req.targetUsername, req.facilityId);
  },
  getDaftarUserFacilityAssignment: function (req) {
    return getDaftarUserFacilityAssignment(req.username);
  },
  importLokasiAktif: function (req) {
    return importLokasiAktif(req.username, req.facilityId, req.daftarLokasi, req.gantiSemua);
  },
  getDaftarLokasiAktif: function (req) {
    return getDaftarLokasiAktif(req.username, req.facilityId);
  },
  copyLokasiDariFacility: function (req) {
    return copyLokasiDariFacility(req.username, req.facilityIdSumber, req.facilityIdTarget, req.gantiSemua);
  },

  // ---------- Facility Switcher Developer (v8.26.2) ----------
  setDeveloperActiveFacility: function (req) {
    return setDeveloperActiveFacility(req.username, req.facilityId);
  }
};

function doPost(e) {
  try {
    resetMemEksekusi_();
    const req = JSON.parse(e.postData.contents || '{}');
    // Aplikasi Android mengirim { action, args: [...] } -- lihat ApiBridge.gs
    const viaApk = apiBridge_(req);
    if (viaApk) return viaApk;
    const action = req.action;
    const handler = API_ACTIONS[action];

    if (!handler) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          message: 'Unknown action: ' + action
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const result = apiResult_(handler(req));

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
