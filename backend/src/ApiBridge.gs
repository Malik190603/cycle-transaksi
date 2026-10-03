/**
 * ApiBridge.gs — pintu masuk aplikasi Android (APK) Cycle Transaksi.
 *
 * APK mengirim POST ke Web App ini dengan badan JSON:
 *   { "action": "<nama fungsi>", "args": [ ...argumen berurutan... ] }
 * dan menerima:
 *   { "ok": true, "result": <hasil fungsi> }   atau   { "ok": false, "error": "<pesan>" }
 *
 * Fungsi dipanggil dengan argumen yang SAMA PERSIS seperti google.script.run di web app
 * (urutan parameter fungsi .gs), jadi tidak ada pemetaan nama field yang bisa meleset.
 * Hanya fungsi di daftar API_BRIDGE_ALLOW_ yang boleh dipanggil.
 *
 * doPost() di Main.gs memanggil apiBridge_(req) lebih dulu; permintaan tanpa "args" diteruskan
 * ke router lama (API_ACTIONS).
 */

const API_BRIDGE_ALLOW_ = [
  // Login, versi, Home
  'getAppVersion', 'getUserRole', 'getHomeSummary', 'getPlusMinusSummary', 'getMyPendingCount', 'getPendingValidasiCount',
  // Cycle Transaksi & Validasi
  'getMyPendingTasks', 'submitCount', 'getPendingValidasi', 'submitValidasi',
  // Upload Data
  'getEquipmentReadyDefaults', 'getAssignableUsers', 'importRawData', 'getImportProgress',
  // Dashboard & Analytics
  'getDashboardData', 'getErrorAnalysisDetail', 'getUserDashboardDetail', 'getPendingBacklog', 'getBacklogDetailByDate',
  'getProblemItemsDetail', 'getAnalyticsRootCauseData', 'getAnalyticsRootCauseDetail',
  // Verifikasi & Penyelesaian Plus Minus
  'getInvestigasiFormData', 'getOpenTasks', 'updateTaskStatus', 'getTaskLog', 'getPlusMinusCandidates', 'closePlusMinusPair',
  // Config: overflow, user, akses, pembagian tugas
  'getSettingOverflow', 'saveSettingOverflow', 'setOverflowForceOffToday',
  'getDaftarUserMaster', 'tambahUserMaster', 'updateRoleUserMaster', 'setStatusUserMaster',
  'getDaftarAksesSetting', 'tambahAksesSetting', 'hapusAksesSetting',
  'getLevelAssignmentConfig', 'saveLevelAssignmentMatrix', 'setModeAssignment', 'getLogPerubahanConfig',
  // Facility Management
  'getDaftarFacility', 'tambahFacility', 'daftarkanFacilityExisting', 'updateNamaFacility', 'setStatusFacility',
  'assignUserKeFacility', 'getDaftarUserFacilityAssignment', 'importLokasiAktif', 'getDaftarLokasiAktif',
  'copyLokasiDariFacility', 'setDeveloperActiveFacility',
  // Aplikasi v2: satu layar, satu panggilan (AppApi.gs)
  'getHomeBundle', 'getUploadFormData', 'getProductivity', 'getOpenTaskCount'
];

/**
 * Mengembalikan TextOutput JSON bila permintaan berasal dari APK (punya "args" berupa array),
 * atau null supaya doPost melanjutkan ke router lama (API_ACTIONS).
 */
function apiBridge_(req) {
  if (!req || !Array.isArray(req.args)) return null;
  const name = String(req.action || '');
  let out;
  if (API_BRIDGE_ALLOW_.indexOf(name) === -1) {
    out = { ok: false, error: 'Fungsi tidak dikenal: ' + name };
  } else {
    const scope = (typeof globalThis !== 'undefined') ? globalThis : this;
    const fn = scope[name];
    if (typeof fn !== 'function') {
      out = { ok: false, error: 'Fungsi "' + name + '" tidak ada di backend ini.' };
    } else {
      try {
        const result = fn.apply(null, req.args);
        out = { ok: true, result: result === undefined ? null : result };
      } catch (err) {
        out = { ok: false, error: (err && err.message) ? err.message : String(err) };
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
