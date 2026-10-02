/**
 * Setup.gs — penyiapan pertama untuk spreadsheet baru (satu spreadsheet untuk semuanya).
 *
 * CARA PAKAI: di editor Apps Script pilih fungsi "setupAwal" lalu klik Run (sekali saja).
 * Aman dijalankan ulang: sheet/baris yang sudah ada tidak dibuat dua kali dan isinya tidak dihapus.
 *
 * Yang dibuat:
 * 1. Semua sheet beserta header: Master, Master_User, Facility, User_Facility_Assignment,
 *    Config_Akses_Setting, Config_Sistem, Config_Level_Assignment, Log_Perubahan_Config, Log_Sistem,
 *    seluruh sheet operasional facility (Data Count, Riwayat, Lokasi_Aktif, ringkasan, antrean, log),
 *    dan tiga sheet antrean tulis (Queue_Submit_Cycle, Queue_Validasi, Queue_Task_Investigasi).
 * 2. Facility pertama yang menunjuk ke spreadsheet ini sendiri, dan user admin pertama.
 * 3. Satu trigger tiap 1 menit (workerSemua) yang memproses semua antrean.
 *
 * Isi tiga konstanta di bawah SEBELUM menjalankan setupAwal (NIK admin tidak disimpan di repo publik).
 */
const SETUP_FACILITY_NAMA_ = 'ISI NAMA FACILITY';
const SETUP_FACILITY_KODE_ = 'ISI-KODE';
const SETUP_ADMIN_NIK_ = 'ISI.NIK.ADMIN';

function setupEnsureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  let dibuat = false;
  if (!sheet) { sheet = ss.insertSheet(name); dibuat = true; }
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return { sheet: sheet, dibuat: dibuat };
}

/**
 * Kolom tanggal (yyyy-MM-dd), SKU, dan username disimpan sebagai TEKS. Tanpa ini Google Sheets
 * mengubah "2026-10-02" menjadi sel tanggal dan "00123" menjadi angka 123, sedangkan kode
 * membandingkan nilai-nilai itu sebagai string (ringkasan harian, backlog per tanggal, log akses).
 * Dipanggil untuk setiap spreadsheet yang dipakai sistem (setup awal & facility baru).
 */
function terapkanFormatKolomTeks_(ss) {
  const kolom = [
    [SHEET_NAME, ['B2:B', 'E2:E']],
    [RIWAYAT_SHEET_NAME, ['B2:B', 'D2:D']],
    [RIWAYAT_ARCHIVE_SHEET_NAME, ['B2:B', 'D2:D']],
    [RINGKASAN_HARIAN_SHEET_NAME, ['A2:A']],
    [QUEUE_RINGKASAN_SHEET_NAME, ['B2:B']],
    [LOG_AKSES_SHEET_NAME, ['A2:A', 'B2:B']],
    [LOKASI_AKTIF_SHEET_NAME, ['A2:A']],
    [MASTER_USER_SHEET_NAME, ['A2:A']]
  ];
  kolom.forEach(function (k) {
    const sheet = ss.getSheetByName(k[0]);
    if (!sheet) return;
    k[1].forEach(function (a1) { sheet.getRange(a1).setNumberFormat('@'); });
  });
}

function setupCariBaris_(sheet, kolom, nilai) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const target = String(nilai).trim().toLowerCase();
  const values = sheet.getRange(2, kolom, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toLowerCase() === target) return i + 2;
  }
  return 0;
}

function setupAwal() {
  if (/^ISI[ .-]/.test(SETUP_ADMIN_NIK_) || /^ISI[ .-]/.test(SETUP_FACILITY_NAMA_) || /^ISI[ .-]/.test(SETUP_FACILITY_KODE_)) {
    throw new Error('Isi dulu SETUP_FACILITY_NAMA_, SETUP_FACILITY_KODE_, dan SETUP_ADMIN_NIK_ di bagian Setup, lalu jalankan lagi.');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hasil = [];
  const now = new Date();
  let jumlahBaru = 0;

  // ---------- 1. Sheet pusat ----------
  const master = setupEnsureSheet_(ss, MASTER_SHEET_NAME,
    ['ACTIVE LOCATION', '', 'USER', 'ROLE', 'STATUS', 'LEVEL GROUP', 'PRODUKTIVITAS (qty/orang)', '', 'KATEGORI SELISIH', 'USER PIC', 'ROLE PIC']);
  const m = master.sheet;
  if (master.dibuat) jumlahBaru++;
  if (!String(m.getRange('F2').getValue() || '').trim()) {
    // Produktivitas per grup level (dibaca getLevelSettings_ lewat teks "1-2"/"3-4"/"5-6") dan jumlah
    // alat siap bawaan untuk layar Upload (dibaca lewat teks "reach truck"/"tangga").
    const rows = LEVEL_GROUPS.map(function (g) { return ['Level ' + g.min + '-' + g.max, g.defaultProd]; });
    rows.push(['Reach Truck Ready (unit)', 0]);
    rows.push(['Tangga Pesawat Ready (unit)', 0]);
    m.getRange(2, 6, rows.length, 2).setValues(rows);
  }
  if (!String(m.getRange('I2').getValue() || '').trim()) {
    m.getRange(2, 9, TASK_KATEGORI_LIST.length, 1).setValues(TASK_KATEGORI_LIST.map(function (k) { return [k]; }));
  }
  if (!setupCariBaris_(m, 3, SETUP_ADMIN_NIK_)) {
    const lastUserRow = Math.max(1, m.getRange('C:C').getValues().filter(function (r) { return String(r[0]).trim(); }).length);
    m.getRange(lastUserRow + 1, 3, 1, 3).setValues([[SETUP_ADMIN_NIK_, 'admin', 'Aktif']]);
  }

  const pusat = [
    [MASTER_USER_SHEET_NAME, MASTER_USER_HEADERS],
    [FACILITY_SHEET_NAME, FACILITY_HEADERS],
    [USER_FACILITY_ASSIGNMENT_SHEET_NAME, USER_FACILITY_ASSIGNMENT_HEADERS],
    [CONFIG_AKSES_SETTING_SHEET_NAME, CONFIG_AKSES_SETTING_HEADERS],
    [LOG_PERUBAHAN_CONFIG_SHEET_NAME, LOG_PERUBAHAN_CONFIG_HEADERS],
    [LOG_SISTEM_SHEET_NAME, LOG_SISTEM_HEADERS],
    [SUBMIT_QUEUE_SHEET_NAME, SUBMIT_QUEUE_HEADERS],
    [VALIDASI_QUEUE_SHEET_NAME, VALIDASI_QUEUE_HEADERS],
    [TASK_QUEUE_SHEET_NAME, TASK_QUEUE_HEADERS]
  ];
  const semua = pusat.concat(FACILITY_OPERATIONAL_SHEETS.map(function (d) { return [d.name, d.headers]; }));
  const sheets = {};
  semua.forEach(function (d) {
    const r = setupEnsureSheet_(ss, d[0], d[1]);
    sheets[d[0]] = r.sheet;
    if (r.dibuat) jumlahBaru++;
  });
  terapkanFormatKolomTeks_(ss);
  hasil.push(jumlahBaru + ' sheet baru dibuat (total ' + ss.getSheets().length + ' sheet).');

  // ---------- 2. Facility pertama: menunjuk ke spreadsheet ini sendiri ----------
  const facSheet = sheets[FACILITY_SHEET_NAME];
  let facRow = setupCariBaris_(facSheet, 5, ss.getId());
  let facId;
  if (facRow) {
    facId = String(facSheet.getRange(facRow, 1).getValue()).trim();
    hasil.push('Facility sudah ada: ' + facSheet.getRange(facRow, 2).getValue() + ' (' + facId + ').');
  } else {
    facId = 'FAC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    facSheet.appendRow([facId, SETUP_FACILITY_NAMA_, SETUP_FACILITY_KODE_, ss.getName(), ss.getId(), 'Aktif', now, SETUP_ADMIN_NIK_, ss.getUrl()]);
    hasil.push('Facility dibuat: ' + SETUP_FACILITY_NAMA_ + ' (' + facId + ').');
  }

  // ---------- 3. Admin pertama + akses Config ----------
  const userSheet = sheets[MASTER_USER_SHEET_NAME];
  if (!setupCariBaris_(userSheet, 1, SETUP_ADMIN_NIK_)) {
    userSheet.appendRow([SETUP_ADMIN_NIK_, 'admin', 'Aktif', facId]);
    hasil.push('User admin dibuat: ' + SETUP_ADMIN_NIK_ + '.');
  } else {
    hasil.push('User ' + SETUP_ADMIN_NIK_ + ' sudah ada.');
  }
  // Worker ringkasan mencari satu user per facility di sheet ini untuk menemukan spreadsheet-nya
  const assignSheet = sheets[USER_FACILITY_ASSIGNMENT_SHEET_NAME];
  if (!setupCariBaris_(assignSheet, 1, SETUP_ADMIN_NIK_)) {
    assignSheet.appendRow([SETUP_ADMIN_NIK_, facId, now, SETUP_ADMIN_NIK_]);
  }
  const aksesSheet = sheets[CONFIG_AKSES_SETTING_SHEET_NAME];
  if (!setupCariBaris_(aksesSheet, 1, SETUP_ADMIN_NIK_)) {
    aksesSheet.appendRow([SETUP_ADMIN_NIK_, 'Admin pertama', Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd')]);
  }

  // ---------- 4. Sheet konfigurasi dengan nilai bawaan ----------
  getConfigSistemSheet_();
  getLevelAssignmentSheet_();

  // ---------- 5. Sheet kosong bawaan Google ----------
  ['Sheet1', 'Lembar1'].forEach(function (nama) {
    const kosong = ss.getSheetByName(nama);
    if (kosong && kosong.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(kosong);
  });

  // ---------- 6. Trigger pemroses antrean ----------
  hasil.push(pasangTriggerWorker());

  try { clearMasterCache_(); } catch (e) { /* abaikan */ }
  hasil.push('Lokasi_Aktif masih kosong: isi kolom A sheet "Lokasi_Aktif" (satu lokasi per baris) atau impor lewat menu Config di aplikasi.');
  hasil.push('Langkah berikutnya: Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone), lalu salin URL /exec ke aplikasi.');

  const pesan = hasil.join('\n');
  Logger.log(pesan);
  try { SpreadsheetApp.getUi().alert('Setup Awal selesai:\n\n' + pesan); } catch (e) { /* dijalankan dari editor: lihat Execution log */ }
  return pesan;
}

// =========================================================================
// WORKER ANTREAN
// =========================================================================
// Satu trigger untuk keempat antrean (bukan empat trigger terpisah): urutan prosesnya pasti
// (hitung -> validasi -> task -> ringkasan), dan jatah waktu trigger harian akun Google gratis
// (90 menit/hari) tidak habis oleh empat eksekusi kosong tiap menit.
function workerSemua() {
  const langkah = [
    ['Submit', processSubmitQueue_],
    ['Validasi', processValidasiQueue_],
    ['Task', processTaskQueue_],
    ['Ringkasan', processDeferredQueueAllFacilities_]
  ];
  langkah.forEach(function (l) {
    try { l[1](); } catch (e) { catatLogSistem_('Worker ' + l[0], 'ERROR: ' + ((e && e.message) || e)); }
  });
}

function pasangTriggerWorker() {
  const lama = ['workerSemua', 'processSubmitQueue_', 'processValidasiQueue_', 'processTaskQueue_', 'processDeferredQueueAllFacilities_'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (lama.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('workerSemua').timeBased().everyMinutes(1).create();
  return 'Trigger terpasang: workerSemua tiap 1 menit.';
}

// Menu spreadsheet tidak bisa memanggil fungsi berakhiran "_" -- tiga pembungkus publik ini
// yang dipanggil menu "Cycle Count" (lihat onOpen di Main.gs).
function cekArchiveRiwayatLama() { return cekArchiveRiwayatLama_(); }
function rebuildRingkasanHarianMenu() { return rebuildRingkasanHarianMenu_(); }
function rebuildAntrianAktifMenu() { return rebuildAntrianAktifMenu_(); }
