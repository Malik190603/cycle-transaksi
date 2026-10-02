/**
 * FacilityManagement.gs
 * v8.26.0 -- Manajemen Facility (RDC Makassar, RDC Manado, RDC Kendari, dst.)
 *
 * Arsitektur:
 *   - master_db   : tetap terpusat, hanya menyimpan Master User
 *   - config_db   : tetap terpusat, DITAMBAH sheet "Facility" & "User_Facility_Assignment"
 *   - Setiap Facility memiliki spreadsheet operasional SENDIRI (terisolasi)
 *
 * Sheet "Facility" di config_db:
 *   Kolom: ID_Facility, Nama_Facility, Kode_Facility, Nama_Spreadsheet,
 *          Spreadsheet_ID, Status, Tanggal_Dibuat, Dibuat_Oleh, URL_Spreadsheet
 *
 * Sheet "User_Facility_Assignment" di config_db:
 *   Kolom: Username, ID_Facility, Tanggal_Diassign, Diassign_Oleh
 *
 * Sheet "Lokasi_Aktif" di setiap spreadsheet facility:
 *   Kolom: Lokasi (satu lokasi per baris)
 */

// =====================================================================
// KONSTANTA
// =====================================================================
const FACILITY_SHEET_NAME = 'Facility';
const FACILITY_HEADERS = ['ID_Facility', 'Nama_Facility', 'Kode_Facility', 'Nama_Spreadsheet',
  'Spreadsheet_ID', 'Status', 'Tanggal_Dibuat', 'Dibuat_Oleh', 'URL_Spreadsheet'];

const USER_FACILITY_ASSIGNMENT_SHEET_NAME = 'User_Facility_Assignment';
const USER_FACILITY_ASSIGNMENT_HEADERS = ['Username', 'ID_Facility', 'Tanggal_Diassign', 'Diassign_Oleh'];

const LOKASI_AKTIF_SHEET_NAME = 'Lokasi_Aktif';
const LOKASI_AKTIF_HEADERS = ['Lokasi'];

// Daftar sheet operasional yang harus ada di setiap facility spreadsheet
const FACILITY_OPERATIONAL_SHEETS = [
  { name: SHEET_NAME, headers: HEADERS },
  { name: RIWAYAT_SHEET_NAME, headers: RIWAYAT_HEADERS },
  { name: LOG_AKSES_SHEET_NAME, headers: LOG_AKSES_HEADERS },
  { name: LOG_ANOMALI_SHEET_NAME, headers: LOG_ANOMALI_HEADERS },
  { name: LOG_UNASSIGNED_SHEET_NAME, headers: LOG_UNASSIGNED_HEADERS },
  { name: LOG_BUKTI_SHEET_NAME, headers: LOG_BUKTI_HEADERS },
  { name: RINGKASAN_HARIAN_SHEET_NAME, headers: RINGKASAN_HARIAN_HEADERS },
  { name: ANTRIAN_AKTIF_SHEET_NAME, headers: ANTRIAN_AKTIF_HEADERS },
  { name: RIWAYAT_ARCHIVE_SHEET_NAME, headers: RIWAYAT_HEADERS },
  { name: LOG_SISTEM_SHEET_NAME, headers: LOG_SISTEM_HEADERS },
  { name: LOKASI_AKTIF_SHEET_NAME, headers: LOKASI_AKTIF_HEADERS },
  // v8.29.2: sheet antrian delta (performa submitCount) -- lihat Config.gs & QueueDelta.gs
  { name: QUEUE_RINGKASAN_SHEET_NAME, headers: QUEUE_RINGKASAN_HEADERS },
  { name: QUEUE_COUNTER_SHEET_NAME, headers: QUEUE_COUNTER_HEADERS }
];

// Cache key untuk facility lookup
const CACHE_FACILITY_LIST = 'facilityList';
const CACHE_USER_FACILITY_PREFIX = 'userFacility_';

// v8.26.2: Cache key untuk facility switcher Developer (lihat setDeveloperActiveFacility()).
// TTL 21600 detik (6 jam, maksimum CacheService) -- otomatis "lupa" kalau Developer tidak
// aktif lama, balik ke assignment normalnya.
const CACHE_DEV_ACTIVE_FACILITY_PREFIX = 'devActiveFacility_';
const DEV_ACTIVE_FACILITY_TTL_SEC = 21600;

// =====================================================================
// HELPER: AKSES SHEET FACILITY DI config_db
// =====================================================================

/**
 * Sheet "Facility" di config_db. Kalau belum ada, dibuat otomatis.
 */
function getFacilitySheet_() {
  const ss = getConfigTargetSpreadsheet_();
  let sheet = ss.getSheetByName(FACILITY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FACILITY_SHEET_NAME);
    sheet.appendRow(FACILITY_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, FACILITY_HEADERS);
  }
  return sheet;
}

/**
 * Sheet "User_Facility_Assignment" di config_db. Kalau belum ada, dibuat otomatis.
 */
function getUserFacilityAssignmentSheet_() {
  const ss = getConfigTargetSpreadsheet_();
  let sheet = ss.getSheetByName(USER_FACILITY_ASSIGNMENT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USER_FACILITY_ASSIGNMENT_SHEET_NAME);
    sheet.appendRow(USER_FACILITY_ASSIGNMENT_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, USER_FACILITY_ASSIGNMENT_HEADERS);
  }
  return sheet;
}

function clearFacilityCache_() {
  try {
    CacheService.getScriptCache().remove(CACHE_FACILITY_LIST);
  } catch (e) { /* abaikan */ }
}

function clearUserFacilityCache_(username) {
  try {
    CacheService.getScriptCache().remove(CACHE_USER_FACILITY_PREFIX + String(username || '').toLowerCase());
  } catch (e) { /* abaikan */ }
}

// =====================================================================
// DAFTAR FACILITY
// =====================================================================

/**
 * Ambil semua daftar facility dari config_db. Di-cache 5 menit.
 */
function getDaftarFacility(username) {
  if (!isNikPunyaAksesSetting_(username)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Facility Management.' };
  }
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_FACILITY_LIST);
  if (cached) {
    return { success: true, facilities: JSON.parse(cached) };
  }
  const sheet = getFacilitySheet_();
  const lastRow = sheet.getLastRow();
  const facilities = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, FACILITY_HEADERS.length).getValues();
    values.forEach(function (r) {
      const id = String(r[0] || '').trim();
      if (!id) return;
      facilities.push({
        id: id,
        nama: String(r[1] || '').trim(),
        kode: String(r[2] || '').trim(),
        namaSpreadsheet: String(r[3] || '').trim(),
        spreadsheetId: String(r[4] || '').trim(),
        status: String(r[5] || '').trim(),
        tanggalDibuat: r[6] instanceof Date ? Utilities.formatDate(r[6], 'Asia/Jakarta', 'yyyy-MM-dd HH:mm') : String(r[6] || ''),
        dibuatOleh: String(r[7] || '').trim(),
        url: String(r[8] || '').trim()
      });
    });
  }
  try { cache.put(CACHE_FACILITY_LIST, JSON.stringify(facilities), 300); } catch (e) { /* abaikan */ }
  return { success: true, facilities: facilities };
}

// =====================================================================
// BUAT SPREADSHEET FACILITY BARU
// =====================================================================

/**
 * Buat spreadsheet baru untuk facility, lengkap dengan semua sheet operasional.
 * Return: { spreadsheetId, url }
 */
function buatSpreadsheetFacility_(namaSpreadsheet) {
  const ss = SpreadsheetApp.create(namaSpreadsheet);
  
  // Hapus sheet default "Sheet1" nanti di akhir
  const sheet1 = ss.getSheetByName('Sheet1');
  
  // Buat semua sheet operasional
  FACILITY_OPERATIONAL_SHEETS.forEach(function (def) {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
    }
    sheet.appendRow(def.headers);
    sheet.setFrozenRows(1);
  });
  
  // Hapus Sheet1 default kalau ada dan bukan satu-satunya
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }
  terapkanFormatKolomTeks_(ss);
  
  return {
    spreadsheetId: ss.getId(),
    url: ss.getUrl()
  };
}

// =====================================================================
// TAMBAH FACILITY BARU
// =====================================================================

/**
 * Tambah facility baru.
 * @param {string} username - NIK admin yang melakukan
 * @param {string} namaFacility - Nama facility (mis. "RDC Manado")
 * @param {string} kodeFacility - Kode facility (mis. "RDC-MDC")
 * @param {string} namaSpreadsheet - Nama spreadsheet yang akan dibuat
 * @param {Array<string>} daftarLokasi - Array string lokasi aktif untuk facility ini
 * @returns {Object} hasil operasi
 */
function tambahFacility(username, namaFacility, kodeFacility, namaSpreadsheet, daftarLokasi, jobId) {
  // v8.26.1: Menambah facility baru HANYA boleh role Developer -- beda dengan menu Config
  // Hub lain yang berbasis NIK Akses Setting (bisa termasuk Admin). Admin TIDAK boleh
  // membuat facility baru walau NIK-nya terdaftar di Config_Akses_Setting.
  const info = getUserRole(username);
  if (!info || info.role !== 'developer') {
    return { success: false, message: 'Hanya role Developer yang boleh menambah facility baru.' };
  }
  reportProgress_(jobId, 5, 'Validasi input');
  
  const nama = String(namaFacility || '').trim();
  const kode = String(kodeFacility || '').trim();
  const namaSs = String(namaSpreadsheet || '').trim();
  
  if (!nama) return { success: false, message: 'Nama Facility tidak boleh kosong.' };
  if (!kode) return { success: false, message: 'Kode Facility tidak boleh kosong.' };
  if (!namaSs) return { success: false, message: 'Nama Spreadsheet tidak boleh kosong.' };
  
  // Validasi daftar lokasi
  if (!daftarLokasi || !Array.isArray(daftarLokasi) || daftarLokasi.length === 0) {
    return { success: false, message: 'Daftar lokasi aktif wajib diisi. Minimal 1 lokasi.' };
  }
  
  // Cek duplikasi nama atau kode
  reportProgress_(jobId, 15, 'Cek duplikasi nama/kode facility');
  const sheet = getFacilitySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < existing.length; i++) {
      if (String(existing[i][1] || '').trim().toLowerCase() === nama.toLowerCase()) {
        return { success: false, message: 'Nama Facility "' + nama + '" sudah ada. Gunakan nama lain.' };
      }
      if (String(existing[i][2] || '').trim().toLowerCase() === kode.toLowerCase()) {
        return { success: false, message: 'Kode Facility "' + kode + '" sudah ada. Gunakan kode lain.' };
      }
    }
  }
  
  // Generate ID Facility
  const idFacility = 'FAC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  
  // Buat spreadsheet baru
  reportProgress_(jobId, 25, 'Membuat spreadsheet facility baru');
  const ssBaru = buatSpreadsheetFacility_(namaSs);
  
  // Tulis daftar lokasi aktif ke sheet "Lokasi_Aktif" di spreadsheet baru
  reportProgress_(jobId, 60, 'Menulis daftar lokasi aktif');
  const facilitySs = SpreadsheetApp.openById(ssBaru.spreadsheetId);
  const lokasiSheet = facilitySs.getSheetByName(LOKASI_AKTIF_SHEET_NAME);
  // Bersihkan lokasi duplikat & kosong (dideklarasikan di luar blok supaya terbaca saat menyusun pesan hasil)
  const lokasiUnik = {};
  const lokasiBersih = [];
  if (lokasiSheet && daftarLokasi.length > 0) {
    daftarLokasi.forEach(function (lok) {
      const l = String(lok || '').trim();
      if (l && !lokasiUnik[l]) {
        lokasiUnik[l] = true;
        lokasiBersih.push([l]);
      }
    });
    if (lokasiBersih.length > 0) {
      lokasiSheet.getRange(2, 1, lokasiBersih.length, 1).setValues(lokasiBersih);
    }
  }
  
  // Simpan ke sheet Facility di config_db
  reportProgress_(jobId, 85, 'Menyimpan info facility');
  const targetRow = Math.max(lastRow + 1, 2);
  const now = new Date();
  sheet.getRange(targetRow, 1).setValue(idFacility);
  sheet.getRange(targetRow, 2).setValue(nama);
  sheet.getRange(targetRow, 3).setValue(kode);
  sheet.getRange(targetRow, 4).setValue(namaSs);
  sheet.getRange(targetRow, 5).setValue(ssBaru.spreadsheetId);
  sheet.getRange(targetRow, 6).setValue('Aktif');
  sheet.getRange(targetRow, 7).setValue(now);
  sheet.getRange(targetRow, 8).setValue(username);
  sheet.getRange(targetRow, 9).setValue(ssBaru.url);
  
  clearFacilityCache_();
  catatLogPerubahanConfig_(username, 'Facility Management',
    'Tambah facility "' + nama + '" (Kode: ' + kode + ', ID: ' + idFacility + '). ' +
    'Spreadsheet dibuat: ' + ssBaru.url + '. ' +
    'Lokasi aktif diimpor: ' + (lokasiBersih ? lokasiBersih.length : daftarLokasi.length) + ' lokasi.');
  reportProgress_(jobId, 100, 'Selesai');
  
  return {
    success: true,
    message: 'Facility "' + nama + '" berhasil dibuat.',
    facility: {
      id: idFacility,
      nama: nama,
      kode: kode,
      spreadsheetId: ssBaru.spreadsheetId,
      url: ssBaru.url,
      jumlahLokasi: lokasiBersih ? lokasiBersih.length : daftarLokasi.length
    }
  };
}

/**
 * v8.29.0: Daftarkan facility baru dengan MEMAKAI spreadsheet yang SUDAH ADA (bukan bikin
 * spreadsheet kosong baru seperti tambahFacility()). Dipakai untuk facility yang datanya
 * sudah eksis dari sebelum sistem multi-facility ini dibuat (mis. RDC Tallo/Makassar --
 * spreadsheet terpisah yang sudah punya sheet "Data Count"/"Riwayat" berisi data lama).
 *
 * PENTING -- ini NON-DESTRUKTIF:
 * - Sheet yang SUDAH ADA di spreadsheet existing (Data Count, Riwayat, dst dari data lama)
 *   TIDAK PERNAH ditimpa/dikosongkan. Cuma ditambahkan kolom header yang kurang lewat
 *   ensureHeaderColumns_ (sama seperti getSheet_/getRiwayatSheet_ dkk di SheetHelpers.gs),
 *   supaya kompatibel dengan fitur2 yang lebih baru dari data lama itu (mis. kolom
 *   AddWho_Transaksi, Assigned_Validator, dst).
 * - Sheet operasional yang BELUM ada di spreadsheet lama (Log_Anomali, Ringkasan_Harian,
 *   Antrian_Aktif, Riwayat_Archive, Log_Sistem, Lokasi_Aktif, dst -- semua fitur yang lahir
 *   SETELAH data lama itu dibuat) otomatis ditambahkan sebagai sheet baru KOSONG.
 * - Lokasi_Aktif: kalau daftarLokasi diisi, lokasi BARU ditambahkan (append, dicek dulu biar
 *   tidak dobel) lewat importLokasiAktif() yang sudah ada -- lokasi yang mungkin sudah pernah
 *   diisi manual sebelumnya TETAP AMAN, tidak ditimpa.
 *
 * @param {string} username - NIK Developer yang melakukan (sama seperti tambahFacility)
 * @param {string} namaFacility - Nama facility (mis. "RDC Tallo Makassar")
 * @param {string} kodeFacility - Kode facility (mis. "RDC-MKS")
 * @param {string} spreadsheetIdAtauUrl - ID spreadsheet existing, ATAU link lengkap Google
 *   Sheets-nya (boleh dipaste apa adanya, ID-nya diekstrak otomatis)
 * @param {Array<string>} [daftarLokasi] - Opsional: daftar lokasi aktif tambahan. Boleh
 *   dikosongkan kalau mau isi lokasi belakangan lewat menu "Import Lokasi Aktif" terpisah.
 */
function daftarkanFacilityExisting(username, namaFacility, kodeFacility, spreadsheetIdAtauUrl, daftarLokasi, jobId) {
  // v8.29.0: sama seperti tambahFacility() -- HANYA role Developer yang boleh.
  const info = getUserRole(username);
  if (!info || info.role !== 'developer') {
    return { success: false, message: 'Hanya role Developer yang boleh mendaftarkan facility baru.' };
  }
  reportProgress_(jobId, 5, 'Validasi input');

  const nama = String(namaFacility || '').trim();
  const kode = String(kodeFacility || '').trim();
  const ssId = extractSpreadsheetId_(spreadsheetIdAtauUrl);

  if (!nama) return { success: false, message: 'Nama Facility tidak boleh kosong.' };
  if (!kode) return { success: false, message: 'Kode Facility tidak boleh kosong.' };
  if (!ssId) return { success: false, message: 'Spreadsheet ID/URL tidak boleh kosong atau tidak valid.' };

  // Cek duplikasi nama, kode, ATAU spreadsheet yang sudah kepakai facility lain (mencegah
  // 1 spreadsheet fisik ke-assign ke lebih dari 1 facility -- bisa bikin data campur aduk).
  reportProgress_(jobId, 15, 'Cek duplikasi facility');
  const sheet = getFacilitySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    for (let i = 0; i < existing.length; i++) {
      if (String(existing[i][1] || '').trim().toLowerCase() === nama.toLowerCase()) {
        return { success: false, message: 'Nama Facility "' + nama + '" sudah ada. Gunakan nama lain.' };
      }
      if (String(existing[i][2] || '').trim().toLowerCase() === kode.toLowerCase()) {
        return { success: false, message: 'Kode Facility "' + kode + '" sudah ada. Gunakan kode lain.' };
      }
      if (String(existing[i][4] || '').trim() === ssId) {
        return { success: false, message: 'Spreadsheet ini sudah terdaftar sebagai facility "' + existing[i][1] + '". Tidak boleh 1 spreadsheet dipakai 2 facility.' };
      }
    }
  }

  // Coba buka spreadsheet existing -- pastikan ID valid & script punya akses.
  reportProgress_(jobId, 30, 'Membuka spreadsheet existing');
  let facilitySs;
  try {
    facilitySs = SpreadsheetApp.openById(ssId);
  } catch (e) {
    return { success: false, message: 'Gagal membuka spreadsheet dengan ID/URL tersebut (' + e.message + '). Pastikan ID/URL-nya benar dan spreadsheet sudah di-share ke akun yang menjalankan script ini.' };
  }

  // Pastikan SEMUA sheet operasional yang dibutuhkan sistem ADA di spreadsheet ini -- yang
  // SUDAH ADA (data lama) sama sekali tidak disentuh/dikosongkan, cuma ditambah kolom header
  // yang kurang. Yang belum ada (fitur2 baru) dibuat baru dalam keadaan kosong.
  reportProgress_(jobId, 50, 'Menyiapkan sheet operasional (data lama tidak diubah)');
  FACILITY_OPERATIONAL_SHEETS.forEach(function (def) {
    let s = facilitySs.getSheetByName(def.name);
    if (!s) {
      s = facilitySs.insertSheet(def.name);
      s.appendRow(def.headers);
      s.setFrozenRows(1);
    } else {
      ensureHeaderColumns_(s, def.headers);
    }
  });
  terapkanFormatKolomTeks_(facilitySs);

  // Simpan ke sheet Facility di config_db dulu (sebelum import lokasi, karena
  // importLokasiAktif butuh facilityId sudah terdaftar).
  reportProgress_(jobId, 70, 'Menyimpan info facility');
  const idFacility = 'FAC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const targetRow = Math.max(lastRow + 1, 2);
  const now = new Date();
  sheet.getRange(targetRow, 1, 1, FACILITY_HEADERS.length).setValues([[
    idFacility, nama, kode, facilitySs.getName(), ssId, 'Aktif', now, username, facilitySs.getUrl()
  ]]);
  clearFacilityCache_();

  // Lokasi aktif tambahan (opsional) -- pakai importLokasiAktif() yang sudah ada supaya
  // logika merge/anti-dobelnya konsisten dengan menu "Import Lokasi Aktif" biasa.
  let jumlahLokasiDitambah = 0;
  if (daftarLokasi && Array.isArray(daftarLokasi) && daftarLokasi.length > 0) {
    reportProgress_(jobId, 85, 'Menambahkan lokasi aktif');
    const hasilLokasi = importLokasiAktif(username, idFacility, daftarLokasi, false);
    if (hasilLokasi && hasilLokasi.success) {
      jumlahLokasiDitambah = hasilLokasi.ditambahkan || 0;
    }
  }

  catatLogPerubahanConfig_(username, 'Facility Management',
    'Daftarkan facility "' + nama + '" (Kode: ' + kode + ', ID: ' + idFacility + ') dari SPREADSHEET EXISTING: ' +
    facilitySs.getUrl() + ' (data lama di sheet Data Count/Riwayat TIDAK diubah). Lokasi baru ditambahkan: ' + jumlahLokasiDitambah + '.');
  reportProgress_(jobId, 100, 'Selesai');

  return {
    success: true,
    message: 'Facility "' + nama + '" berhasil didaftarkan dari spreadsheet existing. Data lama di sheet Data Count/Riwayat tetap seperti semula.',
    facility: {
      id: idFacility,
      nama: nama,
      kode: kode,
      spreadsheetId: ssId,
      url: facilitySs.getUrl(),
      lokasiBaruDitambah: jumlahLokasiDitambah
    }
  };
}

// =====================================================================
// UBAH NAMA FACILITY
// =====================================================================

function updateNamaFacility(username, facilityId, namaBaru, kodeBaru) {
  if (!isNikPunyaAksesSetting_(username)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Facility Management.' };
  }
  const id = String(facilityId || '').trim();
  const nama = String(namaBaru || '').trim();
  const kode = String(kodeBaru || '').trim();
  if (!id) return { success: false, message: 'ID Facility tidak valid.' };
  if (!nama) return { success: false, message: 'Nama Facility tidak boleh kosong.' };
  
  const sheet = getFacilitySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'Tidak ada data facility.' };
  
  const values = sheet.getRange(2, 1, lastRow - 1, FACILITY_HEADERS.length).getValues();
  let targetIdx = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === id) {
      targetIdx = i;
      break;
    }
  }
  if (targetIdx === -1) return { success: false, message: 'Facility dengan ID tersebut tidak ditemukan.' };
  
  // Cek duplikasi nama/kode dengan facility lain
  for (let i = 0; i < values.length; i++) {
    if (i === targetIdx) continue;
    if (String(values[i][1] || '').trim().toLowerCase() === nama.toLowerCase()) {
      return { success: false, message: 'Nama Facility "' + nama + '" sudah dipakai facility lain.' };
    }
    if (kode && String(values[i][2] || '').trim().toLowerCase() === kode.toLowerCase()) {
      return { success: false, message: 'Kode Facility "' + kode + '" sudah dipakai facility lain.' };
    }
  }
  
  const rowIndex = targetIdx + 2;
  const namaLama = String(values[targetIdx][1] || '').trim();
  sheet.getRange(rowIndex, 2).setValue(nama);
  if (kode) sheet.getRange(rowIndex, 3).setValue(kode);
  
  clearFacilityCache_();
  catatLogPerubahanConfig_(username, 'Facility Management',
    'Ubah facility "' + namaLama + '" → "' + nama + '"' + (kode ? ' (Kode: ' + kode + ')' : '') + '.');
  
  return { success: true, message: 'Data facility berhasil diperbarui.' };
}

// =====================================================================
// AKTIFKAN / NONAKTIFKAN FACILITY
// =====================================================================

function setStatusFacility(username, facilityId, statusBaru) {
  if (!isNikPunyaAksesSetting_(username)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Facility Management.' };
  }
  const id = String(facilityId || '').trim();
  const status = (String(statusBaru || '').trim().toLowerCase() === 'nonaktif') ? 'Nonaktif' : 'Aktif';
  if (!id) return { success: false, message: 'ID Facility tidak valid.' };
  
  const sheet = getFacilitySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'Tidak ada data facility.' };
  
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  let targetIdx = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === id) {
      targetIdx = i;
      break;
    }
  }
  if (targetIdx === -1) return { success: false, message: 'Facility dengan ID tersebut tidak ditemukan.' };
  
  const namaFacility = String(values[targetIdx][1] || '').trim();
  sheet.getRange(targetIdx + 2, 6).setValue(status);
  
  clearFacilityCache_();
  catatLogPerubahanConfig_(username, 'Facility Management',
    'Status facility "' + namaFacility + '" diubah menjadi ' + status + '.');
  
  return {
    success: true,
    message: 'Facility "' + namaFacility + '" statusnya sekarang: ' + status + '.'
  };
}

// =====================================================================
// ASSIGN USER KE FACILITY
// =====================================================================

/**
 * Cari detail facility berdasarkan ID_Facility. Dipakai oleh getUserFacility() (assignment
 * normal) maupun oleh facility switcher Developer (override sementara).
 */
function getFacilityInfoById_(facilityId) {
  const id = String(facilityId || '').trim();
  if (!id) return null;
  const facSheet = getFacilitySheet_();
  const facLastRow = facSheet.getLastRow();
  if (facLastRow < 2) return null;
  const facValues = facSheet.getRange(2, 1, facLastRow - 1, FACILITY_HEADERS.length).getValues();
  for (let i = 0; i < facValues.length; i++) {
    if (String(facValues[i][0] || '').trim() === id) {
      return {
        id: id,
        nama: String(facValues[i][1] || '').trim(),
        kode: String(facValues[i][2] || '').trim(),
        spreadsheetId: String(facValues[i][4] || '').trim(),
        status: String(facValues[i][5] || '').trim(),
        url: String(facValues[i][8] || '').trim()
      };
    }
  }
  return null;
}

/**
 * Dapatkan info facility yang di-assign ke user tertentu.
 * Return null kalau belum di-assign.
 *
 * v8.26.2: Kalau user role Developer dan sedang punya "facility switcher" aktif (lihat
 * setDeveloperActiveFacility()), override ini SELALU dipakai duluan -- assignment asli
 * Developer (kalau ada) diabaikan sementara. Bukan Developer, tidak terpengaruh sama sekali.
 */
function getUserFacility(username) {
  const uname = String(username || '').trim().toLowerCase();
  if (!uname) return null;

  if (isUsernameRoleDeveloper_(uname)) {
    const overrideId = CacheService.getScriptCache().get(CACHE_DEV_ACTIVE_FACILITY_PREFIX + uname);
    if (overrideId) {
      const overrideInfo = getFacilityInfoById_(overrideId);
      if (overrideInfo) return overrideInfo;
      CacheService.getScriptCache().remove(CACHE_DEV_ACTIVE_FACILITY_PREFIX + uname);
    }
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = CACHE_USER_FACILITY_PREFIX + uname;
  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* abaikan */ }
  }
  
  let facilityId = null;
  
  // 1. Coba baca dari Master_User (sumber utama baru)
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    const muSheet = getMasterUserSheetWithFallback_();
    if (muSheet && muSheet.getLastRow() >= 2) {
      const values = muSheet.getRange(2, 1, muSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][0] || '').trim().toLowerCase() === uname) {
          facilityId = String(values[i][3] || '').trim();
          break;
        }
      }
    }
  }
  
  // 2. Fallback: baca dari User_Facility_Assignment
  if (!facilityId) {
    const sheet = getUserFacilityAssignmentSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let i = values.length - 1; i >= 0; i--) {
        if (String(values[i][0] || '').trim().toLowerCase() === uname) {
          facilityId = String(values[i][1] || '').trim();
          break;
        }
      }
    }
  }
  
  if (!facilityId) return null;
  
  const info = getFacilityInfoById_(facilityId);
  if (info) {
    try { cache.put(cacheKey, JSON.stringify(info), 300); } catch (e) { /* abaikan */ }
  }
  return info;
}

/**
 * Facility switcher untuk Developer: mengganti facility "aktif" Developer secara sementara
 * (cache, TTL 6 jam), TANPA mengubah data assignment permanen di User_Facility_Assignment.
 * HANYA role Developer yang boleh pakai. facilityId kosong ('' atau null) = kembali ke
 * assignment normal Developer (hapus override).
 */
function setDeveloperActiveFacility(username, facilityId) {
  const uname = String(username || '').trim().toLowerCase();
  if (!uname) return { success: false, message: 'Username tidak valid.' };
  if (!isUsernameRoleDeveloper_(uname)) {
    return { success: false, message: 'Hanya role Developer yang boleh pakai facility switcher.' };
  }

  const cacheKey = CACHE_DEV_ACTIVE_FACILITY_PREFIX + uname;
  const id = String(facilityId || '').trim();

  if (!id) {
    CacheService.getScriptCache().remove(cacheKey);
    return { success: true, message: 'Kembali ke facility assignment normal.', facility: getUserFacility(uname) };
  }

  const info = getFacilityInfoById_(id);
  if (!info) return { success: false, message: 'Facility tidak ditemukan.' };
  if (info.status === 'Nonaktif') {
    return { success: false, message: 'Facility "' + info.nama + '" sedang nonaktif, tidak bisa dipilih.' };
  }

  CacheService.getScriptCache().put(cacheKey, id, DEV_ACTIVE_FACILITY_TTL_SEC);
  return { success: true, message: 'Berpindah ke facility "' + info.nama + '".', facility: info };
}

/**
 * Assign user ke facility tertentu. Bisa dipanggil oleh admin saja.
 */
function assignUserKeFacility(username, targetUsername, facilityId) {
  if (!isNikPunyaAksesSetting_(username)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Facility Management.' };
  }
  const targetUname = String(targetUsername || '').trim();
  const facId = String(facilityId || '').trim();
  if (!targetUname) return { success: false, message: 'Username target tidak boleh kosong.' };
  if (!facId) return { success: false, message: 'ID Facility tidak boleh kosong.' };
  
  // Validasi user ada di master
  const userInfo = getUserRole(targetUname);
  if (!userInfo) {
    return { success: false, message: 'Username "' + targetUname + '" tidak terdaftar di Master.' };
  }
  
  // Validasi facility ada
  const facSheet = getFacilitySheet_();
  const facLastRow = facSheet.getLastRow();
  let namaFacility = '';
  let facStatus = '';
  if (facLastRow >= 2) {
    const facValues = facSheet.getRange(2, 1, facLastRow - 1, 6).getValues();
    let ketemu = false;
    for (let i = 0; i < facValues.length; i++) {
      if (String(facValues[i][0] || '').trim() === facId) {
        namaFacility = String(facValues[i][1] || '').trim();
        facStatus = String(facValues[i][5] || '').trim();
        ketemu = true;
        break;
      }
    }
    if (!ketemu) return { success: false, message: 'Facility dengan ID tersebut tidak ditemukan.' };
  } else {
    return { success: false, message: 'Belum ada facility yang terdaftar.' };
  }
  
  if (facStatus === 'Nonaktif') {
    return { success: false, message: 'Facility "' + namaFacility + '" sedang Nonaktif. Tidak bisa di-assign.' };
  }
  
  // Cek apakah user sudah punya assignment sebelumnya
  const sheet = getUserFacilityAssignmentSheet_();
  const lastRow = sheet.getLastRow();
  let existingRow = 0;
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim().toLowerCase() === targetUname.toLowerCase()) {
        existingRow = i + 2;
        break;
      }
    }
  }
  
  const now = new Date();
  if (existingRow > 0) {
    // Update assignment yang sudah ada
    sheet.getRange(existingRow, 2).setValue(facId);
    sheet.getRange(existingRow, 3).setValue(now);
    sheet.getRange(existingRow, 4).setValue(username);
  } else {
    // Tambah baris baru
    const targetRow = Math.max(lastRow + 1, 2);
    sheet.getRange(targetRow, 1).setValue(targetUname);
    sheet.getRange(targetRow, 2).setValue(facId);
    sheet.getRange(targetRow, 3).setValue(now);
    sheet.getRange(targetRow, 4).setValue(username);
  }
  
  // Dual write ke Master_User
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    const muSheet = getMasterUserSheetWithFallback_();
    if (muSheet) {
      const muLastRow = muSheet.getLastRow();
      if (muLastRow >= 2) {
        const muValues = muSheet.getRange(2, 1, muLastRow - 1, 1).getValues();
        for (let i = 0; i < muValues.length; i++) {
          if (String(muValues[i][0] || '').trim().toLowerCase() === targetUname.toLowerCase()) {
            muSheet.getRange(i + 2, 4).setValue(facId);
            break;
          }
        }
      }
    }
  }
  
  clearUserFacilityCache_(targetUname);
  catatLogPerubahanConfig_(username, 'Facility Management',
    'User "' + targetUname + '" di-assign ke facility "' + namaFacility + '".');
  
  return {
    success: true,
    message: 'User "' + targetUname + '" berhasil di-assign ke facility "' + namaFacility + '".'
  };
}

/**
 * Dapatkan daftar assignment user-facility untuk ditampilkan di UI.
 */
function getDaftarUserFacilityAssignment(username) {
  if (!isNikPunyaAksesSetting_(username)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Facility Management.' };
  }
  const sheet = getUserFacilityAssignmentSheet_();
  const lastRow = sheet.getLastRow();
  const assignments = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, USER_FACILITY_ASSIGNMENT_HEADERS.length).getValues();
    values.forEach(function (r, idx) {
      const uname = String(r[0] || '').trim();
      if (!uname) return;
      assignments.push({
        rowIndex: idx + 2,
        username: uname,
        facilityId: String(r[1] || '').trim(),
        tanggalDiassign: r[2] instanceof Date ? Utilities.formatDate(r[2], 'Asia/Jakarta', 'yyyy-MM-dd HH:mm') : String(r[2] || ''),
        diassignOleh: String(r[3] || '').trim()
      });
    });
  }
  return { success: true, assignments: assignments };
}

// =====================================================================
// FUNGSI SENTRAL: DAPATKAN SPREADSHEET OPERASIONAL BERDASARKAN USER
// =====================================================================

/**
 * Fungsi SENTRAL: mengembalikan Spreadsheet object untuk facility yang
 * di-assign ke user tersebut.
 *
 * FALLBACK: kalau user belum di-assign ke facility manapun, atau facility-nya
 * tidak bisa dibuka, kembalikan SpreadsheetApp.getActiveSpreadsheet() (kompatibilitas
 * dengan sistem lama sebelum ada Facility Management).
 *
 * Ini adalah fungsi yang akan dipanggil oleh SheetHelpers.gs untuk menggantikan
 * pemanggilan SpreadsheetApp.getActiveSpreadsheet() secara langsung.
 */
var _memSpreadsheetCache_ = {};

function getOperasionalSpreadsheet_(username) {
  const facInfo = getUserFacility(username);
  
  // Kalau user tidak punya assignment atau facility nonaktif, fallback ke active spreadsheet
  if (!facInfo || facInfo.status === 'Nonaktif') {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  
  const sid = facInfo.spreadsheetId;
  if (_memSpreadsheetCache_[sid]) {
    return _memSpreadsheetCache_[sid];
  }
  
  // Coba buka spreadsheet facility
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'facSsDown_' + facInfo.id;
    if (cache.get(cacheKey)) {
      // Baru saja gagal, langsung fallback
      return SpreadsheetApp.getActiveSpreadsheet();
    }
    const ss = SpreadsheetApp.openById(sid);
    _memSpreadsheetCache_[sid] = ss;
    return ss;
  } catch (e) {
    try { CacheService.getScriptCache().put('facSsDown_' + facInfo.id, '1', 60); } catch (e2) { /* abaikan */ }
    catatLogSistem_('Facility DB',
      'Gagal membuka spreadsheet facility "' + (facInfo.nama || facInfo.id) +
      '" (' + e.message + '). Fallback ke spreadsheet aktif.');
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

/**
 * v8.29.0: Wajib punya facility yang valid (assigned & aktif) sebelum boleh submit
 * count/validasi/task/upload. Sebelum versi ini, user tanpa assignment DIAM-DIAM fallback
 * ke spreadsheet Master DB sendiri (lewat getOperasionalSpreadsheet_ di atas) -- data
 * nyasar ke sheet Master DB alih-alih ke DC yang benar, dan biasanya baru ketahuan lewat
 * pesan membingungkan "Item tidak ditemukan" belakangan. Sekarang ditolak eksplisit di
 * awal fungsi, SEBELUM masuk antrean lock / nyentuh sheet manapun.
 *
 * @param {string} username
 * @return {{id:string, nama:string, kode:string, spreadsheetId:string, status:string,
 *   url:string}|null} - Info facility kalau valid & aktif, ATAU null kalau tidak. Caller
 *   HARUS langsung return pesan error ke user kalau hasilnya null, jangan lanjut proses.
 */
function requireUserFacility_(username) {
  const facInfo = getUserFacility(username);
  if (!facInfo || facInfo.status === 'Nonaktif') return null;
  return facInfo;
}

// =====================================================================
// LOCK PER-FACILITY (v8.29.0)
// =====================================================================

/**
 * Apps Script cuma nyediain LockService.getScriptLock() yang sifatnya GLOBAL untuk SATU
 * project -- semua facility, semua user, ikut antre di lock yang SAMA. Kalau 1 proses submit
 * agak lama (buka spreadsheet facility, beberapa kali flush(), dst), user DC LAIN yang
 * kebetulan submit bersamaan ikut ke-block >10 detik dan dapat "Sistem sedang sibuk",
 * padahal gak ada hubungannya sama facility mereka sendiri.
 *
 * Fungsi ini bikin lock "virtual" per-facility: penanda busy/tidaknya slot suatu facility
 * disimpan di CacheService (key beda per facilityId), sedangkan ScriptLock GLOBAL cuma
 * dipegang SEKILAS (maks 1 detik) buat cek-dan-set slot itu secara atomic -- BUKAN dipegang
 * selama proses submit yang sebenarnya berjalan. Efeknya: user DC A cuma ngantre kalau ada
 * user DC A LAIN yang lagi proses; user DC B/C/dst tetap jalan bebas bersamaan.
 *
 * @param {string} facilityId - ID facility (dari requireUserFacility_().id). Kalau kosong,
 *   dianggap 'default' -- tetap dapat slot sendiri, gak nyampur sama facility manapun.
 * @param {number} [timeoutMs=10000] - Total waktu tunggu maksimal sebelum caller dianggap
 *   gagal dapat giliran (tampilkan "Sistem sedang sibuk" ke user).
 * @return {{release: function()}|null} - Panggil release() setelah proses selesai (WAJIB,
 *   taruh di blok finally). null kalau gagal dapat slot dalam batas waktu.
 */
/**
 * v8.29.1: Multi-slot lock per facility.
 *
 * MASALAH SEBELUMNYA: 1 slot per facility → 50 user submit bersamaan = 49 user antri!
 * Setiap submit makan ~1-3 detik (buka SS, flush, tulis Riwayat), jadi user ke-50 bisa
 * nunggu sampai 50×3 = 150 detik = timeout pasti.
 *
 * SOLUSI: Setiap facility punya MAX_CONCURRENT_SLOTS slot. Setiap user mencari slot
 * KOSONG pertama (key: facLock_<facilityId>_s<0..N-1>). Kalau ada slot kosong, langsung
 * ambil dan lanjut proses tanpa nunggu. Hanya nunggu kalau SEMUA slot penuh (facility
 * benar-benar overload). Dengan 10 slot dan submit ~2 detik, facility bisa handle
 * 10/2 = ~5 submit/detik → 50 user submit dalam ~10 detik total, jauh lebih cepat.
 *
 * @param {string} facilityId
 * @param {number} [timeoutMs=10000]
 * @return {{release: function()}|null}
 */
function acquireFacilityLock_(facilityId, timeoutMs) {
  // v8.29.3 (performa submitCount @ 30-50 user bersamaan): MAX_CONCURRENT_SLOTS naik dari 10
  // ke 16. Alasan: testKecepatan50User() dipakai buat ukur rata-rata durasi 1 submit; dengan
  // rata-rata submit ~1-1.5 detik, 10 slot cuma sanggup ~6-10 submit/detik per facility, jadi
  // 40-50 user yang submit dalam window sempit (mis. shift dimulai bersamaan) tetap saling
  // antre & gampang nabrak timeout. 16 slot dipilih supaya masih realistis dibanding limit
  // eksekusi simultan Apps Script itu sendiri per project (di luar kendali kode ini -- lihat
  // catatan di README/CATATAN performa) sambil tetap ngasih headroom jauh lebih besar dari 10.
  const MAX_CONCURRENT_SLOTS = 16; // Slot paralel per facility
  const baseKey = 'facLock_' + (facilityId ? String(facilityId).trim() : 'default');
  const cache = CacheService.getScriptCache();
  const deadline = Date.now() + (timeoutMs || 10000);
  const scriptLock = LockService.getScriptLock();

  // v8.29.3: Siapkan daftar slot key SEKALI di luar loop (dipakai utk cache.getAll() batch).
  const allSlotKeys = [];
  for (let s = 0; s < MAX_CONCURRENT_SLOTS; s++) allSlotKeys.push(baseKey + '_s' + s);

  while (Date.now() < deadline) {
    try {
      // Lock global SEBENTAR saja — cuma buat cek+ambil slot secara atomic
      scriptLock.waitLock(1000);
    } catch (e) {
      // Lock global sedang dipakai proses lain — coba lagi
      continue;
    }
    try {
      // v8.29.3 (FIX bottleneck utama): SEBELUMNYA loop ini manggil cache.get(slotKey) SATU-SATU
      // sampai 16x di dalam genggaman scriptLock GLOBAL -- tiap cache.get() adalah 1 round-trip
      // jaringan ke CacheService, jadi kalau semua/sebagian besar slot penuh, lock global bisa
      // digenggam sampai 16x waktu round-trip sebelum nyerah, dan SEMUA user (lintas facility
      // sekalipun) ikut antre di belakangnya. cache.getAll() membaca SEMUA slot dalam SATU
      // round-trip, jadi waktu genggam lock global turun drastis (dari ~16 RPC jadi ~1 RPC +
      // 1 RPC put) -- ini yang paling berdampak buat throughput di beban 30-50 user bersamaan,
      // karena lock global ini dipakai BERSAMA oleh semua facility.
      const slotStates = cache.getAll(allSlotKeys); // { key: value } cuma utk key yang terisi
      let chosenKey = null;
      for (let s = 0; s < MAX_CONCURRENT_SLOTS; s++) {
        const slotKey = allSlotKeys[s];
        if (!slotStates[slotKey]) { chosenKey = slotKey; break; }
      }
      if (chosenKey) {
        // TTL 30 detik = jaring pengaman kalau release() tidak sempat dipanggil
        cache.put(chosenKey, '1', 30);
        const capturedSlotKey = chosenKey; // capture untuk closure release()
        return {
          release: function () {
            try { cache.remove(capturedSlotKey); } catch (e2) { /* abaikan */ }
          }
        };
      }
      // Semua slot penuh — lepas lock global dan tunggu sebentar
    } finally {
      scriptLock.releaseLock();
    }
    Utilities.sleep(100); // Lebih cepat dari 150ms sebelumnya
  }
  return null; // Timeout — semua slot penuh > batas waktu
}

// =====================================================================
// IMPORT LOKASI AKTIF
// =====================================================================

/**
 * Import lokasi aktif ke facility tertentu. Bisa untuk menambah atau mengganti
 * seluruh daftar lokasi.
 * @param {string} username - NIK admin
 * @param {string} facilityId - ID facility target
 * @param {Array<string>} daftarLokasi - Array string lokasi
 * @param {boolean} gantiSemua - true = hapus yang lama, ganti total; false = tambah (append)
 */
function importLokasiAktif(username, facilityId, daftarLokasi, gantiSemua) {
  if (!isNikPunyaAksesSetting_(username)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Facility Management.' };
  }
  const facId = String(facilityId || '').trim();
  if (!facId) return { success: false, message: 'ID Facility tidak boleh kosong.' };
  if (!daftarLokasi || !Array.isArray(daftarLokasi) || daftarLokasi.length === 0) {
    return { success: false, message: 'Daftar lokasi tidak boleh kosong.' };
  }
  
  // Cari detail facility
  const facSheet = getFacilitySheet_();
  const facLastRow = facSheet.getLastRow();
  let namaFacility = '';
  let spreadsheetId = '';
  if (facLastRow >= 2) {
    const facValues = facSheet.getRange(2, 1, facLastRow - 1, 5).getValues();
    let ketemu = false;
    for (let i = 0; i < facValues.length; i++) {
      if (String(facValues[i][0] || '').trim() === facId) {
        namaFacility = String(facValues[i][1] || '').trim();
        spreadsheetId = String(facValues[i][4] || '').trim();
        ketemu = true;
        break;
      }
    }
    if (!ketemu) return { success: false, message: 'Facility tidak ditemukan.' };
  } else {
    return { success: false, message: 'Belum ada facility yang terdaftar.' };
  }
  
  // Buka spreadsheet facility
  let facilitySs;
  try {
    facilitySs = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    return { success: false, message: 'Gagal membuka spreadsheet facility: ' + e.message };
  }
  
  let lokasiSheet = facilitySs.getSheetByName(LOKASI_AKTIF_SHEET_NAME);
  if (!lokasiSheet) {
    // Buat sheet kalau belum ada (kasus facility lama sebelum ada Lokasi_Aktif)
    lokasiSheet = facilitySs.insertSheet(LOKASI_AKTIF_SHEET_NAME);
    lokasiSheet.appendRow(LOKASI_AKTIF_HEADERS);
    lokasiSheet.setFrozenRows(1);
  }
  
  // Bersihkan & validasi lokasi
  const lokasiUnik = {};
  const lokasiBersih = [];
  daftarLokasi.forEach(function (lok) {
    const l = String(lok || '').trim();
    if (l && !lokasiUnik[l]) {
      lokasiUnik[l] = true;
      lokasiBersih.push([l]);
    }
  });
  
  if (lokasiBersih.length === 0) {
    return { success: false, message: 'Tidak ada lokasi valid setelah dibersihkan.' };
  }
  
  const lastRowLokasi = lokasiSheet.getLastRow();
  
  if (gantiSemua) {
    // Hapus data lama (baris 2 ke bawah), lalu tulis yang baru
    if (lastRowLokasi >= 2) {
      lokasiSheet.getRange(2, 1, lastRowLokasi - 1, 1).clearContent();
    }
    lokasiSheet.getRange(2, 1, lokasiBersih.length, 1).setValues(lokasiBersih);
  } else {
    // Append: tambahkan lokasi yang BELUM ADA
    const lokasiExisting = {};
    if (lastRowLokasi >= 2) {
      const existing = lokasiSheet.getRange(2, 1, lastRowLokasi - 1, 1).getValues();
      existing.forEach(function (r) {
        const l = String(r[0] || '').trim();
        if (l) lokasiExisting[l] = true;
      });
    }
    const lokasiBaru = lokasiBersih.filter(function (r) { return !lokasiExisting[r[0]]; });
    if (lokasiBaru.length === 0) {
      return { success: true, message: 'Semua lokasi sudah ada. Tidak ada yang ditambahkan.', ditambahkan: 0 };
    }
    const startRow = Math.max(lastRowLokasi + 1, 2);
    lokasiSheet.getRange(startRow, 1, lokasiBaru.length, 1).setValues(lokasiBaru);
    lokasiBersih.length = lokasiBaru.length;
  }
  
  catatLogPerubahanConfig_(username, 'Facility Management',
    'Import lokasi aktif ke facility "' + namaFacility + '": ' +
    (gantiSemua ? 'GANTI SEMUA' : 'TAMBAH') + ', ' + lokasiBersih.length + ' lokasi.');
  
  return {
    success: true,
    message: 'Berhasil mengimpor ' + lokasiBersih.length + ' lokasi ke facility "' + namaFacility + '".',
    ditambahkan: lokasiBersih.length
  };
}

/**
 * Dapatkan daftar lokasi aktif dari facility tertentu (untuk ditampilkan di UI).
 */
function getDaftarLokasiAktif(username, facilityId) {
  if (!isNikPunyaAksesSetting_(username)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Facility Management.' };
  }
  const facId = String(facilityId || '').trim();
  if (!facId) return { success: false, message: 'ID Facility tidak boleh kosong.' };
  
  // Cari spreadsheetId dari registry
  const facSheet = getFacilitySheet_();
  const facLastRow = facSheet.getLastRow();
  let spreadsheetId = '';
  if (facLastRow >= 2) {
    const facValues = facSheet.getRange(2, 1, facLastRow - 1, 5).getValues();
    for (let i = 0; i < facValues.length; i++) {
      if (String(facValues[i][0] || '').trim() === facId) {
        spreadsheetId = String(facValues[i][4] || '').trim();
        break;
      }
    }
  }
  if (!spreadsheetId) return { success: false, message: 'Facility tidak ditemukan.' };
  
  try {
    const facilitySs = SpreadsheetApp.openById(spreadsheetId);
    const lokasiSheet = facilitySs.getSheetByName(LOKASI_AKTIF_SHEET_NAME);
    if (!lokasiSheet) return { success: true, lokasi: [], total: 0 };
    
    const lastRow = lokasiSheet.getLastRow();
    const lokasi = [];
    if (lastRow >= 2) {
      const values = lokasiSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      values.forEach(function (r) {
        const l = String(r[0] || '').trim();
        if (l) lokasi.push(l);
      });
    }
    return { success: true, lokasi: lokasi, total: lokasi.length };
  } catch (e) {
    return { success: false, message: 'Gagal membaca lokasi: ' + e.message };
  }
}

// =====================================================================
// COPY LOKASI DARI FACILITY LAIN
// =====================================================================

/**
 * Copy daftar lokasi dari facility sumber ke facility target.
 * Opsi: ganti semua atau tambah.
 */
function copyLokasiDariFacility(username, facilityIdSumber, facilityIdTarget, gantiSemua) {
  if (!isNikPunyaAksesSetting_(username)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Facility Management.' };
  }
  
  // Ambil lokasi dari sumber
  const hasilSumber = getDaftarLokasiAktif(username, facilityIdSumber);
  if (!hasilSumber.success) {
    return { success: false, message: 'Gagal baca lokasi sumber: ' + hasilSumber.message };
  }
  if (!hasilSumber.lokasi || hasilSumber.lokasi.length === 0) {
    return { success: false, message: 'Facility sumber tidak memiliki lokasi aktif.' };
  }
  
  // Tulis ke target
  return importLokasiAktif(username, facilityIdTarget, hasilSumber.lokasi, gantiSemua);
}

// =====================================================================
// MIGRASI: SEED FACILITY AWAL DARI SISTEM LAMA
// =====================================================================

/**
 * Untuk sistem yang sudah berjalan (RDC Makassar existing):
 * Panggil fungsi ini SEKALI untuk mendaftarkan facility pertama (RDC Makassar)
 * menggunakan spreadsheet yang sedang aktif sekarang, dan menyalin lokasi aktif
 * dari master_db kolom A ke sheet "Lokasi_Aktif" di spreadsheet ini.
 *
 * Aman dipanggil berkali-kali (idempotent): kalau facility dengan kode RDC-MKS
 * sudah ada, tidak akan dibuat duplikat.
 */
function seedFacilityAwalDariSistemLama() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('FACILITY_SEEDED')) {
    Logger.log('Seed facility sudah pernah dijalankan. Lewati.');
    return;
  }
  
  const ssAktif = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getFacilitySheet_();
  const lastRow = sheet.getLastRow();
  
  // Cek apakah RDC Makassar sudah terdaftar
  let sudahAda = false;
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim().toUpperCase() === 'RDC-MKS') {
        sudahAda = true;
        break;
      }
    }
  }
  
  if (!sudahAda) {
    const idFacility = 'FAC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const targetRow = Math.max(lastRow + 1, 2);
    const now = new Date();
    
    sheet.getRange(targetRow, 1).setValue(idFacility);
    sheet.getRange(targetRow, 2).setValue('RDC Makassar');
    sheet.getRange(targetRow, 3).setValue('RDC-MKS');
    sheet.getRange(targetRow, 4).setValue(ssAktif.getName());
    sheet.getRange(targetRow, 5).setValue(ssAktif.getId());
    sheet.getRange(targetRow, 6).setValue('Aktif');
    sheet.getRange(targetRow, 7).setValue(now);
    sheet.getRange(targetRow, 8).setValue('system-migration');
    sheet.getRange(targetRow, 9).setValue(ssAktif.getUrl());
    
    Logger.log('Facility RDC Makassar terdaftar dengan ID: ' + idFacility);
  }
  
  // Pastikan sheet "Lokasi_Aktif" ada di spreadsheet aktif
  let lokasiSheet = ssAktif.getSheetByName(LOKASI_AKTIF_SHEET_NAME);
  if (!lokasiSheet) {
    lokasiSheet = ssAktif.insertSheet(LOKASI_AKTIF_SHEET_NAME);
    lokasiSheet.appendRow(LOKASI_AKTIF_HEADERS);
    lokasiSheet.setFrozenRows(1);
    
    // Salin lokasi dari master_db sheet Master kolom A
    const masterSheet = getMasterSheet_();
    if (masterSheet) {
      const mLastRow = masterSheet.getLastRow();
      if (mLastRow >= 2) {
        const lokasiValues = masterSheet.getRange(2, 1, mLastRow - 1, 1).getValues();
        const lokasiBersih = [];
        const lokasiUnik = {};
        lokasiValues.forEach(function (r) {
          const l = String(r[0] || '').trim();
          if (l && !lokasiUnik[l]) {
            lokasiUnik[l] = true;
            lokasiBersih.push([l]);
          }
        });
        if (lokasiBersih.length > 0) {
          lokasiSheet.getRange(2, 1, lokasiBersih.length, 1).setValues(lokasiBersih);
          Logger.log('Lokasi aktif disalin dari master_db: ' + lokasiBersih.length + ' lokasi.');
        }
      }
    }
  }
  
  props.setProperty('FACILITY_SEEDED', '1');
  clearFacilityCache_();
  Logger.log('Seed facility awal SELESAI.');
}
