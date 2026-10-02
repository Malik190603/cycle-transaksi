/**
 * OverflowConfig.gs
 * v8.24.0 (FASE OVERFLOW) -- Setting "Overflow": saat demand Level 5-6
 * (Reach Truck)
 * RENDAH dibanding kapasitas user Storing yang standby, kelebihan orang
 * dialihkan bantu
 * Level 3-4 (Tangga). Semua parameter (ON/OFF, batas minimal sisa, batas
 * maks orang)
 * disimpan di sheet key-value "Config_Sistem" supaya bisa diubah dari UI
 * App, TANPA edit
 * kode. Akses menu Setting (berbasis NIK) disimpan terpisah di
 * "Config_Akses_Setting".
 * (Hasil pemecahan dari Code.gs -- semua file .gs berbagi global scope yang
 * sama di Apps
 * Script, jadi urutan file tidak masalah, fungsi di file lain tetap bisa
 * saling panggil.)
 *
 * v8.26.4 (PER-FACILITY): Setting Overflow ini langsung mempengaruhi
 * PEMBAGIAN TUGAS harian
 * (redirect Storing ke Tangga), jadi SEHARUSNYA per-facility juga --
 * SEBELUMNYA sheet
 * "Config_Sistem" disimpan di Config DB TERPUSAT (1 setting untuk SEMUA
 * facility), SEKARANG
 * pindah ke spreadsheet OPERASIONAL tiap facility (sama pola seperti
 * Config_Level_Assignment
 * di v8.26.3). "Config_Akses_Setting" (siapa boleh buka Config Hub) TETAP
 * terpusat -- itu
 * soal hak akses admin, bukan data operasional per-facility.
 *
 * MIGRASI: sama seperti Config_Level_Assignment -- facility yang SUDAH
 * PERNAH mengaktifkan
 * Overflow manual sebelum v8.26.4 TIDAK otomatis ikut pindah settingnya.
 * Begitu facility itu
 * buka menu Setting Overflow lagi, yang muncul DEFAULT (overflow OFF) --
 * perlu diaktifkan
 * ulang manual per facility kalau sebelumnya memang sudah di-ON-kan.
 */

// ---------- Sheet Config_Sistem ----------

/**
 * v8.26.4: Menerima parameter opsional 'username' -- sheet Config_Sistem
 * SEKARANG di
 * spreadsheet OPERASIONAL facility user tersebut (bukan Config DB terpusat
 * lagi).
 */
function getConfigSistemSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SISTEM_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SISTEM_SHEET_NAME);
    sheet.appendRow(CONFIG_SISTEM_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, CONFIG_SISTEM_HEADERS);
  }
  // v8.24.0.1: seed default dicek dari LAST ROW, bukan dari exists-nya
  // sheet, supaya tetap
  // keisi walau percobaan sebelumnya sempat timeout di tengah jalan (header
  // sudah ke-appendRow
  // tapi baris default belum).
  if (sheet.getLastRow() < 2) {
    seedConfigSistemDefaults_(sheet);
  }
  return sheet;
}

function seedConfigSistemDefaults_(sheet) {
  const now = new Date();
  const rows = Object.keys(CONFIG_SISTEM_DEFAULTS_).map(function (key) {
    return [key, CONFIG_SISTEM_DEFAULTS_[key], CONFIG_SISTEM_KETERANGAN_[key] || '', now, 'system'];
  });
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, CONFIG_SISTEM_HEADERS.length).setValues(rows);
  }
}

/**
 * Baca semua parameter Config_Sistem jadi object { key: value }. Key yang
 * belum ada
 * barisnya di sheet (mis. sheet lama yang belum sempat diisi lengkap)
 * otomatis dilengkapi
 * dari CONFIG_SISTEM_DEFAULTS_ -- supaya kode pemanggil TIDAK PERLU cek
 * undefined satu-satu.
 * Di-cache 5 menit (cukup singkat karena ini parameter yang bisa diubah
 * admin kapan saja
 * lewat UI, beda dari Master yang cache-nya di-clear manual lewat onEdit).
 *
 * v8.26.4: Menerima parameter opsional 'username' untuk baca dari facility
 * yang benar.
 * Cache key diprefix facilityId supaya tidak tercampur antar facility (pola
 * sama seperti
 * Fix cache Antrian_Aktif/Ringkasan_Harian di v8.26.2).
 */
function getConfigSistem_(username) {
  const cache = CacheService.getScriptCache();
  const facPrefix = username ? (function () { const f = getUserFacility(username); return f ? f.id : 'legacy'; })() : 'legacy';
  const cacheKey = 'configSistem_' + facPrefix;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const sheet = getConfigSistemSheet_(username);
  const lastRow = sheet.getLastRow();
  const result = {};
  Object.keys(CONFIG_SISTEM_DEFAULTS_).forEach(function (k) { result[k] = CONFIG_SISTEM_DEFAULTS_[k]; });

  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    data.forEach(function (row) {
      const key = String(row[0] || '').trim();
      if (!key) return;
      result[key] = (row[1] === null || row[1] === undefined) ? '' : String(row[1]).trim();
    });
  }

  try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (e) { /* abaikan */ }
  return result;
}

/**
 * v8.26.4: Menerima parameter opsional 'username' untuk clear cache facility
 * yang benar.
 */
function clearConfigSistemCache_(username) {
  try {
    const facPrefix = username ? (function () { const f = getUserFacility(username); return f ? f.id : 'legacy'; })() : 'legacy';
    CacheService.getScriptCache().remove('configSistem_' + facPrefix);
  } catch (e) { /* abaikan */ }
}

/**
 * Tulis satu parameter Config_Sistem. Kalau key belum ada barisnya (sheet
 * lama), baris
 * baru ditambahkan di bawah -- TIDAK PERNAH menghapus/menggeser baris lain.
 *
 * v8.26.4: 'username' SEKARANG dipakai DUA hal -- (a) tulis ke sheet
 * Config_Sistem facility
 * user ini (bukan Config DB terpusat lagi), (b) dicatat di kolom Diubah_Oleh
 * (seperti sebelumnya).
 */
function setConfigSistemValue_(key, value, username) {
  const sheet = getConfigSistemSheet_(username);
  const lastRow = sheet.getLastRow();
  const now = new Date();
  let targetRow = 0;

  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0] || '').trim() === key) { targetRow = i + 2; break; }
    }
  }

  if (!targetRow) {
    targetRow = lastRow + 1;
    sheet.getRange(targetRow, 1).setValue(key);
    sheet.getRange(targetRow, 3).setValue(CONFIG_SISTEM_KETERANGAN_[key] || '');
  }

  sheet.getRange(targetRow, 2).setValue(value);
  sheet.getRange(targetRow, 4).setValue(now);
  sheet.getRange(targetRow, 5).setValue(username || '');
  clearConfigSistemCache_(username);
}

// ---------- Sheet Config_Akses_Setting (v8.26.4: TETAP terpusat -- soal hak akses admin,
// bukan data operasional, jadi TIDAK ikut pindah per-facility) ----------

function getConfigAksesSettingSheet_() {
  const ss = getConfigTargetSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG_AKSES_SETTING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_AKSES_SETTING_SHEET_NAME);
    sheet.appendRow(CONFIG_AKSES_SETTING_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, CONFIG_AKSES_SETTING_HEADERS);
  }
  // v8.24.0.1: seed NIK pertama dicek terpisah dari "sheet baru atau belum"
  // -- supaya
  // tetap terisi walau sheet-nya sempat kebuat duluan (mis. percobaan
  // sebelumnya timeout
  // di tengah jalan, header-nya sempat ke-appendRow tapi baris NIK belum
  // sempat). Dicek
  // dari LAST ROW (bukan exists-nya sheet), jadi aman dijalankan
  // berkali-kali (idempotent)
  // -- tidak akan dobel-nambah NIK yang sama kalau baris data sudah ada.
  if (sheet.getLastRow() < 2) {
    sheet.appendRow([SETUP_ADMIN_NIK_, '', Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd')]);
  }
  return sheet;
}

function getNikAksesSettingList_() {
  // v8.24.0.1: cache dihapus -- sheet ini kecil (segelintir NIK), risiko
  // "kelihatan belum
  // muncul padahal sheet sudah diisi" (stale cache sampai 5 menit) lebih
  // mengganggu
  // daripada manfaat performanya. Baca langsung tiap panggil.
  const sheet = getConfigAksesSettingSheet_();
  const lastRow = sheet.getLastRow();
  const list = [];
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    data.forEach(function (row) {
      const nik = String(row[0] || '').trim();
      if (nik) list.push(nik.toLowerCase());
    });
  }
  return list;
}

function clearNikAksesSettingCache_() {
  // Dipertahankan sebagai no-op supaya pemanggil lama (onEdit,
  // setupOverflowConfigSheets)
  // tidak perlu diubah -- cache-nya sendiri sudah tidak dipakai lagi (lihat
  // getNikAksesSettingList_).
}

function isNikPunyaAksesSetting_(username) {
  const nik = String(username || '').trim().toLowerCase();
  if (!nik) return false;
  // v8.26.1: Developer = super user, otomatis punya akses Config Hub (Kelola
  // User,
  // Pembagian Tugas, Akses Setting, Facility Management) walau NIK-nya belum
  // didaftarkan manual di sheet Config_Akses_Setting.
  // PENTING: pakai isUsernameRoleDeveloper_ (baca role mentah langsung
  // dari Master),
  // BUKAN getUserRole(), karena getUserRole() memanggil
  // isNikPunyaAksesSetting_() ini
  // untuk mengisi field aksesSettingOverflow -- kalau dipanggil balik akan
  // infinite loop.
  if (isUsernameRoleDeveloper_(nik)) return true;
  return getNikAksesSettingList_().indexOf(nik) !== -1;
}

/**
 * Cek role mentah user langsung dari sheet Master, TANPA lewat getUserRole
 * () --
 * isNikPunyaAksesSetting_() -- supaya aman dipakai di dalam
 * isNikPunyaAksesSetting_()
 * sendiri tanpa risiko infinite recursion.
 */
function isUsernameRoleDeveloper_(usernameLowercase) {
  const sheet = getMasterSheet_();
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const values = sheet.getRange(2, 3, lastRow - 1, 2).getValues(); // kolom C (username), D (role)
  for (let i = 0; i < values.length; i++) {
    const u = String(values[i][0] || '').trim().toLowerCase();
    if (u === usernameLowercase) {
      const roleRaw = String(values[i][1] || '').trim().toLowerCase();
      return roleRaw === 'developer' || roleRaw === 'dev' || roleRaw === 'dewa';
    }
  }
  return false;
}

/**
 * Daftar NIK yang punya akses menu Config (Setting Overflow + Kelola User),
 * ditampilkan
 * apa adanya (bukan lowercase) di UI Akses Setting.
 */
function getDaftarAksesSetting(requesterUsername) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke menu Akses Setting.' };
  }
  const sheet = getConfigAksesSettingSheet_();
  const lastRow = sheet.getLastRow();
  const daftar = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    values.forEach(function (r, idx) {
      const nik = String(r[0] || '').trim();
      if (!nik) return;
      daftar.push({ rowIndex: idx + 2, nik: nik, nama: String(r[1] || '').trim(), tanggal: String(r[2] || '').trim() });
    });
  }
  return { success: true, daftar: daftar };
}

function tambahAksesSetting(requesterUsername, nikBaru, namaBaru) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke menu Akses Setting.' };
  }
  const nik = String(nikBaru || '').trim();
  if (!nik) return { success: false, message: 'NIK tidak boleh kosong.' };

  const sheet = getConfigAksesSettingSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < existing.length; i++) {
      if (String(existing[i][0] || '').trim().toLowerCase() === nik.toLowerCase()) {
        return { success: false, message: 'NIK "' + nik + '" sudah terdaftar.' };
      }
    }
  }

  sheet.appendRow([nik, String(namaBaru || '').trim(), todayJakarta_()]);
  catatLogPerubahanConfig_(requesterUsername, 'Akses Setting', 'Tambah akses NIK "' + nik + '".');
  return { success: true, message: 'NIK "' + nik + '" berhasil ditambahkan ke akses Config.' };
}

/**
 * Cabut akses. Sengaja TIDAK bisa cabut akses NIK sendiri yang sedang login
 * -- supaya
 * tidak ada admin yang gak sengaja kunci diri sendiri keluar dari menu
 * Config kalau
 * kebetulan dia satu-satunya NIK terdaftar.
 */
function hapusAksesSetting(requesterUsername, rowIndex, nikKonfirmasi) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke menu Akses Setting.' };
  }
  const sheet = getConfigAksesSettingSheet_();
  const actualNik = String(sheet.getRange(rowIndex, 1).getValue() || '').trim();

  if (!actualNik || actualNik.toLowerCase() !== String(nikKonfirmasi || '').trim().toLowerCase()) {
    return { success: false, message: 'Data sudah berubah sejak halaman dibuka. Refresh dulu, lalu coba lagi.' };
  }

  if (actualNik.toLowerCase() === String(requesterUsername || '').trim().toLowerCase()) {
    return { success: false, message: 'Tidak bisa mencabut akses NIK Anda sendiri yang sedang login. Minta NIK lain yang punya akses untuk melakukan ini.' };
  }

  sheet.deleteRow(rowIndex);
  catatLogPerubahanConfig_(requesterUsername, 'Akses Setting', 'Cabut akses NIK "' + actualNik + '".');
  return { success: true, message: 'Akses NIK "' + actualNik + '" dicabut.' };
}

// ---------- Setup (dipanggil dari menu Cycle Count di Spreadsheet, lihat
// onOpen di Main.gs) ----------

function setupOverflowConfigSheets() {
  getConfigSistemSheet_();
  getConfigAksesSettingSheet_();
  clearConfigSistemCache_();
  clearNikAksesSettingCache_();
  const pesan = 'Sheet "Config_Sistem" & "Config_Akses_Setting" siap. NIK ' + SETUP_ADMIN_NIK_ + ' sudah didaftarkan sebagai akses pertama menu Setting Overflow.';
  // getUi() cuma valid kalau dipanggil dari klik menu di spreadsheet --
  // kalau dijalankan
  // langsung dari tombol Run di Apps Script editor, tidak ada context UI
  // sama sekali dan
  // getUi() akan melempar error. Dibungkus try/catch supaya fungsi tetap
  // SUKSES jalan di
  // kedua cara, cuma beda cara kasih tahu hasilnya (alert vs Logger).
  try {
    SpreadsheetApp.getUi().alert(pesan);
  } catch (e) {
    Logger.log(pesan);
  }
}

// ---------- API: dipanggil dari menu Setting Overflow di app ----------

function getSettingOverflow(requesterUsername) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Setting Overflow.' };
  }

  // v8.26.4: Teruskan requesterUsername agar baca dari Config_Sistem
  // facility user
  const cfg = getConfigSistem_(requesterUsername);
  return {
    success: true,
    overflowEnabled: String(cfg.overflow_enabled).toUpperCase() === 'TRUE',
    minSisaReachTruck: Number(cfg.overflow_min_sisa_reach_truck) || 0,
    maksOrang: cfg.overflow_maks_orang === '' || cfg.overflow_maks_orang === undefined ? null : Number(cfg.overflow_maks_orang) || 0,
    forceOffToday: cfg.overflow_force_off_today === todayJakarta_(),
    keterangan: CONFIG_SISTEM_KETERANGAN_
  };
}

function saveSettingOverflow(requesterUsername, overflowEnabled, minSisaReachTruck, maksOrang) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Setting Overflow.' };
  }

  setConfigSistemValue_('overflow_enabled', overflowEnabled ? 'TRUE' : 'FALSE', requesterUsername);
  setConfigSistemValue_('overflow_min_sisa_reach_truck', String(Math.max(0, Number(minSisaReachTruck) || 0)), requesterUsername);
  setConfigSistemValue_('overflow_maks_orang', (maksOrang === null || maksOrang === '' || maksOrang === undefined) ? '' : String(Math.max(0, Number(maksOrang) || 0)), requesterUsername);
  return { success: true, message: 'Setting Overflow tersimpan.' };
}

/**
 * Tombol "Nonaktifkan Overflow Hari Ini" -- rem darurat yang HANYA berlaku
 * untuk tanggal
 * hari ini (Asia/Jakarta), otomatis "lupa" sendiri besok paginya karena
 * selalu dibandingkan
 * ke todayJakarta_() saat dicek (lihat applyOverflowStoringKeTangga_). TIDAK
 * mengubah
 * overflow_enabled -- setting permanen di Setting screen tetap seperti apa
 * adanya.
 */
function setOverflowForceOffToday(requesterUsername) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Setting Overflow.' };
  }
  setConfigSistemValue_('overflow_force_off_today', todayJakarta_(), requesterUsername);
  return { success: true, message: 'Overflow dinonaktifkan untuk hari ini saja. Besok otomatis aktif lagi (kalau setting permanen ON).' };
}

function todayJakarta_() {
  return Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
}

// ---------- Logic inti: pindahkan kelebihan user Storing ke pool Tangga ----------

/**
 * Dipanggil dari buildAssignedRowsEquipmentAware_ (ImportData.gs) SEBELUM
 * Tahap 1 & 2
 * jalan, begitu jumlah item Level 5-6 (reachTruckItemsCount) sudah diketahui
 * dari data
 * upload hari ini. Reach Truck TETAP prioritas #1 -- overflow baru terjadi
 * kalau kapasitas
 * Reach Truck (jumlah user x produktivitas/orang dari Master kolom G) LEBIH
 * BESAR dari
 * demand hari ini, dan minimal overflow_min_sisa_reach_truck orang WAJIB
 * tetap di Reach
 * Truck supaya slot itu tidak pernah kosong sama sekali.
 * Return: { reachTruckUsers, tanggaUsers, overflowCount, pesan } -- pesan
 * kosong kalau
 * overflow tidak terjadi/tidak aktif (dipakai jadi notifikasi ke admin
 * setelah import).
 *
 * v8.26.4: Menerima parameter 'username' agar baca Setting Overflow dari
 * facility yang benar
 * (bukan Config DB terpusat lagi).
 */
function applyOverflowStoringKeTangga_(reachTruckItemsCount, reachTruckUsers, tanggaUsers, username) {
  const result = { reachTruckUsers: reachTruckUsers, tanggaUsers: tanggaUsers, overflowCount: 0, pesan: '' };

  if (!reachTruckUsers.length) return result;

  const cfg = getConfigSistem_(username);
  const enabled = String(cfg.overflow_enabled).toUpperCase() === 'TRUE';
  if (!enabled) return result;

  if (cfg.overflow_force_off_today === todayJakarta_()) {
    result.pesan = 'Overflow OFF hari ini (dinonaktifkan manual lewat tombol darurat).';
    return result;
  }

  const prodPerOrang = getEquipmentReadyDefaults_().reachTruck || 0;
  if (!prodPerOrang) return result; // produktivitas belum di-setting di Master kolom G -- tidak bisa hitung kapasitas dengan aman

  const kapasitas = reachTruckUsers.length * prodPerOrang;
  if (reachTruckItemsCount >= kapasitas) return result; // demand tinggi/pas, tidak ada kelebihan kapasitas

  const kelebihanOrang = Math.floor((kapasitas - reachTruckItemsCount) / prodPerOrang);
  if (kelebihanOrang <= 0) return result;

  const minSisa = Math.max(0, Number(cfg.overflow_min_sisa_reach_truck) || 0);
  const maksBolehPindah = reachTruckUsers.length - minSisa;
  if (maksBolehPindah <= 0) return result;

  const batasCfg = (cfg.overflow_maks_orang === '' || cfg.overflow_maks_orang === undefined)
    ? Infinity
    : (Number(cfg.overflow_maks_orang) || 0);

  const jumlahPindah = Math.min(kelebihanOrang, maksBolehPindah, batasCfg);
  if (jumlahPindah <= 0) return result;

  // Diambil dari antrian PALING BELAKANG (bukan yang pertama dicentang) --
  // konsisten
  // dengan pola "sisa antrian" yang sudah dipakai di logic Tangga/Bawah yang
  // lama.
  const dipindah = reachTruckUsers.slice(reachTruckUsers.length - jumlahPindah);
  const sisaReachTruck = reachTruckUsers.slice(0, reachTruckUsers.length - jumlahPindah);

  result.reachTruckUsers = sisaReachTruck;
  result.tanggaUsers = tanggaUsers.concat(dipindah);
  result.overflowCount = jumlahPindah;
  result.pesan = 'Overflow aktif: ' + jumlahPindah + ' user Storing dialihkan bantu Level 3-4 karena demand Level 5 & 6 rendah hari ini.';
  return result;
}
