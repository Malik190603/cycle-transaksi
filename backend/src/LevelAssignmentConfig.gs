/**
 * LevelAssignmentConfig.gs
 * v8.26.0 (TAHAP 1 dari rombak Pembagian Tugas) -- config matrix Role x Grup
 * Alat
 * (Bawah=L1-2, Tangga=L3-4, Reach Truck=L5-6) yang bisa diatur dari web app,
 * menggantikan
 * total fitur Setting Overflow (v8.24.0, khusus Storing->Tangga) dengan
 * sesuatu yang
 * general: role apapun bisa jadi backup role apapun, di grup alat manapun.
 *
 * v8.26.3 (PER-FACILITY): SEBELUMNYA matrix ini (+ mode_assignment +
 * maks_grup_alat_per_role)
 * disimpan di Config DB TERPUSAT (getConfigTargetSpreadsheet_()), artinya
 * SATU matrix
 * dipakai SEMUA facility. Ternyata tiap facility butuh pembagian tugas
 * SENDIRI-SENDIRI --
 * sekarang sheet "Config_Level_Assignment" pindah ke spreadsheet OPERASIONAL
 * tiap facility
 * (sama seperti Lokasi_Aktif), dan mode_assignment/maks_grup_alat_per_role
 * yang tadinya
 * baris di Config_Sistem (terpusat) sekarang ikut disimpan sebagai 2 baris
 * "meta" di sheet
 * Config_Level_Assignment yang sama (lihat LEVEL_ASSIGNMENT_META_*_KEY_ di
 * bawah) --
 * supaya tetap 1 sheet per facility, tidak perlu sheet baru.
 *
 * MIGRASI: facility yang SUDAH PERNAH diatur manual sebelum v8.26.3 (matrix/
 * mode di Config DB
 * lama) TIDAK otomatis ikut pindah -- begitu facility itu buka Pembagian
 * Tugas lagi, yang
 * muncul adalah DEFAULT (replikasi logic hardcode lama, lihat
 * LEVEL_ASSIGNMENT_DEFAULTS_,
 * node_assignment='legacy'). Kalau sebelumnya ada facility yang SUDAH pernah
 * pindah ke Mode
 * Matrix dengan matrix custom, itu perlu diatur ULANG manual per facility
 * setelah update ini.
 *
 * PENTING -- INI BARU LAPISAN CONFIG/PENYIMPANAN SAJA. ImportData.gs SUDAH
 * pakai
 * buildAssignedRowsMatrixAware_() kalau mode_assignment facility ybs =
 * 'matrix'.
 *
 * 3 grup alat itu FIXED (tidak bisa ditambah/dikurang dari UI) karena nempel
 * ke alat bantu
 * fisik yang memang cuma ada 3 jenis di gudang -- yang configurable adalah
 * ROLE mana yang
 * eligible di tiap grup, urutan prioritasnya, dan mode pembagiannya.
 */

const GRUP_ALAT_LIST_ = [
  { key: 'bawah', label: 'Bawah (Level 1-2)' },
  { key: 'tangga', label: 'Tangga Pesawat (Level 3-4)' },
  { key: 'reach_truck', label: 'Reach Truck (Level 5-6)' }
];

const LEVEL_ASSIGNMENT_SHEET_NAME = 'Config_Level_Assignment';
const LEVEL_ASSIGNMENT_HEADERS = ['GrupAlat', 'RolesCSV', 'Mode'];

// v8.26.3: 2 baris "meta" di sheet Config_Level_Assignment yang SAMA (kolom
// GrupAlat diisi
// key khusus ini, BUKAN salah satu dari GRUP_ALAT_LIST_) -- dipakai nyimpan
// mode_assignment
// & maks_grup_alat_per_role PER FACILITY. bacaLevelAssignmentMatrix_()
// otomatis mengabaikan
// baris ini (cuma proses key yang match GRUP_ALAT_LIST_), jadi aman 1 sheet
// dipakai bareng.
const LEVEL_ASSIGNMENT_META_MODE_KEY_ = '_meta_mode_assignment_';
const LEVEL_ASSIGNMENT_META_MAKSGRUP_KEY_ = '_meta_maks_grup_alat_';

// Default -- REPLIKASI PERSIS logic hardcode sekarang (Storing eksklusif
// Reach Truck,
// role lain eksklusif Bawah+Tangga) -- dipakai sebagai starting template
// begitu admin
// pertama kali buka layar Pembagian Tugas, SEBELUM dia ubah apa pun.
const LEVEL_ASSIGNMENT_DEFAULTS_ = {
  bawah: { roles: ['outbound', 'inbound', 'inventory', 'admin'], mode: 'pemerataan' },
  tangga: { roles: ['outbound', 'inbound', 'inventory', 'admin'], mode: 'pemerataan' },
  reach_truck: { roles: ['storing'], mode: 'pemerataan' }
};

const CONFIG_SISTEM_MODE_ASSIGNMENT_KEY_ = 'mode_assignment'; // 'legacy' | 'matrix' -- v8.26.3: tidak dipakai lagi (lihat LEVEL_ASSIGNMENT_META_MODE_KEY_), dipertahankan sebagai nama const untuk kompatibilitas referensi lama
const CONFIG_SISTEM_MAKS_GRUP_ALAT_KEY_ = 'maks_grup_alat_per_role'; // v8.26.3: tidak dipakai lagi (lihat LEVEL_ASSIGNMENT_META_MAKSGRUP_KEY_)

/**
 * v8.26.3: Menerima parameter opsional 'username' -- sheet
 * Config_Level_Assignment
 * SEKARANG di spreadsheet OPERASIONAL facility user tersebut (bukan Config
 * DB terpusat lagi).
 */
function getLevelAssignmentSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LEVEL_ASSIGNMENT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LEVEL_ASSIGNMENT_SHEET_NAME);
    sheet.appendRow(LEVEL_ASSIGNMENT_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, LEVEL_ASSIGNMENT_HEADERS);
  }
  if (sheet.getLastRow() < 2) {
    const rows = GRUP_ALAT_LIST_.map(function (g) {
      const d = LEVEL_ASSIGNMENT_DEFAULTS_[g.key];
      return [g.key, d.roles.join(','), d.mode];
    });
    sheet.getRange(2, 1, rows.length, LEVEL_ASSIGNMENT_HEADERS.length).setValues(rows);
  }
  return sheet;
}

/**
 * Cari baris meta (mode_assignment / maks_grup_alat) di sheet
 * Config_Level_Assignment.
 * Return { rowIndex, value } atau null kalau belum ada barisnya.
 */
function findLevelAssignmentMetaRow_(sheet, metaKey) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === metaKey) {
      return { rowIndex: i + 2, value: String(values[i][1] || '').trim() };
    }
  }
  return null;
}

function getLevelAssignmentMeta_(sheet) {
  const modeRow = findLevelAssignmentMetaRow_(sheet, LEVEL_ASSIGNMENT_META_MODE_KEY_);
  const maksRow = findLevelAssignmentMetaRow_(sheet, LEVEL_ASSIGNMENT_META_MAKSGRUP_KEY_);
  return {
    mode: (modeRow && modeRow.value === 'matrix') ? 'matrix' : 'legacy',
    maksGrupAlat: (maksRow && Number(maksRow.value) > 0) ? Number(maksRow.value) : 2
  };
}

function setLevelAssignmentMetaValue_(sheet, metaKey, value) {
  const existing = findLevelAssignmentMetaRow_(sheet, metaKey);
  if (existing) {
    sheet.getRange(existing.rowIndex, 2).setValue(String(value));
  } else {
    sheet.appendRow([metaKey, String(value), '']);
  }
}

/**
 * Baca matrix apa adanya dari sheet, dilengkapi default kalau ada grup alat
 * yang barisnya
 * belum ada (mis. sheet lama yang cuma sebagian keisi).
 *
 * v8.26.3: Menerima parameter opsional 'username' untuk baca dari facility
 * yang benar.
 */
function bacaLevelAssignmentMatrix_(username) {
  const sheet = getLevelAssignmentSheet_(username);
  const lastRow = sheet.getLastRow();
  const byKey = {};
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 3).getValues().forEach(function (r) {
      const key = String(r[0] || '').trim();
      // v8.26.3: lewati 2 baris meta (mode_assignment/maks_grup_alat) --
      // bukan grup alat.
      if (!key || key === LEVEL_ASSIGNMENT_META_MODE_KEY_ || key === LEVEL_ASSIGNMENT_META_MAKSGRUP_KEY_) return;
      byKey[key] = {
        roles: String(r[1] || '').split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(function (s) { return s; }),
        mode: (String(r[2] || '').trim().toLowerCase() === 'prioritas') ? 'prioritas' : 'pemerataan'
      };
    });
  }

  return GRUP_ALAT_LIST_.map(function (g) {
    const data = byKey[g.key] || LEVEL_ASSIGNMENT_DEFAULTS_[g.key];
    return { grupAlat: g.key, label: g.label, roles: data.roles, mode: data.mode };
  });
}

/**
 * Dipanggil dari layar Pembagian Tugas. Balikin matrix + mode assignment
 * aktif (legacy/
 * matrix) + batas maks grup alat per role, supaya form ke-render lengkap
 * sekali panggil.
 *
 * v8.26.3: Matrix, mode, & maksGrupAlat SEKARANG per-facility (baca dari
 * spreadsheet
 * operasional requesterUsername, bukan Config DB terpusat lagi).
 */
function getLevelAssignmentConfig(requesterUsername) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Pembagian Tugas.' };
  }

  const sheet = getLevelAssignmentSheet_(requesterUsername);
  const meta = getLevelAssignmentMeta_(sheet);
  return {
    success: true,
    matrix: bacaLevelAssignmentMatrix_(requesterUsername),
    modeAssignment: meta.mode,
    maksGrupAlat: meta.maksGrupAlat,
    roleValid: MASTER_ROLE_VALID_,
    grupAlatList: GRUP_ALAT_LIST_
  };
}

/**
 * Simpan matrix baru. TIDAK PERNAH menolak simpan (batas maks grup alat cuma
 * warning,
 * bukan validasi keras -- sesuai keputusan) -- tapi warning-nya dikembalikan
 * jelas di
 * response supaya UI bisa tampilin notif ke admin. Semua perubahan tercatat
 * ke
 * Log_Perubahan_Config, termasuk detail role yang berubah per grup alat.
 *
 * v8.26.3: Menyimpan ke sheet Config_Level_Assignment MILIK FACILITY
 * requesterUsername.
 */
function saveLevelAssignmentMatrix(requesterUsername, matrixBaru, maksGrupAlatBaru) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Pembagian Tugas.' };
  }

  if (!Array.isArray(matrixBaru) || matrixBaru.length !== GRUP_ALAT_LIST_.length) {
    return { success: false, message: 'Data matrix tidak lengkap (harus ada semua ' + GRUP_ALAT_LIST_.length + ' grup alat).' };
  }

  const matrixLama = bacaLevelAssignmentMatrix_(requesterUsername);
  const sheet = getLevelAssignmentSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const rowIndexByKey = {};
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function (r, idx) {
      rowIndexByKey[String(r[0] || '').trim()] = idx + 2;
    });
  }

  const perubahan = [];
  const hitungGrupPerRole = {}; // role -> jumlah grup alat yang dicentang, buat cek batas

  matrixBaru.forEach(function (g) {
    const key = String(g.grupAlat || '').trim();
    if (!key) return;
    const rolesBersih = (g.roles || []).map(function (r) { return String(r).trim().toLowerCase(); })
      .filter(function (r) { return MASTER_ROLE_VALID_.indexOf(r) !== -1; });
    const mode = (String(g.mode || '').trim().toLowerCase() === 'prioritas') ? 'prioritas' : 'pemerataan';

    rolesBersih.forEach(function (r) { hitungGrupPerRole[r] = (hitungGrupPerRole[r] || 0) + 1; });

    const rowIndex = rowIndexByKey[key];
    if (rowIndex) {
      sheet.getRange(rowIndex, 2).setValue(rolesBersih.join(','));
      sheet.getRange(rowIndex, 3).setValue(mode);
    } else {
      sheet.appendRow([key, rolesBersih.join(','), mode]);
    }

    const grupLama = matrixLama.filter(function (m) { return m.grupAlat === key; })[0];
    if (grupLama && (grupLama.roles.join(',') !== rolesBersih.join(',') || grupLama.mode !== mode)) {
      perubahan.push(key + ': [' + grupLama.roles.join(',') + ' / ' + grupLama.mode + '] -> [' + rolesBersih.join(',') + ' / ' + mode + ']');
    }
  });

  const maksGrupAlat = Math.max(1, Number(maksGrupAlatBaru) || 2);
  // v8.26.3: maks_grup_alat_per_role SEKARANG disimpan sebagai baris meta di
  // sheet
  // Config_Level_Assignment facility ini (bukan lagi setConfigSistemValue_
  // terpusat).
  setLevelAssignmentMetaValue_(sheet, LEVEL_ASSIGNMENT_META_MAKSGRUP_KEY_, maksGrupAlat);

  if (perubahan.length) {
    catatLogPerubahanConfig_(requesterUsername, 'Level Assignment', perubahan.join('; '));
  }

  const roleMelebihiBatas = Object.keys(hitungGrupPerRole).filter(function (r) { return hitungGrupPerRole[r] > maksGrupAlat; });
  const warning = roleMelebihiBatas.length
    ? 'Role ' + roleMelebihiBatas.join(', ') + ' dicentang di lebih dari ' + maksGrupAlat + ' grup alat -- tersimpan, tapi cek lagi apakah ini disengaja.'
    : '';

  return { success: true, message: 'Matrix Pembagian Tugas tersimpan.', warning: warning };
}

/**
 * Ganti mode assignment aktif ('legacy' = logic hardcode lama, 'matrix' =
 * pakai
 * Config_Level_Assignment). Tombol 'Kembali ke Mode Lama' di UI cuma memanggil
 * ini dengan
 * 'legacy' -- config matrix yang sudah diisi TIDAK dihapus, jadi aman dicoba
 * lagi nanti.
 *
 * v8.26.3: Mode assignment SEKARANG per-facility (baris meta di
 * Config_Level_Assignment
 * facility requesterUsername, BUKAN Config_Sistem terpusat lagi).
 */
function setModeAssignment(requesterUsername, mode) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Pembagian Tugas.' };
  }

  const modeBersih = (String(mode || '').trim().toLowerCase() === 'matrix') ? 'matrix' : 'legacy';
  const sheet = getLevelAssignmentSheet_(requesterUsername);
  setLevelAssignmentMetaValue_(sheet, LEVEL_ASSIGNMENT_META_MODE_KEY_, modeBersih);

  catatLogPerubahanConfig_(requesterUsername, 'Level Assignment', 'Mode assignment diubah jadi: ' + modeBersih);
  return {
    success: true,
    message: modeBersih === 'legacy'
      ? 'Kembali ke Mode Lama (By Alat) -- assignment harian pakai logic Storing/non-Storing seperti sebelumnya.'
      : 'Mode Matrix aktif -- assignment harian sekarang mengikuti Config_Level_Assignment facility ini.'
  };
}

/**
 * Riwayat perubahan config, ditampilkan di layar Config (biar audit trail
 * kelihatan tanpa
 * buka sheet manual). Dibatasi 50 terbarus supaya ringan (aplikasi lapangan,
 * performa
 * dijaga -- baca sheet secukupnya, bukan semua histori).
 */
function getLogPerubahanConfig(requesterUsername) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke log perubahan.' };
  }

  const sheet = getLogPerubahanConfigSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, log: [] };

  const ambilDari = Math.max(2, lastRow - 49);
  const jumlah = lastRow - ambilDari + 1;
  const values = sheet.getRange(ambilDari, 1, jumlah, 4).getValues();
  const log = values.map(function (r) {
    return { waktu: r[0] instanceof Date ? r[0].toISOString() : String(r[0] || ''), nik: r[1], sumber: r[2], detail: r[3] };
  }).reverse();
  return { success: true, log: log };
}

// ----------- TAHAP 2: assignment sungguhan berbasis matrix (dipanggil dari ImportData.gs,
// HANYA kalau mode_assignment = 'matrix') -----------

const GRUP_ALAT_TO_LEVEL_GROUP_KEY_ = { bawah: 'tanpa_alat', tangga: 'tangga', reach_truck: 'reach_truck' };

/**
 * Bangun pool user buat 1 grup alat sesuai Mode-nya:
 * - Pemerataan: semua role eligible digabung jadi 1 kolam, urutan asli
 * selectedUsernames
 *   dipertahankan (round-robin di buildAssignedRows_ jalan rata ke semua
 *   orang tanpa
 *   pandang role).
 * - Prioritas: role ditambahkan SATU PER SATU sesuai urutan di config --
 *   begitu kapasitas
 *   kumulatif (jumlah orang x produktivitas grup alat itu) sudah >= demand
 *   hari ini,
 *   berhenti nambah role berikutnya. Role belakangan cuma "backup" kalau
 *   role di depannya
 *   gak cukup nutup demand -- generalisasi dari fitur Overflow yang lama
 *   (Storing->Tangga).
 *   sekarang berlaku buat kombinasi role & grup alat manapun.
 */
function bangunPoolGrupAlat_(grupConfig, selectedUsernames, usernameRoleMap, demandCount, produktivitas) {
  if (grupConfig.mode === 'prioritas') {
    let pool = [];
    for (let i = 0; i < grupConfig.roles.length; i++) {
      const roleIni = grupConfig.roles[i];
      const roleUsers = selectedUsernames.filter(function (u) { return usernameRoleMap[u.toLowerCase()] === roleIni; });
      pool = pool.concat(roleUsers);
      const kapasitas = pool.length * (produktivitas > 0 ? produktivitas : 1);
      if (kapasitas >= demandCount) break;
    }
    return pool;
  }

  // Pemerataan -- gabung semua role eligible jadi 1 kolam, urutan asli dipertahankan.
  return selectedUsernames.filter(function (u) {
    return grupConfig.roles.indexOf(usernameRoleMap[u.toLowerCase()]) !== -1;
  });
}

/**
 * Versi Mode Matrix dari buildAssignedRowsEquipmentAware_ -- items
 * dikelompokkan ke 3 grup
 * alat berdasarkan lokasi (persis logic getLevelGroupByLocation_ yang sudah
 * ada), lalu tiap
 * grup dapat pool user sesuai Config_Level_Assignment (bukan hardcode
 * Storing/non-Storing).
 * TIDAK dipanggil sama sekali kecuali mode_assignment = 'matrix' -- lihat
 * writeItemsToSheet_.
 */
function buildAssignedRowsMatrixAware_(items, startingNo, tanggalUpload, selectedUsernames, usernameRoleMap, existingBacklogEffort, username) {
  const warnings = [];
  // v8.26.3: matrix SEKARANG per-facility -- teruskan username agar baca
  // dari facility yang benar.
  const matrix = bacaLevelAssignmentMatrix_(username);
  const prod = getLevelSettings_();

  const itemsByGrup = { bawah: [], tangga: [], reach_truck: [] };
  items.forEach(function (it) {
    const g = getLevelGroupByLocation_(it.lokasi); // v8.10.0: exception rollcage, tetap dipakai apa adanya
    const grupKey = (g.key === 'tanpa_alat') ? 'bawah' : g.key;
    itemsByGrup[grupKey].push(it);
  });

  let no = startingNo;
  let output = [];

  GRUP_ALAT_LIST_.forEach(function (g) {
    const grupItems = itemsByGrup[g.key];
    if (!grupItems.length) return;

    const grupConfig = matrix.filter(function (m) { return m.grupAlat === g.key; })[0];
    const produktivitas = prod[GRUP_ALAT_TO_LEVEL_GROUP_KEY_[g.key]] || 1;
    const pool = bangunPoolGrupAlat_(grupConfig, selectedUsernames, usernameRoleMap, grupItems.length, produktivitas);

    if (!pool.length) {
      warnings.push(grupItems.length + ' item ' + g.label + ' TIDAK dibagi karena tidak ada user dengan role yang eligible (cek Config > Pembagian Tugas).');
      return;
    }

    // Grup alat selain Bawah tetap pakai urutan section genap-dulu-baru-ganjil (sama
    // seperti Mode Legacy) -- ini soal urutan baca rak fisik, tidak berubah oleh matrix.
    const compareFn = (g.key === 'bawah') ? undefined : compareLocationEvenOddSection_;
    const rows = buildAssignedRows_(grupItems, no, tanggalUpload, pool, existingBacklogEffort, compareFn);
    output = output.concat(rows);
    no += rows.length;
  });

  return { rows: output, warnings: warnings, overflowPesan: '', overflowCount: 0 };
}
