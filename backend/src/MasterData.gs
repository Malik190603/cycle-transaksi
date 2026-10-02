/**
 * MasterData.gs
 * Baca data Master: lokasi aktif, user/role, assignment validator & PIC round-robin
 * (Hasil pemecahan dari Code.gs -- semua file .gs berbagi global scope yang sama di Apps Script,
 *  jadi urutan file tidak masalah, fungsi di file lain tetap bisa saling panggil.)
 *
 * v8.26.0 (Facility Management):
 * - getActiveLocations_() SEKARANG membaca dari sheet "Lokasi_Aktif" PER-FACILITY, bukan
 *   dari master_db kolom A (GLOBAL). Fallback ke master_db untuk backward compatibility.
 * - getUserRole() SEKARANG juga mengembalikan info facility user (jika sudah di-assign).
 * - v8.26.0 Strict Isolation: getAssignableUsers_(), getInventoryUsernames_(),
 *   getInvestigationOwnerPool_() SEKARANG menerima parameter opsional facilityId.
 *   Kalau disediakan, HANYA user yang di-assign ke facility tersebut yang dikembalikan.
 * - Role 'developer' TIDAK PERNAH masuk daftar assignable (tidak pernah diberi tugas).
 */

// v8.25.0: kolom E di Master = Status user ("Aktif"/"Nonaktif"). Baris lama yang kolom
// E-nya belum pernah diisi (blank) dianggap Aktif -- backward compatible, tidak perlu
// migrasi data lama satu-satu. Dipakai di semua fungsi baca user (getUserRole,
// getAssignableUsers_, dst) supaya user yang di-"Nonaktifkan" lewat menu Kelola User
// otomatis tidak bisa login lagi & tidak ikut ke-assign tugas baru, TANPA hapus
// baris/riwayatnya
// di-hapus (lihat UserManagement.gs).
function isBarisUserAktif_(statusRaw) {
  const s = String(statusRaw || '').trim().toLowerCase();
  return s !== 'nonaktif';
}

// ---------- Master: lokasi aktif & user/role ----------

/**
 * v8.26.0 (Facility Management):
 * Sebelumnya: lokasi aktif dibaca dari master_db sheet Master kolom A (GLOBAL untuk semua).
 * Sekarang: lokasi aktif PER-FACILITY, dibaca dari sheet "Lokasi_Aktif" di spreadsheet
 * operasional facility user tersebut.
 *
 * BACKWARD COMPATIBLE: kalau username tidak disediakan, atau sheet "Lokasi_Aktif" belum ada
 * di spreadsheet target, fallback membaca dari master_db sheet Master kolom A (sistem lama).
 *
 * @param {string} [username] - Opsional: username untuk menentukan facility spreadsheet
 */
function getActiveLocations_(username) {
  const cache = CacheService.getScriptCache();
  // Cache key per-user supaya setiap facility punya cache sendiri
  const cacheKey = username ? ('activeLocations_' + String(username).toLowerCase()) : 'activeLocations';
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const set = {};

  // Coba baca dari sheet "Lokasi_Aktif" di spreadsheet facility user
  if (username) {
    try {
      const lokasiSheet = getLokasiAktifSheet_(username);
      if (lokasiSheet) {
        const lastRow = lokasiSheet.getLastRow();
        if (lastRow >= 2) {
          const values = lokasiSheet.getRange(2, 1, lastRow - 1, 1).getValues();
          values.forEach(function (r) {
            const loc = String(r[0] || '').trim();
            if (loc) set[loc] = true;
          });
        }
      }
    } catch (e) {
      // Kalau gagal baca Lokasi_Aktif, fallback ke master_db di bawah
    }
  }

  // Fallback: kalau belum ada lokasi dari facility (jumlahnya 0), coba baca dari master_db
  // (kompatibilitas dengan sistem lama sebelum ada Facility Management)
  if (Object.keys(set).length === 0) {
    const sheet = getMasterSheet_();
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        values.forEach(function (r) {
          const loc = String(r[0] || '').trim();
          if (loc) set[loc] = true;
        });
      }
    }
  }

  try { cache.put(cacheKey, JSON.stringify(set), 300); } catch (e) { /* abaikan */ }
  return set;
}

var _memUserRoleCache_ = {};

function clearUserRoleCache_(username) {
  try {
    const uname = String(username || '').trim().toLowerCase();
    delete _memUserRoleCache_[uname];
    CacheService.getScriptCache().remove('userRole_' + uname);
  } catch (e) { /* abaikan */ }
}

/**
 * Mendapatkan role dan info user dari Master_User (fallback ke Master lama).
 * v8.26.0: Juga mengembalikan info facility user (jika sudah di-assign).
 * v8.29.0: Membaca dari sheet Master_User.
 */
function getUserRole(username) {
  const uname = String(username || '').trim().toLowerCase();
  if (!uname) return null;

  // Cek in-memory cache
  if (_memUserRoleCache_[uname]) {
    return _memUserRoleCache_[uname];
  }

  // Cek CacheService
  const cacheKey = 'userRole_' + uname;
  try {
    const cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) {
      const parsed = JSON.parse(cached);
      _memUserRoleCache_[uname] = parsed;
      return parsed;
    }
  } catch (e) { /* abaikan */ }

  // 1. Coba baca dari Master_User
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    const sheet = getMasterUserSheetWithFallback_();
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
        for (let i = 0; i < values.length; i++) {
          const u = String(values[i][0] || '').trim();
          if (u.toLowerCase() === uname) {
            if (!isBarisUserAktif_(values[i][2])) return null;

            const roleRaw = String(values[i][1] || '').trim().toLowerCase();
            let role = roleRaw;
            if (roleRaw === 'adm' || roleRaw === 'administrator') role = 'admin';
            else if (roleRaw === 'inventory') role = 'inventory';
            else if (roleRaw === 'cycle') role = 'cycle';
            else if (roleRaw === 'outbound') role = 'outbound';
            else if (roleRaw === 'storing') role = 'storing';
            else if (roleRaw === 'inbound') role = 'inbound';
            else if (roleRaw === 'developer' || roleRaw === 'dev' || roleRaw === 'dewa') role = 'developer';

            let facInfo = null;
            const facId = String(values[i][3] || '').trim();
            if (facId && typeof getFacilityInfoById_ === 'function') {
              facInfo = getFacilityInfoById_(facId);
            }

            // Fallback facility (jika kosong, mungkin user dari migrasi belum diassign)
            if (!facInfo && typeof getUserFacility === 'function') {
              // Perlu hati-hati infinite loop, jangan pakai getUserFacility langsung jika memanggil ini
            }

            const result = {
              role: role,
              displayName: u,
              aksesSettingOverflow: isNikPunyaAksesSetting_(u)
            };

            if (facInfo) {
              result.facilityId = facInfo.id;
              result.facilityName = facInfo.nama;
              result.facilityCode = facInfo.kode;
              result.facilityStatus = facInfo.status;
              result.facilityUrl = facInfo.url;
            }

            _memUserRoleCache_[uname] = result;
            try {
              CacheService.getScriptCache().put('userRole_' + uname, JSON.stringify(result), 300);
            } catch (e) { /* abaikan */ }
            return result;
          }
        }
      }
    }
  }

  // 2. Fallback: baca dari Master lama
  const sheetOld = getMasterSheet_();
  if (!sheetOld) return null;
  const lastRow = sheetOld.getLastRow();
  if (lastRow < 2) return null;

  const values = sheetOld.getRange(2, 3, lastRow - 1, 3).getValues();
  for (let i = 0; i < values.length; i++) {
    const u = String(values[i][0] || '').trim();
    if (u.toLowerCase() === uname) {
      if (!isBarisUserAktif_(values[i][2])) return null;
      const roleRaw = String(values[i][1] || '').trim().toLowerCase();
      let role = roleRaw;
      if (roleRaw === 'adm' || roleRaw === 'administrator') role = 'admin';
      else if (roleRaw === 'inventory') role = 'inventory';
      else if (roleRaw === 'cycle') role = 'cycle';
      else if (roleRaw === 'outbound') role = 'outbound';
      else if (roleRaw === 'storing') role = 'storing';
      else if (roleRaw === 'inbound') role = 'inbound';
      else if (roleRaw === 'developer' || roleRaw === 'dev' || roleRaw === 'dewa') role = 'developer';

      let facInfo = null;
      // Gunakan manual assignment lookup jika belum terintegrasi di sheet
      if (typeof getUserFacilityAssignmentSheet_ === 'function') {
        const assignSheet = getUserFacilityAssignmentSheet_();
        if (assignSheet && assignSheet.getLastRow() >= 2) {
          const assignData = assignSheet.getRange(2, 1, assignSheet.getLastRow() - 1, 2).getValues();
          for (let j = assignData.length - 1; j >= 0; j--) {
            if (String(assignData[j][0]).trim().toLowerCase() === u.toLowerCase()) {
              const fid = String(assignData[j][1]).trim();
              if (fid && typeof getFacilityInfoById_ === 'function') {
                facInfo = getFacilityInfoById_(fid);
              }
              break;
            }
          }
        }
      }

      const result = {
        role: role,
        displayName: u,
        aksesSettingOverflow: isNikPunyaAksesSetting_(u)
      };

      if (facInfo) {
        result.facilityId = facInfo.id;
        result.facilityName = facInfo.nama;
        result.facilityCode = facInfo.kode;
        result.facilityStatus = facInfo.status;
        result.facilityUrl = facInfo.url;
      }

      _memUserRoleCache_[uname] = result;
      try {
        CacheService.getScriptCache().put('userRole_' + uname, JSON.stringify(result), 300);
      } catch (e) { /* abaikan */ }
      return result;
    }
  }

  return null;
}

function getMasterUsernames() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('masterUsernames');
  if (cached) return JSON.parse(cached);

  let list = [];
  let useLegacy = true;

  // Coba Master_User dulu
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    const sheet = getMasterUserSheetWithFallback_();
    if (sheet && sheet.getLastRow() >= 2) {
      useLegacy = false;
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues(); // A=User, B=Role, C=Status
      list = values
        .filter(function (r) { return isBarisUserAktif_(r[2]); })
        .map(function (r) { return String(r[0] || '').trim(); })
        .filter(function (v) { return v; });
    }
  }

  // Fallback
  if (useLegacy) {
    const sheet = getMasterSheet_();
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const values = sheet.getRange(2, 3, lastRow - 1, 3).getValues(); // C=User, D=Role, E=Status
        list = values
          .filter(function (r) { return isBarisUserAktif_(r[2]); })
          .map(function (r) { return String(r[0] || '').trim(); })
          .filter(function (v) { return v; });
      }
    }
  }

  cache.put('masterUsernames', JSON.stringify(list), 300);
  return list;
}

/**
 * v8.26.0 Strict Isolation:
 * Dapatkan daftar user yang BISA diberi tugas (assignable).
 * - Kalau facilityId disediakan: HANYA user yang di-assign ke facility tersebut.
 * - Role 'developer' TIDAK PERNAH masuk daftar ini.
 *
 * @param {string} [facilityId] - Opsional: filter user per facility
 */
function getAssignableUsers_(facilityId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = facilityId ? ('assignableUsers_fac_' + facilityId) : 'assignableUsers';
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  let list = [];
  let useLegacy = true;

  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    const sheet = getMasterUserSheetWithFallback_();
    if (sheet && sheet.getLastRow() >= 2) {
      useLegacy = false;
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
      // User, Role, Status, ID_Facility
      list = values.map(function (r) {
        if (!isBarisUserAktif_(r[2])) return null;
        const username = String(r[0] || '').trim();
        const roleRaw = String(r[1] || '').trim().toLowerCase();
        const facId = String(r[3] || '').trim();

        if (roleRaw === 'developer' || roleRaw === 'dev' || roleRaw === 'dewa') return null;
        if (facilityId && facId !== facilityId) return null; // Pakai kolom D langsung

        let role = roleRaw;
        if (roleRaw === 'adm' || roleRaw === 'administrator') role = 'admin';

        return username ? { username: username, role: role } : null;
      }).filter(function (v) { return v; });
    }
  }

  // Fallback (legacy)
  if (useLegacy) {
    const sheet = getMasterSheet_();
    let userDiFacility = null;
    if (facilityId) {
      try {
        const assignSheet = getUserFacilityAssignmentSheet_();
        const lastRowAssign = assignSheet.getLastRow();
        if (lastRowAssign >= 2) {
          const assignValues = assignSheet.getRange(2, 1, lastRowAssign - 1, 2).getValues();
          userDiFacility = {};
          assignValues.forEach(function (r) {
            const uname = String(r[0] || '').trim().toLowerCase();
            const facId = String(r[1] || '').trim();
            if (uname && facId === facilityId) {
              userDiFacility[uname] = true;
            }
          });
        }
      } catch (e) {
        userDiFacility = null;
      }
    }

    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const values = sheet.getRange(2, 3, lastRow - 1, 3).getValues(); // C=User, D=Role, E=Status
        list = values.map(function (r) {
          if (!isBarisUserAktif_(r[2])) return null;
          const username = String(r[0] || '').trim();
          const roleRaw = String(r[1] || '').trim().toLowerCase();

          if (roleRaw === 'developer' || roleRaw === 'dev' || roleRaw === 'dewa') return null;
          if (userDiFacility && !userDiFacility[username.toLowerCase()]) return null;

          let role = roleRaw;
          if (roleRaw === 'adm' || roleRaw === 'administrator') role = 'admin';

          return username ? { username: username, role: role } : null;
        }).filter(function (v) { return v; });
      }
    }
  }

  try { cache.put(cacheKey, JSON.stringify(list), 300); } catch (e) { /* abaikan */ }
  return list;
}

function getAssignableUsers(requesterUsername) {
  requireRole_(requesterUsername, ['admin', 'developer']);
  // v8.26.0: Dapatkan facility requester, lalu filter assignable users per facility
  const facInfo = getUserFacility(requesterUsername);
  const facilityId = facInfo ? facInfo.id : null;
  return getAssignableUsers_(facilityId);
}

/**
 * Lookup cepat username(lowercase) -> role(lowercase, dinormalisasi) untuk semua user
 * Master. Dipakai ImportData.gs buat nentuin siapa yang role Storing (otomatis eligible
 * Reach Truck, TIDAK dibatasi jumlah -- lihat catatan di importRawData).
 *
 * v8.26.0: Role 'developer' DILUARKAN dari map ini (tidak pernah ikut assignment).
 */
function getUserRoleMap_() {
  // v8.26.0: Gunakan getAssignableUsers_() yang sudah otomatis exclude developer
  const list = getAssignableUsers_();
  const map = {};
  list.forEach(function (u) {
    let role = String(u.role || '').trim().toLowerCase();
    if (role === 'adm' || role === 'administrator') role = 'admin';
    map[u.username.trim().toLowerCase()] = role;
  });
  return map;
}

// ---------- Pembagian Validator otomatis (round-robin antar role Inventory) ----------

/**
 * v8.26.0 Strict Isolation:
 * Dapatkan daftar user role Inventory yang BISA di-assign jadi validator.
 * - Kalau facilityId disediakan: HANYA user Inventory yang di-assign ke facility tersebut.
 * - Role 'developer' otomatis tidak masuk (karena getAssignableUsers_ sudah mengecualikan).
 *
 * @param {string} [facilityId] - Opsional: filter per facility
 */
function getInventoryUsernames_(facilityId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = facilityId ? ('inventoryUsers_fac_' + facilityId) : 'inventoryUsers';
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  // v8.26.0: Gunakan getAssignableUsers_ yang sudah difilter facility & exclude developer
  const allAssignable = getAssignableUsers_(facilityId);
  const list = allAssignable
    .filter(function (u) { return String(u.role || '').trim().toLowerCase() === 'inventory'; })
    .map(function (u) { return u.username; });

  try { cache.put(cacheKey, JSON.stringify(list), 300); } catch (e) { /* abaikan */ }
  return list;
}

/**
 * Menunjuk 1 validator secara round-robin (bergilir merata) dari daftar user role Inventory.
 * v8.26.0: Menerima parameter opsional facilityId untuk strict isolation per facility.
 * Pointer round-robin DISIMPAN PER-FACILITY supaya giliran antar facility tidak saling ganggu.
 *
 * Kalau facility tidak punya user Inventory sama sekali, kembalikan string kosong --
 * item tetap masuk antrian Pending tapi tidak ke siapa-siapa; Admin tetap bisa melihat &
 * validasi (fallback).
 *
 * @param {string} [facilityId] - Opsional: filter validator per facility
 */
function assignNextValidator_(facilityId) {
  const users = getInventoryUsernames_(facilityId);
  if (!users.length) return '';

  // v8.29.2 (performa submitCount): pindah dari PropertiesService ke CacheService.
  // PropertiesService.getProperty/setProperty di spreadsheet facility eksternal bisa makan
  // 200-500ms+ per panggilan (2x dipanggil di sini = read+write) -- salah satu penyumbang
  // lambatnya submitCount() sejak facility split. Pointer round-robin ini cuma butuh "cukup
  // adil" antar validator, TIDAK perlu persisten permanen, jadi aman pindah ke CacheService
  // (jauh lebih cepat). TTL 21600 detik (6 jam, maksimum CacheService) -- kalau cache kosong
  // (baru deploy / lama tidak dipakai facility-nya), otomatis mulai dari 0 lagi; efeknya cuma
  // reset urutan giliran, TIDAK merusak/menghapus data apapun.
  const cache = CacheService.getScriptCache();
  const cacheKey = facilityId
    ? ('validasiRoundRobinPointer_' + facilityId)
    : 'validasiRoundRobinPointer';

  let pointer = Number(cache.get(cacheKey) || '0');
  if (isNaN(pointer) || pointer < 0) pointer = 0;
  const username = users[pointer % users.length];
  try { cache.put(cacheKey, String(pointer + 1), 21600); } catch (e) { /* abaikan */ }
  return username;
}

/**
 * Pool "pemilik awal" task investigasi (role Inventory + Admin) -- dipakai untuk
 * assignInitialPIC_ saat task baru Open.
 *
 * v8.26.0 Strict Isolation:
 * - Kalau facilityId disediakan: HANYA user dari facility tersebut.
 * - Role 'developer' otomatis tidak masuk (getAssignableUsers_ sudah mengecualikan).
 *
 * @param {string} [facilityId] - Opsional: filter per facility
 */
function getInvestigationOwnerPool_(facilityId) {
  // v8.26.0: Gunakan getAssignableUsers_ yang sudah difilter facility & exclude developer
  const allAssignable = getAssignableUsers_(facilityId);
  const list = allAssignable
    .filter(function (u) {
      const role = String(u.role || '').trim().toLowerCase();
      return role === 'inventory' || role === 'admin';
    })
    .map(function (u) { return u.username; });
  return list;
}

/**
 * v8.26.0: Menerima parameter opsional facilityId untuk strict isolation.
 * Pointer round-robin DISIMPAN PER-FACILITY.
 *
 * @param {string} [facilityId] - Opsional: filter PIC per facility
 */
function assignInitialPIC_(facilityId) {
  const users = getInvestigationOwnerPool_(facilityId);
  if (!users.length) return '';

  const props = PropertiesService.getScriptProperties();
  // v8.26.0: Pointer terpisah per facility
  const propKey = facilityId
    ? ('picRoundRobinPointer_' + facilityId)
    : 'picRoundRobinPointer';

  let pointer = Number(props.getProperty(propKey) || '0');
  if (isNaN(pointer) || pointer < 0) pointer = 0;
  const username = users[pointer % users.length];
  props.setProperty(propKey, String(pointer + 1));
  return username;
}

function getKategoriMapping_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('kategoriMapping');
  if (cached) return JSON.parse(cached);
  const sheet = getMasterSheet_();
  let list = [];
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 9, lastRow - 1, 3).getValues(); // I=Kategori, J=User, K=Role
      values.forEach(function (r) {
        const kategori = String(r[0] || '').trim();
        const username = String(r[1] || '').trim();
        const role = String(r[2] || '').trim();
        if (kategori && username) list.push({ kategori: kategori, username: username, role: role });
      });
    }
  }
  cache.put('kategoriMapping', JSON.stringify(list), 300);
  return list;
}

/**
 * Data lengkap untuk form Task Investigasi (dipanggil sekali saat layar dibuka):
 * daftar kategori wajib, mapping PIC per kategori, dan seluruh user Master
 * (fallback kalau suatu kategori belum ada PIC-nya di mapping).
 */
function getInvestigasiFormData(requesterUsername) {
  requireRole_(requesterUsername, ['inventory', 'admin', 'developer']);
  const sheet = getMasterSheet_();
  let allUsers = [];

  // v8.26.0: Dapatkan facility requester untuk filter user
  const facInfo = getUserFacility(requesterUsername);
  const facilityId = facInfo ? facInfo.id : null;

  // v8.26.0: Gunakan getAssignableUsers_ yang sudah difilter facility & exclude developer
  allUsers = getAssignableUsers_(facilityId);

  return {
    kategoriList: TASK_KATEGORI_LIST,
    mapping: getKategoriMapping_(),
    allUsers: allUsers,
    statusFlow: TASK_STATUS_FLOW,
    slaHari: TASK_SLA_HARI
  };
}

function getLevelSettings_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('levelSettings');
  if (cached) return JSON.parse(cached);
  const prod = {};
  LEVEL_GROUPS.forEach(function (g) { prod[g.key] = g.defaultProd; });
  const sheet = getMasterSheet_();
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 6, lastRow - 1, 2).getValues();
      data.forEach(function (row) {
        const label = String(row[0] || '').trim();
        const value = row[1];
        if (!label || value === '' || isNaN(value)) return;
        const g = LEVEL_GROUPS.filter(function (gr) { return label.indexOf(gr.min + '-' + gr.max) !== -1; })[0];
        if (g) prod[g.key] = Number(value);
      });
    }
  }
  cache.put('levelSettings', JSON.stringify(prod), 300);
  return prod;
}
