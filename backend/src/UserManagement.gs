/**
 * UserManagement.gs
 * v8.29.0 -- CRUD user Master (Username, Role, Status Aktif/Nonaktif, ID_Facility)
 * diarahkan ke sheet Master_User.
 */

const MASTER_ROLE_VALID_ = ['admin', 'inventory', 'outbound', 'storing', 'inbound', 'lp', 'maintenance', 'developer'];

/**
 * Daftar lengkap user Master dari Master_User (termasuk yg Nonaktif).
 */
function getDaftarUserMaster(requesterUsername) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Kelola User.' };
  }

  // Baca dari Master_User
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    const muSheet = getMasterUserSheetWithFallback_();
    if (muSheet) {
      const lastRow = muSheet.getLastRow();
      const users = [];
      if (lastRow >= 2) {
        const values = muSheet.getRange(2, 1, lastRow - 1, 4).getValues(); // A=User, B=Role, C=Status, D=Facility
        values.forEach(function (r, idx) {
          const username = String(r[0] || '').trim();
          if (!username) return;
          const facId = String(r[3] || '').trim();
          let facNama = '';
          if (facId && typeof getFacilityInfoById_ === 'function') {
            const facInfo = getFacilityInfoById_(facId);
            if (facInfo) facNama = facInfo.nama;
          }
          users.push({
            rowIndex: idx + 2,
            username: username,
            role: String(r[1] || '').trim(),
            status: isBarisUserAktif_(r[2]) ? 'Aktif' : 'Nonaktif',
            facilityId: facId,
            facilityNama: facNama
          });
        });
      }
      return { success: true, users: users, roleValid: MASTER_ROLE_VALID_ };
    }
  }

  // Fallback ke lama
  const sheet = getMasterSheet_();
  if (!sheet) return { success: false, message: 'Sheet Master tidak ditemukan.' };

  const lastRow = sheet.getLastRow();
  const users = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 3, lastRow - 1, 3).getValues(); // C=User, D=Role, E=Status
    values.forEach(function (r, idx) {
      const username = String(r[0] || '').trim();
      if (!username) return;
      users.push({
        rowIndex: idx + 2,
        username: username,
        role: String(r[1] || '').trim(),
        status: isBarisUserAktif_(r[2]) ? 'Aktif' : 'Nonaktif',
        facilityId: '',
        facilityNama: ''
      });
    });
  }
  return { success: true, users: users, roleValid: MASTER_ROLE_VALID_ };
}

const PESAN_KHUSUS_DEVELOPER_ = 'Akun dan peran Developer hanya bisa diubah oleh Developer.';
function roleDeveloper_(roleRaw) {
  const r = String(roleRaw || '').trim().toLowerCase();
  return r === 'developer' || r === 'dev' || r === 'dewa';
}
function requesterDeveloper_(requesterUsername) {
  return isUsernameRoleDeveloper_(String(requesterUsername || '').trim().toLowerCase());
}

/**
 * Tambah user baru (dan facility-nya) ke Master_User.
 */
function tambahUserMaster(requesterUsername, usernameBaru, roleBaru, facilityId) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Kelola User.' };
  }

  const uname = String(usernameBaru || '').trim();
  const role = String(roleBaru || '').trim().toLowerCase();
  let facId = String(facilityId || '').trim();
  let facNama = ''; // nama facility untuk pesan hasil (bukan kode internal FAC-...)

  if (!uname) return { success: false, message: 'NIK/Username tidak boleh kosong.' };
  if (MASTER_ROLE_VALID_.indexOf(role) === -1) {
    return {
      success: false,
      message: 'Role tidak valid. Pilih salah satu: ' + MASTER_ROLE_VALID_.join(', ') + '.'
    };
  }
  if (role === 'developer' && !requesterDeveloper_(requesterUsername)) {
    return { success: false, message: PESAN_KHUSUS_DEVELOPER_ };
  }

  // Improvement: Auto-assign user baru ke facility yang sedang aktif dipakai oleh admin/requester
  if (!facId && role !== 'developer') {
    if (typeof getUserFacility === 'function') {
      const requesterFac = getUserFacility(requesterUsername);
      if (requesterFac) { facId = requesterFac.id; facNama = requesterFac.nama || requesterFac.name || ''; }
    }
  }

  let sheet = null;
  let isNewFormat = false;
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    sheet = getMasterUserSheetWithFallback_();
    if (sheet) isNewFormat = true;
  }
  if (!sheet) sheet = getMasterSheet_();
  if (!sheet) return { success: false, message: 'Sheet Master tidak ditemukan.' };

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const colIdx = isNewFormat ? 1 : 3;
    const existing = sheet.getRange(2, colIdx, lastRow - 1, 1).getValues();
    for (let i = 0; i < existing.length; i++) {
      if (String(existing[i][0] || '').trim().toLowerCase() === uname.toLowerCase()) {
        return { success: false, message: 'NIK/Username "' + uname + '" sudah terdaftar (baris ' + (i + 2) + ').' };
      }
    }
  }

  const targetRow = Math.max(lastRow + 1, 2);
  if (isNewFormat) {
    sheet.getRange(targetRow, 1).setValue(uname);
    sheet.getRange(targetRow, 2).setValue(role);
    sheet.getRange(targetRow, 3).setValue('Aktif');
    sheet.getRange(targetRow, 4).setValue(facId);

    // Sync juga ke User_Facility_Assignment
    if (facId && typeof assignUserKeFacility === 'function') {
      assignUserKeFacility(requesterUsername, uname, facId);
    }
  } else {
    sheet.getRange(targetRow, 3).setValue(uname);
    sheet.getRange(targetRow, 4).setValue(role);
    sheet.getRange(targetRow, 5).setValue('Aktif');
  }

  clearMasterCache_();
  catatLogPerubahanConfig_(requesterUsername, 'Manajemen User', 'Tambah user "' + uname + '" role ' + role + (facId ? ' ke facility ' + facId : '') + '.');

  return { success: true, message: 'User "' + uname + '" berhasil ditambahkan' + (facId ? ' ke facility ' + (facNama || facId) + '.' : '.') };
}

function updateRoleUserMaster(requesterUsername, rowIndex, usernameKonfirmasi, roleBaru, facilityId) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Kelola User.' };
  }

  const role = String(roleBaru || '').trim().toLowerCase();
  // facilityId hanya diubah bila memang dikirim. Tanpa pengecekan ini, mengganti peran saja
  // (tanpa argumen facility) mengosongkan kolom ID_Facility dan user itu hilang dari daftar petugas.
  const facDikirim = facilityId !== undefined && facilityId !== null;
  const facId = String(facilityId || '').trim();

  if (MASTER_ROLE_VALID_.indexOf(role) === -1) {
    return { success: false, message: 'Role tidak valid.' };
  }

  let sheet = null;
  let isNewFormat = false;
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    sheet = getMasterUserSheetWithFallback_();
    if (sheet) isNewFormat = true;
  }
  if (!sheet) sheet = getMasterSheet_();
  if (!sheet) return { success: false, message: 'Sheet Master tidak ditemukan.' };

  const colIdx = isNewFormat ? 1 : 3;
  const roleColIdx = isNewFormat ? 2 : 4;
  const actualUsername = String(sheet.getRange(rowIndex, colIdx).getValue() || '').trim();

  if (!actualUsername || actualUsername.toLowerCase() !== String(usernameKonfirmasi || '').trim().toLowerCase()) {
    return { success: false, message: 'Data user sudah berubah sejak halaman dibuka.' };
  }

  const roleLama = String(sheet.getRange(rowIndex, roleColIdx).getValue() || '').trim();
  // Developer adalah super user lintas facility: peran itu hanya boleh diberikan, atau dicabut,
  // oleh Developer lain. Tanpa ini pemegang akses Config mana pun bisa mengangkat dirinya sendiri.
  if ((role === 'developer' || roleDeveloper_(roleLama)) && roleDeveloper_(roleLama) !== (role === 'developer') && !requesterDeveloper_(requesterUsername)) {
    return { success: false, message: PESAN_KHUSUS_DEVELOPER_ };
  }
  sheet.getRange(rowIndex, roleColIdx).setValue(role);

  if (isNewFormat && facDikirim) {
    sheet.getRange(rowIndex, 4).setValue(facId);
    // Sync juga ke User_Facility_Assignment
    if (facId && typeof assignUserKeFacility === 'function') {
      assignUserKeFacility(requesterUsername, actualUsername, facId);
    }
  }

  clearMasterCache_();
  clearUserRoleCache_(actualUsername);
  catatLogPerubahanConfig_(requesterUsername, 'Manajemen User', 'Data "' + actualUsername + '" diperbarui.');

  return { success: true, message: 'Data "' + actualUsername + '" berhasil diperbarui.' };
}

function setStatusUserMaster(requesterUsername, rowIndex, usernameKonfirmasi, statusBaru) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return { success: false, message: 'NIK ini tidak punya akses ke Kelola User.' };
  }

  const status = (String(statusBaru || '').trim().toLowerCase() === 'nonaktif') ? 'Nonaktif' : 'Aktif';

  let sheet = null;
  let isNewFormat = false;
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    sheet = getMasterUserSheetWithFallback_();
    if (sheet) isNewFormat = true;
  }
  if (!sheet) sheet = getMasterSheet_();
  if (!sheet) return { success: false, message: 'Sheet Master tidak ditemukan.' };

  const colIdx = isNewFormat ? 1 : 3;
  const statusColIdx = isNewFormat ? 3 : 5;
  const actualUsername = String(sheet.getRange(rowIndex, colIdx).getValue() || '').trim();

  if (!actualUsername || actualUsername.toLowerCase() !== String(usernameKonfirmasi || '').trim().toLowerCase()) {
    return { success: false, message: 'Data user sudah berubah.' };
  }
  if (status === 'Nonaktif') {
    // Menonaktifkan akun sendiri = terkunci di luar (tidak bisa masuk lagi untuk mengaktifkannya).
    if (actualUsername.toLowerCase() === String(requesterUsername || '').trim().toLowerCase()) {
      return { success: false, message: 'Akun sendiri tidak bisa dinonaktifkan. Minta pemegang akses Config lain.' };
    }
    const roleTarget = String(sheet.getRange(rowIndex, isNewFormat ? 2 : 4).getValue() || '');
    if (roleDeveloper_(roleTarget) && !requesterDeveloper_(requesterUsername)) {
      return { success: false, message: PESAN_KHUSUS_DEVELOPER_ };
    }
  }

  sheet.getRange(rowIndex, statusColIdx).setValue(status);
  clearMasterCache_();
  clearUserRoleCache_(actualUsername); // tanpa ini user yang dinonaktifkan masih bisa login sampai cache 5 menit habis
  catatLogPerubahanConfig_(requesterUsername, 'Manajemen User', 'Status "' + actualUsername + '" diubah jadi ' + status + '.');

  return {
    success: true,
    message: 'Status "' + actualUsername + '" berhasil diubah menjadi ' + status + '.'
  };
}
