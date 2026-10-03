/**
 * Validasi.gs
 * Alur Validasi discrepancy (role Inventory/Admin)
 * (Hasil pemecahan dari Code.gs -- semua file .gs berbagi global scope yang
 * sama di Apps Script,
 * jadi urutan file tidak masalah, fungsi di file lain tetap bisa saling
 * panggil.)
 * 
 * v8.26.0 (Facility Management - Strict Isolation):
 * * Semua pemanggilan getRiwayatSheet_() SEKARANG meneruskan parameter username
 * agar
 * * membaca dari spreadsheet facility user tersebut.
 * * assignInitialPIC_() juga menerima facilityId agar PIC hanya dari facility
 * yang sama.
 */

// ---------- Validasi (role: inventory/admin) ----------
// Sejak v6: setiap item DISCREPANCY sudah ditunjuk 1 validator lewat
// assignNextValidator_()
// (round-robin antar role Inventory) saat submitCount. Validator (role
// Inventory) HANYA
// melihat item yang ditugaskan ke dirinya sendiri. Admin tetap bisa melihat
// SEMUA item Pending
// (peran pengawas/cadangan kalau validator berhalangan).
function getPendingValidasi(requesterUsername) {
  const info = requireRole_(requesterUsername, ['inventory', 'admin', 'developer']);
  // v8.26.0: Meneruskan requesterUsername agar baca dari spreadsheet facility
  // user
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  const items = [];
  data.forEach(function (row) {
    if (row[11] !== 'Pending') return;
    const assignedValidator = String(row[22] || '').trim();
    if (info.role === 'inventory' && assignedValidator !== info.displayName)
      return;
    items.push({ id: row[0], lokasi: row[2], article: row[3], description: row[4], qtySystem: Number(row[6]) });
  });
  return items;
}

function getPendingValidasiCountData_(info, username) {
  // v8.26.0: Meneruskan username agar baca dari spreadsheet facility user
  const sheet = getRiwayatSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const data = sheet.getRange(2, 12, lastRow - 1, 12).getValues(); // kolom L (12) s/d kolom W (23): cukup ambil status(12) & assigned(23)
  let count = 0;
  data.forEach(function (r) {
    const status = r[0]; // kolom L relatif = Status_Validasi
    const assignedValidator = String(r[11] || '').trim(); // kolom W relatif = Assigned_Validator (index 11 dari offset 12)
    if (status !== 'Pending') return;
    if (info.role === 'inventory' && assignedValidator !== info.displayName)
      return;
    count++;
  });
  return count;
}

function getPendingValidasiCount(requesterUsername) {
  const info = requireRole_(requesterUsername, ['inventory', 'admin', 'developer']);
  // v8.26.0: Meneruskan requesterUsername
  return getPendingValidasiCountData_(info, requesterUsername);
}

function submitValidasi(id, namaValidator, qtyValidasi, waktuKlien, facKlien) {
  requireRole_(namaValidator, ['inventory', 'admin', 'developer']);
  const aktual = Number(qtyValidasi);
  if (qtyValidasi === '' || qtyValidasi === null || qtyValidasi === undefined || isNaN(aktual) || aktual < 0) {
    return { success: false, message: 'Qty hasil validasi tidak valid.' };
  }

  const facInfo = requireUserFacility_(namaValidator);
  if (!facInfo) {
    return { success: false, message: PESAN_TANPA_FACILITY_ };
  }
  if (!facilityKlienCocok_(facKlien, facInfo)) {
    return { success: false, kode: 'FACILITY_BERUBAH', message: PESAN_FACILITY_BERUBAH_ };
  }

  // v8.31.0 (Asynchronous Queue Processing, sama pola dgn submitCount --
  // lihat ValidasiQueue.gs):
  // submitValidasi() TIDAK LAGI membuka spreadsheet facility di jalur ini.
  // Cukup tulis payload
  // ke Queue_Validasi (spreadsheet BOUND, bukan openById) lalu langsung
  // return. Worker
  // (processValidasiQueue_(), trigger 1 menit) yang benar-benar baca Riwayat,
  // tulis hasil,
  // assignInitialPIC_, dst -- membuka spreadsheet facility 1x per siklus utk
  // semua item yang
  // menumpuk, bukan 1x per item.
  try {
    queueValidasiPayload_(id, namaValidator, aktual, facInfo.id, waktuKlienSah_(waktuKlien));
  } catch (e) {
    return { success: false, message: 'Gagal masuk antrian validasi: ' + e.message };
  }

  return { success: true, queued: true, message: 'Validasi diterima.' };
}

// getAlasanList & saveAlasanDiscrepancy (v5) DIHAPUS di v6 -- alasan/kategori
// discrepancy
// sekarang dipilih bertahap di Task Investigasi (updateTaskStatus), bukan
// lagi di langkah Validasi.
// Lihat getInvestigasiFormData() & updateTaskStatus() di bawah.
