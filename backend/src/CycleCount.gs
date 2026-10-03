/**
 * CycleCount.gs
 * Alur Cycle Transaksi (role Outbound/Storing/Inbound/dst): ambil & submit count
 * (Hasil pemecahan dari Code.gs -- semua file .gs berbagi global scope yang sama di Apps Script,
 *  jadi urutan file tidak masalah, fungsi di file lain tetap bisa saling panggil.)
 *
 * v8.26.0 (Facility Management - Strict Isolation):
 * Semua pemanggilan getSheet_(), getRiwayatSheet_(), getLogAnomaliSheet_() SEKARANG
 * meneruskan parameter username agar membaca/menulis ke spreadsheet facility user tersebut.
 * assignNextValidator_() juga menerima facilityId agar validator hanya dari facility yang sama.
 */
// ---------- Cycle Count (role: cycle/inventory/admin) ----------
function getMyPendingTasks(username) {
  requireRole_(username, ALL_CYCLE_LIKE_ROLES);
  // v8.26.0: Meneruskan username agar baca dari spreadsheet facility user
  const sheet = getSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const uname = String(username || '').trim();
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const items = [];
  data.forEach(function (row) {
    const namaPetugas = String(row[9] || '').trim();
    const status = row[10];
    if (namaPetugas === uname && status === 'Pending') {
      const lokasi = row[2];
      // v8.10.0: area + grup alat dikirim per item -- dipakai client buat dropdown filter
      // "Area Gudang" & "Level Alat" di layar counting (lihat JsCycleValidasi.html), supaya
      // user gak lompat-lompat gudang begitu pindah antar batch/tahap equipment.
      const group = getLevelGroupByLocation_(lokasi);
      items.push({
        no: row[0], lokasi: lokasi, article: row[4], description: row[5],
        area: extractArea_(lokasi), levelGroupKey: group.key, levelGroupLabel: group.label,
        // v8.30.0 (Asynchronous Queue Processing): dikirim supaya CLIENT bisa estimasi
        // HIT/DISCREPANCY sendiri secara instan setelah submit -- server (Worker) tidak lagi
        // menghitung & mengembalikan hasil ini secara sinkron, lihat submitCount() di bawah.
        qtySystem: Number(row[7])
      });
    }
  });
  return items;
}

function getMyPendingCount(username) {
  requireRole_(username, ALL_CYCLE_LIKE_ROLES);
  // v8.26.0: Meneruskan username agar baca dari spreadsheet facility user
  const sheet = getSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const uname = String(username || '').trim();
  const data = sheet.getRange(2, 10, lastRow - 1, 2).getValues(); // J..K: Nama_Petugas, Status
  let count = 0;
  data.forEach(function (row) {
    if (String(row[0] || '').trim() === uname && row[1] === 'Pending') count++;
  });
  return count;
}

/**
 * Mencari nomor baris di "Data Count" untuk suatu No secara O(1) memakai asumsi
 * No berurutan mulai dari baris 2. Mengambil langsung seluruh baris data dalam 1 API call.
 */
function findDataCountRowAndValues_(sheet, no) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  
  // Estimasi cepat: No 1 di baris 2, No 2 di baris 3, dst
  let estimatedRow = Number(no) + 1;
  if (estimatedRow >= 2 && estimatedRow <= lastRow) {
    const rowValues = sheet.getRange(estimatedRow, 1, 1, HEADERS.length).getValues()[0];
    if (Number(rowValues[0]) === Number(no)) {
      return { rowIndex: estimatedRow, rowValues: rowValues };
    }
  }

  // Cek jika No tidak mulai dari 1: estimasi berdasarkan firstNo
  const firstNo = Number(sheet.getRange(2, 1).getValue());
  estimatedRow = (Number(no) - firstNo) + 2;
  if (estimatedRow >= 2 && estimatedRow <= lastRow) {
    const rowValues = sheet.getRange(estimatedRow, 1, 1, HEADERS.length).getValues()[0];
    if (Number(rowValues[0]) === Number(no)) {
      return { rowIndex: estimatedRow, rowValues: rowValues };
    }
  }

  // Fallback: scan kolom A
  const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (Number(colA[i][0]) === Number(no)) {
      const r = i + 2;
      const rowValues = sheet.getRange(r, 1, 1, HEADERS.length).getValues()[0];
      return { rowIndex: r, rowValues: rowValues };
    }
  }
  return null;
}

function submitCount(no, namaPetugas, qtyCount, waktuKlien, facKlien) {
  requireRole_(namaPetugas, ALL_CYCLE_LIKE_ROLES);
  const aktual = Number(qtyCount);
  if (qtyCount === '' || qtyCount === null || qtyCount === undefined || isNaN(aktual) || aktual < 0) {
    return { success: false, message: 'Qty hasil hitung tidak valid.' };
  }

  const facInfo = requireUserFacility_(namaPetugas);
  if (!facInfo) {
    return { success: false, message: PESAN_TANPA_FACILITY_ };
  }
  if (!facilityKlienCocok_(facKlien, facInfo)) {
    return { success: false, kode: 'FACILITY_BERUBAH', message: PESAN_FACILITY_BERUBAH_ };
  }

  // v8.30.0 (Asynchronous Queue Processing): submitCount() SEKARANG TIDAK LAGI membuka
  // spreadsheet facility SAMA SEKALI di jalur ini (dulu: SpreadsheetApp.openById() ke
  // spreadsheet eksternal facility -- inilah penyumbang terbesar 20 detik/submit, karena
  // TIDAK bisa di-cache antar eksekusi di Apps Script, mau lock-nya seefisien apapun).
  // Cukup nulis payload ke Queue_Submit_Cycle (spreadsheet BOUND/aktif milik script ini
  // sendiri -- getActiveSpreadsheet(), bukan openById, makanya cepat) lalu langsung
  // return ke user. Worker (processSubmitQueue_() di SubmitQueue.gs, trigger 1 menit) yang
  // nanti BENAR-BENAR baca Data Count, tulis "Selesai", & appendRiwayatCycle_ -- membuka
  // spreadsheet facility 1x per SIKLUS (bukan 1x per item), jadi biayanya diamortisasi ke
  // banyak submit sekaligus.
  //
  // KONSEKUENSI (sudah didiskusikan & disetujui):
  // 1. Hasil HIT/DISCREPANCY TIDAK lagi dihitung server secara sinkron di titik ini -- client
  //    yang estimasi sendiri dari qtySystem yang sudah dikirim lewat getMyPendingTasks().
  //    Worker tetap jadi single source of truth saat benar-benar menulis ke Riwayat.
  // 2. Status "Selesai" di Data Count baru ter-tulis setelah Worker jalan (maks ~1 menit) --
  //    item TIDAK langsung hilang dari sheet, tapi client sudah menyembunyikannya dari layar
  //    user ybs secara optimistic begitu masuk antrian sukses (lihat submitItem() di
  //    JsCycleValidasi.html), jadi tidak terasa oleh user yang submit.
  // 3. Item duplikat/double-submit disaring oleh Worker (cek status masih "Pending" sebelum
  //    proses), bukan lagi oleh itemClaimKey/facLock synchronous di sini.
  try {
    queueSubmitPayload_(no, namaPetugas, aktual, facInfo.id, waktuKlienSah_(waktuKlien));
  } catch (e) {
    return { success: false, message: 'Gagal masuk antrian submit: ' + e.message };
  }

  return { success: true, queued: true, message: 'Submission diterima.' };
}

function appendRiwayatCycle_(lokasi, article, description, qtyTransaksi, qtySystem, namaPetugas, qtyCount, selisih, hasilAwal, statusValidasi, hasilFinal, waktu, assignedValidator, addWhoTransaksi) {
  const sheet = getRiwayatSheet_(namaPetugas);
  const id = Utilities.getUuid();
  const tanggal = Utilities.formatDate(waktu, 'Asia/Jakarta', 'yyyy-MM-dd');
  const statusTaskAwal = hasilAwal === 'HIT' ? 'Selesai' : '';
  const newRow = [
    id, tanggal, lokasi, article, description, qtyTransaksi, qtySystem,
    namaPetugas, qtyCount, selisih, hasilAwal, statusValidasi, '', '', hasilFinal, waktu, '', '',
    statusTaskAwal, '', '', '', assignedValidator || '', '', waktu, addWhoTransaksi || '', ''
  ];
  // v8.29.3 (FIX row-collision risk + hemat 1 RPC): SEBELUMNYA pola-nya
  // "rowIndex = getLastRow()+1" (baca) lalu "getRange(rowIndex,...).setValues()" (tulis) --
  // 2 panggilan TERPISAH dengan jeda di antaranya. Sejak lock naik jadi multi-slot (sampai 16
  // proses submitCount jalan BERSAMAAN per facility, lihat FacilityManagement.gs), rowIndex
  // yang dihitung di sisi client itu gampang jadi BASI kalau proses lain sempat nyisip append
  // di antara baca & tulis kita -- 2 user bisa menghitung rowIndex yang SAMA dan salah satu
  // MENIMPA baris Riwayat milik yang lain (data hilang). appendRow() menyelesaikan posisi baris
  // di server Sheets pada saat permintaan diproses (bukan dari nilai lastRow yang kita baca
  // duluan), jadi race window-nya jauh lebih sempit -- sekaligus memangkas 1 RPC (getLastRow)
  // dari jalur kritis submitCount di SETIAP submit (bukan cuma yang discrepancy).
  sheet.appendRow(newRow);
  // rowIndex utk cache HANYA dipakai sebagai "titik awal pencarian" oleh findRiwayatRow_ --
  // fungsi itu SUDAH memvalidasi ID di rowIndex sebelum dipakai, dan otomatis fallback scan
  // kolom A kalau meleset (lihat SheetHelpers.gs), jadi aman walau baris ini kadang bergeser
  // karena race langka dengan proses lain yang appendRow() persis di detik yang sama.
  const rowIndex = sheet.getLastRow();
  riwayatRowCachePut_(id, rowIndex);

  // v8.29.2 (performa submitCount): Background caches (precompute) SEKARANG dititipkan ke
  // antrian (QueueDelta.gs), BUKAN dieksekusi langsung di sini. Sebelumnya applyRingkasanDelta_/
  // incrementCounterKey_ di titik ini = baca+tulis langsung ke spreadsheet facility (jalur
  // kritis submitCount, penyumbang lambat terbesar sejak facility split). Sekarang cukup
  // append cepat; diterapkan beneran oleh processDeferredQueueAllFacilities_() via trigger
  // tiap 1 menit -- lihat QueueDelta.gs untuk detail & cara pasang trigger-nya.
  try { queueRingkasanDeltaNewRow_(rowDateTag_(newRow), newRow, namaPetugas); } catch (e) { /* abaikan */ }
  try {
    if (newRow[11] === 'Pending' && newRow[22]) {
      queueCounterIncrements_([
        { key: 'pending_total', delta: 1 },
        { key: 'pending_validator:' + newRow[22], delta: 1 }
      ], namaPetugas);
    }
  } catch (e) { /* abaikan */ }
}

// Cache ringan ID Riwayat -> nomor baris, supaya lookup submitValidasi/saveAlasanDiscrepancy/
// closeTask (dipanggil segera setelah appendRiwayatCycle_/submitValidasi) tidak perlu scan
// penuh sheet Riwayat yang terus bertambah besar. TTL pendek + fallback scan kalau meleset
// (mis. baris dihapus/disisipkan manual di spreadsheet).
function getPendingBacklog(requesterUsername) {
  requireRole_(requesterUsername, ['admin', 'developer']);
  // v8.26.0: Meneruskan requesterUsername agar baca dari spreadsheet facility admin
  const sheet = getSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const map = {};
  data.forEach(function (row) {
    const tgl = String(row[1]);
    const status = row[10];
    if (!map[tgl]) map[tgl] = { total: 0, pending: 0 };
    map[tgl].total++;
    if (status === 'Pending') map[tgl].pending++;
  });
  return Object.keys(map)
    .filter(function (k) { return map[k].pending > 0; })
    .sort()
    .map(function (k) { return { tanggal: k, pending: map[k].pending, total: map[k].total }; });
}

function getBacklogDetailByDate(tanggal, requesterUsername) {
  requireRole_(requesterUsername, ['admin', 'developer']);
  // v8.26.0: Meneruskan requesterUsername agar baca dari spreadsheet facility admin
  const sheet = getSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const tglTarget = String(tanggal);
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const map = {};
  data.forEach(function (row) {
    if (String(row[1]) !== tglTarget) return;
    const namaPetugas = String(row[9] || '').trim() || '(belum ditugaskan)';
    const status = row[10];
    if (!map[namaPetugas]) map[namaPetugas] = { namaPetugas: namaPetugas, total: 0, pending: 0 };
    map[namaPetugas].total++;
    if (status === 'Pending') map[namaPetugas].pending++;
  });
  return Object.keys(map)
    .map(function (k) { return map[k]; })
    .filter(function (u) { return u.pending > 0; })
    .sort(function (a, b) { return b.pending - a.pending; });
}
