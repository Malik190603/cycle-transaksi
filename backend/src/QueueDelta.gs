/**
 * QueueDelta.gs (v8.29.2 - performa submitCount)
 *
 * Mekanisme "deferred queue" untuk update Ringkasan_Harian & Antrian_Aktif,
 * supaya jalur
 * kritis submitCount() TIDAK perlu baca+tulis langsung ke 2 sheet precompute
 * itu.
 * (Lihat catatan lengkap di Config.gs bagian 'QUEUE DEFERRED UPDATES'.)
 *
 * ALUR:
 * 1. submitCount() -> appendRiwayatCycle_() -> queueRingkasanDeltaNewRow_() &
 *    queueCounterIncrements_() -- HANYA menulis (append), TIDAK baca apapun
 *    dulu.
 *    Ini 1-2 API call cepat, menggantikan 6-8 call baca+tulis yang
 *    sebelumnya ada
 *    di jalur kritis.
 * 2. Trigger time-driven (tiap 1 menit, lihat installDeferredQueueTrigger_()
 *    di bawah)
 *    menjalankan processDeferredQueueAllFacilities_() -- untuk SETIAP
 *    facility aktif, baca
 *    semua baris antrian yang menumpuk, terapkan ke Ringkasan_Harian/
 *    Antrian_Aktif via fungsi
 *    LAMA yang SUDAH ADA (applyRingkasanDelta_, incrementCounterKey_ --
 *    TIDAK diubah sama
 *    sekali, strictly additive), lalu HAPUS baris yang sudah diproses.
 *
 * PENTING -- WAJIB dijalankan sekali manual setelah deploy:
 * - Buka Apps Script editor -> pilih fungsi "installDeferredQueueTrigger_"
 *   dari dropdown ->
 * - klik Run. Ini memasang trigger time-driven yang menjalankan
 *   processDeferredQueueAllFacilities_() tiap 1 menit. Tanpa langkah ini,
 *   antrian akan terus
 *   menumpuk dan Ringkasan_Harian/Antrian_Aktif TIDAK akan ter-update
 *   (dashboard KPI jadi
 *   basi/stuck), walaupun submitCount() sendiri tetap sukses & cepat.
 *
 * SAFETY NET: rebuildRingkasanHarian_() & rebuildAntrianAktif_() (sudah ada
 * sebelumnya, tidak
 * disentuh sama sekali) tetap bisa dipakai kapan saja untuk resync total
 * dari nol kalau
 * dicurigai ada delta yang somehow ke-skip/gagal proses.
 */

// =========================================================================
// SISI PENULIS (dipanggil dari submitCount / appendRiwayatCycle_)
// =========================================================================

/**
 * Penitipkan 1 baris Riwayat BARU (newRow; oldRow selalu null di titik
 * submitCount) ke
 * antrian, untuk nanti diterapkan ke Ringkasan_Harian oleh
 * processDeferredQueueAllFacilities_(). HANYA 1x appendRow -- tanpa baca
 * apapun.
 *
 * @param {string} tanggal - dari rowDateTag_(newRow)
 * @param {Array} newRow - array 1 baris Riwayat lengkap (urutan sesuai
 * RIWAYAT_HEADERS)
 * @param {string} [username]
 */
function queueRingkasanDeltaNewRow_(tanggal, newRow, username) {
  if (!tanggal) return;
  const sheet = getQueueRingkasanSheet_(username);
  // v8.31.1 (FIX bug konkurensi @ 15 akun): appendRow dibungkus lock --
  // lihat catatan lengkap
  // di appendRowLocked_ (Utils.gs). Sheet ini dipanggil dari DALAM Worker
  // yang SUDAH memegang
  // facLock (16-slot, BUKAN exclusive) -- 2 slot-holder berbeda utk facility
  // yang SAMA (mis.
  // Worker submit + closePlusMinusPair sinkron) tetap bisa appendRow
  // bersamaan ke sheet ini.
  appendRowLocked_(sheet, [new Date(), tanggal, JSON.stringify(newRow)]);
}

/**
 * Penitipkan 1 atau lebih perubahan counter Antrian_Aktif (mis.
 * pending_total,
 * pending_validator:<nama>) ke antrian, tanpa baca apapun dulu. Ditulis
 * DALAM SATU kali
 * setValues (bukan appendRow satu-satu per item) supaya tetap 1 round-trip
 * meskipun ada
 * >1 counter yang berubah bersamaan (kasus submitCount: pending_total +
 * pending_validator:X sekaligus).
 *
 * @param {Array<{key:string, delta:number}>} items
 * @param {string} [username]
 */
function queueCounterIncrements_(items, username) {
  if (!items || !items.length) return;
  const sheet = getQueueCounterSheet_(username);
  const now = new Date();
  const rows = items.map(function (it) { return [now, it.key, it.delta]; });
  // v8.31.1 (FIX bug konkurensi @ 15 akun): dulu getLastRow()+setValues()
  // TANPA lock -- 2
  // eksekusi bersamaan bisa sama-sama simpulkan "baris kosong berikutnya =
  // row X" & saling
  // timpa. Sekarang dibungkus lock (appendRowsLocked_ di Utils.gs, hitung
  // startRow DI DALAM
  // genggaman lock, bukan sebelumnya).
  appendRowsLocked_(sheet, rows);
}

// =========================================================================
// SISI PEMROSES (dipanggil dari trigger time-driven, BUKAN dari submitCount)
// =========================================================================

/**
 * Proses antrian delta Ringkasan_Harian & Antrian_Aktif untuk SATU facility.
 * Dibungkus lock
 * per-facility yang SAMA dengan submitCount (acquireFacilityLock_) supaya
 * tidak balapan
 * dengan submit yang sedang berjalan di facility itu.
 *
 * @param {string} facilityUsernameHint - username APAPUN yang valid &
 * ter-assign ke facility
 *        ini (dipakai getOperasionalSpreadsheet_/getUserFacility untuk resolve
 * spreadsheet yang
 *        benar).
 */
function processDeferredQueueForFacility_(facilityUsernameHint) {
  const facInfo = getUserFacility(facilityUsernameHint);
  const facilityId = facInfo ? facInfo.id : null;

  // v8.31.1 (konsistensi dengan Submit/Validasi/Task Queue -- cegah 2 worker
  // run tumpang
  // tindih memproses facility yang SAMA, yang bisa menyebabkan delta yang
  // sama diterapkan
  // 2x ke Ringkasan_Harian/Antrian_Aktif sebelum salah satunya sempat
  // menghapus baris queue).
  const cache = CacheService.getScriptCache();
  const workerGuardKey = 'deferredQueueWorkerLock_' + (facilityId || 'default');
  if (cache.get(workerGuardKey)) return;
  try { cache.put(workerGuardKey, '1', 55); } catch (e) { /* abaikan */ }

  const facLock = acquireFacilityLock_(facilityId, 15000);
  if (!facLock) { try { cache.remove(workerGuardKey); } catch (e) { /* abaikan */ } return; } // facility sedang sibuk -- coba lagi di jadwal trigger berikutnya (1 menit lagi)
  try {
    // --- Ringkasan_Harian ---
    const qrSheet = getQueueRingkasanSheet_(facilityUsernameHint);
    const qrLastRow = qrSheet.getLastRow();
    if (qrLastRow >= 2) {
      const qrData = qrSheet.getRange(2, 1, qrLastRow - 1, QUEUE_RINGKASAN_HEADERS.length).getValues();
      qrData.forEach(function (row) {
        const tanggal = row[1];
        try {
          const newRow = JSON.parse(row[2]);
          applyRingkasanDelta_(tanggal, null, newRow, facilityUsernameHint);
        } catch (e) { /* baris korup -- lewati, jangan sampai macetin baris lain */ }
      });
      qrSheet.deleteRows(2, qrLastRow - 1);
    }

    // --- Antrian_Aktif (counter) ---
    const qcSheet = getQueueCounterSheet_(facilityUsernameHint);
    const qcLastRow = qcSheet.getLastRow();
    if (qcLastRow >= 2) {
      const qcData = qcSheet.getRange(2, 1, qcLastRow - 1, QUEUE_COUNTER_HEADERS.length).getValues();
      const totals = {}; // key -> total delta, digabung dulu SEBELUM ditulis
      // hemat round-trip
      qcData.forEach(function (row) {
        const key = String(row[1] || '').trim();
        const delta = Number(row[2]) || 0;
        if (!key) return;
        totals[key] = (totals[key] || 0) + delta;
      });
      Object.keys(totals).forEach(function (key) {
        if (totals[key] !== 0) incrementCounterKey_(key, totals[key], facilityUsernameHint);
      });
      qcSheet.deleteRows(2, qcLastRow - 1);
    }
  } finally {
    facLock.release();
    try { cache.remove(workerGuardKey); } catch (e) { /* abaikan */ }
  }
}

/**
 * Dipanggil trigger time-driven (tiap 1 menit) -- proses antrian SEMUA
 * facility aktif
 * satu-satu. Berjalan sebagai proses TERPISAH dari submitCount, jadi TIDAK
 * menambah beban
 * ke request submit user manapun.
 */
function processDeferredQueueAllFacilities_() {
  const facSheet = getFacilitySheet_();
  const lastRow = facSheet.getLastRow();
  if (lastRow < 2) return;
  const facValues = facSheet.getRange(2, 1, lastRow - 1, FACILITY_HEADERS.length).getValues();

  // Butuh 1 username yang ter-assign ke tiap facility supaya
  // getOperasionalSpreadsheet_ bisa
  // resolve spreadsheet-nya. Ambil dari User_Facility_Assignment (baris
  // TERAKHIR per facility
  // menang -- assignment paling baru).
  const assignSheet = getUserFacilityAssignmentSheet_();
  const assignLastRow = assignSheet.getLastRow();
  const usernameByFacility = {};
  if (assignLastRow >= 2) {
    const assignValues = assignSheet.getRange(2, 1, assignLastRow - 1, 2).getValues();
    assignValues.forEach(function (row) {
      const uname = String(row[0] || '').trim();
      const facId = String(row[1] || '').trim();
      if (uname && facId) usernameByFacility[facId] = uname;
    });
  }

  facValues.forEach(function (row) {
    const facId = String(row[0] || '').trim();
    const status = String(row[5] || '').trim();
    if (!facId || status === 'Nonaktif') return;
    const uname = usernameByFacility[facId];
    if (!uname) return; // facility belum punya user ter-assign -- tidak ada antrian utk diproses

    try {
      processDeferredQueueForFacility_(uname);
    } catch (e) {
      try { catatLogSistem_('Queue Deferred', 'Gagal proses antrian facility ' + facId + ': ' + e.message); } catch (e2) { /* abaikan */ }
    }
  });
}

// =========================================================================
// SETUP (dijalankan MANUAL SEKALI dari Apps Script editor setelah deploy patch
// ini)
// =========================================================================

/**
 * Jalankan fungsi ini SEKALI SAJA secara manual (pilih dari dropdown fungsi
 * di Apps Script
 * editor, klik Run) untuk memasang trigger time-driven yang menjalankan
 * processDeferredQueueAllFacilities_() tiap 1 menit. Aman dijalankan ulang
 * -- otomatis
 * menghapus trigger lama dengan nama fungsi yang sama dulu supaya tidak
 * dobel.
 */
function installDeferredQueueTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processDeferredQueueAllFacilities_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processDeferredQueueAllFacilities_')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('Trigger terpasang: processDeferredQueueAllFacilities_ tiap 1 menit.');
}
