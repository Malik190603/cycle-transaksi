/**
 * SubmitQueue.gs (v0.10.0 - Asynchronous Queue Processing untuk submitCount)
 * 
 * Lihat catatan arsitektur lengkap di Config.gs bagian "SUBMIT QUEUE".
 * 
 * ALUR:
 * 1. submitCount() (CycleCount.gs) -> queueSubmitPayload_() -- HANYA 1x appendRow ke
 *    Queue_Submit_Cycle, sheet yang hidup di spreadsheet BOUND/aktif milik script ini
 *    sendiri (getActiveSpreadsheet(), BUKAN openById ke facility). Langsung return ke user.
 * 2. Trigger time-driven (tiap 1 menit, lihat installSubmitQueueTrigger_() di bawah)
 *    menjalankan processSubmitQueue_() -- baca SEMUA baris antrian SEKALI, kelompokkan per
 *    facility, lalu untuk TIAP facility: buka spreadsheet-nya SEKALI (bukan 1x per item).
 *    proses semua item yang menumpuk lewat fungsi LAMA yang SUDAH ADA
 *    (findDataCountRowAndValues_, appendRiwayatCycle_, assignNextValidator_ -- TIDAK diubah
 *    sama sekali, strictly additive).
 * 3. Item yang gagal permanen (tidak ketemu / sudah "Selesai" sebelumnya -- duplikat) DICATAT
 *    ke Log_Sistem lalu tetap dibuang dari antrian (tidak di-retry selamanya). Item yang gagal
 *    karena facility SEDANG SIBUK (lock penuh) dikembalikan ke antrian utk dicoba lagi siklus
 *    berikutnya.
 * 
 * PENTING -- WAJIB dijalankan sekali manual setelah deploy:
 * * Buka Apps Script editor -> pilih fungsi "installSubmitQueueTrigger_" dari dropdown ->
 * * klik Run. Tanpa langkah ini, submitCount() akan tetap CEPAT (antrian tetap ditulis),
 * * TAPI tidak ada yang memproses -- item tidak akan pernah benar-benar tercatat "Selesai"
 *   di Data Count / Riwayat.
 */

// =========================================================================
// SISI PENULIS (dipanggil dari submitCount())
// =========================================================================

/**
 * Menitipkan 1 payload submit ke antrian. HANYA 1x appendRow ke spreadsheet BOUND -- tidak
 * membuka spreadsheet facility sama sekali.
 */
function queueSubmitPayload_(no, namaPetugas, qtyCount, facilityId) {
  const sheet = getSubmitQueueSheet_();
  appendRowLocked_(sheet, [new Date(), no, namaPetugas, qtyCount, facilityId]);
}


// =========================================================================
// SISI PEMROSES (dipanggil dari trigger time-driven, BUKAN dari submitCount)
// =========================================================================

/**
 * Dipanggil trigger time-driven (tiap 1 menit). Baca antrian SEKALI, kelompokkan per
 * facility, proses tiap facility dengan spreadsheet-nya dibuka SEKALI SAJA.
 */
function processSubmitQueue_() {
  const qSheet = getSubmitQueueSheet_();
  const lastRow = qSheet.getLastRow();
  if (lastRow < 2) return;

  const values = qSheet.getRange(2, 1, lastRow - 1, SUBMIT_QUEUE_HEADERS.length).getValues();
  const byFacility = {};
  values.forEach(function (row, idx) {
    const facilityId = String(row[4] || '').trim();
    if (!byFacility[facilityId]) byFacility[facilityId] = [];
    byFacility[facilityId].push(idx);
  });

  // v8.31.1 (FIX bug konkurensi @ 15 akun -- lihat catatan deleteProcessedQueueRows_ di
  // Utils.gs): TIDAK LAGI 'baca-semua -> clear-semua -> tulis-ulang-sisa' (rawan menghapus
  // baris baru yang masuk SAAT Worker sedang jalan). Sekarang HANYA kumpulkan nomor baris
  // AKTUAL yang benar-benar sudah dipproses, lalu dihapus SATU-SATU di akhir -- baris yang
  // belum diproses (facility sibuk) ATAU baris baru yang masuk belakangan TIDAK PERNAH disentuh.
  const processedRowNumbers = []; // nomor baris AKTUAL (1-based di sheet) yang sudah diproses
  const cache = CacheService.getScriptCache();

  Object.keys(byFacility).forEach(function (facilityId) {
    const indices = byFacility[facilityId];
    if (!facilityId) {
      // FacilityId kosong (data korup/lama) -- log & buang, jangan sampai nyangkut selamanya
      indices.forEach(function (i) {
        catatLogSistem_('Worker Submit', 'Baris antrian tanpa FacilityId dibuang: ' + JSON.stringify(values[i]));
        processedRowNumbers.push(i + 2);
      });
      return;
    }

    // Guard tambahan: cegah 2 EKSEKUSI WORKER yang tumpang tindih memproses facility yang
    // SAMA secara bersamaan (beda dgn facLock 16-slot yang dipakai submitValidasi dkk --
    // ini khusus memastikan HANYA 1 worker ran yang pegang facility ini sekaligus, supaya
    // tidak ada 2 worker yang sama-sama baca status "Pending" sebelum salah satu sempat
    // menulis "Selesai"). TTL 55 detik -- sedikit di bawah interval trigger (1 menit).
    const workerGuardKey = 'submitQueueWorkerLock_' + facilityId;
    if (cache.get(workerGuardKey)) {
      return; // facility sedang diproses worker run lain -- baris DIBIARKAN, coba lagi siklus berikutnya
    }
    try { cache.put(workerGuardKey, '1', 55); } catch (e) { /* abaikan */ }

    const facLock = acquireFacilityLock_(facilityId, 10000);
    if (!facLock) {
      // Facility sedang penuh (16 slot terpakai oleh submitValidasi dkk) -- baris DIBIARKAN, coba lagi siklus berikutnya
      try { cache.remove(workerGuardKey); } catch (e) { /* abaikan */ }
      return;
    }

    try {
      // Urutkan berdasarkan Timestamp (kolom 0) supaya diproses FIFO sesuai urutan submit asli
      const sortedIndices = indices.slice().sort(function (a, b) {
        return new Date(values[a][0]) - new Date(values[b][0]);
      });

      sortedIndices.forEach(function (i) {
        const row = values[i];
        const no = row[1];
        const namaPetugas = String(row[2] || '').trim();
        const qtyCount = row[3];

        try {
          const result = processSingleQueuedSubmit_(no, namaPetugas, qtyCount, facilityId);
          if (!result.success) {
            catatLogSistem_('Worker Submit', 'Skip No ' + no + ' (' + namaPetugas + '): ' + result.message);
          }
        } catch (e) {
          catatLogSistem_('Worker Submit', 'ERROR No ' + no + ' (' + namaPetugas + '): ' + e.message);
        }

        // Baik sukses, skip (duplikat/tidak ketemu), maupun error tak terduga -- baris dianggap
        // SELESAI diproses (hindari infinite retry loop utk item yang memang gagal permanen).
        processedRowNumbers.push(i + 2); // +2: idx 0-based -> baris sheet 1-based, +1 lagi utk header
      });
    } finally {
      facLock.release();
      try { cache.remove(workerGuardKey); } catch (e) { /* abaikan */ }
    }
  });

  // Hapus HANYA baris yang benar-benar sudah diproses (aman terhadap baris baru yang masuk
  // selagi Worker jalan -- lihat catatan deleteProcessedQueueRows_ di Utils.gs).
  deleteProcessedQueueRows_(qSheet, processedRowNumbers);
}

/**
 * Proses SATU item queue: reuse PERSIS logic yang dulu ada langsung di submitCount()
 * (findDataCountRowAndValues_, tulis Data Count, appendRiwayatCycle_, assignNextValidator_).
 * Dipanggil dari dalam processSubmitQueue_() yang SUDAH memegang facLock facility terkait --
 * TIDAK perlu lock lagi di sini.
 */
function processSingleQueuedSubmit_(no, namaPetugas, qtyCount, facilityId) {
  const aktual = Number(qtyCount);
  const sheet = getSheet_(namaPetugas);
  const itemData = findDataCountRowAndValues_(sheet, no);
  if (!itemData) {
    return { success: false, message: 'Item tidak ditemukan (kemungkinan dihapus/direvisi Admin).' };
  }

  const rowIndex = itemData.rowIndex;
  const rowValues = itemData.rowValues;
  if (rowValues[10] !== 'Pending') {
    return { success: false, message: 'Sudah diproses sebelumnya (duplikat submit).' };
  }

  const lokasi = rowValues[2];
  const article = rowValues[4];
  const description = rowValues[5];
  const qtyTransaksi = rowValues[6];
  const qtySystem = Number(rowValues[7]);
  const addWhoTransaksi = rowValues[15];
  const selisih = aktual - qtySystem;
  const hasilAwal = selisih === 0 ? 'HIT' : 'DISCREPANCY';
  const waktu = new Date();

  // Tulis ke Data Count (1 API call)
  sheet.getRange(rowIndex, 10, 1, 6).setValues([[namaPetugas, 'Selesai', aktual, selisih, hasilAwal, waktu]]);

  const statusValidasi = hasilAwal === 'HIT' ? 'Tidak Perlu' : 'Pending';
  const hasilFinal = hasilAwal === 'HIT' ? 'HIT' : '';
  const assignedValidator = hasilAwal === 'HIT' ? '' : assignNextValidator_(facilityId);

  appendRiwayatCycle_(lokasi, article, description, qtyTransaksi, qtySystem, namaPetugas, aktual, selisih, hasilAwal, statusValidasi, hasilFinal, waktu, assignedValidator, addWhoTransaksi);

  return { success: true, hasil: hasilAwal, selisih: selisih };
}

/**
 * v8.31.1: rewriteSubmitQueueSheet_() DIHAPUS (diganti deleteProcessedQueueRows_ di Utils.gs)
 * -- pola lama "clear semua + tulis ulang sisa" rawan menghapus baris baru yang masuk PAS
 * Worker sedang jalan. Lihat catatan lengkap di Utils.gs.
 */

// =========================================================================
// SETUP (jalankan MANUAL SEKALI dari Apps Script editor setelah deploy patch ini)
// =========================================================================

/**
 * Jalankan fungsi ini SEKALI SAJA secara manual (pilih dari dropdown fungsi di Apps Script
 * editor, klik Run) untuk memasang trigger time-driven yang menjalankan processSubmitQueue_()
 * tiap 1 menit. Aman dijalankan ulang -- otomatis menghapus trigger lama dengan nama fungsi
 * yang sama dulu supaya tidak dobel.
 */
function installSubmitQueueTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processSubmitQueue_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processSubmitQueue_')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('Trigger terpasang: processSubmitQueue_ tiap 1 menit.');
}
