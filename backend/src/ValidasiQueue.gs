/**
 * ValidasiQueue.gs (v8.31.0 - Asynchronous Queue Processing untuk submitValidasi)
 * 
 * Pola identik dengan SubmitQueue.gs (lihat catatan lengkap di sana & di Config.gs bagian
 * "SUBMIT QUEUE"). submitValidasi() (Validasi.gs) SEKARANG cuma menitipkan payload ke
 * Queue_Validasi (spreadsheet BOUND, bukan openById ke facility), lalu processValidasiQueue_()
 * (trigger 1 menit) yang benar-benar baca Riwayat, tulis hasil, assignInitialPIC_, dst --
 * membuka spreadsheet facility 1x per siklus, bukan 1x per item.
 * 
 * processSingleQueuedValidasi_() adalah body PERSIS dari submitValidasi() versi lama (v8.29.x)
 * -- tidak ada logic yang berubah, cuma dipindah dari jalur sinkron ke Worker.
 * 
 * PENTING -- WAJIB dijalankan sekali manual setelah deploy:
 * * Buka Apps Script editor -> pilih fungsi "installValidasiQueueTrigger_" dari dropdown ->
 * * klik Run.
 */

// =========================================================================
// SISI PENULIS (dipanggil dari submitValidasi())
// =========================================================================

function queueValidasiPayload_(id, namaValidator, qtyValidasi, facilityId, waktuHitung) {
  const sheet = getValidasiQueueSheet_();
  appendRowLocked_(sheet, [new Date(), id, namaValidator, qtyValidasi, facilityId, waktuHitung || '']);
}


// =========================================================================
// SISI PEMROSES (dipanggil dari trigger time-driven, BUKAN dari submitValidasi)
// =========================================================================

/**
 * Dipanggil trigger time-driven (tiap 1 menit). Baca antrian SEKALI, kelompokkan per
 * facility, proses tiap facility dengan spreadsheet-nya dibuka SEKALI SAJA.
 */
function processValidasiQueue_() {
  const qSheet = getValidasiQueueSheet_();
  const lastRow = qSheet.getLastRow();
  if (lastRow < 2) return;

  const values = qSheet.getRange(2, 1, lastRow - 1, VALIDASI_QUEUE_HEADERS.length).getValues();
  const byFacility = {};
  values.forEach(function (row, idx) {
    const facilityId = String(row[4] || '').trim();
    if (!byFacility[facilityId]) byFacility[facilityId] = [];
    byFacility[facilityId].push(idx);
  });

  const processedRowNumbers = [];
  const cache = CacheService.getScriptCache();

  Object.keys(byFacility).forEach(function (facilityId) {
    const indices = byFacility[facilityId];
    if (!facilityId) {
      indices.forEach(function (i) {
        catatLogSistem_('Worker Validasi', 'Baris antrian tanpa FacilityId dibuang: ' + JSON.stringify(values[i]));
        processedRowNumbers.push(i + 2);
      });
      return;
    }

    // Guard "1 worker per facility" -- lihat penjelasan lengkap di SubmitQueue.gs
    const workerGuardKey = 'validasiQueueWorkerLock_' + facilityId;
    if (cache.get(workerGuardKey)) {
      return; // baris DIBIARKAN, coba lagi siklus berikutnya
    }
    try { cache.put(workerGuardKey, '1', 55); } catch (e) { /* abaikan */ }

    const facLock = acquireFacilityLock_(facilityId, 10000);
    if (!facLock) {
      try { cache.remove(workerGuardKey); } catch (e) { /* abaikan */ }
      return;
    }

    try {
      const sortedIndices = indices.slice().sort(function (a, b) {
        return new Date(values[a][0]) - new Date(values[b][0]);
      });

      sortedIndices.forEach(function (i) {
        const row = values[i];
        const id = row[1];
        const namaValidator = String(row[2] || '').trim();
        const qtyValidasi = row[3];

        try {
          const result = processSingleQueuedValidasi_(id, namaValidator, qtyValidasi, facilityId, waktuBarisAntrean_(row[5], row[0]));
          if (!result.success) {
            catatLogSistem_('Worker Validasi', 'Skip Id=' + id + ' (' + namaValidator + '): ' + result.message);
          }
        } catch (e) {
          catatLogSistem_('Worker Validasi', 'ERROR Id=' + id + ' (' + namaValidator + '): ' + e.message);
        }
        processedRowNumbers.push(i + 2);
      });
    } finally {
      facLock.release();
      try { cache.remove(workerGuardKey); } catch (e) { /* abaikan */ }
    }
  });

  // v8.31.1 (FIX bug konkurensi @ 15 akun) -- lihat catatan lengkap di Utils.gs / SubmitQueue.gs
  deleteProcessedQueueRows_(qSheet, processedRowNumbers);
}

/**
 * Body PERSIS dari submitValidasi() versi lama (v8.29.x) -- lihat riwayat perubahan di
 * Validasi.gs. Dipanggil dari dalam processValidasiQueue_() yang SUDAH memegang facLock
 * facility terkait -- TIDAK perlu lock lagi di sini.
 */
function processSingleQueuedValidasi_(id, namaValidator, qtyValidasi, facilityId, waktuHitung) {
  const aktual = Number(qtyValidasi);
  const sheet = getRiwayatSheet_(namaValidator);
  const rowIndex = findRiwayatRow_(sheet, id);
  if (rowIndex === -1) {
    return { success: false, message: 'Item tidak ditemukan (kemungkinan sudah divalidasi orang lain).' };
  }

  const rowValues = sheet.getRange(rowIndex, 1, 1, RIWAYAT_HEADERS.length).getValues()[0];
  if (rowValues[11] !== 'Pending') {
    return { success: false, message: 'Sudah divalidasi sebelumnya (duplikat submit).' };
  }

  const qtySystem = Number(rowValues[6]);
  const hasilFinal = aktual === qtySystem ? 'HIT' : 'DISCREPANCY';
  const waktu = waktuHitung || new Date();
  const newRow = rowValues.slice();
  sheet.getRange(rowIndex, 12, 1, 4).setValues([['Selesai', namaValidator, aktual, hasilFinal]]);
  sheet.getRange(rowIndex, 14).setNumberFormat('0');
  sheet.getRange(rowIndex, 17).setValue(waktu);
  newRow[11] = 'Selesai'; newRow[12] = namaValidator; newRow[13] = aktual;
  newRow[14] = hasilFinal; newRow[16] = waktu;

  if (hasilFinal === 'HIT') {
    const catatanOtomatis = 'Otomatis: hasil validasi ulang sesuai stok system, tidak ada discrepancy nyata (kesalahan hitung awal petugas).';
    sheet.getRange(rowIndex, 18, 1, 4).setValues([[
      'Salah Hitung', 'Selesai', catatanOtomatis, namaValidator
    ]]);
    sheet.getRange(rowIndex, 22).setValue(waktu);
    sheet.getRange(rowIndex, 25).setValue(waktu);
    newRow[17] = 'Salah Hitung'; newRow[18] = 'Selesai'; newRow[19] = catatanOtomatis;
    newRow[20] = namaValidator; newRow[21] = waktu; newRow[24] = waktu;
  } else {
    const initialPic = assignInitialPIC_(facilityId);
    sheet.getRange(rowIndex, 19).setValue('Open');
    sheet.getRange(rowIndex, 24).setValue(initialPic);
    sheet.getRange(rowIndex, 25).setValue(waktu);
    newRow[18] = 'Open'; newRow[23] = initialPic; newRow[24] = waktu;
  }

  try { applyRingkasanDelta_(rowDateTag_(rowValues), rowValues, newRow, namaValidator); } catch (e) { /* lihat rebuildRingkasanHarian_ */ }
  try {
    incrementCounterKey_('pending_total', -1, namaValidator);
    if (rowValues[22]) incrementCounterKey_('pending_validator:' + rowValues[22], -1, namaValidator);
    if (hasilFinal === 'DISCREPANCY') {
      adjustPlusMinusTracker_(rowValues[9], rowValues[3], rowValues[4], rowValues[2], true, namaValidator);
    }
  } catch (e) { /* lihat rebuildAntrianAktif_ */ }

  return { success: true, hasilFinal: hasilFinal };
}

// =========================================================================
// SETUP (jalankan MANUAL SEKALI dari Apps Script editor setelah deploy patch ini)
// =========================================================================

function installValidasiQueueTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processValidasiQueue_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processValidasiQueue_')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('Trigger terpasang: processValidasiQueue_ tiap 1 menit.');
}
