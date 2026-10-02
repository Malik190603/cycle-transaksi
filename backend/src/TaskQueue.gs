/**
 * TaskQueue.gs (v8.31.0 - Asynchronous Queue Processing untuk updateTaskStatus)
 * 
 * Pola sama dengan SubmitQueue.gs / ValidasiQueue.gs. updateTaskStatus() (TaskInvestigasi.gs)
 * SEKARANG memvalidasi bukti WMS secara SINKRON (pakai lokasi/article/selisih yang dikirim
 * client, murni CPU tanpa baca sheet -- lihat validateBuktiTransaksi_), lalu cuma menitipkan
 * payload ke Queue_Task_Investigasi (spreadsheet BOUND, bukan openById).
 * Worker (processTaskQueue_(), trigger 1 menit) yang benar-benar buka spreadsheet facility & menulis
 * -- SEKALIGUS validasi ULANG bukti pakai data SEGAR dari sheet (defense-in-depth, client
 * tidak dipercaya buta untuk keputusan final penutupan task).
 * 
 * PENTING -- WAJIB dijalankan sekali manual setelah deploy:
 * * Buka Apps Script editor -> pilih fungsi "installTaskQueueTrigger_" dari dropdown ->
 * * Klik Run.
 */

// =========================================================================
// SISI PENULIS (dipanggil dari updateTaskStatus())
// =========================================================================

function queueTaskUpdatePayload_(id, namaUser, newStatusTrim, catatanTrim, kategoriTrim, picTrim, buktiRows, facilityId) {
  const sheet = getTaskQueueSheet_();
  appendRowLocked_(sheet, [new Date(), id, namaUser, newStatusTrim, catatanTrim, kategoriTrim, picTrim, JSON.stringify(buktiRows || []), facilityId]);
}


// =========================================================================
// SISI PEMROSES (dipanggil dari trigger time-driven, BUKAN dari updateTaskStatus)
// =========================================================================

function processTaskQueue_() {
  const qSheet = getTaskQueueSheet_();
  const lastRow = qSheet.getLastRow();
  if (lastRow < 2) return;

  const values = qSheet.getRange(2, 1, lastRow - 1, TASK_QUEUE_HEADERS.length).getValues();
  const byFacility = {};
  values.forEach(function (row, idx) {
    const facilityId = String(row[8] || '').trim();
    if (!byFacility[facilityId]) byFacility[facilityId] = [];
    byFacility[facilityId].push(idx);
  });

  const processedRowNumbers = [];
  const cache = CacheService.getScriptCache();

  Object.keys(byFacility).forEach(function (facilityId) {
    const indices = byFacility[facilityId];
    if (!facilityId) {
      indices.forEach(function (i) {
        catatLogSistem_('Worker Task', 'Baris antrian tanpa FacilityId dibuang: ' + JSON.stringify(values[i]));
        processedRowNumbers.push(i + 2);
      });
      return;
    }

    const workerGuardKey = 'taskQueueWorkerLock_' + facilityId;
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
        const namaUser = String(row[2] || '').trim();
        const newStatusTrim = String(row[3] || '').trim();
        const catatanTrim = String(row[4] || '').trim();
        const kategoriTrim = String(row[5] || '').trim();
        const picTrim = String(row[6] || '').trim();
        let buktiRows = [];
        try { buktiRows = JSON.parse(row[7] || '[]'); } catch (e) { /* abaikan -- treat sbg kosong */ }
        try {
          const result = processSingleQueuedTaskUpdate_(id, namaUser, newStatusTrim, catatanTrim, kategoriTrim, picTrim, buktiRows);
          if (!result.success) {
            catatLogSistem_('Worker Task', 'Skip Id=' + id + ' (' + namaUser + '): ' + result.message);
          }
        } catch (e) {
          catatLogSistem_('Worker Task', 'ERROR Id=' + id + ' (' + namaUser + '): ' + e.message);
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
 * Body PERSIS dari updateTaskStatus() versi lama (v8.29.x, bagian PENULISAN saja -- validasi
 * ringan/pure sudah dilakukan sinkron di updateTaskStatus() sebelum masuk antrian). SATU
 * TAMBAHAN: bukti WMS divalidasi ULANG di sini pakai lokasi/article/selisih SEGAR dari sheet
 * (bukan dari client) -- defense-in-depth, supaya kalau data client sempat basi (jarang
 * terjadi), Worker tetap menolak & tidak salah menutup task.
 */
function processSingleQueuedTaskUpdate_(id, namaUser, newStatusTrim, catatanTrim, kategoriTrim, picTrim, buktiRows) {
  const targetIdx = TASK_STATUS_FLOW.indexOf(newStatusTrim);
  if (targetIdx <= 0) return { success: false, message: 'Status tujuan tidak valid (data antrian korup).' };
  const isFinal = newStatusTrim === 'Selesai';

  const sheet = getRiwayatSheet_(namaUser);
  const rowIndex = findRiwayatRow_(sheet, id);
  if (rowIndex === -1) return { success: false, message: 'Task tidak ditemukan, kemungkinan sudah diproses.' };

  const rowValues = sheet.getRange(rowIndex, 1, 1, RIWAYAT_HEADERS.length).getValues()[0];
  const currentStatus = String(rowValues[18]);
  const currentIdx = TASK_STATUS_FLOW.indexOf(currentStatus);
  if (currentIdx === -1 || currentStatus === 'Selesai') {
    return { success: false, message: 'Task ini sudah tidak aktif (mungkin sudah Selesai oleh proses lain).' };
  }
  if (targetIdx <= currentIdx) {
    return { success: false, message: 'Status tidak boleh mundur/sama (sekarang: ' + currentStatus + ') -- kemungkinan sudah diupdate proses lain.' };
  }

  let lokasiTask = '', articleTask = '';
  if (isFinal && KATEGORI_TANPA_BUKTI_.indexOf(kategoriTrim) === -1) {
    lokasiTask = String(rowValues[2] || '').trim();
    articleTask = String(rowValues[3] || '').trim();
    const selisihAbs = Math.abs(Number(rowValues[9]) || 0);
    // v8.31.0: validasi ULANG pakai data SEGAR (bukan echo dari client) -- defense-in-depth.
    const cek = validateBuktiTransaksi_(buktiRows, articleTask, selisihAbs, kategoriTrim, lokasiTask);
    if (!cek.valid) return { success: false, message: 'Bukti tidak valid saat diproses ulang (data mungkin sudah berubah): ' + cek.message };
  }

  const waktu = new Date();
  const prevCatatan = String(rowValues[19] || '');
  const stamp = Utilities.formatDate(waktu, 'Asia/Jakarta', 'dd/MM HH:mm');
  let newCatatan = (prevCatatan ? prevCatatan + '\n' : '') + '[' + stamp + ' - ' + newStatusTrim + '] ' + catatanTrim;
  if (isFinal && KATEGORI_TANPA_BUKTI_.indexOf(kategoriTrim) === -1 && buktiRows && buktiRows.length) {
    newCatatan += '\n[' + stamp + '] Bukti transaksi WMS terlampir (' + buktiRows.length + ' baris) -- lihat sheet Log_Bukti_Investigasi.';
  }

  const newRow = rowValues.slice();

  sheet.getRange(rowIndex, 19).setValue(newStatusTrim);
  sheet.getRange(rowIndex, 20).setValue(newCatatan);
  sheet.getRange(rowIndex, 25).setValue(waktu);
  newRow[18] = newStatusTrim; newRow[19] = newCatatan; newRow[24] = waktu;

  if (kategoriTrim && TASK_KATEGORI_LIST.indexOf(kategoriTrim) !== -1) {
    sheet.getRange(rowIndex, 18).setValue(kategoriTrim);
    newRow[17] = kategoriTrim;
  }
  if (picTrim && getUserRole(picTrim)) {
    sheet.getRange(rowIndex, 24).setValue(picTrim);
    newRow[23] = picTrim;
  }

  if (isFinal) {
    sheet.getRange(rowIndex, 21).setValue(namaUser);
    sheet.getRange(rowIndex, 22).setValue(waktu);
    newRow[20] = namaUser; newRow[21] = waktu;
  }

  if (isFinal && KATEGORI_TANPA_BUKTI_.indexOf(kategoriTrim) === -1 && buktiRows && buktiRows.length) {
    const buktiSheet = getLogBuktiSheet_(namaUser);
    const buktiStartRow = buktiSheet.getLastRow() + 1;
    const buktiOut = buktiRows.map(function (r) {
      return [waktu, id, lokasiTask, articleTask, namaUser].concat(r.slice(0, 15));
    });
    buktiSheet.getRange(buktiStartRow, 1, buktiOut.length, LOG_BUKTI_HEADERS.length).setValues(buktiOut);
  }

  try { applyRingkasanDelta_(rowDateTag_(rowValues), rowValues, newRow, namaUser); } catch (e) { /* lihat rebuildRingkasanHarian_ */ }

  if (isFinal && rowValues[14] === 'DISCREPANCY') {
    try { adjustPlusMinusTracker_(rowValues[9], rowValues[3], rowValues[4], rowValues[2], false, namaUser); } catch (e) { /* lihat rebuildAntrianAktif_ */ }
  }

  return { success: true };
}

// =========================================================================
// SETUP (jalankan MANUAL SEKALI dari Apps Script editor setelah deploy patch ini)
// =========================================================================

function installTaskQueueTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processTaskQueue_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processTaskQueue_')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('Trigger terpasang: processTaskQueue_ tiap 1 menit.');
}
