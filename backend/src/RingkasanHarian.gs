/**
 * RingkasanHarian.gs
 * Fase 22 - performa: sheet precompute "Ringkasan_Harian" (1 baris = 1
 * tanggal cycle count)
 * Yang di-update INCREMENTAL (delta) tiap kali ada transaksi baru, supaya
 * Dashboard tidak
 * perlu scan penuh sheet Riwayat setiap kali dibuka.
 *
 * v8.26.0 (Facility Management - Strict Isolation):
 * - SELURUH fungsi di file ini SEKARANG menerima parameter opsional 'username'.
 * - Kalau disediakan, Ringkasan_Harian dibaca/ditulis ke spreadsheet facility
 * user tersebut.
 * - Fungsi 'getRingkasanHarianSheet_()' TIDAK LAGI didefinisikan di sini --
 * memakai versi
 * facility-aware dari SheetHelpers.gs yang menerima parameter username.
 *
 * SUMBER KEBENARAN (single source of truth) untuk "kontribusi 1 baris
 * Riwayat ke summary"
 * ada di computeRowContribution_() -- fungsi ini MENIRU PERSIS logika yang
 * sudah ada di
 * Dashboard.gs (computeSummaryFromData_ & computeErrorAnalysisFromData_),
 * supaya angka di
 * Ringkasan_Harian selalu konsisten dengan hasil scan langsung. Kalau logika
 * penghitungan di
 * Dashboard.gs pernah diubah, computeRowContribution_() WAJIB diubah juga di
 * waktu yang sama.
 *
 * MEKANISME DELTA: karena status 1 baris Riwayat bisa berubah SETELAH baris
 * itu dibuat
 * (Pending -> divalidasi -> diinvestigasi -> Selesai), kontribusinya ke
 * Ringkasan_Harian juga
 * berubah. Solusinya: di setiap titik tulis (submitCount, submitValidasi,
 * updateTaskStatus,
 * closePlusMinusPair), kita hitung kontribusi versi LAMA (sebelum ditulis)
 * dan versi BARU
 * (sesudah ditulis), lalu terapkan SELISIHnya (delta) ke baris tanggal yang
 * bersangkutan di
 * Ringkasan_Harian -- BUKAN scan ulang Riwayat. Baris baru (belum pernah ada
 * sebelumnya, dari
 * submitCount) dianggap "versi lama"-nya kosong (semua kontribusi 0).
 *
 * SAFETY NET: kalau delta ini pernah drift/tidak sinkron (mis. ada bug, atau
 * sheet Riwayat
 * diedit manual lewat Google Sheets tanpa lewat aplikasi), tersedia
 * rebuildRingkasanHarian_()
 * yang membangun ULANG seluruh Ringkasan_Harian dari nol berdasarkan isi
 * Riwayat saat ini --
 * bisa dipanggil kapan saja lewat menu "Cycle Count > Rebuild Ringkasan
 * Harian" untuk
 * mengembalikan konsistensi tanpa perlu debug delta satu-satu.
 */

// ---------- Sheet & row cache (pola sama seperti findRiwayatRow_ di
// SheetHelpers.gs) ----------

// PENTING v8.26.0: getRingkasanHarianSheet_() TIDAK didefinisikan di sini
// lagi.
// Memakai versi facility-aware dari SheetHelpers.gs: getRingkasanHarianSheet_
// (username)

// v8.26.2 (FIX cross-facility cache corruption -- PALING KRITIKAL dari 2
// cache yang diperbaiki
// di refactor ini): key di sini adalah TANGGAL ('2026-08-17'), dan SETIAP
// facility PASTI punya
// baris untuk tanggal yang sama tiap hari. Row-index cache SEBELUM v8.26.2
// cuma di-key oleh
// 'tanggal' (tanpa konteks Facility) -- kalau 2 orang di 2 facility berbeda
// submitCount() di
// hari yang sama (skenario yang PASTI terjadi tiap hari kerja), salah satu
// bisa dapat rowIndex
// baris MILIK FACILITY LAIN dari cache, dan applyRingkasanDelta_ menulis
// delta ke baris yang
// salah -- KPI harian salah satu facility jadi ikut kecampur/rusak. Fix:
// cache key sekarang
// diprefix facilityId (via getUserFacility(username), sendirinya sudah
// di-cache 300 detik).

function ringkasanCacheFacilityPrefix_(username) {
  if (!username) return 'legacy';
  const facInfo = getUserFacility(username);
  return facInfo ? facInfo.id : 'legacy';
}

function ringkasanHarianRowCachePut_(tanggal, rowIndex, username) {
  try {
    const prefix = ringkasanCacheFacilityPrefix_(username);
    CacheService.getScriptCache().put('ringkasanRow_' + prefix + '_' +
      tanggal, String(rowIndex), 21600);
  } catch (e) { /* cache opsional, aman diabaikan */ }
}

function ringkasanHarianRowCacheGet_(tanggal, username) {
  try {
    const prefix = ringkasanCacheFacilityPrefix_(username);
    const v = CacheService.getScriptCache().get('ringkasanRow_' + prefix +
      '_' + tanggal);
    return v ? Number(v) : -1;
  } catch (e) { return -1; }
}

function findRingkasanHarianRow_(sheet, tanggal, username) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const cachedRow = ringkasanHarianRowCacheGet_(tanggal, username);
  if (cachedRow >= 2 && cachedRow <= lastRow) {
    if (String(sheet.getRange(cachedRow, 1).getValue()) === String(tanggal)) {
      return cachedRow;
    }
  }

  // fallback: scan kolom A saja (bukan seluruh sheet) -- sheet ini kecil
  // (maks ~1 baris/hari
  // dalam ARCHIVE_AGE_MONTHS terakhir), jadi tetap murah meskipun cache
  // meleset.
  const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]) === String(tanggal)) {
      const rowIndex = i + 2;
      ringkasanHarianRowCachePut_(tanggal, rowIndex, username);
      return rowIndex;
    }
  }
  return -1;
}

/**
 * v8.26.0: Menerima parameter opsional 'username' untuk baca/tulis ke
 * facility yang benar.
 */
function getOrCreateRingkasanHarianRow_(sheet, tanggal, username) {
  let rowIndex = findRingkasanHarianRow_(sheet, tanggal, username);
  if (rowIndex === -1) {
    rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 1, 1, RINGKASAN_HARIAN_HEADERS.length).setValues(
      [[tanggal, 0, 0, 0, '{}', '{}', '{}', '{}']]
    );
    ringkasanHarianRowCachePut_(tanggal, rowIndex, username);
  } else {
    ringkasanHarianRowCachePut_(tanggal, rowIndex, username);
  }
  const values = sheet.getRange(rowIndex, 1, 1, RINGKASAN_HARIAN_HEADERS.length).getValues()[0];
  return { rowIndex: rowIndex, values: values };
}

// ---------- Kontribusi 1 baris Riwayat (sumber kebenaran tunggal) ----------

/**
 * MENIRU PERSIS logika di computeSummaryFromData_ &
 * computeErrorAnalysisFromData_ (Dashboard.gs).
 * row = 1 baris Riwayat (array sesuai urutan RIWAYAT_HEADERS).
 */
function computeRowContribution_(row) {
  const namaPetugas = row[7];
  const hasilAwal = row[10], statusValidasi = row[11], namaValidator = row[12], hasilFinal = row[14];
  const kategoriSelisih = String(row[17] || '').trim();
  const statusTask = row[18];
  const effectiveHasil = statusValidasi === 'Pending' ? 'PENDING' : (statusValidasi === 'Tidak Perlu' ? hasilAwal : hasilFinal);
  const kpi = {
    total: 1,
    hit: effectiveHasil === 'HIT' ? 1 : 0,
    disc: effectiveHasil === 'DISCREPANCY' ? 1 : 0
  };
  let petugas = null;
  if (statusValidasi === 'Tidak Perlu') {
    petugas = { nama: namaPetugas, total: 1, hit: 1 };
  } else if (statusValidasi === 'Selesai') {
    petugas = { nama: namaPetugas, total: 1, hit: (hasilAwal === hasilFinal) ? 1 : 0 };
  }
  // Fase 23: petugasRaw = SEMUA item yang dihitung petugas ini, TANPA PEDULI
  // status validasi
  // (Pending ikut dihitung) -- dipakai Home "Total Cycle Hari Ini" /
  // "Selesai Hari Ini". Beda
  // dari 'petugas' di atas (Leaderboard Dashboard) yang CUMA menghitung item
  // yang sudah final.
  const petugasRaw = namaPetugas
    ? { nama: namaPetugas, rawTotal: 1, rawSelesai: statusValidasi !== 'Pending' ? 1 : 0 }
    : null;
  let validator = null;
  if (namaValidator && statusTask === 'Selesai') {
    const validatorSalah = (hasilFinal === 'DISCREPANCY' && kategoriSelisih === 'Salah Hitung');
    validator = { nama: namaValidator, total: 1, hit: validatorSalah ? 0 : 1 };
  }

  return { kpi: kpi, petugas: petugas, petugasRaw: petugasRaw, validator: validator, kategori: kategoriSelisih || null };
}

// ---------- Penerapan delta ----------

function applyKeyedDelta_(map, oldEntry, newEntry) {
  if (oldEntry) {
    const k = oldEntry.nama;
    if (map[k]) {
      map[k].total -= oldEntry.total;
      map[k].hit -= oldEntry.hit;
      if (map[k].total <= 0) delete map[k];
    }
  }
  if (newEntry) {
    const k = newEntry.nama;
    if (!map[k]) map[k] = { total: 0, hit: 0 };
    map[k].total += newEntry.total;
    map[k].hit += newEntry.hit;
  }
}

/**
 * Fase 23: delta khusus untuk Petugas_Raw_JSON (field rawTotal/rawSelesai,
 * BEDA nama field
 * dari applyKeyedDelta_ yang pakai total/hit -- dibikin fungsi terpisah
 * supaya tidak ketukar).
 */
function applyKeyedRawDelta_(map, oldEntry, newEntry) {
  if (oldEntry) {
    const k = oldEntry.nama;
    if (map[k]) {
      map[k].rawTotal -= oldEntry.rawTotal;
      map[k].rawSelesai -= oldEntry.rawSelesai;
      if (map[k].rawTotal <= 0) delete map[k];
    }
  }
  if (newEntry) {
    const k = newEntry.nama;
    if (!map[k]) map[k] = { rawTotal: 0, rawSelesai: 0 };
    map[k].rawTotal += newEntry.rawTotal;
    map[k].rawSelesai += newEntry.rawSelesai;
  }
}

function applyKategoriDelta_(map, oldKategori, newKategori) {
  if (oldKategori) {
    map[oldKategori] = (map[oldKategori] || 0) - 1;
    if (map[oldKategori] <= 0) delete map[oldKategori];
  }
  if (newKategori) {
    map[newKategori] = (map[newKategori] || 0) + 1;
  }
}

/**
 * Terapkan delta kontribusi 1 baris Riwayat (dari 'oldRow' ke 'newRow') ke
 * baris tanggal
 * yang bersangkutan di Ringkasan_Harian. 'oldRow' = null kalau ini baris
 * BARU (belum pernah
 * dihitung sebelumnya, mis. dari submitCount). 'newRow' = null kalau baris
 * itu dihapus
 * (saat ini tidak pernah terjadi di alur aplikasi, disediakan untuk
 * kelengkapan/masa depan).
 * 'tanggal' HARUS dari rowDateTag_(oldRow || newRow) supaya konsisten dengan
 * konvensi
 * tanggal yang dipakai Dashboard.
 *
 * v8.26.0: Menerima parameter opsional 'username' untuk tulis ke Facility
 * yang benar.
 */
function applyRingkasanDelta_(tanggal, oldRow, newRow, username) {
  if (!tanggal) return;
  const zero = { kpi: { total: 0, hit: 0, disc: 0 }, petugas: null, petugasRaw: null, validator: null, kategori: null };
  const oldC = oldRow ? computeRowContribution_(oldRow) : zero;
  const newC = newRow ? computeRowContribution_(newRow) : zero;
  // Tidak ada perubahan sama sekali -- skip supaya tidak menulis sheet tanpa
  // perlu.
  const kpiSame = oldC.kpi.total === newC.kpi.total && oldC.kpi.hit === newC.kpi.hit && oldC.kpi.disc === newC.kpi.disc;
  const petugasSame = JSON.stringify(oldC.petugas) === JSON.stringify(newC.petugas);
  const petugasRawSame = JSON.stringify(oldC.petugasRaw) === JSON.stringify(newC.petugasRaw);
  const validatorSame = JSON.stringify(oldC.validator) === JSON.stringify(newC.validator);
  const kategoriSame = oldC.kategori === newC.kategori;
  if (kpiSame && petugasSame && petugasRawSame && validatorSame && kategoriSame) return;

  // v8.26.0: Baca/tulis ke Ringkasan_Harian milik facility user
  const sheet = getRingkasanHarianSheet_(username);
  const rowInfo = getOrCreateRingkasanHarianRow_(sheet, tanggal, username);
  const vals = rowInfo.values;
  const kpiTotal = (Number(vals[1]) || 0) + (newC.kpi.total - oldC.kpi.total);
  const kpiHit = (Number(vals[2]) || 0) + (newC.kpi.hit - oldC.kpi.hit);
  const kpiDisc = (Number(vals[3]) || 0) + (newC.kpi.disc - oldC.kpi.disc);
  const petugasMap = parseJsonSafe_(vals[4], {});
  applyKeyedDelta_(petugasMap, oldC.petugas, newC.petugas);
  const validatorMap = parseJsonSafe_(vals[5], {});
  applyKeyedDelta_(validatorMap, oldC.validator, newC.validator);
  const kategoriMap = parseJsonSafe_(vals[6], {});
  applyKategoriDelta_(kategoriMap, oldC.kategori, newC.kategori);
  const petugasRawMap = parseJsonSafe_(vals[7], {});
  applyKeyedRawDelta_(petugasRawMap, oldC.petugasRaw, newC.petugasRaw);
  sheet.getRange(rowInfo.rowIndex, 2, 1, 7).setValues([[
    kpiTotal, kpiHit, kpiDisc, JSON.stringify(petugasMap), JSON.stringify(validatorMap), JSON.stringify(kategoriMap), JSON.stringify(petugasRawMap)
  ]]);
}

// ---------- Rebuild total (safety net / reconciliation) ----------

/**
 * v8.26.0: Menerima parameter opsional 'username'.
 * Membaca Riwayat dari facility user tersebut, lalu membangun ulang
 * Ringkasan_Harian.
 */
function rebuildRingkasanHarian_(username) {
  // v8.29.0: lock PER-FACILITY (bukan global) -- HARUS pakai key facility
  // yang SAMA dengan
  // submitCount/submitValidasi/dst (lihat acquireFacilityLock_ di
  // FacilityManagement.gs)
  // supaya rebuild tetap mutually-exclusive terhadap submit di facility yang
  // sama, tapi
  // rebuild facility lain tidak ikut ke-block.
  const facInfoForLock = getUserFacility(username);
  const facLock = acquireFacilityLock_(facInfoForLock ? facInfoForLock.id : null, 20000);
  if (!facLock) throw new Error('Sistem sedang sibuk, coba lagi nanti.');
  try {
    // v8.26.0: Baca Riwayat dari facility user
    const riwayatSheet = getRiwayatSheet_(username);
    const lastRow = riwayatSheet.getLastRow();
    const acc = {};
    if (lastRow >= 2) {
      const data = riwayatSheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
      data.forEach(function (row) {
        const tgl = rowDateTag_(row);
        if (!tgl) return;
        if (!acc[tgl]) acc[tgl] = { kpiTotal: 0, kpiHit: 0, kpiDisc: 0, petugas: {}, validator: {}, kategori: {}, petugasRaw: {} };
        const c = computeRowContribution_(row);
        const bucket = acc[tgl];
        bucket.kpiTotal += c.kpi.total;
        bucket.kpiHit += c.kpi.hit;
        bucket.kpiDisc += c.kpi.disc;
        if (c.petugas) {
          if (!bucket.petugas[c.petugas.nama]) bucket.petugas[c.petugas.nama] = { total: 0, hit: 0 };
          bucket.petugas[c.petugas.nama].total += c.petugas.total;
          bucket.petugas[c.petugas.nama].hit += c.petugas.hit;
        }
        if (c.petugasRaw) {
          if (!bucket.petugasRaw[c.petugasRaw.nama]) bucket.petugasRaw[c.petugasRaw.nama] = { rawTotal: 0, rawSelesai: 0 };
          bucket.petugasRaw[c.petugasRaw.nama].rawTotal += c.petugasRaw.rawTotal;
          bucket.petugasRaw[c.petugasRaw.nama].rawSelesai += c.petugasRaw.rawSelesai;
        }
        if (c.validator) {
          if (!bucket.validator[c.validator.nama]) bucket.validator[c.validator.nama] = { total: 0, hit: 0 };
          bucket.validator[c.validator.nama].total += c.validator.total;
          bucket.validator[c.validator.nama].hit += c.validator.hit;
        }
        if (c.kategori) {
          bucket.kategori[c.kategori] = (bucket.kategori[c.kategori] || 0) + 1;
        }
      });
    }

    // v8.26.0: Tulis ke Ringkasan_Harian milik facility user
    const sheet = getRingkasanHarianSheet_(username);
    const outRows = Object.keys(acc).sort().map(function (tgl) {
      const a = acc[tgl];
      return [tgl, a.kpiTotal, a.kpiHit, a.kpiDisc, JSON.stringify(a.petugas), JSON.stringify(a.validator), JSON.stringify(a.kategori), JSON.stringify(a.petugasRaw)];
    });
    const oldLastRow = sheet.getLastRow();
    if (oldLastRow > 1) sheet.getRange(2, 1, oldLastRow - 1, RINGKASAN_HARIAN_HEADERS.length).clearContent();
    if (outRows.length > 0) {
      sheet.getRange(2, 1, outRows.length, RINGKASAN_HARIAN_HEADERS.length).setValues(outRows);
    }
    return { success: true, hari: outRows.length };
  } finally {
    facLock.release();
  }
}

/**
 * Menu "Cycle Count > Rebuild Ringkasan Harian" -- aman dipanggil kapan
 * saja, cuma menghitung
 * ulang dari Riwayat (tidak mengubah Riwayat sama sekali).
 */
function rebuildRingkasanHarianMenu_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Ini akan menghitung ULANG seluruh sheet "' + RINGKASAN_HARIAN_SHEET_NAME +
    '" dari isi Riwayat saat ini (dipakai kalau angka Dashboard terasa tidak sinkron). Proses ini TIDAK mengubah data Riwayat sama sekali.\n\nLanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  // v8.26.0: Menu dipanggil dari spreadsheet aktif, pakai null/fallback.
  const hasil = rebuildRingkasanHarian_(null);
  ui.alert('Selesai. Ringkasan_Harian dibangun ulang untuk ' + hasil.hari + ' tanggal.');
}
