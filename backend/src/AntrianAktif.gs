/**
 * AntrianAktif.gs
 * Fase 23 - performa: counter LIVE (bukan agregat historis per-tanggal seperti Ringkasan_Harian)
 * untuk 2 kebutuhan Home yang sifatnya "kondisi SEKARANG", bukan "kondisi hari tertentu":
 *   1. Antrian belum divalidasi (belumValidasi) -- 1 item bisa dihitung berhari-hari lalu dan
 *      baru divalidasi sekarang, jadi tidak bisa dikelompokkan per tanggal.
 *   2. Summary Plus/Minus aktif -- SKU yang statusnya masih aktif (Status_Task belum Selesai),
 *      sama sifatnya, tidak terikat 1 tanggal tertentu.
 *
 * v8.26.0 (Facility Management - Strict Isolation):
 * SELURUH fungsi di file ini SEKARANG menerima parameter opsional `username`.
 * Kalau disediakan, counter & tracker dibaca/ditulis ke spreadsheet facility user tersebut.
 * Fungsi `getAntrianAktifSheet_()` TIDAK LAGI didefinisikan di sini -- memakai versi
 * facility-aware dari SheetHelpers.gs yang menerima parameter username.
 *
 * Di-update +1/-1 di titik yang sama seperti Ringkasan_Harian (appendRiwayatCycle_,
 * submitValidasi, updateTaskStatus, closePlusMinusPair) -- SELALU dibungkus try/catch supaya
 * kegagalan di sini TIDAK PERNAH menggagalkan transaksi utama. Ada rebuildAntrianAktif_()
 * sebagai safety net, sama seperti rebuildRingkasanHarian_().
 */
// ---------- Sheet & row cache (pola sama seperti RingkasanHarian.gs) ----------
//
// PENTING v8.26.0: getAntrianAktifSheet_() TIDAK didefinisikan di sini lagi.
// Memakai versi facility-aware dari SheetHelpers.gs: getAntrianAktifSheet_(username)
//
// v8.26.2 (FIX cross-facility cache corruption): key counter (mis. 'pending_total',
// 'plusminus_plus') PERSIS SAMA di semua facility -- Antrian_Aktif facility A dan facility
// B SAMA-SAMA punya baris dengan key 'pending_total', biasanya di nomor baris yang mirip/sama
// juga. Row-index cache SEBELUM v8.26.2 cuma di-key oleh `key` (tanpa konteks facility), jadi
// kalau 2 orang di 2 facility berbeda sama-sama trigger cache put untuk key yang sama, salah
// satu bisa dapat rowIndex baris MILIK FACILITY LAIN -- baca/tulis ke baris yang salah.
// Fix: cache key sekarang diprefix facilityId (via getUserFacility(username), yang sendirinya
// sudah di-cache 300 detik jadi lookup tambahan ini murah).
function antrianCacheFacilityPrefix_(username) {
  if (!username) return 'legacy';
  const facInfo = getUserFacility(username);
  return facInfo ? facInfo.id : 'legacy';
}
function antrianRowCachePut_(key, rowIndex, username) {
  try {
    const prefix = antrianCacheFacilityPrefix_(username);
    CacheService.getScriptCache().put('antrianRow_' + prefix + '_' + key, String(rowIndex), 21600);
  } catch (e) { /* abaikan */ }
}
function antrianRowCacheGet_(key, username) {
  try {
    const prefix = antrianCacheFacilityPrefix_(username);
    const v = CacheService.getScriptCache().get('antrianRow_' + prefix + '_' + key);
    return v ? Number(v) : -1;
  } catch (e) { return -1; }
}
function findAntrianRow_(sheet, key, username) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const cachedRow = antrianRowCacheGet_(key, username);
  if (cachedRow >= 2 && cachedRow <= lastRow) {
    if (String(sheet.getRange(cachedRow, 1).getValue()) === String(key)) return cachedRow;
  }
  const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]) === String(key)) {
      const rowIndex = i + 2;
      antrianRowCachePut_(key, rowIndex, username);
      return rowIndex;
    }
  }
  return -1;
}

/**
 * v8.26.0: Menerima parameter opsional `username` untuk baca dari facility yang benar.
 */
function readAntrianValue_(key, fallback, username) {
  const sheet = getAntrianAktifSheet_(username);
  const rowIndex = findAntrianRow_(sheet, key, username);
  if (rowIndex === -1) return fallback;
  return parseJsonSafe_(sheet.getRange(rowIndex, 2).getValue(), fallback);
}

/**
 * v8.26.0: Menerima parameter opsional `username` untuk tulis ke facility yang benar.
 */
function writeAntrianValue_(key, value, username) {
  const sheet = getAntrianAktifSheet_(username);
  let rowIndex = findAntrianRow_(sheet, key, username);
  if (rowIndex === -1) {
    rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 1, 1, 2).setValues([[key, JSON.stringify(value)]]);
    antrianRowCachePut_(key, rowIndex, username);
  } else {
    sheet.getRange(rowIndex, 2).setValue(JSON.stringify(value));
  }
}

/**
 * Naik/turunkan counter sederhana {"count": N}. delta boleh negatif. Diclamp minimal 0 --
 * seharusnya tidak pernah terjadi kalau delta seimbang, tapi jaga-jaga drift kecil supaya
 * angka yang tampil ke user tidak pernah negatif (aneh dilihat).
 *
 * v8.26.0: Menerima parameter opsional `username` untuk counter per-facility.
 */
function incrementCounterKey_(key, delta, username) {
  const current = readAntrianValue_(key, { count: 0 }, username);
  const next = Math.max(0, (Number(current.count) || 0) + delta);
  writeAntrianValue_(key, { count: next }, username);
}

// ---------- Tracker Plus/Minus aktif per SKU ----------
const PLUSMINUS_KEY_PLUS_ = 'plusminus_plus';
const PLUSMINUS_KEY_MINUS_ = 'plusminus_minus';

/**
 * add=true: 1 task discrepancy BARU jadi aktif (dipanggil dari submitValidasi begitu
 * Hasil_Final=DISCREPANCY dikonfirmasi). add=false: 1 task discrepancy SELESAI/ditutup
 * (dipanggil dari updateTaskStatus/closePlusMinusPair) -- hapus kontribusinya.
 *
 * v8.26.0: Menerima parameter opsional `username` untuk tracker per-facility.
 */
function adjustPlusMinusTracker_(selisih, article, description, lokasi, add, username) {
  const key = Number(selisih) > 0 ? PLUSMINUS_KEY_PLUS_ : PLUSMINUS_KEY_MINUS_;
  const qtyAbs = Math.abs(Number(selisih) || 0);
  if (!qtyAbs || !article) return;
  const map = readAntrianValue_(key, {}, username);
  if (!map[article]) map[article] = { description: description, qty: 0, lokasiCounts: {} };
  const entry = map[article];
  if (add) {
    entry.qty += qtyAbs;
    entry.lokasiCounts[lokasi] = (entry.lokasiCounts[lokasi] || 0) + 1;
  } else {
    entry.qty -= qtyAbs;
    if (entry.lokasiCounts[lokasi]) {
      entry.lokasiCounts[lokasi] -= 1;
      if (entry.lokasiCounts[lokasi] <= 0) delete entry.lokasiCounts[lokasi];
    }
  }
  if (entry.qty <= 0 || Object.keys(entry.lokasiCounts).length === 0) {
    delete map[article];
  }
  writeAntrianValue_(key, map, username);
}

/**
 * Pengganti getPlusMinusSummary() versi lama (yang scan penuh Riwayat) -- baca 2 key kecil
 * ini langsung, TIDAK ada scan sama sekali.
 *
 * v8.26.0: Menerima parameter opsional `username` untuk summary per-facility.
 */
function getPlusMinusSummaryFromAntrian_(username) {
  function toItems(map) {
    const items = Object.keys(map).map(function (article) {
      const e = map[article];
      const lokasiOrder = Object.keys(e.lokasiCounts || {});
      return { article: article, description: e.description, qty: e.qty, lokasi: lokasiOrder.join(', '), jumlahLokasi: lokasiOrder.length };
    }).sort(function (a, b) { return b.qty - a.qty; });
    return { totalSku: items.length, totalQty: items.reduce(function (s, it) { return s + it.qty; }, 0), items: items };
  }
  const plusMap = readAntrianValue_(PLUSMINUS_KEY_PLUS_, {}, username);
  const minusMap = readAntrianValue_(PLUSMINUS_KEY_MINUS_, {}, username);
  return { plus: toItems(plusMap), minus: toItems(minusMap) };
}

// ---------- Rebuild total (safety net / reconciliation) ----------

/**
 * v8.26.0: Menerima parameter opsional `username`.
 * Membaca Riwayat dari facility user tersebut, lalu membangun ulang Antrian_Aktif.
 */
function rebuildAntrianAktif_(username) {
  // v8.29.0: lock PER-FACILITY (bukan global) -- HARUS pakai key facility yang SAMA dengan
  // submitCount/submitValidasi/dst (lihat acquireFacilityLock_ di FacilityManagement.gs)
  // supaya rebuild tetap mutually-exclusive terhadap submit di facility yang sama, tapi
  // rebuild facility lain tidak ikut ke-block.
  const facInfoForLock_ = getUserFacility(username);
  const facLock = acquireFacilityLock_(facInfoForLock_ ? facInfoForLock_.id : null, 20000);
  if (!facLock) throw new Error('Sistem sedang sibuk, coba lagi nanti.');
  try {
    // v8.26.0: Baca Riwayat dari facility user
    const riwayatSheet = getRiwayatSheet_(username);
    const lastRow = riwayatSheet.getLastRow();
    let pendingTotal = 0;
    const pendingByValidator = {};
    const plusMap = {}, minusMap = {};
    const activeStatuses = TASK_STATUS_FLOW.slice(0, TASK_STATUS_FLOW.length - 1); // semua kecuali 'Selesai'
    if (lastRow >= 2) {
      const data = riwayatSheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
      data.forEach(function (row) {
        const statusValidasi = row[11], hasilFinal = row[14], statusTask = row[18];
        const assignedValidator = row[22];
        if (statusValidasi === 'Pending') {
          pendingTotal++;
          if (assignedValidator) pendingByValidator[assignedValidator] = (pendingByValidator[assignedValidator] || 0) + 1;
        }
        if (hasilFinal === 'DISCREPANCY' && activeStatuses.indexOf(statusTask) !== -1) {
          const article = row[3], description = row[4], lokasi = row[2], selisih = row[9];
          const map = Number(selisih) > 0 ? plusMap : minusMap;
          const qtyAbs = Math.abs(Number(selisih) || 0);
          if (!qtyAbs || !article) return;
          if (!map[article]) map[article] = { description: description, qty: 0, lokasiCounts: {} };
          map[article].qty += qtyAbs;
          map[article].lokasiCounts[lokasi] = (map[article].lokasiCounts[lokasi] || 0) + 1;
        }
      });
    }
    // v8.26.0: Tulis ke Antrian_Aktif milik facility user
    const sheet = getAntrianAktifSheet_(username);
    const oldLastRow = sheet.getLastRow();
    if (oldLastRow > 1) sheet.getRange(2, 1, oldLastRow - 1, ANTRIAN_AKTIF_HEADERS.length).clearContent();
    const rows = [
      ['pending_total', JSON.stringify({ count: pendingTotal })],
      [PLUSMINUS_KEY_PLUS_, JSON.stringify(plusMap)],
      [PLUSMINUS_KEY_MINUS_, JSON.stringify(minusMap)]
    ];
    Object.keys(pendingByValidator).forEach(function (nama) {
      rows.push(['pending_validator:' + nama, JSON.stringify({ count: pendingByValidator[nama] })]);
    });
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    // reset semua cache row-index lama (nomor baris sudah pasti berubah abis rewrite total)
    // v8.26.2: prefix facility supaya yang dibersihkan cache milik facility ini saja
    try {
      const prefix = antrianCacheFacilityPrefix_(username);
      CacheService.getScriptCache().removeAll(rows.map(function (r) { return 'antrianRow_' + prefix + '_' + r[0]; }));
    } catch (e) { /* abaikan */ }
    return { success: true, pendingTotal: pendingTotal, jumlahValidator: Object.keys(pendingByValidator).length, jumlahSkuPlus: Object.keys(plusMap).length, jumlahSkuMinus: Object.keys(minusMap).length };
  } finally {
    facLock.release();
  }
}

function rebuildAntrianAktifMenu_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Ini akan menghitung ULANG seluruh sheet "' + ANTRIAN_AKTIF_SHEET_NAME + '" (antrian belum divalidasi & Summary Plus/Minus aktif) dari isi Riwayat saat ini. Dipakai untuk setup awal atau kalau angkanya dicurigai tidak sinkron. Lanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  // v8.26.0: Menu dipanggil dari spreadsheet aktif, jadi username tidak diketahui.
  // Pakai null/fallback ke active spreadsheet (sesuai konteks menu).
  const hasil = rebuildAntrianAktif_(null);
  ui.alert('Selesai. Antrian Pending: ' + hasil.pendingTotal + ' item (' + hasil.jumlahValidator + ' validator). ' +
    'SKU Plus aktif: ' + hasil.jumlahSkuPlus + ', SKU Minus aktif: ' + hasil.jumlahSkuMinus + '.');
}
