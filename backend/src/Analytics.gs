/**
 * Analytics.gs
 * Menu "Analytics" -- Accuracy Root Cause: kenapa akurasi turun, siapa/apa penyebabnya, dan
 * detail investigasi per item. Beda dari Dashboard.gs (yang fokus leaderboard akurasi per
 * user/periode) -- Analytics fokus ke ROOT CAUSE dari discrepancy: kategori penyebab, siapa
 * user/validator yang paling banyak terlibat, dan breakdown jenis masalah operasional.
 *
 * Sumber data: scan Riwayat langsung (BUKAN Ringkasan_Harian) -- karena breakdown per
 * User Cycle/Validator/Jenis Masalah butuh data per-baris yang tidak direpresentasikan
 * sebagai angka agregat di Ringkasan_Harian (pola sama seperti getProblemItemsDetail di
 * Dashboard.gs). Rentang tanggal (dateFrom/dateTo) bebas dipilih user, tidak dibatasi
 * daily/monthly seperti Dashboard.
 */

// ---------- Pemetaan Kategori_Selisih (yang sudah ada) -> Root Cause (5 bucket, sesuai UI) ----------
// Disepakati bareng user: kategori teknis yang sudah ada di sistem dikelompokkan jadi 5
// root cause yang lebih mudah dibaca manager. "Tidak diketahui" = task discrepancy yang
// belum/tidak punya Kategori_Selisih (biasanya task masih Open, belum diinvestigasi tuntas).
const ANALYTICS_ROOT_CAUSE_ORDER_ = [
  'Salah Picking / Move',
  'Salah Putaway',
  'Human Error Cycle',
  'Master Data / Location',
  'Tidak diketahui'
];

const ANALYTICS_ROOT_CAUSE_MAP_ = {
  'Lebih Picking': 'Salah Picking / Move',
  'Kurang Picking': 'Salah Picking / Move',
  'Lebih Move': 'Salah Picking / Move',
  'Kurang Move': 'Salah Picking / Move',
  'Barang ketemu di lokasi lain': 'Salah Putaway',
  'Pindah Lokasi (Plus-Minus)': 'Salah Putaway',
  'Salah Hitung': 'Human Error Cycle',
  'Barang Sudah di Picking': 'Human Error Cycle',
  'Adjustment Plus': 'Master Data / Location',
  'Adjustment Minus': 'Master Data / Location'
};

// Sub-kategori "masalah operasional" (tab ke-3 "Who is the main contributor" -- pengganti
// "By WMS User" sesuai keputusan user: lebih relevan lihat jenis masalah operasional
// picking/move mana yang paling sering, karena itu memang murni masalah operasional).
const ANALYTICS_OPERASIONAL_KATEGORI_ = ['Lebih Picking', 'Kurang Picking', 'Lebih Move', 'Kurang Move'];

function mapKategoriKeRootCause_(kategori) {
  const k = String(kategori || '').trim();
  if (!k) return 'Tidak diketahui';
  return ANALYTICS_ROOT_CAUSE_MAP_[k] || 'Tidak diketahui';
}

function analyticsInRange_(tgl, dateFrom, dateTo) {
  if (dateFrom && tgl < dateFrom) return false;
  if (dateTo && tgl > dateTo) return false;
  return true;
}

function analyticsEffectiveHasil_(row) {
  const statusValidasi = row[11], hasilAwal = row[10], hasilFinal = row[14];
  return statusValidasi === 'Pending' ? 'PENDING' : (statusValidasi === 'Tidak Perlu' ? hasilAwal : hasilFinal);
}

/**
 * Bundle utama layar Analytics: KPI accuracy overview, trend harian dalam rentang tanggal,
 * root cause breakdown (5 bucket), dan 3 tab kontributor (User Cycle / Validator / Jenis
 * Masalah Operasional). 1x scan Riwayat untuk semuanya (pola sama seperti getDashboardData).
 */
function getAnalyticsRootCauseData(dateFrom, dateTo, requesterUsername) {
  requireRole_(requesterUsername, ['admin', 'developer']);

  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];

  let kpiTotal = 0, kpiHit = 0, kpiDisc = 0;
  const trendAcc = {}; // tgl -> {total,hit,discrepancy}
  const rootCauseCounts = {};
  const contribUserCycle = {};
  const contribValidator = {};
  const contribOperasional = {};

  data.forEach(function (row) {
    const tgl = rowDateTag_(row);
    if (!analyticsInRange_(tgl, dateFrom, dateTo)) return;

    const effectiveHasil = analyticsEffectiveHasil_(row);
    kpiTotal++;
    if (effectiveHasil === 'HIT') kpiHit++;
    else if (effectiveHasil === 'DISCREPANCY') kpiDisc++;

    if (!trendAcc[tgl]) trendAcc[tgl] = { total: 0, hit: 0, discrepancy: 0 };
    trendAcc[tgl].total++;
    if (effectiveHasil === 'HIT') trendAcc[tgl].hit++;
    else if (effectiveHasil === 'DISCREPANCY') trendAcc[tgl].discrepancy++;

    if (effectiveHasil !== 'DISCREPANCY') return; // sisanya (root cause & kontributor) cuma untuk discrepancy nyata

    const kategoriSelisih = String(row[17] || '').trim();
    const rootCause = mapKategoriKeRootCause_(kategoriSelisih);
    rootCauseCounts[rootCause] = (rootCauseCounts[rootCause] || 0) + 1;

    const namaPetugas = String(row[7] || '').trim();
    if (namaPetugas) {
      contribUserCycle[namaPetugas] = (contribUserCycle[namaPetugas] || 0) + 1;
    }

    const namaValidator = String(row[12] || '').trim();
    if (namaValidator) {
      contribValidator[namaValidator] = (contribValidator[namaValidator] || 0) + 1;
    }

    if (ANALYTICS_OPERASIONAL_KATEGORI_.indexOf(kategoriSelisih) !== -1) {
      contribOperasional[kategoriSelisih] = (contribOperasional[kategoriSelisih] || 0) + 1;
    }
  });

  const akurasi = kpiTotal ? Math.round((kpiHit / kpiTotal) * 1000) / 10 : 0;
  const errorPersen = kpiTotal ? Math.round((kpiDisc / kpiTotal) * 1000) / 10 : 0;

  const trend = Object.keys(trendAcc).sort().map(function (tgl) {
    const b = trendAcc[tgl];
    return {
      tanggal: tgl, total: b.total, hit: b.hit, discrepancy: b.discrepancy,
      akurasi: b.total ? Math.round((b.hit / b.total) * 1000) / 10 : 0
    };
  });

  function toContribList_(map, limit) {
    const list = Object.keys(map).map(function (nama) {
      return { nama: nama, jumlah: map[nama], persen: kpiDisc ? Math.round((map[nama] / kpiDisc) * 1000) / 10 : 0 };
    }).sort(function (a, b) { return b.jumlah - a.jumlah; });
    return limit ? list.slice(0, limit) : list;
  }

  const rootCause = ANALYTICS_ROOT_CAUSE_ORDER_.map(function (label) {
    const jumlah = rootCauseCounts[label] || 0;
    return { label: label, jumlah: jumlah, persen: kpiDisc ? Math.round((jumlah / kpiDisc) * 1000) / 10 : 0 };
  });

  return {
    kpi: { total: kpiTotal, hit: kpiHit, discrepancy: kpiDisc, akurasi: akurasi, errorPersen: errorPersen },
    trend: trend,
    rootCause: rootCause,
    contributors: {
      userCycle: toContribList_(contribUserCycle, 5),
      validator: toContribList_(contribValidator, 5),
      operasional: toContribList_(contribOperasional, 5)
    }
  };
}

/**
 * Detail investigasi per root cause (lazy-load, dipanggil saat user klik salah satu bucket
 * root cause) -- pola sama seperti getErrorAnalysisDetail di Dashboard.gs. Kolom mengikuti
 * alur UX yang diminta: SKU -> Lokasi -> Qty -> User Cycle -> Validator -> User WMS ->
 * Transaksi -> Hasil Investigasi (Catatan_Penyelesaian).
 */
function getAnalyticsRootCauseDetail(rootCauseLabel, dateFrom, dateTo, requesterUsername) {
  requireRole_(requesterUsername, ['admin', 'developer']);

  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  const result = [];

  data.forEach(function (row) {
    const tgl = rowDateTag_(row);
    if (!analyticsInRange_(tgl, dateFrom, dateTo)) return;
    if (analyticsEffectiveHasil_(row) !== 'DISCREPANCY') return;

    const kategoriSelisih = String(row[17] || '').trim();
    if (mapKategoriKeRootCause_(kategoriSelisih) !== rootCauseLabel) return;

    const addWho = decodeAddWhoByType_(row[25]);
    const wmsUsers = Object.keys(addWho.move).concat(Object.keys(addWho.picking)).concat(Object.keys(addWho.lainnya));
    const transaksiTypes = [];
    if (Object.keys(addWho.move).length) transaksiTypes.push('Move');
    if (Object.keys(addWho.picking).length) transaksiTypes.push('Picking');
    if (Object.keys(addWho.lainnya).length) transaksiTypes.push('Lainnya');

    result.push({
      tanggal: tgl,
      sku: row[3],
      description: row[4],
      lokasi: row[2],
      qtyError: Math.abs(Number(row[9]) || 0),
      userCycle: row[7],
      validator: row[12] || '-',
      wmsUser: wmsUsers.length ? wmsUsers.join(', ') : '-',
      transaksi: transaksiTypes.length ? transaksiTypes.join(' + ') : '-',
      kategori: kategoriSelisih || '-',
      statusTask: row[18],
      hasilInvestigasi: row[19] || ''
    });
  });

  result.sort(function (a, b) { return a.tanggal < b.tanggal ? 1 : (a.tanggal > b.tanggal ? -1 : 0); });
  return result;
}
