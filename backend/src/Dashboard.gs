/**
 * Dashboard.gs
 * Dashboard analytics: trend, summary per periode, leaderboard user, error analysis.
 * (Hasil pemecahan dari Code.gs -- semua file .gs berbagi global scope yang sama di Apps Script,
 *  jadi urutan file tidak masalah, fungsi di file lain tetap bisa saling panggil.)
 *
 * v8.16: setiap fungsi publik di sini SEBELUMNYA baca ulang seluruh sheet Riwayat sendiri-
 * sendiri (getSummaryByPeriod, getErrorAnalysis, getTrendData masing-masing 1x scan penuh --
 * padahal dipanggil bareng tiap kali Dashboard dibuka/refresh, jadi 3x scan untuk data yang
 * SAMA). Sekarang logika hitungnya dipecah jadi compute*FromData_(data, ...) yang menerima
 * data yang SUDAH dibaca -- fungsi publik individual tetap ada (baca sendiri lalu panggil
 * helper, untuk kompatibilitas), TAPI ditambahkan 2 fungsi BUNDLE (getDashboardData,
 * getUserDashboardDetail) yang baca sheet SATU KALI lalu hitung semuanya sekaligus. Dashboard
 * (JsDashboard.html) sekarang pakai bundle ini -- dari 3x+2x scan per interaksi jadi cuma 1x+1x.
 */

// ---------- Fase 22: baca Ringkasan_Harian (precompute) untuk periode tertentu ----------
// Sheet ini KECIL (1 baris/tanggal, dibatasi ARCHIVE_AGE_MONTHS oleh Tahap 1 archiving) --
// baca-semua-lalu-filter di JS tetap jauh lebih murah dibanding scan penuh Riwayat, dan TIDAK
// tergantung berapa banyak baris Riwayat (yang bisa berisi ratusan/ribuan baris PER HARI).

/**
 * v8.26.0: Menerima parameter opsional `username` untuk baca dari facility yang benar.
 */
function readRingkasanHarianRows_(periodType, periodValue, username) {
  // v8.26.0: Meneruskan username agar baca dari spreadsheet facility user
  const sheet = getRingkasanHarianSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, RINGKASAN_HARIAN_HEADERS.length).getValues();
  return data.filter(function (row) {
    const tgl = String(row[0] || '');
    return periodType === 'monthly' ? tgl.substring(0, 7) === periodValue : tgl === periodValue;
  });
}

function computeSummaryFromRingkasan_(rows, periodValue) {
  let kpiTotal = 0, kpiHit = 0, kpiDisc = 0;
  const acc = {}, accValidator = {};

  rows.forEach(function (r) {
    kpiTotal += Number(r[1]) || 0;
    kpiHit += Number(r[2]) || 0;
    kpiDisc += Number(r[3]) || 0;

    const petugasMap = parseJsonSafe_(r[4], {});
    Object.keys(petugasMap).forEach(function (nama) {
      if (!acc[nama]) acc[nama] = { nama: nama, total: 0, hit: 0 };
      acc[nama].total += petugasMap[nama].total || 0;
      acc[nama].hit += petugasMap[nama].hit || 0;
    });

    const validatorMap = parseJsonSafe_(r[5], {});
    Object.keys(validatorMap).forEach(function (nama) {
      if (!accValidator[nama]) accValidator[nama] = { nama: nama, total: 0, hit: 0 };
      accValidator[nama].total += validatorMap[nama].total || 0;
      accValidator[nama].hit += validatorMap[nama].hit || 0;
    });
  });

  function toLeaderboard(a) {
    return Object.keys(a).map(function (k) {
      const p = a[k];
      const akurasi = p.total ? Math.round((p.hit / p.total) * 1000) / 10 : 0;
      return {
        nama: p.nama, total: p.total, hit: p.hit, akurasi: akurasi,
        persenDiscrepancy: p.total ? Math.round(((p.total - p.hit) / p.total) * 1000) / 10 : 0
      };
    }).sort(function (x, y) { return y.akurasi - x.akurasi; });
  }

  return {
    period: periodValue, leaderboard: toLeaderboard(acc), leaderboardValidator: toLeaderboard(accValidator),
    kpi: { total: kpiTotal, hit: kpiHit, discrepancy: kpiDisc }
  };
}

function computeErrorAnalysisFromRingkasan_(rows) {
  const counts = {};
  let totalWithAlasan = 0;
  rows.forEach(function (r) {
    const kategoriMap = parseJsonSafe_(r[6], {});
    Object.keys(kategoriMap).forEach(function (k) {
      counts[k] = (counts[k] || 0) + (kategoriMap[k] || 0);
      totalWithAlasan += (kategoriMap[k] || 0);
    });
  });
  return Object.keys(counts).map(function (k) {
    return { alasan: k, jumlah: counts[k], persen: totalWithAlasan ? Math.round((counts[k] / totalWithAlasan) * 1000) / 10 : 0 };
  }).sort(function (a, b) { return b.jumlah - a.jumlah; });
}

/**
 * v8.26.0: Menerima parameter opsional `username` untuk baca dari facility yang benar.
 */
function fillTrendFromRingkasan_(periodType, acc, username) {
  // v8.26.0: Meneruskan username agar baca dari spreadsheet facility user
  const sheet = getRingkasanHarianSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const data = sheet.getRange(2, 1, lastRow - 1, RINGKASAN_HARIAN_HEADERS.length).getValues();
  data.forEach(function (r) {
    const tgl = String(r[0] || '');
    const key = periodType === 'monthly' ? tgl.substring(0, 7) : tgl;
    const bucket = acc[key];
    if (!bucket) return;
    bucket.total += Number(r[1]) || 0;
    bucket.hit += Number(r[2]) || 0;
    bucket.discrepancy += Number(r[3]) || 0;
  });
}

// ---------- Trend (beberapa periode sekaligus) ----------

function buildTrendSkeleton_(periodType, count) {
  const n = count ? Number(count) : (periodType === 'monthly' ? 6 : 7);
  const now = new Date();
  const periods = [];
  const acc = {};

  for (let i = n - 1; i >= 0; i--) {
    let key;
    if (periodType === 'monthly') {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      key = Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy-MM');
    } else {
      const d = new Date(now.getTime() - i * 86400000);
      key = Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy-MM-dd');
    }
    periods.push(key);
    acc[key] = { total: 0, hit: 0, discrepancy: 0 };
  }
  return { periods: periods, acc: acc };
}

function fillTrendFromData_(data, periodType, acc) {
  data.forEach(function (row) {
    const tgl = rowDateTag_(row);
    const key = periodType === 'monthly' ? tgl.substring(0, 7) : tgl;
    const bucket = acc[key];
    if (!bucket) return; // di luar rentang N periode -- diabaikan (bukan error, cuma bukan bagian tren yang diminta)

    const statusValidasi = row[11], hasilAwal = row[10], hasilFinal = row[14];
    const effectiveHasil = statusValidasi === 'Pending' ? 'PENDING' : (statusValidasi === 'Tidak Perlu' ? hasilAwal : hasilFinal);
    bucket.total++;
    if (effectiveHasil === 'HIT') bucket.hit++;
    else if (effectiveHasil === 'DISCREPANCY') bucket.discrepancy++;
  });
}

function finalizeTrend_(periods, acc) {
  return periods.map(function (p) {
    const b = acc[p];
    return {
      period: p, total: b.total, hit: b.hit, discrepancy: b.discrepancy,
      akurasi: b.total ? Math.round((b.hit / b.total) * 1000) / 10 : 0
    };
  });
}

function getTrendData(periodType, count, requesterUsername) {
  requireRole_(requesterUsername, ['admin']);

  const skeleton = buildTrendSkeleton_(periodType, count);
  // v8.26.0: Teruskan requesterUsername agar baca dari facility user
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
    fillTrendFromData_(data, periodType, skeleton.acc);
  }
  return finalizeTrend_(skeleton.periods, skeleton.acc);
}

// ---------- Summary per periode (leaderboard petugas & validator, detail, KPI) ----------

function computeSummaryFromData_(data, periodType, periodValue) {
  const details = [];
  const acc = {};
  const accValidator = {};
  let kpiTotal = 0, kpiHit = 0, kpiDisc = 0;

  data.forEach(function (row) {
    const tgl = rowDateTag_(row);
    const match = periodType === 'monthly' ? (tgl.substring(0, 7) === periodValue) : (tgl === periodValue);
    if (!match) return;

    const lokasi = row[2], article = row[3], description = row[4], namaPetugas = row[7];
    const selisih = row[9], hasilAwal = row[10], statusValidasi = row[11], hasilFinal = row[14], statusTask = row[18];
    const namaValidator = row[12], kategoriSelisih = row[17];

    const effectiveHasil = statusValidasi === 'Pending' ? 'PENDING' : (statusValidasi === 'Tidak Perlu' ? hasilAwal : hasilFinal);

    kpiTotal++;
    if (effectiveHasil === 'HIT') kpiHit++;
    else if (effectiveHasil === 'DISCREPANCY') kpiDisc++;

    if (effectiveHasil !== 'HIT') {
      let statusLabel;
      if (effectiveHasil === 'PENDING') statusLabel = 'Menunggu Validasi';
      else if (statusTask === 'Open') statusLabel = 'Verifikasi (Investigasi)'; // v8.18.0: "Open" -> "Verifikasi" (nama tampilan, nilai tersimpan tetap 'Open')
      else if (statusTask === 'Selesai') statusLabel = 'Selesai (Ditindaklanjuti)';
      else statusLabel = 'Discrepancy';

      details.push({ tanggal: tgl, lokasi: lokasi, article: article, description: description, namaPetugas: namaPetugas, selisih: selisih, status: statusLabel });
    }

    if (statusValidasi === 'Tidak Perlu') {
      if (!acc[namaPetugas]) acc[namaPetugas] = { nama: namaPetugas, total: 0, hit: 0 };
      acc[namaPetugas].total++;
      acc[namaPetugas].hit++;
    } else if (statusValidasi === 'Selesai') {
      if (!acc[namaPetugas]) acc[namaPetugas] = { nama: namaPetugas, total: 0, hit: 0 };
      acc[namaPetugas].total++;
      if (hasilAwal === hasilFinal) acc[namaPetugas].hit++;
    }

    // Ranking Validator (v8.7): item baru dihitung kalau task-nya sudah BENAR-BENAR selesai
    // (Status_Task='Selesai') -- sama pola seperti ranking petugas yang exclude 'Pending',
    // item yang task-nya masih Open/Sedang Dicari/Menunggu Konfirmasi belum dihitung karena
    // belum jelas siapa yang salah. "Salah" validator = ketahuan saat investigasi bahwa
    // discrepancy sebenarnya bukan discrepancy nyata, tapi validator yang salah hitung saat
    // blind recount (Hasil_Final tetap DISCREPANCY + Kategori_Selisih='Salah Hitung' manual).
    // Kasus AUTO (Hasil_Final=HIT, petugas yang salah hitung) TIDAK menghitung sebagai
    // kesalahan validator -- justru validator benar karena berhasil menangkap salah hitung petugas.
    if (namaValidator && statusTask === 'Selesai') {
      if (!accValidator[namaValidator]) accValidator[namaValidator] = { nama: namaValidator, total: 0, hit: 0 };
      accValidator[namaValidator].total++;
      const validatorSalah = (hasilFinal === 'DISCREPANCY' && kategoriSelisih === 'Salah Hitung');
      if (!validatorSalah) accValidator[namaValidator].hit++;
    }
  });

  const leaderboard = Object.keys(acc).map(function (k) {
    const p = acc[k];
    const akurasi = p.total ? Math.round((p.hit / p.total) * 1000) / 10 : 0;
    return {
      nama: p.nama, total: p.total, hit: p.hit, akurasi: akurasi,
      persenDiscrepancy: p.total ? Math.round(((p.total - p.hit) / p.total) * 1000) / 10 : 0
    };
  }).sort(function (a, b) { return b.akurasi - a.akurasi; });

  const leaderboardValidator = Object.keys(accValidator).map(function (k) {
    const p = accValidator[k];
    const akurasi = p.total ? Math.round((p.hit / p.total) * 1000) / 10 : 0;
    return {
      nama: p.nama, total: p.total, hit: p.hit, akurasi: akurasi,
      persenDiscrepancy: p.total ? Math.round(((p.total - p.hit) / p.total) * 1000) / 10 : 0
    };
  }).sort(function (a, b) { return b.akurasi - a.akurasi; });

  return { period: periodValue, leaderboard: leaderboard, leaderboardValidator: leaderboardValidator, details: details, kpi: { total: kpiTotal, hit: kpiHit, discrepancy: kpiDisc } };
}

function getSummaryByPeriod(periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, ['admin']);

  // v8.26.0: Teruskan requesterUsername agar baca dari facility user
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];
  return computeSummaryFromData_(data, periodType, periodValue);
}

// ---------- Ringkasan per user (akurasi hitung) ----------

function computeUserSummaryFromData_(data, nama, periodType, periodValue) {
  let total = 0, hit = 0;

  data.forEach(function (row) {
    if (String(row[7]) !== String(nama)) return;
    const statusValidasi = row[11];
    if (statusValidasi === 'Pending') return;

    const tgl = rowDateTag_(row);
    const match = periodType === 'monthly' ? (tgl.substring(0, 7) === periodValue) : (tgl === periodValue);
    if (!match) return;

    const hasilAwal = row[10], hasilFinal = row[14];
    const petugasBenar = statusValidasi === 'Tidak Perlu' ? true : (hasilAwal === hasilFinal);
    total++;
    if (petugasBenar) hit++;
  });

  const akurasi = total ? Math.round((hit / total) * 1000) / 10 : 0;
  // persenHit sama nilainya dengan akurasi (alias eksplisit); persenDiscrepancy = kebalikannya --
  // dipisah jadi field sendiri sesuai permintaan (bukan cuma diturunkan di client dari akurasi).
  return {
    nama: nama, total: total, hit: hit, kesalahanHitung: total - hit, akurasi: akurasi,
    persenHit: akurasi,
    persenDiscrepancy: total ? Math.round(((total - hit) / total) * 1000) / 10 : 0
  };
}

function getUserSummary(nama, periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, ['admin']);

  // v8.26.0: Teruskan requesterUsername agar baca dari facility user
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];
  return computeUserSummaryFromData_(data, nama, periodType, periodValue);
}

// ---------- Analisis Kesalahan (agregat per kategori penyebab + drill-down) ----------

function computeErrorAnalysisFromData_(data, periodType, periodValue) {
  const counts = {};
  let totalWithAlasan = 0;

  data.forEach(function (row) {
    const alasan = String(row[17] || '').trim();
    if (!alasan) return;
    const tgl = rowDateTag_(row);
    const match = periodType === 'monthly' ? (tgl.substring(0, 7) === periodValue) : (tgl === periodValue);
    if (!match) return;
    counts[alasan] = (counts[alasan] || 0) + 1;
    totalWithAlasan++;
  });

  return Object.keys(counts).map(function (k) {
    return { alasan: k, jumlah: counts[k], persen: totalWithAlasan ? Math.round((counts[k] / totalWithAlasan) * 1000) / 10 : 0 };
  }).sort(function (a, b) { return b.jumlah - a.jumlah; });
}

function getErrorAnalysis(periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, ['admin']);

  // v8.26.0: Teruskan requesterUsername agar baca dari facility user
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];
  return computeErrorAnalysisFromData_(data, periodType, periodValue);
}

/**
 * v8.15: detail item per kategori alasan pada card "Analisis Kesalahan" -- dipanggil lazy saat
 * baris kategori diklik (pola sama seperti getBacklogDetailByDate), supaya payload getErrorAnalysis
 * tetap ringkas (cuma agregat per kategori) sementara detail per-item baru diambil kalau memang
 * dibuka. Field yang dikembalikan sesuai kebutuhan tampilan: Lokasi, Item(Article), Deskripsi,
 * Qty_System, Qty_Count, dan Nama_Petugas (user yang melakukan cycle count-nya).
 */
function getErrorAnalysisDetail(alasan, periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, ['admin', 'developer']);

  // v8.26.0: Teruskan requesterUsername agar baca dari facility user
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  const result = [];

  data.forEach(function (row) {
    const rowAlasan = String(row[17] || '').trim();
    if (rowAlasan !== alasan) return;
    const tgl = rowDateTag_(row);
    const match = periodType === 'monthly' ? (tgl.substring(0, 7) === periodValue) : (tgl === periodValue);
    if (!match) return;

    result.push({
      lokasi: row[2], article: row[3], description: row[4],
      qtySystem: row[6], qtyCount: row[8], namaPetugas: row[7]
    });
  });

  return result;
}

// ---------- Performa user dari sisi Task Investigasi ----------

/**
 * FASE 3 (v8): performa user dari sisi Task Investigasi (beda dari getUserSummary yang
 * fokus akurasi hitung) -- jumlah task yang DISELESAIKAN user ini (kolom Diselesaikan_Oleh)
 * pada periode ini (difilter dari Waktu_Selesai_Task, BUKAN Waktu_Cycle, supaya konsisten
 * dengan makna "diselesaikan pada periode X"), plus rata-rata durasi Waktu_Cycle -> Waktu_
 * Selesai_Task dalam jam. PENTING: durasi ini "dari task dibuka sampai ditutup", BUKAN murni
 * waktu kerja aktif -- bisa lama karena antri, bukan karena lambat. Disclaimer ini wajib
 * ditampilkan di UI (lihat catatan di Index.html) supaya tidak menyesatkan penilaian individu.
 */
function computeUserInvestigasiFromData_(data, nama, periodType, periodValue) {
  let totalSelesai = 0, totalDurasiJam = 0, withDurasi = 0;

  data.forEach(function (row) {
    if (row[18] !== 'Selesai') return; // Status_Task
    if (String(row[20] || '').trim() !== String(nama).trim()) return; // Diselesaikan_Oleh

    const waktuSelesai = row[21]; // Waktu_Selesai_Task
    if (Object.prototype.toString.call(waktuSelesai) !== '[object Date]') return;
    const tgl = Utilities.formatDate(waktuSelesai, 'Asia/Jakarta', 'yyyy-MM-dd');
    const match = periodType === 'monthly' ? (tgl.substring(0, 7) === periodValue) : (tgl === periodValue);
    if (!match) return;

    totalSelesai++;
    const waktuCycle = row[15]; // Waktu_Cycle
    if (Object.prototype.toString.call(waktuCycle) === '[object Date]') {
      totalDurasiJam += (waktuSelesai - waktuCycle) / 3600000;
      withDurasi++;
    }
  });

  return {
    totalSelesai: totalSelesai,
    avgDurasiJam: withDurasi ? Math.round((totalDurasiJam / withDurasi) * 10) / 10 : 0
  };
}

function getUserInvestigasiSummary(nama, periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, ['admin']);

  // v8.26.0: Teruskan requesterUsername agar baca dari facility user
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];
  return computeUserInvestigasiFromData_(data, nama, periodType, periodValue);
}

// ---------- BUNDLE (v8.16): 1x baca sheet untuk beberapa hasil sekaligus ----------

/**
 * Fase 22: DIUBAH -- dulu 1x scan penuh Riwayat (sudah dioptimalkan v8.16 dari 3x jadi 1x),
 * SEKARANG baca sheet "Ringkasan_Harian" yang kecil & precomputed (lihat RingkasanHarian.gs).
 * Kecepatan buka Dashboard SEKARANG TIDAK LAGI tergantung berapa banyak baris Riwayat --
 * cuma tergantung berapa banyak TANGGAL yang tercakup periode ini (maks puluhan baris).
 * Field `details` (daftar item bermasalah per baris) SENGAJA TIDAK dibundel lagi di sini --
 * sekarang lazy-load lewat getProblemItemsDetail() (dipanggil cuma kalau kartu "Detail Item
 * Bermasalah" diklik/dibuka), pola sama seperti getErrorAnalysisDetail/getBacklogDetailByDate.
 */
function getDashboardData(periodType, periodValue, trendCount, requesterUsername) {
  requireRole_(requesterUsername, ['admin', 'developer']);

  // v8.26.0: Teruskan requesterUsername agar baca dari facility user
  const rows = readRingkasanHarianRows_(periodType, periodValue, requesterUsername);
  const summary = computeSummaryFromRingkasan_(rows, periodValue);
  const errorAnalysis = computeErrorAnalysisFromRingkasan_(rows);

  const skeleton = buildTrendSkeleton_(periodType, trendCount);
  // v8.26.0: Teruskan requesterUsername agar baca dari facility user
  fillTrendFromRingkasan_(periodType, skeleton.acc, requesterUsername);
  const trend = finalizeTrend_(skeleton.periods, skeleton.acc);

  return { summary: summary, errorAnalysis: errorAnalysis, trend: trend };
}

/**
 * Fase 22 (pengganti field `details` yang dulu dibundel di getDashboardData): daftar item
 * NON-HIT (discrepancy/pending/dalam investigasi) untuk kartu "Detail Item Bermasalah" --
 * lazy-load, cuma dipanggil kalau kartunya dibuka. Ini SATU-SATUNYA bagian Dashboard yang
 * masih scan penuh Riwayat, karena butuh data mentah per baris (lokasi, SKU, status individual)
 * yang tidak direpresentasikan sebagai angka agregat di Ringkasan_Harian. Karena Riwayat sudah
 * diramping-kan oleh archiving (Tahap 1), scan ini tetap murah, dan sekarang cuma jalan kalau
 * memang dibutuhkan (bukan tiap buka Dashboard).
 */
function getProblemItemsDetail(periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, ['admin', 'developer']);

  // v8.26.0: Teruskan requesterUsername agar baca dari facility user
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  const details = [];

  data.forEach(function (row) {
    const tgl = rowDateTag_(row);
    const match = periodType === 'monthly' ? (tgl.substring(0, 7) === periodValue) : (tgl === periodValue);
    if (!match) return;

    const selisih = row[9], hasilAwal = row[10], statusValidasi = row[11], hasilFinal = row[14], statusTask = row[18];
    const effectiveHasil = statusValidasi === 'Pending' ? 'PENDING' : (statusValidasi === 'Tidak Perlu' ? hasilAwal : hasilFinal);
    if (effectiveHasil === 'HIT') return;

    let statusLabel;
    if (effectiveHasil === 'PENDING') statusLabel = 'Menunggu Validasi';
    else if (statusTask === 'Open') statusLabel = 'Verifikasi (Investigasi)';
    else if (statusTask === 'Selesai') statusLabel = 'Selesai (Ditindaklanjuti)';
    else statusLabel = 'Discrepancy';

    details.push({
      tanggal: tgl, lokasi: row[2], article: row[3], description: row[4],
      namaPetugas: row[7], selisih: selisih, status: statusLabel
    });
  });

  return details;
}

/**
 * Fase 22: bagian `summary` sekarang baca Ringkasan_Harian (cepat, precomputed). Bagian
 * `investigasi` TETAP scan Riwayat -- dimensi tanggalnya beda (berdasar Waktu_Selesai_Task,
 * kapan task DITUTUP, bukan Tanggal cycle count), jadi tidak bisa dipetakan ke Ringkasan_Harian
 * yang di-key oleh tanggal cycle. Ini tetap murah karena (a) Riwayat sudah diramping-kan Tahap
 * 1, dan (b) cuma jalan sekali per klik baris leaderboard, bukan tiap buka Dashboard.
 */
function getUserDashboardDetail(nama, periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, ['admin', 'developer']);

  // v8.26.0: Teruskan requesterUsername agar baca dari facility user
  const rows = readRingkasanHarianRows_(periodType, periodValue, requesterUsername);
  let total = 0, hit = 0;
  rows.forEach(function (r) {
    const petugasMap = parseJsonSafe_(r[4], {});
    const p = petugasMap[nama];
    if (p) { total += p.total || 0; hit += p.hit || 0; }
  });
  const akurasi = total ? Math.round((hit / total) * 1000) / 10 : 0;
  const summary = {
    nama: nama, total: total, hit: hit, kesalahanHitung: total - hit, akurasi: akurasi,
    persenHit: akurasi, persenDiscrepancy: total ? Math.round(((total - hit) / total) * 1000) / 10 : 0
  };

  // v8.26.0: Teruskan requesterUsername agar baca dari facility user
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];
  const investigasi = computeUserInvestigasiFromData_(data, nama, periodType, periodValue);

  return { summary: summary, investigasi: investigasi };
}

// NOTE (v8.15): getTopProblemLocations() & getTopProblemSKU() (card "Lokasi Bermasalah" &
// "SKU Bermasalah") dihapus -- sudah tercakup oleh card "Detail Item Bermasalah" yang lebih
// lengkap (per-item, bisa dicari & di-export), jadi 2 card ringkasan ini jadi redundan.
