/**
 * AppApi.gs — fungsi untuk aplikasi Android v2: satu layar, satu panggilan.
 *
 * Tiap panggilan ke Web App memakan 1-3 detik, jadi layar yang dulu memanggil 3-5 fungsi
 * (Home: getHomeSummary + getMyPendingCount + getPendingValidasiCount + getPlusMinusSummary +
 * getPendingBacklog) sekarang cukup satu, dan tiap sheet dibaca paling banyak satu kali.
 * Semua fungsi di sini hanya MEMBACA (kecuali catatan Log_Akses yang sama dengan getHomeSummary);
 * tidak ada skema sheet baru.
 */

/** Tanggal 'yyyy-MM-dd' menurut Asia/Jakarta (UTC+7, tanpa DST) -- setara Utilities.formatDate, tanpa memanggilnya per baris. */
function tagWib_(d) {
  return new Date(d.getTime() + 7 * 3600000).toISOString().slice(0, 10);
}

/** Seluruh isi Antrian_Aktif (sheet kecil: beberapa kunci) dalam satu kali baca: { key: teks JSON }. */
function bacaAntrianSemua_(username) {
  const sheet = getAntrianAktifSheet_(username);
  const last = sheet.getLastRow();
  const map = {};
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) { map[String(r[0])] = r[1]; });
  }
  return map;
}

// ---------- Home ----------

/**
 * Semua data Home dalam satu panggilan.
 *
 * - totalCycleHariIni / selesaiHariIni / belumValidasi: persis getHomeSummary (admin & developer
 *   melihat angka facility, peran lain melihat angka miliknya).
 * - myPending: tugas cycle milik user ini yang belum dihitung.
 * - outstanding: admin & developer = seluruh item Pending di facility; peran lain = myPending.
 * - pendingValidasi, openTasks, plusMinus: hanya untuk inventory, admin, developer.
 * - perPetugas: hanya admin & developer -- per orang: sudah dihitung hari ini, sudah final, sisa.
 * - belumDitugaskan: item Pending tanpa nama petugas (seharusnya 0; ada bila sheet diedit manual).
 * - serverTime: epoch ms, supaya jam "diperbarui" ditampilkan dalam zona waktu HP.
 * - hariIni: tanggal kerja server (Asia/Jakarta); aplikasi memakainya sebagai "hari ini".
 * - tanpaFacility: true bila akun belum punya facility aktif (semua angka 0, tidak ada sheet dibaca).
 */
function getHomeBundle(username) {
  const info = requireRole_(username, ALL_CYCLE_LIKE_ROLES);
  const isAdmin = info.role === 'admin' || info.role === 'developer';
  const isInventory = info.role === 'inventory';
  const now = new Date();
  const today = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd');
  const out = {
    role: info.role,
    version: APP_VERSION,
    lastUpdated: Utilities.formatDate(now, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm'),
    serverTime: now.getTime(),
    hariIni: today,
    totalCycleHariIni: 0,
    selesaiHariIni: 0,
    belumValidasi: 0,
    myPending: 0,
    outstanding: 0,
    pendingValidasi: 0,
    openTasks: 0,
    plusMinus: null,
    perPetugas: [],
    belumDitugaskan: 0
  };

  // Tanpa facility aktif, pengambil sheet per-user jatuh ke spreadsheet pusat: itu data facility lain.
  if (!requireUserFacility_(username)) {
    out.tanpaFacility = true;
    return out;
  }
  catatAksesLog_(username, info);

  // 1) Ringkasan hari ini: satu baris Ringkasan_Harian (sama dengan getHomeSummary)
  const todayRows = readRingkasanHarianRows_('daily', today, username);
  const petugasRaw = todayRows.length ? parseJsonSafe_(todayRows[0][7], {}) : {};
  if (todayRows.length) {
    if (isAdmin) {
      out.totalCycleHariIni = Number(todayRows[0][1]) || 0;                                  // KPI_Total
      out.selesaiHariIni = (Number(todayRows[0][2]) || 0) + (Number(todayRows[0][3]) || 0);   // KPI_Hit + KPI_Disc
    } else {
      const mine = petugasRaw[info.displayName];
      if (mine) { out.totalCycleHariIni = mine.rawTotal || 0; out.selesaiHariIni = mine.rawSelesai || 0; }
    }
  }

  // 2) Antrian_Aktif: penghitung validasi dan pelacak plus minus, satu kali baca
  if (isAdmin || isInventory) {
    const antrian = bacaAntrianSemua_(username);
    const kunci = isAdmin ? 'pending_total' : ('pending_validator:' + info.displayName);
    out.belumValidasi = (parseJsonSafe_(antrian[kunci], { count: 0 }).count) || 0;
    out.pendingValidasi = out.belumValidasi;
    out.plusMinus = plusMinusDariMap_(parseJsonSafe_(antrian[PLUSMINUS_KEY_PLUS_], {}), parseJsonSafe_(antrian[PLUSMINUS_KEY_MINUS_], {}));
  }

  // 3) Data Count kolom Nama_Petugas & Status (J..K): semua hitungan "belum di cycle"
  const sisaPerPetugas = {};
  const sheet = getSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const me = String(info.displayName || '').trim();
    sheet.getRange(2, 10, lastRow - 1, 2).getValues().forEach(function (r) {
      if (r[1] !== 'Pending') return;
      const nama = String(r[0] || '').trim();
      out.outstanding++;
      if (!nama) { out.belumDitugaskan++; return; }
      if (nama === me) out.myPending++;
      sisaPerPetugas[nama] = (sisaPerPetugas[nama] || 0) + 1;
    });
  }
  if (!isAdmin) { out.outstanding = out.myPending; out.belumDitugaskan = 0; }

  // 4) Riwayat kolom Status_Task (S): selisih yang belum selesai
  if (isAdmin || isInventory) {
    const riwayat = getRiwayatSheet_(username);
    const lastRiwayat = riwayat.getLastRow();
    if (lastRiwayat >= 2) {
      const aktif = TASK_STATUS_FLOW.slice(0, TASK_STATUS_FLOW.length - 1);
      riwayat.getRange(2, 19, lastRiwayat - 1, 1).getValues().forEach(function (r) {
        if (aktif.indexOf(r[0]) !== -1) out.openTasks++;
      });
    }
  }

  if (isAdmin) {
    const nama = {};
    Object.keys(petugasRaw).forEach(function (k) { nama[k] = true; });
    Object.keys(sisaPerPetugas).forEach(function (k) { nama[k] = true; });
    out.perPetugas = Object.keys(nama).map(function (k) {
      const r = petugasRaw[k] || {};
      return { nama: k, dihitung: Number(r.rawTotal) || 0, final: Number(r.rawSelesai) || 0, sisa: sisaPerPetugas[k] || 0 };
    }).sort(function (a, b) { return (b.sisa - a.sisa) || (b.dihitung - a.dihitung) || (a.nama < b.nama ? -1 : 1); });
  }

  return out;
}

// ---------- Upload Data ----------

/**
 * Isi formulir Upload Data dalam satu panggilan: petugas yang bisa ditugaskan, alat bantu bawaan,
 * jumlah lokasi aktif (0 = unggah pasti ditolak, jadi aplikasi bisa mengingatkan lebih dulu), dan
 * mode pembagian tugas yang sedang aktif.
 */
function getUploadFormData(username) {
  requireRole_(username, ['admin', 'developer']);
  const facInfo = requireUserFacility_(username);
  if (!facInfo) {
    // importRawData menolak akun tanpa facility aktif; formulirnya tidak boleh menampilkan user facility lain.
    return { users: [], equipment: { reachTruck: 0, tangga: 0 }, lokasiAktif: 0, modeAssignment: 'legacy', facilityId: '', facilityName: '', tanpaFacility: true };
  }
  return {
    users: getAssignableUsers_(facInfo.id),
    equipment: getEquipmentReadyDefaults_(),
    lokasiAktif: Object.keys(getActiveLocations_(username)).length,
    modeAssignment: getLevelAssignmentMeta_(getLevelAssignmentSheet_(username)).mode,
    facilityId: facInfo.id,
    facilityName: facInfo.nama
  };
}

// ---------- Report: Productivity ----------

/**
 * Produktivitas satu hari: siapa menghitung berapa item, kapan mulai dan terakhir, dan berapa per jam.
 *
 * Sumber: Riwayat (1 baris = 1 item yang sudah dihitung; Waktu_Cycle = jam hitung, Waktu_Validasi =
 * jam validasi). "Jam aktif" = banyaknya jam berbeda yang berisi minimal satu hitungan, jadi
 * istirahat panjang tidak ikut menurunkan angka per jam.
 *
 * Satu "hari" selalu hari kerja server (Asia/Jakarta), sama dengan Dashboard. Jamnya ditampilkan
 * menurut zona waktu HP (tzMenit), berurutan sesuai kejadian: untuk WITA, hari kerja berjalan dari
 * jam 01 sampai jam 00 berikutnya.
 *
 * @param {string} tanggal 'yyyy-MM-dd' (kosong = hari ini)
 * @param {string} username admin / developer
 * @param {number} tzMenit selisih zona waktu HP dari UTC dalam menit (WITA = 480); bawaan 420 (WIB).
 */
function getProductivity(tanggal, username, tzMenit) {
  requireRole_(username, ['admin', 'developer']);
  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  const tgl = String(tanggal || '').trim() || today;
  const hasil = {
    tanggal: tgl,
    hariIni: tgl === today,
    ringkasan: { totalItem: 0, petugasAktif: 0, perJam: 0, sisa: 0, totalValidasi: 0 },
    petugas: [], validator: [], perJam: []
  };
  const facInfo = requireUserFacility_(username);
  if (!facInfo) { hasil.tanpaFacility = true; return hasil; }

  const tz = Number(tzMenit);
  const offMs = (tzMenit === null || tzMenit === undefined || tzMenit === '' || isNaN(tz) ? 420 : tz) * 60000;
  const isDate = function (v) { return Object.prototype.toString.call(v) === '[object Date]'; };
  // Jam dinding HP sebagai nomor urut jam sejak epoch: berurutan sesuai kejadian, juga saat melewati tengah malam.
  const kunciJam = function (t) { return Math.floor((t + offMs) / 3600000); };

  const petugas = {}, validator = {}, perJam = {};
  const sheet = getRiwayatSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    // Kolom B..Q saja: Tanggal, ..., Nama_Petugas, ..., Hasil_Awal, ..., Nama_Validator, ..., Hasil_Final, Waktu_Cycle, Waktu_Validasi
    sheet.getRange(2, 2, lastRow - 1, 16).getValues().forEach(function (r) {
      const w = r[14];
      const tagBaris = isDate(w) ? tagWib_(w) : String(r[0]);
      if (tagBaris === tgl) {
        const nama = String(r[6] || '').trim() || '(tanpa nama)';
        if (!petugas[nama]) petugas[nama] = { nama: nama, total: 0, hit: 0, selisih: 0, mulai: 0, akhir: 0, jam: {} };
        const p = petugas[nama];
        p.total++;
        if (r[9] === 'HIT') p.hit++; else p.selisih++;
        if (isDate(w)) {
          const t = w.getTime(), k = kunciJam(t);
          if (!p.mulai || t < p.mulai) p.mulai = t;
          if (t > p.akhir) p.akhir = t;
          p.jam[k] = true;
          perJam[k] = (perJam[k] || 0) + 1;
        }
      }
      const wv = r[15], namaV = String(r[11] || '').trim();
      if (namaV && isDate(wv) && tagWib_(wv) === tgl) {
        if (!validator[namaV]) validator[namaV] = { nama: namaV, total: 0, hit: 0, mulai: 0, akhir: 0 };
        const v = validator[namaV], tv = wv.getTime();
        v.total++;
        if (r[13] === 'HIT') v.hit++;
        if (!v.mulai || tv < v.mulai) v.mulai = tv;
        if (tv > v.akhir) v.akhir = tv;
      }
    });
  }

  // Sisa tugas (hanya bermakna untuk hari ini) dan peran tiap orang
  const sisa = {};
  if (tgl === today) {
    const dc = getSheet_(username);
    const dcLast = dc.getLastRow();
    if (dcLast >= 2) {
      dc.getRange(2, 10, dcLast - 1, 2).getValues().forEach(function (r) {
        if (r[1] !== 'Pending') return;
        const nama = String(r[0] || '').trim();
        if (nama) sisa[nama] = (sisa[nama] || 0) + 1;
      });
    }
  }
  const peran = {};
  getAssignableUsers_(facInfo.id).forEach(function (u) { peran[u.username] = u.role; });

  Object.keys(sisa).forEach(function (nama) {
    if (!petugas[nama]) petugas[nama] = { nama: nama, total: 0, hit: 0, selisih: 0, mulai: 0, akhir: 0, jam: {} };
  });

  let totalItem = 0, totalJamOrang = 0;
  hasil.petugas = Object.keys(petugas).map(function (k) {
    const p = petugas[k];
    const jamAktif = Object.keys(p.jam).length;
    totalItem += p.total;
    totalJamOrang += jamAktif;
    return {
      nama: p.nama, role: peran[p.nama] || '', total: p.total, hit: p.hit, selisih: p.selisih,
      mulai: p.mulai, akhir: p.akhir, jamAktif: jamAktif,
      perJam: jamAktif ? Math.round((p.total / jamAktif) * 10) / 10 : 0,
      sisa: sisa[p.nama] || 0
    };
  }).sort(function (a, b) { return (b.total - a.total) || (b.sisa - a.sisa) || (a.nama < b.nama ? -1 : 1); });

  hasil.validator = Object.keys(validator).map(function (k) { return validator[k]; })
    .sort(function (a, b) { return b.total - a.total; });

  const kunci = Object.keys(perJam).map(Number).sort(function (a, b) { return a - b; });
  if (kunci.length) {
    for (let k = kunci[0]; k <= kunci[kunci.length - 1]; k++) {
      hasil.perJam.push({ jam: ((k % 24) + 24) % 24, total: perJam[k] || 0 });
    }
  }

  hasil.ringkasan = {
    totalItem: totalItem,
    petugasAktif: hasil.petugas.filter(function (p) { return p.total > 0; }).length,
    perJam: totalJamOrang ? Math.round((totalItem / totalJamOrang) * 10) / 10 : 0,
    sisa: Object.keys(sisa).reduce(function (s, k) { return s + sisa[k]; }, 0),
    totalValidasi: hasil.validator.reduce(function (s, v) { return s + v.total; }, 0)
  };
  return hasil;
}
