/**
 * Cycle Transaksi - backend Apps Script v8.31.1-apk2 (satu file, gabungan 29 file di backend/src).
 * JANGAN diedit di sini: ubah file di backend/src lalu jalankan "npm run backend".
 *
 * PASANG PERTAMA KALI:
 * 1. Di spreadsheet: Extensions > Apps Script. Hapus isi Code.gs, tempel SELURUH file ini, simpan.
 * 2. Pilih fungsi "setupAwal" di bilah atas, klik Run, izinkan akses. Semua sheet, facility,
 *    admin pertama, dan trigger antrean dibuat otomatis.
 * 3. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Kirim URL /exec
 *    ke pengembang aplikasi: alamat itu ditanam di aplikasi (app/server.json), bukan diisi pengguna.
 *
 * PERBARUI BACKEND: tempel ulang file ini, lalu Deploy > Manage deployments > Edit > Version: New version.
 * URL /exec tidak berubah, jadi aplikasi tidak perlu disetel ulang.
 */
function setupMasterLevelSettings() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getMasterSheet_();
  if (!sheet) {
    ui.alert('Sheet "Master" belum ada. Buat dulu sheet Master (kolom A=ACTIVE LOCATION, C=USER, D=ROLE).');
    return;
  }
  if (!sheet.getRange('F1').getValue()) sheet.getRange('F1').setValue('LEVEL GROUP');
  if (!sheet.getRange('G1').getValue()) sheet.getRange('G1').setValue('PRODUKTIVITAS (qty/orang)');
  LEVEL_GROUPS.forEach(function(g, idx) {
    const row = idx + 2;
    const cellF = sheet.getRange(row, 6);
    const cellG = sheet.getRange(row, 7);
    if (!cellF.getValue()) {
      cellF.setValue('Level ' + g.min + '-' + g.max + ' (' + g.label + ')');
      cellG.setValue(g.defaultProd);
    }
  });
  sheet.getRange('F1:G1').setFontWeight('bold');
  clearMasterCache_();
  ui.alert('Setting level ditambahkan di sheet Master kolom F-G. Sesuaikan angka produktivitas, lalu upload data lagi.');
}

function setupMasterAssignmentMapping() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getMasterSheet_();
  if (!sheet) {
    ui.alert('Sheet "Master" belum ada. Buat dulu sheet Master (kolom A=ACTIVE LOCATION, C=USER, D=ROLE).');
    return;
  }
  if (!sheet.getRange('I1').getValue()) sheet.getRange('I1').setValue('KATEGORI SELISIH');
  if (!sheet.getRange('J1').getValue()) sheet.getRange('J1').setValue('USER PIC');
  if (!sheet.getRange('K1').getValue()) sheet.getRange('K1').setValue('ROLE PIC');
  const existingKategori = {};
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 9, lastRow - 1, 1).getValues().forEach(function(r) {
      const v = String(r[0] || '').trim();
      if (v) existingKategori[v] = true;
    });
  }
  let nextRow = Math.max(lastRow + 1, 2);
  TASK_KATEGORI_LIST.forEach(function(k) {
    if (existingKategori[k]) return;
    sheet.getRange(nextRow, 9).setValue(k);
    nextRow++;
  });
  sheet.getRange('I1:K1').setFontWeight('bold');
  clearMasterCache_();
  ui.alert('Kolom I/J/K (Kategori Selisih / User PIC / Role PIC) sudah disiapkan di Master. Isi kolom J & K manual untuk tiap baris kategori -- boleh lebih dari 1 baris per kategori kalau PIC-nya lebih dari 1 orang.');
}

function perbaikiHeader() {
  fixHeaderForSheet_(getSheet_(), HEADERS);
  fixHeaderForSheet_(getRiwayatSheet_(), RIWAYAT_HEADERS);
  SpreadsheetApp.getUi().alert('Header sudah diperiksa/diperbaiki di sheet "Data Count" dan "Riwayat".');
}

function resetCount() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('Reset semua hasil count di sheet Data Count (Riwayat tidak terhapus)?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
  ui.alert('Selesai direset.');
}

function repairWaktuInputDariRiwayat_(dryRun) {
  dryRun = dryRun === undefined ? true : dryRun;
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {
    totalBlank: 0,
    matched: 0,
    unmatched: 0,
    unmatchedNos: []
  };
  const dcData = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const riwayatSheet = getRiwayatSheet_();
  const rLastRow = riwayatSheet.getLastRow();
  const rData = rLastRow >= 2 ? riwayatSheet.getRange(2, 1, rLastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];
  const riwayatQueueByKey = {};
  rData.forEach(function(row) {
    const key = [ row[2], row[3], row[7], row[8], row[9], row[10] ].join('||');
    if (!riwayatQueueByKey[key]) riwayatQueueByKey[key] = [];
    riwayatQueueByKey[key].push(row[15]);
  });
  let totalBlank = 0, matched = 0;
  const unmatchedNos = [];
  dcData.forEach(function(row) {
    const status = row[10];
    const waktuInput = row[14];
    const valid = waktuInput instanceof Date && !isNaN(waktuInput.getTime());
    if (status !== 'Selesai' || valid) return;
    totalBlank++;
    const key = [ row[2], row[4], row[9], row[11], row[12], row[13] ].join('||');
    const queue = riwayatQueueByKey[key];
    if (queue && queue.length) {
      row[14] = queue.shift();
      matched++;
    } else {
      unmatchedNos.push(row[0]);
    }
  });
  Logger.log('Total baris Selesai dengan Waktu_Input kosong: ' + totalBlank);
  Logger.log('Berhasil dicocokkan ke Riwayat: ' + matched);
  Logger.log('TIDAK ketemu pasangannya di Riwayat: ' + unmatchedNos.length);
  if (unmatchedNos.length) {
    Logger.log('No yang tidak ketemu (cek manual): ' + unmatchedNos.slice(0, 100).join(', ') + (unmatchedNos.length > 100 ? ' ...(+' + (unmatchedNos.length - 100) + ' lagi)' : ''));
  }
  if (!dryRun && matched) {
    const waktuInputColumn = dcData.map(function(row) {
      return [ row[14] ];
    });
    sheet.getRange(2, 15, waktuInputColumn.length, 1).setValues(waktuInputColumn);
    SpreadsheetApp.flush();
  }
  return {
    totalBlank: totalBlank,
    matched: matched,
    unmatched: unmatchedNos.length,
    unmatchedNos: unmatchedNos
  };
}

function cekPerbaikiWaktuInputKosong() {
  const ui = SpreadsheetApp.getUi();
  const cek = repairWaktuInputDariRiwayat_(true);
  if (!cek.totalBlank) {
    ui.alert('Tidak ada baris Waktu_Input kosong yang perlu diperbaiki. ');
    return;
  }
  const pesan = 'Ditemukan ' + cek.totalBlank + ' baris Selesai dengan Waktu_Input kosong.\n\n' + '- Bisa dicocokkan & diperbaiki dari Riwayat: ' + cek.matched + '\n' + '- TIDAK ketemu pasangannya (perlu cek manual, lihat Execution Log): ' + cek.unmatched + '\n\n' + 'Lanjutkan perbaikan untuk ' + cek.matched + ' baris yang cocok?';
  const resp = ui.alert(pesan, ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  const hasil = repairWaktuInputDariRiwayat_(false);
  ui.alert('Selesai. ' + hasil.matched + ' baris Waktu_Input berhasil diperbaiki dari Riwayat.' + (hasil.unmatched ? ' ' + hasil.unmatched + ' baris tidak ketemu pasangannya, cek manual (lihat No-nya di menu Extensions > Apps Script > Executions, atau View > Logs).' : ''));
}

const ANALYTICS_ROOT_CAUSE_ORDER_ = [ 'Salah Picking / Move', 'Salah Putaway', 'Human Error Cycle', 'Master Data / Location', 'Tidak diketahui' ];

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

const ANALYTICS_OPERASIONAL_KATEGORI_ = [ 'Lebih Picking', 'Kurang Picking', 'Lebih Move', 'Kurang Move' ];

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
  return statusValidasi === 'Pending' ? 'PENDING' : statusValidasi === 'Tidak Perlu' ? hasilAwal : hasilFinal;
}

function getAnalyticsRootCauseData(dateFrom, dateTo, requesterUsername) {
  requireRole_(requesterUsername, [ 'admin', 'developer' ]);
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];
  let kpiTotal = 0, kpiHit = 0, kpiDisc = 0;
  const trendAcc = {};
  const rootCauseCounts = {};
  const contribUserCycle = {};
  const contribValidator = {};
  const contribOperasional = {};
  data.forEach(function(row) {
    const tgl = rowDateTag_(row);
    if (!analyticsInRange_(tgl, dateFrom, dateTo)) return;
    const effectiveHasil = analyticsEffectiveHasil_(row);
    kpiTotal++;
    if (effectiveHasil === 'HIT') kpiHit++; else if (effectiveHasil === 'DISCREPANCY') kpiDisc++;
    if (!trendAcc[tgl]) trendAcc[tgl] = {
      total: 0,
      hit: 0,
      discrepancy: 0
    };
    trendAcc[tgl].total++;
    if (effectiveHasil === 'HIT') trendAcc[tgl].hit++; else if (effectiveHasil === 'DISCREPANCY') trendAcc[tgl].discrepancy++;
    if (effectiveHasil !== 'DISCREPANCY') return;
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
  const akurasi = kpiTotal ? Math.round(kpiHit / kpiTotal * 1e3) / 10 : 0;
  const errorPersen = kpiTotal ? Math.round(kpiDisc / kpiTotal * 1e3) / 10 : 0;
  const trend = Object.keys(trendAcc).sort().map(function(tgl) {
    const b = trendAcc[tgl];
    return {
      tanggal: tgl,
      total: b.total,
      hit: b.hit,
      discrepancy: b.discrepancy,
      akurasi: b.total ? Math.round(b.hit / b.total * 1e3) / 10 : 0
    };
  });
  function toContribList_(map, limit) {
    const list = Object.keys(map).map(function(nama) {
      return {
        nama: nama,
        jumlah: map[nama],
        persen: kpiDisc ? Math.round(map[nama] / kpiDisc * 1e3) / 10 : 0
      };
    }).sort(function(a, b) {
      return b.jumlah - a.jumlah;
    });
    return limit ? list.slice(0, limit) : list;
  }
  const rootCause = ANALYTICS_ROOT_CAUSE_ORDER_.map(function(label) {
    const jumlah = rootCauseCounts[label] || 0;
    return {
      label: label,
      jumlah: jumlah,
      persen: kpiDisc ? Math.round(jumlah / kpiDisc * 1e3) / 10 : 0
    };
  });
  return {
    kpi: {
      total: kpiTotal,
      hit: kpiHit,
      discrepancy: kpiDisc,
      akurasi: akurasi,
      errorPersen: errorPersen
    },
    trend: trend,
    rootCause: rootCause,
    contributors: {
      userCycle: toContribList_(contribUserCycle, 5),
      validator: toContribList_(contribValidator, 5),
      operasional: toContribList_(contribOperasional, 5)
    }
  };
}

function getAnalyticsRootCauseDetail(rootCauseLabel, dateFrom, dateTo, requesterUsername) {
  requireRole_(requesterUsername, [ 'admin', 'developer' ]);
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  const result = [];
  data.forEach(function(row) {
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
  result.sort(function(a, b) {
    return a.tanggal < b.tanggal ? 1 : a.tanggal > b.tanggal ? -1 : 0;
  });
  return result;
}

function antrianCacheFacilityPrefix_(username) {
  if (!username) return 'legacy';
  const facInfo = getUserFacility(username);
  return facInfo ? facInfo.id : 'legacy';
}

function antrianRowCachePut_(key, rowIndex, username) {
  try {
    const prefix = antrianCacheFacilityPrefix_(username);
    CacheService.getScriptCache().put('antrianRow_' + prefix + '_' + key, String(rowIndex), 21600);
  } catch (e) {}
}

function antrianRowCacheGet_(key, username) {
  try {
    const prefix = antrianCacheFacilityPrefix_(username);
    const v = CacheService.getScriptCache().get('antrianRow_' + prefix + '_' + key);
    return v ? Number(v) : -1;
  } catch (e) {
    return -1;
  }
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

function readAntrianValue_(key, fallback, username) {
  const sheet = getAntrianAktifSheet_(username);
  const rowIndex = findAntrianRow_(sheet, key, username);
  if (rowIndex === -1) return fallback;
  return parseJsonSafe_(sheet.getRange(rowIndex, 2).getValue(), fallback);
}

function writeAntrianValue_(key, value, username) {
  const sheet = getAntrianAktifSheet_(username);
  let rowIndex = findAntrianRow_(sheet, key, username);
  if (rowIndex === -1) {
    rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 1, 1, 2).setValues([ [ key, JSON.stringify(value) ] ]);
    antrianRowCachePut_(key, rowIndex, username);
  } else {
    sheet.getRange(rowIndex, 2).setValue(JSON.stringify(value));
  }
}

function incrementCounterKey_(key, delta, username) {
  const current = readAntrianValue_(key, {
    count: 0
  }, username);
  const next = Math.max(0, (Number(current.count) || 0) + delta);
  writeAntrianValue_(key, {
    count: next
  }, username);
}

const PLUSMINUS_KEY_PLUS_ = 'plusminus_plus';

const PLUSMINUS_KEY_MINUS_ = 'plusminus_minus';

function adjustPlusMinusTracker_(selisih, article, description, lokasi, add, username) {
  const key = Number(selisih) > 0 ? PLUSMINUS_KEY_PLUS_ : PLUSMINUS_KEY_MINUS_;
  const qtyAbs = Math.abs(Number(selisih) || 0);
  if (!qtyAbs || !article) return;
  const map = readAntrianValue_(key, {}, username);
  if (!map[article]) map[article] = {
    description: description,
    qty: 0,
    lokasiCounts: {}
  };
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

function getPlusMinusSummaryFromAntrian_(username) {
  const plusMap = readAntrianValue_(PLUSMINUS_KEY_PLUS_, {}, username);
  const minusMap = readAntrianValue_(PLUSMINUS_KEY_MINUS_, {}, username);
  return plusMinusDariMap_(plusMap, minusMap);
}

function plusMinusDariMap_(plusMap, minusMap) {
  function toItems(map) {
    const items = Object.keys(map || {}).map(function(article) {
      const e = map[article] || {};
      const lokasiOrder = Object.keys(e.lokasiCounts || {});
      return {
        article: String(article),
        description: String(e.description == null ? '' : e.description),
        qty: Number(e.qty) || 0,
        lokasi: lokasiOrder.join(', '),
        jumlahLokasi: lokasiOrder.length
      };
    }).sort(function(a, b) {
      return b.qty - a.qty;
    });
    return {
      totalSku: items.length,
      totalQty: items.reduce(function(s, it) {
        return s + it.qty;
      }, 0),
      items: items
    };
  }
  return {
    plus: toItems(plusMap),
    minus: toItems(minusMap)
  };
}

function rebuildAntrianAktif_(username) {
  const facInfoForLock_ = getUserFacility(username);
  const facLock = acquireFacilityLock_(facInfoForLock_ ? facInfoForLock_.id : null, 2e4);
  if (!facLock) throw new Error('Sistem sedang sibuk, coba lagi nanti.');
  try {
    const riwayatSheet = getRiwayatSheet_(username);
    const lastRow = riwayatSheet.getLastRow();
    let pendingTotal = 0;
    const pendingByValidator = {};
    const plusMap = {}, minusMap = {};
    const activeStatuses = TASK_STATUS_FLOW.slice(0, TASK_STATUS_FLOW.length - 1);
    if (lastRow >= 2) {
      const data = riwayatSheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
      data.forEach(function(row) {
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
          if (!map[article]) map[article] = {
            description: description,
            qty: 0,
            lokasiCounts: {}
          };
          map[article].qty += qtyAbs;
          map[article].lokasiCounts[lokasi] = (map[article].lokasiCounts[lokasi] || 0) + 1;
        }
      });
    }
    const sheet = getAntrianAktifSheet_(username);
    const oldLastRow = sheet.getLastRow();
    if (oldLastRow > 1) sheet.getRange(2, 1, oldLastRow - 1, ANTRIAN_AKTIF_HEADERS.length).clearContent();
    const rows = [ [ 'pending_total', JSON.stringify({
      count: pendingTotal
    }) ], [ PLUSMINUS_KEY_PLUS_, JSON.stringify(plusMap) ], [ PLUSMINUS_KEY_MINUS_, JSON.stringify(minusMap) ] ];
    Object.keys(pendingByValidator).forEach(function(nama) {
      rows.push([ 'pending_validator:' + nama, JSON.stringify({
        count: pendingByValidator[nama]
      }) ]);
    });
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    try {
      const prefix = antrianCacheFacilityPrefix_(username);
      CacheService.getScriptCache().removeAll(rows.map(function(r) {
        return 'antrianRow_' + prefix + '_' + r[0];
      }));
    } catch (e) {}
    return {
      success: true,
      pendingTotal: pendingTotal,
      jumlahValidator: Object.keys(pendingByValidator).length,
      jumlahSkuPlus: Object.keys(plusMap).length,
      jumlahSkuMinus: Object.keys(minusMap).length
    };
  } finally {
    facLock.release();
  }
}

function rebuildAntrianAktifMenu_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('Ini akan menghitung ULANG seluruh sheet "' + ANTRIAN_AKTIF_SHEET_NAME + '" (antrian belum divalidasi & Summary Plus/Minus aktif) dari isi Riwayat saat ini. Dipakai untuk setup awal atau kalau angkanya dicurigai tidak sinkron. Lanjutkan?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  const hasil = rebuildAntrianAktif_(null);
  ui.alert('Selesai. Antrian Pending: ' + hasil.pendingTotal + ' item (' + hasil.jumlahValidator + ' validator). ' + 'SKU Plus aktif: ' + hasil.jumlahSkuPlus + ', SKU Minus aktif: ' + hasil.jumlahSkuMinus + '.');
}

const API_BRIDGE_ALLOW_ = [ 'getAppVersion', 'getUserRole', 'getHomeSummary', 'getPlusMinusSummary', 'getMyPendingCount', 'getPendingValidasiCount', 'getMyPendingTasks', 'submitCount', 'getPendingValidasi', 'submitValidasi', 'getEquipmentReadyDefaults', 'getAssignableUsers', 'importRawData', 'getImportProgress', 'getDashboardData', 'getErrorAnalysisDetail', 'getUserDashboardDetail', 'getPendingBacklog', 'getBacklogDetailByDate', 'getProblemItemsDetail', 'getAnalyticsRootCauseData', 'getAnalyticsRootCauseDetail', 'getInvestigasiFormData', 'getOpenTasks', 'updateTaskStatus', 'getTaskLog', 'getPlusMinusCandidates', 'closePlusMinusPair', 'getSettingOverflow', 'saveSettingOverflow', 'setOverflowForceOffToday', 'getDaftarUserMaster', 'tambahUserMaster', 'updateRoleUserMaster', 'setStatusUserMaster', 'getDaftarAksesSetting', 'tambahAksesSetting', 'hapusAksesSetting', 'getLevelAssignmentConfig', 'saveLevelAssignmentMatrix', 'setModeAssignment', 'getLogPerubahanConfig', 'getDaftarFacility', 'tambahFacility', 'daftarkanFacilityExisting', 'updateNamaFacility', 'setStatusFacility', 'assignUserKeFacility', 'getDaftarUserFacilityAssignment', 'importLokasiAktif', 'getDaftarLokasiAktif', 'copyLokasiDariFacility', 'setDeveloperActiveFacility', 'getHomeBundle', 'getUploadFormData', 'getProductivity', 'getOpenTaskCount' ];

function apiBridge_(req) {
  if (!req || !Array.isArray(req.args)) return null;
  const name = String(req.action || '');
  let out;
  if (API_BRIDGE_ALLOW_.indexOf(name) === -1) {
    out = {
      ok: false,
      error: 'Fungsi tidak dikenal: ' + name
    };
  } else {
    const scope = typeof globalThis !== 'undefined' ? globalThis : this;
    const fn = scope[name];
    if (typeof fn !== 'function') {
      out = {
        ok: false,
        error: 'Fungsi "' + name + '" tidak ada di backend ini.'
      };
    } else {
      try {
        const result = fn.apply(null, req.args);
        out = {
          ok: true,
          result: result === undefined ? null : result
        };
      } catch (err) {
        out = {
          ok: false,
          error: err && err.message ? err.message : String(err)
        };
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function tagWib_(d) {
  return new Date(d.getTime() + 7 * 36e5).toISOString().slice(0, 10);
}

function bacaAntrianSemua_(username) {
  const sheet = getAntrianAktifSheet_(username);
  const last = sheet.getLastRow();
  const map = {};
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, 2).getValues().forEach(function(r) {
      map[String(r[0])] = r[1];
    });
  }
  return map;
}

function getHomeBundle(username) {
  const info = requireRole_(username, ALL_CYCLE_LIKE_ROLES);
  const isAdmin = info.role === 'admin' || info.role === 'developer';
  const isInventory = info.role === 'inventory';
  const now = new Date;
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
  if (!requireUserFacility_(username)) {
    out.tanpaFacility = true;
    return out;
  }
  catatAksesLog_(username, info);
  const todayRows = readRingkasanHarianRows_('daily', today, username);
  const petugasRaw = todayRows.length ? parseJsonSafe_(todayRows[0][7], {}) : {};
  if (todayRows.length) {
    if (isAdmin) {
      out.totalCycleHariIni = Number(todayRows[0][1]) || 0;
      out.selesaiHariIni = (Number(todayRows[0][2]) || 0) + (Number(todayRows[0][3]) || 0);
    } else {
      const mine = petugasRaw[info.displayName];
      if (mine) {
        out.totalCycleHariIni = mine.rawTotal || 0;
        out.selesaiHariIni = mine.rawSelesai || 0;
      }
    }
  }
  if (isAdmin || isInventory) {
    const antrian = bacaAntrianSemua_(username);
    const kunci = isAdmin ? 'pending_total' : 'pending_validator:' + info.displayName;
    out.belumValidasi = parseJsonSafe_(antrian[kunci], {
      count: 0
    }).count || 0;
    out.pendingValidasi = out.belumValidasi;
    out.plusMinus = plusMinusDariMap_(parseJsonSafe_(antrian[PLUSMINUS_KEY_PLUS_], {}), parseJsonSafe_(antrian[PLUSMINUS_KEY_MINUS_], {}));
  }
  const sisaPerPetugas = {};
  const sheet = getSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const me = String(info.displayName || '').trim();
    sheet.getRange(2, 10, lastRow - 1, 2).getValues().forEach(function(r) {
      if (r[1] !== 'Pending') return;
      const nama = String(r[0] || '').trim();
      out.outstanding++;
      if (!nama) {
        out.belumDitugaskan++;
        return;
      }
      if (nama === me) out.myPending++;
      sisaPerPetugas[nama] = (sisaPerPetugas[nama] || 0) + 1;
    });
  }
  if (!isAdmin) {
    out.outstanding = out.myPending;
    out.belumDitugaskan = 0;
  }
  if (isAdmin || isInventory) {
    const riwayat = getRiwayatSheet_(username);
    const lastRiwayat = riwayat.getLastRow();
    if (lastRiwayat >= 2) {
      const aktif = TASK_STATUS_FLOW.slice(0, TASK_STATUS_FLOW.length - 1);
      riwayat.getRange(2, 19, lastRiwayat - 1, 1).getValues().forEach(function(r) {
        if (aktif.indexOf(r[0]) !== -1) out.openTasks++;
      });
    }
  }
  if (isAdmin) {
    const nama = {};
    Object.keys(petugasRaw).forEach(function(k) {
      nama[k] = true;
    });
    Object.keys(sisaPerPetugas).forEach(function(k) {
      nama[k] = true;
    });
    out.perPetugas = Object.keys(nama).map(function(k) {
      const r = petugasRaw[k] || {};
      return {
        nama: k,
        dihitung: Number(r.rawTotal) || 0,
        final: Number(r.rawSelesai) || 0,
        sisa: sisaPerPetugas[k] || 0
      };
    }).sort(function(a, b) {
      return b.sisa - a.sisa || b.dihitung - a.dihitung || (a.nama < b.nama ? -1 : 1);
    });
  }
  return out;
}

function getUploadFormData(username) {
  requireRole_(username, [ 'admin', 'developer' ]);
  const facInfo = requireUserFacility_(username);
  if (!facInfo) {
    return {
      users: [],
      equipment: {
        reachTruck: 0,
        tangga: 0
      },
      lokasiAktif: 0,
      modeAssignment: 'legacy',
      facilityId: '',
      facilityName: '',
      tanpaFacility: true
    };
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

function getProductivity(tanggal, username, tzMenit) {
  requireRole_(username, [ 'admin', 'developer' ]);
  const today = Utilities.formatDate(new Date, 'Asia/Jakarta', 'yyyy-MM-dd');
  const tgl = String(tanggal || '').trim() || today;
  const hasil = {
    tanggal: tgl,
    hariIni: tgl === today,
    ringkasan: {
      totalItem: 0,
      petugasAktif: 0,
      perJam: 0,
      sisa: 0,
      totalValidasi: 0
    },
    petugas: [],
    validator: [],
    perJam: []
  };
  const facInfo = requireUserFacility_(username);
  if (!facInfo) {
    hasil.tanpaFacility = true;
    return hasil;
  }
  const tz = Number(tzMenit);
  const offMs = (tzMenit === null || tzMenit === undefined || tzMenit === '' || isNaN(tz) ? 420 : tz) * 6e4;
  const isDate = function(v) {
    return Object.prototype.toString.call(v) === '[object Date]';
  };
  const kunciJam = function(t) {
    return Math.floor((t + offMs) / 36e5);
  };
  const petugas = {}, validator = {}, perJam = {};
  const sheet = getRiwayatSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 2, lastRow - 1, 16).getValues().forEach(function(r) {
      const w = r[14];
      const tagBaris = isDate(w) ? tagWib_(w) : String(r[0]);
      if (tagBaris === tgl) {
        const nama = String(r[6] || '').trim() || '(tanpa nama)';
        if (!petugas[nama]) petugas[nama] = {
          nama: nama,
          total: 0,
          hit: 0,
          selisih: 0,
          mulai: 0,
          akhir: 0,
          jam: {}
        };
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
        if (!validator[namaV]) validator[namaV] = {
          nama: namaV,
          total: 0,
          hit: 0,
          mulai: 0,
          akhir: 0
        };
        const v = validator[namaV], tv = wv.getTime();
        v.total++;
        if (r[13] === 'HIT') v.hit++;
        if (!v.mulai || tv < v.mulai) v.mulai = tv;
        if (tv > v.akhir) v.akhir = tv;
      }
    });
  }
  const sisa = {};
  if (tgl === today) {
    const dc = getSheet_(username);
    const dcLast = dc.getLastRow();
    if (dcLast >= 2) {
      dc.getRange(2, 10, dcLast - 1, 2).getValues().forEach(function(r) {
        if (r[1] !== 'Pending') return;
        const nama = String(r[0] || '').trim();
        if (nama) sisa[nama] = (sisa[nama] || 0) + 1;
      });
    }
  }
  const peran = {};
  getAssignableUsers_(facInfo.id).forEach(function(u) {
    peran[u.username] = u.role;
  });
  Object.keys(sisa).forEach(function(nama) {
    if (!petugas[nama]) petugas[nama] = {
      nama: nama,
      total: 0,
      hit: 0,
      selisih: 0,
      mulai: 0,
      akhir: 0,
      jam: {}
    };
  });
  let totalItem = 0, totalJamOrang = 0;
  hasil.petugas = Object.keys(petugas).map(function(k) {
    const p = petugas[k];
    const jamAktif = Object.keys(p.jam).length;
    totalItem += p.total;
    totalJamOrang += jamAktif;
    return {
      nama: p.nama,
      role: peran[p.nama] || '',
      total: p.total,
      hit: p.hit,
      selisih: p.selisih,
      mulai: p.mulai,
      akhir: p.akhir,
      jamAktif: jamAktif,
      perJam: jamAktif ? Math.round(p.total / jamAktif * 10) / 10 : 0,
      sisa: sisa[p.nama] || 0
    };
  }).sort(function(a, b) {
    return b.total - a.total || b.sisa - a.sisa || (a.nama < b.nama ? -1 : 1);
  });
  hasil.validator = Object.keys(validator).map(function(k) {
    return validator[k];
  }).sort(function(a, b) {
    return b.total - a.total;
  });
  const kunci = Object.keys(perJam).map(Number).sort(function(a, b) {
    return a - b;
  });
  if (kunci.length) {
    for (let k = kunci[0]; k <= kunci[kunci.length - 1]; k++) {
      hasil.perJam.push({
        jam: (k % 24 + 24) % 24,
        total: perJam[k] || 0
      });
    }
  }
  hasil.ringkasan = {
    totalItem: totalItem,
    petugasAktif: hasil.petugas.filter(function(p) {
      return p.total > 0;
    }).length,
    perJam: totalJamOrang ? Math.round(totalItem / totalJamOrang * 10) / 10 : 0,
    sisa: Object.keys(sisa).reduce(function(s, k) {
      return s + sisa[k];
    }, 0),
    totalValidasi: hasil.validator.reduce(function(s, v) {
      return s + v.total;
    }, 0)
  };
  return hasil;
}

function hitungCutoffArchive_() {
  const now = new Date;
  return new Date(now.getFullYear(), now.getMonth() - ARCHIVE_AGE_MONTHS, now.getDate());
}

function archiveOldRiwayat_(dryRun, username) {
  const facInfoForLock_ = getUserFacility(username);
  const facLock = acquireFacilityLock_(facInfoForLock_ ? facInfoForLock_.id : null, 2e4);
  if (!facLock) throw new Error('Sistem sedang sibuk, coba lagi nanti.');
  try {
    const sheet = getRiwayatSheet_(username);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {
      success: true,
      archived: 0,
      sisa: 0,
      dryRun: !!dryRun
    };
    const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
    const cutoff = hitungCutoffArchive_();
    const keep = [];
    const toArchive = [];
    data.forEach(function(row) {
      const statusTask = row[18];
      const tglStr = rowDateTag_(row);
      const tgl = tglStr ? new Date(String(tglStr) + 'T00:00:00') : null;
      const eligible = statusTask === 'Selesai' && tgl && !isNaN(tgl.getTime()) && tgl < cutoff;
      if (eligible) toArchive.push(row); else keep.push(row);
    });
    if (toArchive.length === 0) {
      return {
        success: true,
        archived: 0,
        sisa: keep.length,
        dryRun: !!dryRun
      };
    }
    if (dryRun) {
      return {
        success: true,
        archived: toArchive.length,
        sisa: keep.length,
        dryRun: true
      };
    }
    const archiveSheet = getRiwayatArchiveSheet_(username);
    const archiveLastRow = archiveSheet.getLastRow();
    archiveSheet.getRange(archiveLastRow + 1, 1, toArchive.length, RIWAYAT_HEADERS.length).setValues(toArchive);
    sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).clearContent();
    if (keep.length > 0) {
      sheet.getRange(2, 1, keep.length, RIWAYAT_HEADERS.length).setValues(keep);
    }
    return {
      success: true,
      archived: toArchive.length,
      sisa: keep.length,
      dryRun: false
    };
  } finally {
    facLock.release();
  }
}

function cekArchiveRiwayatLama_() {
  const ui = SpreadsheetApp.getUi();
  const cek = archiveOldRiwayat_(true, null);
  if (!cek.archived) {
    ui.alert('Tidak ada baris Riwayat yang perlu diarsipkan saat ini (kriteria: Status_Task=Selesai & usia > ' + ARCHIVE_AGE_MONTHS + ' bulan).');
    return;
  }
  const pesan = 'Ditemukan ' + cek.archived + ' baris Riwayat (Status_Task=Selesai, usia > ' + ARCHIVE_AGE_MONTHS + ' bulan) siap diarsipkan.\n\n' + 'Sisa yang TETAP di Riwayat (masih aktif atau masih baru): ' + cek.sisa + ' baris.\n\n' + 'Baris yang diarsipkan akan dipindah ke sheet "' + RIWAYAT_ARCHIVE_SHEET_NAME + '" -- data TIDAK dihapus, hanya dipindah, tetap bisa dibuka manual kapan saja.\n\n' + 'Lanjutkan?';
  const resp = ui.alert(pesan, ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  const hasil = archiveOldRiwayat_(false, null);
  ui.alert('Selesai. ' + hasil.archived + ' baris dipindah ke "' + RIWAYAT_ARCHIVE_SHEET_NAME + '". Sisa ' + hasil.sisa + ' baris aktif di Riwayat.');
}

const APP_VERSION = 'v8.31.1-apk2';

const MASTER_SHEET_NAME = 'Master';

const SHEET_NAME = 'Data Count';

const RIWAYAT_SHEET_NAME = 'Riwayat';

const HEADERS = [ 'No', 'Tanggal_Upload', 'Lokasi', 'Level', 'Article', 'Description', 'Qty_Transaksi', 'Qty_System', 'Batch', 'Nama_Petugas', 'Status', 'Qty_Count', 'Selisih', 'Hasil_Awal', 'Waktu_Input', 'AddWho_Transaksi' ];

const LOG_AKSES_SHEET_NAME = 'Log_Akses';

const LOG_AKSES_HEADERS = [ 'Tanggal', 'Username', 'Nama_Petugas', 'Role', 'Versi_App', 'Waktu_Pertama', 'Waktu_Terakhir', 'Jumlah_Buka' ];

const LOG_ANOMALI_SHEET_NAME = 'Log_Anomali';

const LOG_ANOMALI_HEADERS = [ 'Waktu_Deteksi', 'No', 'Row_Index', 'Nama_Petugas', 'Waktu_Seharusnya', 'Nilai_Terbaca_Saat_Cek', 'Status_Retry' ];

const LOG_UNASSIGNED_SHEET_NAME = 'Log_Unassigned';

const LOG_UNASSIGNED_HEADERS = [ 'Waktu', 'Admin', 'Pesan_Warning' ];

const LOG_BUKTI_SHEET_NAME = 'Log_Bukti_Investigasi';

const LOG_BUKTI_HEADERS = [ 'Waktu_Submit', 'ID_Riwayat', 'Lokasi_Task', 'Article_Task', 'Nama_User', 'No', 'STORERKEY', 'TRANTYPE', 'SKU', 'Description', 'SKUGROUP', 'LOT', 'FROMLOC', 'FROMID', 'TOLOC', 'TOID', 'SOURCEKEY', 'QTY', 'ADDDATE', 'ADDWHO' ];

const LEVEL_GROUPS = [ {
  key: 'tanpa_alat',
  label: 'Tanpa Alat Bantu',
  min: 1,
  max: 2,
  defaultProd: 150
}, {
  key: 'tangga',
  label: 'Tangga Pesawat',
  min: 3,
  max: 4,
  defaultProd: 100
}, {
  key: 'reach_truck',
  label: 'Reach Truck',
  min: 5,
  max: 6,
  defaultProd: 60
} ];

const CONFIG_SISTEM_SHEET_NAME = 'Config_Sistem';

const CONFIG_SISTEM_HEADERS = [ 'Key', 'Value', 'Keterangan', 'Waktu_Update', 'Diubah_Oleh' ];

const CONFIG_SISTEM_DEFAULTS_ = {
  overflow_enabled: 'FALSE',
  overflow_min_sisa_reach_truck: '1',
  overflow_maks_orang: '',
  overflow_force_off_today: ''
};

const CONFIG_SISTEM_KETERANGAN_ = {
  overflow_enabled: 'Jika diaktifkan, saat demand Level 5 & 6 rendah maka tugas Level 3 dan 4 akan di-assign ke Storing.',
  overflow_min_sisa_reach_truck: 'Minimal user Storing yang wajib tetap di Level 5-6, tidak ikut dialihkan.',
  overflow_maks_orang: 'Maks user Storing yang boleh dialihkan per import. Kosongkan = tanpa batas.',
  overflow_force_off_today: '(otomatis) Tanggal rem darurat overflow dari tombol "Nonaktifkan Overflow Hari Ini".'
};

const CONFIG_AKSES_SETTING_SHEET_NAME = 'Config_Akses_Setting';

const CONFIG_AKSES_SETTING_HEADERS = [ 'NIK', 'Nama', 'Tanggal_Ditambahkan' ];

const CYCLE_ROLES = [ 'outbound', 'storing', 'inbound', 'lp', 'maintenance' ];

const CYCLE_ROLES_WITH_LEGACY = CYCLE_ROLES.concat([ 'cycle' ]);

const ALL_CYCLE_LIKE_ROLES = CYCLE_ROLES_WITH_LEGACY.concat([ 'inventory', 'admin', 'developer' ]);

const DEVELOPER_ROLE = 'developer';

const TASK_KATEGORI_LIST = [ 'Lebih Picking', 'Kurang Picking', 'Lebih Move', 'Kurang Move', 'Barang ketemu di lokasi lain', 'Adjustment Plus', 'Adjustment Minus', 'Salah Hitung', 'Barang Sudah di Picking' ];

const KATEGORI_TANPA_BUKTI_ = [ 'Salah Hitung', 'Barang Sudah di Picking' ];

const KATEGORI_ADJUSTMENT_ = [ 'Adjustment Plus', 'Adjustment Minus' ];

const KATEGORI_PLUS_MINUS_PAIR_ = 'Pindah Lokasi (Plus-Minus)';

const TASK_STATUS_FLOW = [ 'Open', 'Sedang Dicari', 'Menunggu Konfirmasi', 'Selesai' ];

const TASK_SLA_HARI = 1;

const PRIORITAS_BOBOT_UMUR = 5;

const RIWAYAT_HEADERS = [ 'ID', 'Tanggal', 'Lokasi', 'Article', 'Description', 'Qty_Transaksi', 'Qty_System', 'Nama_Petugas', 'Qty_Count', 'Selisih', 'Hasil_Awal', 'Status_Validasi', 'Nama_Validator', 'Qty_Validasi', 'Hasil_Final', 'Waktu_Cycle', 'Waktu_Validasi', 'Kategori_Selisih', 'Status_Task', 'Catatan_Penyelesaian', 'Diselesaikan_Oleh', 'Waktu_Selesai_Task', 'Assigned_Validator', 'PIC_Investigasi', 'Waktu_Update_Status', 'AddWho_Transaksi', 'Pasangan_Task_ID' ];

const RIWAYAT_COL_PASANGAN_TASK_ID_ = RIWAYAT_HEADERS.indexOf('Pasangan_Task_ID') + 1;

const RIWAYAT_ARCHIVE_SHEET_NAME = 'Riwayat_Archive';

const ARCHIVE_AGE_MONTHS = 3;

const RINGKASAN_HARIAN_SHEET_NAME = 'Ringkasan_Harian';

const RINGKASAN_HARIAN_HEADERS = [ 'Tanggal', 'KPI_Total', 'KPI_Hit', 'KPI_Disc', 'Petugas_JSON', 'Validator_JSON', 'Kategori_JSON', 'Petugas_Raw_JSON' ];

const ANTRIAN_AKTIF_SHEET_NAME = 'Antrian_Aktif';

const ANTRIAN_AKTIF_HEADERS = [ 'Key', 'Value_JSON' ];

const QUEUE_RINGKASAN_SHEET_NAME = 'Queue_Ringkasan_Delta';

const QUEUE_RINGKASAN_HEADERS = [ 'Timestamp', 'Tanggal', 'RowJSON' ];

const QUEUE_COUNTER_SHEET_NAME = 'Queue_Antrian_Counter';

const QUEUE_COUNTER_HEADERS = [ 'Timestamp', 'Key', 'Delta' ];

const SUBMIT_QUEUE_SHEET_NAME = 'Queue_Submit_Cycle';

const SUBMIT_QUEUE_HEADERS = [ 'Timestamp', 'No', 'NamaPetugas', 'QtyCount', 'FacilityId', 'WaktuHitung' ];

const VALIDASI_QUEUE_SHEET_NAME = 'Queue_Validasi';

const VALIDASI_QUEUE_HEADERS = [ 'Timestamp', 'Id', 'NamaValidator', 'QtyValidasi', 'FacilityId', 'WaktuHitung' ];

const TASK_QUEUE_SHEET_NAME = 'Queue_Task_Investigasi';

const TASK_QUEUE_HEADERS = [ 'Timestamp', 'Id', 'NamaUser', 'NewStatus', 'Catatan', 'Kategori', 'PicUsername', 'BuktiRowsJSON', 'FacilityId' ];

function getMyPendingTasks(username) {
  requireRole_(username, ALL_CYCLE_LIKE_ROLES);
  const sheet = getSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const uname = String(username || '').trim();
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const items = [];
  data.forEach(function(row) {
    const namaPetugas = String(row[9] || '').trim();
    const status = row[10];
    if (namaPetugas === uname && status === 'Pending') {
      const lokasi = row[2];
      const group = getLevelGroupByLocation_(lokasi);
      items.push({
        no: row[0],
        lokasi: lokasi,
        article: row[4],
        description: row[5],
        area: extractArea_(lokasi),
        levelGroupKey: group.key,
        levelGroupLabel: group.label,
        qtySystem: Number(row[7])
      });
    }
  });
  return items;
}

function getMyPendingCount(username) {
  requireRole_(username, ALL_CYCLE_LIKE_ROLES);
  const sheet = getSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const uname = String(username || '').trim();
  const data = sheet.getRange(2, 10, lastRow - 1, 2).getValues();
  let count = 0;
  data.forEach(function(row) {
    if (String(row[0] || '').trim() === uname && row[1] === 'Pending') count++;
  });
  return count;
}

function findDataCountRowAndValues_(sheet, no) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  let estimatedRow = Number(no) + 1;
  if (estimatedRow >= 2 && estimatedRow <= lastRow) {
    const rowValues = sheet.getRange(estimatedRow, 1, 1, HEADERS.length).getValues()[0];
    if (Number(rowValues[0]) === Number(no)) {
      return {
        rowIndex: estimatedRow,
        rowValues: rowValues
      };
    }
  }
  const firstNo = Number(sheet.getRange(2, 1).getValue());
  estimatedRow = Number(no) - firstNo + 2;
  if (estimatedRow >= 2 && estimatedRow <= lastRow) {
    const rowValues = sheet.getRange(estimatedRow, 1, 1, HEADERS.length).getValues()[0];
    if (Number(rowValues[0]) === Number(no)) {
      return {
        rowIndex: estimatedRow,
        rowValues: rowValues
      };
    }
  }
  const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (Number(colA[i][0]) === Number(no)) {
      const r = i + 2;
      const rowValues = sheet.getRange(r, 1, 1, HEADERS.length).getValues()[0];
      return {
        rowIndex: r,
        rowValues: rowValues
      };
    }
  }
  return null;
}

function submitCount(no, namaPetugas, qtyCount, waktuKlien, facKlien) {
  requireRole_(namaPetugas, ALL_CYCLE_LIKE_ROLES);
  const aktual = Number(qtyCount);
  if (qtyCount === '' || qtyCount === null || qtyCount === undefined || isNaN(aktual) || aktual < 0) {
    return {
      success: false,
      message: 'Qty hasil hitung tidak valid.'
    };
  }
  const facInfo = requireUserFacility_(namaPetugas);
  if (!facInfo) {
    return {
      success: false,
      message: PESAN_TANPA_FACILITY_
    };
  }
  if (!facilityKlienCocok_(facKlien, facInfo)) {
    return {
      success: false,
      kode: 'FACILITY_BERUBAH',
      message: PESAN_FACILITY_BERUBAH_
    };
  }
  try {
    queueSubmitPayload_(no, namaPetugas, aktual, facInfo.id, waktuKlienSah_(waktuKlien));
  } catch (e) {
    return {
      success: false,
      message: 'Gagal masuk antrian submit: ' + e.message
    };
  }
  return {
    success: true,
    queued: true,
    message: 'Submission diterima.'
  };
}

function appendRiwayatCycle_(lokasi, article, description, qtyTransaksi, qtySystem, namaPetugas, qtyCount, selisih, hasilAwal, statusValidasi, hasilFinal, waktu, assignedValidator, addWhoTransaksi) {
  const sheet = getRiwayatSheet_(namaPetugas);
  const id = Utilities.getUuid();
  const tanggal = Utilities.formatDate(waktu, 'Asia/Jakarta', 'yyyy-MM-dd');
  const statusTaskAwal = hasilAwal === 'HIT' ? 'Selesai' : '';
  const newRow = [ id, tanggal, lokasi, article, description, qtyTransaksi, qtySystem, namaPetugas, qtyCount, selisih, hasilAwal, statusValidasi, '', '', hasilFinal, waktu, '', '', statusTaskAwal, '', '', '', assignedValidator || '', '', waktu, addWhoTransaksi || '', '' ];
  sheet.appendRow(newRow);
  const rowIndex = sheet.getLastRow();
  riwayatRowCachePut_(id, rowIndex);
  try {
    queueRingkasanDeltaNewRow_(rowDateTag_(newRow), newRow, namaPetugas);
  } catch (e) {}
  try {
    if (newRow[11] === 'Pending' && newRow[22]) {
      queueCounterIncrements_([ {
        key: 'pending_total',
        delta: 1
      }, {
        key: 'pending_validator:' + newRow[22],
        delta: 1
      } ], namaPetugas);
    }
  } catch (e) {}
}

function getPendingBacklog(requesterUsername) {
  requireRole_(requesterUsername, [ 'admin', 'developer' ]);
  const sheet = getSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const map = {};
  data.forEach(function(row) {
    const tgl = String(row[1]);
    const status = row[10];
    if (!map[tgl]) map[tgl] = {
      total: 0,
      pending: 0
    };
    map[tgl].total++;
    if (status === 'Pending') map[tgl].pending++;
  });
  return Object.keys(map).filter(function(k) {
    return map[k].pending > 0;
  }).sort().map(function(k) {
    return {
      tanggal: k,
      pending: map[k].pending,
      total: map[k].total
    };
  });
}

function getBacklogDetailByDate(tanggal, requesterUsername) {
  requireRole_(requesterUsername, [ 'admin', 'developer' ]);
  const sheet = getSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const tglTarget = String(tanggal);
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const map = {};
  data.forEach(function(row) {
    if (String(row[1]) !== tglTarget) return;
    const namaPetugas = String(row[9] || '').trim() || '(belum ditugaskan)';
    const status = row[10];
    if (!map[namaPetugas]) map[namaPetugas] = {
      namaPetugas: namaPetugas,
      total: 0,
      pending: 0
    };
    map[namaPetugas].total++;
    if (status === 'Pending') map[namaPetugas].pending++;
  });
  return Object.keys(map).map(function(k) {
    return map[k];
  }).filter(function(u) {
    return u.pending > 0;
  }).sort(function(a, b) {
    return b.pending - a.pending;
  });
}

function readRingkasanHarianRows_(periodType, periodValue, username) {
  const sheet = getRingkasanHarianSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, RINGKASAN_HARIAN_HEADERS.length).getValues();
  return data.filter(function(row) {
    const tgl = String(row[0] || '');
    return periodType === 'monthly' ? tgl.substring(0, 7) === periodValue : tgl === periodValue;
  });
}

function computeSummaryFromRingkasan_(rows, periodValue) {
  let kpiTotal = 0, kpiHit = 0, kpiDisc = 0;
  const acc = {}, accValidator = {};
  rows.forEach(function(r) {
    kpiTotal += Number(r[1]) || 0;
    kpiHit += Number(r[2]) || 0;
    kpiDisc += Number(r[3]) || 0;
    const petugasMap = parseJsonSafe_(r[4], {});
    Object.keys(petugasMap).forEach(function(nama) {
      if (!acc[nama]) acc[nama] = {
        nama: nama,
        total: 0,
        hit: 0
      };
      acc[nama].total += petugasMap[nama].total || 0;
      acc[nama].hit += petugasMap[nama].hit || 0;
    });
    const validatorMap = parseJsonSafe_(r[5], {});
    Object.keys(validatorMap).forEach(function(nama) {
      if (!accValidator[nama]) accValidator[nama] = {
        nama: nama,
        total: 0,
        hit: 0
      };
      accValidator[nama].total += validatorMap[nama].total || 0;
      accValidator[nama].hit += validatorMap[nama].hit || 0;
    });
  });
  function toLeaderboard(a) {
    return Object.keys(a).map(function(k) {
      const p = a[k];
      const akurasi = p.total ? Math.round(p.hit / p.total * 1e3) / 10 : 0;
      return {
        nama: p.nama,
        total: p.total,
        hit: p.hit,
        akurasi: akurasi,
        persenDiscrepancy: p.total ? Math.round((p.total - p.hit) / p.total * 1e3) / 10 : 0
      };
    }).sort(function(x, y) {
      return y.akurasi - x.akurasi;
    });
  }
  return {
    period: periodValue,
    leaderboard: toLeaderboard(acc),
    leaderboardValidator: toLeaderboard(accValidator),
    kpi: {
      total: kpiTotal,
      hit: kpiHit,
      discrepancy: kpiDisc
    }
  };
}

function computeErrorAnalysisFromRingkasan_(rows) {
  const counts = {};
  let totalWithAlasan = 0;
  rows.forEach(function(r) {
    const kategoriMap = parseJsonSafe_(r[6], {});
    Object.keys(kategoriMap).forEach(function(k) {
      counts[k] = (counts[k] || 0) + (kategoriMap[k] || 0);
      totalWithAlasan += kategoriMap[k] || 0;
    });
  });
  return Object.keys(counts).map(function(k) {
    return {
      alasan: k,
      jumlah: counts[k],
      persen: totalWithAlasan ? Math.round(counts[k] / totalWithAlasan * 1e3) / 10 : 0
    };
  }).sort(function(a, b) {
    return b.jumlah - a.jumlah;
  });
}

function fillTrendFromRingkasan_(periodType, acc, username) {
  const sheet = getRingkasanHarianSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const data = sheet.getRange(2, 1, lastRow - 1, RINGKASAN_HARIAN_HEADERS.length).getValues();
  data.forEach(function(r) {
    const tgl = String(r[0] || '');
    const key = periodType === 'monthly' ? tgl.substring(0, 7) : tgl;
    const bucket = acc[key];
    if (!bucket) return;
    bucket.total += Number(r[1]) || 0;
    bucket.hit += Number(r[2]) || 0;
    bucket.discrepancy += Number(r[3]) || 0;
  });
}

function buildTrendSkeleton_(periodType, count) {
  const n = count ? Number(count) : periodType === 'monthly' ? 6 : 7;
  const now = new Date;
  const periods = [];
  const acc = {};
  for (let i = n - 1; i >= 0; i--) {
    let key;
    if (periodType === 'monthly') {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      key = Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy-MM');
    } else {
      const d = new Date(now.getTime() - i * 864e5);
      key = Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy-MM-dd');
    }
    periods.push(key);
    acc[key] = {
      total: 0,
      hit: 0,
      discrepancy: 0
    };
  }
  return {
    periods: periods,
    acc: acc
  };
}

function fillTrendFromData_(data, periodType, acc) {
  data.forEach(function(row) {
    const tgl = rowDateTag_(row);
    const key = periodType === 'monthly' ? tgl.substring(0, 7) : tgl;
    const bucket = acc[key];
    if (!bucket) return;
    const statusValidasi = row[11], hasilAwal = row[10], hasilFinal = row[14];
    const effectiveHasil = statusValidasi === 'Pending' ? 'PENDING' : statusValidasi === 'Tidak Perlu' ? hasilAwal : hasilFinal;
    bucket.total++;
    if (effectiveHasil === 'HIT') bucket.hit++; else if (effectiveHasil === 'DISCREPANCY') bucket.discrepancy++;
  });
}

function finalizeTrend_(periods, acc) {
  return periods.map(function(p) {
    const b = acc[p];
    return {
      period: p,
      total: b.total,
      hit: b.hit,
      discrepancy: b.discrepancy,
      akurasi: b.total ? Math.round(b.hit / b.total * 1e3) / 10 : 0
    };
  });
}

function getTrendData(periodType, count, requesterUsername) {
  requireRole_(requesterUsername, [ 'admin' ]);
  const skeleton = buildTrendSkeleton_(periodType, count);
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
    fillTrendFromData_(data, periodType, skeleton.acc);
  }
  return finalizeTrend_(skeleton.periods, skeleton.acc);
}

function computeSummaryFromData_(data, periodType, periodValue) {
  const details = [];
  const acc = {};
  const accValidator = {};
  let kpiTotal = 0, kpiHit = 0, kpiDisc = 0;
  data.forEach(function(row) {
    const tgl = rowDateTag_(row);
    const match = periodType === 'monthly' ? tgl.substring(0, 7) === periodValue : tgl === periodValue;
    if (!match) return;
    const lokasi = row[2], article = row[3], description = row[4], namaPetugas = row[7];
    const selisih = row[9], hasilAwal = row[10], statusValidasi = row[11], hasilFinal = row[14], statusTask = row[18];
    const namaValidator = row[12], kategoriSelisih = row[17];
    const effectiveHasil = statusValidasi === 'Pending' ? 'PENDING' : statusValidasi === 'Tidak Perlu' ? hasilAwal : hasilFinal;
    kpiTotal++;
    if (effectiveHasil === 'HIT') kpiHit++; else if (effectiveHasil === 'DISCREPANCY') kpiDisc++;
    if (effectiveHasil !== 'HIT') {
      let statusLabel;
      if (effectiveHasil === 'PENDING') statusLabel = 'Menunggu Validasi'; else if (statusTask === 'Open') statusLabel = 'Verifikasi (Investigasi)'; else if (statusTask === 'Selesai') statusLabel = 'Selesai (Ditindaklanjuti)'; else statusLabel = 'Discrepancy';
      details.push({
        tanggal: tgl,
        lokasi: lokasi,
        article: article,
        description: description,
        namaPetugas: namaPetugas,
        selisih: selisih,
        status: statusLabel
      });
    }
    if (statusValidasi === 'Tidak Perlu') {
      if (!acc[namaPetugas]) acc[namaPetugas] = {
        nama: namaPetugas,
        total: 0,
        hit: 0
      };
      acc[namaPetugas].total++;
      acc[namaPetugas].hit++;
    } else if (statusValidasi === 'Selesai') {
      if (!acc[namaPetugas]) acc[namaPetugas] = {
        nama: namaPetugas,
        total: 0,
        hit: 0
      };
      acc[namaPetugas].total++;
      if (hasilAwal === hasilFinal) acc[namaPetugas].hit++;
    }
    if (namaValidator && statusTask === 'Selesai') {
      if (!accValidator[namaValidator]) accValidator[namaValidator] = {
        nama: namaValidator,
        total: 0,
        hit: 0
      };
      accValidator[namaValidator].total++;
      const validatorSalah = hasilFinal === 'DISCREPANCY' && kategoriSelisih === 'Salah Hitung';
      if (!validatorSalah) accValidator[namaValidator].hit++;
    }
  });
  const leaderboard = Object.keys(acc).map(function(k) {
    const p = acc[k];
    const akurasi = p.total ? Math.round(p.hit / p.total * 1e3) / 10 : 0;
    return {
      nama: p.nama,
      total: p.total,
      hit: p.hit,
      akurasi: akurasi,
      persenDiscrepancy: p.total ? Math.round((p.total - p.hit) / p.total * 1e3) / 10 : 0
    };
  }).sort(function(a, b) {
    return b.akurasi - a.akurasi;
  });
  const leaderboardValidator = Object.keys(accValidator).map(function(k) {
    const p = accValidator[k];
    const akurasi = p.total ? Math.round(p.hit / p.total * 1e3) / 10 : 0;
    return {
      nama: p.nama,
      total: p.total,
      hit: p.hit,
      akurasi: akurasi,
      persenDiscrepancy: p.total ? Math.round((p.total - p.hit) / p.total * 1e3) / 10 : 0
    };
  }).sort(function(a, b) {
    return b.akurasi - a.akurasi;
  });
  return {
    period: periodValue,
    leaderboard: leaderboard,
    leaderboardValidator: leaderboardValidator,
    details: details,
    kpi: {
      total: kpiTotal,
      hit: kpiHit,
      discrepancy: kpiDisc
    }
  };
}

function getSummaryByPeriod(periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, [ 'admin' ]);
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];
  return computeSummaryFromData_(data, periodType, periodValue);
}

function computeUserSummaryFromData_(data, nama, periodType, periodValue) {
  let total = 0, hit = 0;
  data.forEach(function(row) {
    if (String(row[7]) !== String(nama)) return;
    const statusValidasi = row[11];
    if (statusValidasi === 'Pending') return;
    const tgl = rowDateTag_(row);
    const match = periodType === 'monthly' ? tgl.substring(0, 7) === periodValue : tgl === periodValue;
    if (!match) return;
    const hasilAwal = row[10], hasilFinal = row[14];
    const petugasBenar = statusValidasi === 'Tidak Perlu' ? true : hasilAwal === hasilFinal;
    total++;
    if (petugasBenar) hit++;
  });
  const akurasi = total ? Math.round(hit / total * 1e3) / 10 : 0;
  return {
    nama: nama,
    total: total,
    hit: hit,
    kesalahanHitung: total - hit,
    akurasi: akurasi,
    persenHit: akurasi,
    persenDiscrepancy: total ? Math.round((total - hit) / total * 1e3) / 10 : 0
  };
}

function getUserSummary(nama, periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, [ 'admin' ]);
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];
  return computeUserSummaryFromData_(data, nama, periodType, periodValue);
}

function computeErrorAnalysisFromData_(data, periodType, periodValue) {
  const counts = {};
  let totalWithAlasan = 0;
  data.forEach(function(row) {
    const alasan = String(row[17] || '').trim();
    if (!alasan) return;
    const tgl = rowDateTag_(row);
    const match = periodType === 'monthly' ? tgl.substring(0, 7) === periodValue : tgl === periodValue;
    if (!match) return;
    counts[alasan] = (counts[alasan] || 0) + 1;
    totalWithAlasan++;
  });
  return Object.keys(counts).map(function(k) {
    return {
      alasan: k,
      jumlah: counts[k],
      persen: totalWithAlasan ? Math.round(counts[k] / totalWithAlasan * 1e3) / 10 : 0
    };
  }).sort(function(a, b) {
    return b.jumlah - a.jumlah;
  });
}

function getErrorAnalysis(periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, [ 'admin' ]);
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];
  return computeErrorAnalysisFromData_(data, periodType, periodValue);
}

function getErrorAnalysisDetail(alasan, periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, [ 'admin', 'developer' ]);
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  const result = [];
  data.forEach(function(row) {
    const rowAlasan = String(row[17] || '').trim();
    if (rowAlasan !== alasan) return;
    const tgl = rowDateTag_(row);
    const match = periodType === 'monthly' ? tgl.substring(0, 7) === periodValue : tgl === periodValue;
    if (!match) return;
    result.push({
      lokasi: row[2],
      article: row[3],
      description: row[4],
      qtySystem: row[6],
      qtyCount: row[8],
      namaPetugas: row[7]
    });
  });
  return result;
}

function computeUserInvestigasiFromData_(data, nama, periodType, periodValue) {
  let totalSelesai = 0, totalDurasiJam = 0, withDurasi = 0;
  data.forEach(function(row) {
    if (row[18] !== 'Selesai') return;
    if (String(row[20] || '').trim() !== String(nama).trim()) return;
    const waktuSelesai = row[21];
    if (Object.prototype.toString.call(waktuSelesai) !== '[object Date]') return;
    const tgl = Utilities.formatDate(waktuSelesai, 'Asia/Jakarta', 'yyyy-MM-dd');
    const match = periodType === 'monthly' ? tgl.substring(0, 7) === periodValue : tgl === periodValue;
    if (!match) return;
    totalSelesai++;
    const waktuCycle = row[15];
    if (Object.prototype.toString.call(waktuCycle) === '[object Date]') {
      totalDurasiJam += (waktuSelesai - waktuCycle) / 36e5;
      withDurasi++;
    }
  });
  return {
    totalSelesai: totalSelesai,
    avgDurasiJam: withDurasi ? Math.round(totalDurasiJam / withDurasi * 10) / 10 : 0
  };
}

function getUserInvestigasiSummary(nama, periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, [ 'admin' ]);
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];
  return computeUserInvestigasiFromData_(data, nama, periodType, periodValue);
}

function getDashboardData(periodType, periodValue, trendCount, requesterUsername) {
  requireRole_(requesterUsername, [ 'admin', 'developer' ]);
  const rows = readRingkasanHarianRows_(periodType, periodValue, requesterUsername);
  const summary = computeSummaryFromRingkasan_(rows, periodValue);
  const errorAnalysis = computeErrorAnalysisFromRingkasan_(rows);
  const skeleton = buildTrendSkeleton_(periodType, trendCount);
  fillTrendFromRingkasan_(periodType, skeleton.acc, requesterUsername);
  const trend = finalizeTrend_(skeleton.periods, skeleton.acc);
  return {
    summary: summary,
    errorAnalysis: errorAnalysis,
    trend: trend
  };
}

function getProblemItemsDetail(periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, [ 'admin', 'developer' ]);
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  const details = [];
  data.forEach(function(row) {
    const tgl = rowDateTag_(row);
    const match = periodType === 'monthly' ? tgl.substring(0, 7) === periodValue : tgl === periodValue;
    if (!match) return;
    const selisih = row[9], hasilAwal = row[10], statusValidasi = row[11], hasilFinal = row[14], statusTask = row[18];
    const effectiveHasil = statusValidasi === 'Pending' ? 'PENDING' : statusValidasi === 'Tidak Perlu' ? hasilAwal : hasilFinal;
    if (effectiveHasil === 'HIT') return;
    let statusLabel;
    if (effectiveHasil === 'PENDING') statusLabel = 'Menunggu Validasi'; else if (statusTask === 'Open') statusLabel = 'Verifikasi (Investigasi)'; else if (statusTask === 'Selesai') statusLabel = 'Selesai (Ditindaklanjuti)'; else statusLabel = 'Discrepancy';
    details.push({
      tanggal: tgl,
      lokasi: row[2],
      article: row[3],
      description: row[4],
      namaPetugas: row[7],
      selisih: selisih,
      status: statusLabel
    });
  });
  return details;
}

function getUserDashboardDetail(nama, periodType, periodValue, requesterUsername) {
  requireRole_(requesterUsername, [ 'admin', 'developer' ]);
  const rows = readRingkasanHarianRows_(periodType, periodValue, requesterUsername);
  let total = 0, hit = 0;
  rows.forEach(function(r) {
    const petugasMap = parseJsonSafe_(r[4], {});
    const p = petugasMap[nama];
    if (p) {
      total += p.total || 0;
      hit += p.hit || 0;
    }
  });
  const akurasi = total ? Math.round(hit / total * 1e3) / 10 : 0;
  const summary = {
    nama: nama,
    total: total,
    hit: hit,
    kesalahanHitung: total - hit,
    akurasi: akurasi,
    persenHit: akurasi,
    persenDiscrepancy: total ? Math.round((total - hit) / total * 1e3) / 10 : 0
  };
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues() : [];
  const investigasi = computeUserInvestigasiFromData_(data, nama, periodType, periodValue);
  return {
    summary: summary,
    investigasi: investigasi
  };
}

const SCRIPT_PROP_MASTER_DB_ID = 'MASTER_DB_ID';

const SCRIPT_PROP_CONFIG_DB_ID = 'CONFIG_DB_ID';

const LOG_SISTEM_SHEET_NAME = 'Log_Sistem';

const LOG_SISTEM_HEADERS = [ 'Waktu', 'Sumber', 'Pesan' ];

function catatLogSistem_(sumber, pesan) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(LOG_SISTEM_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_SISTEM_SHEET_NAME);
      sheet.appendRow(LOG_SISTEM_HEADERS);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([ new Date, sumber, pesan ]);
  } catch (e) {}
}

function tryOpenMasterDb_() {
  const id = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_MASTER_DB_ID);
  if (!id) return null;
  const cache = CacheService.getScriptCache();
  if (cache.get('masterDbDown')) return null;
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    try {
      cache.put('masterDbDown', '1', 60);
    } catch (e2) {}
    catatLogSistem_('Master DB', 'Gagal membuka Master DB (' + e.message + '). Fallback ke Master sheet lokal.');
    return null;
  }
}

function getMasterSheetWithFallback_() {
  const masterDb = tryOpenMasterDb_();
  if (masterDb) {
    const sheet = masterDb.getSheetByName(MASTER_SHEET_NAME);
    if (sheet) return sheet;
    catatLogSistem_('Master DB', 'Master DB terbuka tapi sheet "Master" tidak ditemukan di dalamnya. Fallback ke Master sheet lokal.');
  }
  return getLegacyMasterSheetLocal_();
}

function getLegacyMasterSheetLocal_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(MASTER_SHEET_NAME);
}

function tryOpenConfigDb_() {
  const id = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_CONFIG_DB_ID);
  if (!id) return null;
  const cache = CacheService.getScriptCache();
  if (cache.get('configDbDown')) return null;
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    try {
      cache.put('configDbDown', '1', 60);
    } catch (e2) {}
    catatLogSistem_('Config DB', 'Gagal membuka Config DB (' + e.message + '). Fallback ke sheet config lokal.');
    return null;
  }
}

function getConfigTargetSpreadsheet_() {
  const configDb = tryOpenConfigDb_();
  return configDb || SpreadsheetApp.getActiveSpreadsheet();
}

const LOG_PERUBAHAN_CONFIG_SHEET_NAME = 'Log_Perubahan_Config';

const LOG_PERUBAHAN_CONFIG_HEADERS = [ 'Waktu', 'NIK', 'Sumber', 'Detail' ];

function getLogPerubahanConfigSheet_() {
  const ss = getConfigTargetSpreadsheet_();
  let sheet = ss.getSheetByName(LOG_PERUBAHAN_CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_PERUBAHAN_CONFIG_SHEET_NAME);
    sheet.appendRow(LOG_PERUBAHAN_CONFIG_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function catatLogPerubahanConfig_(nik, sumber, detail) {
  try {
    getLogPerubahanConfigSheet_().appendRow([ new Date, nik || '(tidak diketahui)', sumber, detail ]);
  } catch (e) {}
}

function setupExternalDbMasterDanConfig() {
  const props = PropertiesService.getScriptProperties();
  const hasil = [];
  let masterDbId = props.getProperty(SCRIPT_PROP_MASTER_DB_ID);
  let masterDb;
  if (masterDbId) {
    try {
      masterDb = SpreadsheetApp.openById(masterDbId);
    } catch (e) {
      masterDb = null;
    }
  }
  if (!masterDb) {
    masterDb = SpreadsheetApp.create('Master DB - Cycle Count DC');
    props.setProperty(SCRIPT_PROP_MASTER_DB_ID, masterDb.getId());
    hasil.push('Master DB baru dibuat: ' + masterDb.getUrl());
  } else {
    hasil.push('Master DB sudah ada, dipakai lagi: ' + masterDb.getUrl());
  }
  let masterSheetBaru = masterDb.getSheetByName(MASTER_SHEET_NAME);
  const masterLokal = getLegacyMasterSheetLocal_();
  if (!masterSheetBaru && masterLokal) {
    masterSheetBaru = masterLokal.copyTo(masterDb);
    masterSheetBaru.setName(MASTER_SHEET_NAME);
    const sheet1 = masterDb.getSheetByName('Sheet1');
    if (sheet1 && masterDb.getSheets().length > 1) masterDb.deleteSheet(sheet1);
    hasil.push('Data sheet "Master" (' + (masterLokal.getLastRow() - 1) + ' baris) berhasil dipindah ke Master DB.');
  } else if (masterSheetBaru) {
    hasil.push('Sheet "Master" sudah ada di Master DB, tidak ditimpa (data existing di Master DB dipertahankan).');
  } else {
    hasil.push('PERINGATAN: sheet "Master" lokal tidak ditemukan, Master DB dibuat kosong.');
  }
  let configDbId = props.getProperty(SCRIPT_PROP_CONFIG_DB_ID);
  let configDb;
  if (configDbId) {
    try {
      configDb = SpreadsheetApp.openById(configDbId);
    } catch (e) {
      configDb = null;
    }
  }
  if (!configDb) {
    configDb = SpreadsheetApp.create('Config DB - Cycle Count DC');
    props.setProperty(SCRIPT_PROP_CONFIG_DB_ID, configDb.getId());
    hasil.push('Config DB baru dibuat: ' + configDb.getUrl());
  } else {
    hasil.push('Config DB sudah ada, dipakai lagi: ' + configDb.getUrl());
  }
  [ CONFIG_SISTEM_SHEET_NAME, CONFIG_AKSES_SETTING_SHEET_NAME ].forEach(function(namaSheet) {
    let target = configDb.getSheetByName(namaSheet);
    const lokal = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet);
    if (!target && lokal) {
      target = lokal.copyTo(configDb);
      target.setName(namaSheet);
      hasil.push('Sheet "' + namaSheet + '" berhasil dipindah ke Config DB.');
    } else if (target) {
      hasil.push('Sheet "' + namaSheet + '" sudah ada di Config DB, tidak ditimpa.');
    }
  });
  const configSheet1 = configDb.getSheetByName('Sheet1');
  if (configSheet1 && configDb.getSheets().length > 1) configDb.deleteSheet(configSheet1);
  getConfigSistemSheet_();
  getConfigAksesSettingSheet_();
  const pesan = hasil.join('\n');
  try {
    SpreadsheetApp.getUi().alert('Setup Master DB & Config DB selesai:\n\n' + pesan);
  } catch (e) {
    Logger.log(pesan);
  }
}

const MASTER_USER_SHEET_NAME = 'Master_User';

const MASTER_USER_HEADERS = [ 'Username', 'Role', 'Status', 'ID_Facility' ];

function getMasterUserSheetWithFallback_() {
  const masterDb = tryOpenMasterDb_();
  let sheet = null;
  if (masterDb) {
    sheet = masterDb.getSheetByName(MASTER_USER_SHEET_NAME);
  }
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MASTER_USER_SHEET_NAME);
  }
  return sheet;
}

function migrasiMasterKeSheetBaru() {
  const masterSheet = getMasterSheetWithFallback_();
  const assignSheet = getUserFacilityAssignmentSheet_();
  const masterDb = tryOpenMasterDb_() || SpreadsheetApp.getActiveSpreadsheet();
  let newUserSheet = masterDb.getSheetByName(MASTER_USER_SHEET_NAME);
  if (!newUserSheet) {
    newUserSheet = masterDb.insertSheet(MASTER_USER_SHEET_NAME);
    newUserSheet.appendRow(MASTER_USER_HEADERS);
    newUserSheet.setFrozenRows(1);
  }
  const assignments = {};
  if (assignSheet && assignSheet.getLastRow() >= 2) {
    const data = assignSheet.getRange(2, 1, assignSheet.getLastRow() - 1, 2).getValues();
    data.forEach(function(r) {
      const uname = String(r[0] || '').trim().toLowerCase();
      if (uname) assignments[uname] = String(r[1] || '').trim();
    });
  }
  if (masterSheet && masterSheet.getLastRow() >= 2) {
    const data = masterSheet.getRange(2, 3, masterSheet.getLastRow() - 1, 3).getValues();
    const rows = [];
    data.forEach(function(r) {
      const uname = String(r[0] || '').trim();
      const role = String(r[1] || '').trim();
      const status = String(r[2] || '').trim();
      if (uname) {
        const facId = assignments[uname.toLowerCase()] || '';
        rows.push([ uname, role, status || 'Aktif', facId ]);
      }
    });
    if (rows.length > 0) {
      const existingLastRow = newUserSheet.getLastRow();
      if (existingLastRow >= 2) {
        newUserSheet.getRange(2, 1, existingLastRow - 1, 4).clearContent();
      }
      newUserSheet.getRange(2, 1, rows.length, 4).setValues(rows);
    }
  }
  return "Migrasi berhasil dilakukan ke sheet Master_User. Silakan refresh web app.";
}

const FACILITY_SHEET_NAME = 'Facility';

const FACILITY_HEADERS = [ 'ID_Facility', 'Nama_Facility', 'Kode_Facility', 'Nama_Spreadsheet', 'Spreadsheet_ID', 'Status', 'Tanggal_Dibuat', 'Dibuat_Oleh', 'URL_Spreadsheet' ];

const USER_FACILITY_ASSIGNMENT_SHEET_NAME = 'User_Facility_Assignment';

const USER_FACILITY_ASSIGNMENT_HEADERS = [ 'Username', 'ID_Facility', 'Tanggal_Diassign', 'Diassign_Oleh' ];

const LOKASI_AKTIF_SHEET_NAME = 'Lokasi_Aktif';

const LOKASI_AKTIF_HEADERS = [ 'Lokasi' ];

const FACILITY_OPERATIONAL_SHEETS = [ {
  name: SHEET_NAME,
  headers: HEADERS
}, {
  name: RIWAYAT_SHEET_NAME,
  headers: RIWAYAT_HEADERS
}, {
  name: LOG_AKSES_SHEET_NAME,
  headers: LOG_AKSES_HEADERS
}, {
  name: LOG_ANOMALI_SHEET_NAME,
  headers: LOG_ANOMALI_HEADERS
}, {
  name: LOG_UNASSIGNED_SHEET_NAME,
  headers: LOG_UNASSIGNED_HEADERS
}, {
  name: LOG_BUKTI_SHEET_NAME,
  headers: LOG_BUKTI_HEADERS
}, {
  name: RINGKASAN_HARIAN_SHEET_NAME,
  headers: RINGKASAN_HARIAN_HEADERS
}, {
  name: ANTRIAN_AKTIF_SHEET_NAME,
  headers: ANTRIAN_AKTIF_HEADERS
}, {
  name: RIWAYAT_ARCHIVE_SHEET_NAME,
  headers: RIWAYAT_HEADERS
}, {
  name: LOG_SISTEM_SHEET_NAME,
  headers: LOG_SISTEM_HEADERS
}, {
  name: LOKASI_AKTIF_SHEET_NAME,
  headers: LOKASI_AKTIF_HEADERS
}, {
  name: QUEUE_RINGKASAN_SHEET_NAME,
  headers: QUEUE_RINGKASAN_HEADERS
}, {
  name: QUEUE_COUNTER_SHEET_NAME,
  headers: QUEUE_COUNTER_HEADERS
} ];

const CACHE_FACILITY_LIST = 'facilityList';

const CACHE_USER_FACILITY_PREFIX = 'userFacility_';

const CACHE_DEV_ACTIVE_FACILITY_PREFIX = 'devActiveFacility_';

const DEV_ACTIVE_FACILITY_TTL_SEC = 21600;

function getFacilitySheet_() {
  const ss = getConfigTargetSpreadsheet_();
  let sheet = ss.getSheetByName(FACILITY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FACILITY_SHEET_NAME);
    sheet.appendRow(FACILITY_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, FACILITY_HEADERS);
  }
  return sheet;
}

function getUserFacilityAssignmentSheet_() {
  const ss = getConfigTargetSpreadsheet_();
  let sheet = ss.getSheetByName(USER_FACILITY_ASSIGNMENT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USER_FACILITY_ASSIGNMENT_SHEET_NAME);
    sheet.appendRow(USER_FACILITY_ASSIGNMENT_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, USER_FACILITY_ASSIGNMENT_HEADERS);
  }
  return sheet;
}

function clearFacilityCache_() {
  _memUserFacility_ = {};
  try {
    CacheService.getScriptCache().remove(CACHE_FACILITY_LIST);
  } catch (e) {}
}

function clearUserFacilityCache_(username) {
  delete _memUserFacility_[String(username || '').trim().toLowerCase()];
  try {
    CacheService.getScriptCache().remove(CACHE_USER_FACILITY_PREFIX + String(username || '').toLowerCase());
  } catch (e) {}
}

function getDaftarFacility(username) {
  if (!isNikPunyaAksesSetting_(username)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Facility Management.'
    };
  }
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_FACILITY_LIST);
  if (cached) {
    return {
      success: true,
      facilities: JSON.parse(cached)
    };
  }
  const sheet = getFacilitySheet_();
  const lastRow = sheet.getLastRow();
  const facilities = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, FACILITY_HEADERS.length).getValues();
    values.forEach(function(r) {
      const id = String(r[0] || '').trim();
      if (!id) return;
      facilities.push({
        id: id,
        nama: String(r[1] || '').trim(),
        kode: String(r[2] || '').trim(),
        namaSpreadsheet: String(r[3] || '').trim(),
        spreadsheetId: String(r[4] || '').trim(),
        status: String(r[5] || '').trim(),
        tanggalDibuat: r[6] instanceof Date ? Utilities.formatDate(r[6], 'Asia/Jakarta', 'yyyy-MM-dd HH:mm') : String(r[6] || ''),
        dibuatOleh: String(r[7] || '').trim(),
        url: String(r[8] || '').trim()
      });
    });
  }
  try {
    cache.put(CACHE_FACILITY_LIST, JSON.stringify(facilities), 300);
  } catch (e) {}
  return {
    success: true,
    facilities: facilities
  };
}

function buatSpreadsheetFacility_(namaSpreadsheet) {
  const ss = SpreadsheetApp.create(namaSpreadsheet);
  const sheet1 = ss.getSheetByName('Sheet1');
  FACILITY_OPERATIONAL_SHEETS.forEach(function(def) {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
    }
    sheet.appendRow(def.headers);
    sheet.setFrozenRows(1);
  });
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }
  terapkanFormatKolomTeks_(ss);
  return {
    spreadsheetId: ss.getId(),
    url: ss.getUrl()
  };
}

function tambahFacility(username, namaFacility, kodeFacility, namaSpreadsheet, daftarLokasi, jobId) {
  const info = getUserRole(username);
  if (!info || info.role !== 'developer') {
    return {
      success: false,
      message: 'Hanya role Developer yang boleh menambah facility baru.'
    };
  }
  reportProgress_(jobId, 5, 'Validasi input');
  const nama = String(namaFacility || '').trim();
  const kode = String(kodeFacility || '').trim();
  const namaSs = String(namaSpreadsheet || '').trim();
  if (!nama) return {
    success: false,
    message: 'Nama Facility tidak boleh kosong.'
  };
  if (!kode) return {
    success: false,
    message: 'Kode Facility tidak boleh kosong.'
  };
  if (!namaSs) return {
    success: false,
    message: 'Nama Spreadsheet tidak boleh kosong.'
  };
  if (!daftarLokasi || !Array.isArray(daftarLokasi) || daftarLokasi.length === 0) {
    return {
      success: false,
      message: 'Daftar lokasi aktif wajib diisi. Minimal 1 lokasi.'
    };
  }
  reportProgress_(jobId, 15, 'Cek duplikasi nama/kode facility');
  const sheet = getFacilitySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < existing.length; i++) {
      if (String(existing[i][1] || '').trim().toLowerCase() === nama.toLowerCase()) {
        return {
          success: false,
          message: 'Nama Facility "' + nama + '" sudah ada. Gunakan nama lain.'
        };
      }
      if (String(existing[i][2] || '').trim().toLowerCase() === kode.toLowerCase()) {
        return {
          success: false,
          message: 'Kode Facility "' + kode + '" sudah ada. Gunakan kode lain.'
        };
      }
    }
  }
  const idFacility = 'FAC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  reportProgress_(jobId, 25, 'Membuat spreadsheet facility baru');
  const ssBaru = buatSpreadsheetFacility_(namaSs);
  reportProgress_(jobId, 60, 'Menulis daftar lokasi aktif');
  const facilitySs = SpreadsheetApp.openById(ssBaru.spreadsheetId);
  const lokasiSheet = facilitySs.getSheetByName(LOKASI_AKTIF_SHEET_NAME);
  const lokasiUnik = {};
  const lokasiBersih = [];
  if (lokasiSheet && daftarLokasi.length > 0) {
    daftarLokasi.forEach(function(lok) {
      const l = String(lok || '').trim();
      if (l && !lokasiUnik[l]) {
        lokasiUnik[l] = true;
        lokasiBersih.push([ l ]);
      }
    });
    if (lokasiBersih.length > 0) {
      lokasiSheet.getRange(2, 1, lokasiBersih.length, 1).setValues(lokasiBersih);
    }
  }
  reportProgress_(jobId, 85, 'Menyimpan info facility');
  const targetRow = Math.max(lastRow + 1, 2);
  const now = new Date;
  sheet.getRange(targetRow, 1).setValue(idFacility);
  sheet.getRange(targetRow, 2).setValue(nama);
  sheet.getRange(targetRow, 3).setValue(kode);
  sheet.getRange(targetRow, 4).setValue(namaSs);
  sheet.getRange(targetRow, 5).setValue(ssBaru.spreadsheetId);
  sheet.getRange(targetRow, 6).setValue('Aktif');
  sheet.getRange(targetRow, 7).setValue(now);
  sheet.getRange(targetRow, 8).setValue(username);
  sheet.getRange(targetRow, 9).setValue(ssBaru.url);
  clearFacilityCache_();
  catatLogPerubahanConfig_(username, 'Facility Management', 'Tambah facility "' + nama + '" (Kode: ' + kode + ', ID: ' + idFacility + '). ' + 'Spreadsheet dibuat: ' + ssBaru.url + '. ' + 'Lokasi aktif diimpor: ' + (lokasiBersih ? lokasiBersih.length : daftarLokasi.length) + ' lokasi.');
  reportProgress_(jobId, 100, 'Selesai');
  return {
    success: true,
    message: 'Facility "' + nama + '" berhasil dibuat.',
    facility: {
      id: idFacility,
      nama: nama,
      kode: kode,
      spreadsheetId: ssBaru.spreadsheetId,
      url: ssBaru.url,
      jumlahLokasi: lokasiBersih ? lokasiBersih.length : daftarLokasi.length
    }
  };
}

function daftarkanFacilityExisting(username, namaFacility, kodeFacility, spreadsheetIdAtauUrl, daftarLokasi, jobId) {
  const info = getUserRole(username);
  if (!info || info.role !== 'developer') {
    return {
      success: false,
      message: 'Hanya role Developer yang boleh mendaftarkan facility baru.'
    };
  }
  reportProgress_(jobId, 5, 'Validasi input');
  const nama = String(namaFacility || '').trim();
  const kode = String(kodeFacility || '').trim();
  const ssId = extractSpreadsheetId_(spreadsheetIdAtauUrl);
  if (!nama) return {
    success: false,
    message: 'Nama Facility tidak boleh kosong.'
  };
  if (!kode) return {
    success: false,
    message: 'Kode Facility tidak boleh kosong.'
  };
  if (!ssId) return {
    success: false,
    message: 'Spreadsheet ID/URL tidak boleh kosong atau tidak valid.'
  };
  reportProgress_(jobId, 15, 'Cek duplikasi facility');
  const sheet = getFacilitySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    for (let i = 0; i < existing.length; i++) {
      if (String(existing[i][1] || '').trim().toLowerCase() === nama.toLowerCase()) {
        return {
          success: false,
          message: 'Nama Facility "' + nama + '" sudah ada. Gunakan nama lain.'
        };
      }
      if (String(existing[i][2] || '').trim().toLowerCase() === kode.toLowerCase()) {
        return {
          success: false,
          message: 'Kode Facility "' + kode + '" sudah ada. Gunakan kode lain.'
        };
      }
      if (String(existing[i][4] || '').trim() === ssId) {
        return {
          success: false,
          message: 'Spreadsheet ini sudah terdaftar sebagai facility "' + existing[i][1] + '". Tidak boleh 1 spreadsheet dipakai 2 facility.'
        };
      }
    }
  }
  reportProgress_(jobId, 30, 'Membuka spreadsheet existing');
  let facilitySs;
  try {
    facilitySs = SpreadsheetApp.openById(ssId);
  } catch (e) {
    return {
      success: false,
      message: 'Gagal membuka spreadsheet dengan ID/URL tersebut (' + e.message + '). Pastikan ID/URL-nya benar dan spreadsheet sudah di-share ke akun yang menjalankan script ini.'
    };
  }
  reportProgress_(jobId, 50, 'Menyiapkan sheet operasional (data lama tidak diubah)');
  FACILITY_OPERATIONAL_SHEETS.forEach(function(def) {
    let s = facilitySs.getSheetByName(def.name);
    if (!s) {
      s = facilitySs.insertSheet(def.name);
      s.appendRow(def.headers);
      s.setFrozenRows(1);
    } else {
      ensureHeaderColumns_(s, def.headers);
    }
  });
  terapkanFormatKolomTeks_(facilitySs);
  reportProgress_(jobId, 70, 'Menyimpan info facility');
  const idFacility = 'FAC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const targetRow = Math.max(lastRow + 1, 2);
  const now = new Date;
  sheet.getRange(targetRow, 1, 1, FACILITY_HEADERS.length).setValues([ [ idFacility, nama, kode, facilitySs.getName(), ssId, 'Aktif', now, username, facilitySs.getUrl() ] ]);
  clearFacilityCache_();
  let jumlahLokasiDitambah = 0;
  if (daftarLokasi && Array.isArray(daftarLokasi) && daftarLokasi.length > 0) {
    reportProgress_(jobId, 85, 'Menambahkan lokasi aktif');
    const hasilLokasi = importLokasiAktif(username, idFacility, daftarLokasi, false);
    if (hasilLokasi && hasilLokasi.success) {
      jumlahLokasiDitambah = hasilLokasi.ditambahkan || 0;
    }
  }
  catatLogPerubahanConfig_(username, 'Facility Management', 'Daftarkan facility "' + nama + '" (Kode: ' + kode + ', ID: ' + idFacility + ') dari SPREADSHEET EXISTING: ' + facilitySs.getUrl() + ' (data lama di sheet Data Count/Riwayat TIDAK diubah). Lokasi baru ditambahkan: ' + jumlahLokasiDitambah + '.');
  reportProgress_(jobId, 100, 'Selesai');
  return {
    success: true,
    message: 'Facility "' + nama + '" berhasil didaftarkan dari spreadsheet existing. Data lama di sheet Data Count/Riwayat tetap seperti semula.',
    facility: {
      id: idFacility,
      nama: nama,
      kode: kode,
      spreadsheetId: ssId,
      url: facilitySs.getUrl(),
      lokasiBaruDitambah: jumlahLokasiDitambah
    }
  };
}

function updateNamaFacility(username, facilityId, namaBaru, kodeBaru) {
  if (!isNikPunyaAksesSetting_(username)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Facility Management.'
    };
  }
  const id = String(facilityId || '').trim();
  const nama = String(namaBaru || '').trim();
  const kode = String(kodeBaru || '').trim();
  if (!id) return {
    success: false,
    message: 'ID Facility tidak valid.'
  };
  if (!nama) return {
    success: false,
    message: 'Nama Facility tidak boleh kosong.'
  };
  const sheet = getFacilitySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {
    success: false,
    message: 'Tidak ada data facility.'
  };
  const values = sheet.getRange(2, 1, lastRow - 1, FACILITY_HEADERS.length).getValues();
  let targetIdx = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === id) {
      targetIdx = i;
      break;
    }
  }
  if (targetIdx === -1) return {
    success: false,
    message: 'Facility dengan ID tersebut tidak ditemukan.'
  };
  for (let i = 0; i < values.length; i++) {
    if (i === targetIdx) continue;
    if (String(values[i][1] || '').trim().toLowerCase() === nama.toLowerCase()) {
      return {
        success: false,
        message: 'Nama Facility "' + nama + '" sudah dipakai facility lain.'
      };
    }
    if (kode && String(values[i][2] || '').trim().toLowerCase() === kode.toLowerCase()) {
      return {
        success: false,
        message: 'Kode Facility "' + kode + '" sudah dipakai facility lain.'
      };
    }
  }
  const rowIndex = targetIdx + 2;
  const namaLama = String(values[targetIdx][1] || '').trim();
  sheet.getRange(rowIndex, 2).setValue(nama);
  if (kode) sheet.getRange(rowIndex, 3).setValue(kode);
  clearFacilityCache_();
  catatLogPerubahanConfig_(username, 'Facility Management', 'Ubah facility "' + namaLama + '" → "' + nama + '"' + (kode ? ' (Kode: ' + kode + ')' : '') + '.');
  return {
    success: true,
    message: 'Data facility berhasil diperbarui.'
  };
}

function setStatusFacility(username, facilityId, statusBaru) {
  if (!isNikPunyaAksesSetting_(username)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Facility Management.'
    };
  }
  const id = String(facilityId || '').trim();
  const status = String(statusBaru || '').trim().toLowerCase() === 'nonaktif' ? 'Nonaktif' : 'Aktif';
  if (!id) return {
    success: false,
    message: 'ID Facility tidak valid.'
  };
  const sheet = getFacilitySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {
    success: false,
    message: 'Tidak ada data facility.'
  };
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  let targetIdx = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === id) {
      targetIdx = i;
      break;
    }
  }
  if (targetIdx === -1) return {
    success: false,
    message: 'Facility dengan ID tersebut tidak ditemukan.'
  };
  const namaFacility = String(values[targetIdx][1] || '').trim();
  sheet.getRange(targetIdx + 2, 6).setValue(status);
  clearFacilityCache_();
  catatLogPerubahanConfig_(username, 'Facility Management', 'Status facility "' + namaFacility + '" diubah menjadi ' + status + '.');
  return {
    success: true,
    message: 'Facility "' + namaFacility + '" statusnya sekarang: ' + status + '.'
  };
}

function getFacilityInfoById_(facilityId) {
  const id = String(facilityId || '').trim();
  if (!id) return null;
  const facSheet = getFacilitySheet_();
  const facLastRow = facSheet.getLastRow();
  if (facLastRow < 2) return null;
  const facValues = facSheet.getRange(2, 1, facLastRow - 1, FACILITY_HEADERS.length).getValues();
  for (let i = 0; i < facValues.length; i++) {
    if (String(facValues[i][0] || '').trim() === id) {
      return {
        id: id,
        nama: String(facValues[i][1] || '').trim(),
        kode: String(facValues[i][2] || '').trim(),
        spreadsheetId: String(facValues[i][4] || '').trim(),
        status: String(facValues[i][5] || '').trim(),
        url: String(facValues[i][8] || '').trim()
      };
    }
  }
  return null;
}

function getUserFacility(username) {
  const uname = String(username || '').trim().toLowerCase();
  if (!uname) return null;
  if (Object.prototype.hasOwnProperty.call(_memUserFacility_, uname)) return _memUserFacility_[uname];
  const info = getUserFacilityTanpaMemo_(uname);
  _memUserFacility_[uname] = info;
  return info;
}

function getUserFacilityTanpaMemo_(uname) {
  if (isUsernameRoleDeveloper_(uname)) {
    const overrideId = CacheService.getScriptCache().get(CACHE_DEV_ACTIVE_FACILITY_PREFIX + uname);
    if (overrideId) {
      const overrideInfo = getFacilityInfoById_(overrideId);
      if (overrideInfo) return overrideInfo;
      CacheService.getScriptCache().remove(CACHE_DEV_ACTIVE_FACILITY_PREFIX + uname);
    }
  }
  const cache = CacheService.getScriptCache();
  const cacheKey = CACHE_USER_FACILITY_PREFIX + uname;
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }
  let facilityId = null;
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    const muSheet = getMasterUserSheetWithFallback_();
    if (muSheet && muSheet.getLastRow() >= 2) {
      const values = muSheet.getRange(2, 1, muSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][0] || '').trim().toLowerCase() === uname) {
          facilityId = String(values[i][3] || '').trim();
          break;
        }
      }
    }
  }
  if (!facilityId) {
    const sheet = getUserFacilityAssignmentSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let i = values.length - 1; i >= 0; i--) {
        if (String(values[i][0] || '').trim().toLowerCase() === uname) {
          facilityId = String(values[i][1] || '').trim();
          break;
        }
      }
    }
  }
  if (!facilityId) return null;
  const info = getFacilityInfoById_(facilityId);
  if (info) {
    try {
      cache.put(cacheKey, JSON.stringify(info), 300);
    } catch (e) {}
  }
  return info;
}

function setDeveloperActiveFacility(username, facilityId) {
  const uname = String(username || '').trim().toLowerCase();
  if (!uname) return {
    success: false,
    message: 'Username tidak valid.'
  };
  if (!isUsernameRoleDeveloper_(uname)) {
    return {
      success: false,
      message: 'Hanya role Developer yang boleh pakai facility switcher.'
    };
  }
  const cacheKey = CACHE_DEV_ACTIVE_FACILITY_PREFIX + uname;
  const id = String(facilityId || '').trim();
  clearUserRoleCache_(uname);
  if (!id) {
    CacheService.getScriptCache().remove(cacheKey);
    return {
      success: true,
      message: 'Kembali ke facility assignment normal.',
      facility: getUserFacility(uname)
    };
  }
  const info = getFacilityInfoById_(id);
  if (!info) return {
    success: false,
    message: 'Facility tidak ditemukan.'
  };
  if (info.status === 'Nonaktif') {
    return {
      success: false,
      message: 'Facility "' + info.nama + '" sedang nonaktif, tidak bisa dipilih.'
    };
  }
  CacheService.getScriptCache().put(cacheKey, id, DEV_ACTIVE_FACILITY_TTL_SEC);
  return {
    success: true,
    message: 'Berpindah ke facility "' + info.nama + '".',
    facility: info
  };
}

function assignUserKeFacility(username, targetUsername, facilityId) {
  if (!isNikPunyaAksesSetting_(username)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Facility Management.'
    };
  }
  const targetUname = String(targetUsername || '').trim();
  const facId = String(facilityId || '').trim();
  if (!targetUname) return {
    success: false,
    message: 'Username target tidak boleh kosong.'
  };
  if (!facId) return {
    success: false,
    message: 'ID Facility tidak boleh kosong.'
  };
  const userInfo = getUserRole(targetUname);
  if (!userInfo) {
    return {
      success: false,
      message: 'Username "' + targetUname + '" tidak terdaftar di Master.'
    };
  }
  const facSheet = getFacilitySheet_();
  const facLastRow = facSheet.getLastRow();
  let namaFacility = '';
  let facStatus = '';
  if (facLastRow >= 2) {
    const facValues = facSheet.getRange(2, 1, facLastRow - 1, 6).getValues();
    let ketemu = false;
    for (let i = 0; i < facValues.length; i++) {
      if (String(facValues[i][0] || '').trim() === facId) {
        namaFacility = String(facValues[i][1] || '').trim();
        facStatus = String(facValues[i][5] || '').trim();
        ketemu = true;
        break;
      }
    }
    if (!ketemu) return {
      success: false,
      message: 'Facility dengan ID tersebut tidak ditemukan.'
    };
  } else {
    return {
      success: false,
      message: 'Belum ada facility yang terdaftar.'
    };
  }
  if (facStatus === 'Nonaktif') {
    return {
      success: false,
      message: 'Facility "' + namaFacility + '" sedang Nonaktif. Tidak bisa di-assign.'
    };
  }
  const sheet = getUserFacilityAssignmentSheet_();
  const lastRow = sheet.getLastRow();
  let existingRow = 0;
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim().toLowerCase() === targetUname.toLowerCase()) {
        existingRow = i + 2;
        break;
      }
    }
  }
  const now = new Date;
  if (existingRow > 0) {
    sheet.getRange(existingRow, 2).setValue(facId);
    sheet.getRange(existingRow, 3).setValue(now);
    sheet.getRange(existingRow, 4).setValue(username);
  } else {
    const targetRow = Math.max(lastRow + 1, 2);
    sheet.getRange(targetRow, 1).setValue(targetUname);
    sheet.getRange(targetRow, 2).setValue(facId);
    sheet.getRange(targetRow, 3).setValue(now);
    sheet.getRange(targetRow, 4).setValue(username);
  }
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    const muSheet = getMasterUserSheetWithFallback_();
    if (muSheet) {
      const muLastRow = muSheet.getLastRow();
      if (muLastRow >= 2) {
        const muValues = muSheet.getRange(2, 1, muLastRow - 1, 1).getValues();
        for (let i = 0; i < muValues.length; i++) {
          if (String(muValues[i][0] || '').trim().toLowerCase() === targetUname.toLowerCase()) {
            muSheet.getRange(i + 2, 4).setValue(facId);
            break;
          }
        }
      }
    }
  }
  clearUserFacilityCache_(targetUname);
  clearUserRoleCache_(targetUname);
  catatLogPerubahanConfig_(username, 'Facility Management', 'User "' + targetUname + '" di-assign ke facility "' + namaFacility + '".');
  return {
    success: true,
    message: 'User "' + targetUname + '" berhasil di-assign ke facility "' + namaFacility + '".'
  };
}

function getDaftarUserFacilityAssignment(username) {
  if (!isNikPunyaAksesSetting_(username)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Facility Management.'
    };
  }
  const sheet = getUserFacilityAssignmentSheet_();
  const lastRow = sheet.getLastRow();
  const assignments = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, USER_FACILITY_ASSIGNMENT_HEADERS.length).getValues();
    values.forEach(function(r, idx) {
      const uname = String(r[0] || '').trim();
      if (!uname) return;
      assignments.push({
        rowIndex: idx + 2,
        username: uname,
        facilityId: String(r[1] || '').trim(),
        tanggalDiassign: r[2] instanceof Date ? Utilities.formatDate(r[2], 'Asia/Jakarta', 'yyyy-MM-dd HH:mm') : String(r[2] || ''),
        diassignOleh: String(r[3] || '').trim()
      });
    });
  }
  return {
    success: true,
    assignments: assignments
  };
}

var _memSpreadsheetCache_ = {};

function getOperasionalSpreadsheet_(username) {
  const facInfo = getUserFacility(username);
  if (!facInfo || facInfo.status === 'Nonaktif') {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  const sid = facInfo.spreadsheetId;
  if (_memSpreadsheetCache_[sid]) {
    return _memSpreadsheetCache_[sid];
  }
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'facSsDown_' + facInfo.id;
    if (cache.get(cacheKey)) {
      return SpreadsheetApp.getActiveSpreadsheet();
    }
    const ss = SpreadsheetApp.openById(sid);
    _memSpreadsheetCache_[sid] = ss;
    return ss;
  } catch (e) {
    try {
      CacheService.getScriptCache().put('facSsDown_' + facInfo.id, '1', 60);
    } catch (e2) {}
    catatLogSistem_('Facility DB', 'Gagal membuka spreadsheet facility "' + (facInfo.nama || facInfo.id) + '" (' + e.message + '). Fallback ke spreadsheet aktif.');
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

const PESAN_TANPA_FACILITY_ = 'Akun Anda belum punya facility aktif. Minta admin mengaturnya di Config.';

const PESAN_FACILITY_BERUBAH_ = 'Facility akun ini berubah sebelum hasil terkirim. Hitung ulang dari daftar tugas yang baru.';

function facilityKlienCocok_(facKlien, facInfo) {
  const f = String(facKlien === null || facKlien === undefined ? '' : facKlien).trim();
  return !f || f === '-' || f === String(facInfo.id);
}

function requireUserFacility_(username) {
  const facInfo = getUserFacility(username);
  if (!facInfo || facInfo.status === 'Nonaktif') return null;
  return facInfo;
}

function acquireFacilityLock_(facilityId, timeoutMs) {
  const MAX_CONCURRENT_SLOTS = 16;
  const baseKey = 'facLock_' + (facilityId ? String(facilityId).trim() : 'default');
  const cache = CacheService.getScriptCache();
  const deadline = Date.now() + (timeoutMs || 1e4);
  const scriptLock = LockService.getScriptLock();
  const allSlotKeys = [];
  for (let s = 0; s < MAX_CONCURRENT_SLOTS; s++) allSlotKeys.push(baseKey + '_s' + s);
  while (Date.now() < deadline) {
    try {
      scriptLock.waitLock(1e3);
    } catch (e) {
      continue;
    }
    try {
      const slotStates = cache.getAll(allSlotKeys);
      let chosenKey = null;
      for (let s = 0; s < MAX_CONCURRENT_SLOTS; s++) {
        const slotKey = allSlotKeys[s];
        if (!slotStates[slotKey]) {
          chosenKey = slotKey;
          break;
        }
      }
      if (chosenKey) {
        cache.put(chosenKey, '1', 30);
        const capturedSlotKey = chosenKey;
        return {
          release: function() {
            try {
              cache.remove(capturedSlotKey);
            } catch (e2) {}
          }
        };
      }
    } finally {
      scriptLock.releaseLock();
    }
    Utilities.sleep(100);
  }
  return null;
}

function importLokasiAktif(username, facilityId, daftarLokasi, gantiSemua) {
  if (!isNikPunyaAksesSetting_(username)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Facility Management.'
    };
  }
  const facId = String(facilityId || '').trim();
  if (!facId) return {
    success: false,
    message: 'ID Facility tidak boleh kosong.'
  };
  if (!daftarLokasi || !Array.isArray(daftarLokasi) || daftarLokasi.length === 0) {
    return {
      success: false,
      message: 'Daftar lokasi tidak boleh kosong.'
    };
  }
  const facSheet = getFacilitySheet_();
  const facLastRow = facSheet.getLastRow();
  let namaFacility = '';
  let spreadsheetId = '';
  if (facLastRow >= 2) {
    const facValues = facSheet.getRange(2, 1, facLastRow - 1, 5).getValues();
    let ketemu = false;
    for (let i = 0; i < facValues.length; i++) {
      if (String(facValues[i][0] || '').trim() === facId) {
        namaFacility = String(facValues[i][1] || '').trim();
        spreadsheetId = String(facValues[i][4] || '').trim();
        ketemu = true;
        break;
      }
    }
    if (!ketemu) return {
      success: false,
      message: 'Facility tidak ditemukan.'
    };
  } else {
    return {
      success: false,
      message: 'Belum ada facility yang terdaftar.'
    };
  }
  let facilitySs;
  try {
    facilitySs = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    return {
      success: false,
      message: 'Gagal membuka spreadsheet facility: ' + e.message
    };
  }
  let lokasiSheet = facilitySs.getSheetByName(LOKASI_AKTIF_SHEET_NAME);
  if (!lokasiSheet) {
    lokasiSheet = facilitySs.insertSheet(LOKASI_AKTIF_SHEET_NAME);
    lokasiSheet.appendRow(LOKASI_AKTIF_HEADERS);
    lokasiSheet.setFrozenRows(1);
  }
  const lokasiUnik = {};
  const lokasiBersih = [];
  daftarLokasi.forEach(function(lok) {
    const l = String(lok || '').trim();
    if (l && !lokasiUnik[l]) {
      lokasiUnik[l] = true;
      lokasiBersih.push([ l ]);
    }
  });
  if (lokasiBersih.length === 0) {
    return {
      success: false,
      message: 'Tidak ada lokasi valid setelah dibersihkan.'
    };
  }
  const lastRowLokasi = lokasiSheet.getLastRow();
  if (gantiSemua) {
    if (lastRowLokasi >= 2) {
      lokasiSheet.getRange(2, 1, lastRowLokasi - 1, 1).clearContent();
    }
    lokasiSheet.getRange(2, 1, lokasiBersih.length, 1).setValues(lokasiBersih);
  } else {
    const lokasiExisting = {};
    if (lastRowLokasi >= 2) {
      const existing = lokasiSheet.getRange(2, 1, lastRowLokasi - 1, 1).getValues();
      existing.forEach(function(r) {
        const l = String(r[0] || '').trim();
        if (l) lokasiExisting[l] = true;
      });
    }
    const lokasiBaru = lokasiBersih.filter(function(r) {
      return !lokasiExisting[r[0]];
    });
    if (lokasiBaru.length === 0) {
      return {
        success: true,
        message: 'Semua lokasi sudah ada. Tidak ada yang ditambahkan.',
        ditambahkan: 0
      };
    }
    const startRow = Math.max(lastRowLokasi + 1, 2);
    lokasiSheet.getRange(startRow, 1, lokasiBaru.length, 1).setValues(lokasiBaru);
    lokasiBersih.length = lokasiBaru.length;
  }
  clearActiveLocationsCache_(facId);
  catatLogPerubahanConfig_(username, 'Facility Management', 'Import lokasi aktif ke facility "' + namaFacility + '": ' + (gantiSemua ? 'GANTI SEMUA' : 'TAMBAH') + ', ' + lokasiBersih.length + ' lokasi.');
  return {
    success: true,
    message: 'Berhasil mengimpor ' + lokasiBersih.length + ' lokasi ke facility "' + namaFacility + '".',
    ditambahkan: lokasiBersih.length
  };
}

function getDaftarLokasiAktif(username, facilityId) {
  if (!isNikPunyaAksesSetting_(username)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Facility Management.'
    };
  }
  const facId = String(facilityId || '').trim();
  if (!facId) return {
    success: false,
    message: 'ID Facility tidak boleh kosong.'
  };
  const facSheet = getFacilitySheet_();
  const facLastRow = facSheet.getLastRow();
  let spreadsheetId = '';
  if (facLastRow >= 2) {
    const facValues = facSheet.getRange(2, 1, facLastRow - 1, 5).getValues();
    for (let i = 0; i < facValues.length; i++) {
      if (String(facValues[i][0] || '').trim() === facId) {
        spreadsheetId = String(facValues[i][4] || '').trim();
        break;
      }
    }
  }
  if (!spreadsheetId) return {
    success: false,
    message: 'Facility tidak ditemukan.'
  };
  try {
    const facilitySs = SpreadsheetApp.openById(spreadsheetId);
    const lokasiSheet = facilitySs.getSheetByName(LOKASI_AKTIF_SHEET_NAME);
    if (!lokasiSheet) return {
      success: true,
      lokasi: [],
      total: 0
    };
    const lastRow = lokasiSheet.getLastRow();
    const lokasi = [];
    if (lastRow >= 2) {
      const values = lokasiSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      values.forEach(function(r) {
        const l = String(r[0] || '').trim();
        if (l) lokasi.push(l);
      });
    }
    return {
      success: true,
      lokasi: lokasi,
      total: lokasi.length
    };
  } catch (e) {
    return {
      success: false,
      message: 'Gagal membaca lokasi: ' + e.message
    };
  }
}

function copyLokasiDariFacility(username, facilityIdSumber, facilityIdTarget, gantiSemua) {
  if (!isNikPunyaAksesSetting_(username)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Facility Management.'
    };
  }
  const hasilSumber = getDaftarLokasiAktif(username, facilityIdSumber);
  if (!hasilSumber.success) {
    return {
      success: false,
      message: 'Gagal baca lokasi sumber: ' + hasilSumber.message
    };
  }
  if (!hasilSumber.lokasi || hasilSumber.lokasi.length === 0) {
    return {
      success: false,
      message: 'Facility sumber tidak memiliki lokasi aktif.'
    };
  }
  return importLokasiAktif(username, facilityIdTarget, hasilSumber.lokasi, gantiSemua);
}

function seedFacilityAwalDariSistemLama() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('FACILITY_SEEDED')) {
    Logger.log('Seed facility sudah pernah dijalankan. Lewati.');
    return;
  }
  const ssAktif = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getFacilitySheet_();
  const lastRow = sheet.getLastRow();
  let sudahAda = false;
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim().toUpperCase() === 'RDC-MKS') {
        sudahAda = true;
        break;
      }
    }
  }
  if (!sudahAda) {
    const idFacility = 'FAC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const targetRow = Math.max(lastRow + 1, 2);
    const now = new Date;
    sheet.getRange(targetRow, 1).setValue(idFacility);
    sheet.getRange(targetRow, 2).setValue('RDC Makassar');
    sheet.getRange(targetRow, 3).setValue('RDC-MKS');
    sheet.getRange(targetRow, 4).setValue(ssAktif.getName());
    sheet.getRange(targetRow, 5).setValue(ssAktif.getId());
    sheet.getRange(targetRow, 6).setValue('Aktif');
    sheet.getRange(targetRow, 7).setValue(now);
    sheet.getRange(targetRow, 8).setValue('system-migration');
    sheet.getRange(targetRow, 9).setValue(ssAktif.getUrl());
    Logger.log('Facility RDC Makassar terdaftar dengan ID: ' + idFacility);
  }
  let lokasiSheet = ssAktif.getSheetByName(LOKASI_AKTIF_SHEET_NAME);
  if (!lokasiSheet) {
    lokasiSheet = ssAktif.insertSheet(LOKASI_AKTIF_SHEET_NAME);
    lokasiSheet.appendRow(LOKASI_AKTIF_HEADERS);
    lokasiSheet.setFrozenRows(1);
    const masterSheet = getMasterSheet_();
    if (masterSheet) {
      const mLastRow = masterSheet.getLastRow();
      if (mLastRow >= 2) {
        const lokasiValues = masterSheet.getRange(2, 1, mLastRow - 1, 1).getValues();
        const lokasiBersih = [];
        const lokasiUnik = {};
        lokasiValues.forEach(function(r) {
          const l = String(r[0] || '').trim();
          if (l && !lokasiUnik[l]) {
            lokasiUnik[l] = true;
            lokasiBersih.push([ l ]);
          }
        });
        if (lokasiBersih.length > 0) {
          lokasiSheet.getRange(2, 1, lokasiBersih.length, 1).setValues(lokasiBersih);
          Logger.log('Lokasi aktif disalin dari master_db: ' + lokasiBersih.length + ' lokasi.');
        }
      }
    }
  }
  props.setProperty('FACILITY_SEEDED', '1');
  clearFacilityCache_();
  Logger.log('Seed facility awal SELESAI.');
}

function getHomeSummary(username) {
  const info = requireRole_(username, ALL_CYCLE_LIKE_ROLES);
  catatAksesLog_(username, info);
  const today = Utilities.formatDate(new Date, 'Asia/Jakarta', 'yyyy-MM-dd');
  let totalCycleHariIni = 0, selesaiHariIni = 0;
  const todayRows = readRingkasanHarianRows_('daily', today, username);
  if (todayRows.length) {
    const r = todayRows[0];
    if (info.role === 'admin' || info.role === 'developer') {
      totalCycleHariIni = Number(r[1]) || 0;
      selesaiHariIni = (Number(r[2]) || 0) + (Number(r[3]) || 0);
    } else {
      const petugasRawMap = parseJsonSafe_(r[7], {});
      const mine = petugasRawMap[info.displayName];
      if (mine) {
        totalCycleHariIni = mine.rawTotal || 0;
        selesaiHariIni = mine.rawSelesai || 0;
      }
    }
  }
  let belumValidasi = 0;
  if (info.role === 'admin' || info.role === 'developer') {
    belumValidasi = readAntrianValue_('pending_total', {
      count: 0
    }, username).count || 0;
  } else if (info.role === 'inventory') {
    belumValidasi = readAntrianValue_('pending_validator:' + info.displayName, {
      count: 0
    }, username).count || 0;
  }
  return {
    totalCycleHariIni: totalCycleHariIni,
    selesaiHariIni: selesaiHariIni,
    belumValidasi: belumValidasi,
    lastUpdated: Utilities.formatDate(new Date, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm'),
    version: APP_VERSION
  };
}

function catatAksesLog_(username, info) {
  try {
    const sheet = getLogAksesSheet_(username);
    const tz = 'Asia/Jakarta';
    const now = new Date;
    const tanggal = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const jamStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
    const uname = String(username || '').trim().toLowerCase();
    if (!uname) return;
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        if (String(data[i][0]) === tanggal && String(data[i][1]) === uname) {
          const rowIndex = i + 2;
          const jumlahCell = sheet.getRange(rowIndex, 8);
          sheet.getRange(rowIndex, 7).setValue(jamStr);
          jumlahCell.setValue(Number(jumlahCell.getValue() || 0) + 1);
          return;
        }
      }
    }
    sheet.appendRow([ tanggal, uname, info.displayName || uname, info.role || '', APP_VERSION, jamStr, jamStr, 1 ]);
  } catch (e) {}
}

function getPlusMinusSummary(requesterUsername) {
  requireRole_(requesterUsername, [ 'inventory', 'admin', 'developer' ]);
  return getPlusMinusSummaryFromAntrian_(requesterUsername);
}

function testSummary() {
  const result = getSummaryByPeriod('daily', '2026-07-04', SETUP_ADMIN_NIK_);
  Logger.log(JSON.stringify(result, null, 2));
}

function getExistingBacklogEffort_(username) {
  const sheet = getSheet_(username);
  const lastRow = sheet.getLastRow();
  const prod = getLevelSettings_();
  const result = {};
  if (lastRow < 2) return result;
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  data.forEach(function(row) {
    const status = row[10];
    if (status !== 'Pending') return;
    const namaPetugas = String(row[9] || '').trim();
    if (!namaPetugas) return;
    const group = getLevelGroupByLocation_(row[2]);
    const productivity = prod[group.key] > 0 ? prod[group.key] : 1;
    result[namaPetugas] = (result[namaPetugas] || 0) + 1 / productivity;
  });
  return result;
}

function getExistingPendingMap_(username) {
  const sheet = getSheet_(username);
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  data.forEach(function(row, idx) {
    if (row[10] !== 'Pending') return;
    const key = String(row[2]).trim() + '||' + String(row[4]).trim();
    map[key] = {
      row: idx + 2,
      qtyTransaksi: Number(row[6]) || 0,
      addWho: row[15]
    };
  });
  return map;
}

function buildAssignedRows_(items, startingNo, tanggalUpload, usernames, existingBacklogEffort, compareFn) {
  const prod = getLevelSettings_();
  items.forEach(function(it) {
    const level = extractLevel_(it.lokasi);
    it.level = level;
    const group = getLevelGroupByLocation_(it.lokasi);
    it.levelGroupLabel = group.label;
    const productivity = prod[group.key] > 0 ? prod[group.key] : 1;
    it.effort = 1 / productivity;
    it.locationSort = parseLocation_(it.lokasi);
  });
  items.sort(compareFn || compareLocation_);
  const locGroups = [];
  items.forEach(function(it) {
    const last = locGroups[locGroups.length - 1];
    if (last && last.lokasi === it.lokasi) {
      last.items.push(it);
      last.effort += it.effort;
    } else {
      locGroups.push({
        lokasi: it.lokasi,
        items: [ it ],
        effort: it.effort
      });
    }
  });
  const userCount = usernames.length;
  const totalNewEffort = items.reduce(function(s, it) {
    return s + it.effort;
  }, 0);
  const totalBacklogEffort = usernames.reduce(function(s, u) {
    return s + (existingBacklogEffort[u] || 0);
  }, 0);
  const targetPerUser = userCount > 0 ? (totalNewEffort + totalBacklogEffort) / userCount : totalNewEffort;
  const output = [];
  let no = startingNo;
  let userIdx = 0;
  let accEffort = existingBacklogEffort[usernames[0]] || 0;
  let chunk = [];
  function flushChunk() {
    if (!chunk.length) return;
    const username = usernames[Math.min(userIdx, userCount - 1)];
    const firstLoc = chunk[0].lokasi;
    const lastLoc = chunk[chunk.length - 1].lokasi;
    const mixLabel = getGroupMixLabel_(chunk);
    const batchLabel = 'Tugas - ' + tanggalUpload + ' - ' + username + ' (' + mixLabel + ') [' + firstLoc + ' s/d ' + lastLoc + ']';
    chunk.forEach(function(it) {
      output.push([ no, tanggalUpload, it.lokasi, it.level !== null ? it.level : '', it.article, it.description, it.qtyTransaksi, it.qtySystem, batchLabel, username, 'Pending', '', '', '', '', it.addWhoList || '' ]);
      no++;
    });
    chunk = [];
  }
  locGroups.forEach(function(grp, idx) {
    chunk = chunk.concat(grp.items);
    accEffort += grp.effort;
    const isLastGroup = idx === locGroups.length - 1;
    const isLastUser = userIdx >= userCount - 1;
    if (isLastGroup) {
      flushChunk();
    } else if (!isLastUser && accEffort >= targetPerUser) {
      flushChunk();
      userIdx++;
      accEffort = existingBacklogEffort[usernames[userIdx]] || 0;
    }
  });
  return output;
}

function getEquipmentReadyDefaults_() {
  const result = {
    reachTruck: 0,
    tangga: 0
  };
  const sheet = getMasterSheet_();
  if (!sheet) return result;
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return result;
  const data = sheet.getRange(1, 6, lastRow, 2).getValues();
  data.forEach(function(row) {
    const label = String(row[0] || '').trim().toLowerCase();
    const value = Number(row[1]) || 0;
    if (!label) return;
    if (label.indexOf('reach truck') !== -1) result.reachTruck = value; else if (label.indexOf('tangga') !== -1) result.tangga = value;
  });
  return result;
}

function getEquipmentReadyDefaults(requesterUsername) {
  requireRole_(requesterUsername, [ 'admin', 'developer' ]);
  return getEquipmentReadyDefaults_();
}

function buildAssignedRowsEquipmentAware_(items, startingNo, tanggalUpload, reachTruckUsers, tanggaUsers, bawahUsers, existingBacklogEffort, username) {
  const warnings = [];
  const reachTruckItems = [];
  const tanggaItems = [];
  const bawahItems = [];
  items.forEach(function(it) {
    const group = getLevelGroupByLocation_(it.lokasi);
    if (group.key === 'reach_truck') reachTruckItems.push(it); else if (group.key === 'tangga') tanggaItems.push(it); else bawahItems.push(it);
  });
  const overflowResult = applyOverflowStoringKeTangga_(reachTruckItems.length, reachTruckUsers, tanggaUsers, username);
  reachTruckUsers = overflowResult.reachTruckUsers;
  tanggaUsers = overflowResult.tanggaUsers;
  const overflowPesan = overflowResult.pesan;
  let no = startingNo;
  let output = [];
  if (reachTruckItems.length) {
    if (!reachTruckUsers.length) {
      warnings.push(reachTruckItems.length + ' item Level 5-6 (reach truck) tidak dibagi karena tidak ada petugas Storing yang dipilih.');
    } else {
      const rows1 = buildAssignedRows_(reachTruckItems, no, tanggalUpload, reachTruckUsers, existingBacklogEffort, compareLocationEvenOddSection_);
      output = output.concat(rows1);
      no += rows1.length;
    }
  }
  if (tanggaItems.length) {
    if (!tanggaUsers.length) {
      warnings.push(tanggaItems.length + ' item Level 3-4 (tangga pesawat) tidak dibagi karena jumlah tangga pesawat 0.');
    } else {
      const rows2 = buildAssignedRows_(tanggaItems, no, tanggalUpload, tanggaUsers, existingBacklogEffort, compareLocationEvenOddSection_);
      output = output.concat(rows2);
      no += rows2.length;
    }
  }
  if (bawahItems.length) {
    if (!bawahUsers.length) {
      warnings.push(bawahItems.length + ' item Level 1-2 tidak dibagi karena tidak ada petugas tersisa setelah reach truck dan tangga pesawat.');
    } else {
      const rows3 = buildAssignedRows_(bawahItems, no, tanggalUpload, bawahUsers, existingBacklogEffort);
      output = output.concat(rows3);
      no += rows3.length;
    }
  }
  return {
    rows: output,
    warnings: warnings,
    overflowPesan: overflowPesan,
    overflowCount: overflowResult.overflowCount
  };
}

function getGroupMixLabel_(items) {
  const present = {};
  items.forEach(function(it) {
    present[it.levelGroupLabel] = true;
  });
  const labels = Object.keys(present);
  return labels.length > 1 ? 'Campuran' : labels[0];
}

function writeItemsToSheet_(items, reachTruckUsers, tanggaUsers, bawahUsers, selectedUsernames, usernameRoleMap, username) {
  const sheet = getSheet_(username);
  const lastRow = sheet.getLastRow();
  const startingNo = lastRow >= 2 ? Number(sheet.getRange(lastRow, 1).getValue() || 1) + 1 : 1;
  const tanggalUpload = Utilities.formatDate(new Date, 'Asia/Jakarta', 'yyyy-MM-dd');
  const existingBacklogEffort = getExistingBacklogEffort_(username);
  const modeAssignment = getLevelAssignmentMeta_(getLevelAssignmentSheet_(username)).mode;
  const result = modeAssignment === 'matrix' ? buildAssignedRowsMatrixAware_(items, startingNo, tanggalUpload, selectedUsernames, usernameRoleMap, existingBacklogEffort, username) : buildAssignedRowsEquipmentAware_(items, startingNo, tanggalUpload, reachTruckUsers, tanggaUsers, bawahUsers, existingBacklogEffort, username);
  const rows = result.rows;
  if (rows.length > 0) {
    const startRow = lastRow + 1;
    sheet.getRange(startRow, 2, rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 5, rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);
  }
  return {
    total: rows.length,
    warnings: result.warnings,
    overflowPesan: result.overflowPesan,
    overflowCount: result.overflowCount,
    modeAssignment: modeAssignment
  };
}

function catatLogUnassigned_(adminUsername, warningMessages, username) {
  try {
    const sheet = getLogUnassignedSheet_(username);
    const now = new Date;
    warningMessages.forEach(function(msg) {
      sheet.appendRow([ now, adminUsername, msg ]);
    });
  } catch (e) {}
}

function importRawData(transaksiRawRows, stockRows, selectedUsernames, requesterUsername, reachTruckReady, tanggaReady, jobId) {
  requireRole_(requesterUsername, [ 'admin', 'developer' ]);
  reportProgress_(jobId, 5, 'Validasi input');
  if (!transaksiRawRows || !transaksiRawRows.length) {
    return {
      success: false,
      message: 'Data Transaksi (raw) kosong.'
    };
  }
  if (!stockRows || !stockRows.length) {
    return {
      success: false,
      message: 'Data Stock by Location kosong.'
    };
  }
  const usernames = (selectedUsernames || []).map(function(u) {
    return String(u).trim();
  }).filter(function(u) {
    return u;
  });
  if (!usernames.length) {
    return {
      success: false,
      message: 'Pilih minimal 1 user yang bertugas.'
    };
  }
  const facInfo = requireUserFacility_(requesterUsername);
  if (!facInfo) {
    return {
      success: false,
      message: PESAN_TANPA_FACILITY_
    };
  }
  const facilityId = facInfo.id;
  const facilityName = facInfo.nama || facInfo.name || '';
  const assignableInFacility = getAssignableUsers_(facilityId);
  const assignableMap = {};
  assignableInFacility.forEach(function(u) {
    assignableMap[u.username.toLowerCase()] = true;
  });
  const userTidakValid = [];
  usernames.forEach(function(u) {
    if (!assignableMap[u.toLowerCase()]) {
      userTidakValid.push(u);
    }
  });
  if (userTidakValid.length > 0) {
    return {
      success: false,
      message: 'User berikut tidak terdaftar di facility "' + facilityName + '" atau tidak bisa diberi tugas: ' + userTidakValid.join(', ') + '. Muat ulang daftar petugas lalu pilih lagi.'
    };
  }
  const roleMap = getUserRoleMap_();
  const reachTruckUsers = usernames.filter(function(u) {
    return roleMap[u.toLowerCase()] === 'storing';
  });
  const nonStoring = usernames.filter(function(u) {
    return roleMap[u.toLowerCase()] !== 'storing';
  });
  const m = Math.max(0, Math.min(Number(tanggaReady) || 0, nonStoring.length));
  const tanggaUsers = nonStoring.slice(0, m);
  const bawahUsers = nonStoring.slice(m);
  const reachTruckReadyNum = Number(reachTruckReady) || 0;
  const modeAssignmentSaatIni = getLevelAssignmentMeta_(getLevelAssignmentSheet_(requesterUsername)).mode;
  const reachTruckNote = modeAssignmentSaatIni === 'legacy' && reachTruckReadyNum && reachTruckReadyNum !== reachTruckUsers.length ? 'Jumlah reach truck (' + reachTruckReadyNum + ') berbeda dari jumlah petugas Storing yang dipilih (' + reachTruckUsers.length + '). Semua petugas Storing yang dipilih tetap mendapat tugas Level 5-6.' : '';
  reportProgress_(jobId, 15, 'Memuat lokasi aktif facility');
  const activeLocations = getActiveLocations_(requesterUsername);
  if (Object.keys(activeLocations).length === 0) {
    return {
      success: false,
      message: 'Tidak ada lokasi aktif untuk facility "' + facilityName + '". Impor lokasi aktif dulu lewat Config, tab Facility.'
    };
  }
  const facLock = acquireFacilityLock_(facilityId, 2e4);
  if (!facLock) {
    return {
      success: false,
      message: 'Sistem sedang sibuk memproses upload lain, coba lagi sebentar.'
    };
  }
  try {
    reportProgress_(jobId, 25, 'Memproses data stock');
    const stockMap = {};
    stockRows.forEach(function(r) {
      const lokasi = r[0] === undefined || r[0] === null ? '' : String(r[0]).trim();
      const article = r[1] === undefined || r[1] === null ? '' : String(r[1]).trim();
      if (!lokasi && !article) return;
      const key = lokasi + '||' + article;
      stockMap[key] = (stockMap[key] || 0) + (Number(r[2]) || 0);
    });
    reportProgress_(jobId, 40, 'Memproses data transaksi');
    const itemMap = {};
    let skippedBlankLokasi = 0;
    let skippedInactive = 0;
    let skippedOtherType = 0;
    transaksiRawRows.forEach(function(r) {
      const type = String(r[0] || '').trim().toLowerCase();
      if (type !== 'move' && type !== 'picking') {
        if (type) skippedOtherType++;
        return;
      }
      const article = String(r[1] || '').trim();
      const description = String(r[2] || '').trim();
      const lokasiH = String(r[3] || '').trim();
      const lokasiJ = String(r[4] || '').trim();
      const qtyTransaksi = Number(r[5]) || 0;
      const addWho = String(r[6] || '').trim();
      const lokasi = type === 'move' ? lokasiJ : lokasiH;
      if (!lokasi) {
        skippedBlankLokasi++;
        return;
      }
      if (!activeLocations[lokasi]) {
        skippedInactive++;
        return;
      }
      const key = lokasi + '||' + article;
      if (!itemMap[key]) {
        const qtySystem = stockMap.hasOwnProperty(key) ? Number(stockMap[key]) : 0;
        itemMap[key] = {
          lokasi: lokasi,
          article: article,
          description: description,
          qtyTransaksi: 0,
          qtySystem: qtySystem,
          addWhoByType: {
            move: {},
            picking: {},
            lainnya: {}
          }
        };
      }
      itemMap[key].qtyTransaksi += qtyTransaksi;
      if (addWho) itemMap[key].addWhoByType[type][addWho] = true;
    });
    let items = Object.keys(itemMap).map(function(k) {
      const it = itemMap[k];
      it.addWhoList = encodeAddWhoByType_(it.addWhoByType);
      return it;
    });
    if (!items.length) {
      return {
        success: false,
        message: 'Tidak ada item yang valid untuk dibagi tugas. ' + (skippedBlankLokasi ? skippedBlankLokasi + ' baris lokasi kosong. ' : '') + (skippedInactive ? skippedInactive + ' baris lokasi tidak aktif di facility "' + facilityName + '". ' : '') + (skippedOtherType ? skippedOtherType + ' baris tipe selain Move/Picking. ' : '')
      };
    }
    reportProgress_(jobId, 60, 'Menggabungkan tugas existing (dedup)');
    const existingMap = getExistingPendingMap_(requesterUsername);
    const mergedKeys = {};
    let mergedCount = 0;
    items = items.filter(function(it) {
      const key = it.lokasi + '||' + it.article;
      const existing = existingMap[key];
      if (!existing) return true;
      mergedKeys[key] = existing;
      mergedCount++;
      return false;
    });
    if (mergedCount > 0) {
      const sheet = getSheet_(requesterUsername);
      const mergedRows = Object.keys(mergedKeys).map(function(k) {
        return mergedKeys[k];
      });
      mergedRows.sort(function(a, b) {
        return a.row - b.row;
      });
      const minRow = mergedRows[0].row;
      const maxRow = mergedRows[mergedRows.length - 1].row;
      const numRows = maxRow - minRow + 1;
      const rangeQty = sheet.getRange(minRow, 7, numRows, 1);
      const rangeAddWho = sheet.getRange(minRow, 16, numRows, 1);
      const qtyValues = rangeQty.getValues();
      const addWhoValues = rangeAddWho.getValues();
      Object.keys(mergedKeys).forEach(function(key) {
        const existing = mergedKeys[key];
        const it = itemMap[key];
        const relRow = existing.row - minRow;
        const qtyLama = Number(qtyValues[relRow][0]) || 0;
        qtyValues[relRow][0] = qtyLama + it.qtyTransaksi;
        const addWhoLama = String(addWhoValues[relRow][0] || '');
        const addWhoBaru = mergeAddWho_(addWhoLama, it.addWhoByType);
        addWhoValues[relRow][0] = addWhoBaru;
      });
      rangeQty.setValues(qtyValues);
      rangeAddWho.setValues(addWhoValues);
    }
    if (!items.length) {
      return {
        success: true,
        message: mergedCount + ' item digabung dengan tugas Pending yang sudah ada (qty ditambahkan). Tidak ada tugas baru dibuat.',
        merged: mergedCount,
        ditugaskan: 0,
        peringatan: [],
        catatan: []
      };
    }
    const tulis = writeItemsToSheet_(items, reachTruckUsers, tanggaUsers, bawahUsers, usernames, roleMap, requesterUsername);
    reportProgress_(jobId, 95, 'Menyimpan hasil pembagian tugas');
    let pesan = 'Sukses. ' + tulis.total + ' item ditugaskan ke ' + usernames.length + ' user.';
    if (mergedCount) pesan += ' ' + mergedCount + ' item digabung dengan tugas existing.';
    if (skippedInactive) pesan += ' ' + skippedInactive + ' lokasi tidak aktif di facility ini.';
    if (skippedBlankLokasi) pesan += ' ' + skippedBlankLokasi + ' lokasi kosong.';
    if (skippedOtherType) pesan += ' ' + skippedOtherType + ' tipe selain Move/Picking.';
    if (reachTruckNote) pesan += ' ' + reachTruckNote;
    if (tulis.overflowPesan) pesan += ' ' + tulis.overflowPesan;
    let warningGabungan = '';
    if (tulis.warnings && tulis.warnings.length) {
      warningGabungan = tulis.warnings.join(' ');
      pesan += ' ⚠️ ' + warningGabungan;
      catatLogUnassigned_(requesterUsername, tulis.warnings, requesterUsername);
    }
    return {
      success: true,
      message: pesan,
      ditugaskan: tulis.total,
      merged: mergedCount,
      skippedInactive: skippedInactive,
      skippedBlankLokasi: skippedBlankLokasi,
      skippedOtherType: skippedOtherType,
      warning: warningGabungan,
      peringatan: tulis.warnings || [],
      catatan: [ reachTruckNote, tulis.overflowPesan ].filter(function(x) {
        return x;
      }),
      perluKonfirmasi: !!(tulis.warnings && tulis.warnings.length),
      modeAssignment: tulis.modeAssignment,
      facilityName: facilityName
    };
  } finally {
    facLock.release();
  }
}

const GRUP_ALAT_LIST_ = [ {
  key: 'bawah',
  label: 'Bawah (Level 1-2)'
}, {
  key: 'tangga',
  label: 'Tangga Pesawat (Level 3-4)'
}, {
  key: 'reach_truck',
  label: 'Reach Truck (Level 5-6)'
} ];

const LEVEL_ASSIGNMENT_SHEET_NAME = 'Config_Level_Assignment';

const LEVEL_ASSIGNMENT_HEADERS = [ 'GrupAlat', 'RolesCSV', 'Mode' ];

const LEVEL_ASSIGNMENT_META_MODE_KEY_ = '_meta_mode_assignment_';

const LEVEL_ASSIGNMENT_META_MAKSGRUP_KEY_ = '_meta_maks_grup_alat_';

const LEVEL_ASSIGNMENT_DEFAULTS_ = {
  bawah: {
    roles: [ 'outbound', 'inbound', 'lp', 'maintenance', 'inventory', 'admin' ],
    mode: 'pemerataan'
  },
  tangga: {
    roles: [ 'outbound', 'inbound', 'lp', 'maintenance', 'inventory', 'admin' ],
    mode: 'pemerataan'
  },
  reach_truck: {
    roles: [ 'storing' ],
    mode: 'pemerataan'
  }
};

const CONFIG_SISTEM_MODE_ASSIGNMENT_KEY_ = 'mode_assignment';

const CONFIG_SISTEM_MAKS_GRUP_ALAT_KEY_ = 'maks_grup_alat_per_role';

function getLevelAssignmentSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LEVEL_ASSIGNMENT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LEVEL_ASSIGNMENT_SHEET_NAME);
    sheet.appendRow(LEVEL_ASSIGNMENT_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, LEVEL_ASSIGNMENT_HEADERS);
  }
  if (sheet.getLastRow() < 2) {
    const rows = GRUP_ALAT_LIST_.map(function(g) {
      const d = LEVEL_ASSIGNMENT_DEFAULTS_[g.key];
      return [ g.key, d.roles.join(','), d.mode ];
    });
    sheet.getRange(2, 1, rows.length, LEVEL_ASSIGNMENT_HEADERS.length).setValues(rows);
  }
  return sheet;
}

function findLevelAssignmentMetaRow_(sheet, metaKey) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === metaKey) {
      return {
        rowIndex: i + 2,
        value: String(values[i][1] || '').trim()
      };
    }
  }
  return null;
}

function getLevelAssignmentMeta_(sheet) {
  const modeRow = findLevelAssignmentMetaRow_(sheet, LEVEL_ASSIGNMENT_META_MODE_KEY_);
  const maksRow = findLevelAssignmentMetaRow_(sheet, LEVEL_ASSIGNMENT_META_MAKSGRUP_KEY_);
  return {
    mode: modeRow && modeRow.value === 'matrix' ? 'matrix' : 'legacy',
    maksGrupAlat: maksRow && Number(maksRow.value) > 0 ? Number(maksRow.value) : 2
  };
}

function setLevelAssignmentMetaValue_(sheet, metaKey, value) {
  const existing = findLevelAssignmentMetaRow_(sheet, metaKey);
  if (existing) {
    sheet.getRange(existing.rowIndex, 2).setValue(String(value));
  } else {
    sheet.appendRow([ metaKey, String(value), '' ]);
  }
}

function bacaLevelAssignmentMatrix_(username) {
  const sheet = getLevelAssignmentSheet_(username);
  const lastRow = sheet.getLastRow();
  const byKey = {};
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 3).getValues().forEach(function(r) {
      const key = String(r[0] || '').trim();
      if (!key || key === LEVEL_ASSIGNMENT_META_MODE_KEY_ || key === LEVEL_ASSIGNMENT_META_MAKSGRUP_KEY_) return;
      byKey[key] = {
        roles: String(r[1] || '').split(',').map(function(s) {
          return s.trim().toLowerCase();
        }).filter(function(s) {
          return s;
        }),
        mode: String(r[2] || '').trim().toLowerCase() === 'prioritas' ? 'prioritas' : 'pemerataan'
      };
    });
  }
  return GRUP_ALAT_LIST_.map(function(g) {
    const data = byKey[g.key] || LEVEL_ASSIGNMENT_DEFAULTS_[g.key];
    return {
      grupAlat: g.key,
      label: g.label,
      roles: data.roles,
      mode: data.mode
    };
  });
}

function getLevelAssignmentConfig(requesterUsername) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Pembagian Tugas.'
    };
  }
  const sheet = getLevelAssignmentSheet_(requesterUsername);
  const meta = getLevelAssignmentMeta_(sheet);
  return {
    success: true,
    matrix: bacaLevelAssignmentMatrix_(requesterUsername),
    modeAssignment: meta.mode,
    maksGrupAlat: meta.maksGrupAlat,
    roleValid: MASTER_ROLE_VALID_,
    grupAlatList: GRUP_ALAT_LIST_
  };
}

function saveLevelAssignmentMatrix(requesterUsername, matrixBaru, maksGrupAlatBaru) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Pembagian Tugas.'
    };
  }
  if (!Array.isArray(matrixBaru) || matrixBaru.length !== GRUP_ALAT_LIST_.length) {
    return {
      success: false,
      message: 'Data matrix tidak lengkap (harus ada semua ' + GRUP_ALAT_LIST_.length + ' grup alat).'
    };
  }
  const matrixLama = bacaLevelAssignmentMatrix_(requesterUsername);
  const sheet = getLevelAssignmentSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const rowIndexByKey = {};
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function(r, idx) {
      rowIndexByKey[String(r[0] || '').trim()] = idx + 2;
    });
  }
  const perubahan = [];
  const hitungGrupPerRole = {};
  matrixBaru.forEach(function(g) {
    const key = String(g.grupAlat || '').trim();
    if (!key) return;
    const rolesBersih = (g.roles || []).map(function(r) {
      return String(r).trim().toLowerCase();
    }).filter(function(r) {
      return MASTER_ROLE_VALID_.indexOf(r) !== -1;
    });
    const mode = String(g.mode || '').trim().toLowerCase() === 'prioritas' ? 'prioritas' : 'pemerataan';
    rolesBersih.forEach(function(r) {
      hitungGrupPerRole[r] = (hitungGrupPerRole[r] || 0) + 1;
    });
    const rowIndex = rowIndexByKey[key];
    if (rowIndex) {
      sheet.getRange(rowIndex, 2).setValue(rolesBersih.join(','));
      sheet.getRange(rowIndex, 3).setValue(mode);
    } else {
      sheet.appendRow([ key, rolesBersih.join(','), mode ]);
    }
    const grupLama = matrixLama.filter(function(m) {
      return m.grupAlat === key;
    })[0];
    if (grupLama && (grupLama.roles.join(',') !== rolesBersih.join(',') || grupLama.mode !== mode)) {
      perubahan.push(key + ': [' + grupLama.roles.join(',') + ' / ' + grupLama.mode + '] -> [' + rolesBersih.join(',') + ' / ' + mode + ']');
    }
  });
  const maksGrupAlat = Math.max(1, Number(maksGrupAlatBaru) || 2);
  setLevelAssignmentMetaValue_(sheet, LEVEL_ASSIGNMENT_META_MAKSGRUP_KEY_, maksGrupAlat);
  if (perubahan.length) {
    catatLogPerubahanConfig_(requesterUsername, 'Level Assignment', perubahan.join('; '));
  }
  const roleMelebihiBatas = Object.keys(hitungGrupPerRole).filter(function(r) {
    return hitungGrupPerRole[r] > maksGrupAlat;
  });
  const warning = roleMelebihiBatas.length ? 'Role ' + roleMelebihiBatas.join(', ') + ' dicentang di lebih dari ' + maksGrupAlat + ' grup alat -- tersimpan, tapi cek lagi apakah ini disengaja.' : '';
  return {
    success: true,
    message: 'Matrix Pembagian Tugas tersimpan.',
    warning: warning
  };
}

function setModeAssignment(requesterUsername, mode) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Pembagian Tugas.'
    };
  }
  const modeBersih = String(mode || '').trim().toLowerCase() === 'matrix' ? 'matrix' : 'legacy';
  const sheet = getLevelAssignmentSheet_(requesterUsername);
  setLevelAssignmentMetaValue_(sheet, LEVEL_ASSIGNMENT_META_MODE_KEY_, modeBersih);
  catatLogPerubahanConfig_(requesterUsername, 'Level Assignment', 'Mode assignment diubah jadi: ' + modeBersih);
  return {
    success: true,
    message: modeBersih === 'legacy' ? 'Kembali ke Mode Lama (By Alat) -- assignment harian pakai logic Storing/non-Storing seperti sebelumnya.' : 'Mode Matrix aktif -- assignment harian sekarang mengikuti Config_Level_Assignment facility ini.'
  };
}

function getLogPerubahanConfig(requesterUsername) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke log perubahan.'
    };
  }
  const sheet = getLogPerubahanConfigSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {
    success: true,
    log: []
  };
  const ambilDari = Math.max(2, lastRow - 49);
  const jumlah = lastRow - ambilDari + 1;
  const values = sheet.getRange(ambilDari, 1, jumlah, 4).getValues();
  const log = values.map(function(r) {
    return {
      waktu: r[0] instanceof Date ? r[0].toISOString() : String(r[0] || ''),
      nik: r[1],
      sumber: r[2],
      detail: r[3]
    };
  }).reverse();
  return {
    success: true,
    log: log
  };
}

const GRUP_ALAT_TO_LEVEL_GROUP_KEY_ = {
  bawah: 'tanpa_alat',
  tangga: 'tangga',
  reach_truck: 'reach_truck'
};

function bangunPoolGrupAlat_(grupConfig, selectedUsernames, usernameRoleMap, demandCount, produktivitas) {
  if (grupConfig.mode === 'prioritas') {
    let pool = [];
    for (let i = 0; i < grupConfig.roles.length; i++) {
      const roleIni = grupConfig.roles[i];
      const roleUsers = selectedUsernames.filter(function(u) {
        return usernameRoleMap[u.toLowerCase()] === roleIni;
      });
      pool = pool.concat(roleUsers);
      const kapasitas = pool.length * (produktivitas > 0 ? produktivitas : 1);
      if (kapasitas >= demandCount) break;
    }
    return pool;
  }
  return selectedUsernames.filter(function(u) {
    return grupConfig.roles.indexOf(usernameRoleMap[u.toLowerCase()]) !== -1;
  });
}

function buildAssignedRowsMatrixAware_(items, startingNo, tanggalUpload, selectedUsernames, usernameRoleMap, existingBacklogEffort, username) {
  const warnings = [];
  const matrix = bacaLevelAssignmentMatrix_(username);
  const prod = getLevelSettings_();
  const itemsByGrup = {
    bawah: [],
    tangga: [],
    reach_truck: []
  };
  items.forEach(function(it) {
    const g = getLevelGroupByLocation_(it.lokasi);
    const grupKey = g.key === 'tanpa_alat' ? 'bawah' : g.key;
    itemsByGrup[grupKey].push(it);
  });
  let no = startingNo;
  let output = [];
  GRUP_ALAT_LIST_.forEach(function(g) {
    const grupItems = itemsByGrup[g.key];
    if (!grupItems.length) return;
    const grupConfig = matrix.filter(function(m) {
      return m.grupAlat === g.key;
    })[0];
    const produktivitas = prod[GRUP_ALAT_TO_LEVEL_GROUP_KEY_[g.key]] || 1;
    const pool = bangunPoolGrupAlat_(grupConfig, selectedUsernames, usernameRoleMap, grupItems.length, produktivitas);
    if (!pool.length) {
      warnings.push(grupItems.length + ' item ' + g.label + ' tidak dibagi karena tidak ada petugas terpilih dengan peran yang sesuai. Periksa Config, tab Pembagian tugas.');
      return;
    }
    const compareFn = g.key === 'bawah' ? undefined : compareLocationEvenOddSection_;
    const rows = buildAssignedRows_(grupItems, no, tanggalUpload, pool, existingBacklogEffort, compareFn);
    output = output.concat(rows);
    no += rows.length;
  });
  const peranBerGrup = {};
  matrix.forEach(function(m) {
    (m.roles || []).forEach(function(r) {
      peranBerGrup[r] = true;
    });
  });
  const tanpaGrup = selectedUsernames.filter(function(u) {
    return !peranBerGrup[usernameRoleMap[String(u).toLowerCase()]];
  });
  const catatanTanpaGrup = tanpaGrup.length ? tanpaGrup.length + ' petugas terpilih tidak mendapat tugas karena perannya belum dicentang di Config, tab Pembagian tugas: ' + tanpaGrup.join(', ') + '.' : '';
  return {
    rows: output,
    warnings: warnings,
    overflowPesan: catatanTanpaGrup,
    overflowCount: 0
  };
}

function doGet(e) {
  if (typeof PANEL_HTML_ === 'string' && PANEL_HTML_) {
    return HtmlService.createHtmlOutput(PANEL_HTML_).setTitle('Cycle Transaksi - Panel Admin').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutput('<div style="font-family:sans-serif;padding:24px;line-height:1.5">' + '<h3>Cycle Transaksi — server aktif (' + APP_VERSION + ')</h3>' + '<p>Halaman ini adalah server aplikasi Android Cycle Transaksi. Alamat halaman ini (yang berakhiran <b>/exec</b>) ' + 'ditanam di aplikasi oleh pengembang; pengguna cukup masuk dengan NIK.</p></div>').setTitle('Cycle Transaksi').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function panelApi(action, argsJson) {
  var args;
  try {
    args = JSON.parse(String(argsJson || '[]'));
  } catch (err) {
    args = null;
  }
  if (!Array.isArray(args)) return JSON.stringify({
    ok: false,
    error: 'Permintaan tidak valid.'
  });
  resetMemEksekusi_();
  return apiBridge_({
    action: String(action || ''),
    args: args
  }).getContent();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAppVersion() {
  return APP_VERSION;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Cycle Count').addItem('Setup Awal (buat semua sheet, facility & admin pertama)', 'setupAwal').addItem('Setup Setting Level Rack (Master)', 'setupMasterLevelSettings').addItem('Setup Mapping PIC Kategori Selisih (Master)', 'setupMasterAssignmentMapping').addItem('Setup Config Overflow (Config_Sistem & Config_Akses_Setting)', 'setupOverflowConfigSheets').addItem('Setup Master DB & Config DB (Pindah ke Spreadsheet Terpisah)', 'setupExternalDbMasterDanConfig').addItem('Setup Facility Awal (Seed RDC Makassar dari sistem lama)', 'seedFacilityAwalDariSistemLama').addItem('Perbaiki Header (Data Count & Riwayat)', 'perbaikiHeader').addItem('Cek & Perbaiki Waktu_Input Kosong', 'cekPerbaikiWaktuInputKosong').addItem('Cek & Jalankan Archive Riwayat Lama', 'cekArchiveRiwayatLama').addItem('Rebuild Ringkasan Harian (kalau Dashboard terasa tidak sinkron)', 'rebuildRingkasanHarianMenu').addItem('Rebuild Antrian Aktif (kalau Home terasa tidak sinkron)', 'rebuildAntrianAktifMenu').addItem('Reset Semua Hasil Count Hari Ini', 'resetCount').addToUi();
}

function onEdit(e) {
  try {
    if (e && e.range && e.range.getSheet().getName() === MASTER_SHEET_NAME) {
      clearMasterCache_();
    }
    if (e && e.range && e.range.getSheet().getName() === CONFIG_SISTEM_SHEET_NAME) {
      clearConfigSistemCache_();
    }
    if (e && e.range && e.range.getSheet().getName() === CONFIG_AKSES_SETTING_SHEET_NAME) {
      clearNikAksesSettingCache_();
    }
  } catch (err) {}
}

function debugCekAksesConfig() {
  const nik = SETUP_ADMIN_NIK_;
  Logger.log('--- Cek Script Properties ---');
  Logger.log('MASTER_DB_ID: ' + PropertiesService.getScriptProperties().getProperty('MASTER_DB_ID'));
  Logger.log('CONFIG_DB_ID: ' + PropertiesService.getScriptProperties().getProperty('CONFIG_DB_ID'));
  Logger.log('--- Cek Master DB sheet Master ---');
  const masterSheet = getMasterSheet_();
  Logger.log('getMasterSheet_() ketemu: ' + !!masterSheet + (masterSheet ? ' (di spreadsheet: ' + masterSheet.getParent().getUrl() + ')' : ''));
  Logger.log('--- Cek sheet Config_Akses_Setting ---');
  const aksesSheet = getConfigAksesSettingSheet_();
  Logger.log('getConfigAksesSettingSheet_() di spreadsheet: ' + aksesSheet.getParent().getUrl());
  Logger.log('Isi baris 2 dst: ' + JSON.stringify(aksesSheet.getRange(2, 1, Math.max(aksesSheet.getLastRow() - 1, 0), 3).getValues()));
  Logger.log('--- Cek isNikPunyaAksesSetting_("' + nik + '") ---');
  Logger.log('Hasil: ' + isNikPunyaAksesSetting_(nik));
  Logger.log('--- Cek getUserRole("' + nik + '") lengkap ---');
  Logger.log(JSON.stringify(getUserRole(nik)));
}

function clearMasterCache_() {
  resetMemEksekusi_();
  const kunci = [ 'activeLocations', 'masterUsernames', 'assignableUsers', 'levelSettings', 'kategoriMapping', 'inventoryUsers' ];
  try {
    const fs = getFacilitySheet_();
    const last = fs.getLastRow();
    if (last >= 2) {
      fs.getRange(2, 1, last - 1, 1).getValues().forEach(function(r) {
        const id = String(r[0] || '').trim();
        if (id) {
          kunci.push('assignableUsers_fac_' + id);
          kunci.push('inventoryUsers_fac_' + id);
        }
      });
    }
  } catch (e) {}
  CacheService.getScriptCache().removeAll(kunci);
  clearFacilityCache_();
}

function requireRole_(username, allowedRoles) {
  const info = getUserRole(username);
  if (!info) {
    throw new Error('User tidak dikenali. Silakan login ulang.');
  }
  if (allowedRoles && allowedRoles.indexOf(info.role) === -1) {
    throw new Error('Akses ditolak untuk role "' + info.role + '".');
  }
  return info;
}

function apiResult_(raw) {
  if (raw && typeof raw === 'object' && !Array.isArray(raw) && Object.prototype.hasOwnProperty.call(raw, 'success')) {
    return raw;
  }
  return {
    success: true,
    data: raw
  };
}

const API_ACTIONS = {
  login: function(req) {
    const info = getUserRole(req.username);
    if (info) info.aksesSettingOverflow = isNikPunyaAksesSetting_(req.username);
    return info;
  },
  getAppVersion: function() {
    return getAppVersion();
  },
  getHomeSummary: function(req) {
    return getHomeSummary(req.username);
  },
  getPlusMinusSummary: function(req) {
    return getPlusMinusSummary(req.username);
  },
  getMyPendingCount: function(req) {
    return getMyPendingCount(req.username);
  },
  getPendingValidasiCount: function(req) {
    return getPendingValidasiCount(req.username);
  },
  getMyPendingTasks: function(req) {
    return getMyPendingTasks(req.username);
  },
  submitCount: function(req) {
    return submitCount(req.no, req.username, req.qty);
  },
  getPendingValidasi: function(req) {
    return getPendingValidasi(req.username);
  },
  submitValidasi: function(req) {
    return submitValidasi(req.id, req.username, req.qty);
  },
  getEquipmentReadyDefaults: function(req) {
    return getEquipmentReadyDefaults(req.username);
  },
  getAssignableUsers: function(req) {
    return getAssignableUsers(req.username);
  },
  importRawData: function(req) {
    return importRawData(req.transaksiRawRows, req.stockRows, req.selectedUsernames, req.username, req.reachTruckReady, req.tanggaReady, req.jobId);
  },
  getImportProgress: function(req) {
    return getImportProgress(req.jobId);
  },
  getDashboardData: function(req) {
    return getDashboardData(req.periodType, req.periodValue, req.trendCount, req.username);
  },
  getErrorAnalysisDetail: function(req) {
    return getErrorAnalysisDetail(req.alasan, req.periodType, req.periodValue, req.username);
  },
  getUserDashboardDetail: function(req) {
    return getUserDashboardDetail(req.nama, req.periodType, req.periodValue, req.username);
  },
  getPendingBacklog: function(req) {
    return getPendingBacklog(req.username);
  },
  getBacklogDetailByDate: function(req) {
    return getBacklogDetailByDate(req.tanggal, req.username);
  },
  getProblemItemsDetail: function(req) {
    return getProblemItemsDetail(req.periodType, req.periodValue, req.username);
  },
  getAnalyticsRootCauseData: function(req) {
    return getAnalyticsRootCauseData(req.dateFrom, req.dateTo, req.username);
  },
  getAnalyticsRootCauseDetail: function(req) {
    return getAnalyticsRootCauseDetail(req.rootCauseLabel, req.dateFrom, req.dateTo, req.username);
  },
  getInvestigasiFormData: function(req) {
    return getInvestigasiFormData(req.username);
  },
  getOpenTasks: function(req) {
    return getOpenTasks(req.username);
  },
  updateTaskStatus: function(req) {
    return updateTaskStatus(req.id, req.username, req.newStatus, req.catatan, req.kategori, req.picUsername, req.buktiRows);
  },
  getTaskLog: function(req) {
    return getTaskLog(req.limit, req.username, req.dateFrom, req.dateTo);
  },
  getPlusMinusCandidates: function(req) {
    return getPlusMinusCandidates(req.username);
  },
  closePlusMinusPair: function(req) {
    return closePlusMinusPair(req.minusTaskId, req.plusTaskId, req.username, req.catatan, req.picUsername, req.buktiRows);
  },
  getSettingOverflow: function(req) {
    return getSettingOverflow(req.username);
  },
  saveSettingOverflow: function(req) {
    return saveSettingOverflow(req.username, req.overflowEnabled, req.minSisaReachTruck, req.maksOrang);
  },
  setOverflowForceOffToday: function(req) {
    return setOverflowForceOffToday(req.username);
  },
  getDaftarUserMaster: function(req) {
    return getDaftarUserMaster(req.username);
  },
  tambahUserMaster: function(req) {
    return tambahUserMaster(req.username, req.usernameBaru, req.roleBaru);
  },
  updateRoleUserMaster: function(req) {
    return updateRoleUserMaster(req.username, req.rowIndex, req.usernameKonfirmasi, req.roleBaru);
  },
  setStatusUserMaster: function(req) {
    return setStatusUserMaster(req.username, req.rowIndex, req.usernameKonfirmasi, req.statusBaru);
  },
  getDaftarAksesSetting: function(req) {
    return getDaftarAksesSetting(req.username);
  },
  tambahAksesSetting: function(req) {
    return tambahAksesSetting(req.username, req.nikBaru, req.namaBaru);
  },
  hapusAksesSetting: function(req) {
    return hapusAksesSetting(req.username, req.rowIndex, req.nikKonfirmasi);
  },
  getLevelAssignmentConfig: function(req) {
    return getLevelAssignmentConfig(req.username);
  },
  saveLevelAssignmentMatrix: function(req) {
    return saveLevelAssignmentMatrix(req.username, req.matrixBaru, req.maksGrupAlatBaru);
  },
  setModeAssignment: function(req) {
    return setModeAssignment(req.username, req.mode);
  },
  getLogPerubahanConfig: function(req) {
    return getLogPerubahanConfig(req.username);
  },
  getDaftarFacility: function(req) {
    return getDaftarFacility(req.username);
  },
  tambahFacility: function(req) {
    return tambahFacility(req.username, req.namaFacility, req.kodeFacility, req.namaSpreadsheet, req.daftarLokasi, req.jobId);
  },
  daftarkanFacilityExisting: function(req) {
    return daftarkanFacilityExisting(req.username, req.namaFacility, req.kodeFacility, req.spreadsheetIdAtauUrl, req.daftarLokasi, req.jobId);
  },
  updateNamaFacility: function(req) {
    return updateNamaFacility(req.username, req.facilityId, req.namaBaru, req.kodeBaru);
  },
  setStatusFacility: function(req) {
    return setStatusFacility(req.username, req.facilityId, req.statusBaru);
  },
  assignUserKeFacility: function(req) {
    return assignUserKeFacility(req.username, req.targetUsername, req.facilityId);
  },
  getDaftarUserFacilityAssignment: function(req) {
    return getDaftarUserFacilityAssignment(req.username);
  },
  importLokasiAktif: function(req) {
    return importLokasiAktif(req.username, req.facilityId, req.daftarLokasi, req.gantiSemua);
  },
  getDaftarLokasiAktif: function(req) {
    return getDaftarLokasiAktif(req.username, req.facilityId);
  },
  copyLokasiDariFacility: function(req) {
    return copyLokasiDariFacility(req.username, req.facilityIdSumber, req.facilityIdTarget, req.gantiSemua);
  },
  setDeveloperActiveFacility: function(req) {
    return setDeveloperActiveFacility(req.username, req.facilityId);
  }
};

function doPost(e) {
  try {
    resetMemEksekusi_();
    const req = JSON.parse(e.postData.contents || '{}');
    const viaApk = apiBridge_(req);
    if (viaApk) return viaApk;
    const action = req.action;
    const handler = API_ACTIONS[action];
    if (!handler) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'Unknown action: ' + action
      })).setMimeType(ContentService.MimeType.JSON);
    }
    const result = apiResult_(handler(req));
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function isBarisUserAktif_(statusRaw) {
  const s = String(statusRaw || '').trim().toLowerCase();
  return s !== 'nonaktif';
}

function getActiveLocations_(username) {
  const cache = CacheService.getScriptCache();
  const facCache = username ? getUserFacility(username) : null;
  const cacheKey = facCache ? 'activeLocations_fac_' + facCache.id : username ? 'activeLocations_' + String(username).toLowerCase() : 'activeLocations';
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);
  const set = {};
  if (username) {
    try {
      const lokasiSheet = getLokasiAktifSheet_(username);
      if (lokasiSheet) {
        const lastRow = lokasiSheet.getLastRow();
        if (lastRow >= 2) {
          const values = lokasiSheet.getRange(2, 1, lastRow - 1, 1).getValues();
          values.forEach(function(r) {
            const loc = String(r[0] || '').trim();
            if (loc) set[loc] = true;
          });
        }
      }
    } catch (e) {}
  }
  if (Object.keys(set).length === 0) {
    const sheet = getMasterSheet_();
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        values.forEach(function(r) {
          const loc = String(r[0] || '').trim();
          if (loc) set[loc] = true;
        });
      }
    }
  }
  try {
    cache.put(cacheKey, JSON.stringify(set), 300);
  } catch (e) {}
  return set;
}

function clearActiveLocationsCache_(facilityId) {
  try {
    CacheService.getScriptCache().removeAll([ 'activeLocations', 'activeLocations_fac_' + facilityId ]);
  } catch (e) {}
}

var _memUserRoleCache_ = {};

var _memIsDeveloper_ = {};

var _memUserFacility_ = {};

function resetMemEksekusi_() {
  _memUserRoleCache_ = {};
  _memIsDeveloper_ = {};
  _memUserFacility_ = {};
}

function clearUserRoleCache_(username) {
  try {
    const uname = String(username || '').trim().toLowerCase();
    delete _memUserRoleCache_[uname];
    delete _memIsDeveloper_[uname];
    delete _memUserFacility_[uname];
    CacheService.getScriptCache().remove('userRole_' + uname);
  } catch (e) {}
}

function getUserRole(username) {
  const uname = String(username || '').trim().toLowerCase();
  if (!uname) return null;
  if (_memUserRoleCache_[uname]) {
    return _memUserRoleCache_[uname];
  }
  const cacheKey = 'userRole_' + uname;
  try {
    const cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) {
      const parsed = JSON.parse(cached);
      _memUserRoleCache_[uname] = parsed;
      return parsed;
    }
  } catch (e) {}
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    const sheet = getMasterUserSheetWithFallback_();
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
        for (let i = 0; i < values.length; i++) {
          const u = String(values[i][0] || '').trim();
          if (u.toLowerCase() === uname) {
            if (!isBarisUserAktif_(values[i][2])) return null;
            const roleRaw = String(values[i][1] || '').trim().toLowerCase();
            let role = roleRaw;
            if (roleRaw === 'adm' || roleRaw === 'administrator') role = 'admin'; else if (roleRaw === 'inventory') role = 'inventory'; else if (roleRaw === 'cycle') role = 'cycle'; else if (roleRaw === 'outbound') role = 'outbound'; else if (roleRaw === 'storing') role = 'storing'; else if (roleRaw === 'inbound') role = 'inbound'; else if (roleRaw === 'developer' || roleRaw === 'dev' || roleRaw === 'dewa') role = 'developer';
            let facInfo = null;
            const facId = String(values[i][3] || '').trim();
            if (facId && typeof getFacilityInfoById_ === 'function') {
              facInfo = getFacilityInfoById_(facId);
            }
            if (role === 'developer' && typeof getUserFacility === 'function') {
              const facDev = getUserFacility(u);
              if (facDev) facInfo = facDev;
            }
            const result = {
              role: role,
              displayName: u,
              aksesSettingOverflow: isNikPunyaAksesSetting_(u)
            };
            if (facInfo) {
              result.facilityId = facInfo.id;
              result.facilityName = facInfo.nama;
              result.facilityCode = facInfo.kode;
              result.facilityStatus = facInfo.status;
              result.facilityUrl = facInfo.url;
            }
            _memUserRoleCache_[uname] = result;
            try {
              CacheService.getScriptCache().put('userRole_' + uname, JSON.stringify(result), 300);
            } catch (e) {}
            return result;
          }
        }
      }
    }
  }
  const sheetOld = getMasterSheet_();
  if (!sheetOld) return null;
  const lastRow = sheetOld.getLastRow();
  if (lastRow < 2) return null;
  const values = sheetOld.getRange(2, 3, lastRow - 1, 3).getValues();
  for (let i = 0; i < values.length; i++) {
    const u = String(values[i][0] || '').trim();
    if (u.toLowerCase() === uname) {
      if (!isBarisUserAktif_(values[i][2])) return null;
      const roleRaw = String(values[i][1] || '').trim().toLowerCase();
      let role = roleRaw;
      if (roleRaw === 'adm' || roleRaw === 'administrator') role = 'admin'; else if (roleRaw === 'inventory') role = 'inventory'; else if (roleRaw === 'cycle') role = 'cycle'; else if (roleRaw === 'outbound') role = 'outbound'; else if (roleRaw === 'storing') role = 'storing'; else if (roleRaw === 'inbound') role = 'inbound'; else if (roleRaw === 'developer' || roleRaw === 'dev' || roleRaw === 'dewa') role = 'developer';
      let facInfo = null;
      if (typeof getUserFacilityAssignmentSheet_ === 'function') {
        const assignSheet = getUserFacilityAssignmentSheet_();
        if (assignSheet && assignSheet.getLastRow() >= 2) {
          const assignData = assignSheet.getRange(2, 1, assignSheet.getLastRow() - 1, 2).getValues();
          for (let j = assignData.length - 1; j >= 0; j--) {
            if (String(assignData[j][0]).trim().toLowerCase() === u.toLowerCase()) {
              const fid = String(assignData[j][1]).trim();
              if (fid && typeof getFacilityInfoById_ === 'function') {
                facInfo = getFacilityInfoById_(fid);
              }
              break;
            }
          }
        }
      }
      const result = {
        role: role,
        displayName: u,
        aksesSettingOverflow: isNikPunyaAksesSetting_(u)
      };
      if (facInfo) {
        result.facilityId = facInfo.id;
        result.facilityName = facInfo.nama;
        result.facilityCode = facInfo.kode;
        result.facilityStatus = facInfo.status;
        result.facilityUrl = facInfo.url;
      }
      _memUserRoleCache_[uname] = result;
      try {
        CacheService.getScriptCache().put('userRole_' + uname, JSON.stringify(result), 300);
      } catch (e) {}
      return result;
    }
  }
  return null;
}

function getMasterUsernames() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('masterUsernames');
  if (cached) return JSON.parse(cached);
  let list = [];
  let useLegacy = true;
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    const sheet = getMasterUserSheetWithFallback_();
    if (sheet && sheet.getLastRow() >= 2) {
      useLegacy = false;
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
      list = values.filter(function(r) {
        return isBarisUserAktif_(r[2]);
      }).map(function(r) {
        return String(r[0] || '').trim();
      }).filter(function(v) {
        return v;
      });
    }
  }
  if (useLegacy) {
    const sheet = getMasterSheet_();
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const values = sheet.getRange(2, 3, lastRow - 1, 3).getValues();
        list = values.filter(function(r) {
          return isBarisUserAktif_(r[2]);
        }).map(function(r) {
          return String(r[0] || '').trim();
        }).filter(function(v) {
          return v;
        });
      }
    }
  }
  cache.put('masterUsernames', JSON.stringify(list), 300);
  return list;
}

function getAssignableUsers_(facilityId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = facilityId ? 'assignableUsers_fac_' + facilityId : 'assignableUsers';
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);
  let list = [];
  let useLegacy = true;
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    const sheet = getMasterUserSheetWithFallback_();
    if (sheet && sheet.getLastRow() >= 2) {
      useLegacy = false;
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
      list = values.map(function(r) {
        if (!isBarisUserAktif_(r[2])) return null;
        const username = String(r[0] || '').trim();
        const roleRaw = String(r[1] || '').trim().toLowerCase();
        const facId = String(r[3] || '').trim();
        if (roleRaw === 'developer' || roleRaw === 'dev' || roleRaw === 'dewa') return null;
        if (facilityId && facId !== facilityId) return null;
        let role = roleRaw;
        if (roleRaw === 'adm' || roleRaw === 'administrator') role = 'admin';
        return username ? {
          username: username,
          role: role
        } : null;
      }).filter(function(v) {
        return v;
      });
    }
  }
  if (useLegacy) {
    const sheet = getMasterSheet_();
    let userDiFacility = null;
    if (facilityId) {
      try {
        const assignSheet = getUserFacilityAssignmentSheet_();
        const lastRowAssign = assignSheet.getLastRow();
        if (lastRowAssign >= 2) {
          const assignValues = assignSheet.getRange(2, 1, lastRowAssign - 1, 2).getValues();
          userDiFacility = {};
          assignValues.forEach(function(r) {
            const uname = String(r[0] || '').trim().toLowerCase();
            const facId = String(r[1] || '').trim();
            if (uname && facId === facilityId) {
              userDiFacility[uname] = true;
            }
          });
        }
      } catch (e) {
        userDiFacility = null;
      }
    }
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const values = sheet.getRange(2, 3, lastRow - 1, 3).getValues();
        list = values.map(function(r) {
          if (!isBarisUserAktif_(r[2])) return null;
          const username = String(r[0] || '').trim();
          const roleRaw = String(r[1] || '').trim().toLowerCase();
          if (roleRaw === 'developer' || roleRaw === 'dev' || roleRaw === 'dewa') return null;
          if (userDiFacility && !userDiFacility[username.toLowerCase()]) return null;
          let role = roleRaw;
          if (roleRaw === 'adm' || roleRaw === 'administrator') role = 'admin';
          return username ? {
            username: username,
            role: role
          } : null;
        }).filter(function(v) {
          return v;
        });
      }
    }
  }
  try {
    cache.put(cacheKey, JSON.stringify(list), 300);
  } catch (e) {}
  return list;
}

function getAssignableUsers(requesterUsername) {
  requireRole_(requesterUsername, [ 'admin', 'developer' ]);
  const facInfo = getUserFacility(requesterUsername);
  const facilityId = facInfo ? facInfo.id : null;
  return getAssignableUsers_(facilityId);
}

function getUserRoleMap_() {
  const list = getAssignableUsers_();
  const map = {};
  list.forEach(function(u) {
    let role = String(u.role || '').trim().toLowerCase();
    if (role === 'adm' || role === 'administrator') role = 'admin';
    map[u.username.trim().toLowerCase()] = role;
  });
  return map;
}

function getInventoryUsernames_(facilityId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = facilityId ? 'inventoryUsers_fac_' + facilityId : 'inventoryUsers';
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);
  const allAssignable = getAssignableUsers_(facilityId);
  const list = allAssignable.filter(function(u) {
    return String(u.role || '').trim().toLowerCase() === 'inventory';
  }).map(function(u) {
    return u.username;
  });
  try {
    cache.put(cacheKey, JSON.stringify(list), 300);
  } catch (e) {}
  return list;
}

function assignNextValidator_(facilityId) {
  const users = getInventoryUsernames_(facilityId);
  if (!users.length) return '';
  const cache = CacheService.getScriptCache();
  const cacheKey = facilityId ? 'validasiRoundRobinPointer_' + facilityId : 'validasiRoundRobinPointer';
  let pointer = Number(cache.get(cacheKey) || '0');
  if (isNaN(pointer) || pointer < 0) pointer = 0;
  const username = users[pointer % users.length];
  try {
    cache.put(cacheKey, String(pointer + 1), 21600);
  } catch (e) {}
  return username;
}

function getInvestigationOwnerPool_(facilityId) {
  const allAssignable = getAssignableUsers_(facilityId);
  const list = allAssignable.filter(function(u) {
    const role = String(u.role || '').trim().toLowerCase();
    return role === 'inventory' || role === 'admin';
  }).map(function(u) {
    return u.username;
  });
  return list;
}

function assignInitialPIC_(facilityId) {
  const users = getInvestigationOwnerPool_(facilityId);
  if (!users.length) return '';
  const props = PropertiesService.getScriptProperties();
  const propKey = facilityId ? 'picRoundRobinPointer_' + facilityId : 'picRoundRobinPointer';
  let pointer = Number(props.getProperty(propKey) || '0');
  if (isNaN(pointer) || pointer < 0) pointer = 0;
  const username = users[pointer % users.length];
  props.setProperty(propKey, String(pointer + 1));
  return username;
}

function getKategoriMapping_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('kategoriMapping');
  if (cached) return JSON.parse(cached);
  const sheet = getMasterSheet_();
  let list = [];
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 9, lastRow - 1, 3).getValues();
      values.forEach(function(r) {
        const kategori = String(r[0] || '').trim();
        const username = String(r[1] || '').trim();
        const role = String(r[2] || '').trim();
        if (kategori && username) list.push({
          kategori: kategori,
          username: username,
          role: role
        });
      });
    }
  }
  cache.put('kategoriMapping', JSON.stringify(list), 300);
  return list;
}

function getInvestigasiFormData(requesterUsername) {
  requireRole_(requesterUsername, [ 'inventory', 'admin', 'developer' ]);
  const sheet = getMasterSheet_();
  let allUsers = [];
  const facInfo = getUserFacility(requesterUsername);
  const facilityId = facInfo ? facInfo.id : null;
  allUsers = getAssignableUsers_(facilityId);
  return {
    kategoriList: TASK_KATEGORI_LIST,
    mapping: getKategoriMapping_(),
    allUsers: allUsers,
    statusFlow: TASK_STATUS_FLOW,
    slaHari: TASK_SLA_HARI
  };
}

function getLevelSettings_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('levelSettings');
  if (cached) return JSON.parse(cached);
  const prod = {};
  LEVEL_GROUPS.forEach(function(g) {
    prod[g.key] = g.defaultProd;
  });
  const sheet = getMasterSheet_();
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 6, lastRow - 1, 2).getValues();
      data.forEach(function(row) {
        const label = String(row[0] || '').trim();
        const value = row[1];
        if (!label || value === '' || isNaN(value)) return;
        const g = LEVEL_GROUPS.filter(function(gr) {
          return label.indexOf(gr.min + '-' + gr.max) !== -1;
        })[0];
        if (g) prod[g.key] = Number(value);
      });
    }
  }
  cache.put('levelSettings', JSON.stringify(prod), 300);
  return prod;
}

function getConfigSistemSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SISTEM_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SISTEM_SHEET_NAME);
    sheet.appendRow(CONFIG_SISTEM_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, CONFIG_SISTEM_HEADERS);
  }
  if (sheet.getLastRow() < 2) {
    seedConfigSistemDefaults_(sheet);
  }
  return sheet;
}

function seedConfigSistemDefaults_(sheet) {
  const now = new Date;
  const rows = Object.keys(CONFIG_SISTEM_DEFAULTS_).map(function(key) {
    return [ key, CONFIG_SISTEM_DEFAULTS_[key], CONFIG_SISTEM_KETERANGAN_[key] || '', now, 'system' ];
  });
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, CONFIG_SISTEM_HEADERS.length).setValues(rows);
  }
}

function getConfigSistem_(username) {
  const cache = CacheService.getScriptCache();
  const facPrefix = username ? function() {
    const f = getUserFacility(username);
    return f ? f.id : 'legacy';
  }() : 'legacy';
  const cacheKey = 'configSistem_' + facPrefix;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);
  const sheet = getConfigSistemSheet_(username);
  const lastRow = sheet.getLastRow();
  const result = {};
  Object.keys(CONFIG_SISTEM_DEFAULTS_).forEach(function(k) {
    result[k] = CONFIG_SISTEM_DEFAULTS_[k];
  });
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    data.forEach(function(row) {
      const key = String(row[0] || '').trim();
      if (!key) return;
      result[key] = row[1] === null || row[1] === undefined ? '' : String(row[1]).trim();
    });
  }
  try {
    cache.put(cacheKey, JSON.stringify(result), 300);
  } catch (e) {}
  return result;
}

function clearConfigSistemCache_(username) {
  try {
    const facPrefix = username ? function() {
      const f = getUserFacility(username);
      return f ? f.id : 'legacy';
    }() : 'legacy';
    CacheService.getScriptCache().remove('configSistem_' + facPrefix);
  } catch (e) {}
}

function setConfigSistemValue_(key, value, username) {
  const sheet = getConfigSistemSheet_(username);
  const lastRow = sheet.getLastRow();
  const now = new Date;
  let targetRow = 0;
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0] || '').trim() === key) {
        targetRow = i + 2;
        break;
      }
    }
  }
  if (!targetRow) {
    targetRow = lastRow + 1;
    sheet.getRange(targetRow, 1).setValue(key);
    sheet.getRange(targetRow, 3).setValue(CONFIG_SISTEM_KETERANGAN_[key] || '');
  }
  sheet.getRange(targetRow, 2).setValue(value);
  sheet.getRange(targetRow, 4).setValue(now);
  sheet.getRange(targetRow, 5).setValue(username || '');
  clearConfigSistemCache_(username);
}

function getConfigAksesSettingSheet_() {
  const ss = getConfigTargetSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG_AKSES_SETTING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_AKSES_SETTING_SHEET_NAME);
    sheet.appendRow(CONFIG_AKSES_SETTING_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, CONFIG_AKSES_SETTING_HEADERS);
  }
  if (sheet.getLastRow() < 2) {
    sheet.appendRow([ SETUP_ADMIN_NIK_, '', Utilities.formatDate(new Date, 'Asia/Jakarta', 'yyyy-MM-dd') ]);
  }
  return sheet;
}

function getNikAksesSettingList_() {
  const sheet = getConfigAksesSettingSheet_();
  const lastRow = sheet.getLastRow();
  const list = [];
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    data.forEach(function(row) {
      const nik = String(row[0] || '').trim();
      if (nik) list.push(nik.toLowerCase());
    });
  }
  return list;
}

function clearNikAksesSettingCache_() {}

function isNikPunyaAksesSetting_(username) {
  const nik = String(username || '').trim().toLowerCase();
  if (!nik) return false;
  if (isUsernameRoleDeveloper_(nik)) return true;
  return getNikAksesSettingList_().indexOf(nik) !== -1;
}

function isUsernameRoleDeveloper_(usernameLowercase) {
  if (Object.prototype.hasOwnProperty.call(_memIsDeveloper_, usernameLowercase)) return _memIsDeveloper_[usernameLowercase];
  const peranDev = function(roleRaw) {
    const r = String(roleRaw || '').trim().toLowerCase();
    return r === 'developer' || r === 'dev' || r === 'dewa';
  };
  let hasil = false, ketemu = false;
  const mu = typeof getMasterUserSheetWithFallback_ === 'function' ? getMasterUserSheetWithFallback_() : null;
  const muLast = mu ? mu.getLastRow() : 0;
  if (muLast >= 2) {
    const v = mu.getRange(2, 1, muLast - 1, 3).getValues();
    for (let i = 0; i < v.length; i++) {
      if (String(v[i][0] || '').trim().toLowerCase() === usernameLowercase) {
        ketemu = true;
        hasil = isBarisUserAktif_(v[i][2]) && peranDev(v[i][1]);
        break;
      }
    }
  }
  if (!ketemu) {
    const sheet = getMasterSheet_();
    const lastRow = sheet ? sheet.getLastRow() : 0;
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 3, lastRow - 1, 3).getValues();
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][0] || '').trim().toLowerCase() === usernameLowercase) {
          hasil = isBarisUserAktif_(values[i][2]) && peranDev(values[i][1]);
          break;
        }
      }
    }
  }
  _memIsDeveloper_[usernameLowercase] = hasil;
  return hasil;
}

function getDaftarAksesSetting(requesterUsername) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke menu Akses Setting.'
    };
  }
  const sheet = getConfigAksesSettingSheet_();
  const lastRow = sheet.getLastRow();
  const daftar = [];
  if (lastRow >= 2) {
    let tz = 'Asia/Jakarta';
    try {
      tz = sheet.getParent().getSpreadsheetTimeZone() || tz;
    } catch (e) {}
    const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    values.forEach(function(r, idx) {
      const nik = String(r[0] || '').trim();
      if (!nik) return;
      daftar.push({
        rowIndex: idx + 2,
        nik: nik,
        nama: String(r[1] || '').trim(),
        tanggal: Object.prototype.toString.call(r[2]) === '[object Date]' ? Utilities.formatDate(r[2], tz, 'yyyy-MM-dd') : String(r[2] || '').trim()
      });
    });
  }
  return {
    success: true,
    daftar: daftar
  };
}

function tambahAksesSetting(requesterUsername, nikBaru, namaBaru) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke menu Akses Setting.'
    };
  }
  const nik = String(nikBaru || '').trim();
  if (!nik) return {
    success: false,
    message: 'NIK tidak boleh kosong.'
  };
  const sheet = getConfigAksesSettingSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < existing.length; i++) {
      if (String(existing[i][0] || '').trim().toLowerCase() === nik.toLowerCase()) {
        return {
          success: false,
          message: 'NIK "' + nik + '" sudah terdaftar.'
        };
      }
    }
  }
  sheet.appendRow([ nik, String(namaBaru || '').trim(), todayJakarta_() ]);
  clearUserRoleCache_(nik);
  catatLogPerubahanConfig_(requesterUsername, 'Akses Setting', 'Tambah akses NIK "' + nik + '".');
  return {
    success: true,
    message: 'NIK "' + nik + '" berhasil ditambahkan ke akses Config.'
  };
}

function hapusAksesSetting(requesterUsername, rowIndex, nikKonfirmasi) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke menu Akses Setting.'
    };
  }
  const sheet = getConfigAksesSettingSheet_();
  const actualNik = String(sheet.getRange(rowIndex, 1).getValue() || '').trim();
  if (!actualNik || actualNik.toLowerCase() !== String(nikKonfirmasi || '').trim().toLowerCase()) {
    return {
      success: false,
      message: 'Data sudah berubah sejak halaman dibuka. Refresh dulu, lalu coba lagi.'
    };
  }
  if (actualNik.toLowerCase() === String(requesterUsername || '').trim().toLowerCase()) {
    return {
      success: false,
      message: 'Tidak bisa mencabut akses NIK Anda sendiri yang sedang login. Minta NIK lain yang punya akses untuk melakukan ini.'
    };
  }
  sheet.deleteRow(rowIndex);
  clearUserRoleCache_(actualNik);
  catatLogPerubahanConfig_(requesterUsername, 'Akses Setting', 'Cabut akses NIK "' + actualNik + '".');
  return {
    success: true,
    message: 'Akses NIK "' + actualNik + '" dicabut.'
  };
}

function setupOverflowConfigSheets() {
  getConfigSistemSheet_();
  getConfigAksesSettingSheet_();
  clearConfigSistemCache_();
  clearNikAksesSettingCache_();
  const pesan = 'Sheet "Config_Sistem" & "Config_Akses_Setting" siap. NIK ' + SETUP_ADMIN_NIK_ + ' sudah didaftarkan sebagai akses pertama menu Setting Overflow.';
  try {
    SpreadsheetApp.getUi().alert(pesan);
  } catch (e) {
    Logger.log(pesan);
  }
}

function getSettingOverflow(requesterUsername) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Setting Overflow.'
    };
  }
  const cfg = getConfigSistem_(requesterUsername);
  return {
    success: true,
    overflowEnabled: String(cfg.overflow_enabled).toUpperCase() === 'TRUE',
    minSisaReachTruck: Number(cfg.overflow_min_sisa_reach_truck) || 0,
    maksOrang: cfg.overflow_maks_orang === '' || cfg.overflow_maks_orang === undefined ? null : Number(cfg.overflow_maks_orang) || 0,
    forceOffToday: cfg.overflow_force_off_today === todayJakarta_(),
    keterangan: CONFIG_SISTEM_KETERANGAN_
  };
}

function saveSettingOverflow(requesterUsername, overflowEnabled, minSisaReachTruck, maksOrang) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Setting Overflow.'
    };
  }
  setConfigSistemValue_('overflow_enabled', overflowEnabled ? 'TRUE' : 'FALSE', requesterUsername);
  setConfigSistemValue_('overflow_min_sisa_reach_truck', String(Math.max(0, Number(minSisaReachTruck) || 0)), requesterUsername);
  setConfigSistemValue_('overflow_maks_orang', maksOrang === null || maksOrang === '' || maksOrang === undefined ? '' : String(Math.max(0, Number(maksOrang) || 0)), requesterUsername);
  return {
    success: true,
    message: 'Setting Overflow tersimpan.'
  };
}

function setOverflowForceOffToday(requesterUsername) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Setting Overflow.'
    };
  }
  setConfigSistemValue_('overflow_force_off_today', todayJakarta_(), requesterUsername);
  return {
    success: true,
    message: 'Overflow dinonaktifkan untuk hari ini saja. Besok otomatis aktif lagi (kalau setting permanen ON).'
  };
}

function todayJakarta_() {
  return Utilities.formatDate(new Date, 'Asia/Jakarta', 'yyyy-MM-dd');
}

function applyOverflowStoringKeTangga_(reachTruckItemsCount, reachTruckUsers, tanggaUsers, username) {
  const result = {
    reachTruckUsers: reachTruckUsers,
    tanggaUsers: tanggaUsers,
    overflowCount: 0,
    pesan: ''
  };
  if (!reachTruckUsers.length) return result;
  const cfg = getConfigSistem_(username);
  const enabled = String(cfg.overflow_enabled).toUpperCase() === 'TRUE';
  if (!enabled) return result;
  if (cfg.overflow_force_off_today === todayJakarta_()) {
    result.pesan = 'Overflow OFF hari ini (dinonaktifkan manual lewat tombol darurat).';
    return result;
  }
  const prodPerOrang = getEquipmentReadyDefaults_().reachTruck || 0;
  if (!prodPerOrang) return result;
  const kapasitas = reachTruckUsers.length * prodPerOrang;
  if (reachTruckItemsCount >= kapasitas) return result;
  const kelebihanOrang = Math.floor((kapasitas - reachTruckItemsCount) / prodPerOrang);
  if (kelebihanOrang <= 0) return result;
  const minSisa = Math.max(0, Number(cfg.overflow_min_sisa_reach_truck) || 0);
  const maksBolehPindah = reachTruckUsers.length - minSisa;
  if (maksBolehPindah <= 0) return result;
  const batasCfg = cfg.overflow_maks_orang === '' || cfg.overflow_maks_orang === undefined ? Infinity : Number(cfg.overflow_maks_orang) || 0;
  const jumlahPindah = Math.min(kelebihanOrang, maksBolehPindah, batasCfg);
  if (jumlahPindah <= 0) return result;
  const dipindah = reachTruckUsers.slice(reachTruckUsers.length - jumlahPindah);
  const sisaReachTruck = reachTruckUsers.slice(0, reachTruckUsers.length - jumlahPindah);
  result.reachTruckUsers = sisaReachTruck;
  result.tanggaUsers = tanggaUsers.concat(dipindah);
  result.overflowCount = jumlahPindah;
  result.pesan = 'Overflow aktif: ' + jumlahPindah + ' user Storing dialihkan bantu Level 3-4 karena demand Level 5 & 6 rendah hari ini.';
  return result;
}

function testKecepatan50User() {
  const USERNAME_TEST = SETUP_ADMIN_NIK_;
  const JUMLAH_ITERASI = 10;
  const OFFSET_SLOT = 0;
  Logger.log('=== PERF TEST: ' + JUMLAH_ITERASI + ' sequential submit (offset ' + OFFSET_SLOT + ') ===');
  const facInfo = getUserFacility(USERNAME_TEST);
  if (!facInfo) {
    Logger.log('ERROR: User tidak punya facility. Ganti USERNAME_TEST.');
    return;
  }
  Logger.log('Facility: ' + facInfo.nama + ' (' + facInfo.id + ')');
  const sheet = getSheet_(USERNAME_TEST);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('Tidak ada data.');
    return;
  }
  const allRows = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const pendingNos = allRows.filter(function(r) {
    return String(r[9] || '').trim().toLowerCase() === USERNAME_TEST.toLowerCase() && r[10] === 'Pending';
  }).slice(OFFSET_SLOT, OFFSET_SLOT + JUMLAH_ITERASI).map(function(r) {
    return r[0];
  });
  if (!pendingNos.length) {
    Logger.log('Tidak ada item Pending untuk user ini di offset ' + OFFSET_SLOT + '. Upload data lebih banyak, atau kecilkan OFFSET_SLOT.');
    return;
  }
  const waktuMulai = Date.now();
  let berhasil = 0, gagal = 0;
  const detailMs = [];
  pendingNos.forEach(function(no, idx) {
    const t0 = Date.now();
    try {
      const res = submitCount(no, USERNAME_TEST, Math.floor(Math.random() * 100));
      const ms = Date.now() - t0;
      detailMs.push(ms);
      if (res && res.success) {
        berhasil++;
        Logger.log('[' + (idx + 1) + '] No ' + no + ' : OK (' + ms + ' ms)');
      } else {
        gagal++;
        Logger.log('[' + (idx + 1) + '] No ' + no + ' : FAIL ' + ms + ' ms - ' + (res ? res.message : ''));
      }
    } catch (e) {
      gagal++;
      detailMs.push(Date.now() - t0);
      Logger.log('[' + (idx + 1) + '] No ' + no + ' : ERROR - ' + e.message);
    }
  });
  const totalMs = Date.now() - waktuMulai;
  const rataMs = detailMs.length ? Math.round(detailMs.reduce(function(a, b) {
    return a + b;
  }, 0) / detailMs.length) : 0;
  Logger.log('');
  Logger.log('=== HASIL ===');
  Logger.log('Berhasil: ' + berhasil + ' / ' + pendingNos.length + ' | Gagal: ' + gagal);
  Logger.log('Total: ' + totalMs + ' ms | Rata-rata: ' + rataMs + ' ms/submit');
  Logger.log('Min: ' + (detailMs.length ? Math.min.apply(null, detailMs) : 0) + ' ms | Max: ' + (detailMs.length ? Math.max.apply(null, detailMs) : 0) + ' ms');
  Logger.log('Throughput: ~' + (berhasil / (totalMs / 1e3)).toFixed(2) + ' submit/detik (sequential)');
  Logger.log('Estimasi 50 user paralel (10-slot lock): ~' + Math.ceil(50 / 10 * rataMs / 1e3) + ' detik');
}

function testLockPerformance() {
  const FACILITY_ID = null;
  const JUMLAH = 20;
  Logger.log('=== LOCK TEST: ' + JUMLAH + ' iterasi, sleep 200ms/iterasi ===');
  const t0 = Date.now();
  let ok = 0;
  for (let i = 0; i < JUMLAH; i++) {
    const lt = Date.now();
    const lk = acquireFacilityLock_(FACILITY_ID, 5e3);
    if (lk) {
      ok++;
      Utilities.sleep(200);
      lk.release();
      Logger.log('#' + (i + 1) + ' OK (' + (Date.now() - lt) + ' ms)');
    } else {
      Logger.log('#' + (i + 1) + ' TIMEOUT (' + (Date.now() - lt) + ' ms)');
    }
  }
  Logger.log('Total: ' + (Date.now() - t0) + ' ms | OK: ' + ok + ' / ' + JUMLAH);
}

function queueRingkasanDeltaNewRow_(tanggal, newRow, username) {
  if (!tanggal) return;
  const sheet = getQueueRingkasanSheet_(username);
  appendRowLocked_(sheet, [ new Date, tanggal, JSON.stringify(newRow) ]);
}

function queueCounterIncrements_(items, username) {
  if (!items || !items.length) return;
  const sheet = getQueueCounterSheet_(username);
  const now = new Date;
  const rows = items.map(function(it) {
    return [ now, it.key, it.delta ];
  });
  appendRowsLocked_(sheet, rows);
}

function processDeferredQueueForFacility_(facilityUsernameHint) {
  const facInfo = getUserFacility(facilityUsernameHint);
  const facilityId = facInfo ? facInfo.id : null;
  const cache = CacheService.getScriptCache();
  const workerGuardKey = 'deferredQueueWorkerLock_' + (facilityId || 'default');
  if (cache.get(workerGuardKey)) return;
  try {
    cache.put(workerGuardKey, '1', 55);
  } catch (e) {}
  const facLock = acquireFacilityLock_(facilityId, 15e3);
  if (!facLock) {
    try {
      cache.remove(workerGuardKey);
    } catch (e) {}
    return;
  }
  try {
    const qrSheet = getQueueRingkasanSheet_(facilityUsernameHint);
    const qrLastRow = qrSheet.getLastRow();
    if (qrLastRow >= 2) {
      const qrData = qrSheet.getRange(2, 1, qrLastRow - 1, QUEUE_RINGKASAN_HEADERS.length).getValues();
      qrData.forEach(function(row) {
        const tanggal = row[1];
        try {
          const newRow = JSON.parse(row[2]);
          applyRingkasanDelta_(tanggal, null, newRow, facilityUsernameHint);
        } catch (e) {}
      });
      qrSheet.deleteRows(2, qrLastRow - 1);
    }
    const qcSheet = getQueueCounterSheet_(facilityUsernameHint);
    const qcLastRow = qcSheet.getLastRow();
    if (qcLastRow >= 2) {
      const qcData = qcSheet.getRange(2, 1, qcLastRow - 1, QUEUE_COUNTER_HEADERS.length).getValues();
      const totals = {};
      qcData.forEach(function(row) {
        const key = String(row[1] || '').trim();
        const delta = Number(row[2]) || 0;
        if (!key) return;
        totals[key] = (totals[key] || 0) + delta;
      });
      Object.keys(totals).forEach(function(key) {
        if (totals[key] !== 0) incrementCounterKey_(key, totals[key], facilityUsernameHint);
      });
      qcSheet.deleteRows(2, qcLastRow - 1);
    }
  } finally {
    facLock.release();
    try {
      cache.remove(workerGuardKey);
    } catch (e) {}
  }
}

function processDeferredQueueAllFacilities_() {
  const facSheet = getFacilitySheet_();
  const lastRow = facSheet.getLastRow();
  if (lastRow < 2) return;
  const facValues = facSheet.getRange(2, 1, lastRow - 1, FACILITY_HEADERS.length).getValues();
  const assignSheet = getUserFacilityAssignmentSheet_();
  const assignLastRow = assignSheet.getLastRow();
  const usernameByFacility = {};
  if (assignLastRow >= 2) {
    const assignValues = assignSheet.getRange(2, 1, assignLastRow - 1, 2).getValues();
    assignValues.forEach(function(row) {
      const uname = String(row[0] || '').trim();
      const facId = String(row[1] || '').trim();
      if (uname && facId) usernameByFacility[facId] = uname;
    });
  }
  facValues.forEach(function(row) {
    const facId = String(row[0] || '').trim();
    const status = String(row[5] || '').trim();
    if (!facId || status === 'Nonaktif') return;
    const uname = usernameByFacility[facId];
    if (!uname) return;
    try {
      processDeferredQueueForFacility_(uname);
    } catch (e) {
      try {
        catatLogSistem_('Queue Deferred', 'Gagal proses antrian facility ' + facId + ': ' + e.message);
      } catch (e2) {}
    }
  });
}

function installDeferredQueueTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processDeferredQueueAllFacilities_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processDeferredQueueAllFacilities_').timeBased().everyMinutes(1).create();
  Logger.log('Trigger terpasang: processDeferredQueueAllFacilities_ tiap 1 menit.');
}

function ringkasanCacheFacilityPrefix_(username) {
  if (!username) return 'legacy';
  const facInfo = getUserFacility(username);
  return facInfo ? facInfo.id : 'legacy';
}

function ringkasanHarianRowCachePut_(tanggal, rowIndex, username) {
  try {
    const prefix = ringkasanCacheFacilityPrefix_(username);
    CacheService.getScriptCache().put('ringkasanRow_' + prefix + '_' + tanggal, String(rowIndex), 21600);
  } catch (e) {}
}

function ringkasanHarianRowCacheGet_(tanggal, username) {
  try {
    const prefix = ringkasanCacheFacilityPrefix_(username);
    const v = CacheService.getScriptCache().get('ringkasanRow_' + prefix + '_' + tanggal);
    return v ? Number(v) : -1;
  } catch (e) {
    return -1;
  }
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

function getOrCreateRingkasanHarianRow_(sheet, tanggal, username) {
  let rowIndex = findRingkasanHarianRow_(sheet, tanggal, username);
  if (rowIndex === -1) {
    rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 1, 1, RINGKASAN_HARIAN_HEADERS.length).setValues([ [ tanggal, 0, 0, 0, '{}', '{}', '{}', '{}' ] ]);
    ringkasanHarianRowCachePut_(tanggal, rowIndex, username);
  } else {
    ringkasanHarianRowCachePut_(tanggal, rowIndex, username);
  }
  const values = sheet.getRange(rowIndex, 1, 1, RINGKASAN_HARIAN_HEADERS.length).getValues()[0];
  return {
    rowIndex: rowIndex,
    values: values
  };
}

function computeRowContribution_(row) {
  const namaPetugas = row[7];
  const hasilAwal = row[10], statusValidasi = row[11], namaValidator = row[12], hasilFinal = row[14];
  const kategoriSelisih = String(row[17] || '').trim();
  const statusTask = row[18];
  const effectiveHasil = statusValidasi === 'Pending' ? 'PENDING' : statusValidasi === 'Tidak Perlu' ? hasilAwal : hasilFinal;
  const kpi = {
    total: 1,
    hit: effectiveHasil === 'HIT' ? 1 : 0,
    disc: effectiveHasil === 'DISCREPANCY' ? 1 : 0
  };
  let petugas = null;
  if (statusValidasi === 'Tidak Perlu') {
    petugas = {
      nama: namaPetugas,
      total: 1,
      hit: 1
    };
  } else if (statusValidasi === 'Selesai') {
    petugas = {
      nama: namaPetugas,
      total: 1,
      hit: hasilAwal === hasilFinal ? 1 : 0
    };
  }
  const petugasRaw = namaPetugas ? {
    nama: namaPetugas,
    rawTotal: 1,
    rawSelesai: statusValidasi !== 'Pending' ? 1 : 0
  } : null;
  let validator = null;
  if (namaValidator && statusTask === 'Selesai') {
    const validatorSalah = hasilFinal === 'DISCREPANCY' && kategoriSelisih === 'Salah Hitung';
    validator = {
      nama: namaValidator,
      total: 1,
      hit: validatorSalah ? 0 : 1
    };
  }
  return {
    kpi: kpi,
    petugas: petugas,
    petugasRaw: petugasRaw,
    validator: validator,
    kategori: kategoriSelisih || null
  };
}

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
    if (!map[k]) map[k] = {
      total: 0,
      hit: 0
    };
    map[k].total += newEntry.total;
    map[k].hit += newEntry.hit;
  }
}

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
    if (!map[k]) map[k] = {
      rawTotal: 0,
      rawSelesai: 0
    };
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

function applyRingkasanDelta_(tanggal, oldRow, newRow, username) {
  if (!tanggal) return;
  const zero = {
    kpi: {
      total: 0,
      hit: 0,
      disc: 0
    },
    petugas: null,
    petugasRaw: null,
    validator: null,
    kategori: null
  };
  const oldC = oldRow ? computeRowContribution_(oldRow) : zero;
  const newC = newRow ? computeRowContribution_(newRow) : zero;
  const kpiSame = oldC.kpi.total === newC.kpi.total && oldC.kpi.hit === newC.kpi.hit && oldC.kpi.disc === newC.kpi.disc;
  const petugasSame = JSON.stringify(oldC.petugas) === JSON.stringify(newC.petugas);
  const petugasRawSame = JSON.stringify(oldC.petugasRaw) === JSON.stringify(newC.petugasRaw);
  const validatorSame = JSON.stringify(oldC.validator) === JSON.stringify(newC.validator);
  const kategoriSame = oldC.kategori === newC.kategori;
  if (kpiSame && petugasSame && petugasRawSame && validatorSame && kategoriSame) return;
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
  sheet.getRange(rowInfo.rowIndex, 2, 1, 7).setValues([ [ kpiTotal, kpiHit, kpiDisc, JSON.stringify(petugasMap), JSON.stringify(validatorMap), JSON.stringify(kategoriMap), JSON.stringify(petugasRawMap) ] ]);
}

function rebuildRingkasanHarian_(username) {
  const facInfoForLock = getUserFacility(username);
  const facLock = acquireFacilityLock_(facInfoForLock ? facInfoForLock.id : null, 2e4);
  if (!facLock) throw new Error('Sistem sedang sibuk, coba lagi nanti.');
  try {
    const riwayatSheet = getRiwayatSheet_(username);
    const lastRow = riwayatSheet.getLastRow();
    const acc = {};
    if (lastRow >= 2) {
      const data = riwayatSheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
      data.forEach(function(row) {
        const tgl = rowDateTag_(row);
        if (!tgl) return;
        if (!acc[tgl]) acc[tgl] = {
          kpiTotal: 0,
          kpiHit: 0,
          kpiDisc: 0,
          petugas: {},
          validator: {},
          kategori: {},
          petugasRaw: {}
        };
        const c = computeRowContribution_(row);
        const bucket = acc[tgl];
        bucket.kpiTotal += c.kpi.total;
        bucket.kpiHit += c.kpi.hit;
        bucket.kpiDisc += c.kpi.disc;
        if (c.petugas) {
          if (!bucket.petugas[c.petugas.nama]) bucket.petugas[c.petugas.nama] = {
            total: 0,
            hit: 0
          };
          bucket.petugas[c.petugas.nama].total += c.petugas.total;
          bucket.petugas[c.petugas.nama].hit += c.petugas.hit;
        }
        if (c.petugasRaw) {
          if (!bucket.petugasRaw[c.petugasRaw.nama]) bucket.petugasRaw[c.petugasRaw.nama] = {
            rawTotal: 0,
            rawSelesai: 0
          };
          bucket.petugasRaw[c.petugasRaw.nama].rawTotal += c.petugasRaw.rawTotal;
          bucket.petugasRaw[c.petugasRaw.nama].rawSelesai += c.petugasRaw.rawSelesai;
        }
        if (c.validator) {
          if (!bucket.validator[c.validator.nama]) bucket.validator[c.validator.nama] = {
            total: 0,
            hit: 0
          };
          bucket.validator[c.validator.nama].total += c.validator.total;
          bucket.validator[c.validator.nama].hit += c.validator.hit;
        }
        if (c.kategori) {
          bucket.kategori[c.kategori] = (bucket.kategori[c.kategori] || 0) + 1;
        }
      });
    }
    const sheet = getRingkasanHarianSheet_(username);
    const outRows = Object.keys(acc).sort().map(function(tgl) {
      const a = acc[tgl];
      return [ tgl, a.kpiTotal, a.kpiHit, a.kpiDisc, JSON.stringify(a.petugas), JSON.stringify(a.validator), JSON.stringify(a.kategori), JSON.stringify(a.petugasRaw) ];
    });
    const oldLastRow = sheet.getLastRow();
    if (oldLastRow > 1) sheet.getRange(2, 1, oldLastRow - 1, RINGKASAN_HARIAN_HEADERS.length).clearContent();
    if (outRows.length > 0) {
      sheet.getRange(2, 1, outRows.length, RINGKASAN_HARIAN_HEADERS.length).setValues(outRows);
    }
    return {
      success: true,
      hari: outRows.length
    };
  } finally {
    facLock.release();
  }
}

function rebuildRingkasanHarianMenu_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('Ini akan menghitung ULANG seluruh sheet "' + RINGKASAN_HARIAN_SHEET_NAME + '" dari isi Riwayat saat ini (dipakai kalau angka Dashboard terasa tidak sinkron). Proses ini TIDAK mengubah data Riwayat sama sekali.\n\nLanjutkan?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  const hasil = rebuildRingkasanHarian_(null);
  ui.alert('Selesai. Ringkasan_Harian dibangun ulang untuk ' + hasil.hari + ' tanggal.');
}

const SETUP_FACILITY_NAMA_ = 'ISI NAMA FACILITY';

const SETUP_FACILITY_KODE_ = 'ISI-KODE';

const SETUP_ADMIN_NIK_ = 'ISI.NIK.ADMIN';

function setupEnsureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  let dibuat = false;
  if (!sheet) {
    sheet = ss.insertSheet(name);
    dibuat = true;
  }
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([ headers ]);
  }
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return {
    sheet: sheet,
    dibuat: dibuat
  };
}

function terapkanFormatKolomTeks_(ss) {
  const kolom = [ [ SHEET_NAME, [ 'B2:B', 'E2:E' ] ], [ RIWAYAT_SHEET_NAME, [ 'B2:B', 'D2:D' ] ], [ RIWAYAT_ARCHIVE_SHEET_NAME, [ 'B2:B', 'D2:D' ] ], [ RINGKASAN_HARIAN_SHEET_NAME, [ 'A2:A' ] ], [ QUEUE_RINGKASAN_SHEET_NAME, [ 'B2:B' ] ], [ LOG_AKSES_SHEET_NAME, [ 'A2:A', 'B2:B' ] ], [ LOKASI_AKTIF_SHEET_NAME, [ 'A2:A' ] ], [ MASTER_USER_SHEET_NAME, [ 'A2:A' ] ] ];
  kolom.forEach(function(k) {
    const sheet = ss.getSheetByName(k[0]);
    if (!sheet) return;
    k[1].forEach(function(a1) {
      sheet.getRange(a1).setNumberFormat('@');
    });
  });
}

function setupCariBaris_(sheet, kolom, nilai) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const target = String(nilai).trim().toLowerCase();
  const values = sheet.getRange(2, kolom, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toLowerCase() === target) return i + 2;
  }
  return 0;
}

function setupAwal() {
  if (/^ISI[ .-]/.test(SETUP_ADMIN_NIK_) || /^ISI[ .-]/.test(SETUP_FACILITY_NAMA_) || /^ISI[ .-]/.test(SETUP_FACILITY_KODE_)) {
    throw new Error('Isi dulu SETUP_FACILITY_NAMA_, SETUP_FACILITY_KODE_, dan SETUP_ADMIN_NIK_ di bagian Setup, lalu jalankan lagi.');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hasil = [];
  const now = new Date;
  let jumlahBaru = 0;
  const master = setupEnsureSheet_(ss, MASTER_SHEET_NAME, [ 'ACTIVE LOCATION', '', 'USER', 'ROLE', 'STATUS', 'LEVEL GROUP', 'PRODUKTIVITAS (qty/orang)', '', 'KATEGORI SELISIH', 'USER PIC', 'ROLE PIC' ]);
  const m = master.sheet;
  if (master.dibuat) jumlahBaru++;
  if (!String(m.getRange('F2').getValue() || '').trim()) {
    const rows = LEVEL_GROUPS.map(function(g) {
      return [ 'Level ' + g.min + '-' + g.max, g.defaultProd ];
    });
    rows.push([ 'Reach Truck Ready (unit)', 0 ]);
    rows.push([ 'Tangga Pesawat Ready (unit)', 0 ]);
    m.getRange(2, 6, rows.length, 2).setValues(rows);
  }
  if (!String(m.getRange('I2').getValue() || '').trim()) {
    m.getRange(2, 9, TASK_KATEGORI_LIST.length, 1).setValues(TASK_KATEGORI_LIST.map(function(k) {
      return [ k ];
    }));
  }
  if (!setupCariBaris_(m, 3, SETUP_ADMIN_NIK_)) {
    const lastUserRow = Math.max(1, m.getRange('C:C').getValues().filter(function(r) {
      return String(r[0]).trim();
    }).length);
    m.getRange(lastUserRow + 1, 3, 1, 3).setValues([ [ SETUP_ADMIN_NIK_, 'admin', 'Aktif' ] ]);
  }
  const pusat = [ [ MASTER_USER_SHEET_NAME, MASTER_USER_HEADERS ], [ FACILITY_SHEET_NAME, FACILITY_HEADERS ], [ USER_FACILITY_ASSIGNMENT_SHEET_NAME, USER_FACILITY_ASSIGNMENT_HEADERS ], [ CONFIG_AKSES_SETTING_SHEET_NAME, CONFIG_AKSES_SETTING_HEADERS ], [ LOG_PERUBAHAN_CONFIG_SHEET_NAME, LOG_PERUBAHAN_CONFIG_HEADERS ], [ LOG_SISTEM_SHEET_NAME, LOG_SISTEM_HEADERS ], [ SUBMIT_QUEUE_SHEET_NAME, SUBMIT_QUEUE_HEADERS ], [ VALIDASI_QUEUE_SHEET_NAME, VALIDASI_QUEUE_HEADERS ], [ TASK_QUEUE_SHEET_NAME, TASK_QUEUE_HEADERS ] ];
  const semua = pusat.concat(FACILITY_OPERATIONAL_SHEETS.map(function(d) {
    return [ d.name, d.headers ];
  }));
  const sheets = {};
  semua.forEach(function(d) {
    const r = setupEnsureSheet_(ss, d[0], d[1]);
    sheets[d[0]] = r.sheet;
    if (r.dibuat) jumlahBaru++;
  });
  terapkanFormatKolomTeks_(ss);
  hasil.push(jumlahBaru + ' sheet baru dibuat (total ' + ss.getSheets().length + ' sheet).');
  const facSheet = sheets[FACILITY_SHEET_NAME];
  let facRow = setupCariBaris_(facSheet, 5, ss.getId());
  let facId;
  if (facRow) {
    facId = String(facSheet.getRange(facRow, 1).getValue()).trim();
    hasil.push('Facility sudah ada: ' + facSheet.getRange(facRow, 2).getValue() + ' (' + facId + ').');
  } else {
    facId = 'FAC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    facSheet.appendRow([ facId, SETUP_FACILITY_NAMA_, SETUP_FACILITY_KODE_, ss.getName(), ss.getId(), 'Aktif', now, SETUP_ADMIN_NIK_, ss.getUrl() ]);
    hasil.push('Facility dibuat: ' + SETUP_FACILITY_NAMA_ + ' (' + facId + ').');
  }
  const userSheet = sheets[MASTER_USER_SHEET_NAME];
  if (!setupCariBaris_(userSheet, 1, SETUP_ADMIN_NIK_)) {
    userSheet.appendRow([ SETUP_ADMIN_NIK_, 'admin', 'Aktif', facId ]);
    hasil.push('User admin dibuat: ' + SETUP_ADMIN_NIK_ + '.');
  } else {
    hasil.push('User ' + SETUP_ADMIN_NIK_ + ' sudah ada.');
  }
  const assignSheet = sheets[USER_FACILITY_ASSIGNMENT_SHEET_NAME];
  if (!setupCariBaris_(assignSheet, 1, SETUP_ADMIN_NIK_)) {
    assignSheet.appendRow([ SETUP_ADMIN_NIK_, facId, now, SETUP_ADMIN_NIK_ ]);
  }
  const aksesSheet = sheets[CONFIG_AKSES_SETTING_SHEET_NAME];
  if (!setupCariBaris_(aksesSheet, 1, SETUP_ADMIN_NIK_)) {
    aksesSheet.appendRow([ SETUP_ADMIN_NIK_, 'Admin pertama', Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd') ]);
  }
  getConfigSistemSheet_();
  getLevelAssignmentSheet_();
  [ 'Sheet1', 'Lembar1' ].forEach(function(nama) {
    const kosong = ss.getSheetByName(nama);
    if (kosong && kosong.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(kosong);
  });
  hasil.push(pasangTriggerWorker());
  try {
    clearMasterCache_();
  } catch (e) {}
  hasil.push('Lokasi_Aktif masih kosong: isi kolom A sheet "Lokasi_Aktif" (satu lokasi per baris) atau impor lewat menu Config di aplikasi.');
  hasil.push('Langkah berikutnya: Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone), lalu kirim URL /exec ke pengembang aplikasi untuk ditanam.');
  const pesan = hasil.join('\n');
  Logger.log(pesan);
  try {
    SpreadsheetApp.getUi().alert('Setup Awal selesai:\n\n' + pesan);
  } catch (e) {}
  return pesan;
}

function workerSemua() {
  resetMemEksekusi_();
  const langkah = [ [ 'Submit', processSubmitQueue_ ], [ 'Validasi', processValidasiQueue_ ], [ 'Task', processTaskQueue_ ], [ 'Ringkasan', processDeferredQueueAllFacilities_ ] ];
  langkah.forEach(function(l) {
    try {
      l[1]();
    } catch (e) {
      catatLogSistem_('Worker ' + l[0], 'ERROR: ' + (e && e.message || e));
    }
  });
}

function pasangTriggerWorker() {
  const lama = [ 'workerSemua', 'processSubmitQueue_', 'processValidasiQueue_', 'processTaskQueue_', 'processDeferredQueueAllFacilities_' ];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (lama.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('workerSemua').timeBased().everyMinutes(1).create();
  return 'Trigger terpasang: workerSemua tiap 1 menit.';
}

function cekArchiveRiwayatLama() {
  return cekArchiveRiwayatLama_();
}

function rebuildRingkasanHarianMenu() {
  return rebuildRingkasanHarianMenu_();
}

function rebuildAntrianAktifMenu() {
  return rebuildAntrianAktifMenu_();
}

function getSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, HEADERS);
  }
  return sheet;
}

function getRiwayatSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RIWAYAT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RIWAYAT_SHEET_NAME);
    sheet.appendRow(RIWAYAT_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, RIWAYAT_HEADERS);
  }
  return sheet;
}

function getRiwayatArchiveSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RIWAYAT_ARCHIVE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RIWAYAT_ARCHIVE_SHEET_NAME);
    sheet.appendRow(RIWAYAT_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, RIWAYAT_HEADERS);
  }
  return sheet;
}

var _ensuredHeadersCache_ = {};

function ensureHeaderColumns_(sheet, expectedHeaders) {
  try {
    const sheetId = sheet.getSheetId();
    if (_ensuredHeadersCache_[sheetId]) return;
    const lastCol = sheet.getLastColumn();
    if (lastCol < expectedHeaders.length) {
      const missingHeaders = expectedHeaders.slice(lastCol);
      sheet.getRange(1, lastCol + 1, 1, missingHeaders.length).setValues([ missingHeaders ]);
    }
    _ensuredHeadersCache_[sheetId] = true;
  } catch (e) {}
}

function getLogAksesSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_AKSES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_AKSES_SHEET_NAME);
    sheet.appendRow(LOG_AKSES_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, LOG_AKSES_HEADERS);
  }
  return sheet;
}

function getLogAnomaliSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_ANOMALI_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_ANOMALI_SHEET_NAME);
    sheet.appendRow(LOG_ANOMALI_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, LOG_ANOMALI_HEADERS);
  }
  return sheet;
}

function getLogUnassignedSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_UNASSIGNED_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_UNASSIGNED_SHEET_NAME);
    sheet.appendRow(LOG_UNASSIGNED_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, LOG_UNASSIGNED_HEADERS);
  }
  return sheet;
}

function getLogBuktiSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_BUKTI_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_BUKTI_SHEET_NAME);
    sheet.appendRow(LOG_BUKTI_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, LOG_BUKTI_HEADERS);
  }
  return sheet;
}

function getRingkasanHarianSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RINGKASAN_HARIAN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RINGKASAN_HARIAN_SHEET_NAME);
    sheet.appendRow(RINGKASAN_HARIAN_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, RINGKASAN_HARIAN_HEADERS);
  }
  return sheet;
}

function getAntrianAktifSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ANTRIAN_AKTIF_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ANTRIAN_AKTIF_SHEET_NAME);
    sheet.appendRow(ANTRIAN_AKTIF_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, ANTRIAN_AKTIF_HEADERS);
  }
  return sheet;
}

function getLokasiAktifSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOKASI_AKTIF_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOKASI_AKTIF_SHEET_NAME);
    sheet.appendRow(LOKASI_AKTIF_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, LOKASI_AKTIF_HEADERS);
  }
  return sheet;
}

function getQueueRingkasanSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(QUEUE_RINGKASAN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(QUEUE_RINGKASAN_SHEET_NAME);
    sheet.appendRow(QUEUE_RINGKASAN_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, QUEUE_RINGKASAN_HEADERS);
  }
  return sheet;
}

function getQueueCounterSheet_(username) {
  const ss = username ? getOperasionalSpreadsheet_(username) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(QUEUE_COUNTER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(QUEUE_COUNTER_SHEET_NAME);
    sheet.appendRow(QUEUE_COUNTER_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, QUEUE_COUNTER_HEADERS);
  }
  return sheet;
}

function getSubmitQueueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SUBMIT_QUEUE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SUBMIT_QUEUE_SHEET_NAME);
    sheet.appendRow(SUBMIT_QUEUE_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, SUBMIT_QUEUE_HEADERS);
  }
  return sheet;
}

function getValidasiQueueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(VALIDASI_QUEUE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(VALIDASI_QUEUE_SHEET_NAME);
    sheet.appendRow(VALIDASI_QUEUE_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, VALIDASI_QUEUE_HEADERS);
  }
  return sheet;
}

function getTaskQueueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TASK_QUEUE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TASK_QUEUE_SHEET_NAME);
    sheet.appendRow(TASK_QUEUE_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaderColumns_(sheet, TASK_QUEUE_HEADERS);
  }
  return sheet;
}

function getMasterSheet_() {
  return getMasterSheetWithFallback_();
}

function fixHeaderForSheet_(sheet, headers) {
  const currentHeaderRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const isMatch = currentHeaderRow.every(function(val, idx) {
    return String(val).trim() === headers[idx];
  });
  if (isMatch) {
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return;
  }
  const firstCell = String(sheet.getRange(1, 1).getValue()).trim();
  if (headers.indexOf(firstCell) === -1 && firstCell !== '') {
    sheet.insertRowBefore(1);
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([ headers ]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function riwayatRowCachePut_(id, rowIndex) {
  try {
    CacheService.getScriptCache().put('riw_' + id, String(rowIndex), 21600);
  } catch (e) {}
}

function riwayatRowCacheGet_(id) {
  try {
    const v = CacheService.getScriptCache().get('riw_' + id);
    return v ? Number(v) : -1;
  } catch (e) {
    return -1;
  }
}

function findRiwayatRow_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const cachedRow = riwayatRowCacheGet_(id);
  if (cachedRow >= 2 && cachedRow <= lastRow) {
    if (String(sheet.getRange(cachedRow, 1).getValue()) === String(id)) {
      riwayatRowCachePut_(id, cachedRow);
      return cachedRow;
    }
  }
  const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]) === String(id)) {
      const rowIndex = i + 2;
      riwayatRowCachePut_(id, rowIndex);
      return rowIndex;
    }
  }
  return -1;
}

function queueSubmitPayload_(no, namaPetugas, qtyCount, facilityId, waktuHitung) {
  const sheet = getSubmitQueueSheet_();
  appendRowLocked_(sheet, [ new Date, no, namaPetugas, qtyCount, facilityId, waktuHitung || '' ]);
}

function processSubmitQueue_() {
  const qSheet = getSubmitQueueSheet_();
  const lastRow = qSheet.getLastRow();
  if (lastRow < 2) return;
  const values = qSheet.getRange(2, 1, lastRow - 1, SUBMIT_QUEUE_HEADERS.length).getValues();
  const byFacility = {};
  values.forEach(function(row, idx) {
    const facilityId = String(row[4] || '').trim();
    if (!byFacility[facilityId]) byFacility[facilityId] = [];
    byFacility[facilityId].push(idx);
  });
  const processedRowNumbers = [];
  const cache = CacheService.getScriptCache();
  Object.keys(byFacility).forEach(function(facilityId) {
    const indices = byFacility[facilityId];
    if (!facilityId) {
      indices.forEach(function(i) {
        catatLogSistem_('Worker Submit', 'Baris antrian tanpa FacilityId dibuang: ' + JSON.stringify(values[i]));
        processedRowNumbers.push(i + 2);
      });
      return;
    }
    const workerGuardKey = 'submitQueueWorkerLock_' + facilityId;
    if (cache.get(workerGuardKey)) {
      return;
    }
    try {
      cache.put(workerGuardKey, '1', 55);
    } catch (e) {}
    const facLock = acquireFacilityLock_(facilityId, 1e4);
    if (!facLock) {
      try {
        cache.remove(workerGuardKey);
      } catch (e) {}
      return;
    }
    try {
      const sortedIndices = indices.slice().sort(function(a, b) {
        return new Date(values[a][0]) - new Date(values[b][0]);
      });
      sortedIndices.forEach(function(i) {
        const row = values[i];
        const no = row[1];
        const namaPetugas = String(row[2] || '').trim();
        const qtyCount = row[3];
        try {
          const result = processSingleQueuedSubmit_(no, namaPetugas, qtyCount, facilityId, waktuBarisAntrean_(row[5], row[0]));
          if (!result.success) {
            catatLogSistem_('Worker Submit', 'Skip No ' + no + ' (' + namaPetugas + '): ' + result.message);
          }
        } catch (e) {
          catatLogSistem_('Worker Submit', 'ERROR No ' + no + ' (' + namaPetugas + '): ' + e.message);
        }
        processedRowNumbers.push(i + 2);
      });
    } finally {
      facLock.release();
      try {
        cache.remove(workerGuardKey);
      } catch (e) {}
    }
  });
  deleteProcessedQueueRows_(qSheet, processedRowNumbers);
}

function processSingleQueuedSubmit_(no, namaPetugas, qtyCount, facilityId, waktuHitung) {
  const aktual = Number(qtyCount);
  const sheet = getSheet_(namaPetugas);
  const itemData = findDataCountRowAndValues_(sheet, no);
  if (!itemData) {
    return {
      success: false,
      message: 'Item tidak ditemukan (kemungkinan dihapus/direvisi Admin).'
    };
  }
  const rowIndex = itemData.rowIndex;
  const rowValues = itemData.rowValues;
  if (rowValues[10] !== 'Pending') {
    return {
      success: false,
      message: 'Sudah diproses sebelumnya (duplikat submit).'
    };
  }
  const lokasi = rowValues[2];
  const article = rowValues[4];
  const description = rowValues[5];
  const qtyTransaksi = rowValues[6];
  const qtySystem = Number(rowValues[7]);
  const addWhoTransaksi = rowValues[15];
  const selisih = aktual - qtySystem;
  const hasilAwal = selisih === 0 ? 'HIT' : 'DISCREPANCY';
  const waktu = waktuHitung || new Date;
  sheet.getRange(rowIndex, 10, 1, 6).setValues([ [ namaPetugas, 'Selesai', aktual, selisih, hasilAwal, waktu ] ]);
  const statusValidasi = hasilAwal === 'HIT' ? 'Tidak Perlu' : 'Pending';
  const hasilFinal = hasilAwal === 'HIT' ? 'HIT' : '';
  const assignedValidator = hasilAwal === 'HIT' ? '' : assignNextValidator_(facilityId);
  appendRiwayatCycle_(lokasi, article, description, qtyTransaksi, qtySystem, namaPetugas, aktual, selisih, hasilAwal, statusValidasi, hasilFinal, waktu, assignedValidator, addWhoTransaksi);
  return {
    success: true,
    hasil: hasilAwal,
    selisih: selisih
  };
}

function installSubmitQueueTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processSubmitQueue_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processSubmitQueue_').timeBased().everyMinutes(1).create();
  Logger.log('Trigger terpasang: processSubmitQueue_ tiap 1 menit.');
}

function canManageTasks_(_username) {
  const info = getUserRole(_username);
  return !!info && (info.role === 'admin' || info.role === 'inventory' || info.role === 'developer');
}

function getOpenTasksData_(username) {
  const sheet = getRiwayatSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const activeStatuses = TASK_STATUS_FLOW.slice(0, TASK_STATUS_FLOW.length - 1);
  const now = new Date;
  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  const tasks = [];
  data.forEach(function(row) {
    const statusTask = row[18];
    if (activeStatuses.indexOf(statusTask) === -1) return;
    const waktuCycle = row[15];
    const umurHari = Object.prototype.toString.call(waktuCycle) === '[object Date]' ? Math.floor((now - waktuCycle) / 864e5) : 0;
    const selisihAbs = Math.abs(Number(row[9]) || 0);
    const skorPrioritas = umurHari * PRIORITAS_BOBOT_UMUR + selisihAbs;
    tasks.push({
      id: row[0],
      tanggal: row[1],
      lokasi: row[2],
      article: row[3],
      description: row[4],
      selisih: row[9],
      namaPetugas: row[7],
      kategori: row[17],
      alasan: row[17],
      statusTask: statusTask,
      picInvestigasi: row[23],
      addWhoTransaksi: row[25],
      namaValidator: row[12],
      nameValidator: row[12],
      umurHari: umurHari,
      isCritical: umurHari > TASK_SLA_HARI,
      skorPrioritas: skorPrioritas
    });
  });
  tasks.sort(function(a, b) {
    if (a.isCritical !== b.isCritical) return a.isCritical ? -1 : 1;
    return b.skorPrioritas - a.skorPrioritas;
  });
  return tasks;
}

function getOpenTasks(requesterUsername) {
  requireRole_(requesterUsername, [ 'inventory', 'admin', 'developer' ]);
  return getOpenTasksData_(requesterUsername);
}

function getOpenTaskCount(requesterUsername) {
  requireRole_(requesterUsername, [ 'inventory', 'admin', 'developer' ]);
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const activeStatuses = TASK_STATUS_FLOW.slice(0, TASK_STATUS_FLOW.length - 1);
  const col = sheet.getRange(2, 19, lastRow - 1, 1).getValues();
  let count = 0;
  col.forEach(function(r) {
    if (activeStatuses.indexOf(r[0]) !== -1) count++;
  });
  return count;
}

function validateBuktiTransaksi_(buktiRows, article, selisihAbs, kategori, lokasiTask) {
  const isAdjustment = KATEGORI_ADJUSTMENT_.indexOf(String(kategori || '').trim()) !== -1;
  const jenisLabel = isAdjustment ? 'Adjustment' : 'Move/Picking';
  if (!buktiRows || !buktiRows.length) {
    return {
      valid: false,
      message: 'Bukti transaksi WMS (' + jenisLabel + ') wajib di-paste untuk menutup task ini (kecuali kategori Salah Hitung / Barang Sudah di Picking).'
    };
  }
  const articleNorm = String(article || '').trim().toUpperCase();
  if (isAdjustment) {
    const relevant = buktiRows.filter(function(r) {
      const sku = String(r[3] || '').trim().toUpperCase();
      const tipe = String(r[2] || '').trim().toUpperCase();
      return sku === articleNorm && tipe === 'ADJUSTMENT';
    });
    if (!relevant.length) {
      return {
        valid: false,
        message: 'Tidak ada baris bukti dengan SKU ' + article + ' dan jenis transaksi ADJUSTMENT. Kategori ini wajib bukti transaksi ADJUSTMENT dari WMS.'
      };
    }
    const totalQtySigned = relevant.reduce(function(s, r) {
      return s + (Number(r[12]) || 0);
    }, 0);
    const expectedSigned = String(kategori).trim() === 'Adjustment Plus' ? selisihAbs : -selisihAbs;
    if (Math.round(totalQtySigned) !== Math.round(expectedSigned)) {
      return {
        valid: false,
        message: 'Total Qty adjustment (' + totalQtySigned + ') harus tepat ' + expectedSigned + ' sesuai kategori ' + kategori + ' & selisih task ini. Qty di data WMS harus sesuai tanda (minus utk Adjustment Minus, plus utk Adjustment Plus).'
      };
    }
    const lokasiNorm = String(lokasiTask || '').trim().toUpperCase();
    const hasLokasiMatch = relevant.some(function(r) {
      return String(r[9] || '').trim().toUpperCase() === lokasiNorm;
    });
    if (!hasLokasiMatch) {
      return {
        valid: false,
        message: 'Bukti adjustment harus mencantumkan TOLOC yang sama dengan lokasi task ini (' + lokasiTask + ').'
      };
    }
    return {
      valid: true
    };
  }
  const relevant = buktiRows.filter(function(r) {
    const sku = String(r[3] || '').trim().toUpperCase();
    const tipe = String(r[2] || '').trim().toUpperCase();
    return sku === articleNorm && (tipe === 'MOVE' || tipe === 'PICKING');
  });
  if (!relevant.length) {
    return {
      valid: false,
      message: 'Tidak ada baris bukti dengan SKU ' + article + ' dan jenis transaksi MOVE/PICKING. Cek lagi data yang di-paste.'
    };
  }
  const totalQty = relevant.reduce(function(s, r) {
    return s + Math.abs(Number(r[12]) || 0);
  }, 0);
  if (Math.round(totalQty) !== Math.round(selisihAbs)) {
    return {
      valid: false,
      message: 'Total Qty di bukti transaksi (' + totalQty + ') tidak sama dengan besar selisih task ini (' + selisihAbs + '). Qty harus sama persis.'
    };
  }
  const hasLokasi = relevant.some(function(r) {
    return String(r[7] || '').trim() || String(r[9] || '').trim();
  });
  if (!hasLokasi) {
    return {
      valid: false,
      message: 'Bukti transaksi harus mencantumkan nomor lokasi (FROMLOC/TOLOC).'
    };
  }
  return {
    valid: true
  };
}

function updateTaskStatus(id, namaUser, newStatus, catatan, kategori, picUsername, buktiRows, lokasiTask, articleTask, selisihTask) {
  const catatanTrim = String(catatan || '').trim();
  const newStatusTrim = String(newStatus || '').trim();
  const kategoriTrim = String(kategori || '').trim();
  const picTrim = String(picUsername || '').trim();
  if (!canManageTasks_(namaUser)) {
    return {
      success: false,
      message: 'Tidak punya akses update task (khusus role Inventory/Admin).'
    };
  }
  const targetIdx = TASK_STATUS_FLOW.indexOf(newStatusTrim);
  if (targetIdx <= 0) {
    return {
      success: false,
      message: 'Status tujuan tidak valid.'
    };
  }
  if (!catatanTrim) return {
    success: false,
    message: 'Catatan progres wajib diisi.'
  };
  const isFinal = newStatusTrim === 'Selesai';
  if (isFinal) {
    if (TASK_KATEGORI_LIST.indexOf(kategoriTrim) === -1) {
      return {
        success: false,
        message: 'Kategori selisih wajib dipilih (Lebih/Kurang Picking, Lebih/Kurang Move, Adjustment Plus/Minus, Salah Hitung, atau Barang Sudah di Picking).'
      };
    }
    if (!picTrim) return {
      success: false,
      message: 'User PIC wajib dipilih.'
    };
    if (!getUserRole(picTrim)) {
      return {
        success: false,
        message: 'User PIC "' + picTrim + '" tidak dikenali di Master.'
      };
    }
    if (KATEGORI_TANPA_BUKTI_.indexOf(kategoriTrim) === -1) {
      const selisihAbs = Math.abs(Number(selisihTask) || 0);
      const cek = validateBuktiTransaksi_(buktiRows, articleTask, selisihAbs, kategoriTrim, lokasiTask);
      if (!cek.valid) return {
        success: false,
        message: cek.message
      };
    }
  }
  const facInfo = requireUserFacility_(namaUser);
  if (!facInfo) {
    return {
      success: false,
      message: PESAN_TANPA_FACILITY_
    };
  }
  try {
    queueTaskUpdatePayload_(id, namaUser, newStatusTrim, catatanTrim, kategoriTrim, picTrim, buktiRows, facInfo.id);
  } catch (e) {
    return {
      success: false,
      message: 'Gagal masuk antrian: ' + e.message
    };
  }
  return {
    success: true,
    queued: true,
    message: 'Update diterima.'
  };
}

function getPlusMinusCandidates(requesterUsername) {
  requireRole_(requesterUsername, [ 'inventory', 'admin', 'developer' ]);
  const all = getOpenTasksData_(requesterUsername);
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  const plus = [], minus = [];
  if (lastRow < 2) return {
    plus: plus,
    minus: minus
  };
  const hasilFinalById = {};
  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  data.forEach(function(row) {
    hasilFinalById[row[0]] = row[14];
  });
  all.forEach(function(t) {
    if (hasilFinalById[t.id] !== 'DISCREPANCY') return;
    const s = Number(t.selisih) || 0;
    if (s > 0) plus.push(t); else if (s < 0) minus.push(t);
  });
  return {
    plus: plus,
    minus: minus
  };
}

function closePlusMinusPair(minusTaskId, plusTaskId, namaUser, catatan, picUsername, buktiRows) {
  const catatanTrim = String(catatan || '').trim();
  const picTrim = String(picUsername || '').trim();
  if (!canManageTasks_(namaUser)) {
    return {
      success: false,
      message: 'Tidak punya akses (khusus role Inventory/Admin).'
    };
  }
  if (!minusTaskId || !plusTaskId) {
    return {
      success: false,
      message: 'Task Minus & task Plus wajib dipilih.'
    };
  }
  if (String(minusTaskId) === String(plusTaskId)) {
    return {
      success: false,
      message: 'Task Minus & Plus tidak boleh task yang sama.'
    };
  }
  if (!catatanTrim) return {
    success: false,
    message: 'Catatan penyelesaian wajib diisi.'
  };
  if (!picTrim) return {
    success: false,
    message: 'User PIC wajib dipilih.'
  };
  if (!getUserRole(picTrim)) {
    return {
      success: false,
      message: 'User PIC "' + picTrim + '" tidak dikenali di Master.'
    };
  }
  if (!buktiRows || !buktiRows.length) {
    return {
      success: false,
      message: 'Bukti transaksi MOVE dari WMS wajib di-paste untuk menutup pasangan task ini.'
    };
  }
  const facInfo = requireUserFacility_(namaUser);
  if (!facInfo) {
    return {
      success: false,
      message: PESAN_TANPA_FACILITY_
    };
  }
  const facLock = acquireFacilityLock_(facInfo.id);
  if (!facLock) {
    return {
      success: false,
      message: 'Sistem sedang sibuk, coba lagi.'
    };
  }
  try {
    const sheet = getRiwayatSheet_(namaUser);
    const rowMinusIdx = findRiwayatRow_(sheet, minusTaskId);
    const rowPlusIdx = findRiwayatRow_(sheet, plusTaskId);
    if (rowMinusIdx === -1 || rowPlusIdx === -1) {
      return {
        success: false,
        message: 'Salah satu task tidak ditemukan, kemungkinan sudah diproses.'
      };
    }
    const rowMinus = sheet.getRange(rowMinusIdx, 1, 1, RIWAYAT_HEADERS.length).getValues()[0];
    const rowPlus = sheet.getRange(rowPlusIdx, 1, 1, RIWAYAT_HEADERS.length).getValues()[0];
    if (String(rowMinus[18]) === 'Selesai' || String(rowPlus[18]) === 'Selesai') {
      return {
        success: false,
        message: 'Salah satu task sudah Selesai, tidak bisa dipasangkan lagi.'
      };
    }
    const selisihMinus = Number(rowMinus[9]) || 0;
    const selisihPlus = Number(rowPlus[9]) || 0;
    if (selisihMinus >= 0) return {
      success: false,
      message: 'Task Minus yang dipilih ternyata bukan selisih minus.'
    };
    if (selisihPlus <= 0) return {
      success: false,
      message: 'Task Plus yang dipilih ternyata bukan selisih plus.'
    };
    const articleMinus = String(rowMinus[3] || '').trim().toUpperCase();
    const articlePlus = String(rowPlus[3] || '').trim().toUpperCase();
    if (articleMinus !== articlePlus) {
      return {
        success: false,
        message: 'SKU/Article kedua task harus sama (Minus: ' + rowMinus[3] + ', Plus: ' + rowPlus[3] + ').'
      };
    }
    if (Math.abs(selisihMinus) !== Math.abs(selisihPlus)) {
      return {
        success: false,
        message: 'Qty selisih kedua task harus sama persis (pairing 1-ke-1). Minus: ' + Math.abs(selisihMinus) + ', Plus: ' + Math.abs(selisihPlus) + '.'
      };
    }
    const qtyAbs = Math.abs(selisihMinus);
    const lokasiMinus = String(rowMinus[2] || '').trim().toUpperCase();
    const lokasiPlus = String(rowPlus[2] || '').trim().toUpperCase();
    const relevant = buktiRows.filter(function(r) {
      const sku = String(r[3] || '').trim().toUpperCase();
      const tipe = String(r[2] || '').trim().toUpperCase();
      return sku === articleMinus && tipe === 'MOVE';
    });
    if (!relevant.length) {
      return {
        success: false,
        message: 'Tidak ada baris bukti dengan SKU ' + rowMinus[3] + ' dan jenis transaksi MOVE.'
      };
    }
    const totalQty = relevant.reduce(function(s, r) {
      return s + Math.abs(Number(r[12]) || 0);
    }, 0);
    if (Math.round(totalQty) !== Math.round(qtyAbs)) {
      return {
        success: false,
        message: 'Total Qty di bukti MOVE (' + totalQty + ') tidak sama dengan qty pasangan ini (' + qtyAbs + ').'
      };
    }
    const hasLinkedRoute = relevant.some(function(r) {
      return String(r[7] || '').trim().toUpperCase() === lokasiMinus && String(r[9] || '').trim().toUpperCase() === lokasiPlus;
    });
    if (!hasLinkedRoute) {
      return {
        success: false,
        message: 'Bukti MOVE harus punya baris dengan FROMLOC=' + rowMinus[2] + ' (lokasi Minus) dan TOLOC=' + rowPlus[2] + ' (lokasi Plus) -- memastikan bukti ini benar menghubungkan kedua lokasi.'
      };
    }
    const waktu = new Date;
    const stamp = Utilities.formatDate(waktu, 'Asia/Jakarta', 'dd/MM HH:mm');
    const noteSuffix = '\n[' + stamp + ' - Selesai] ' + catatanTrim + '\n[' + stamp + '] Dipasangkan & ditutup via Penyelesaian Plus Minus, bukti MOVE terlampir -- lihat sheet Log_Bukti_Investigasi.';
    function closeRow(rowIdx, pairId, rowSnapshot) {
      const prevCatatan = String(rowSnapshot[19] || '');
      sheet.getRange(rowIdx, 19).setValue('Selesai');
      sheet.getRange(rowIdx, 20).setValue(prevCatatan + noteSuffix);
      sheet.getRange(rowIdx, 18).setValue(KATEGORI_PLUS_MINUS_PAIR_);
      sheet.getRange(rowIdx, 21).setValue(namaUser);
      sheet.getRange(rowIdx, 22).setValue(waktu);
      sheet.getRange(rowIdx, 24).setValue(picTrim);
      sheet.getRange(rowIdx, 25).setValue(waktu);
      sheet.getRange(rowIdx, RIWAYAT_COL_PASANGAN_TASK_ID_).setValue(pairId);
      const newRow = rowSnapshot.slice();
      newRow[17] = KATEGORI_PLUS_MINUS_PAIR_;
      newRow[18] = 'Selesai';
      newRow[19] = prevCatatan + noteSuffix;
      newRow[20] = namaUser;
      newRow[21] = waktu;
      newRow[23] = picTrim;
      newRow[24] = waktu;
      newRow[26] = pairId;
      return newRow;
    }
    const newRowMinus = closeRow(rowMinusIdx, plusTaskId, rowMinus);
    const newRowPlus = closeRow(rowPlusIdx, minusTaskId, rowPlus);
    try {
      applyRingkasanDelta_(rowDateTag_(rowMinus), rowMinus, newRowMinus, namaUser);
      applyRingkasanDelta_(rowDateTag_(rowPlus), rowPlus, newRowPlus, namaUser);
    } catch (e) {}
    try {
      adjustPlusMinusTracker_(rowMinus[9], rowMinus[3], rowMinus[4], rowMinus[2], false, namaUser);
      adjustPlusMinusTracker_(rowPlus[9], rowPlus[3], rowPlus[4], rowPlus[2], false, namaUser);
    } catch (e) {}
    const buktiSheet = getLogBuktiSheet_(namaUser);
    const buktiStartRow = buktiSheet.getLastRow() + 1;
    const buktiOut = [];
    buktiRows.forEach(function(r) {
      buktiOut.push([ waktu, minusTaskId, rowMinus[2], rowMinus[3], namaUser ].concat(r.slice(0, 15)));
    });
    buktiRows.forEach(function(r) {
      buktiOut.push([ waktu, plusTaskId, rowPlus[2], rowPlus[3], namaUser ].concat(r.slice(0, 15)));
    });
    buktiSheet.getRange(buktiStartRow, 1, buktiOut.length, LOG_BUKTI_HEADERS.length).setValues(buktiOut);
    return {
      success: true
    };
  } finally {
    facLock.release();
  }
}

function getTaskLog(limit, requesterUsername, dateFrom, dateTo) {
  requireRole_(requesterUsername, [ 'inventory', 'admin', 'developer' ]);
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const fromDate = dateFrom ? new Date(dateFrom + 'T00:00:00+07:00') : null;
  const toDate = dateTo ? new Date(dateTo + 'T23:59:59+07:00') : null;
  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  const logs = [];
  data.forEach(function(row) {
    const statusTask = row[18];
    const alasan = row[17];
    const hasilFinal = row[14];
    if (statusTask !== 'Selesai' || !alasan) return;
    if (alasan === 'Salah Hitung' && hasilFinal === 'HIT') return;
    const waktuSelesaiDate = Object.prototype.toString.call(row[21]) === '[object Date]' ? row[21] : null;
    if (fromDate && (!waktuSelesaiDate || waktuSelesaiDate < fromDate)) return;
    if (toDate && (!waktuSelesaiDate || waktuSelesaiDate > toDate)) return;
    logs.push({
      tanggal: row[1],
      lokasi: row[2],
      article: row[3],
      description: row[4],
      selisih: row[9],
      namaPetugas: row[7],
      alasan: alasan,
      kategori: alasan,
      picUsername: row[23],
      catatan: row[19],
      diselesaikanOleh: row[20],
      addWhoTransaksi: row[25],
      namaValidator: row[12],
      waktuSelesai: waktuSelesaiDate ? Utilities.formatDate(waktuSelesaiDate, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm') : String(row[21] || '')
    });
  });
  logs.sort(function(a, b) {
    return String(b.waktuSelesai).localeCompare(String(a.waktuSelesai));
  });
  const max = limit ? Number(limit) : 200;
  return logs.slice(0, max);
}

function queueTaskUpdatePayload_(id, namaUser, newStatusTrim, catatanTrim, kategoriTrim, picTrim, buktiRows, facilityId) {
  const sheet = getTaskQueueSheet_();
  appendRowLocked_(sheet, [ new Date, id, namaUser, newStatusTrim, catatanTrim, kategoriTrim, picTrim, JSON.stringify(buktiRows || []), facilityId ]);
}

function processTaskQueue_() {
  const qSheet = getTaskQueueSheet_();
  const lastRow = qSheet.getLastRow();
  if (lastRow < 2) return;
  const values = qSheet.getRange(2, 1, lastRow - 1, TASK_QUEUE_HEADERS.length).getValues();
  const byFacility = {};
  values.forEach(function(row, idx) {
    const facilityId = String(row[8] || '').trim();
    if (!byFacility[facilityId]) byFacility[facilityId] = [];
    byFacility[facilityId].push(idx);
  });
  const processedRowNumbers = [];
  const cache = CacheService.getScriptCache();
  Object.keys(byFacility).forEach(function(facilityId) {
    const indices = byFacility[facilityId];
    if (!facilityId) {
      indices.forEach(function(i) {
        catatLogSistem_('Worker Task', 'Baris antrian tanpa FacilityId dibuang: ' + JSON.stringify(values[i]));
        processedRowNumbers.push(i + 2);
      });
      return;
    }
    const workerGuardKey = 'taskQueueWorkerLock_' + facilityId;
    if (cache.get(workerGuardKey)) {
      return;
    }
    try {
      cache.put(workerGuardKey, '1', 55);
    } catch (e) {}
    const facLock = acquireFacilityLock_(facilityId, 1e4);
    if (!facLock) {
      try {
        cache.remove(workerGuardKey);
      } catch (e) {}
      return;
    }
    try {
      const sortedIndices = indices.slice().sort(function(a, b) {
        return new Date(values[a][0]) - new Date(values[b][0]);
      });
      sortedIndices.forEach(function(i) {
        const row = values[i];
        const id = row[1];
        const namaUser = String(row[2] || '').trim();
        const newStatusTrim = String(row[3] || '').trim();
        const catatanTrim = String(row[4] || '').trim();
        const kategoriTrim = String(row[5] || '').trim();
        const picTrim = String(row[6] || '').trim();
        let buktiRows = [];
        try {
          buktiRows = JSON.parse(row[7] || '[]');
        } catch (e) {}
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
      try {
        cache.remove(workerGuardKey);
      } catch (e) {}
    }
  });
  deleteProcessedQueueRows_(qSheet, processedRowNumbers);
}

function processSingleQueuedTaskUpdate_(id, namaUser, newStatusTrim, catatanTrim, kategoriTrim, picTrim, buktiRows) {
  const targetIdx = TASK_STATUS_FLOW.indexOf(newStatusTrim);
  if (targetIdx <= 0) return {
    success: false,
    message: 'Status tujuan tidak valid (data antrian korup).'
  };
  const isFinal = newStatusTrim === 'Selesai';
  const sheet = getRiwayatSheet_(namaUser);
  const rowIndex = findRiwayatRow_(sheet, id);
  if (rowIndex === -1) return {
    success: false,
    message: 'Task tidak ditemukan, kemungkinan sudah diproses.'
  };
  const rowValues = sheet.getRange(rowIndex, 1, 1, RIWAYAT_HEADERS.length).getValues()[0];
  const currentStatus = String(rowValues[18]);
  const currentIdx = TASK_STATUS_FLOW.indexOf(currentStatus);
  if (currentIdx === -1 || currentStatus === 'Selesai') {
    return {
      success: false,
      message: 'Task ini sudah tidak aktif (mungkin sudah Selesai oleh proses lain).'
    };
  }
  if (targetIdx <= currentIdx) {
    return {
      success: false,
      message: 'Status tidak boleh mundur/sama (sekarang: ' + currentStatus + ') -- kemungkinan sudah diupdate proses lain.'
    };
  }
  let lokasiTask = '', articleTask = '';
  if (isFinal && KATEGORI_TANPA_BUKTI_.indexOf(kategoriTrim) === -1) {
    lokasiTask = String(rowValues[2] || '').trim();
    articleTask = String(rowValues[3] || '').trim();
    const selisihAbs = Math.abs(Number(rowValues[9]) || 0);
    const cek = validateBuktiTransaksi_(buktiRows, articleTask, selisihAbs, kategoriTrim, lokasiTask);
    if (!cek.valid) return {
      success: false,
      message: 'Bukti tidak valid saat diproses ulang (data mungkin sudah berubah): ' + cek.message
    };
  }
  const waktu = new Date;
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
  newRow[18] = newStatusTrim;
  newRow[19] = newCatatan;
  newRow[24] = waktu;
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
    newRow[20] = namaUser;
    newRow[21] = waktu;
  }
  if (isFinal && KATEGORI_TANPA_BUKTI_.indexOf(kategoriTrim) === -1 && buktiRows && buktiRows.length) {
    const buktiSheet = getLogBuktiSheet_(namaUser);
    const buktiStartRow = buktiSheet.getLastRow() + 1;
    const buktiOut = buktiRows.map(function(r) {
      return [ waktu, id, lokasiTask, articleTask, namaUser ].concat(r.slice(0, 15));
    });
    buktiSheet.getRange(buktiStartRow, 1, buktiOut.length, LOG_BUKTI_HEADERS.length).setValues(buktiOut);
  }
  try {
    applyRingkasanDelta_(rowDateTag_(rowValues), rowValues, newRow, namaUser);
  } catch (e) {}
  if (isFinal && rowValues[14] === 'DISCREPANCY') {
    try {
      adjustPlusMinusTracker_(rowValues[9], rowValues[3], rowValues[4], rowValues[2], false, namaUser);
    } catch (e) {}
  }
  return {
    success: true
  };
}

function installTaskQueueTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processTaskQueue_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processTaskQueue_').timeBased().everyMinutes(1).create();
  Logger.log('Trigger terpasang: processTaskQueue_ tiap 1 menit.');
}

const MASTER_ROLE_VALID_ = [ 'admin', 'inventory', 'outbound', 'storing', 'inbound', 'lp', 'maintenance', 'developer' ];

function getDaftarUserMaster(requesterUsername) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Kelola User.'
    };
  }
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    const muSheet = getMasterUserSheetWithFallback_();
    if (muSheet) {
      const lastRow = muSheet.getLastRow();
      const users = [];
      if (lastRow >= 2) {
        const values = muSheet.getRange(2, 1, lastRow - 1, 4).getValues();
        values.forEach(function(r, idx) {
          const username = String(r[0] || '').trim();
          if (!username) return;
          const facId = String(r[3] || '').trim();
          let facNama = '';
          if (facId && typeof getFacilityInfoById_ === 'function') {
            const facInfo = getFacilityInfoById_(facId);
            if (facInfo) facNama = facInfo.nama;
          }
          users.push({
            rowIndex: idx + 2,
            username: username,
            role: String(r[1] || '').trim(),
            status: isBarisUserAktif_(r[2]) ? 'Aktif' : 'Nonaktif',
            facilityId: facId,
            facilityNama: facNama
          });
        });
      }
      return {
        success: true,
        users: users,
        roleValid: MASTER_ROLE_VALID_
      };
    }
  }
  const sheet = getMasterSheet_();
  if (!sheet) return {
    success: false,
    message: 'Sheet Master tidak ditemukan.'
  };
  const lastRow = sheet.getLastRow();
  const users = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 3, lastRow - 1, 3).getValues();
    values.forEach(function(r, idx) {
      const username = String(r[0] || '').trim();
      if (!username) return;
      users.push({
        rowIndex: idx + 2,
        username: username,
        role: String(r[1] || '').trim(),
        status: isBarisUserAktif_(r[2]) ? 'Aktif' : 'Nonaktif',
        facilityId: '',
        facilityNama: ''
      });
    });
  }
  return {
    success: true,
    users: users,
    roleValid: MASTER_ROLE_VALID_
  };
}

const PESAN_KHUSUS_DEVELOPER_ = 'Akun dan peran Developer hanya bisa diubah oleh Developer.';

function roleDeveloper_(roleRaw) {
  const r = String(roleRaw || '').trim().toLowerCase();
  return r === 'developer' || r === 'dev' || r === 'dewa';
}

function requesterDeveloper_(requesterUsername) {
  return isUsernameRoleDeveloper_(String(requesterUsername || '').trim().toLowerCase());
}

function tambahUserMaster(requesterUsername, usernameBaru, roleBaru, facilityId) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Kelola User.'
    };
  }
  const uname = String(usernameBaru || '').trim();
  const role = String(roleBaru || '').trim().toLowerCase();
  let facId = String(facilityId || '').trim();
  let facNama = '';
  if (!uname) return {
    success: false,
    message: 'NIK/Username tidak boleh kosong.'
  };
  if (MASTER_ROLE_VALID_.indexOf(role) === -1) {
    return {
      success: false,
      message: 'Role tidak valid. Pilih salah satu: ' + MASTER_ROLE_VALID_.join(', ') + '.'
    };
  }
  if (role === 'developer' && !requesterDeveloper_(requesterUsername)) {
    return {
      success: false,
      message: PESAN_KHUSUS_DEVELOPER_
    };
  }
  if (!facId && role !== 'developer') {
    if (typeof getUserFacility === 'function') {
      const requesterFac = getUserFacility(requesterUsername);
      if (requesterFac) {
        facId = requesterFac.id;
        facNama = requesterFac.nama || requesterFac.name || '';
      }
    }
  }
  let sheet = null;
  let isNewFormat = false;
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    sheet = getMasterUserSheetWithFallback_();
    if (sheet) isNewFormat = true;
  }
  if (!sheet) sheet = getMasterSheet_();
  if (!sheet) return {
    success: false,
    message: 'Sheet Master tidak ditemukan.'
  };
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const colIdx = isNewFormat ? 1 : 3;
    const existing = sheet.getRange(2, colIdx, lastRow - 1, 1).getValues();
    for (let i = 0; i < existing.length; i++) {
      if (String(existing[i][0] || '').trim().toLowerCase() === uname.toLowerCase()) {
        return {
          success: false,
          message: 'NIK/Username "' + uname + '" sudah terdaftar (baris ' + (i + 2) + ').'
        };
      }
    }
  }
  const targetRow = Math.max(lastRow + 1, 2);
  if (isNewFormat) {
    sheet.getRange(targetRow, 1).setValue(uname);
    sheet.getRange(targetRow, 2).setValue(role);
    sheet.getRange(targetRow, 3).setValue('Aktif');
    sheet.getRange(targetRow, 4).setValue(facId);
    if (facId && typeof assignUserKeFacility === 'function') {
      assignUserKeFacility(requesterUsername, uname, facId);
    }
  } else {
    sheet.getRange(targetRow, 3).setValue(uname);
    sheet.getRange(targetRow, 4).setValue(role);
    sheet.getRange(targetRow, 5).setValue('Aktif');
  }
  clearMasterCache_();
  catatLogPerubahanConfig_(requesterUsername, 'Manajemen User', 'Tambah user "' + uname + '" role ' + role + (facId ? ' ke facility ' + facId : '') + '.');
  return {
    success: true,
    message: 'User "' + uname + '" berhasil ditambahkan' + (facId ? ' ke facility ' + (facNama || facId) + '.' : '.')
  };
}

function updateRoleUserMaster(requesterUsername, rowIndex, usernameKonfirmasi, roleBaru, facilityId) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Kelola User.'
    };
  }
  const role = String(roleBaru || '').trim().toLowerCase();
  const facDikirim = facilityId !== undefined && facilityId !== null;
  const facId = String(facilityId || '').trim();
  if (MASTER_ROLE_VALID_.indexOf(role) === -1) {
    return {
      success: false,
      message: 'Role tidak valid.'
    };
  }
  let sheet = null;
  let isNewFormat = false;
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    sheet = getMasterUserSheetWithFallback_();
    if (sheet) isNewFormat = true;
  }
  if (!sheet) sheet = getMasterSheet_();
  if (!sheet) return {
    success: false,
    message: 'Sheet Master tidak ditemukan.'
  };
  const colIdx = isNewFormat ? 1 : 3;
  const roleColIdx = isNewFormat ? 2 : 4;
  const actualUsername = String(sheet.getRange(rowIndex, colIdx).getValue() || '').trim();
  if (!actualUsername || actualUsername.toLowerCase() !== String(usernameKonfirmasi || '').trim().toLowerCase()) {
    return {
      success: false,
      message: 'Data user sudah berubah sejak halaman dibuka.'
    };
  }
  const roleLama = String(sheet.getRange(rowIndex, roleColIdx).getValue() || '').trim();
  if ((role === 'developer' || roleDeveloper_(roleLama)) && roleDeveloper_(roleLama) !== (role === 'developer') && !requesterDeveloper_(requesterUsername)) {
    return {
      success: false,
      message: PESAN_KHUSUS_DEVELOPER_
    };
  }
  sheet.getRange(rowIndex, roleColIdx).setValue(role);
  if (isNewFormat && facDikirim) {
    sheet.getRange(rowIndex, 4).setValue(facId);
    if (facId && typeof assignUserKeFacility === 'function') {
      assignUserKeFacility(requesterUsername, actualUsername, facId);
    }
  }
  clearMasterCache_();
  clearUserRoleCache_(actualUsername);
  catatLogPerubahanConfig_(requesterUsername, 'Manajemen User', 'Data "' + actualUsername + '" diperbarui.');
  return {
    success: true,
    message: 'Data "' + actualUsername + '" berhasil diperbarui.'
  };
}

function setStatusUserMaster(requesterUsername, rowIndex, usernameKonfirmasi, statusBaru) {
  if (!isNikPunyaAksesSetting_(requesterUsername)) {
    return {
      success: false,
      message: 'NIK ini tidak punya akses ke Kelola User.'
    };
  }
  const status = String(statusBaru || '').trim().toLowerCase() === 'nonaktif' ? 'Nonaktif' : 'Aktif';
  let sheet = null;
  let isNewFormat = false;
  if (typeof getMasterUserSheetWithFallback_ === 'function') {
    sheet = getMasterUserSheetWithFallback_();
    if (sheet) isNewFormat = true;
  }
  if (!sheet) sheet = getMasterSheet_();
  if (!sheet) return {
    success: false,
    message: 'Sheet Master tidak ditemukan.'
  };
  const colIdx = isNewFormat ? 1 : 3;
  const statusColIdx = isNewFormat ? 3 : 5;
  const actualUsername = String(sheet.getRange(rowIndex, colIdx).getValue() || '').trim();
  if (!actualUsername || actualUsername.toLowerCase() !== String(usernameKonfirmasi || '').trim().toLowerCase()) {
    return {
      success: false,
      message: 'Data user sudah berubah.'
    };
  }
  if (status === 'Nonaktif') {
    if (actualUsername.toLowerCase() === String(requesterUsername || '').trim().toLowerCase()) {
      return {
        success: false,
        message: 'Akun sendiri tidak bisa dinonaktifkan. Minta pemegang akses Config lain.'
      };
    }
    const roleTarget = String(sheet.getRange(rowIndex, isNewFormat ? 2 : 4).getValue() || '');
    if (roleDeveloper_(roleTarget) && !requesterDeveloper_(requesterUsername)) {
      return {
        success: false,
        message: PESAN_KHUSUS_DEVELOPER_
      };
    }
  }
  sheet.getRange(rowIndex, statusColIdx).setValue(status);
  clearMasterCache_();
  clearUserRoleCache_(actualUsername);
  catatLogPerubahanConfig_(requesterUsername, 'Manajemen User', 'Status "' + actualUsername + '" diubah jadi ' + status + '.');
  return {
    success: true,
    message: 'Status "' + actualUsername + '" berhasil diubah menjadi ' + status + '.'
  };
}

function encodeAddWhoByType_(byType) {
  const order = [ 'move', 'picking', 'lainnya' ];
  const parts = [];
  order.forEach(function(t) {
    const names = Object.keys(byType && byType[t] || {}).filter(function(n) {
      return n;
    }).sort();
    if (names.length) parts.push(t.toUpperCase() + ':' + names.join(','));
  });
  return parts.join(';');
}

function decodeAddWhoByType_(str) {
  const s = String(str || '').trim();
  const result = {
    move: {},
    picking: {},
    lainnya: {}
  };
  if (!s) return result;
  if (/(MOVE|PICKING|LAINNYA):/i.test(s)) {
    s.split(';').forEach(function(part) {
      const idx = part.indexOf(':');
      if (idx === -1) return;
      const type = part.substring(0, idx).trim().toLowerCase();
      const names = part.substring(idx + 1).split(',').map(function(x) {
        return x.trim();
      }).filter(function(x) {
        return x;
      });
      if (!result[type]) result[type] = {};
      names.forEach(function(n) {
        result[type][n] = true;
      });
    });
  } else {
    s.split(',').map(function(x) {
      return x.trim();
    }).filter(function(x) {
      return x;
    }).forEach(function(n) {
      result.lainnya[n] = true;
    });
  }
  return result;
}

function mergeAddWho_(existingStr, newByType) {
  const merged = decodeAddWhoByType_(existingStr);
  [ 'move', 'picking', 'lainnya' ].forEach(function(t) {
    Object.keys(newByType && newByType[t] || {}).forEach(function(n) {
      if (!merged[t]) merged[t] = {};
      merged[t][n] = true;
    });
  });
  return encodeAddWhoByType_(merged);
}

function extractSpreadsheetId_(idAtauUrl) {
  const s = String(idAtauUrl || '').trim();
  if (!s) return '';
  const match = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  return s;
}

function extractLevel_(lokasi) {
  const parts = String(lokasi || '').split('.');
  const last = parts[parts.length - 1] || '';
  const match = last.match(/\d+/);
  if (!match) return null;
  return parseInt(match[0], 10);
}

function isRollcageLocation_(lokasi) {
  const parts = String(lokasi || '').split('.');
  const last = parts[parts.length - 1] || '';
  return /K/i.test(last);
}

function getLevelGroupByLocation_(lokasi) {
  if (isRollcageLocation_(lokasi)) return LEVEL_GROUPS[0];
  return getLevelGroup_(extractLevel_(lokasi));
}

function extractArea_(lokasi) {
  const raw = String(lokasi || '').trim().toUpperCase();
  const m = raw.match(/^([A-Z]+)/);
  if (!m) return '';
  const alpha = m[1];
  if (alpha.length >= 2 && alpha.charAt(0) === 'W') return alpha.charAt(1);
  return alpha.charAt(0);
}

function parseLocation_(lokasi) {
  const raw = String(lokasi || '').trim().toUpperCase();
  const parts = raw.split('.');
  const first = parts[0] || '';
  const firstMatch = first.match(/^([A-Z]+)?(\d+)?/);
  const gudang = firstMatch && firstMatch[1] ? firstMatch[1] : '';
  const lorong = firstMatch && firstMatch[2] ? parseInt(firstMatch[2], 10) : 999999;
  const sectionRaw = parts.length > 1 ? parts[1] : '';
  const sectionMatch = sectionRaw.match(/\d+/);
  const section = sectionMatch ? parseInt(sectionMatch[0], 10) : 999999;
  const level = extractLevel_(raw);
  return {
    raw: raw,
    gudang: gudang,
    lorong: lorong,
    section: section,
    level: level === null ? 999999 : level
  };
}

function compareLocation_(a, b) {
  const la = a.locationSort || parseLocation_(a.lokasi);
  const lb = b.locationSort || parseLocation_(b.lokasi);
  if (la.gudang !== lb.gudang) return la.gudang < lb.gudang ? -1 : 1;
  if (la.lorong !== lb.lorong) return la.lorong - lb.lorong;
  if (la.section !== lb.section) return la.section - lb.section;
  if (la.level !== lb.level) return la.level - lb.level;
  return la.raw < lb.raw ? -1 : la.raw > lb.raw ? 1 : 0;
}

function compareLocationEvenOddSection_(a, b) {
  const la = a.locationSort || parseLocation_(a.lokasi);
  const lb = b.locationSort || parseLocation_(b.lokasi);
  if (la.gudang !== lb.gudang) return la.gudang < lb.gudang ? -1 : 1;
  if (la.lorong !== lb.lorong) return la.lorong - lb.lorong;
  const parityA = la.section % 2 === 0 ? 0 : 1;
  const parityB = lb.section % 2 === 0 ? 0 : 1;
  if (parityA !== parityB) return parityA - parityB;
  if (la.section !== lb.section) return la.section - lb.section;
  if (la.level !== lb.level) return la.level - lb.level;
  return la.raw < lb.raw ? -1 : la.raw > lb.raw ? 1 : 0;
}

function getLevelGroup_(level) {
  if (level === null || level === undefined || isNaN(level) || level < 1) return LEVEL_GROUPS[0];
  for (let i = 0; i < LEVEL_GROUPS.length; i++) {
    const g = LEVEL_GROUPS[i];
    if (level >= g.min && level <= g.max) return g;
  }
  return LEVEL_GROUPS[LEVEL_GROUPS.length - 1];
}

function waktuKlienSah_(ms) {
  const t = Number(ms);
  if (!t || isNaN(t)) return null;
  const now = (new Date).getTime();
  if (t > now + 2 * 6e4 || t < now - 12 * 36e5) return null;
  return new Date(t);
}

function waktuBarisAntrean_(waktuHitung, timestampAntrean) {
  const isDate = function(v) {
    return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
  };
  if (isDate(waktuHitung)) return waktuHitung;
  if (isDate(timestampAntrean)) return timestampAntrean;
  return new Date;
}

function rowDateTag_(row) {
  const waktuCycle = row[15];
  if (Object.prototype.toString.call(waktuCycle) === '[object Date]') {
    return Utilities.formatDate(waktuCycle, 'Asia/Jakarta', 'yyyy-MM-dd');
  }
  const tglRaw = row[1];
  return Object.prototype.toString.call(tglRaw) === '[object Date]' ? Utilities.formatDate(tglRaw, 'Asia/Jakarta', 'yyyy-MM-dd') : String(tglRaw);
}

function parseJsonSafe_(str, fallback) {
  if (!str) return fallback;
  try {
    return JSON.parse(str);
  } catch (e) {
    return fallback;
  }
}

function reportProgress_(jobId, percent, label) {
  if (!jobId) return;
  try {
    CacheService.getScriptCache().put('progress_' + jobId, JSON.stringify({
      percent: percent,
      label: label || ''
    }), 300);
  } catch (e) {}
}

function getImportProgress(jobId) {
  if (!jobId) return {
    percent: 0,
    label: ''
  };
  try {
    const raw = CacheService.getScriptCache().get('progress_' + jobId);
    return raw ? JSON.parse(raw) : {
      percent: 0,
      label: ''
    };
  } catch (e) {
    return {
      percent: 0,
      label: ''
    };
  }
}

function rewriteQueueSheetGeneric_(sheet, keepRows, numCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, numCols).clearContent();
  }
  if (keepRows.length) {
    sheet.getRange(2, 1, keepRows.length, numCols).setValues(keepRows);
  }
}

function appendRowLocked_(sheet, rowValues) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(1e4);
  } catch (e) {
    throw new Error('Sistem sedang sibuk, coba lagi sebentar.');
  }
  try {
    sheet.appendRow(rowValues);
  } finally {
    lock.releaseLock();
  }
}

function appendRowsLocked_(sheet, rowsArray) {
  if (!rowsArray || !rowsArray.length) return;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(1e4);
  } catch (e) {
    throw new Error('Sistem sedang sibuk, coba lagi sebentar.');
  }
  try {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsArray.length, rowsArray[0].length).setValues(rowsArray);
  } finally {
    lock.releaseLock();
  }
}

function deleteProcessedQueueRows_(sheet, processedRowNumbers) {
  if (!processedRowNumbers || !processedRowNumbers.length) return;
  const sorted = processedRowNumbers.slice().sort(function(a, b) {
    return b - a;
  });
  sorted.forEach(function(rowNum) {
    try {
      sheet.deleteRow(rowNum);
    } catch (e) {}
  });
}

function getPendingValidasi(requesterUsername) {
  const info = requireRole_(requesterUsername, [ 'inventory', 'admin', 'developer' ]);
  const sheet = getRiwayatSheet_(requesterUsername);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, RIWAYAT_HEADERS.length).getValues();
  const items = [];
  data.forEach(function(row) {
    if (row[11] !== 'Pending') return;
    const assignedValidator = String(row[22] || '').trim();
    if (info.role === 'inventory' && assignedValidator !== info.displayName) return;
    items.push({
      id: row[0],
      lokasi: row[2],
      article: row[3],
      description: row[4],
      qtySystem: Number(row[6])
    });
  });
  return items;
}

function getPendingValidasiCountData_(info, username) {
  const sheet = getRiwayatSheet_(username);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const data = sheet.getRange(2, 12, lastRow - 1, 12).getValues();
  let count = 0;
  data.forEach(function(r) {
    const status = r[0];
    const assignedValidator = String(r[11] || '').trim();
    if (status !== 'Pending') return;
    if (info.role === 'inventory' && assignedValidator !== info.displayName) return;
    count++;
  });
  return count;
}

function getPendingValidasiCount(requesterUsername) {
  const info = requireRole_(requesterUsername, [ 'inventory', 'admin', 'developer' ]);
  return getPendingValidasiCountData_(info, requesterUsername);
}

function submitValidasi(id, namaValidator, qtyValidasi, waktuKlien, facKlien) {
  requireRole_(namaValidator, [ 'inventory', 'admin', 'developer' ]);
  const aktual = Number(qtyValidasi);
  if (qtyValidasi === '' || qtyValidasi === null || qtyValidasi === undefined || isNaN(aktual) || aktual < 0) {
    return {
      success: false,
      message: 'Qty hasil validasi tidak valid.'
    };
  }
  const facInfo = requireUserFacility_(namaValidator);
  if (!facInfo) {
    return {
      success: false,
      message: PESAN_TANPA_FACILITY_
    };
  }
  if (!facilityKlienCocok_(facKlien, facInfo)) {
    return {
      success: false,
      kode: 'FACILITY_BERUBAH',
      message: PESAN_FACILITY_BERUBAH_
    };
  }
  try {
    queueValidasiPayload_(id, namaValidator, aktual, facInfo.id, waktuKlienSah_(waktuKlien));
  } catch (e) {
    return {
      success: false,
      message: 'Gagal masuk antrian validasi: ' + e.message
    };
  }
  return {
    success: true,
    queued: true,
    message: 'Validasi diterima.'
  };
}

function queueValidasiPayload_(id, namaValidator, qtyValidasi, facilityId, waktuHitung) {
  const sheet = getValidasiQueueSheet_();
  appendRowLocked_(sheet, [ new Date, id, namaValidator, qtyValidasi, facilityId, waktuHitung || '' ]);
}

function processValidasiQueue_() {
  const qSheet = getValidasiQueueSheet_();
  const lastRow = qSheet.getLastRow();
  if (lastRow < 2) return;
  const values = qSheet.getRange(2, 1, lastRow - 1, VALIDASI_QUEUE_HEADERS.length).getValues();
  const byFacility = {};
  values.forEach(function(row, idx) {
    const facilityId = String(row[4] || '').trim();
    if (!byFacility[facilityId]) byFacility[facilityId] = [];
    byFacility[facilityId].push(idx);
  });
  const processedRowNumbers = [];
  const cache = CacheService.getScriptCache();
  Object.keys(byFacility).forEach(function(facilityId) {
    const indices = byFacility[facilityId];
    if (!facilityId) {
      indices.forEach(function(i) {
        catatLogSistem_('Worker Validasi', 'Baris antrian tanpa FacilityId dibuang: ' + JSON.stringify(values[i]));
        processedRowNumbers.push(i + 2);
      });
      return;
    }
    const workerGuardKey = 'validasiQueueWorkerLock_' + facilityId;
    if (cache.get(workerGuardKey)) {
      return;
    }
    try {
      cache.put(workerGuardKey, '1', 55);
    } catch (e) {}
    const facLock = acquireFacilityLock_(facilityId, 1e4);
    if (!facLock) {
      try {
        cache.remove(workerGuardKey);
      } catch (e) {}
      return;
    }
    try {
      const sortedIndices = indices.slice().sort(function(a, b) {
        return new Date(values[a][0]) - new Date(values[b][0]);
      });
      sortedIndices.forEach(function(i) {
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
      try {
        cache.remove(workerGuardKey);
      } catch (e) {}
    }
  });
  deleteProcessedQueueRows_(qSheet, processedRowNumbers);
}

function processSingleQueuedValidasi_(id, namaValidator, qtyValidasi, facilityId, waktuHitung) {
  const aktual = Number(qtyValidasi);
  const sheet = getRiwayatSheet_(namaValidator);
  const rowIndex = findRiwayatRow_(sheet, id);
  if (rowIndex === -1) {
    return {
      success: false,
      message: 'Item tidak ditemukan (kemungkinan sudah divalidasi orang lain).'
    };
  }
  const rowValues = sheet.getRange(rowIndex, 1, 1, RIWAYAT_HEADERS.length).getValues()[0];
  if (rowValues[11] !== 'Pending') {
    return {
      success: false,
      message: 'Sudah divalidasi sebelumnya (duplikat submit).'
    };
  }
  const qtySystem = Number(rowValues[6]);
  const hasilFinal = aktual === qtySystem ? 'HIT' : 'DISCREPANCY';
  const waktu = waktuHitung || new Date;
  const newRow = rowValues.slice();
  sheet.getRange(rowIndex, 12, 1, 4).setValues([ [ 'Selesai', namaValidator, aktual, hasilFinal ] ]);
  sheet.getRange(rowIndex, 14).setNumberFormat('0');
  sheet.getRange(rowIndex, 17).setValue(waktu);
  newRow[11] = 'Selesai';
  newRow[12] = namaValidator;
  newRow[13] = aktual;
  newRow[14] = hasilFinal;
  newRow[16] = waktu;
  if (hasilFinal === 'HIT') {
    const catatanOtomatis = 'Otomatis: hasil validasi ulang sesuai stok system, tidak ada discrepancy nyata (kesalahan hitung awal petugas).';
    sheet.getRange(rowIndex, 18, 1, 4).setValues([ [ 'Salah Hitung', 'Selesai', catatanOtomatis, namaValidator ] ]);
    sheet.getRange(rowIndex, 22).setValue(waktu);
    sheet.getRange(rowIndex, 25).setValue(waktu);
    newRow[17] = 'Salah Hitung';
    newRow[18] = 'Selesai';
    newRow[19] = catatanOtomatis;
    newRow[20] = namaValidator;
    newRow[21] = waktu;
    newRow[24] = waktu;
  } else {
    const initialPic = assignInitialPIC_(facilityId);
    sheet.getRange(rowIndex, 19).setValue('Open');
    sheet.getRange(rowIndex, 24).setValue(initialPic);
    sheet.getRange(rowIndex, 25).setValue(waktu);
    newRow[18] = 'Open';
    newRow[23] = initialPic;
    newRow[24] = waktu;
  }
  try {
    applyRingkasanDelta_(rowDateTag_(rowValues), rowValues, newRow, namaValidator);
  } catch (e) {}
  try {
    incrementCounterKey_('pending_total', -1, namaValidator);
    if (rowValues[22]) incrementCounterKey_('pending_validator:' + rowValues[22], -1, namaValidator);
    if (hasilFinal === 'DISCREPANCY') {
      adjustPlusMinusTracker_(rowValues[9], rowValues[3], rowValues[4], rowValues[2], true, namaValidator);
    }
  } catch (e) {}
  return {
    success: true,
    hasilFinal: hasilFinal
  };
}

function installValidasiQueueTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processValidasiQueue_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processValidasiQueue_').timeBased().everyMinutes(1).create();
  Logger.log('Trigger terpasang: processValidasiQueue_ tiap 1 menit.');
}
const PANEL_HTML_ = "<!DOCTYPE html><html lang=\"id\"><head><meta charset=\"utf-8\"><base target=\"_top\">\n<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\"><link rel=\"stylesheet\" href=\"https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap\">\n<script>try{var t=JSON.parse(localStorage.getItem('ct.tema'));if(t==='gelap')document.documentElement.setAttribute('data-tema','gelap');}catch(e){}</script>\n<style>/* Token desain Cycle Transaksi. Lihat docs/DESAIN.md untuk alasan tiap pilihan. */\n:root {\n  /* Dasar */\n  --lantai: #F3F4F7;      /* latar halaman: lantai gudang */\n  --lantai-2: #E9EBF0;    /* isian tertekan, jalur progres */\n  --kertas: #FFFFFF;      /* kartu, lembar */\n  --tinta: #191B20;       /* teks utama */\n  --abu: #646B7A;         /* teks pendukung */\n  --abu-2: #8D94A3;       /* teks redup, ikon pasif */\n  --garis: #E2E5EB;       /* garis tipis */\n  --garis-kuat: #CDD2DB;\n\n  /* Aksi */\n  --biru: #0A66E4;        /* Biru Rak: tombol utama, Cycle */\n  --biru-tua: #0850B8;\n  --biru-muda: #E7F0FE;\n  --di-warna: #FFFFFF;    /* teks & ikon di atas isian berwarna (tombol, ubin ikon, lencana) */\n  --di-jingga: #201400;   /* jingga cukup terang: angka lencana di atasnya gelap, di kedua tema */\n  --seg-aktif: #FFFFFF;   /* pilihan aktif pada kontrol segmen */\n\n  /* Label rak: satu-satunya elemen yang \"keras\" */\n  --kuning: #FFD326;\n  --kuning-tepi: #E3B700;\n  --tinta-label: #17150A;\n\n  /* Status: warna = arti, dipakai sama di KPI, lencana, grafik */\n  --jingga: #E8830C;  --jingga-teks: #A85A00;  --jingga-muda: #FDF0DE;   /* belum di cycle */\n  --ungu: #7447E0;    --ungu-teks: #5B31C4;    --ungu-muda: #EFEAFD;     /* validasi */\n  --hijau: #1A9D53;   --hijau-teks: #11783D;   --hijau-muda: #E4F5EA;    /* selesai, cocok, plus */\n  --merah: #DC3A3A;   --merah-teks: #BF2626;   --merah-muda: #FDEBEB;    /* minus, kritis, galat */\n\n  /* Bentuk: radius mengikuti hierarki, bukan satu angka untuk semua */\n  --r-kartu: 20px;\n  --r-ubin: 13px;\n  --r-isian: 14px;\n  --r-tombol: 16px;\n  --r-lembar: 26px;\n  --bayang-melayang: 0 10px 30px rgba(22, 27, 45, 0.13), 0 2px 6px rgba(22, 27, 45, 0.06);\n  --bayang-lembar: 0 -10px 44px rgba(16, 20, 34, 0.20);\n\n  /* Huruf */\n  --huruf: 'Barlow', system-ui, -apple-system, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;\n  --huruf-rapat: 'Barlow Condensed', 'Barlow', 'Arial Narrow', system-ui, sans-serif;\n\n  /* Gerak */\n  --luwes: cubic-bezier(0.2, 0.8, 0.2, 1);\n\n  /* Area aman (bilah status & gestur Android) */\n  --sat: var(--safe-area-inset-top, env(safe-area-inset-top, 0px));\n  --sab: var(--safe-area-inset-bottom, env(safe-area-inset-bottom, 0px));\n\n  color-scheme: light;\n}\n\nhtml[data-tema='gelap'] {\n  --lantai: #0F1115;\n  --lantai-2: #23272F;\n  --kertas: #191C22;\n  --tinta: #ECEEF3;\n  --abu: #A0A7B5;\n  --abu-2: #7A8191;\n  --garis: #2A2F39;\n  --garis-kuat: #3A404D;\n\n  --biru: #5B9DFF;\n  --biru-tua: #3F86F2;\n  --biru-muda: #16263F;\n  --di-warna: #0B1220;    /* di tema gelap warna isian lebih terang, jadi isinya gelap */\n  --seg-aktif: #3A4150;\n\n  --jingga: #F29A2E;  --jingga-teks: #F6AE55;  --jingga-muda: #33240F;\n  --ungu: #9A78F2;    --ungu-teks: #B39BF7;    --ungu-muda: #251C40;\n  --hijau: #35B86D;   --hijau-teks: #5FCB8C;   --hijau-muda: #132C1E;\n  --merah: #F06060;   --merah-teks: #F58585;   --merah-muda: #371A1A;\n\n  --bayang-melayang: 0 10px 30px rgba(0, 0, 0, 0.5), 0 2px 6px rgba(0, 0, 0, 0.3);\n  --bayang-lembar: 0 -10px 44px rgba(0, 0, 0, 0.55);\n\n  color-scheme: dark;\n}\n\n/* Dasar: reset, tata letak cangkang, dan komponen yang dipakai semua layar. */\n*, *::before, *::after { box-sizing: border-box; }\nhtml { -webkit-text-size-adjust: 100%; text-size-adjust: 100%; }\nbody {\n  margin: 0; background: var(--lantai); color: var(--tinta);\n  font: 400 15px/1.45 var(--huruf); -webkit-font-smoothing: antialiased;\n  -webkit-tap-highlight-color: transparent; overscroll-behavior-y: none;\n}\nbutton, input, select, textarea { font: inherit; color: inherit; }\nbutton { border: 0; background: none; padding: 0; cursor: pointer; text-align: inherit; }\nh1, h2, h3, p { margin: 0; }\na { color: var(--biru); text-decoration: none; }\n[hidden] { display: none !important; }\n:focus { outline: none; }\n:focus-visible { outline: 2.5px solid var(--biru); outline-offset: 2px; }\n.angka { font-variant-numeric: tabular-nums; }\n.redup { color: var(--abu); }\n.sr { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); white-space: nowrap; }\n\n/* ---------- cangkang ---------- */\n#app { max-width: 520px; margin: 0 auto; min-height: 100vh; min-height: 100dvh; padding: calc(var(--sat) + 14px) 16px calc(var(--sab) + 108px); }\nbody.tanpa-nav #app { padding-bottom: calc(var(--sab) + 24px); }\nbody.ada-pita #app { padding-top: calc(var(--sat) + 14px + var(--pita-t, 28px)); }\n.scr { display: none; }\n.scr.is-on { display: block; animation: muncul 0.16s ease-out; }\n@keyframes muncul { from { opacity: 0; } to { opacity: 1; } }\n\n.hd { display: flex; align-items: center; gap: 10px; min-height: 44px; margin-bottom: 14px; }\n.hd__teks { flex: 1; min-width: 0; }\n.hd__judul { font: 700 26px/1.15 var(--huruf); letter-spacing: -0.01em; }\n.hd__ket { font-size: 13.5px; color: var(--abu); margin-top: 2px; }\n.bag { font: 700 16px/1.3 var(--huruf); margin: 22px 2px 10px; display: flex; align-items: baseline; justify-content: space-between; gap: 10px; }\n.bag small { font: 500 13px/1.3 var(--huruf); color: var(--abu); }\n.tumpuk > * + * { margin-top: 12px; }\n\n/* ---------- ikon ---------- */\n.ic { width: 22px; height: 22px; flex: none; fill: none; stroke: currentColor; stroke-width: 1.9; stroke-linecap: round; stroke-linejoin: round; }\n.ic--kecil { width: 18px; height: 18px; }\n.ic--besar { width: 28px; height: 28px; }\n\n/* ---------- tombol ---------- */\n.btn {\n  display: flex; align-items: center; justify-content: center; gap: 8px; width: 100%; min-height: 52px; padding: 0 18px;\n  border-radius: var(--r-tombol); background: var(--biru); color: var(--di-warna); font: 600 16px/1.2 var(--huruf); text-align: center;\n  transition: transform 0.08s, background-color 0.15s, opacity 0.15s;\n}\n.btn:active { transform: scale(0.985); background: var(--biru-tua); }\n.btn[disabled] { opacity: 0.42; pointer-events: none; }\n.btn--tenang { background: var(--biru-muda); color: var(--biru); }\n.btn--tenang:active { background: var(--biru-muda); }\n.btn--garis { background: var(--kertas); color: var(--tinta); box-shadow: inset 0 0 0 1px var(--garis-kuat); }\n.btn--garis:active { background: var(--lantai); }\n.btn--bahaya { background: var(--merah); }\n.btn--bahaya:active { background: var(--merah-teks); }\n.btn--teks { background: none; color: var(--biru); min-height: 44px; width: auto; padding: 0 8px; }\n.btn--teks:active { background: none; opacity: 0.6; }\n.btn--kecil { min-height: 40px; width: auto; padding: 0 14px; border-radius: 12px; font-size: 14.5px; }\n.btn.is-sibuk { pointer-events: none; }\n.btn.is-sibuk > * { visibility: hidden; }\n.btn.is-sibuk { position: relative; color: transparent; }\n.btn.is-sibuk::after { content: ''; position: absolute; width: 20px; height: 20px; border-radius: 50%; border: 2.5px solid var(--di-warna); border-right-color: transparent; animation: putar 0.7s linear infinite; }\n.btn--tenang.is-sibuk::after, .btn--garis.is-sibuk::after { border-color: var(--biru); border-right-color: transparent; }\n@keyframes putar { to { transform: rotate(360deg); } }\n.baris-tombol { display: flex; gap: 10px; }\n.baris-tombol > * { flex: 1; }\n\n.iconbtn { width: 44px; height: 44px; border-radius: 14px; display: grid; place-items: center; color: var(--tinta); background: var(--kertas); box-shadow: inset 0 0 0 1px var(--garis); flex: none; }\n.iconbtn:active { background: var(--lantai-2); }\n.iconbtn--polos { background: none; box-shadow: none; }\n.tautan { color: var(--biru); font-weight: 600; padding: 6px 2px; }\n\n/* ---------- isian ---------- */\n.lbl { display: block; font: 600 13.5px/1.3 var(--huruf); margin: 16px 0 6px; }\n.lbl:first-child { margin-top: 0; }\n.lbl small { font-weight: 400; color: var(--abu); }\n.field {\n  display: block; width: 100%; min-height: 52px; padding: 0 14px; border-radius: var(--r-isian); border: 1.5px solid var(--garis-kuat);\n  background: var(--kertas); font-size: 16px; -webkit-appearance: none; appearance: none; transition: border-color 0.15s, box-shadow 0.15s;\n}\n.field::placeholder { color: var(--abu-2); }\n.field:focus { border-color: var(--biru); box-shadow: 0 0 0 4px rgba(10, 102, 228, 0.15); outline: none; }\n.field.is-galat { border-color: var(--merah); }\ntextarea.field { padding: 12px 14px; min-height: 88px; line-height: 1.4; resize: vertical; }\n/* tempelan baris dari WMS/Excel: baris panjang tidak dilipat supaya tetap terbaca per baris */\ntextarea.field--tempel { font-size: 13.5px; line-height: 1.5; white-space: pre; overflow-x: auto; tab-size: 4; font-variant-numeric: tabular-nums; }\ntextarea.field--tempel::placeholder { font-size: 16px; }\n.blok { margin-top: 16px; }\nselect.field { padding-right: 40px; background-image: url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23646B7A' stroke-width='2.2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='m6 9 6 6 6-6'/%3E%3C/svg%3E\"); background-repeat: no-repeat; background-position: right 12px center; background-size: 20px; }\ninput[type='date'].field, input[type='month'].field { min-width: 0; }\n.field--kecil { min-height: 44px; font-size: 15px; border-radius: 12px; }\n.ket { font-size: 13px; color: var(--abu); margin-top: 6px; line-height: 1.4; }\n.pesan { font-size: 13.5px; margin-top: 8px; min-height: 0; line-height: 1.4; }\n.pesan:empty { display: none; }\n.pesan--galat { color: var(--merah-teks); }\n.pesan--ok { color: var(--hijau-teks); }\n.cari { position: relative; }\n.cari .ic { position: absolute; left: 13px; top: 50%; transform: translateY(-50%); color: var(--abu-2); pointer-events: none; }\n.cari .field { padding-left: 42px; min-height: 46px; border-radius: 13px; border-width: 1px; }\n\n.cek { -webkit-appearance: none; appearance: none; width: 24px; height: 24px; margin: 0; border-radius: 7px; border: 2px solid var(--garis-kuat); background: var(--kertas); display: inline-grid; place-content: center; flex: none; transition: background-color 0.12s, border-color 0.12s; }\n.cek:checked { background: var(--biru); border-color: var(--biru); }\n.cek:checked::after { content: ''; width: 12px; height: 7px; border: 2.5px solid var(--di-warna); border-top: 0; border-right: 0; transform: translateY(-1.5px) rotate(-45deg); }\n.saklar { width: 50px; height: 30px; border-radius: 15px; background: var(--garis-kuat); position: relative; flex: none; transition: background-color 0.15s; }\n.saklar::after { content: ''; position: absolute; top: 3px; left: 3px; width: 24px; height: 24px; border-radius: 50%; background: #fff; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.28); transition: transform 0.18s var(--luwes); }\n.saklar[aria-checked='true'] { background: var(--biru); }\n.saklar[aria-checked='true']::after { transform: translateX(20px); }\n\n.langkah { display: flex; align-items: center; gap: 0; border-radius: var(--r-isian); box-shadow: inset 0 0 0 1.5px var(--garis-kuat); background: var(--kertas); height: 52px; overflow: hidden; }\n.langkah button { width: 52px; height: 52px; display: grid; place-items: center; color: var(--biru); flex: none; }\n.langkah button:active { background: var(--lantai); }\n.langkah output { flex: 1; text-align: center; font: 700 22px/1 var(--huruf-rapat); }\n\n/* ---------- kartu, daftar ---------- */\n.card { background: var(--kertas); border: 1px solid var(--garis); border-radius: var(--r-kartu); padding: 16px; }\n.card__judul { font: 700 17px/1.25 var(--huruf); }\n.card__ket { font-size: 13.5px; color: var(--abu); margin-top: 3px; line-height: 1.4; }\n.list { background: var(--kertas); border: 1px solid var(--garis); border-radius: var(--r-kartu); overflow: hidden; }\n.row { display: flex; align-items: center; gap: 12px; width: 100%; padding: 12px 16px; min-height: 58px; text-align: left; }\n.row + .row { border-top: 1px solid var(--garis); }\nbutton.row:active, label.row:active { background: var(--lantai); }\n.row__isi { flex: 1; min-width: 0; display: block; }\n.row__t, .row__s { display: block; }\n.row__t { font: 600 15px/1.3 var(--huruf); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }\n.row__t--lipat { white-space: normal; }\n.row__s { font-size: 13px; color: var(--abu); margin-top: 2px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }\n.row__s--lipat { white-space: normal; }\n.row__ekor { flex: none; display: flex; align-items: center; gap: 8px; color: var(--abu-2); }\n.row__angka { font: 700 20px/1 var(--huruf-rapat); color: var(--tinta); }\n.list__grup { display: flex; align-items: baseline; justify-content: space-between; gap: 10px; padding: 9px 16px 8px; background: var(--lantai); font: 600 13px/1.3 var(--huruf); color: var(--abu); }\n.list__grup::first-letter { text-transform: uppercase; }\n.list__grup b { font-weight: 600; font-variant-numeric: tabular-nums; }\n.row + .list__grup, .list__grup + .row { border-top: 1px solid var(--garis); }\n\n/* ---------- pil, cip, lencana ---------- */\n.pill { display: inline-flex; align-items: center; gap: 5px; height: 24px; padding: 0 9px; border-radius: 999px; font: 600 12.5px/1 var(--huruf); background: var(--lantai-2); color: var(--abu); white-space: nowrap; vertical-align: middle; }\n.pill--biru { background: var(--biru-muda); color: var(--biru); }\n.pill--jingga { background: var(--jingga-muda); color: var(--jingga-teks); }\n.pill--ungu { background: var(--ungu-muda); color: var(--ungu-teks); }\n.pill--hijau { background: var(--hijau-muda); color: var(--hijau-teks); }\n.pill--merah { background: var(--merah-muda); color: var(--merah-teks); }\n.pill--kuning { background: var(--kuning); color: var(--tinta-label); }\n.chips { display: flex; gap: 8px; overflow-x: auto; margin: 0 -16px; padding: 2px 16px; scrollbar-width: none; }\n.chips::-webkit-scrollbar { display: none; }\n.chip { display: inline-flex; align-items: center; gap: 6px; height: 36px; padding: 0 13px; border-radius: 999px; box-shadow: inset 0 0 0 1px var(--garis-kuat); background: var(--kertas); font: 500 14px/1 var(--huruf); white-space: nowrap; flex: none; }\n.chip b { font-weight: 600; color: var(--abu); }\n.chip.is-on { background: var(--tinta); box-shadow: none; color: var(--kertas); }\n.chip.is-on b { color: inherit; opacity: 0.7; }\n.seg { display: flex; background: var(--lantai-2); border-radius: 13px; padding: 3px; gap: 2px; }\n.seg button { flex: 1; min-width: 0; min-height: 38px; padding: 0 8px; border-radius: 10px; font: 600 14px/1.15 var(--huruf); color: var(--abu); text-align: center; }\n.seg button.is-on { background: var(--seg-aktif); color: var(--tinta); box-shadow: 0 1px 3px rgba(20, 24, 40, 0.14); }\n.seg--kecil { border-radius: 11px; padding: 2px; }\n.seg--kecil button { min-height: 34px; border-radius: 9px; font-size: 13px; }\n\n/* ---------- bilah progres ---------- */\n.bar { height: 12px; border-radius: 6px; background: var(--lantai-2); overflow: hidden; }\n.bar > i { display: block; height: 100%; width: 0; border-radius: 6px; background: var(--hijau); transition: width 0.5s var(--luwes); }\n.bar--tipis { height: 6px; border-radius: 3px; }\n.bar--biru > i { background: var(--biru); }\n.bar--ungu > i { background: var(--ungu); }\n\n/* ---------- catatan berwarna ---------- */\n.info { display: flex; gap: 11px; align-items: flex-start; padding: 12px 14px; border-radius: 14px; background: var(--lantai); font-size: 14px; line-height: 1.4; }\n.info .ic { margin-top: 1px; color: var(--abu); }\n.info b { font-weight: 600; }\n.info--jingga { background: var(--jingga-muda); } .info--jingga .ic { color: var(--jingga-teks); }\n.info--merah { background: var(--merah-muda); } .info--merah .ic { color: var(--merah-teks); }\n.info--hijau { background: var(--hijau-muda); } .info--hijau .ic { color: var(--hijau-teks); }\n.info--biru { background: var(--biru-muda); } .info--biru .ic { color: var(--biru); }\nbutton.info { width: 100%; text-align: left; }\nbutton.info:active { opacity: 0.75; }\n\n/* ---------- kosong & kerangka ---------- */\n.kosong { text-align: center; padding: 36px 20px; }\n.kosong__ic { width: 60px; height: 60px; margin: 0 auto 14px; border-radius: 19px; display: grid; place-items: center; background: var(--lantai-2); color: var(--abu); }\n.kosong__t { font: 700 18px/1.3 var(--huruf); }\n.kosong__s { color: var(--abu); margin: 6px auto 0; font-size: 14.5px; max-width: 30ch; }\n.kosong .btn { margin-top: 18px; }\n.kerangka { border-radius: 10px; background: linear-gradient(90deg, var(--lantai-2) 25%, var(--garis) 37%, var(--lantai-2) 63%); background-size: 400% 100%; animation: kilau 1.3s ease infinite; }\n@keyframes kilau { from { background-position: 100% 50%; } to { background-position: 0 50%; } }\n.putar { width: 20px; height: 20px; border-radius: 50%; border: 2.5px solid currentColor; border-right-color: transparent; animation: putar 0.7s linear infinite; flex: none; }\n.muat { display: flex; align-items: center; justify-content: center; gap: 10px; padding: 22px; color: var(--abu); font-size: 14px; }\n\n/* ---------- label rak ---------- */\n.rak { display: inline-flex; align-items: center; height: 28px; padding: 0 9px; border-radius: 8px; background: var(--kuning); color: var(--tinta-label); box-shadow: inset 0 0 0 1px var(--kuning-tepi); font: 700 18px/1 var(--huruf-rapat); letter-spacing: 0.02em; white-space: nowrap; flex: none; }\n.rak-besar { display: flex; align-items: stretch; border-radius: 18px; background: var(--kuning); color: var(--tinta-label); box-shadow: inset 0 0 0 1.5px var(--kuning-tepi); padding: 14px 10px 12px; }\n.rak-besar__bagian { flex: 1; min-width: 0; text-align: center; position: relative; }\n.rak-besar__bagian + .rak-besar__bagian::before { content: ''; position: absolute; left: 0; top: 6px; bottom: 26px; width: 2px; border-radius: 1px; background: rgba(23, 21, 10, 0.22); }\n.rak-besar__nilai { display: block; font: 700 58px/0.95 var(--huruf-rapat); letter-spacing: 0.01em; }\n.rak-besar__nama { display: block; font: 600 12.5px/1 var(--huruf); margin-top: 7px; opacity: 0.72; }\n.rak-besar--utuh .rak-besar__nilai { font-size: 46px; }\n.tingkat { display: grid; grid-template-rows: repeat(6, 1fr); gap: 3px; width: 16px; margin: 4px 4px 22px 2px; flex: none; }\n.tingkat i { border-radius: 2px; background: rgba(23, 21, 10, 0.16); }\n.tingkat i.is-on { background: var(--tinta-label); }\n\n/* ---------- navigasi bawah ---------- */\n.nav {\n  position: fixed; z-index: 30; left: 50%; bottom: calc(var(--sab) + 10px); transform: translateX(-50%);\n  width: min(500px, calc(100% - 20px)); display: flex; gap: 2px; padding: 6px; border-radius: 28px;\n  background: var(--kertas); box-shadow: var(--bayang-melayang), inset 0 0 0 1px var(--garis);\n  transition: transform 0.22s var(--luwes), opacity 0.18s;\n}\nbody.tanpa-nav .nav { transform: translate(-50%, 140%); opacity: 0; pointer-events: none; }\n.nav__item { flex: 1; min-width: 0; display: flex; flex-direction: column; align-items: center; gap: 3px; padding: 8px 2px 7px; border-radius: 22px; color: var(--abu); font: 600 11.5px/1.1 var(--huruf); position: relative; transition: background-color 0.15s, color 0.15s; }\n.nav__item.is-on { background: var(--biru-muda); color: var(--biru); }\n.nav__item .ic { width: 23px; height: 23px; }\n.nav__lencana { position: absolute; top: 3px; left: calc(50% + 5px); min-width: 18px; height: 18px; padding: 0 5px; border-radius: 9px; background: var(--jingga); color: var(--di-jingga); font: 700 11px/18px var(--huruf); text-align: center; box-shadow: 0 0 0 2px var(--kertas); }\n.nav__lencana--ungu { background: var(--ungu); color: var(--di-warna); }\n.nav__lencana--merah { background: var(--merah); color: var(--di-warna); }\n\n/* ---------- lembar bawah ---------- */\n.lembar-wadah { position: fixed; inset: 0; z-index: 50; display: flex; align-items: flex-end; justify-content: center; }\n.lembar-tabir { position: absolute; inset: 0; background: rgba(13, 16, 26, 0.46); opacity: 0; transition: opacity 0.2s; }\n.lembar { position: relative; width: min(520px, 100%); max-height: calc(100vh - var(--sat) - 28px); max-height: calc(100dvh - var(--sat) - 28px); display: flex; flex-direction: column; background: var(--kertas); border-radius: var(--r-lembar) var(--r-lembar) 0 0; box-shadow: var(--bayang-lembar); transform: translateY(102%); transition: transform 0.26s var(--luwes); }\n.lembar-wadah.is-on .lembar { transform: none; }\n.lembar-wadah.is-on .lembar-tabir { opacity: 1; }\n.lembar__pegangan { width: 40px; height: 4px; border-radius: 2px; background: var(--garis-kuat); margin: 9px auto 0; flex: none; }\n.lembar__kepala { display: flex; align-items: center; gap: 8px; padding: 10px 12px 8px 20px; flex: none; }\n.lembar__judul { flex: 1; min-width: 0; font: 700 19px/1.25 var(--huruf); }\n.lembar__isi { overflow-y: auto; padding: 4px 20px calc(var(--sab) + 22px); -webkit-overflow-scrolling: touch; overscroll-behavior: contain; }\n.lembar__kaki { flex: none; padding: 12px 20px calc(var(--sab) + 14px); border-top: 1px solid var(--garis); }\n.lembar--penuh { height: calc(100vh - var(--sat) - 28px); height: calc(100dvh - var(--sat) - 28px); }\n\n/* ---------- toast ---------- */\n.toast {\n  position: fixed; z-index: 70; left: 50%; bottom: calc(var(--sab) + 100px); transform: translate(-50%, 14px); opacity: 0; pointer-events: none;\n  width: max-content; max-width: min(480px, calc(100% - 24px)); display: flex; align-items: center; gap: 10px; padding: 11px 12px 11px 16px;\n  border-radius: 16px; background: #20242C; color: #fff; font: 500 14.5px/1.35 var(--huruf); box-shadow: var(--bayang-melayang);\n  transition: opacity 0.18s, transform 0.22s var(--luwes);\n}\n.toast.is-on { opacity: 1; transform: translate(-50%, 0); pointer-events: auto; }\n.toast--atas { bottom: auto; top: calc(var(--sat) + 10px); transform: translate(-50%, -14px); }\nbody.tanpa-nav .toast:not(.toast--atas) { bottom: calc(var(--sab) + 22px); }\n.toast__teks { flex: 1; min-width: 0; padding: 2px 4px 2px 0; }\n.toast__aksi { flex: none; color: #8DBBFF; font-weight: 700; padding: 6px 8px; border-radius: 10px; }\n.toast__aksi:active { background: rgba(255, 255, 255, 0.12); }\n.toast--galat { background: #8E1F1F; }\nhtml[data-tema='gelap'] .toast { background: #2C313C; }\nhtml[data-tema='gelap'] .toast--galat { background: #8E1F1F; }\n\n/* ---------- pita atas (offline, demo) ---------- */\n.pita { position: fixed; z-index: 60; top: 0; left: 0; right: 0; padding: calc(var(--sat) + 5px) 12px 6px; text-align: center; font: 600 12.5px/1.3 var(--huruf); color: #fff; background: #20242C; }\n.pita--jingga { background: var(--jingga-teks); }\n.pita button { color: #fff; text-decoration: underline; font-weight: 700; margin-left: 6px; }\n\n/* ---------- tabir sibuk (proses panjang) ---------- */\n.sibuk { position: fixed; inset: 0; z-index: 80; display: grid; place-items: center; background: rgba(13, 16, 26, 0.5); }\n.sibuk__kotak { width: min(300px, 82%); padding: 26px 24px 22px; border-radius: 24px; background: var(--kertas); text-align: center; }\n.sibuk__cincin { width: 84px; height: 84px; margin: 0 auto 14px; position: relative; }\n.sibuk__cincin svg { width: 84px; height: 84px; transform: rotate(-90deg); }\n.sibuk__cincin circle { fill: none; stroke-width: 7; }\n.sibuk__cincin .jalur { stroke: var(--lantai-2); }\n.sibuk__cincin .isi { stroke: var(--biru); stroke-linecap: round; transition: stroke-dashoffset 0.4s var(--luwes); }\n.sibuk__persen { position: absolute; inset: 0; display: grid; place-items: center; font: 700 22px/1 var(--huruf-rapat); }\n.sibuk__teks { font: 600 15px/1.35 var(--huruf); }\n.sibuk__ket { font-size: 13px; color: var(--abu); margin-top: 4px; }\n\n/* ---------- grafik ---------- */\n.grafik { width: 100%; display: block; overflow: visible; }\n.grafik text { font-family: var(--huruf); fill: var(--abu); }\n.legenda { display: flex; flex-wrap: wrap; gap: 6px 14px; font-size: 12.5px; color: var(--abu); }\n.legenda i { display: inline-block; width: 10px; height: 10px; border-radius: 3px; margin-right: 5px; vertical-align: -1px; }\n\n@media (prefers-reduced-motion: reduce) {\n  *, *::before, *::after { animation-duration: 0.01ms !important; animation-iteration-count: 1 !important; transition-duration: 0.01ms !important; }\n}\n\n/* Gaya per layar. */\n\n/* ---------- masuk ---------- */\n#scr-login.is-on { display: flex; min-height: calc(100vh - var(--sat) - var(--sab) - 40px); min-height: calc(100dvh - var(--sat) - var(--sab) - 40px); }\n.lg { display: flex; flex-direction: column; width: 100%; padding: 4px 6px 0; }\n.lg__gambar { margin: 6px -6px 22px; }\n.lg__rak { display: block; width: 100%; height: auto; max-height: 30vh; }\n@media (max-height: 700px) { .lg__rak { max-height: 23vh; } .lg__gambar { margin-bottom: 14px; } .lg form { margin-top: 16px; } }\n.lg__label { transform-box: fill-box; transform-origin: center; animation: tempel 0.5s var(--luwes) 0.15s both; }\n.lg__centang { transform-box: fill-box; transform-origin: center; animation: tempel 0.4s var(--luwes) 0.5s both; }\n@keyframes tempel { from { opacity: 0; transform: scale(0.7); } to { opacity: 1; transform: none; } }\n.lg__judul { font: 700 32px/1.1 var(--huruf); letter-spacing: -0.015em; }\n.lg__ket { color: var(--abu); margin-top: 6px; font-size: 15.5px; }\n.lg form { margin-top: 22px; }\n.lg__tombol { margin-top: 14px; }\n.lg__status { display: flex; align-items: center; gap: 8px; margin-top: 16px; font-size: 13.5px; color: var(--abu); min-height: 20px; }\n.lg__titik { width: 9px; height: 9px; border-radius: 50%; background: var(--abu-2); flex: none; }\n.lg__titik--siap { background: var(--hijau); }\n.lg__titik--awas { background: var(--jingga); }\n.lg__titik--demo { background: var(--ungu); }\n.lg__titik--cek { animation: denyut 1s ease-in-out infinite; }\n@keyframes denyut { 50% { opacity: 0.25; } }\n.lg__info { margin-top: 14px; }\n.lg__demo { margin-top: 16px; padding-top: 16px; border-top: 1px solid var(--garis); }\n.lg__demo-judul { font: 600 13.5px/1.3 var(--huruf); margin-bottom: 8px; }\n.lg__demo-pilih { display: flex; gap: 8px; flex-wrap: wrap; }\n.lg__kaki { margin-top: auto; padding-top: 22px; display: flex; align-items: center; justify-content: space-between; font-size: 13px; color: var(--abu-2); }\n.lg__kaki .tautan { font-size: 13px; }\n\n/* ---------- home ---------- */\n.hm-kepala { display: flex; align-items: flex-start; gap: 12px; margin-bottom: 16px; }\n.hm-avatar { width: 52px; height: 52px; border-radius: 16px; background: var(--biru); color: var(--di-warna); display: grid; place-items: center; font: 700 24px/1 var(--huruf-rapat); flex: none; }\nbutton.hm-avatar:active { background: var(--biru-tua); }\n.hm-siapa { flex: 1; min-width: 0; }\n.hm-sapa { font-size: 13.5px; color: var(--abu); line-height: 1.2; }\n.hm-nama { font: 700 20px/1.2 var(--huruf); letter-spacing: -0.005em; overflow-wrap: anywhere; }\n.hm-tempat { font-size: 13.5px; color: var(--abu); margin: 1px 0 6px; }\n.hm-segar { display: inline-flex; align-items: center; gap: 6px; height: 36px; padding: 0 12px 0 10px; border-radius: 18px; background: var(--kertas); box-shadow: inset 0 0 0 1px var(--garis); font: 600 13px/1 var(--huruf); color: var(--abu); flex: none; font-variant-numeric: tabular-nums; }\n.hm-segar:active { background: var(--lantai-2); }\n.is-putar .ic { animation: putar 0.8s linear infinite; }\n\n.lanjut { display: flex; align-items: center; gap: 12px; width: 100%; padding: 13px 12px 13px 13px; border-radius: 18px; background: var(--c-muda); margin-bottom: 12px; text-align: left; color: var(--tinta); }\nbutton.lanjut:active { opacity: 0.8; }\n.lanjut > .ic { color: var(--c-teks); }\n.lanjut__ic { width: 40px; height: 40px; border-radius: var(--r-ubin); background: var(--c); color: var(--di-warna); display: grid; place-items: center; flex: none; }\n.lanjut__teks { flex: 1; min-width: 0; display: flex; flex-direction: column; }\n.lanjut__teks b { font: 600 15.5px/1.3 var(--huruf); }\n.lanjut__teks span { font-size: 13.5px; color: var(--abu); line-height: 1.35; }\n.lanjut--biru, .kpi--biru { --c: var(--biru); --c-teks: var(--biru); --c-muda: var(--biru-muda); }\n.lanjut--jingga, .kpi--jingga { --c: var(--jingga); --c-teks: var(--jingga-teks); --c-muda: var(--jingga-muda); }\n.lanjut--ungu, .kpi--ungu { --c: var(--ungu); --c-teks: var(--ungu-teks); --c-muda: var(--ungu-muda); }\n.lanjut--hijau, .kpi--hijau { --c: var(--hijau); --c-teks: var(--hijau-teks); --c-muda: var(--hijau-muda); }\n.lanjut--merah, .kpi--merah { --c: var(--merah); --c-teks: var(--merah-teks); --c-muda: var(--merah-muda); }\n\n.kpi-kisi { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }\n.kpi { display: flex; flex-direction: column; align-items: flex-start; padding: 13px 14px 14px; border-radius: var(--r-kartu); background: var(--kertas); border: 1px solid var(--garis); text-align: left; min-width: 0; }\n.kpi:active { background: var(--lantai); }\n.kpi__atas { display: flex; align-items: center; gap: 9px; min-height: 36px; }\n.kpi__ic { width: 36px; height: 36px; border-radius: 11px; background: var(--c); color: var(--di-warna); display: grid; place-items: center; flex: none; }\n.kpi__judul { font: 600 14px/1.15 var(--huruf); color: var(--c-teks); }\n.kpi__nilai { font: 700 40px/1 var(--huruf-rapat); margin-top: 10px; letter-spacing: 0.005em; }\n.kpi__ket { font-size: 13px; line-height: 1.3; color: var(--abu); margin-top: 3px; }\n\n.hm-kartu { margin-top: 12px; }\n.hm-sisa { display: flex; align-items: baseline; gap: 8px; margin: 14px 0 10px; font-size: 15px; font-weight: 500; }\n.hm-sisa b { font: 700 44px/0.9 var(--huruf-rapat); color: var(--jingga-teks); }\n.hm-sisa--beres { color: var(--hijau-teks); font-weight: 600; }\n.hm-dua { display: flex; justify-content: space-between; gap: 12px; margin-top: 10px; }\n.hm-dua > div { display: flex; flex-direction: column; }\n.hm-dua > div + div { text-align: right; align-items: flex-end; }\n.hm-dua b { font: 700 28px/1.1 var(--huruf-rapat); }\n.hm-dua span { font-size: 13px; color: var(--abu); }\n.hm-dua .hm-persen { font: 600 13.5px/1.3 var(--huruf); }\n.hm-catatan { margin-top: 14px; align-items: center; }\n.hm-catatan span { flex: 1; }\n.hm-tugas .btn { margin-top: 16px; }\n.hm-dari { display: flex; justify-content: space-between; gap: 10px; margin-top: 8px; font-size: 13.5px; color: var(--abu); }\n.hm-dari b { color: var(--tinta); font-weight: 600; }\n\n.pm-kisi { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 14px; }\n.pm { display: flex; flex-direction: column; align-items: flex-start; gap: 2px; padding: 12px 14px; border-radius: 16px; background: var(--lantai); text-align: left; }\n.pm:not([disabled]):active { background: var(--lantai-2); }\n.pm__nama { font: 600 13.5px/1.2 var(--huruf); color: var(--abu); }\n.pm__nilai { display: flex; align-items: baseline; gap: 5px; }\n.pm__nilai .ic { align-self: center; width: 20px; height: 20px; stroke-width: 2.4; }\n.pm__nilai b { font: 700 32px/1.05 var(--huruf-rapat); }\n.pm__nilai small { font-size: 13px; font-weight: 600; }\n.pm--plus .pm__nilai { color: var(--hijau-teks); }\n.pm--minus .pm__nilai { color: var(--merah-teks); }\n.pm[disabled] .pm__nilai { color: var(--abu-2); }\n.pm__sku { font-size: 13px; color: var(--abu); }\n.pm-lokasi { display: flex; flex-wrap: wrap; gap: 5px; align-items: center; margin-top: 7px; font-size: 12.5px; }\n.pm-lokasi .rak { height: 22px; font-size: 15px; padding: 0 7px; border-radius: 6px; }\n\n.orang { padding: 11px 0; }\n.orang + .orang { border-top: 1px solid var(--garis); }\n.orang__atas { display: flex; justify-content: space-between; align-items: center; gap: 10px; margin-bottom: 7px; }\n.orang__atas > b { font-weight: 600; overflow-wrap: anywhere; display: flex; flex-wrap: wrap; align-items: center; gap: 4px 8px; min-width: 0; }\n.orang__atas > span { font: 700 22px/1 var(--huruf-rapat); flex: none; }\n.orang__ket { font-size: 12.5px; color: var(--abu); margin-top: 6px; }\n.bar--jingga > i { background: var(--jingga); }\n\n/* ---------- daftar tugas (Cycle, Validasi) ---------- */\n.tg-mulai { margin-bottom: 14px; }\n.tg-cip { margin-bottom: 8px; }\n.tg-cari { margin: 6px 0 12px; }\n\n/* ---------- mode fokus ---------- */\n/* Satu lokasi per layar. Semuanya harus muat tanpa menggulir, termasuk di handheld gudang 360x640. */\nbody.mode-fokus #app { padding-top: calc(var(--sat) + 10px); padding-bottom: calc(var(--sab) + 12px); }\nbody.mode-fokus.ada-pita #app { padding-top: calc(var(--sat) + 10px + var(--pita-t, 28px)); }\n#scr-hitung.is-on { display: flex; flex-direction: column; min-height: calc(100vh - var(--sat) - var(--sab) - 22px); min-height: calc(100dvh - var(--sat) - var(--sab) - 22px); }\nbody.ada-pita #scr-hitung.is-on { min-height: calc(100vh - var(--sat) - var(--sab) - 22px - var(--pita-t, 28px)); min-height: calc(100dvh - var(--sat) - var(--sab) - 22px - var(--pita-t, 28px)); }\n.fk__kepala { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }\n.fk__judul { flex: 1; min-width: 0; display: flex; flex-direction: column; line-height: 1.2; }\n.fk__judul b { font: 700 18px/1.2 var(--huruf); }\n.fk__judul span { font-size: 13.5px; color: var(--abu); font-variant-numeric: tabular-nums; }\n.fk__antre { display: inline-flex; align-items: center; gap: 6px; height: 36px; padding: 0 12px; border-radius: 18px; background: var(--kertas); box-shadow: inset 0 0 0 1px var(--garis); font: 600 13.5px/1 var(--huruf); color: var(--abu); font-variant-numeric: tabular-nums; }\n.fk__antre.is-tunggu { background: var(--jingga-muda); color: var(--jingga-teks); box-shadow: none; }\n.fk__kartu { margin-top: 12px; }\n.fk__kartu.is-masuk { animation: geser 0.2s var(--luwes); }\n@keyframes geser { from { opacity: 0; transform: translateX(22px); } to { opacity: 1; transform: none; } }\n.fk__barang { display: flex; align-items: flex-start; gap: 10px; padding: 11px 2px 0; }\n.fk__teks { flex: 1; min-width: 0; }\n.fk__sku { font: 700 20px/1.2 var(--huruf); letter-spacing: 0.01em; }\n.fk__nama { color: var(--abu); font-size: 14.5px; line-height: 1.35; margin-top: 1px; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }\n.fk__lewati { flex: none; display: inline-flex; align-items: center; gap: 5px; height: 36px; padding: 0 13px 0 10px; border-radius: 18px; background: var(--kertas); box-shadow: inset 0 0 0 1px var(--garis-kuat); font: 600 13.5px/1 var(--huruf); color: var(--abu); }\n.fk__lewati:active { background: var(--lantai-2); }\n.fk__cip { display: flex; flex-wrap: wrap; gap: 6px; margin: 9px 2px 0; }\n.fk__cip .pill { height: 26px; gap: 4px; padding: 0 10px 0 7px; }\n.fk__cip.is-baru .pill { animation: sorot 1.4s ease-out; }\n@keyframes sorot { 0%, 45% { background: var(--kuning); color: var(--tinta-label); } }\n.fk__terakhir { margin-top: auto; padding-top: 10px; min-height: 56px; }\n.fk__terakhir--atas { margin-top: 0; min-height: 0; }\n.simpanan { position: relative; display: flex; align-items: center; gap: 10px; padding: 6px 6px 6px 12px; min-height: 46px; border-radius: 14px; background: var(--hijau-muda); overflow: hidden; animation: muncul 0.18s ease-out; }\n.simpanan__ic { color: var(--hijau-teks); display: grid; }\n.simpanan__teks { flex: 1; min-width: 0; font-size: 14px; line-height: 1.3; }\n.simpanan__teks b { font-weight: 600; }\n.simpanan__urung { flex: none; height: 34px; padding: 0 12px; border-radius: 10px; background: var(--kertas); font: 600 14px/1 var(--huruf); color: var(--tinta); }\n.simpanan__urung:active { opacity: 0.7; }\n.simpanan__waktu { position: absolute; left: 0; bottom: 0; height: 3px; width: 100%; background: var(--hijau); transform-origin: left; animation: habis linear forwards; }\n@keyframes habis { to { transform: scaleX(0); } }\n.fk__qty { padding: 10px 0 12px; text-align: center; display: flex; flex-direction: column; align-items: center; }\n.fk__angka { font: 700 60px/0.95 var(--huruf-rapat); letter-spacing: 0.01em; min-width: 2ch; }\n.fk__angka.is-kosong { color: var(--garis-kuat); }\n.fk__satuan { font-size: 13.5px; color: var(--abu); margin-top: 6px; }\n.pad { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; }\n.pad button { height: 58px; border-radius: 16px; background: var(--kertas); box-shadow: inset 0 0 0 1px var(--garis); font: 600 28px/1 var(--huruf-rapat); display: grid; place-items: center; text-align: center; touch-action: manipulation; user-select: none; -webkit-user-select: none; transition: transform 0.06s, background-color 0.1s; }\n.pad button:active { background: var(--lantai-2); transform: scale(0.97); }\n.pad .pad__simpan { background: var(--biru); color: var(--di-warna); box-shadow: none; font: 700 17px/1 var(--huruf); }\n.pad .pad__simpan:active { background: var(--biru-tua); }\n.pad .pad__simpan[disabled] { background: var(--lantai-2); color: var(--abu-2); }\n.fk--ungu .pad .pad__simpan:not([disabled]) { background: var(--ungu); }\n.fk__tuntas { margin: auto 0; text-align: center; padding: 20px 8px; }\n.fk__tuntas-ic { width: 72px; height: 72px; border-radius: 24px; margin: 0 auto 18px; display: grid; place-items: center; background: var(--hijau-muda); color: var(--hijau-teks); }\n.fk__tuntas h2 { font: 700 24px/1.2 var(--huruf); }\n.fk__tuntas p { color: var(--abu); margin: 8px auto 22px; max-width: 32ch; }\n.fk__tuntas .btn + .btn { margin-top: 10px; }\n/* Kerapatan: lihat ukurLayar() di 20-ui.js. Tiap tingkat menambah aturan di atas tingkat sebelumnya. */\nhtml[data-rapat~='1'] .rak-besar { padding: 10px 8px 9px; }\nhtml[data-rapat~='1'] .rak-besar__nilai { font-size: 48px; }\nhtml[data-rapat~='1'] .tingkat { margin-bottom: 20px; }\nhtml[data-rapat~='1'] .fk__angka { font-size: 48px; }\nhtml[data-rapat~='1'] .fk__qty { padding: 6px 0 9px; }\nhtml[data-rapat~='1'] .pad button { height: 52px; }\n\nhtml[data-rapat~='2'] .fk__kepala { margin-bottom: 8px; }\nhtml[data-rapat~='2'] .fk__kartu { margin-top: 10px; }\nhtml[data-rapat~='2'] .rak-besar { padding: 8px; border-radius: 16px; }\nhtml[data-rapat~='2'] .rak-besar__nilai { font-size: 42px; }\nhtml[data-rapat~='2'] .rak-besar__nama { margin-top: 5px; font-size: 12px; }\nhtml[data-rapat~='2'] .tingkat { margin: 2px 4px 18px 2px; gap: 2px; }\nhtml[data-rapat~='2'] .fk__barang { padding-top: 9px; }\nhtml[data-rapat~='2'] .fk__nama { -webkit-line-clamp: 1; }\nhtml[data-rapat~='2'] .fk__cip { margin-top: 7px; }\nhtml[data-rapat~='2'] .fk__terakhir { min-height: 52px; padding-top: 6px; }\nhtml[data-rapat~='2'] .fk__angka { font-size: 40px; }\nhtml[data-rapat~='2'] .fk__satuan { margin-top: 3px; font-size: 12.5px; }\nhtml[data-rapat~='2'] .fk__qty { padding: 4px 0 7px; }\nhtml[data-rapat~='2'] .pad { gap: 6px; }\nhtml[data-rapat~='2'] .pad button { height: 46px; font-size: 24px; border-radius: 13px; }\n\nhtml[data-rapat~='3'] .fk__angka { font-size: 34px; }\nhtml[data-rapat~='3'] .pad button { height: 43px; }\nhtml[data-rapat~='3'] .fk__terakhir { min-height: 50px; padding-top: 4px; }\n\n/* Paling sempit (handheld 360x640 dengan pita \"tidak ada internet\"): keterangan area & alat disembunyikan\n   supaya keypad tetap utuh; lokasi, SKU, dan nama barang tetap tampil. */\nhtml[data-rapat~='4'] .fk__cip:not(.fk__cip--sama) { display: none; }\n\n/* ---------- akun & pembaruan ---------- */\n.ak__siapa { display: flex; align-items: center; gap: 12px; margin-bottom: 16px; }\n.ak__nama { font: 700 18px/1.25 var(--huruf); overflow-wrap: anywhere; }\n.upd__catatan { margin: 0 0 6px; padding-left: 18px; color: var(--abu); font-size: 14.5px; line-height: 1.5; }\n\n/* ---------- verifikasi ---------- */\n.vf-seg, .rp-seg { margin-bottom: 12px; }\n.vf-row { align-items: flex-start; padding-top: 13px; padding-bottom: 13px; }\n.vf-row__atas { display: flex; align-items: center; gap: 8px; margin-bottom: 4px; min-width: 0; }\n.vf-row__atas .row__t { flex: 1; min-width: 0; }\n.vf-row__status { display: flex; align-items: center; flex-wrap: wrap; gap: 7px; margin-top: 7px; font-size: 13px; }\n.tahap { display: inline-flex; gap: 3px; }\n.tahap i { width: 14px; height: 5px; border-radius: 3px; background: var(--lantai-2); }\n.tahap i.is-on { background: var(--biru); }\n.vf-selisih { height: 26px; padding: 0 9px; font: 700 16px/1 var(--huruf-rapat); letter-spacing: 0.01em; }\n.vf-pasang { display: block; width: 100%; text-align: left; }\nbutton.vf-pasang:active { background: var(--lantai); }\n.vf-pasang__rute { display: flex; align-items: center; gap: 10px; margin-top: 11px; color: var(--abu-2); }\n.vf-pasang__rute > span { display: inline-flex; align-items: center; gap: 7px; }\n.vf-pasang__rute b { font: 700 18px/1 var(--huruf-rapat); }\n.vf-pasang__rute--besar { margin: 14px 0 6px; }\n.vf-kepala { display: flex; align-items: center; flex-wrap: wrap; gap: 8px; margin-bottom: 10px; }\n.vf-sku { font: 700 20px/1.2 var(--huruf); }\n.rinci { margin: 14px 0 6px; border-top: 1px solid var(--garis); }\n.rinci__baris { display: flex; justify-content: space-between; gap: 14px; padding: 9px 0; border-bottom: 1px solid var(--garis); font-size: 14px; }\n.rinci__baris span { color: var(--abu); flex: none; }\n.rinci__baris b { font-weight: 600; text-align: right; overflow-wrap: anywhere; }\n.rinci--rapat { margin: 8px 0 0; border-top: 0; }\n.rinci--rapat .rinci__baris { padding: 4px 0; border-bottom: 0; font-size: 13.5px; }\n.vf-catatan { white-space: pre-line; font-size: 14px; line-height: 1.45; background: var(--lantai); border-radius: 12px; padding: 11px 13px; margin-top: 8px; overflow-wrap: anywhere; }\n.rentang { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 12px; }\n.rentang label { min-width: 0; }\n.rentang label span { display: block; font: 600 13px/1.3 var(--huruf); color: var(--abu); margin-bottom: 5px; }\n\n/* ---------- upload ---------- */\n.up-berkas { display: flex; align-items: center; gap: 12px; padding: 13px 14px; border-radius: var(--r-kartu); background: var(--kertas); border: 1.5px dashed var(--garis-kuat); cursor: pointer; }\n.up-berkas:active { background: var(--lantai); }\n.up-berkas.is-ok { border-style: solid; border-color: var(--hijau); }\n.up-berkas.is-galat { border-style: solid; border-color: var(--merah); }\n.up-berkas:focus-within { outline: 2.5px solid var(--biru); outline-offset: 2px; }\n.up-berkas__ic { width: 40px; height: 40px; border-radius: var(--r-ubin); display: grid; place-items: center; background: var(--lantai-2); color: var(--abu); flex: none; }\n.up-berkas.is-ok .up-berkas__ic { background: var(--hijau-muda); color: var(--hijau-teks); }\n.up-berkas.is-galat .up-berkas__ic { background: var(--merah-muda); color: var(--merah-teks); }\n.up-berkas__isi { flex: 1; min-width: 0; display: flex; flex-direction: column; }\n.up-berkas__judul { font: 600 15px/1.3 var(--huruf); }\n.up-berkas__ket { font-size: 13px; line-height: 1.35; color: var(--abu); overflow-wrap: anywhere; }\n.up-berkas__ket--ok { color: var(--tinta); font-weight: 500; }\n.up-berkas__ket--galat { color: var(--merah-teks); }\n.up-berkas__aksi { flex: none; font: 600 14px/1 var(--huruf); color: var(--biru); }\n.up-peran { margin-bottom: 4px; }\n.up-petugas__baris { display: flex; align-items: center; gap: 16px; }\n.up-petugas__buka { margin-left: auto; }\n.up-daftar { margin: 8px -16px -16px; border-top: 1px solid var(--garis); }\n.up-daftar .row { min-height: 52px; cursor: pointer; }\n.up-alat { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }\n.up-alat .lbl { margin-top: 0; }\n.up-proses { margin-top: 14px; }\n.up-hasil { display: flex; align-items: baseline; gap: 10px; margin: 2px 0 14px; }\n.up-hasil b { font: 700 52px/0.95 var(--huruf-rapat); color: var(--biru); }\n.up-hasil span { color: var(--abu); }\n\n/* ---------- report ---------- */\n.rp-periode { display: flex; align-items: center; gap: 10px; margin-bottom: 12px; }\n.rp-periode .seg { flex: 1; }\n.rp-periode .field { flex: 1; min-width: 0; }\n.rp-akurasi__utama { display: flex; align-items: baseline; gap: 10px; }\n.rp-akurasi__utama b { font: 700 56px/0.95 var(--huruf-rapat); }\n.rp-akurasi__utama span { color: var(--abu); font-size: 14px; }\n.rp-tiga { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin-top: 14px; }\n.rp-tiga > div { display: flex; flex-direction: column; min-width: 0; }\n.rp-tiga b { font: 700 26px/1.1 var(--huruf-rapat); }\n.rp-tiga span { font-size: 12.5px; line-height: 1.3; color: var(--abu); }\n.rp-tiga--garis { border-top: 1px solid var(--garis); padding-top: 12px; }\n.rp-tiga--kartu { margin-top: 0; gap: 10px; }\n.rp-tiga--kartu > div { background: var(--kertas); border: 1px solid var(--garis); border-radius: 16px; padding: 12px 12px 11px; }\n.rp-tiga--dua { grid-template-columns: 1fr 1fr; }\n.rp-info { display: flex; flex-wrap: wrap; gap: 3px 14px; min-height: 20px; margin-top: 6px; font-size: 13.5px; color: var(--abu); }\n.rp-info b { color: var(--tinta); font-weight: 600; }\n.rp-daftar { margin-top: 8px; }\n.rp-daftar > :first-child { border-top: 0; }\n.rp-baris { display: flex; align-items: center; gap: 10px; width: 100%; padding: 12px 0; border-top: 1px solid var(--garis); text-align: left; font-weight: 500; }\n.rp-baris > span:first-child { flex: 1; min-width: 0; }\n.rp-baris .ic { color: var(--abu-2); }\n.rp-rank { display: flex; align-items: center; gap: 12px; width: 100%; padding: 11px 0; border-top: 1px solid var(--garis); text-align: left; }\n.rp-rank__no { width: 26px; height: 26px; border-radius: 9px; background: var(--lantai-2); color: var(--abu); display: grid; place-items: center; font: 700 14px/1 var(--huruf-rapat); flex: none; }\n.rp-rank__isi { flex: 1; min-width: 0; display: flex; flex-direction: column; }\n.rp-rank__isi b { font-weight: 600; overflow-wrap: anywhere; }\n.rp-rank__isi span { font-size: 13px; color: var(--abu); }\n.rp-rank__persen { font: 700 20px/1 var(--huruf-rapat); }\n.rp-batang { display: block; width: 100%; padding: 11px 0; border-top: 1px solid var(--garis); text-align: left; }\n.rp-batang__atas { display: flex; justify-content: space-between; gap: 10px; margin-bottom: 7px; font-size: 14px; }\n.rp-batang__atas b { font-weight: 600; }\n.rp-batang__atas span { color: var(--abu); flex: none; }\n.bar--merah > i { background: var(--merah); }\n.rp-item { padding: 12px 0; border-top: 1px solid var(--garis); }\n.rp-item__bawah { display: flex; justify-content: space-between; align-items: center; gap: 8px; margin-top: 7px; font-size: 13px; }\n.rp-empat { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }\n.rp-empat > div { display: flex; flex-direction: column; padding: 12px; border-radius: 14px; background: var(--lantai); }\n.rp-empat b { font: 700 28px/1.1 var(--huruf-rapat); }\n.rp-empat span { font-size: 13px; color: var(--abu); }\n.rp-donat { display: flex; align-items: center; gap: 16px; margin-top: 12px; }\n.rp-donat svg { flex: none; }\n.rp-daftar--rapat { flex: 1; min-width: 0; margin-top: 0; }\n.rp-akar { display: flex; align-items: center; gap: 8px; width: 100%; padding: 7px 0; text-align: left; font-size: 13.5px; }\n.rp-akar i { width: 10px; height: 10px; border-radius: 3px; flex: none; }\n.rp-akar span { flex: 1; min-width: 0; line-height: 1.25; }\n.rp-akar b { font: 700 17px/1 var(--huruf-rapat); }\n.rp-akar[disabled] { opacity: 0.5; }\n.rp-saran { display: flex; gap: 12px; padding: 11px 0; border-top: 1px solid var(--garis); }\n.rp-saran .ic { color: var(--biru); margin-top: 1px; }\n.rp-saran b { display: block; font-weight: 600; }\n.rp-saran span { font-size: 13.5px; color: var(--abu); }\n\n/* ---------- config ---------- */\n.cf-matriks { overflow-x: auto; margin: 12px -4px 0; }\n.cf-matriks table { width: 100%; border-collapse: collapse; font-size: 13.5px; }\n.cf-matriks th, .cf-matriks td { padding: 8px 4px; text-align: center; border-bottom: 1px solid var(--garis); }\n.cf-matriks thead th { font-weight: 600; line-height: 1.2; vertical-align: bottom; }\n.cf-matriks thead small { display: block; font-weight: 400; font-size: 11.5px; color: var(--abu); }\n.cf-matriks th[scope='row'], .cf-matriks thead th:first-child { text-align: left; font-weight: 600; }\n.cf-matriks tbody tr:last-child th, .cf-matriks tbody tr:last-child td { border-bottom: 0; }\n.cf-cara__baris { padding: 9px 0; border-top: 1px solid var(--garis); }\n.cf-cara__baris:last-child { border-bottom: 1px solid var(--garis); }\n.cf-cara__atas { display: flex; align-items: center; justify-content: space-between; gap: 12px; }\n.cf-cara__atas > span { font: 600 14px/1.25 var(--huruf); }\n.cf-cara__atas small { display: block; font-weight: 400; font-size: 12.5px; color: var(--abu); }\n.cf-cara__atas .seg { width: 176px; flex: none; }\n.cf-cara__baris .ket { margin: 8px 0 2px; }\n.cf-cara__baris .ket b { color: var(--tinta); font-weight: 600; }\n.cf-fac__atas { display: flex; justify-content: space-between; align-items: flex-start; gap: 10px; }\n.cf-fac__aksi { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 14px; }\n.cf-fac__db { display: inline-block; margin-top: 10px; font-size: 14px; }\n\n#app{max-width:640px}</style></head><body><main id=\"app\">\n    <section class=\"scr\" id=\"scr-login\" aria-label=\"Masuk\"></section>\n    <section class=\"scr\" id=\"scr-home\" aria-label=\"Home\"></section>\n    <section class=\"scr\" id=\"scr-cycle\" aria-label=\"Cycle\"></section>\n    <section class=\"scr\" id=\"scr-validasi\" aria-label=\"Validasi\"></section>\n    <section class=\"scr\" id=\"scr-hitung\" aria-label=\"Hitung\"></section>\n    <section class=\"scr\" id=\"scr-verifikasi\" aria-label=\"Verifikasi\"></section>\n    <section class=\"scr\" id=\"scr-upload\" aria-label=\"Upload Data\"></section>\n    <section class=\"scr\" id=\"scr-report\" aria-label=\"Report\"></section>\n    <section class=\"scr\" id=\"scr-config\" aria-label=\"Config\"></section>\n  </main>\n  <nav class=\"nav\" id=\"nav\" aria-label=\"Menu utama\" hidden></nav>\n<script>window.CT_CONFIG={\"version\":\"2.0.0\",\"panel\":true,\"updateRepo\":\"\",\"nativeBase\":\"\",\"build\":0,\"serverUrl\":\"\",\"demoSaja\":false,\"xlsxUrl\":\"https://cdnjs.cloudflare.com/ajax/libs/xlsx/undefined/xlsx.full.min.js\"};</script>\n<script>!function(){\"use strict\";const a=(a,t)=>(t||document).querySelector(a),t=(a,t)=>Array.from((t||document).querySelectorAll(a)),i=a=>document.getElementById(a),s={\"&\":\"&amp;\",\"<\":\"&lt;\",\">\":\"&gt;\",'\"':\"&quot;\",\"'\":\"&#39;\"},e=a=>String(null==a?\"\":a).replace(/[&<>\"']/g,a=>s[a]);class n{constructor(a){this.html=a}}const l=a=>new n(String(null==a?\"\":a)),r=a=>null==a||!1===a||!0===a?\"\":Array.isArray(a)?a.map(r).join(\"\"):a instanceof n?a.html:e(a);function u(a,...t){let i=a[0];for(let s=0;s<t.length;s++)i+=r(t[s])+a[s+1];return l(i)}const o=(a,t)=>(a&&(a.innerHTML=r(t)),a),d={get(a){try{const t=localStorage.getItem(a);return null==t?null:JSON.parse(t)}catch(a){return null}},set(a,t){try{localStorage.setItem(a,JSON.stringify(t))}catch(a){}},del(a){try{localStorage.removeItem(a)}catch(a){}}},c=a=>(Number(a)||0).toLocaleString(\"id-ID\"),p=a=>(Number(a)>0?\"+\":\"\")+c(a);function h(a,t){if(!t)return\"0%\";if(a>=t)return\"100%\";if(a<=0)return\"0%\";const i=a/t*100;return i>=99?String(Math.min(99.9,Math.floor(10*i)/10)).replace(\".\",\",\")+\"%\":i<1?String(Math.max(.1,Math.ceil(10*i)/10)).replace(\".\",\",\")+\"%\":Math.round(i)+\"%\"}const m=a=>String(Math.round(10*(Number(a)||0))/10).replace(\".\",\",\"),k=a=>String(a).padStart(2,\"0\"),g=[\"Januari\",\"Februari\",\"Maret\",\"April\",\"Mei\",\"Juni\",\"Juli\",\"Agustus\",\"September\",\"Oktober\",\"November\",\"Desember\"],b=[\"Jan\",\"Feb\",\"Mar\",\"Apr\",\"Mei\",\"Jun\",\"Jul\",\"Agu\",\"Sep\",\"Okt\",\"Nov\",\"Des\"],f=[\"Minggu\",\"Senin\",\"Selasa\",\"Rabu\",\"Kamis\",\"Jumat\",\"Sabtu\"],v=()=>new Date(Date.now()+252e5).toISOString().slice(0,10),y=(a,t)=>{const i=new Date(a+\"T12:00:00\");return i.setDate(i.getDate()+t),(a=>a.getFullYear()+\"-\"+k(a.getMonth()+1)+\"-\"+k(a.getDate()))(i)},$=a=>{const t=new Date(a);return k(t.getHours())+\".\"+k(t.getMinutes())},_=a=>{const t=String(a||\"\").split(\"-\");return t.length<3?String(a||\"\"):Number(t[2])+\" \"+b[Number(t[1])-1]},w=a=>{const t=new Date(a+\"T12:00:00\");return isNaN(t)?String(a||\"\"):f[t.getDay()]+\", \"+t.getDate()+\" \"+g[t.getMonth()]+\" \"+t.getFullYear()};function S(a){const t=String(null==a?\"\":a).trim();if(!t)return\"\";const i=/^(\\d{4})-(\\d{2})-(\\d{2})/.exec(t);if(i)return Number(i[3])+\" \"+b[Number(i[2])-1]+\" \"+i[1];const s=new Date(t);return isNaN(s)?t:s.getDate()+\" \"+b[s.getMonth()]+\" \"+s.getFullYear()}const x=a=>{const t=String(a||\"\").split(\"-\");return t.length<2?String(a||\"\"):b[Number(t[1])-1]+\" \"+t[0].slice(2)};const M={admin:\"Admin\",administrator:\"Admin\",inventory:\"Inventory\",outbound:\"Outbound\",storing:\"Storing\",inbound:\"Inbound\",lp:\"LP\",maintenance:\"Maintenance\",developer:\"Developer\",cycle:\"Cycle\"},j=a=>M[String(a||\"\").toLowerCase()]||(a?String(a).charAt(0).toUpperCase()+String(a).slice(1):\"\");function T(a){const t=String(a||\"\"),i=(t.indexOf(\".\")>=0?t.slice(t.indexOf(\".\")+1):t).match(/[A-Za-z]/)||t.match(/[A-Za-z0-9]/);return i?i[0].toUpperCase():\"?\"}function A(a){const t=String(a||\"\").trim().toUpperCase(),i=t.split(\".\").filter(a=>\"\"!==a),s=i[i.length-1]||\"\",e=/K/.test(s)&&i.length>1,n=s.match(/\\d+/);return{raw:t,bagian:i,level:!e&&n&&i.length>1?parseInt(n[0],10):null,rollcage:e}}const L={tanpa_alat:{nama:\"Tanpa alat bantu\",ikon:\"jalan\"},tangga:{nama:\"Tangga pesawat\",ikon:\"tangga\"},reach_truck:{nama:\"Reach truck\",ikon:\"forklift\"}};function P(a){const t=String(a||\"\").trim().toUpperCase().split(\".\"),i=(t[0]||\"\").match(/^([A-Z]+)?(\\d+)?/)||[],s=(t[1]||\"\").match(/\\d+/),e=(t[t.length-1]||\"\").match(/\\d+/);return[i[1]||\"\",i[2]?parseInt(i[2],10):999999,s?parseInt(s[0],10):999999,e?parseInt(e[0],10):999999]}function C(a){return u`<span class=\"rak\">${a}</span>`}const D=a=>new Promise(t=>setTimeout(t,a));function E(a,t){let i;return function(){const s=arguments,e=this;clearTimeout(i),i=setTimeout(()=>a.apply(e,s),t)}}const B=(a,t)=>-1!==String(a||\"\").toLowerCase().indexOf(t),N={home:'<path d=\"M3.5 10.5 12 3.5l8.5 7\"/><path d=\"M5.5 9.5V20h4.6v-5.6h3.8V20h4.6V9.5\"/>',kotak:'<path d=\"m12 3 8 4v10l-8 4-8-4V7l8-4z\"/><path d=\"m4 7 8 4 8-4\"/><path d=\"M12 11v10\"/>',perisai:'<path d=\"M12 3.2 19 6v5.2c0 4.3-2.8 8-7 9.6-4.2-1.6-7-5.3-7-9.6V6l7-2.8z\"/><path d=\"m8.8 12 2.3 2.3 4.2-4.4\"/>',selisih:'<circle cx=\"11\" cy=\"11\" r=\"7\"/><path d=\"m20 20-4-4\"/><path d=\"M8.3 11h5.4M11 8.3v5.4\"/>',unggah:'<path d=\"M12 16V4.5\"/><path d=\"m7 9 5-5 5 5\"/><path d=\"M4 15.5V19a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3.5\"/>',grafik:'<path d=\"M4 20h16\"/><path d=\"M7.5 20v-6\"/><path d=\"M12 20V6\"/><path d=\"M16.5 20v-9.5\"/>',segar:'<path d=\"M19.5 11A7.6 7.6 0 0 0 6 7.2L4.5 9\"/><path d=\"M4.5 4.5V9H9\"/><path d=\"M4.5 13A7.6 7.6 0 0 0 18 16.8l1.5-1.8\"/><path d=\"M19.5 19.5V15H15\"/>',kanan:'<path d=\"m9 6 6 6-6 6\"/>',kiri:'<path d=\"m15 6-6 6 6 6\"/>',bawah:'<path d=\"m6 9 6 6 6-6\"/>',atas:'<path d=\"m6 15 6-6 6 6\"/>',tutup:'<path d=\"M6 6l12 12M18 6 6 18\"/>',cek:'<path d=\"m5 12.5 4.5 4.5L19 7.5\"/>',jam:'<circle cx=\"12\" cy=\"12\" r=\"8.5\"/><path d=\"M12 7.5V12l3 2\"/>',awas:'<path d=\"M12 4 2.8 19.5h18.4L12 4z\"/><path d=\"M12 10v4.3\"/><path d=\"M12 17v.2\"/>',info:'<circle cx=\"12\" cy=\"12\" r=\"8.5\"/><path d=\"M12 11v5\"/><path d=\"M12 7.7v.2\"/>',naik:'<path d=\"M12 19V5\"/><path d=\"m6 11 6-6 6 6\"/>',turun:'<path d=\"M12 5v14\"/><path d=\"m6 13 6 6 6-6\"/>',urung:'<path d=\"M8 5 3 10l5 5\"/><path d=\"M3 10h11a6 6 0 0 1 0 12h-3\"/>',hapus_kiri:'<path d=\"M21 5H9l-6 7 6 7h12a1 1 0 0 0 1-1V6a1 1 0 0 0-1-1z\"/><path d=\"m12 9.5 5 5m0-5-5 5\"/>',orang:'<circle cx=\"12\" cy=\"8\" r=\"4\"/><path d=\"M4.5 20c1-3.5 4-5 7.5-5s6.5 1.5 7.5 5\"/>',orang2:'<circle cx=\"9\" cy=\"8\" r=\"3.5\"/><path d=\"M2.5 19.5c.8-3 3.3-4.5 6.5-4.5s5.7 1.5 6.5 4.5\"/><path d=\"M16 4.8a3.5 3.5 0 0 1 0 6.4\"/><path d=\"M18.5 15.3c1.5.6 2.6 1.9 3 4.2\"/>',bulan:'<path d=\"M20 14.5A8.5 8.5 0 0 1 9.5 4a7.2 7.2 0 1 0 10.5 10.5z\"/>',keluar:'<path d=\"M10 4H6a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h4\"/><path d=\"m15 8 4 4-4 4\"/><path d=\"M19 12H9\"/>',atur:'<path d=\"M4 7h9M19 7h1M4 17h3M13 17h7\"/><circle cx=\"16\" cy=\"7\" r=\"2.5\"/><circle cx=\"10\" cy=\"17\" r=\"2.5\"/>',unduh:'<path d=\"M12 4v11.5\"/><path d=\"m7 11 5 5 5-5\"/><path d=\"M4 20h16\"/>',saring:'<path d=\"M4 6h16M7 12h10M10 18h4\"/>',cari:'<circle cx=\"11\" cy=\"11\" r=\"7\"/><path d=\"m20 20-4-4\"/>',berkas:'<path d=\"M7 3h7l5 5v12a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z\"/><path d=\"M14 3v5h5\"/>',tangga:'<path d=\"M8 3v18M16 3v18M8 7.5h8M8 12h8M8 16.5h8\"/>',forklift:'<path d=\"M3 16.5h9.5V8H8.5L5 12.5v4z\"/><path d=\"M15.5 4v12.5H21\"/><circle cx=\"6.5\" cy=\"18.5\" r=\"1.8\"/><circle cx=\"11\" cy=\"18.5\" r=\"1.8\"/>',jalan:'<circle cx=\"12.5\" cy=\"4.8\" r=\"1.8\"/><path d=\"m9 20 2.5-6-2-2.5 1-4 3 2 3 1\"/><path d=\"m13.5 14 2 6\"/><path d=\"m10.5 7.5-3 1.5-1 3\"/>',tambah:'<path d=\"M12 5v14M5 12h14\"/>',kurang:'<path d=\"M5 12h14\"/>',kait:'<path d=\"m9 15 6-6\"/><path d=\"m10.5 6.5 1-1a4.2 4.2 0 0 1 6 6l-1 1\"/><path d=\"m13.5 17.5-1 1a4.2 4.2 0 0 1-6-6l1-1\"/>',gedung:'<path d=\"M5 20V5h9v15\"/><path d=\"M14 10h5v10\"/><path d=\"M3 20h18\"/><path d=\"M8 9h3M8 13h3\"/>',kunci:'<circle cx=\"8\" cy=\"15.5\" r=\"4\"/><path d=\"m11 12.5 8.5-8.5\"/><path d=\"m16 7.5 3 3\"/>',sampah:'<path d=\"M5 7h14\"/><path d=\"M10 7V4h4v3\"/><path d=\"m7 7 1 13h8l1-13\"/>',pensil:'<path d=\"m4 20 4-1L19 8l-3-3L5 16l-1 4z\"/>',pin:'<path d=\"M12 21s7-6.2 7-11a7 7 0 1 0-14 0c0 4.8 7 11 7 11z\"/><circle cx=\"12\" cy=\"10\" r=\"2.5\"/>',kilat:'<path d=\"M13 3 5 13.5h6L10 21l8-10.5h-6L13 3z\"/>',tren:'<path d=\"m4 17.5 5-6 4 3 7-8.5\"/><path d=\"M15.5 6H20v4.5\"/>',awan_mati:'<path d=\"M3 4l18 17\"/><path d=\"M8.5 8.2A5.5 5.5 0 0 1 17.8 10 4.2 4.2 0 0 1 20 17.5\"/><path d=\"M17 19H7.5a4.5 4.5 0 0 1-1.5-8.7\"/>',awan_kirim:'<path d=\"M7.5 18.5a4.5 4.5 0 0 1-.5-9A5.5 5.5 0 0 1 17.8 10a4.3 4.3 0 0 1-.8 8.5\"/><path d=\"M12 20v-8\"/><path d=\"m9 14.5 3-3 3 3\"/>',daftar:'<path d=\"M8 6.5h12M8 12h12M8 17.5h12\"/><path d=\"M4 6.5h.1M4 12h.1M4 17.5h.1\"/>',lewati:'<path d=\"m5 5 8 7-8 7V5z\"/><path d=\"M18 5v14\"/>',mata:'<path d=\"M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z\"/><circle cx=\"12\" cy=\"12\" r=\"3\"/>'};function I(a,t){return l('<svg class=\"ic'+(t?\" \"+t:\"\")+'\" viewBox=\"0 0 24 24\" aria-hidden=\"true\">'+(N[a]||\"\")+\"</svg>\")}const U=Object.assign({version:\"dev\",updateRepo:\"\",nativeBase:\"\",build:0,serverUrl:\"\",demoSaja:!1,xlsxUrl:\"\",panel:!1},window.CT_CONFIG||{}),H=window.Capacitor||null,O=!!(H&&H.isNativePlatform&&H.isNativePlatform()),K=H&&H.Plugins||{};function q(a,t){const i=new Error(t);return i.jenis=a,i}function F(a){const t=String(a&&a.message||a||\"\");return/timeout|timed out/i.test(t)?q(\"lambat\",\"Sambungan terlalu lama. Coba lagi.\"):!navigator.onLine||/failed to fetch|network|unable to resolve|connect|ssl|host|load failed/i.test(t)?q(\"offline\",\"Tidak ada koneksi internet. Periksa jaringan lalu coba lagi.\"):q(\"jaringan\",\"Data belum bisa diambil\"+(t?\" (\"+t+\")\":\"\")+\". Coba lagi; bila tetap gagal, hubungi admin.\")}async function V(a,t){if(O&&K.CapacitorHttp){const i=await K.CapacitorHttp.get({url:a,headers:{Accept:\"application/json\"},connectTimeout:15e3,readTimeout:t||2e4});if(200!==i.status)throw new Error(\"HTTP \"+i.status);return\"string\"==typeof i.data?JSON.parse(i.data):i.data}const i=new AbortController,s=setTimeout(()=>i.abort(),t||2e4);try{const t=await fetch(a,{cache:\"no-store\",signal:i.signal});if(!t.ok)throw new Error(\"HTTP \"+t.status);return await t.json()}finally{clearTimeout(s)}}const R=/^https:\\/\\/script\\.google\\.com\\/(?:a\\/macros\\/[^/]+|macros)\\/s\\/[A-Za-z0-9_-]+\\/exec$/,z={_janji:null,url(){if(\"dev\"===U.version){const a=d.get(\"ct.server.dev\");if(a&&R.test(a))return a}const a=d.get(\"ct.server.remote\");if(a&&R.test(a.url||\"\"))return a.url;if(R.test(U.serverUrl||\"\"))return U.serverUrl;const t=String(d.get(\"ct.server\")||\"\").trim().replace(/[?#].*$/,\"\");return R.test(t)?t:\"\"},sinkron(a){if(!U.updateRepo||U.demoSaja||U.panel)return Promise.resolve(this.url());const t=d.get(\"ct.server.remote\");if(!a&&t&&Date.now()-(t.at||0)<18e5)return Promise.resolve(this.url());if(this._janji)return this._janji;const i=\"https://raw.githubusercontent.com/\"+U.updateRepo+\"/main/app/server.json\";return this._janji=V(i,12e3).then(a=>{const t=String(a&&a.url||\"\").trim().replace(/[?#].*$/,\"\");d.set(\"ct.server.remote\",{url:R.test(t)?t:\"\",at:Date.now()})}).catch(()=>{}).then(()=>(this._janji=null,this.url())),this._janji}},J={aktif:!1,info:null,_mesin:null,async mulai(){return this._mesin?(this.aktif=!0,this.info):(window.CT_DEMO||await new Promise((a,t)=>{const i=document.createElement(\"script\");i.src=\"demo/demo-backend.js\",i.onload=a,i.onerror=()=>t(q(\"demo\",\"Data demo tidak bisa dimuat.\")),document.head.appendChild(i)}),await D(30),this._mesin=window.CT_DEMO.buat(),this.info=this._mesin.info,this.aktif=!0,this.info)},berhenti(){this.aktif=!1},kirim(a,t){return this._mesin.kirim(a,t)}},W=/\\b(Exception|Error|failed|Service|Spreadsheets?|timeout|timed out|undefined|null|Cannot|invoked|quota|exceeded|not defined|is not a function|We're sorry|permission|Authorization)\\b/i,G={importRawData:1,tambahFacility:1,daftarkanFacilityExisting:1,importLokasiAktif:1,copyLokasiDariFacility:1};async function Y(a,...t){const i=t.map(a=>void 0===a?null:a);let s;if(J.aktif)s=await J.kirim(a,JSON.parse(JSON.stringify(i)));else if(U.panel)try{s=JSON.parse(await new Promise((t,s)=>{window.google.script.run.withSuccessHandler(t).withFailureHandler(s).panelApi(a,JSON.stringify(i))}))}catch(a){throw F(a)}else{let t=z.url();if(t||(t=await z.sinkron(!0)),!t)throw q(\"tanpa-server\",\"Aplikasi belum tersambung. Hubungi admin aplikasi.\");const e=JSON.stringify({action:a,args:i,client:\"apk\",v:U.version});let n;try{n=await async function(a,t,i){const s={\"Content-Type\":\"text/plain;charset=utf-8\"};if(O&&K.CapacitorHttp){const e=await K.CapacitorHttp.request({method:\"POST\",url:a,headers:s,data:t,connectTimeout:2e4,readTimeout:i});if(e.status<200||e.status>=300)throw new Error(\"HTTP \"+e.status);return e.data}const e=new AbortController,n=setTimeout(()=>e.abort(),i);try{const i=await fetch(a,{method:\"POST\",headers:s,body:t,redirect:\"follow\",signal:e.signal});if(!i.ok)throw new Error(\"HTTP \"+i.status);return await i.text()}catch(a){throw a&&\"AbortError\"===a.name?new Error(\"timeout\"):a}finally{clearTimeout(n)}}(t,e,G[a]?39e4:75e3)}catch(a){const t=F(a);throw\"offline\"!==t.jenis&&z.sinkron(!0),t}if(s=n,\"string\"==typeof n)try{s=JSON.parse(n)}catch(a){throw z.sinkron(!0),q(\"ditolak\",/<html|<!doctype/i.test(n)?'Aplikasi belum diizinkan mengambil data. Hubungi admin aplikasi (akses Web App harus \"Siapa saja\").':\"Data belum bisa dibaca. Coba lagi; bila tetap gagal, hubungi admin.\")}}if(s&&!0===s.ok)return void 0===s.result?null:s.result;if(s&&!1===s.ok){const t=String(s.error||\"\"),i=/^Fungsi /.test(t)?\"fungsi-tak-ada\":/^User tidak dikenali|login ulang/i.test(t)?\"sesi\":/^Akses ditolak/i.test(t)?\"akses\":\"server\";let e=t||\"Permintaan gagal diproses. Coba lagi; bila tetap gagal, hubungi admin.\";\"server\"===i&&W.test(t)&&(console.warn(\"Galat server pada \"+a+\": \"+t),e=\"Sistem sedang terganggu. Coba lagi sebentar lagi.\");const n=q(i,e);throw n.asli=t,n}throw q(\"backend-lama\",\"Sistem perlu diperbarui. Hubungi admin aplikasi.\")}const Q={user:null,muat(){const a=d.get(\"ct.sesi\");return this.user=a&&a.username?a:null,this.user},simpan(a){this.user=a,J.aktif||d.set(\"ct.sesi\",a)},hapus(){this.user=null,d.del(\"ct.sesi\")},dari:a=>({username:a.displayName,role:a.role,akses:!!a.aksesSettingOverflow,facilityId:a.facilityId||\"\",facilityName:a.facilityName||\"\",facilityCode:a.facilityCode||\"\",facilityStatus:a.facilityStatus||\"\"})},X=()=>Q.user?Q.user.role:\"\",Z={admin:()=>\"admin\"===X()||\"developer\"===X(),validasi:()=>\"inventory\"===X()||Z.admin(),config:()=>!!Q.user&&(Q.user.akses||\"developer\"===X()),developer:()=>\"developer\"===X()},aa={kunci:a=>\"ct.data.\"+(Q.user?Q.user.username:\"-\")+\".\"+a,get(a){return J.aktif?null:d.get(this.kunci(a))},set(a,t){J.aktif||d.set(this.kunci(a),t)},bersih(){try{Object.keys(localStorage).forEach(a=>{0===a.indexOf(\"ct.data.\")&&localStorage.removeItem(a)})}catch(a){}}},ta={daftar:[],terkirim:{},_pendengar:[],_timer:0,_sibuk:!1,FUNGSI:{cycle:\"submitCount\",validasi:\"submitValidasi\"},LAMA_SEMBUNYI:24e4,muat(){this.daftar=(d.get(\"ct.antrean\")||[]).map(a=>Object.assign(a,{status:\"kirim\"===a.status?\"antre\":a.status})),this.terkirim=d.get(\"ct.terkirim\")||{},this._rapikan()},_simpan(){J.aktif||(d.set(\"ct.antrean\",this.daftar),d.set(\"ct.terkirim\",this.terkirim))},_rapikan(){const a=Date.now()-this.LAMA_SEMBUNYI;Object.keys(this.terkirim).forEach(t=>{this.terkirim[t]<a&&delete this.terkirim[t]})},kosongkan(){this.daftar=[],this.terkirim={},clearTimeout(this._timer)},dengar(a){this._pendengar.push(a)},_kabar(a,t,i){this._pendengar.forEach(s=>{try{s({tipe:a,entri:t,pesan:i})}catch(a){console.error(a)}})},jumlah(a){return this.daftar.filter(t=>!a||t.jenis===a).length},_fac:()=>Q.user&&Q.user.facilityId||\"-\",sembunyi(a,t){const i=this._fac(),s=a+\":\"+i+\":\"+t;return!!this.daftar.some(s=>s.jenis===a&&(s.fac||\"-\")===i&&String(s.kunci)===String(t))||!!this.terkirim[s]&&Date.now()-this.terkirim[s]<this.LAMA_SEMBUNYI},tambah(a,t){this.daftar.forEach(a=>{\"tahan\"===a.status&&(a.status=\"antre\",a.kirimAt=Date.now())});const i=Object.assign({id:Date.now().toString(36)+Math.random().toString(36).slice(2,7),at:Date.now(),coba:0,fac:this._fac()},a,{status:t?\"tahan\":\"antre\",kirimAt:Date.now()+(t||0)});return this.daftar.push(i),this._simpan(),this._kabar(\"ubah\",i),this.jalan(),i},urung(a){const t=this.daftar.findIndex(t=>t.id===a);if(t<0||\"tahan\"!==this.daftar[t].status)return null;const i=this.daftar.splice(t,1)[0];return this._simpan(),this._kabar(\"ubah\",i),i},lepas(){this.daftar.forEach(a=>{\"tahan\"===a.status&&(a.status=\"antre\",a.kirimAt=Date.now())}),this.jalan()},cobaLagi(){this.daftar.forEach(a=>{\"antre\"===a.status&&(a.kirimAt=Math.min(a.kirimAt,Date.now()))}),this.jalan()},tahanSampai(a,t){this._ditahan=!0;const i=()=>{this._ditahan&&(this._ditahan=!1,this.jalan())};Promise.resolve(a).then(i,i),setTimeout(i,t)},buangFacilityLain(a,t){const i=String(a||\"\").toLowerCase(),s=this.daftar.length;return this.daftar=this.daftar.filter(a=>\"kirim\"===a.status||String(a.args[1]||\"\").toLowerCase()!==i||(a.fac||\"-\")===(t||\"-\")),this.daftar.length!==s&&this._simpan(),s-this.daftar.length},jalan(){if(clearTimeout(this._timer),this._sibuk||this._ditahan||!this.daftar.length)return;const a=Date.now(),t=this.daftar.find(t=>t.kirimAt<=a);if(!t){const t=Math.min.apply(null,this.daftar.map(a=>a.kirimAt));return void(this._timer=setTimeout(()=>this.jalan(),Math.max(50,t-a)))}this._kirim(t)},MAKS_COBA_SERVER:8,SABAR_SERVER_MS:9e5,async _kirim(a){this._sibuk=!0,a.status=\"kirim\",this._kabar(\"ubah\",a);let t=null,i=null;if(Q.user&&String(a.args[1]||\"\").toLowerCase()===Q.user.username.toLowerCase()&&(a.fac||\"-\")!==this._fac())i=q(\"akses\",\"Facility akun ini berubah sebelum hasil terkirim.\");else try{t=await Y.apply(null,[this.FUNGSI[a.jenis]].concat(a.args,[a.at,a.fac&&\"-\"!==a.fac?a.fac:\"\"]))}catch(a){i=a}this._sibuk=!1;const s=()=>{const t=this.daftar.indexOf(a);t>=0&&this.daftar.splice(t,1)},e=t&&t.message||i&&i.message||\"\",n=!!i&&[\"offline\",\"lambat\",\"jaringan\",\"ditolak\",\"tanpa-server\",\"backend-lama\"].indexOf(i.jenis)>=0,l=i?\"server\"===i.jenis:/coba lagi|sibuk|sebentar|gagal masuk antri/i.test(e),r=l&&(a.cobaServer||0)>=this.MAKS_COBA_SERVER&&Date.now()-(a.gagalSejak||Date.now())>=this.SABAR_SERVER_MS;t&&t.success?(s(),this.terkirim[a.jenis+\":\"+(a.fac||\"-\")+\":\"+a.kunci]=Date.now(),this._rapikan(),this._simpan(),this._kabar(\"terkirim\",a)):n||l&&!r?(n||(a.cobaServer=(a.cobaServer||0)+1,a.gagalSejak=a.gagalSejak||Date.now()),a.status=\"antre\",a.coba++,a.kirimAt=Date.now()+Math.min(6e4,3e3*Math.pow(2,Math.min(a.coba-1,5))),this._simpan(),this._kabar(\"tertunda\",a,e)):(s(),this._simpan(),this._kabar(\"ditolak\",a,l?\"Sistem sedang sibuk. Hitung ulang lokasi ini nanti.\":e||\"Hasil tidak diterima.\"),t&&\"FACILITY_BERUBAH\"===t.kode&&this.tahanSampai(lt(),4e3)),this.jalan()}};window.addEventListener(\"online\",()=>ta.cobaLagi());const ia={};document.addEventListener(\"click\",a=>{const t=a.target.closest(\"[data-aksi]\");if(!t||t.disabled)return;const i=ia[t.dataset.aksi];i?(a.preventDefault(),i(t,a)):console.error(\"Aksi tidak dikenal: \"+t.dataset.aksi)});const sa={_t:0,_el:null,tampil(t,i){const s=i||{};this._el||(this._el=document.createElement(\"div\"),document.body.appendChild(this._el));const e=this._el;e.className=\"toast\"+(s.atas?\" toast--atas\":\"\")+(s.galat?\" toast--galat\":\"\"),e.setAttribute(\"role\",s.galat?\"alert\":\"status\"),o(e,u`<span class=\"toast__teks\">${t}</span>${s.aksi?u`<button type=\"button\" class=\"toast__aksi\">${s.aksi}</button>`:\"\"}`);const n=a(\".toast__aksi\",e);n&&n.addEventListener(\"click\",()=>{this.sembunyi(),s.onAksi&&s.onAksi()}),e.offsetWidth,e.classList.add(\"is-on\"),clearTimeout(this._t),this._t=setTimeout(()=>this.sembunyi(),s.lama||(s.aksi?5200:3200))},sembunyi(){clearTimeout(this._t),this._el&&this._el.classList.remove(\"is-on\")},galat(a){this.tampil(a&&a.message||String(a),{galat:!0,lama:4500})}},ea={tumpukan:[],buka(i){const s=i||{},e=document.createElement(\"div\");e.className=\"lembar-wadah\",o(e,u`<div class=\"lembar-tabir\"></div>\n      <div class=\"lembar${s.penuh?\" lembar--penuh\":\"\"}\" role=\"dialog\" aria-modal=\"true\" aria-label=\"${s.judul||\"\"}\">\n        <div class=\"lembar__pegangan\"></div>\n        <div class=\"lembar__kepala\"><h2 class=\"lembar__judul\">${s.judul||\"\"}</h2><button type=\"button\" class=\"iconbtn iconbtn--polos\" data-tutup aria-label=\"Tutup\">${I(\"tutup\")}</button></div>\n        <div class=\"lembar__isi\"></div>\n        ${s.kaki?u`<div class=\"lembar__kaki\">${s.kaki}</div>`:\"\"}\n      </div>`),document.body.appendChild(e);const n={el:e,isi:a(\".lembar__isi\",e),kaki:a(\".lembar__kaki\",e),onTutup:s.onTutup,terkunci:!1,judul(t){a(\".lembar__judul\",e).textContent=t},tutup:()=>this.tutup(n)};return s.isi&&o(n.isi,s.isi),a(\".lembar-tabir\",e).addEventListener(\"click\",()=>{n.terkunci||n.tutup()}),t(\"[data-tutup]\",e).forEach(a=>a.addEventListener(\"click\",()=>{n.terkunci||n.tutup()})),this.tumpukan.push(n),e.offsetWidth,e.classList.add(\"is-on\"),n},tutup(a){const t=a||this.tumpukan[this.tumpukan.length-1];if(!t)return!1;const i=this.tumpukan.indexOf(t);if(i>=0&&this.tumpukan.splice(i,1),t.el.classList.remove(\"is-on\"),setTimeout(()=>t.el.remove(),260),t.onTutup){const a=t.onTutup;t.onTutup=null,a()}return!0},tutupSemua(){for(;this.tumpukan.length;)this.tutup()}};function na(a){const i=a||{};return new Promise(a=>{let s=!1;const e=ea.buka({judul:i.judul||\"Lanjutkan?\",isi:u`${i.pesan?u`<p class=\"redup\" style=\"margin-bottom:18px\">${i.pesan}</p>`:\"\"}\n        <div class=\"baris-tombol\"><button type=\"button\" class=\"btn btn--garis\" data-j=\"0\">${i.batal||\"Batal\"}</button>\n        <button type=\"button\" class=\"btn${i.bahaya?\" btn--bahaya\":\"\"}\" data-j=\"1\">${i.ya||\"Ya, lanjutkan\"}</button></div>`,onTutup:()=>a(s)});t(\"[data-j]\",e.isi).forEach(a=>a.addEventListener(\"click\",()=>{s=\"1\"===a.dataset.j,e.tutup()}))})}const la={_el:null,_poll:0,K:2*Math.PI*34,tampil(a,t){this.sembunyi();const i=document.createElement(\"div\");i.className=\"sibuk\",i.setAttribute(\"role\",\"alertdialog\"),i.setAttribute(\"aria-busy\",\"true\"),o(i,u`<div class=\"sibuk__kotak\"><div class=\"sibuk__cincin\"><svg viewBox=\"0 0 84 84\"><circle class=\"jalur\" cx=\"42\" cy=\"42\" r=\"34\"/><circle class=\"isi\" cx=\"42\" cy=\"42\" r=\"34\" stroke-dasharray=\"${this.K}\" stroke-dashoffset=\"${this.K}\"/></svg><div class=\"sibuk__persen\">0%</div></div>\n      <div class=\"sibuk__teks\">${a}</div><div class=\"sibuk__ket\">${t||\"Jangan tutup aplikasi.\"}</div></div>`),document.body.appendChild(i),this._el=i},persen(t,i){if(!this._el)return;const s=Math.max(0,Math.min(100,Number(t)||0));a(\".isi\",this._el).style.strokeDashoffset=this.K*(1-s/100),a(\".sibuk__persen\",this._el).textContent=Math.round(s)+\"%\",i&&(a(\".sibuk__ket\",this._el).textContent=i)},pantau(a){clearInterval(this._poll),this._poll=setInterval(()=>{Y(\"getImportProgress\",a).then(a=>{a&&a.percent&&this.persen(a.percent,a.label)}).catch(()=>{})},1500)},sembunyi(){clearInterval(this._poll),this._el&&(this._el.remove(),this._el=null)}},ra={_isi:{},pasang(a,t,i){this._isi[a]={html:t,kelas:i},this._gambar()},lepas(a){delete this._isi[a],this._gambar()},_gambar(){let a=i(\"pita\");const t=Object.keys(this._isi);if(!t.length)return a&&a.remove(),document.body.classList.remove(\"ada-pita\"),void oa();a||(a=document.createElement(\"div\"),a.id=\"pita\",document.body.appendChild(a));const s=this._isi.offline?\"offline\":t[0];a.className=\"pita\"+(this._isi[s].kelas?\" \"+this._isi[s].kelas:\"\"),o(a,this._isi[s].html),document.body.classList.add(\"ada-pita\"),document.documentElement.style.setProperty(\"--pita-t\",a.offsetHeight-(parseFloat(getComputedStyle(a).paddingTop)||0)+5+\"px\"),oa()}},ua=[728,668,578,558];function oa(){let a=i(\"ukurAman\");a||(a=document.createElement(\"div\"),a.id=\"ukurAman\",a.style.cssText=\"position:fixed;visibility:hidden;pointer-events:none;top:0;left:0;width:0;height:0;padding-top:var(--sat);padding-bottom:var(--sab)\",document.body.appendChild(a));const t=getComputedStyle(a),s=parseFloat(t.paddingTop)||0,e=parseFloat(t.paddingBottom)||0,n=i(\"pita\")?Math.max(0,i(\"pita\").offsetHeight-s):0,l=window.innerHeight-s-e-n,r=ua.map((a,t)=>l<a?String(t+1):\"\").filter(Boolean).join(\" \");document.documentElement.getAttribute(\"data-rapat\")!==r&&document.documentElement.setAttribute(\"data-rapat\",r)}function da(){navigator.onLine?ra.lepas(\"offline\"):ra.pasang(\"offline\",u`Tidak ada internet. Hasil tetap tersimpan di HP.`,\"pita--jingga\")}window.addEventListener(\"resize\",oa),window.addEventListener(\"online\",da),window.addEventListener(\"offline\",da);const ca={},pa=[{id:\"home\",nama:\"Home\",ikon:\"home\",boleh:()=>!0},{id:\"cycle\",nama:\"Cycle\",ikon:\"kotak\",boleh:()=>!0,lencana:\"jingga\"},{id:\"validasi\",nama:\"Validasi\",ikon:\"perisai\",boleh:()=>Z.validasi(),lencana:\"ungu\"},{id:\"verifikasi\",nama:\"Verifikasi\",ikon:\"selisih\",boleh:()=>Z.validasi(),lencana:\"merah\"},{id:\"upload\",nama:\"Upload\",ikon:\"unggah\",boleh:()=>Z.admin()},{id:\"report\",nama:\"Report\",ikon:\"grafik\",boleh:()=>Z.admin()}],ha=[{id:\"upload\",nama:\"Upload\",ikon:\"unggah\",boleh:()=>Z.admin()},{id:\"config\",nama:\"Config\",ikon:\"atur\",boleh:()=>Z.config()}],ma={tab:\"\",halaman:[],lencana:{},daftarTab:()=>(U.panel?ha:pa).filter(a=>a.boleh()),awal(){const a=this.daftarTab();return U.panel?a.length?a[0].id:\"\":\"home\"},bangun(){const a=i(\"nav\");o(a,u`${this.daftarTab().map(a=>u`<button type=\"button\" class=\"nav__item\" data-aksi=\"tab\" data-tab=\"${a.id}\" id=\"tab-${a.id}\">${I(a.ikon)}<span>${a.nama}</span></button>`)}${U.panel?u`<button type=\"button\" class=\"nav__item\" data-aksi=\"akun\" id=\"tab-akun\">${I(\"orang\")}<span>Akun</span></button>`:\"\"}`),this._lencana()},setLencana(a,t){this.lencana[a]=Number(t)||0,this._lencana()},_lencana(){pa.forEach(t=>{const s=i(\"tab-\"+t.id);if(!s)return;const e=a(\".nav__lencana\",s);e&&e.remove();const n=this.lencana[t.id];t.lencana&&n>0&&s.insertAdjacentHTML(\"beforeend\",'<span class=\"nav__lencana nav__lencana--'+t.lencana+'\">'+(n>99?\"99+\":n)+\"</span>\"),s.setAttribute(\"aria-label\",t.nama+(n>0?\", \"+n:\"\"))})},_tampil(a){t(\".scr\").forEach(t=>t.classList.toggle(\"is-on\",t.id===\"scr-\"+a)),window.scrollTo(0,0)},aktif(){return this.halaman.length?this.halaman[this.halaman.length-1].id:this.tab},ke(a,i){this.daftarTab().some(t=>t.id===a)||(a=this.awal());const s=this.aktif();s&&ca[s]&&ca[s].keluar&&ca[s].keluar(),this.halaman=[],this.tab=a,document.body.classList.remove(\"tanpa-nav\"),t(\".nav__item\").forEach(t=>{const i=t.dataset.tab===a;t.classList.toggle(\"is-on\",i),i?t.setAttribute(\"aria-current\",\"page\"):t.removeAttribute(\"aria-current\")}),this._tampil(a),ca[a]&&ca[a].masuk&&ca[a].masuk(i||{})},dorong(a,t){const i=this.aktif();i&&ca[i]&&ca[i].keluar&&ca[i].keluar(),this.halaman.push({id:a}),document.body.classList.add(\"tanpa-nav\"),this._tampil(a),ca[a]&&ca[a].masuk&&ca[a].masuk(t||{})},kembali(){if(ea.tumpukan.length){return ea.tumpukan[ea.tumpukan.length-1].terkunci||ea.tutup(),!0}if(this.halaman.length){const a=this.halaman.pop();ca[a.id]&&ca[a.id].keluar&&ca[a.id].keluar();const t=this.aktif();return this.halaman.length||document.body.classList.remove(\"tanpa-nav\"),this._tampil(t),ca[t]&&ca[t].masuk&&ca[t].masuk({kembali:!0}),!0}return!(!this.tab||this.tab===this.awal()||\"login\"===this.tab)&&(this.ke(this.awal()),!0)}};ia.tab=a=>ma.ke(a.dataset.tab),ia.kembali=()=>ma.kembali(),ia.ke=a=>{ea.tutupSemua(),ma.ke(a.dataset.tab)};const ka=(a,t,i,s)=>u`<div class=\"kosong\"><div class=\"kosong__ic\">${I(a,\"ic--besar\")}</div><div class=\"kosong__t\">${t}</div>${i?u`<div class=\"kosong__s\">${i}</div>`:\"\"}${s||\"\"}</div>`,ga=a=>u`<div class=\"muat\"><span class=\"putar\"></span>${a||\"Memuat…\"}</div>`,ba=(a,t)=>u`<div class=\"tumpuk\">${Array.from({length:a||3},()=>u`<div class=\"kerangka\" style=\"height:${t||64}px;border-radius:var(--r-kartu)\"></div>`)}</div>`,fa=(a,t)=>ka(a&&\"offline\"===a.jenis?\"awan_mati\":\"awas\",a&&\"offline\"===a.jenis?\"Tidak ada koneksi\":\"Data tidak bisa dimuat\",a&&a.message||\"\",u`<button type=\"button\" class=\"btn btn--tenang\" data-aksi=\"${t}\">Coba lagi</button>`),va=(a,t,i)=>u`<header class=\"hd\"><div class=\"hd__teks\"><h1 class=\"hd__judul\">${a}</h1>${t?u`<div class=\"hd__ket\">${t}</div>`:\"\"}</div>${i||\"\"}</header>`,ya=(a,t,i)=>u`<header class=\"hd\"><button type=\"button\" class=\"iconbtn\" data-aksi=\"kembali\" aria-label=\"Kembali\">${I(\"kiri\")}</button><div class=\"hd__teks\"><h1 class=\"hd__judul\" style=\"font-size:21px\">${a}</h1>${t?u`<div class=\"hd__ket\">${t}</div>`:\"\"}</div>${i||\"\"}</header>`;let $a=null;function _a(){return window.XLSX?Promise.resolve(window.XLSX):($a||($a=new Promise((a,t)=>{const i=document.createElement(\"script\");i.src=U.xlsxUrl||\"vendor/xlsx.full.min.js\",i.onload=()=>a(window.XLSX),i.onerror=()=>{$a=null,t(new Error(\"Pembaca Excel tidak bisa dimuat.\"))},document.head.appendChild(i)})),$a)}async function wa(a,t){if(U.demoSaja)sa.tampil(\"Ekspor ke Excel tersedia di aplikasi Android.\");else try{const i=await _a(),s=i.utils.book_new();t.forEach(a=>i.utils.book_append_sheet(s,i.utils.json_to_sheet(a.baris.length?a.baris:[{Info:\"Tidak ada data\"}]),String(a.nama).slice(0,31)));const e=String(a).replace(/[\\\\/:*?\"<>|]+/g,\"_\");if(O&&K.Filesystem&&K.FileOpener){const a=i.write(s,{bookType:\"xlsx\",type:\"base64\"});await K.Filesystem.writeFile({path:e,data:a,directory:\"CACHE\"});const t=(await K.Filesystem.getUri({path:e,directory:\"CACHE\"})).uri;await K.FileOpener.open({filePath:t,contentType:\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\",openWithDefault:!1})}else i.writeFile(s,e);sa.tampil(\"File \"+e+\" dibuat.\")}catch(a){sa.galat(new Error(\"Ekspor gagal: \"+(a&&a.message||a)))}}const Sa={lebar:a=>Math.max(280,Math.round(a.clientWidth||a.parentNode.clientWidth||320)),tren(i,s,n,l){const r=this.lebar(i),d=196,p=r-6,h=166,k=s.length,g=(p-6)/Math.max(1,k),b=Math.max(1,Math.max.apply(null,s.map(a=>a.total))),f=a=>6+g*(a+.5),v=a=>a/b*106,y=Math.max(8,Math.min(34,.56*g));let $='<line x1=\"6\" y1=\"166\" x2=\"'+p+'\" y2=\"'+h+'\" stroke=\"var(--garis)\" stroke-width=\"1\"/>',_=\"\",w=\"\",S=!1;s.forEach((a,t)=>{const i=f(t)-y/2;if(a.total>0){const e=v(a.hit),n=v(a.total-a.hit-a.discrepancy),l=a.discrepancy>0?Math.max(3,v(a.discrepancy)):0;let r=h;e>0&&(r-=e,$+='<rect x=\"'+i+'\" y=\"'+r+'\" width=\"'+y+'\" height=\"'+e+'\" rx=\"3\" fill=\"var(--hijau)\"/>'),n>.5&&(r-=n,$+='<rect x=\"'+i+'\" y=\"'+r+'\" width=\"'+y+'\" height=\"'+n+'\" fill=\"var(--ungu)\" opacity=\"0.75\"/>'),l>0&&(r-=l,$+='<rect x=\"'+i+'\" y=\"'+r+'\" width=\"'+y+'\" height=\"'+l+'\" rx=\"1.5\" fill=\"var(--merah)\"/>');const u=(s=a.akurasi,26+30*(1-Math.max(0,Math.min(100,s)-80)/20));_+=(S?\" L \":\" M \")+f(t)+\" \"+u,S=!0,w+='<circle cx=\"'+f(t)+'\" cy=\"'+u+'\" r=\"3.6\" fill=\"var(--kertas)\" stroke=\"var(--tinta)\" stroke-width=\"2\"/>'}else S=!1,$+='<line x1=\"'+(f(t)-5)+'\" y1=\"158\" x2=\"'+(f(t)+5)+'\" y2=\"158\" stroke=\"var(--garis-kuat)\" stroke-width=\"2\" stroke-linecap=\"round\"/>';var s;$+='<text x=\"'+f(t)+'\" y=\"186\" font-size=\"11.5\" text-anchor=\"middle\" data-lbl=\"'+t+'\">'+e(a.label)+\"</text>\"}),_&&($+='<path d=\"'+_.trim()+'\" fill=\"none\" stroke=\"var(--tinta)\" stroke-width=\"2\" stroke-linejoin=\"round\" stroke-linecap=\"round\"/>'+w),s.forEach((a,t)=>{$+='<rect data-i=\"'+t+'\" x=\"'+(6+g*t)+'\" y=\"0\" width=\"'+g+'\" height=\"'+d+'\" fill=\"transparent\"/>'}),i.setAttribute(\"viewBox\",\"0 0 \"+r+\" \"+d),i.setAttribute(\"height\",d),i.innerHTML='<rect data-sorot x=\"0\" y=\"6\" width=\"'+g+'\" height=\"184\" rx=\"10\" fill=\"var(--lantai)\" opacity=\"0\"/>'+$;const x=e=>{const l=s[e];if(!l)return;const r=a(\"[data-sorot]\",i);r.setAttribute(\"x\",6+g*e),r.setAttribute(\"opacity\",\"1\"),t(\"[data-lbl]\",i).forEach(a=>{const t=Number(a.dataset.lbl)===e;a.setAttribute(\"font-weight\",t?\"700\":\"400\"),a.style.fill=t?\"var(--tinta)\":\"\"}),n&&o(n,l.total>0?u`<b>${l.judul||l.label}</b><span>Akurasi <b>${m(l.akurasi)}%</b></span><span>Hit <b>${c(l.hit)}</b></span><span>Discrepancy <b>${c(l.discrepancy)}</b></span>`:u`<b>${l.judul||l.label}</b><span>Tidak ada data</span>`)};t(\"rect[data-i]\",i).forEach(a=>a.addEventListener(\"click\",()=>x(Number(a.dataset.i))));let M=l;null==M&&(M=-1,s.forEach((a,t)=>{a.total>0&&(M=t)})),M>=0?x(M):n&&o(n,u`<span>Belum ada data pada periode ini.</span>`)},batang(a,t,i){const s=this.lebar(a),n=150,l=s-4,r=126,u=t.length,o=(l-4)/Math.max(1,u),d=Math.max(8,Math.min(30,.62*o)),p=Math.max(1,Math.max.apply(null,t.map(a=>a.nilai)));let h='<line x1=\"4\" y1=\"126\" x2=\"'+l+'\" y2=\"'+r+'\" stroke=\"var(--garis)\" stroke-width=\"1\"/>';const m=Math.ceil(u/9);t.forEach((a,t)=>{const s=a.nilai/p*106,n=4+o*(t+.5)-d/2;a.nilai>0&&(h+='<rect x=\"'+n+'\" y=\"'+(r-Math.max(2,s))+'\" width=\"'+d+'\" height=\"'+Math.max(2,s)+'\" rx=\"3\" fill=\"'+(i||\"var(--biru)\")+'\"/>',(o>=26||a.nilai===p)&&(h+='<text x=\"'+(n+d/2)+'\" y=\"'+(r-Math.max(2,s)-5)+'\" font-size=\"11\" font-weight=\"600\" text-anchor=\"middle\" style=\"fill:var(--tinta)\">'+c(a.nilai)+\"</text>\")),(t%m===0||u<=9)&&(h+='<text x=\"'+(n+d/2)+'\" y=\"143\" font-size=\"11.5\" text-anchor=\"middle\">'+e(a.label)+\"</text>\")}),a.setAttribute(\"viewBox\",\"0 0 \"+s+\" \"+n),a.setAttribute(\"height\",n),a.innerHTML=h},garis(a,t){const i=this.lebar(a),s=150,n=34,l=i-20,r=t.length,u=Math.max(0,Math.min(90,5*Math.floor((Math.min.apply(null,t.map(a=>a.akurasi))-2)/5))),o=a=>r>1?n+(l-n)*a/(r-1):(n+l)/2,d=a=>126-(a-u)/(100-u)*112;let c=\"\";[u,(u+100)/2,100].forEach(a=>{c+='<line x1=\"34\" y1=\"'+d(a)+'\" x2=\"'+l+'\" y2=\"'+d(a)+'\" stroke=\"var(--garis)\" stroke-width=\"1\" stroke-dasharray=\"3 4\"/><text x=\"28\" y=\"'+(d(a)+4)+'\" font-size=\"11\" text-anchor=\"end\">'+Math.round(a)+\"</text>\"});let p=\"\",h=\"\";const m=Math.ceil(r/Math.max(2,Math.floor((l-n)/46)));t.forEach((a,t)=>{p+=(t?\" L \":\"M \")+o(t)+\" \"+d(a.akurasi),h+='<circle cx=\"'+o(t)+'\" cy=\"'+d(a.akurasi)+'\" r=\"'+(r>20?2:3.5)+'\" fill=\"var(--hijau)\"/>',(t===r-1||t%m===0&&r-1-t>=.75*m)&&(c+='<text x=\"'+o(t)+'\" y=\"143\" font-size=\"11.5\" text-anchor=\"middle\">'+e(a.label)+\"</text>\")}),a.setAttribute(\"viewBox\",\"0 0 \"+i+\" \"+s),a.setAttribute(\"height\",s),a.innerHTML=c+'<path d=\"'+p+'\" fill=\"none\" stroke=\"var(--hijau)\" stroke-width=\"2.5\" stroke-linejoin=\"round\" stroke-linecap=\"round\"/>'+h},donat(a,t,i){const s=2*Math.PI*46,n=a.reduce((a,t)=>a+t.nilai,0);let r=0,u='<circle cx=\"60\" cy=\"60\" r=\"46\" fill=\"none\" stroke=\"var(--lantai-2)\" stroke-width=\"16\"/>';return n>0&&a.forEach(a=>{if(!a.nilai)return;const t=a.nilai/n*s;u+='<circle cx=\"60\" cy=\"60\" r=\"46\" fill=\"none\" stroke=\"'+a.warna+'\" stroke-width=\"16\" stroke-dasharray=\"'+Math.max(0,t-1.5)+\" \"+(s-Math.max(0,t-1.5))+'\" stroke-dashoffset=\"'+-r+'\" transform=\"rotate(-90 60 60)\"/>',r+=t}),l('<svg viewBox=\"0 0 120 120\" width=\"120\" height=\"120\" role=\"img\" aria-label=\"'+e(i||\"\")+'\">'+u+'<text x=\"60\" y=\"60\" text-anchor=\"middle\" font-size=\"26\" font-weight=\"700\" style=\"fill:var(--tinta);font-family:var(--huruf-rapat)\">'+e(t)+'</text><text x=\"60\" y=\"77\" text-anchor=\"middle\" font-size=\"11.5\" style=\"fill:var(--abu);font-family:var(--huruf)\">'+e(i||\"\")+\"</text></svg>\")}},xa=l('<svg class=\"lg__rak\" viewBox=\"0 0 360 208\" role=\"img\" aria-label=\"Rak gudang dengan satu label lokasi yang sudah dicentang\">\\n  <line x1=\"8\" y1=\"196\" x2=\"352\" y2=\"196\" stroke=\"var(--garis-kuat)\" stroke-width=\"2\" stroke-linecap=\"round\"/>\\n  <g fill=\"var(--kertas)\" stroke=\"var(--garis-kuat)\" stroke-width=\"1.5\">\\n    <rect x=\"38\" y=\"30\" width=\"52\" height=\"40\" rx=\"5\"/><rect x=\"96\" y=\"40\" width=\"40\" height=\"30\" rx=\"5\"/><rect x=\"141\" y=\"24\" width=\"27\" height=\"46\" rx=\"5\"/>\\n    <rect x=\"194\" y=\"36\" width=\"60\" height=\"34\" rx=\"5\"/><rect x=\"260\" y=\"26\" width=\"60\" height=\"44\" rx=\"5\"/>\\n    <rect x=\"38\" y=\"90\" width=\"44\" height=\"42\" rx=\"5\"/><rect x=\"88\" y=\"102\" width=\"76\" height=\"30\" rx=\"5\"/>\\n    <rect x=\"194\" y=\"88\" width=\"40\" height=\"44\" rx=\"5\"/><rect x=\"240\" y=\"96\" width=\"40\" height=\"36\" rx=\"5\"/><rect x=\"286\" y=\"104\" width=\"34\" height=\"28\" rx=\"5\"/>\\n    <rect x=\"38\" y=\"152\" width=\"60\" height=\"44\" rx=\"5\"/><rect x=\"104\" y=\"160\" width=\"60\" height=\"36\" rx=\"5\"/><rect x=\"194\" y=\"148\" width=\"126\" height=\"48\" rx=\"5\"/>\\n  </g>\\n  <g stroke=\"var(--garis-kuat)\" stroke-width=\"1.5\" stroke-linecap=\"round\">\\n    <path d=\"M64 30v9M116 40v8M224 36v9M290 26v10M60 90v9M126 102v8M214 88v9M68 152v9M134 160v8M257 148v10\"/>\\n  </g>\\n  <g fill=\"var(--jingga)\"><rect x=\"22\" y=\"70\" width=\"316\" height=\"7\" rx=\"2\"/><rect x=\"22\" y=\"132\" width=\"316\" height=\"7\" rx=\"2\"/></g>\\n  <g fill=\"var(--biru)\"><rect x=\"22\" y=\"12\" width=\"7\" height=\"184\" rx=\"2\"/><rect x=\"176.5\" y=\"12\" width=\"7\" height=\"184\" rx=\"2\"/><rect x=\"331\" y=\"12\" width=\"7\" height=\"184\" rx=\"2\"/></g>\\n  <g class=\"lg__label\">\\n    <rect x=\"52\" y=\"118\" width=\"118\" height=\"36\" rx=\"8\" fill=\"var(--kuning)\" stroke=\"var(--kuning-tepi)\" stroke-width=\"1.5\"/>\\n    <text x=\"111\" y=\"144\" text-anchor=\"middle\" font-size=\"25\" font-weight=\"700\" fill=\"var(--tinta-label)\" style=\"font-family:var(--huruf-rapat);letter-spacing:.02em\">B03.016.2</text>\\n  </g>\\n  <g class=\"lg__centang\">\\n    <circle cx=\"170\" cy=\"118\" r=\"14\" fill=\"var(--biru)\" stroke=\"var(--lantai)\" stroke-width=\"3\"/>\\n    <path d=\"m163.5 118.3 4.3 4.3 8.2-8.7\" fill=\"none\" stroke=\"var(--di-warna)\" stroke-width=\"2.6\" stroke-linecap=\"round\" stroke-linejoin=\"round\"/>\\n  </g>\\n</svg>'),Ma={status:\"cek\",gambar(){o(i(\"scr-login\"),u`<div class=\"lg\">\n      <div class=\"lg__gambar\">${xa}</div>\n      <h1 class=\"lg__judul\">Cycle Transaksi</h1>\n      <p class=\"lg__ket\">${U.panel?\"Panel admin: upload data, facility gudang, dan config. Masuk dengan NIK admin.\":\"Masuk dengan NIK untuk melihat tugas hari ini.\"}</p>\n      <form id=\"lgForm\" novalidate>\n        <label class=\"lbl\" for=\"lgNik\">NIK</label>\n        <input class=\"field\" id=\"lgNik\" type=\"text\" inputmode=\"text\" autocomplete=\"username\" autocapitalize=\"off\" autocorrect=\"off\" spellcheck=\"false\" enterkeyhint=\"go\" placeholder=\"contoh: 123456.nama\">\n        <div class=\"pesan pesan--galat\" id=\"lgPesan\" role=\"alert\"></div>\n        <button class=\"btn lg__tombol\" id=\"lgTombol\" type=\"submit\"><span>Masuk</span></button>\n      </form>\n      <div class=\"lg__status\" id=\"lgStatus\" aria-live=\"polite\"></div>\n      <div id=\"lgDemo\"></div>\n      <div class=\"lg__kaki\"><span>Versi ${U.version}</span>${U.panel?\"\":u`<button type=\"button\" class=\"tautan\" data-aksi=\"cek-versi\">Periksa pembaruan</button>`}</div>\n    </div>`),i(\"lgForm\").addEventListener(\"submit\",a=>{a.preventDefault(),this.kirim()}),i(\"lgNik\").addEventListener(\"input\",()=>this.pesan(\"\")),this.gambarStatus()},pesan(a){i(\"lgPesan\").textContent=a,i(\"lgNik\").classList.toggle(\"is-galat\",!!a)},setStatus(a){this.status=a,this.gambarStatus()},gambarStatus(){const a=i(\"lgStatus\");if(!a)return;const t={cek:[\"lg__titik--cek\",\"Menyambungkan…\"],siap:[\"lg__titik--siap\",\"Tersambung\"],offline:[\"lg__titik--awas\",\"Tidak ada koneksi internet\"],tanpa:[\"lg__titik--awas\",\"Aplikasi belum tersambung\"],galat:[\"lg__titik--awas\",\"Belum bisa tersambung. Coba lagi nanti.\"],demo:[\"lg__titik--demo\",\"Mode demo, memakai data contoh\"]}[this.status];o(a,u`<i class=\"lg__titik ${t[0]}\"></i><span>${t[1]}</span>`);const s=i(\"lgDemo\");U.demoSaja||\"demo\"===this.status?o(s,u`<div class=\"lg__demo\"><div class=\"lg__demo-judul\">Lihat sebagai</div><div class=\"lg__demo-pilih\">\n        <button type=\"button\" class=\"chip\" data-aksi=\"demo-masuk\" data-sebagai=\"admin\">Admin</button>\n        <button type=\"button\" class=\"chip\" data-aksi=\"demo-masuk\" data-sebagai=\"inventory\">Inventory</button>\n        <button type=\"button\" class=\"chip\" data-aksi=\"demo-masuk\" data-sebagai=\"petugas\">Petugas</button></div></div>`):\"tanpa\"===this.status?o(s,u`<div class=\"info info--biru lg__info\">${I(\"info\")}<div>Aplikasi ini belum disambungkan ke data gudang, jadi NIK belum bisa dipakai. Hubungi admin aplikasi, atau lihat dulu tampilannya dengan data contoh.\n        <button type=\"button\" class=\"btn btn--kecil\" data-aksi=\"demo-mulai\" style=\"margin-top:10px\">Coba mode demo</button></div></div>`):o(s,\"\")},async cekServer(){if(U.demoSaja||J.aktif)this.setStatus(\"demo\");else if(this.setStatus(\"cek\"),U.panel)try{await Y(\"getAppVersion\"),this.setStatus(\"siap\")}catch(a){this.setStatus(\"galat\")}else try{if(!(z.url()||await z.sinkron(!0)))return void this.setStatus(navigator.onLine?\"tanpa\":\"offline\");await Y(\"getAppVersion\"),this.setStatus(\"siap\")}catch(a){this.setStatus(\"offline\"===a.jenis?\"offline\":\"tanpa-server\"===a.jenis?\"tanpa\":\"galat\")}},async kirim(a){const t=String(a||i(\"lgNik\").value||\"\").trim();if(!t)return this.pesan(\"NIK belum diisi.\"),void i(\"lgNik\").focus();const s=i(\"lgTombol\");s.classList.add(\"is-sibuk\"),this.pesan(\"\");try{const a=await Y(\"getUserRole\",t);if(!a)return void this.pesan(\"NIK ini belum terdaftar atau sudah dinonaktifkan. Periksa ejaannya atau hubungi admin.\");\"demo\"!==this.status&&this.setStatus(\"siap\"),i(\"lgNik\").value=\"\",st(Q.dari(a))}catch(a){\"tanpa-server\"===a.jenis?this.setStatus(\"tanpa\"):\"offline\"===a.jenis&&this.setStatus(\"offline\"),this.pesan(a.message)}finally{s.classList.remove(\"is-sibuk\")}},async demo(a){const t=i(\"lgTombol\");t.classList.add(\"is-sibuk\"),this.pesan(\"\");try{const t=await J.mulai();ta.kosongkan(),this.setStatus(\"demo\"),a&&await this.kirim(\"admin\"===a?t.admin:\"inventory\"===a?t.inventory:t.petugas)}catch(a){this.pesan(a.message)}finally{t.classList.remove(\"is-sibuk\")}}};ca.login={masuk(){Ma.cekServer()}},ia[\"demo-mulai\"]=()=>Ma.demo(\"\"),ia[\"demo-masuk\"]=a=>Ma.demo(a.dataset.sebagai);const ja={data:null,at:0,sibuk:!1,galat:null,req:0,mulai(){const a=aa.get(\"home\");this.data=a?a.d:null,this.at=a?a.at:0,this.galat=null,this.sibuk=!1,this.req++},async muat(a){if(this.sibuk||!Q.user)return;if(!a&&this.data&&Date.now()-this.at<45e3)return;const t=++this.req,i=Q.user.username;this.sibuk=!0,this.tandaSegar();try{let a;try{a=await Y(\"getHomeBundle\",i)}catch(t){if(\"fungsi-tak-ada\"!==t.jenis)throw t;a=await this.muatLama(i)}if(t!==this.req)return;this.data=a,this.at=Date.now(),this.galat=null,aa.set(\"home\",{d:a,at:this.at})}catch(a){if(t!==this.req)return;if(\"sesi\"===a.jenis)return void et(\"Akun ini tidak lagi terdaftar. Hubungi admin.\");this.galat=a,this.data&&\"home\"===ma.aktif()&&sa.galat(a)}finally{t===this.req&&(this.sibuk=!1,this.lencana(),\"home\"===ma.aktif()&&this.gambar())}},async muatLama(a){const t=[Y(\"getHomeSummary\",a),Y(\"getMyPendingCount\",a)];Z.validasi()&&t.push(Y(\"getPlusMinusSummary\",a)),Z.admin()&&t.push(Y(\"getPendingBacklog\",a));const i=await Promise.all(t),s=Z.admin()?(i[3]||[]).reduce((a,t)=>a+t.pending,0):i[1];return Object.assign({},i[0],{role:X(),serverTime:Date.now(),myPending:i[1],outstanding:s,pendingValidasi:i[0].belumValidasi,openTasks:0,plusMinus:i[2]||null,perPetugas:[],lama:!0})},lencana(){const a=this.data;if(!a)return;const t=Aa.sisa(\"cycle\"),i=Aa.sisa(\"validasi\");ma.setLencana(\"cycle\",null!=t?t:a.myPending),ma.setLencana(\"validasi\",null!=i?i:a.pendingValidasi),ma.setLencana(\"verifikasi\",a.openTasks)},sisaSaya(a){const t=Aa.sisa(\"cycle\");return null!=t?t:Math.max(0,a.myPending-ta.jumlah(\"cycle\"))},tandaSegar(){const a=i(\"hmSegar\");a&&a.classList.toggle(\"is-putar\",this.sibuk)},gambar(){const a=Q.user;if(!a)return;const t=this.data;o(i(\"scr-home\"),u`\n      <header class=\"hm-kepala\">\n        <button type=\"button\" class=\"hm-avatar\" data-aksi=\"akun\" aria-label=\"Akun dan pengaturan\">${T(a.username)}</button>\n        <div class=\"hm-siapa\">\n          <div class=\"hm-sapa\">${function(){const a=(new Date).getHours();return a<11?\"Selamat pagi\":a<15?\"Selamat siang\":a<18?\"Selamat sore\":\"Selamat malam\"}()},</div>\n          <div class=\"hm-nama\">${a.username}</div>\n          <div class=\"hm-tempat\">${a.facilityName||\"Belum punya facility\"}</div>\n          <span class=\"pill pill--biru\">${j(a.role)}</span>\n        </div>\n        <button type=\"button\" class=\"hm-segar${this.sibuk?\" is-putar\":\"\"}\" id=\"hmSegar\" data-aksi=\"home-segar\" aria-label=\"Perbarui data\">${I(\"segar\",\"ic--kecil\")}<span>${this.at?$(this.at):\"Perbarui\"}</span></button>\n      </header>\n      <div id=\"hmPembaruan\"></div>\n      ${a.facilityName?\"\":u`<div class=\"info info--jingga\" style=\"margin-bottom:12px\">${I(\"awas\")}<div>Akun ini belum dimasukkan ke facility mana pun, jadi belum bisa menerima tugas atau mengirim hasil. Minta admin mengaturnya di Config.</div></div>`}\n      ${t?t.tanpaFacility?\"\":this.badan(t):this.galat?fa(this.galat,\"home-segar\"):u`<div class=\"kerangka\" style=\"height:74px;border-radius:var(--r-kartu);margin-bottom:12px\"></div><div class=\"kpi-kisi\">${[1,2,3,4].map(()=>u`<div class=\"kerangka\" style=\"height:112px;border-radius:var(--r-kartu)\"></div>`)}</div>`}\n    `),at.gambarKartu()},badan(a){return Z.admin()?u`${this.berikutnya(a)}${this.kpiAdmin(a)}${this.progres(a)}${this.plusMinus(a)}`:Z.validasi()?u`${this.berikutnya(a)}${this.kpiInventory(a)}${this.plusMinus(a)}`:this.tugasSaya(a)},berikutnya(a){const t=a.totalCycleHariIni+a.outstanding,i=this.sisaSaya(a);let s;if(i>0)s=[\"biru\",\"kotak\",c(i)+\" item menunggu Anda hitung\",\"Lanjutkan cycle dari lokasi berikutnya.\",\"home-mulai-cycle\"];else if(Z.admin()&&0===t)s=[\"biru\",\"unggah\",\"Belum ada tugas hari ini\",\"Upload data transaksi dan stok untuk membagi tugas.\",\"home-ke-upload\"];else if(Z.admin()&&a.outstanding>0)s=[\"jingga\",\"jam\",c(a.outstanding)+\" item belum di cycle\",\"Lihat siapa yang masih punya tugas.\",\"home-rinci-sisa\"];else if(a.pendingValidasi>0)s=[\"ungu\",\"perisai\",c(a.pendingValidasi)+\" item menunggu validasi\",Z.admin()?\"Selisih hitung perlu dihitung ulang oleh Inventory.\":\"Hitung ulang item yang selisih.\",\"home-ke-validasi\"];else if(a.openTasks>0)s=[\"merah\",\"selisih\",c(a.openTasks)+\" selisih belum selesai\",\"Tindak lanjuti di Verifikasi.\",\"home-ke-verifikasi\"];else{if(!(t>0||a.selesaiHariIni>0))return\"\";s=[\"hijau\",\"cek\",\"Semua beres untuk saat ini\",\"Tidak ada tugas, validasi, atau selisih yang menunggu.\",\"\"]}const e=u`<span class=\"lanjut__ic\">${I(s[1])}</span><span class=\"lanjut__teks\"><b>${s[2]}</b><span>${s[3]}</span></span>${s[4]?I(\"kanan\"):\"\"}`;return s[4]?u`<button type=\"button\" class=\"lanjut lanjut--${s[0]}\" data-aksi=\"${s[4]}\">${e}</button>`:u`<div class=\"lanjut lanjut--${s[0]}\">${e}</div>`},kpi:(a,t,i,s,e,n)=>u`<button type=\"button\" class=\"kpi kpi--${a}\" data-aksi=\"${n}\"><span class=\"kpi__atas\"><span class=\"kpi__ic\">${I(t)}</span><span class=\"kpi__judul\">${i}</span></span><span class=\"kpi__nilai angka\">${c(s)}</span><span class=\"kpi__ket\">${e}</span></button>`,kpiAdmin(a){return u`<div class=\"kpi-kisi\">\n      ${this.kpi(\"biru\",\"kotak\",\"Total Cycle\",a.totalCycleHariIni,\"Item sudah di cycle\",\"home-rinci-cycle\")}\n      ${this.kpi(\"jingga\",\"jam\",\"Outstanding Cycle\",a.outstanding,\"Item belum di cycle\",\"home-rinci-sisa\")}\n      ${this.kpi(\"ungu\",\"perisai\",\"Belum Validasi\",a.belumValidasi,\"Item menunggu validasi\",\"home-ke-validasi\")}\n      ${this.kpi(\"hijau\",\"cek\",\"Selesai\",a.selesaiHariIni,\"Item selesai\",\"home-ke-report\")}\n    </div>`},kpiInventory(a){return u`<div class=\"kpi-kisi\">\n      ${this.kpi(\"ungu\",\"perisai\",\"Belum Validasi\",a.pendingValidasi,\"Item menunggu Anda validasi\",\"home-ke-validasi\")}\n      ${this.kpi(\"merah\",\"selisih\",\"Verifikasi\",a.openTasks,\"Selisih belum selesai\",\"home-ke-verifikasi\")}\n      ${this.kpi(\"jingga\",\"jam\",\"Tugas Cycle\",this.sisaSaya(a),\"Item belum Anda hitung\",\"home-ke-cycle\")}\n      ${this.kpi(\"hijau\",\"cek\",\"Sudah Dihitung\",a.totalCycleHariIni,\"Item Anda hitung hari ini\",\"home-ke-cycle\")}\n    </div>`},progres(a){const t=a.totalCycleHariIni,i=a.outstanding,s=t+i;return s?u`<section class=\"card hm-kartu\">\n      <h2 class=\"card__judul\">Progress cycle hari ini</h2>\n      <p class=\"card__ket\">Item yang sudah dan belum di cycle hari ini.</p>\n      <div class=\"hm-sisa ${i?\"\":\"hm-sisa--beres\"}\"><b class=\"angka\">${i?c(i):\"\"}</b><span>${i?\"item belum di cycle\":\"Semua item sudah di cycle\"}</span></div>\n      <div class=\"bar\" role=\"progressbar\" aria-valuemin=\"0\" aria-valuemax=\"${s}\" aria-valuenow=\"${t}\"><i style=\"width:${Math.min(100,t/s*100)}%\"></i></div>\n      <div class=\"hm-dua\">\n        <div><span class=\"hm-persen\" style=\"color:var(--hijau-teks)\">${h(t,s)}</span><b class=\"angka\">${c(t)}</b><span>Item sudah di cycle</span></div>\n        <div><span class=\"hm-persen\" style=\"color:var(--jingga-teks)\">${h(i,s)} tersisa</span><b class=\"angka\">${c(s)}</b><span>Total item antrean</span></div>\n      </div>\n      ${i&&a.perPetugas&&a.perPetugas.length?u`<button type=\"button\" class=\"info info--jingga hm-catatan\" data-aksi=\"home-rinci-sisa\">${I(\"orang2\")}<span>${this.ringkasSisa(a)}</span>${I(\"kanan\",\"ic--kecil\")}</button>`:\"\"}\n    </section>`:u`<section class=\"card hm-kartu\"><h2 class=\"card__judul\">Progress cycle hari ini</h2><p class=\"card__ket\">Belum ada item dalam antrean. Angka di sini muncul setelah data di-upload.</p></section>`},ringkasSisa(a){const t=a.perPetugas.filter(a=>a.sisa>0),i=a.belumDitugaskan?u` ${c(a.belumDitugaskan)} item belum ditugaskan ke siapa pun.`:\"\";if(!t.length)return i||\"Lihat rincian per petugas.\";const s=t[0];return 1===t.length?u`Hanya <b>${s.nama}</b> yang masih punya tugas, ${c(s.sisa)} item.${i}`:u`<b>${t.length} petugas</b> masih punya tugas. Terbanyak <b>${s.nama}</b>, ${c(s.sisa)} item.${i}`},plusMinus(a){const t=a.plusMinus;if(!t)return\"\";const i=(a,t,i,s)=>u`<button type=\"button\" class=\"pm pm--${a}\" data-aksi=\"home-rinci-pm\" data-sisi=\"${a}\" ${s.totalSku?\"\":\"disabled\"}>\n      <span class=\"pm__nama\">${i}</span><span class=\"pm__nilai\">${I(t)}<b class=\"angka\">${c(s.totalQty)}</b><small>qty</small></span><span class=\"pm__sku\">${c(s.totalSku)} SKU</span></button>`;return u`<section class=\"card hm-kartu\">\n      <h2 class=\"card__judul\">Summary plus minus</h2>\n      <p class=\"card__ket\">${!t.plus.totalSku&&!t.minus.totalSku?\"Semua selisih sudah selesai.\":\"Selisih yang belum selesai di Verifikasi. Ketuk untuk melihat per SKU.\"}</p>\n      <div class=\"pm-kisi\">${i(\"plus\",\"naik\",\"Plus\",t.plus)}${i(\"minus\",\"turun\",\"Minus\",t.minus)}</div>\n    </section>`},tugasSaya(a){const t=this.sisaSaya(a),i=Math.max(a.totalCycleHariIni+a.myPending,t),s=i-t;return i?u`<section class=\"card hm-tugas\">\n      <h2 class=\"card__judul\">Tugas Anda hari ini</h2>\n      <div class=\"hm-sisa ${t?\"\":\"hm-sisa--beres\"}\"><b class=\"angka\">${t?c(t):\"\"}</b><span>${t?\"item belum dihitung\":\"Semua item sudah dihitung\"}</span></div>\n      <div class=\"bar bar--biru\" role=\"progressbar\" aria-valuemin=\"0\" aria-valuemax=\"${i}\" aria-valuenow=\"${s}\"><i style=\"width:${s/i*100}%\"></i></div>\n      <div class=\"hm-dari\"><span><b class=\"angka\">${c(s)}</b> dari ${c(i)} item sudah dihitung</span><span>${h(s,i)}</span></div>\n      ${t?u`<button type=\"button\" class=\"btn\" data-aksi=\"home-mulai-cycle\">${s?\"Lanjutkan cycle\":\"Mulai cycle\"}</button>`:\"\"}\n    </section>\n    ${a.selesaiHariIni<a.totalCycleHariIni?u`<div class=\"info\" style=\"margin-top:12px\">${I(\"perisai\")}<div><b>${c(a.totalCycleHariIni-a.selesaiHariIni)} item</b> hasil hitung Anda menunggu validasi oleh Inventory.</div></div>`:\"\"}`:u`<section class=\"card\">${ka(\"daftar\",\"Belum ada tugas untuk Anda\",\"Tugas muncul di sini setelah admin meng-upload dan membagi data.\")}</section>`},rinciPetugas(a){const t=this.data;if(!t)return;const i=(t.perPetugas||[]).filter(t=>\"sisa\"===a?t.sisa>0:t.dihitung>0).sort((t,i)=>\"sisa\"===a?i.sisa-t.sisa:i.dihitung-t.dihitung),s=Math.max(1,Math.max.apply(null,i.map(t=>\"sisa\"===a?t.sisa:t.dihitung).concat([1]))),e=i.map(t=>{const i=\"sisa\"===a?t.sisa:t.dihitung;return u`<div class=\"orang\"><div class=\"orang__atas\"><b>${t.nama}</b><span class=\"angka\">${c(i)}</span></div>\n        <div class=\"bar bar--tipis ${\"sisa\"===a?\"bar--jingga\":\"bar--biru\"}\"><i style=\"width:${i/s*100}%\"></i></div>\n        <div class=\"orang__ket\">${\"sisa\"===a?u`Sudah dihitung ${c(t.dihitung)} item`:u`${t.sisa?u`Sisa ${c(t.sisa)} item`:\"Tugas tuntas\"}${t.final<t.dihitung?u`, ${c(t.dihitung-t.final)} menunggu validasi`:\"\"}`}</div></div>`});ea.buka({judul:\"sisa\"===a?\"Belum di cycle\":\"Cycle hari ini\",isi:i.length?u`<p class=\"redup\" style=\"margin-bottom:14px\">${\"sisa\"===a?u`${c(t.outstanding)} item masih menunggu, termasuk sisa hari sebelumnya.${t.belumDitugaskan?u` ${c(t.belumDitugaskan)} di antaranya belum ditugaskan ke siapa pun.`:\"\"}`:u`${c(t.totalCycleHariIni)} item sudah dihitung oleh ${i.length} petugas.`}</p>${e}`:t.lama?u`<p class=\"redup\">Rincian per petugas tersedia setelah admin aplikasi memperbarui sistem.</p>`:ka(\"cek\",\"sisa\"===a?\"Tidak ada sisa tugas\":\"Belum ada yang dihitung\",\"sisa\"===a?\"Semua item sudah di cycle.\":\"Angka muncul setelah petugas mulai menghitung.\")})},rinciPlusMinus(a){const t=this.data&&this.data.plusMinus;if(!t)return;const i=t[a],s=\"plus\"===a;ea.buka({judul:s?\"Selisih plus\":\"Selisih minus\",isi:u`<p class=\"redup\" style=\"margin-bottom:12px\">${c(i.totalQty)} qty pada ${c(i.totalSku)} SKU, belum selesai di Verifikasi.</p>\n        <div class=\"list\">${i.items.map(a=>u`<div class=\"row\"><div class=\"row__isi\"><div class=\"row__t\">${a.article}</div><div class=\"row__s row__s--lipat\">${a.description}</div>\n          <div class=\"pm-lokasi\">${String(a.lokasi||\"\").split(\", \").slice(0,4).map(a=>C(a))}${a.jumlahLokasi>4?u`<span class=\"redup\">+${a.jumlahLokasi-4} lokasi</span>`:\"\"}</div></div>\n          <div class=\"row__ekor\"><span class=\"pill pill--${s?\"hijau\":\"merah\"}\">${s?\"+\":\"−\"}${c(a.qty)}</span></div></div>`)}</div>`,kaki:u`<button type=\"button\" class=\"btn btn--tenang\" data-aksi=\"ke\" data-tab=\"verifikasi\">Buka Verifikasi</button>`})}};ca.home={masuk(){ja.gambar(),ja.muat(!1)}},ia[\"home-segar\"]=()=>ja.muat(!0),ia[\"home-mulai-cycle\"]=()=>ma.ke(\"cycle\",{mulai:!0}),ia[\"home-ke-cycle\"]=()=>ma.ke(\"cycle\"),ia[\"home-ke-upload\"]=()=>ma.ke(\"upload\"),ia[\"home-ke-validasi\"]=()=>ma.ke(\"validasi\"),ia[\"home-ke-verifikasi\"]=()=>ma.ke(\"verifikasi\"),ia[\"home-ke-report\"]=()=>ma.ke(\"report\"),ia[\"home-rinci-sisa\"]=()=>ja.rinciPetugas(\"sisa\"),ia[\"home-rinci-cycle\"]=()=>ja.rinciPetugas(\"cycle\"),ia[\"home-rinci-pm\"]=a=>ja.rinciPlusMinus(a.dataset.sisi);const Ta={cycle:{judul:\"Cycle\",fungsi:\"getMyPendingTasks\",warna:\"biru\",ikon:\"kotak\",satuan:\"Qty hasil hitung\",kunci:a=>a.no,args:(a,t)=>[a.no,Q.user.username,t],ket:a=>c(a)+\" item menunggu dihitung\",kosong:[\"daftar\",\"Belum ada tugas untuk Anda\",\"Cek lagi setelah admin meng-upload dan membagi data baru.\"],tuntas:[\"Semua item selesai dihitung\",\"Hasil dikirim otomatis. Item yang selisih masuk ke Validasi.\"]},validasi:{judul:\"Validasi\",fungsi:\"getPendingValidasi\",warna:\"ungu\",ikon:\"perisai\",satuan:\"Qty hasil validasi\",kunci:a=>a.id,args:(a,t)=>[a.id,Q.user.username,t],ket:a=>c(a)+\" item selisih menunggu dihitung ulang\",kosong:[\"perisai\",\"Tidak ada item menunggu validasi\",\"Semua selisih sudah divalidasi.\"],tuntas:[\"Semua item sudah divalidasi\",\"Item yang tetap selisih masuk ke Verifikasi.\"]}},Aa={cycle:{items:null,at:0,sibuk:!1,galat:null,area:\"\",alat:\"\",cari:\"\",req:0},validasi:{items:null,at:0,sibuk:!1,galat:null,area:\"\",alat:\"\",cari:\"\",req:0},reset(){[\"cycle\",\"validasi\"].forEach(a=>Object.assign(this[a],{items:null,at:0,sibuk:!1,galat:null,area:\"\",alat:\"\",cari:\"\",req:this[a].req+1}))},susun(a,t){const i={},s=t.map((a,t)=>{const s=Object.assign({},a);return s.area=s.area||function(a){const t=String(a||\"\").trim().toUpperCase().match(/^([A-Z]+)/);return t?t[1].length>=2&&\"W\"===t[1].charAt(0)?t[1].charAt(1):t[1].charAt(0):\"\"}(s.lokasi),s.alat=s.levelGroupKey||function(a){const t=A(a);return t.rollcage||null==t.level||t.level<=2?\"tanpa_alat\":t.level<=4?\"tangga\":\"reach_truck\"}(s.lokasi),void 0===i[s.alat]&&(i[s.alat]=Object.keys(i).length),s._i=t,s});return s.sort((t,s)=>(t.area<s.area?-1:t.area>s.area?1:0)||i[t.alat]-i[s.alat]||(\"validasi\"===a?function(a,t){const i=P(a),s=P(t);for(let a=0;a<4;a++)if(i[a]!==s[a])return i[a]<s[a]?-1:1;return 0}(t.lokasi,s.lokasi):0)||t._i-s._i),s},tampak(a){const t=this[a],i=Ta[a];return(t.items||[]).filter(t=>!ta.sembunyi(a,i.kunci(t)))},tersaring(a){const t=this[a],i=t.cari.trim().toLowerCase();return this.tampak(a).filter(a=>(!t.area||a.area===t.area)&&(!t.alat||a.alat===t.alat)&&(!i||B(a.lokasi+\" \"+a.article+\" \"+a.description,i)))},sisa(a){return this[a].items?this.tampak(a).length:null},async muat(a,t){const i=this[a],s=Ta[a];if(i.sibuk||!Q.user)return;if(!t&&i.items&&Date.now()-i.at<6e4)return;const e=++i.req,n=Q.user.username;i.sibuk=!0,this.gambar(a);try{const t=await Y(s.fungsi,n);if(e!==i.req)return;i.items=this.susun(a,t||[]),i.at=Date.now(),i.galat=null,aa.set(\"tugas-\"+a,{items:t||[],at:i.at})}catch(t){if(e!==i.req)return;if(i.galat=t,!i.items){const t=aa.get(\"tugas-\"+a);t&&t.items&&(i.items=this.susun(a,t.items),i.at=t.at)}}finally{e===i.req&&(i.sibuk=!1,this.lencana(),ma.aktif()===a&&this.gambar(a))}},lencana(){const a=this.sisa(\"cycle\"),t=this.sisa(\"validasi\");null!=a&&ma.setLencana(\"cycle\",a),null!=t&&ma.setLencana(\"validasi\",t)},gambar(t){const s=this[t],e=Ta[t],n=i(\"scr-\"+t),l=this.tampak(t),r=u`<button type=\"button\" class=\"iconbtn${s.sibuk?\" is-putar\":\"\"}\" data-aksi=\"tugas-segar\" data-mode=\"${t}\" aria-label=\"Perbarui daftar\">${I(\"segar\")}</button>`;if(!s.items)return void o(n,u`${va(e.judul,\"\",r)}${s.galat?fa(s.galat,\"tugas-segar-\"+t):ba(5,62)}`);if(!l.length){const a=ta.jumlah(t),i=\"cycle\"===t&&Z.admin();return void o(n,u`${va(e.judul,\"\",r)}<section class=\"card\">${ka(e.kosong[0],e.kosong[1],a?c(a)+\" hasil terakhir sedang dikirim.\":i?\"Tugas dibagi ke petugas saat upload data. Anda ikut mendapat tugas bila memilih diri sendiri sebagai petugas.\":e.kosong[2],i?u`<button type=\"button\" class=\"btn btn--tenang\" data-aksi=\"ke\" data-tab=\"upload\">Upload data</button>`:\"\")}</section>`)}const d=a=>{const t={};return l.forEach(i=>{t[i[a]]=(t[i[a]]||0)+1}),t},p=d(\"area\"),h=d(\"alat\");s.area&&!p[s.area]&&(s.area=\"\"),s.alat&&!h[s.alat]&&(s.alat=\"\");const m=this.tersaring(t),k=(a,i,e,n)=>u`<button type=\"button\" class=\"chip${s[a]===i?\" is-on\":\"\"}\" data-aksi=\"tugas-saring\" data-mode=\"${t}\" data-jenis=\"${a}\" data-nilai=\"${i}\">${e}${null!=n?u` <b>${c(n)}</b>`:\"\"}</button>`,g=Object.keys(p).sort(),b=Object.keys(h);o(n,u`\n      ${va(e.judul,e.ket(l.length),r)}\n      ${s.galat?u`<div class=\"info info--jingga\" style=\"margin-bottom:12px\">${I(\"awan_mati\")}<div>Daftar belum bisa diperbarui: ${s.galat.message} Menampilkan daftar terakhir.</div></div>`:\"\"}\n      <button type=\"button\" class=\"btn tg-mulai\" data-aksi=\"tugas-mulai\" data-mode=\"${t}\" ${m.length?\"\":\"disabled\"}>${I(e.ikon)}<span id=\"tgMulai-${t}\">${this.teksMulai(m.length,l.length)}</span></button>\n      ${g.length>1?u`<div class=\"chips tg-cip\" role=\"group\" aria-label=\"Area gudang\">${k(\"area\",\"\",\"Semua area\")}${g.map(a=>k(\"area\",a,\"Area \"+a,p[a]))}</div>`:\"\"}\n      ${b.length>1?u`<div class=\"chips tg-cip\" role=\"group\" aria-label=\"Alat bantu\">${k(\"alat\",\"\",\"Semua alat\")}${b.map(a=>k(\"alat\",a,(L[a]||{}).nama||a,h[a]))}</div>`:\"\"}\n      ${l.length>8?u`<div class=\"cari tg-cari\">${I(\"cari\",\"ic--kecil\")}<input class=\"field\" type=\"search\" id=\"tgCari-${t}\" placeholder=\"Cari lokasi atau SKU\" value=\"${s.cari}\" autocomplete=\"off\"></div>`:\"\"}\n      <div id=\"tgDaftar-${t}\">${this.daftarHtml(t,m,120)}</div>\n    `);const f=i(\"tgCari-\"+t);f&&f.addEventListener(\"input\",E(()=>{s.cari=f.value;const e=this.tersaring(t),l=a(\".tg-mulai\",n),r=i(\"tgMulai-\"+t);i(\"tgDaftar-\"+t)&&(o(i(\"tgDaftar-\"+t),this.daftarHtml(t,e,120)),r&&(r.textContent=this.teksMulai(e.length,this.tampak(t).length)),l&&(l.disabled=!e.length))},160))},teksMulai:(a,t)=>a?a===t?\"Mulai dari lokasi pertama\":\"Mulai \"+c(a)+\" item terpilih\":\"Tidak ada item yang cocok\",daftarHtml(a,t,i){if(!t.length)return u`<p class=\"redup\" style=\"text-align:center;padding:24px 0\">Tidak ada item yang cocok.</p>`;const s=Ta[a],e=t.slice(0,i),n=a=>a.area+\"|\"+a.alat,l={};t.forEach(a=>{l[n(a)]=(l[n(a)]||0)+1});const r=Object.keys(l).length>1;let o=null;return u`<div class=\"list\">${e.map(t=>{const i=n(t),e=r&&i!==o?u`<div class=\"list__grup\"><span>${t.area?\"Area \"+t.area+\", \":\"\"}${((L[t.alat]||{}).nama||t.alat).toLowerCase()}</span><b>${c(l[i])}</b></div>`:\"\";return o=i,u`${e}<button type=\"button\" class=\"row\" data-aksi=\"tugas-buka\" data-mode=\"${a}\" data-kunci=\"${s.kunci(t)}\">\n      ${C(t.lokasi)}<span class=\"row__isi\"><span class=\"row__t\">${t.article}</span><span class=\"row__s\">${t.description}</span></span>\n      <span class=\"row__ekor\" title=\"${(L[t.alat]||{}).nama||\"\"}\">${I((L[t.alat]||{}).ikon||\"jalan\",\"ic--kecil\")}</span></button>`})}</div>\n      ${t.length>i?u`<p class=\"redup\" style=\"text-align:center;margin-top:12px\">Dan ${c(t.length-i)} item lagi. Gunakan pencarian atau mulai menghitung.</p>`:\"\"}`}},La={mode:\"cycle\",daftar:[],pos:0,awal:0,selesai:0,qty:\"\",mulaiAt:0,terakhir:null,log:[],areaLalu:\"\",alatLalu:\"\",buka(a,t){const i=Ta[a];this.mode=a,this.daftar=Aa.tersaring(a).slice(),this.pos=Math.max(0,this.daftar.findIndex(a=>String(i.kunci(a))===String(t))),this.awal=this.daftar.length,this.selesai=0,this.qty=\"\",this.mulaiAt=Date.now(),this.terakhir=null,this.log=[],this.areaLalu=\"\",this.alatLalu=\"\",this.lokasiLalu=\"\",this.daftar.length?ma.dorong(\"hitung\"):sa.tampil(\"Tidak ada item untuk dihitung.\")},item(){return this.daftar[this.pos]||null},gambar(){const a=Ta[this.mode],t=this.item(),s=i(\"scr-hitung\");if(s.className=\"scr is-on fk fk--\"+a.warna,!t)return void this.gambarTuntas();const e=this.areaLalu&&this.areaLalu!==t.area||this.alatLalu&&this.alatLalu!==t.alat,n=!!this.lokasiLalu&&this.lokasiLalu===t.lokasi;this.areaLalu=t.area,this.alatLalu=t.alat;const l=L[t.alat]||L.tanpa_alat;o(s,u`\n      <header class=\"fk__kepala\">\n        <button type=\"button\" class=\"iconbtn\" data-aksi=\"kembali\" aria-label=\"Kembali ke daftar\">${I(\"kiri\")}</button>\n        <div class=\"fk__judul\"><b>${a.judul}</b><span>Item ${c(Math.min(this.selesai+1,this.awal))} dari ${c(this.awal)}</span></div>\n        <button type=\"button\" class=\"fk__antre\" id=\"fkAntre\" data-aksi=\"fokus-log\" aria-label=\"Riwayat sesi ini\"></button>\n      </header>\n      <div class=\"bar bar--tipis bar--${a.warna}\"><i style=\"width:${this.awal?this.selesai/this.awal*100:0}%\"></i></div>\n      <div class=\"fk__kartu\" id=\"fkKartu\">\n        ${function(a){const t=A(a);if(3!==t.bagian.length)return u`<div class=\"rak-besar rak-besar--utuh\"><div class=\"rak-besar__bagian\"><span class=\"rak-besar__nilai\">${t.raw}</span><span class=\"rak-besar__nama\">Lokasi</span></div></div>`;const i=[\"Lorong\",\"Section\",t.rollcage?\"Rollcage\":\"Level\"],s=t.level&&t.level<=6?u`<div class=\"tingkat\" aria-hidden=\"true\">${[6,5,4,3,2,1].map(a=>u`<i class=\"${a===t.level?\"is-on\":\"\"}\"></i>`)}</div>`:\"\";return u`<div class=\"rak-besar\" aria-label=\"Lokasi ${t.raw}\">${t.bagian.map((a,t)=>u`<div class=\"rak-besar__bagian\"><span class=\"rak-besar__nilai\">${a}</span><span class=\"rak-besar__nama\">${i[t]}</span></div>`)}${s}</div>`}(t.lokasi)}\n        <div class=\"fk__barang\">\n          <div class=\"fk__teks\"><div class=\"fk__sku\">${t.article}</div><div class=\"fk__nama\">${t.description||\"\"}</div></div>\n          ${this.daftar.length>1?u`<button type=\"button\" class=\"fk__lewati\" data-aksi=\"fokus-lewati\" aria-label=\"Lewati item ini dulu\">${I(\"lewati\",\"ic--kecil\")}Lewati</button>`:\"\"}\n        </div>\n        <div class=\"fk__cip${e?\" is-baru\":\"\"}${n?\" fk__cip--sama\":\"\"}\">${n?u`<span class=\"pill pill--kuning\">${I(\"pin\",\"ic--kecil\")}Masih di lokasi ini, SKU berbeda</span>`:u`<span class=\"pill\">${I(\"pin\",\"ic--kecil\")}Area ${t.area||\"-\"}</span><span class=\"pill\">${I(l.ikon,\"ic--kecil\")}${l.nama}</span>`}</div>\n      </div>\n      <div class=\"fk__terakhir\" id=\"fkTerakhir\" aria-live=\"polite\"></div>\n      <div class=\"fk__qty\" aria-live=\"polite\"><output class=\"fk__angka\" id=\"fkAngka\"></output><span class=\"fk__satuan\">${a.satuan}</span></div>\n      <div class=\"pad\" id=\"fkPad\">\n        ${[1,2,3,4,5,6,7,8,9].map(a=>u`<button type=\"button\" data-aksi=\"pad\" data-k=\"${a}\">${a}</button>`)}\n        <button type=\"button\" data-aksi=\"pad\" data-k=\"hapus\" aria-label=\"Hapus angka terakhir\">${I(\"hapus_kiri\")}</button>\n        <button type=\"button\" data-aksi=\"pad\" data-k=\"0\">0</button>\n        <button type=\"button\" class=\"pad__simpan\" id=\"fkSimpan\" data-aksi=\"fokus-simpan\">Simpan</button>\n      </div>\n    `),this.gambarQty(),this.gambarAntre(),this.gambarTerakhir()},gambarTerakhir(){const a=i(\"fkTerakhir\");if(!a)return;const t=this.terakhir,s=t?t.entri.kirimAt-Date.now():0;clearTimeout(this._tTerakhir),!t||s<=0||\"tahan\"!==t.entri.status?o(a,\"\"):(o(a,u`<div class=\"simpanan\"><span class=\"simpanan__ic\">${I(\"cek\",\"ic--kecil\")}</span><span class=\"simpanan__teks\"><b>${t.it.lokasi}</b> tersimpan, qty ${c(t.entri.info.qty)}</span>\n      <button type=\"button\" class=\"simpanan__urung\" data-aksi=\"fokus-urung\">Urungkan</button><i class=\"simpanan__waktu\" style=\"animation-duration:${s}ms\"></i></div>`),this._tTerakhir=setTimeout(()=>this.gambarTerakhir(),s+30))},gambarQty(){const a=i(\"fkAngka\");a&&(a.textContent=\"\"===this.qty?\"0\":c(Number(this.qty)),a.classList.toggle(\"is-kosong\",\"\"===this.qty),i(\"fkSimpan\").disabled=\"\"===this.qty)},gambarAntre(){const a=i(\"fkAntre\");if(!a)return;const t=ta.jumlah();a.classList.toggle(\"is-tunggu\",t>1||t>0&&!navigator.onLine),o(a,u`${I(t>0?\"awan_kirim\":\"daftar\",\"ic--kecil\")}<span>${c(t>0?t:this.log.length)}</span>`)},tekan(a){this.item()&&(\"hapus\"===a?this.qty=this.qty.slice(0,-1):this.qty.length<5&&(this.qty=(\"0\"===this.qty?\"\":this.qty)+a),this.gambarQty())},simpan(){const a=this.item(),t=Ta[this.mode];if(!a||\"\"===this.qty)return;const s=Number(this.qty),e=ta.tambah({jenis:this.mode,kunci:t.kunci(a),args:t.args(a,s),info:{lokasi:a.lokasi,article:a.article,description:a.description,qty:s,sys:a.qtySystem}},5e3);this.terakhir={it:a,entri:e,pos:this.pos},this.lokasiLalu=a.lokasi,this.log.unshift({id:e.id,lokasi:a.lokasi,article:a.article,qty:s,sys:a.qtySystem,status:\"antre\"}),this.daftar.splice(this.pos,1),this.pos>=this.daftar.length&&(this.pos=0),this.selesai++,this.qty=\"\",Aa.lencana(),this.gambar();const n=i(\"fkKartu\");n&&n.classList.add(\"is-masuk\")},urung(){const a=this.terakhir;if(!a)return;const t=ta.urung(a.entri.id);if(!t)return this.terakhir=null,this.gambarTerakhir(),void sa.tampil(\"Hasil ini sudah terkirim, tidak bisa diurungkan.\");this.terakhir=null,this.log=this.log.filter(a=>a.id!==t.id),this.daftar.splice(Math.min(a.pos,this.daftar.length),0,a.it),this.pos=Math.min(a.pos,this.daftar.length-1),this.selesai=Math.max(0,this.selesai-1),this.qty=String(t.info.qty),this.lokasiLalu=\"\",Aa.lencana(),this.gambar()},lewati(){if(this.daftar.length<2)return;this.pos=(this.pos+1)%this.daftar.length,this.qty=\"\",this.lokasiLalu=\"\",this.gambar();const a=i(\"fkKartu\");a&&a.classList.add(\"is-masuk\")},gambarTuntas(){const a=Ta[this.mode],t=Aa.tampak(this.mode).length,s=Math.max(1,Math.round((Date.now()-this.mulaiAt)/6e4));o(i(\"scr-hitung\"),u`\n      <header class=\"fk__kepala\"><button type=\"button\" class=\"iconbtn\" data-aksi=\"kembali\" aria-label=\"Kembali ke daftar\">${I(\"kiri\")}</button>\n        <div class=\"fk__judul\"><b>${a.judul}</b><span>${c(this.selesai)} item selesai</span></div>\n        <button type=\"button\" class=\"fk__antre\" id=\"fkAntre\" data-aksi=\"fokus-log\" aria-label=\"Riwayat sesi ini\"></button></header>\n      <div class=\"bar bar--tipis bar--${a.warna}\"><i style=\"width:100%\"></i></div>\n      <div class=\"fk__terakhir fk__terakhir--atas\" id=\"fkTerakhir\" aria-live=\"polite\"></div>\n      <div class=\"fk__tuntas\">\n        <div class=\"fk__tuntas-ic\">${I(\"cek\",\"ic--besar\")}</div>\n        <h2>${t?\"Bagian ini selesai\":a.tuntas[0]}</h2>\n        <p>${c(this.selesai)} item dalam ${c(s)} menit. ${t?u`Masih ada ${c(t)} item di bagian lain.`:a.tuntas[1]}</p>\n        ${t?u`<button type=\"button\" class=\"btn\" data-aksi=\"fokus-lanjut-semua\">Lanjutkan ${c(t)} item lainnya</button>`:\"\"}\n        <button type=\"button\" class=\"btn ${t?\"btn--garis\":\"\"}\" data-aksi=\"ke\" data-tab=\"home\">Kembali ke Home</button>\n      </div>`),this.gambarAntre(),this.gambarTerakhir()},bukaLog(){const a=this.log.map(a=>{let t;if(\"terkirim\"===a.status){const i=a.qty-Number(a.sys);t=isNaN(i)?u`<span class=\"pill pill--hijau\">Terkirim</span>`:0===i?u`<span class=\"pill pill--hijau\">Cocok</span>`:u`<span class=\"pill pill--jingga\">Selisih ${p(i)}</span>`}else t=\"ditolak\"===a.status?u`<span class=\"pill pill--merah\">Ditolak</span>`:u`<span class=\"pill\">Menunggu kirim</span>`;return u`<div class=\"row\">${C(a.lokasi)}<div class=\"row__isi\"><div class=\"row__t\">${a.article}</div><div class=\"row__s\">Qty ${c(a.qty)}${a.pesan?u`, ${a.pesan}`:\"\"}</div></div><div class=\"row__ekor\">${t}</div></div>`});ea.buka({judul:\"Riwayat sesi ini\",isi:this.log.length?u`<p class=\"redup\" style=\"margin-bottom:12px\">${\"cycle\"===this.mode?\"Item yang selisih otomatis masuk antrean Validasi.\":\"Item yang tetap selisih masuk ke Verifikasi.\"} Hasil muncul setelah terkirim.</p><div class=\"list\">${a}</div>`:ka(\"daftar\",\"Belum ada hasil\",\"Hasil \"+(\"cycle\"===this.mode?\"hitung\":\"validasi\")+\" pada sesi ini muncul di sini.\")})},kembalikan(a){if(!this.log.some(t=>t.id===a.id))return;const t=Ta[this.mode],i=(Aa[this.mode].items||[]).find(i=>String(t.kunci(i))===String(a.kunci));!i||this.daftar.indexOf(i)>=0||(this.daftar.push(i),this.selesai=Math.max(0,this.selesai-1),this.gambar())},kabar(a){const t=this.log.find(t=>t.id===a.entri.id);t&&(\"terkirim\"===a.tipe?t.status=\"terkirim\":\"ditolak\"===a.tipe&&(t.status=\"ditolak\",t.pesan=a.pesan)),\"ditolak\"===a.tipe&&(sa.tampil(u`${(a.entri.info||{}).lokasi||\"Item\"} belum tersimpan: ${a.pesan}`,{galat:!0,atas:\"hitung\"===ma.aktif(),lama:6e3}),\"hitung\"===ma.aktif()&&a.entri.jenis===La.mode&&La.kembalikan(a.entri),Aa[a.entri.jenis].at=0),\"hitung\"===ma.aktif()&&this.gambarAntre(),Aa.lencana()}};[\"cycle\",\"validasi\"].forEach(a=>{ca[a]={masuk(t){Aa.gambar(a);const i=Aa.muat(a,!1);t&&t.mulai&&(Aa[a].items?La.buka(a):Promise.resolve(i).then(()=>{ma.aktif()===a&&Aa.tampak(a).length&&La.buka(a)}))}},ia[\"tugas-segar-\"+a]=()=>Aa.muat(a,!0)}),ca.hitung={masuk(){document.body.classList.add(\"mode-fokus\"),oa(),La.gambar(),document.addEventListener(\"keydown\",La._tombol)},keluar(){document.body.classList.remove(\"mode-fokus\"),document.removeEventListener(\"keydown\",La._tombol),clearTimeout(La._tTerakhir),ta.lepas(),Aa[La.mode].cari=\"\"}},La._tombol=a=>{ea.tumpukan.length||a.ctrlKey||a.metaKey||a.altKey||(/^[0-9]$/.test(a.key)?(La.tekan(a.key),a.preventDefault()):\"Backspace\"===a.key?(La.tekan(\"hapus\"),a.preventDefault()):\"Enter\"===a.key&&(La.simpan(),a.preventDefault()))},ia[\"tugas-segar\"]=a=>Aa.muat(a.dataset.mode,!0),ia[\"tugas-saring\"]=a=>{Aa[a.dataset.mode][a.dataset.jenis]=a.dataset.nilai,Aa.gambar(a.dataset.mode)},ia[\"tugas-mulai\"]=a=>La.buka(a.dataset.mode),ia[\"tugas-buka\"]=a=>La.buka(a.dataset.mode,a.dataset.kunci),ia.pad=a=>La.tekan(a.dataset.k),ia[\"fokus-simpan\"]=()=>La.simpan(),ia[\"fokus-lewati\"]=()=>La.lewati(),ia[\"fokus-urung\"]=()=>La.urung(),ia[\"fokus-log\"]=()=>La.bukaLog(),ia[\"fokus-lanjut-semua\"]=()=>{const a=Aa[La.mode];a.area=\"\",a.alat=\"\",a.cari=\"\",ma.kembali(),La.buka(La.mode)},ta.dengar(a=>La.kabar(a));const Pa=[\"Salah Hitung\",\"Barang Sudah di Picking\"],Ca=[\"Adjustment Plus\",\"Adjustment Minus\"],Da=[\"No\",\"STORERKEY\",\"TRANTYPE\",\"SKU\",\"Description\",\"SKUGROUP\",\"LOT\",\"FROMLOC\",\"FROMID\",\"TOLOC\",\"TOID\",\"SOURCEKEY\",\"QTY\",\"ADDDATE\",\"ADDWHO\"],Ea=[\"Open\",\"Sedang Dicari\",\"Menunggu Konfirmasi\",\"Selesai\"],Ba=a=>\"Open\"===a?\"Verifikasi\":\"Sedang Dicari\"===a?\"Sedang dicari\":\"Menunggu Konfirmasi\"===a?\"Menunggu konfirmasi\":a;function Na(a){const t=String(a||\"\").split(/\\r?\\n/).map(a=>a.replace(/\\s+$/,\"\")).filter(a=>a.trim());if(!t.length)throw new Error(\"Bukti masih kosong.\");const i=t.map(a=>{const t=a.split(\"\\t\");return Da.map((a,i)=>void 0!==t[i]?t[i].trim():\"\")}),s=i.findIndex(a=>!a[2]||!a[3]||!a[12])+1;if(s)throw new Error(\"Baris ke-\"+s+\" tidak lengkap. TRANTYPE, SKU, dan QTY wajib terisi; urutan kolom: \"+Da.join(\", \")+\".\");return i}function Ia(a,t,i,s,e){const n=String(t||\"\").trim().toUpperCase(),l=a=>String(a||\"\").trim().toUpperCase();if(Ca.indexOf(s)>=0){const r=a.filter(a=>l(a[3])===n&&\"ADJUSTMENT\"===l(a[2]));if(!r.length)return\"Belum ada baris ADJUSTMENT untuk SKU \"+t+\".\";const u=r.reduce((a,t)=>a+(Number(t[12])||0),0),o=\"Adjustment Plus\"===s?i:-i;return Math.round(u)!==Math.round(o)?\"Total qty adjustment \"+u+\", seharusnya \"+o+\" (tanda plus/minus harus sesuai).\":r.some(a=>l(a[9])===l(e))?\"\":\"TOLOC harus lokasi task ini (\"+e+\").\"}const r=a.filter(a=>l(a[3])===n&&(\"MOVE\"===l(a[2])||\"PICKING\"===l(a[2])));if(!r.length)return\"Belum ada baris MOVE atau PICKING untuk SKU \"+t+\".\";const u=r.reduce((a,t)=>a+Math.abs(Number(t[12])||0),0);return Math.round(u)!==Math.round(i)?\"Total qty di bukti \"+u+\", harus sama dengan selisih (\"+i+\").\":r.some(a=>a[7]||a[9])?\"\":\"Bukti harus memuat FROMLOC atau TOLOC.\"}function Ua(a,t,i,s,e){const n=String(t||\"\").trim().toUpperCase(),l=a=>String(a||\"\").trim().toUpperCase(),r=a.filter(a=>l(a[3])===n&&\"MOVE\"===l(a[2]));if(!r.length)return\"Belum ada baris MOVE untuk SKU \"+t+\".\";const u=r.reduce((a,t)=>a+Math.abs(Number(t[12])||0),0);return Math.round(u)!==Math.round(i)?\"Total qty MOVE \"+u+\", harus \"+i+\".\":r.some(a=>l(a[7])===l(s)&&l(a[9])===l(e))?\"\":\"Harus ada baris dengan FROMLOC \"+s+\" dan TOLOC \"+e+\".\"}function Ha(a){const t=String(a||\"\").trim();if(!t)return[];if(!/(MOVE|PICKING|LAINNYA):/i.test(t))return[[\"User WMS\",t.split(\",\").map(a=>a.trim()).filter(Boolean).join(\", \")]];const i={move:\"Transaksi Move terakhir\",picking:\"Transaksi Picking terakhir\",lainnya:\"User WMS\"};return t.split(\";\").map(a=>{const t=a.indexOf(\":\");if(t<0)return null;const s=a.slice(0,t).trim().toLowerCase(),e=a.slice(t+1).split(\",\").map(a=>a.trim()).filter(Boolean).join(\", \");return e?[i[s]||s,e]:null}).filter(Boolean)}const Oa={tasks:null,form:null,at:0,sibuk:!1,galat:null,tampilan:\"task\",saring:\"semua\",cari:\"\",tertutup:{},log:null,logDari:\"\",logSampai:\"\",logSibuk:!1,req:0,logReq:0,reset(){Object.assign(this,{tasks:null,form:null,at:0,sibuk:!1,galat:null,tampilan:\"task\",saring:\"semua\",cari:\"\",tertutup:{},log:null,logGalat:null,logDari:\"\",logSampai:\"\",logSibuk:!1,req:this.req+1,logReq:this.logReq+1})},async muat(a){if(this.sibuk||!Q.user)return;if(!a&&this.tasks&&Date.now()-this.at<6e4)return;const t=++this.req,i=Q.user.username;this.sibuk=!0,this.gambar();try{const a=await Promise.all([Y(\"getOpenTasks\",i),this.form?this.form:Y(\"getInvestigasiFormData\",i)]);if(t!==this.req)return;this.tasks=(a[0]||[]).map(a=>Object.assign(a,{namaValidator:a.namaValidator||a.nameValidator||\"\",selisih:Number(a.selisih)||0})),this.form=a[1],this.at=Date.now(),this.galat=null}catch(a){if(t!==this.req)return;this.galat=a}finally{t===this.req&&(this.sibuk=!1,\"verifikasi\"===ma.aktif()&&this.gambar(),this.lencana())}},aktif(){const a=Date.now()-24e4;return(this.tasks||[]).filter(t=>!(this.tertutup[t.id]>a))},lencana(){this.tasks&&(ma.setLencana(\"verifikasi\",this.aktif().length),ja.data&&(ja.data.openTasks=this.aktif().length))},tersaring(){const a=this.cari.trim().toLowerCase(),t=this.saring,i=Q.user.username;return this.aktif().filter(s=>(\"semua\"===t||\"kritis\"===t&&s.isCritical||\"minus\"===t&&s.selisih<0||\"plus\"===t&&s.selisih>0||\"saya\"===t&&s.picInvestigasi===i)&&(!a||B([s.lokasi,s.article,s.description,s.namaPetugas,s.picInvestigasi,s.namaValidator,s.addWhoTransaksi,Ba(s.statusTask)].join(\" \"),a)))},pasangan(){const a=a=>String(a.article).trim().toUpperCase()+\"|\"+Math.abs(a.selisih),t=this.aktif().filter(a=>a.selisih>0),i=[];return this.aktif().filter(a=>a.selisih<0).forEach(s=>t.forEach(t=>{a(t)===a(s)&&i.push({minus:s,plus:t})})),i},jumlahPasangan(a){const t={},i={};return a.forEach(a=>{const s=String(a.minus.article).trim().toUpperCase()+\"|\"+Math.abs(a.minus.selisih);(t[s]=t[s]||{})[a.minus.id]=1,(i[s]=i[s]||{})[a.plus.id]=1}),Object.keys(t).reduce((a,s)=>a+Math.min(Object.keys(t[s]).length,Object.keys(i[s]).length),0)},gambar(){const a=i(\"scr-verifikasi\"),t=u`<button type=\"button\" class=\"iconbtn${this.sibuk?\" is-putar\":\"\"}\" data-aksi=\"verif-segar\" aria-label=\"Perbarui\">${I(\"segar\")}</button>`;if(!this.tasks)return void o(a,u`${va(\"Verifikasi\",\"\",t)}${this.galat?fa(this.galat,\"verif-segar\"):ba(5,84)}`);const s=this.aktif(),e=this.pasangan(),n=this.jumlahPasangan(e),l=u`<div class=\"seg vf-seg\" role=\"tablist\">\n      <button type=\"button\" role=\"tab\" class=\"${\"task\"===this.tampilan?\"is-on\":\"\"}\" data-aksi=\"verif-tampilan\" data-v=\"task\">Task ${s.length?c(s.length):\"\"}</button>\n      <button type=\"button\" role=\"tab\" class=\"${\"pasangan\"===this.tampilan?\"is-on\":\"\"}\" data-aksi=\"verif-tampilan\" data-v=\"pasangan\">Pasangan ${n?c(n):\"\"}</button>\n      <button type=\"button\" role=\"tab\" class=\"${\"riwayat\"===this.tampilan?\"is-on\":\"\"}\" data-aksi=\"verif-tampilan\" data-v=\"riwayat\">Riwayat</button></div>`;let r;r=\"pasangan\"===this.tampilan?this.gambarPasangan(e):\"riwayat\"===this.tampilan?this.gambarRiwayat():this.gambarTask(s),o(a,u`${va(\"Verifikasi\",s.length?c(s.length)+\" selisih belum selesai\":\"Semua selisih sudah selesai\",t)}\n      ${this.galat?u`<div class=\"info info--jingga\" style=\"margin-bottom:12px\">${I(\"awan_mati\")}<div>Belum bisa diperbarui: ${this.galat.message}</div></div>`:\"\"}\n      ${l}${r}`);const d=i(\"vfCari\");d&&d.addEventListener(\"input\",E(()=>{this.cari=d.value,o(i(\"vfDaftar\"),this.daftarTask(this.tersaring()))},160)),\"riwayat\"===this.tampilan&&([\"vfLogDari\",\"vfLogSampai\"].forEach(a=>i(a).addEventListener(\"change\",()=>{this.logDari=i(\"vfLogDari\").value,this.logSampai=i(\"vfLogSampai\").value,this.muatLog()})),this.log||this.logSibuk||this.logGalat||this.muatLog())},gambarTask(a){if(!a.length)return u`<section class=\"card\">${ka(\"cek\",\"Semua selisih sudah selesai\",\"Selisih baru muncul di sini setelah validasi memastikan hasil hitung berbeda dari stok.\")}</section>`;const t=t=>a.filter(t).length,i=Q.user.username,s=(a,t,i)=>u`<button type=\"button\" class=\"chip${this.saring===a?\" is-on\":\"\"}\" data-aksi=\"verif-saring\" data-v=\"${a}\">${t} <b>${c(i)}</b></button>`;return u`<div class=\"cari\" style=\"margin-bottom:10px\">${I(\"cari\",\"ic--kecil\")}<input class=\"field\" type=\"search\" id=\"vfCari\" placeholder=\"Cari lokasi, SKU, atau nama\" value=\"${this.cari}\" autocomplete=\"off\"></div>\n      <div class=\"chips\" style=\"margin-bottom:12px\" role=\"group\" aria-label=\"Saring task\">${s(\"semua\",\"Semua\",a.length)}${s(\"kritis\",\"Kritis\",t(a=>a.isCritical))}${s(\"minus\",\"Minus\",t(a=>a.selisih<0))}${s(\"plus\",\"Plus\",t(a=>a.selisih>0))}${s(\"saya\",\"PIC saya\",t(a=>a.picInvestigasi===i))}</div>\n      <div id=\"vfDaftar\">${this.daftarTask(this.tersaring())}</div>\n      <button type=\"button\" class=\"btn btn--teks\" data-aksi=\"verif-ekspor\" style=\"margin:10px auto 0\">${I(\"unduh\",\"ic--kecil\")}Ekspor daftar ini ke Excel</button>`},tahap(a){const t=Math.max(0,Ea.indexOf(a));return u`<span class=\"tahap\" aria-hidden=\"true\">${[0,1,2].map(a=>u`<i class=\"${a<=t?\"is-on\":\"\"}\"></i>`)}</span>`},daftarTask(a){if(!a.length)return u`<p class=\"redup\" style=\"text-align:center;padding:24px 0\">Tidak ada task yang cocok.</p>`;return u`<div class=\"list\">${a.slice(0,80).map(a=>u`<button type=\"button\" class=\"row vf-row\" data-aksi=\"verif-buka\" data-id=\"${a.id}\">\n        <span class=\"row__isi\"><span class=\"vf-row__atas\">${C(a.lokasi)}<span class=\"row__t\">${a.article}</span></span>\n          <span class=\"row__s\">${a.description}</span>\n          <span class=\"vf-row__status\">${this.tahap(a.statusTask)}<span>${Ba(a.statusTask)}</span>${a.isCritical?u`<span class=\"pill pill--merah\">Kritis, ${c(a.umurHari)} hari</span>`:u`<span class=\"redup\">${a.umurHari?c(a.umurHari)+\" hari\":\"Hari ini\"}</span>`}</span></span>\n        <span class=\"row__ekor\"><span class=\"pill pill--${a.selisih>0?\"hijau\":\"merah\"} vf-selisih\">${a.selisih>0?\"+\":\"−\"}${c(Math.abs(a.selisih))}</span></span></button>`)}</div>\n      ${a.length>80?u`<p class=\"redup\" style=\"text-align:center;margin-top:12px\">Dan ${c(a.length-80)} task lagi. Persempit dengan pencarian atau saringan.</p>`:\"\"}`},gambarPasangan(a){return a.length?u`<p class=\"redup\" style=\"margin:0 2px 12px;font-size:14px\">Barang kurang di satu lokasi dan lebih di lokasi lain. Selesaikan keduanya sekaligus dengan satu bukti Move.${a.length>this.jumlahPasangan(a)?\" Ada selisih dengan beberapa calon pasangan: pilih yang lokasinya sesuai bukti Move.\":\"\"}</p>\n      <div class=\"tumpuk\">${a.map(a=>u`<button type=\"button\" class=\"card vf-pasang\" data-aksi=\"verif-pasang\" data-minus=\"${a.minus.id}\" data-plus=\"${a.plus.id}\">\n        <span class=\"row__t\">${a.minus.article}</span><span class=\"row__s\">${a.minus.description}</span>\n        <span class=\"vf-pasang__rute\"><span>${C(a.minus.lokasi)}<b style=\"color:var(--merah-teks)\">−${c(Math.abs(a.minus.selisih))}</b></span>${I(\"kanan\")}<span>${C(a.plus.lokasi)}<b style=\"color:var(--hijau-teks)\">+${c(Math.abs(a.plus.selisih))}</b></span></span></button>`)}</div>`:u`<section class=\"card\">${ka(\"kait\",\"Belum ada pasangan yang cocok\",\"Pasangan muncul bila ada selisih minus dan plus untuk SKU yang sama dengan qty yang sama persis.\")}</section>`},bukaTask(s){const e=(this.tasks||[]).find(a=>String(a.id)===String(s));if(!e)return;const n=this.form||{kategoriList:[],mapping:[],allUsers:[],statusFlow:Ea},l=n.statusFlow&&n.statusFlow.length?n.statusFlow:Ea,r=l.slice(l.indexOf(e.statusTask)+1),d=(a,t)=>t?u`<div class=\"rinci__baris\"><span>${a}</span><b>${t}</b></div>`:\"\",p=ea.buka({judul:\"Tindak lanjut selisih\",penuh:!0,isi:u`<div class=\"vf-kepala\">${C(e.lokasi)}<span class=\"pill pill--${e.selisih>0?\"hijau\":\"merah\"} vf-selisih\">${e.selisih>0?\"Plus \":\"Minus \"}${c(Math.abs(e.selisih))}</span>${e.isCritical?u`<span class=\"pill pill--merah\">Kritis</span>`:\"\"}</div>\n        <div class=\"vf-sku\">${e.article}</div><div class=\"redup\">${e.description}</div>\n        <div class=\"rinci\">${d(\"Dihitung oleh\",e.namaPetugas)}${d(\"Tanggal cycle\",w(String(e.tanggal).slice(0,10)))}${d(\"Validator\",e.namaValidator)}${d(\"PIC saat ini\",e.picInvestigasi)}${Ha(e.addWhoTransaksi).map(a=>d(a[0],a[1]))}${d(\"Umur task\",e.umurHari?c(e.umurHari)+\" hari\":\"Hari ini\")}</div>\n        <form id=\"vfForm\" novalidate>\n          <span class=\"lbl\">Status berikutnya</span>\n          <div class=\"seg\" id=\"vfStatus\">${r.map(a=>u`<button type=\"button\" data-s=\"${a}\" class=\"${\"Selesai\"===a?\"is-on\":\"\"}\">${\"Selesai\"===a?\"Selesai\":Ba(a)}</button>`)}</div>\n          <label class=\"lbl\" for=\"vfKategori\">Kategori selisih <small id=\"vfWajib1\"></small></label>\n          <select class=\"field\" id=\"vfKategori\"><option value=\"\">Belum diketahui</option>${n.kategoriList.map(a=>u`<option value=\"${a}\" ${e.kategori===a?\"selected\":\"\"}>${a}</option>`)}</select>\n          <label class=\"lbl\" for=\"vfPic\">PIC <small id=\"vfWajib2\"></small></label>\n          <select class=\"field\" id=\"vfPic\"></select>\n          <label class=\"lbl\" for=\"vfCatatan\">Catatan</label>\n          <textarea class=\"field\" id=\"vfCatatan\" rows=\"3\" placeholder=\"Apa yang sudah dicek atau ditemukan. Contoh: sudah cek lokasi B02.031.4, belum ketemu.\"></textarea>\n          <div class=\"blok\" id=\"vfBuktiWadah\" hidden>\n            <label class=\"lbl\" for=\"vfBukti\" id=\"vfBuktiLabel\">Bukti transaksi WMS</label>\n            <textarea class=\"field field--tempel\" id=\"vfBukti\" rows=\"4\" placeholder=\"Tempel baris dari WMS di sini\" spellcheck=\"false\"></textarea>\n            <div class=\"ket\" id=\"vfBuktiKet\"></div>\n            <div class=\"pesan\" id=\"vfBuktiCek\" aria-live=\"polite\"></div>\n          </div>\n          <div class=\"pesan pesan--galat\" id=\"vfPesan\" role=\"alert\"></div>\n        </form>`,kaki:u`<button type=\"button\" class=\"btn\" id=\"vfSimpan\">Tandai selesai</button>`}),h=()=>{const t=a(\"#vfStatus .is-on\",p.el);return t?t.dataset.s:r[r.length-1]},m=i(\"vfKategori\"),k=i(\"vfPic\"),g=i(\"vfBukti\"),b=()=>{const a=m.value;if(\"Salah Hitung\"===a&&e.namaValidator)return o(k,u`<option value=\"${e.namaValidator}\" selected>${e.namaValidator} (validator)</option>`),void(k.disabled=!0);k.disabled=!1;const t=n.mapping.filter(t=>t.kategori===a),i=a&&t.length?t:n.allUsers,s={},l=i.filter(a=>!s[a.username]&&(s[a.username]=!0)),r=l.some(a=>a.username===e.picInvestigasi);o(k,u`<option value=\"\">Pilih PIC</option>${l.map(a=>u`<option value=\"${a.username}\" ${a.username===e.picInvestigasi?\"selected\":\"\"}>${a.username}${a.role?\" (\"+j(a.role)+\")\":\"\"}</option>`)}${e.picInvestigasi&&!r?u`<option value=\"${e.picInvestigasi}\" selected>${e.picInvestigasi} (saat ini)</option>`:\"\"}`)},f=()=>\"Selesai\"===h()&&m.value&&Pa.indexOf(m.value)<0,v=()=>{const a=\"Selesai\"===h();if(i(\"vfSimpan\").textContent=a?\"Tandai selesai\":\"Simpan status\",i(\"vfWajib1\").textContent=a?\"\":\"(boleh dikosongkan dulu)\",i(\"vfWajib2\").textContent=a?\"\":\"(boleh dikosongkan dulu)\",i(\"vfBuktiWadah\").hidden=!f(),f()){const a=Ca.indexOf(m.value)>=0,t=Math.abs(e.selisih);i(\"vfBuktiLabel\").textContent=a?\"Bukti transaksi WMS (Adjustment)\":\"Bukti transaksi WMS (Move atau Picking)\",o(i(\"vfBuktiKet\"),a?u`Salin barisnya dari WMS tanpa judul kolom. Jenis transaksi ADJUSTMENT, total qty <b>${\"Adjustment Minus\"===m.value?\"−\":\"+\"}${c(t)}</b>, dan TOLOC <b>${e.lokasi}</b>.`:u`Salin barisnya dari WMS tanpa judul kolom. Baris MOVE atau PICKING untuk SKU ini dengan total qty <b>${c(t)}</b> dan lokasi asal atau tujuan.`),y()}},y=()=>{const a=i(\"vfBuktiCek\");if(!g.value.trim())return a.textContent=\"\",void(a.className=\"pesan\");try{const t=Na(g.value),i=Ia(t,e.article,Math.abs(e.selisih),m.value,e.lokasi);a.className=\"pesan \"+(i?\"pesan--galat\":\"pesan--ok\"),a.textContent=i||c(t.length)+\" baris terbaca. Bukti sesuai.\"}catch(t){a.className=\"pesan pesan--galat\",a.textContent=t.message}};t(\"#vfStatus button\",p.el).forEach(a=>a.addEventListener(\"click\",()=>{t(\"#vfStatus button\",p.el).forEach(t=>t.classList.toggle(\"is-on\",t===a)),v()})),m.addEventListener(\"change\",()=>{if(b(),\"Barang Sudah di Picking\"===m.value&&(t(\"#vfStatus button\",p.el).forEach(a=>a.classList.toggle(\"is-on\",\"Selesai\"===a.dataset.s)),i(\"vfCatatan\").value.trim()||(i(\"vfCatatan\").value=\"Barang sudah di picking, selisih terkonfirmasi.\"),!k.value)){const a=n.mapping.find(a=>a.kategori===m.value);k.value=a?a.username:Q.user.username}v()}),g.addEventListener(\"input\",E(y,250)),b(),v(),i(\"vfSimpan\").addEventListener(\"click\",()=>this.simpanTask(e,p,h()))},async simpanTask(a,t,s){const e=i(\"vfKategori\").value,n=i(\"vfPic\").value,l=i(\"vfCatatan\").value.trim(),r=\"Selesai\"===s,u=a=>{i(\"vfPesan\").textContent=a,a&&i(\"vfPesan\").scrollIntoView({block:\"nearest\"})};if(!l)return u(\"Catatan belum diisi.\");if(r&&!e)return u(\"Pilih kategori selisih untuk menandai selesai.\");if(r&&!n)return u(\"Pilih PIC untuk menandai selesai.\");let o=null;if(r&&Pa.indexOf(e)<0){if(!i(\"vfBukti\").value.trim())return u(\"Tempel bukti transaksi WMS untuk menandai selesai.\");try{o=Na(i(\"vfBukti\").value)}catch(a){return u(a.message)}const t=Ia(o,a.article,Math.abs(a.selisih),e,a.lokasi);if(t)return u(t)}u(\"\");const d=i(\"vfSimpan\");d.classList.add(\"is-sibuk\"),t.terkunci=!0;try{const i=await Y(\"updateTaskStatus\",a.id,Q.user.username,s,l,e,n,o,a.lokasi,a.article,a.selisih);if(!i.success)return void u(i.message);r?(this.tertutup[a.id]=Date.now(),this.log=null):(a.statusTask=s,e&&(a.kategori=e),n&&(a.picInvestigasi=n)),t.terkunci=!1,t.tutup(),sa.tampil(r?\"Task ditandai selesai.\":\"Status diperbarui: \"+Ba(s)+\".\"),this.gambar(),this.lencana()}catch(a){u(a.message)}finally{d.classList.remove(\"is-sibuk\"),t.terkunci=!1}},bukaPasangan(a,t){const s=this.pasangan().find(i=>String(i.minus.id)===String(a)&&String(i.plus.id)===String(t));if(!s)return;const e=Math.abs(s.minus.selisih),n=this.form||{allUsers:[]},l=ea.buka({judul:\"Selesaikan pasangan plus minus\",penuh:!0,isi:u`<div class=\"vf-sku\">${s.minus.article}</div><div class=\"redup\">${s.minus.description}</div>\n        <div class=\"vf-pasang__rute vf-pasang__rute--besar\"><span>${C(s.minus.lokasi)}<b style=\"color:var(--merah-teks)\">−${c(e)}</b></span>${I(\"kanan\")}<span>${C(s.plus.lokasi)}<b style=\"color:var(--hijau-teks)\">+${c(e)}</b></span></div>\n        <label class=\"lbl\" for=\"vpPic\">PIC</label>\n        <select class=\"field\" id=\"vpPic\"><option value=\"\">Pilih PIC</option>${n.allUsers.map(a=>u`<option value=\"${a.username}\">${a.username}${a.role?\" (\"+j(a.role)+\")\":\"\"}</option>`)}</select>\n        <label class=\"lbl\" for=\"vpCatatan\">Catatan</label>\n        <textarea class=\"field\" id=\"vpCatatan\" rows=\"2\" placeholder=\"Contoh: barang ketemu di lokasi tujuan, sudah dipindahkan sesuai bukti Move.\"></textarea>\n        <label class=\"lbl\" for=\"vpBukti\">Bukti transaksi WMS (Move)</label>\n        <textarea class=\"field field--tempel\" id=\"vpBukti\" rows=\"4\" placeholder=\"Tempel baris dari WMS di sini\" spellcheck=\"false\"></textarea>\n        <div class=\"ket\">Salin barisnya dari WMS tanpa judul kolom. Baris MOVE untuk SKU ini dengan FROMLOC <b>${s.minus.lokasi}</b>, TOLOC <b>${s.plus.lokasi}</b>, dan total qty <b>${c(e)}</b>.</div>\n        <div class=\"pesan\" id=\"vpCek\" aria-live=\"polite\"></div>\n        <div class=\"pesan pesan--galat\" id=\"vpPesan\" role=\"alert\"></div>`,kaki:u`<button type=\"button\" class=\"btn\" id=\"vpSimpan\">Tandai keduanya selesai</button>`});i(\"vpBukti\").addEventListener(\"input\",E(()=>{const a=i(\"vpCek\"),t=i(\"vpBukti\").value;if(t.trim())try{const i=Na(t),n=Ua(i,s.minus.article,e,s.minus.lokasi,s.plus.lokasi);a.className=\"pesan \"+(n?\"pesan--galat\":\"pesan--ok\"),a.textContent=n||c(i.length)+\" baris terbaca. Bukti sesuai.\"}catch(t){a.className=\"pesan pesan--galat\",a.textContent=t.message}else a.textContent=\"\"},250)),i(\"vpSimpan\").addEventListener(\"click\",async()=>{const a=a=>{i(\"vpPesan\").textContent=a},t=i(\"vpPic\").value,n=i(\"vpCatatan\").value.trim();if(!t)return a(\"Pilih PIC.\");if(!n)return a(\"Catatan belum diisi.\");let r;try{r=Na(i(\"vpBukti\").value)}catch(t){return a(\"Bukti masih kosong.\"===t.message?\"Tempel bukti Move dari WMS.\":t.message)}const u=Ua(r,s.minus.article,e,s.minus.lokasi,s.plus.lokasi);if(u)return a(u);a(\"\");const o=i(\"vpSimpan\");o.classList.add(\"is-sibuk\"),l.terkunci=!0;try{const i=await Y(\"closePlusMinusPair\",s.minus.id,s.plus.id,Q.user.username,n,t,r);if(!i.success)return void a(i.message);this.tasks=this.tasks.filter(a=>a.id!==s.minus.id&&a.id!==s.plus.id),this.log=null,l.terkunci=!1,l.tutup(),sa.tampil(\"Kedua task ditandai selesai.\"),this.gambar(),this.lencana(),ja.at=0}catch(t){a(t.message)}finally{o.classList.remove(\"is-sibuk\"),l.terkunci=!1}})},gambarRiwayat(){this.logDari||(this.logSampai=v(),this.logDari=y(this.logSampai,-30));const a=u`<div class=\"rentang\"><label><span>Dari</span><input type=\"date\" class=\"field field--kecil\" id=\"vfLogDari\" value=\"${this.logDari}\"></label><label><span>Sampai</span><input type=\"date\" class=\"field field--kecil\" id=\"vfLogSampai\" value=\"${this.logSampai}\"></label></div>`;return this.logSibuk||!this.log?u`${a}${this.logGalat?fa(this.logGalat,\"verif-log-lagi\"):ga(\"Memuat riwayat…\")}`:this.log.length?u`${a}<div class=\"list\">${this.log.map((a,t)=>u`<button type=\"button\" class=\"row vf-row\" data-aksi=\"verif-log-buka\" data-i=\"${t}\">\n      <span class=\"row__isi\"><span class=\"vf-row__atas\">${C(a.lokasi)}<span class=\"row__t\">${a.article}</span></span>\n        <span class=\"row__s\">${a.kategori||a.alasan}</span><span class=\"row__s\">Diselesaikan ${a.diselesaikanOleh}, ${a.waktuSelesai}</span></span>\n      <span class=\"row__ekor\"><span class=\"pill vf-selisih\">${Number(a.selisih)>0?\"+\":\"−\"}${c(Math.abs(Number(a.selisih)||0))}</span></span></button>`)}</div>`:u`${a}<section class=\"card\">${ka(\"daftar\",\"Belum ada task selesai\",\"Tidak ada task yang selesai pada rentang tanggal ini.\")}</section>`},async muatLog(){const a=++this.logReq;this.logSibuk=!0,this.logGalat=null,\"verifikasi\"===ma.aktif()&&\"riwayat\"===this.tampilan&&this.gambar();try{const t=await Y(\"getTaskLog\",300,Q.user.username,this.logDari,this.logSampai)||[];if(a!==this.logReq)return;this.log=t}catch(t){if(a!==this.logReq)return;this.logGalat=t,this.log=null}finally{a===this.logReq&&(this.logSibuk=!1,\"verifikasi\"===ma.aktif()&&\"riwayat\"===this.tampilan&&this.gambar())}},bukaLog(a){const t=(this.log||[])[a];if(!t)return;const i=(a,t)=>t?u`<div class=\"rinci__baris\"><span>${a}</span><b>${t}</b></div>`:\"\";ea.buka({judul:\"Task selesai\",isi:u`<div class=\"vf-kepala\">${C(t.lokasi)}<span class=\"pill vf-selisih\">${Number(t.selisih)>0?\"Plus \":\"Minus \"}${c(Math.abs(Number(t.selisih)||0))}</span></div>\n        <div class=\"vf-sku\">${t.article}</div><div class=\"redup\">${t.description||\"\"}</div>\n        <div class=\"rinci\">${i(\"Kategori\",t.kategori||t.alasan)}${i(\"PIC\",t.picUsername)}${i(\"Diselesaikan oleh\",t.diselesaikanOleh)}${i(\"Waktu selesai\",t.waktuSelesai)}${i(\"Dihitung oleh\",t.namaPetugas)}${i(\"Validator\",t.namaValidator)}${Ha(t.addWhoTransaksi).map(a=>i(a[0],a[1]))}</div>\n        ${t.catatan?u`<span class=\"lbl\">Catatan</span><div class=\"vf-catatan\">${t.catatan}</div>`:\"\"}`})},ekspor(){const a=this.tersaring().map(a=>({Lokasi:a.lokasi,SKU:a.article,Deskripsi:a.description,\"Qty Selisih\":a.selisih,Tanggal:String(a.tanggal).slice(0,10),Petugas:a.namaPetugas,Status:Ba(a.statusTask),\"Umur (hari)\":a.umurHari,PIC:a.picInvestigasi||\"\",Validator:a.namaValidator||\"\",\"User Transaksi WMS\":a.addWhoTransaksi||\"\"}));wa(\"Verifikasi_\"+v()+\"_\"+a.length+\"_task.xlsx\",[{nama:\"Task\",baris:a}])}};ca.verifikasi={masuk(){Oa.gambar(),Oa.muat(!1)}},ia[\"verif-segar\"]=()=>Oa.muat(!0),ia[\"verif-tampilan\"]=a=>{Oa.tampilan=a.dataset.v,Oa.gambar()},ia[\"verif-saring\"]=a=>{Oa.saring=a.dataset.v,Oa.gambar()},ia[\"verif-buka\"]=a=>Oa.bukaTask(a.dataset.id),ia[\"verif-pasang\"]=a=>Oa.bukaPasangan(a.dataset.minus,a.dataset.plus),ia[\"verif-ekspor\"]=()=>Oa.ekspor(),ia[\"verif-log-lagi\"]=()=>Oa.muatLog(),ia[\"verif-log-buka\"]=a=>Oa.bukaLog(Number(a.dataset.i));const Ka={type:[\"type\",\"trantype\",\"jenis transaksi\",\"transaction type\",\"tipe\"],article:[\"article\",\"sku\",\"kode barang\",\"item code\"],description:[\"description\",\"deskripsi\",\"nama barang\",\"item description\"],lokasiAwal:[\"from location\",\"fromloc\",\"lokasi awal\",\"source location\",\"lokasi picking\"],lokasiTujuan:[\"to location\",\"toloc\",\"lokasi tujuan\",\"destination location\",\"lokasi move\"],qty:[\"qty\",\"quantity\",\"qty transaksi\",\"jumlah\"],addWho:[\"addwho\",\"add who\",\"added by\",\"user wms\",\"created by\"]},qa={lokasi:[\"location\",\"loc\",\"lokasi\"],article:[\"article\",\"sku\",\"kode barang\",\"item code\"],qty:[\"qty\",\"quantity\",\"stock\",\"jumlah\"]},Fa=[\"outbound\",\"inbound\",\"storing\",\"lp\",\"maintenance\",\"inventory\",\"admin\"],Va=a=>{const t=null==a?\"\":String(a).trim(),i=/^=\"(.*)\"$/.exec(t);return i?i[1].replace(/\"\"/g,'\"').trim():t},Ra=/^(-|–|—|null|nan|n\\/a|#n\\/a)$/i;async function za(a,t,i){const s=await _a(),e=await new Promise((t,i)=>{const s=new FileReader;s.onload=()=>t(s.result),s.onerror=()=>i(new Error(\"File tidak bisa dibaca.\")),s.readAsArrayBuffer(a)}),n=s.read(new Uint8Array(e),{type:\"array\",raw:!0}),l=s.utils.sheet_to_json(n.Sheets[n.SheetNames[0]],{header:1});if(!l.length)throw new Error(\"File kosong.\");const r=l[0].map(a=>String(a||\"\").trim().toLowerCase().replace(/[_\\s-]+/g,\" \")),u={},o=[];if(Object.keys(t).forEach(a=>{u[a]=r.findIndex(i=>t[a].indexOf(i)>=0),u[a]<0&&(i||[]).indexOf(a)<0&&o.push(t[a][0])}),o.length)throw new Error(\"Kolom tidak ditemukan: \"+o.join(\", \")+\". Periksa judul kolom di baris pertama.\");const d=Object.keys(t),c=[];return l.forEach((a,t)=>{if(0===t||!a||!a.length)return;const i={};if(d.forEach(t=>{i[t]=Va(u[t]>=0?a[u[t]]:\"\")}),\"qty\"in i){const a=function(a){const t=a.replace(/\\s+/g,\"\");if(\"\"===t||Ra.test(t))return{nilai:\"\"};const i=/^-?\\d{1,3}(,\\d{3})+(\\.\\d+)?$/.test(t)?t.replace(/,/g,\"\"):t;return/^-?\\d+(\\.\\d+)?$/.test(i)?{nilai:String(Number(i))}:{nilai:a,salah:!0}}(i.qty);i.qty=a.nilai,a.salah&&(i.qtySalah=!0,i.baris=t+1)}c.push(i)}),c}const Ja={form:null,sibuk:!1,galat:null,transaksi:null,stok:null,pilih:{},rt:0,tangga:0,terbuka:!1,lama:!1,req:0,reset(){Object.assign(this,{form:null,sibuk:!1,galat:null,transaksi:null,stok:null,pilih:{},terbuka:!1,req:this.req+1})},kunciPilih:()=>\"ct.petugas.\"+(Q.user.facilityId||\"-\"),async muat(a){if(this.sibuk)return;if(this.form&&!a)return;const t=++this.req,i=Q.user.username;this.sibuk=!0,this.galat=null,this.gambar();try{let a;try{a=await Y(\"getUploadFormData\",i),this.lama=!1}catch(t){if(\"fungsi-tak-ada\"!==t.jenis)throw t;const s=await Promise.all([Y(\"getAssignableUsers\",i),Y(\"getEquipmentReadyDefaults\",i)]);a={users:s[0],equipment:s[1],lokasiAktif:null,modeAssignment:\"legacy\",facilityId:Q.user.facilityId,facilityName:Q.user.facilityName},this.lama=!0}if(t!==this.req)return;this.form=a,this.rt=Number(a.equipment.reachTruck)||0,this.tangga=Number(a.equipment.tangga)||0;const s=J.aktif?null:d.get(this.kunciPilih());this.pilih={},s&&s.length&&a.users.forEach(a=>{s.indexOf(a.username)>=0&&(this.pilih[a.username]=!0)}),this.adaTerakhir=Object.keys(this.pilih).length>0}catch(a){if(t!==this.req)return;this.galat=a}finally{t===this.req&&(this.sibuk=!1,\"upload\"===ma.aktif()&&this.gambar())}},terpilih(){return this.form.users.filter(a=>this.pilih[a.username])},peranUser:a=>String(a.role||\"\").trim().toLowerCase(),gambar(){const a=i(\"scr-upload\"),s=this.form,e=u`<button type=\"button\" class=\"iconbtn${this.sibuk?\" is-putar\":\"\"}\" data-aksi=\"upload-segar\" aria-label=\"Perbarui daftar petugas\">${I(\"segar\")}</button>`;if(!s)return void o(a,u`${va(\"Upload Data\",\"\",e)}${this.galat?fa(this.galat,\"upload-segar\"):ba(4,96)}`);if(s.tanpaFacility)return void o(a,u`${va(\"Upload Data\",\"\",e)}<section class=\"card\">${ka(\"gedung\",\"Akun ini belum punya facility\",\"Tugas dibagi per facility. Minta pemegang akses Config memasukkan akun Anda ke facility lebih dulu.\")}</section>`);const n=this.terpilih(),l=\"matrix\"===s.modeAssignment,r=n.filter(a=>\"storing\"!==this.peranUser(a)).length,d={};s.users.forEach(a=>{const t=this.peranUser(a);d[t]||(d[t]={total:0,pilih:0}),d[t].total++,this.pilih[a.username]&&d[t].pilih++});const p=Object.keys(d).sort((a,t)=>(Fa.indexOf(a)<0?99:Fa.indexOf(a))-(Fa.indexOf(t)<0?99:Fa.indexOf(t))),h=(a,t,i,s)=>{const e=this[a],n=e?e.sibuk?u`<span class=\"up-berkas__ket\">Membaca ${e.nama}…</span>`:e.galat?u`<span class=\"up-berkas__ket up-berkas__ket--galat\">${e.galat}</span>`:u`<span class=\"up-berkas__ket up-berkas__ket--ok\">${e.nama}</span><span class=\"up-berkas__ket\">${e.ringkas}</span>`:u`<span class=\"up-berkas__ket\">${s}</span>`;return u`<label class=\"up-berkas${!e||e.galat||e.sibuk?\"\":\" is-ok\"}${e&&e.galat?\" is-galat\":\"\"}\">\n        <input type=\"file\" class=\"sr\" accept=\".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv\" data-berkas=\"${a}\">\n        <span class=\"up-berkas__ic\">${e&&e.sibuk?u`<span class=\"putar\"></span>`:I(e&&!e.galat?\"cek\":\"berkas\")}</span>\n        <span class=\"up-berkas__isi\"><span class=\"up-berkas__judul\">${t}. ${i}</span>${n}</span>\n        <span class=\"up-berkas__aksi\">${e?\"Ganti\":\"Pilih file\"}</span></label>`};o(a,u`${va(\"Upload Data\",s.facilityName?\"Bagi tugas cycle count untuk \"+s.facilityName:\"\",e)}\n      ${0===s.lokasiAktif?u`<div class=\"info info--merah\" style=\"margin-bottom:12px\">${I(\"awas\")}<div><b>Daftar lokasi aktif masih kosong.</b> Tanpa daftar ini semua baris akan dilewati. ${Z.config()?u`<button type=\"button\" class=\"tautan\" data-aksi=\"upload-lokasi\">Impor lokasi aktif</button>`:\"Minta pemegang akses Config mengimpornya.\"}</div></div>`:\"\"}\n      <div class=\"tumpuk\">\n        ${h(\"transaksi\",1,\"Data transaksi\",\"File ekspor dari WMS, .xlsx atau .csv\")}\n        ${h(\"stok\",2,\"Stock by location\",\"Stok per lokasi, .xlsx atau .csv\")}\n      </div>\n      ${J.aktif?u`<button type=\"button\" class=\"btn btn--teks\" data-aksi=\"upload-contoh\" style=\"margin:6px auto 0\">Pakai file contoh</button>`:\"\"}\n      <h2 class=\"bag\"><span>3. Petugas yang bertugas</span><small>${c(n.length)} dari ${c(s.users.length)}</small></h2>\n      <section class=\"card up-petugas\">\n        ${s.users.length?u`\n        <div class=\"chips up-peran\" role=\"group\" aria-label=\"Pilih per peran\">${p.map(a=>u`<button type=\"button\" class=\"chip${d[a].pilih===d[a].total?\" is-on\":\"\"}\" data-aksi=\"upload-peran\" data-peran=\"${a}\">${j(a)} <b>${d[a].pilih}/${d[a].total}</b></button>`)}</div>\n        <div class=\"up-petugas__baris\">\n          <button type=\"button\" class=\"tautan\" data-aksi=\"upload-semua\" data-v=\"1\">Pilih semua</button>\n          <button type=\"button\" class=\"tautan\" data-aksi=\"upload-semua\" data-v=\"0\">Kosongkan</button>\n          <button type=\"button\" class=\"tautan up-petugas__buka\" data-aksi=\"upload-daftar\">${this.terbuka?\"Tutup daftar\":\"Lihat daftar\"}</button>\n        </div>\n        ${this.adaTerakhir&&!this.terbuka?u`<p class=\"ket\" style=\"margin-top:2px\">Pilihan terakhir dipakai lagi. Ubah bila tim hari ini berbeda.</p>`:\"\"}\n        ${this.terbuka?u`<div class=\"up-daftar\">${s.users.map(a=>u`<label class=\"row\"><input type=\"checkbox\" class=\"cek\" data-user=\"${a.username}\" ${this.pilih[a.username]?\"checked\":\"\"}><span class=\"row__isi\"><span class=\"row__t\">${a.username}</span></span><span class=\"pill\">${j(a.role)}</span></label>`)}</div>`:\"\"}`:u`<p class=\"redup\">Belum ada petugas di facility ini. Tambahkan lewat Config.</p>`}\n      </section>\n      ${l?u`<div class=\"info info--biru\" style=\"margin-top:12px\">${I(\"info\")}<div>Pembagian tugas mengikuti <b>Mode Matrix</b> (peran per grup alat). Jumlah alat bantu tidak perlu diisi.</div></div>`:u`\n      <h2 class=\"bag\"><span>4. Alat bantu siap hari ini</span></h2>\n      <section class=\"card\">\n        <div class=\"up-alat\">\n          <div><span class=\"lbl\">Reach truck</span>${this.langkah(\"rt\",this.rt)}</div>\n          <div><span class=\"lbl\">Tangga pesawat</span>${this.langkah(\"tangga\",this.tangga)}</div>\n        </div>\n        <p class=\"ket\">Petugas Storing mengerjakan Level 5-6 dengan reach truck. Petugas lain mengerjakan Level 3-4 sebanyak jumlah tangga; sisanya Level 1-2.</p>\n        ${this.tangga>r&&n.length?u`<div class=\"pesan pesan--galat\">Tangga (${this.tangga}) melebihi petugas terpilih di luar Storing (${r}).</div>`:\"\"}\n      </section>`}\n      <div class=\"pesan pesan--galat\" id=\"upPesan\" role=\"alert\" style=\"margin-top:12px\"></div>\n      <button type=\"button\" class=\"btn up-proses\" id=\"upProses\" data-aksi=\"upload-proses\">Proses dan bagi tugas</button>\n    `),t(\"input[data-berkas]\",a).forEach(a=>a.addEventListener(\"change\",()=>{a.files&&a.files[0]&&this.pilihBerkas(a.dataset.berkas,a.files[0])})),t(\"input[data-user]\",a).forEach(a=>a.addEventListener(\"change\",()=>{a.checked?this.pilih[a.dataset.user]=!0:delete this.pilih[a.dataset.user],this.adaTerakhir=!1,this.gambarSimpanGulir()}))},langkah:(a,t)=>u`<div class=\"langkah\"><button type=\"button\" data-aksi=\"upload-langkah\" data-n=\"${a}\" data-d=\"-1\" aria-label=\"Kurangi\">${I(\"kurang\")}</button><output>${c(t)}</output><button type=\"button\" data-aksi=\"upload-langkah\" data-n=\"${a}\" data-d=\"1\" aria-label=\"Tambah\">${I(\"tambah\")}</button></div>`,gambarSimpanGulir(){const a=window.scrollY;this.gambar(),window.scrollTo(0,a)},ringkasTransaksi(a,t){let i=0,s=0;const e={};if(t.forEach(a=>{const t=a.type.toLowerCase();\"move\"===t?(i++,e[a.lokasiTujuan+\"||\"+a.article]=1):\"picking\"===t&&(s++,e[a.lokasiAwal+\"||\"+a.article]=1)}),!i&&!s)throw new Error(\"Tidak ada baris bertipe Move atau Picking di file ini.\");return{nama:a,rows:t,kunci:e,ringkas:c(t.length)+\" baris: \"+c(i)+\" Move, \"+c(s)+\" Picking\"+(t.length-i-s?\", \"+c(t.length-i-s)+\" tipe lain dilewati\":\"\")}},async pilihBerkas(a,t){this[a]={nama:t.name,sibuk:!0},this.gambarSimpanGulir();try{if(\"transaksi\"===a)this.transaksi=this.ringkasTransaksi(t.name,await za(t,Ka,[\"addWho\"]));else{const a=await za(t,qa,null);this.stok={nama:t.name,rows:a,ringkas:c(a.length)+\" baris stok\"}}}catch(i){this[a]={nama:t.name,galat:i.message}}this.periksaStok(),this.gambarSimpanGulir()},periksaStok(){const a=this.transaksi,t=this.stok;if(!t||!t.rows)return;if(delete t.galat,!a||!a.kunci)return;const i=t.rows.find(t=>t.qtySalah&&a.kunci[t.lokasi+\"||\"+t.article]);i&&(t.galat=\"Baris \"+i.baris+': qty \"'+i.qty+'\" bukan angka. Perbaiki file lalu pilih lagi.')},contoh(){const a=J._mesin.contoh();this.transaksi=this.ringkasTransaksi(\"transaksi-contoh.xlsx\",a.transaksi),this.stok={nama:\"stok-contoh.xlsx\",rows:a.stok,ringkas:c(a.stok.length)+\" baris stok\"},this.terpilih().length||this.form.users.forEach(a=>{\"inventory\"!==this.peranUser(a)&&\"admin\"!==this.peranUser(a)&&(this.pilih[a.username]=!0)}),this.gambarSimpanGulir()},async proses(){const a=this.form,t=a=>{i(\"upPesan\").textContent=a},s=this.terpilih(),e=(a,t)=>a?a.sibuk?\"File \"+t+\" masih dibaca. Tunggu sebentar.\":a.galat?\"File \"+t+\" bermasalah. Pilih file lain.\":\"\":\"Pilih file \"+t+\" dulu.\";this.periksaStok();const n=e(this.transaksi,\"data transaksi\")||e(this.stok,\"stock by location\");if(n)return t(n);if(!s.length)return t(\"Pilih minimal satu petugas.\");const l=\"matrix\"===a.modeAssignment,r=s.filter(a=>\"storing\"!==this.peranUser(a)).length;if(!l&&this.tangga>r)return t(\"Jumlah tangga melebihi petugas terpilih di luar Storing (\"+r+\").\");t(\"\");const u=this.transaksi.rows.map(a=>[a.type,a.article,a.description,a.lokasiAwal,a.lokasiTujuan,a.qty,a.addWho]);let o=this.stok.rows.filter(a=>this.transaksi.kunci[a.lokasi+\"||\"+a.article]).map(a=>[a.lokasi,a.article,a.qty]);o.length||(o=this.stok.rows.slice(0,1).map(a=>[a.lokasi,a.article,a.qty]));const p=s.map(a=>a.username);J.aktif||d.set(this.kunciPilih(),p);const h=\"job_\"+Date.now()+\"_\"+Math.floor(1e5*Math.random());la.tampil(\"Memproses dan membagi tugas\",c(u.length)+\" baris transaksi\"),la.pantau(h);try{const a=await Y(\"importRawData\",u,o,p,Q.user.username,this.rt,this.tangga,h);if(la.persen(100),await D(250),la.sembunyi(),!a.success)return t(a.message),void i(\"upPesan\").scrollIntoView({block:\"center\"});this.transaksi=null,this.stok=null,ja.at=0,Aa.cycle.at=0,this.gambar(),this.hasil(a,p.length)}catch(a){la.sembunyi(),t(a.message),i(\"upPesan\").scrollIntoView({block:\"center\"})}},hasil(a,t){const i=a.peringatan||(a.warning?[a.warning]:[]),s=a.catatan||(a.message?[String(a.message).split(\"⚠\")[0].trim()]:[]),e=(a,t,i)=>t>0?u`<div class=\"row\"><span class=\"row__isi\"><span class=\"row__t row__t--lipat\" style=\"font-weight:500\">${a}</span></span><span class=\"row__angka\" style=\"${i?\"color:var(--\"+i+\"-teks)\":\"\"}\">${c(t)}</span></div>`:\"\";ea.buka({judul:a.ditugaskan?\"Tugas sudah dibagi\":\"Tidak ada item baru\",isi:u`<div class=\"up-hasil\"><b class=\"angka\">${c(a.ditugaskan||0)}</b><span>item baru dibagi ke ${c(t)} petugas</span></div>\n        ${a.merged||a.skippedInactive||a.skippedBlankLokasi||a.skippedOtherType?u`<div class=\"list\" style=\"margin-bottom:14px\">\n          ${e(\"Digabung ke item yang masih menunggu\",a.merged)}\n          ${e(\"Dilewati: lokasi tidak ada di daftar lokasi aktif\",a.skippedInactive)}\n          ${e(\"Dilewati: kolom lokasi kosong\",a.skippedBlankLokasi,\"jingga\")}\n          ${e(\"Dilewati: tipe selain Move dan Picking\",a.skippedOtherType)}\n        </div>`:\"\"}\n        ${i.length?u`<div class=\"info info--jingga\" style=\"margin-bottom:12px\">${I(\"awas\")}<div><b>Ada item yang tidak dibagi.</b> ${i.join(\" \")} Rinciannya tercatat di sheet Log_Unassigned.</div></div>`:\"\"}\n        ${s.map(a=>u`<div class=\"info\" style=\"margin-bottom:10px\">${I(\"info\")}<div>${a}</div></div>`)}`,kaki:U.panel?u`<button type=\"button\" class=\"btn\" data-tutup>${i.length?\"Mengerti\":\"Selesai\"}</button>`:u`<button type=\"button\" class=\"btn\" data-aksi=\"ke\" data-tab=\"home\">${i.length?\"Mengerti, kembali ke Home\":\"Kembali ke Home\"}</button>`})}};ca.upload={masuk(){Ja.gambar(),Ja.muat(!1)}},ia[\"upload-segar\"]=()=>Ja.muat(!0),ia[\"upload-daftar\"]=()=>{Ja.terbuka=!Ja.terbuka,Ja.gambarSimpanGulir()},ia[\"upload-semua\"]=a=>{Ja.pilih={},\"1\"===a.dataset.v&&Ja.form.users.forEach(a=>{Ja.pilih[a.username]=!0}),Ja.adaTerakhir=!1,Ja.gambarSimpanGulir()},ia[\"upload-peran\"]=a=>{const t=a.dataset.peran,i=Ja.form.users.filter(a=>Ja.peranUser(a)===t),s=i.every(a=>Ja.pilih[a.username]);i.forEach(a=>{s?delete Ja.pilih[a.username]:Ja.pilih[a.username]=!0}),Ja.adaTerakhir=!1,Ja.gambarSimpanGulir()},ia[\"upload-langkah\"]=a=>{const t=a.dataset.n;Ja[t]=Math.max(0,Math.min(99,Ja[t]+Number(a.dataset.d))),Ja.gambarSimpanGulir()},ia[\"upload-proses\"]=()=>Ja.proses(),ia[\"upload-contoh\"]=()=>Ja.contoh(),ia[\"upload-lokasi\"]=()=>{Xa.bukaLokasi(Ja.form.facilityId,Ja.form.facilityName,()=>Ja.muat(!0))};const Wa={\"Salah Picking / Move\":\"var(--merah)\",\"Salah Putaway\":\"var(--jingga)\",\"Human Error Cycle\":\"var(--kuning-tepi)\",\"Master Data / Location\":\"var(--biru)\",\"Tidak diketahui\":\"var(--abu-2)\"},Ga={\"Salah Picking / Move\":[\"orang\",\"Coaching petugas picking dan move\",\"Dampingi petugas dengan kesalahan picking atau move terbanyak.\"],\"Salah Putaway\":[\"kotak\",\"Tinjau proses putaway\",\"Periksa penempatan barang di lokasi yang sering bermasalah.\"],\"Human Error Cycle\":[\"daftar\",\"Tinjau cara menghitung\",\"Ulangi pelatihan ketelitian hitung untuk petugas cycle.\"],\"Master Data / Location\":[\"atur\",\"Perbaiki master data\",\"Periksa data lokasi dan stok yang tidak sesuai.\"],\"Tidak diketahui\":[\"cari\",\"Selesaikan verifikasi\",\"Tindak lanjuti selisih yang belum punya kategori penyebab.\"]},Ya={tampilan:\"dashboard\",dash:null,prod:null,an:null,reset(){this.tampilan=\"dashboard\",this.dash={mode:\"daily\",tanggal:\"\",bulan:\"\",data:null,untuk:\"\",sibuk:!1,galat:null,req:0,lb:\"petugas\",backlog:null,backlogBasi:!1,detail:null,detailSibuk:!1,cari:\"\"},this.prod={tanggal:\"\",data:null,untuk:\"\",sibuk:!1,galat:null,req:0},this.an={dari:\"\",sampai:\"\",data:null,untuk:\"\",sibuk:!1,galat:null,req:0,kontrib:\"userCycle\"}},async ambil(a,t,i,s){const e=this[a],n=++e.req;e.untuk!==t&&(e.data=null,e.untuk=t),e.sibuk=!0,e.galat=null,\"report\"===ma.aktif()&&this.gambar();let l=null,r=null;try{l=await i()}catch(a){r=a}this[a]===e&&e.req===n&&(e.sibuk=!1,r?e.galat=r:s(l),\"report\"===ma.aktif()&&this.keadaan()===e&&this.gambar())},keadaan(){const a=this.tampilan;return\"dashboard\"===a?this.dash:\"productivity\"===a?this.prod:this.an},belumSegar:a=>a.galat&&a.data?u`<div class=\"info info--jingga\" style=\"margin-bottom:12px\">${I(\"awan_mati\")}<div>Belum bisa diperbarui: ${a.galat.message}</div></div>`:\"\",awal(){const a=v();this.dash.tanggal||(this.dash.tanggal=a,this.dash.bulan=a.slice(0,7)),this.prod.tanggal||(this.prod.tanggal=a),this.an.sampai||(this.an.sampai=a,this.an.dari=y(a,-6))},gambar(){this.awal();const a=i(\"scr-report\"),t=this.tampilan,s=u`<button type=\"button\" class=\"iconbtn${this.keadaan().sibuk?\" is-putar\":\"\"}\" data-aksi=\"report-segar\" aria-label=\"Perbarui\">${I(\"segar\")}</button>`;o(a,u`${va(\"Report\",\"\",s)}\n      <div class=\"seg rp-seg\" role=\"tablist\">${[[\"dashboard\",\"Dashboard\"],[\"productivity\",\"Productivity\"],[\"analytics\",\"Analytics\"]].map(a=>u`<button type=\"button\" role=\"tab\" class=\"${t===a[0]?\"is-on\":\"\"}\" data-aksi=\"report-tampilan\" data-v=\"${a[0]}\">${a[1]}</button>`)}</div>\n      <div id=\"rpIsi\">${\"dashboard\"===t?this.dashHtml():\"productivity\"===t?this.prodHtml():this.anHtml()}</div>`),this.pasang()},pasang(){const a=this.tampilan,t=(a,t)=>{const s=i(a);s&&s.addEventListener(\"change\",()=>{s.value&&t(s.value)})};if(\"dashboard\"===a){t(\"rpTanggal\",a=>{this.dash.tanggal=a,this.muatDash()}),t(\"rpBulan\",a=>{this.dash.bulan=a,this.muatDash()});const a=this.dash.data;a&&i(\"rpTren\")&&Sa.tren(i(\"rpTren\"),a.trend.map(a=>Object.assign({},a,{label:\"monthly\"===this.dash.mode?x(a.period):_(a.period),judul:\"monthly\"===this.dash.mode?x(a.period):w(a.period)})),i(\"rpTrenInfo\"));const s=i(\"rpCari\");s&&s.addEventListener(\"input\",E(()=>{this.dash.cari=s.value,o(i(\"rpMasalah\"),this.masalahHtml())},160))}else if(\"productivity\"===a){t(\"rpProdTanggal\",a=>{this.prod.tanggal=a,this.muatProd()});const a=this.prod.data;a&&a.perJam.length&&i(\"rpJam\")&&Sa.batang(i(\"rpJam\"),a.perJam.map(a=>({label:k(a.jam),nilai:a.total})))}else{t(\"rpDari\",a=>{this.an.dari=a,this.muatAn()}),t(\"rpSampai\",a=>{this.an.sampai=a,this.muatAn()});const a=this.an.data;a&&a.trend.length>1&&i(\"rpAnTren\")&&Sa.garis(i(\"rpAnTren\"),a.trend.map(a=>({label:_(a.tanggal),akurasi:a.akurasi})))}},muatIni(){const a=this.tampilan;return\"dashboard\"===a?this.muatDash():\"productivity\"===a?this.muatProd():this.muatAn()},segar(){const a=this.keadaan();a.sibuk||(a===this.dash&&(a.backlogBasi=!0),this.muatIni())},buka(){const a=this.keadaan();a.data||a.sibuk||this.muatIni()},muatDash(){const a=this.dash,t=\"daily\"===a.mode?a.tanggal:a.bulan,i=Q.user.username,s=a.mode+\":\"+t,e=a.untuk!==s,n=!e&&!!a.detail;return e&&(a.detail=null,a.cari=\"\"),a.detailSibuk=!1,this.ambil(\"dash\",s,()=>Promise.all([Y(\"getDashboardData\",a.mode,t,\"monthly\"===a.mode?6:7,i),a.backlog&&!a.backlogBasi?a.backlog:Y(\"getPendingBacklog\",i)]),t=>{a.data=t[0],a.backlog=t[1]||[],a.backlogBasi=!1,n&&this.muatMasalah(!0)})},dashHtml(){const a=this.dash,t=a.data,i=u`<div class=\"rp-periode\"><div class=\"seg\"><button type=\"button\" class=\"${\"daily\"===a.mode?\"is-on\":\"\"}\" data-aksi=\"report-mode\" data-v=\"daily\">Harian</button><button type=\"button\" class=\"${\"monthly\"===a.mode?\"is-on\":\"\"}\" data-aksi=\"report-mode\" data-v=\"monthly\">Bulanan</button></div>\n      ${\"daily\"===a.mode?u`<input type=\"date\" class=\"field field--kecil\" id=\"rpTanggal\" value=\"${a.tanggal}\" max=\"${v()}\" aria-label=\"Tanggal\">`:u`<input type=\"month\" class=\"field field--kecil\" id=\"rpBulan\" value=\"${a.bulan}\" aria-label=\"Bulan\">`}</div>`;if(!t)return u`${i}${a.galat?fa(a.galat,\"report-segar\"):ba(3,130)}`;const s=t.summary.kpi,e=Math.max(0,s.total-s.hit-s.discrepancy),n=s.total?Math.round(s.hit/s.total*1e3)/10:0,l=t.trend.filter(a=>a.total>0),r=l.length?Math.round(l.reduce((a,t)=>a+t.akurasi,0)/l.length*10)/10:null,o=l.length>=2?Math.round(10*(l[l.length-1].akurasi-l[l.length-2].akurasi))/10:null,d=\"validator\"===a.lb?t.summary.leaderboardValidator:t.summary.leaderboard;return u`${i}\n      ${this.belumSegar(a)}\n      ${s.total?u`\n      <section class=\"card rp-akurasi\">\n        <div class=\"rp-akurasi__utama\"><b class=\"angka\" style=\"color:${n>=99?\"var(--hijau-teks)\":\"var(--tinta)\"}\">${m(n)}%</b><span>Akurasi ${\"daily\"===a.mode?_(a.tanggal):x(a.bulan)}</span></div>\n        <div class=\"rp-tiga\"><div><b class=\"angka\">${c(s.total)}</b><span>Total hitung</span></div><div><b class=\"angka\" style=\"color:var(--hijau-teks)\">${c(s.hit)}</b><span>Hit</span></div><div><b class=\"angka\" style=\"color:var(--merah-teks)\">${c(s.discrepancy)}</b><span>Discrepancy</span></div></div>\n        ${e?u`<div class=\"ket\">${c(e)} item masih menunggu validasi dan belum dihitung sebagai hit atau discrepancy.</div>`:\"\"}\n      </section>`:u`<section class=\"card\">${ka(\"grafik\",\"Belum ada hitungan\",\"daily\"===a.mode?\"Tidak ada data cycle count pada tanggal ini.\":\"Tidak ada data cycle count pada bulan ini.\")}</section>`}\n      <section class=\"card hm-kartu\">\n        <h2 class=\"card__judul\">Tren akurasi</h2><p class=\"card__ket\">${\"monthly\"===a.mode?\"6 bulan terakhir.\":\"7 hari terakhir.\"} Ketuk batang untuk melihat angkanya.</p>\n        <svg class=\"grafik\" id=\"rpTren\" style=\"margin-top:8px\"></svg>\n        <div class=\"rp-info\" id=\"rpTrenInfo\"></div>\n        <div class=\"legenda\" style=\"margin-top:10px\"><span><i style=\"background:var(--hijau)\"></i>Hit</span><span><i style=\"background:var(--merah)\"></i>Discrepancy</span><span><i style=\"background:var(--tinta);border-radius:50%\"></i>Akurasi</span></div>\n        ${null!=r?u`<div class=\"rp-tiga rp-tiga--garis\"><div><b class=\"angka\">${m(r)}%</b><span>Rata-rata akurasi</span></div><div><b class=\"angka\">${c(t.trend.reduce((a,t)=>a+t.discrepancy,0))}</b><span>Total discrepancy</span></div>\n          <div><b class=\"angka\" style=\"color:${o>0?\"var(--hijau-teks)\":o<0?\"var(--merah-teks)\":\"inherit\"}\">${null==o?\"-\":(o>0?\"+\":\"\")+m(o)}</b><span>Poin dari periode lalu</span></div></div>`:\"\"}\n      </section>\n      ${a.backlog&&a.backlog.length?u`<section class=\"card hm-kartu\"><h2 class=\"card__judul\">Outstanding Cycle</h2><p class=\"card__ket\">Tugas yang belum dihitung, per tanggal upload. Ketuk untuk melihat petugasnya.</p>\n        <div class=\"rp-daftar\">${a.backlog.map(a=>u`<button type=\"button\" class=\"rp-baris\" data-aksi=\"report-backlog\" data-tgl=\"${a.tanggal}\"><span>${w(String(a.tanggal).slice(0,10))}</span><span class=\"pill pill--jingga\">${c(a.pending)} belum</span>${I(\"kanan\",\"ic--kecil\")}</button>`)}</div></section>`:\"\"}\n      ${s.total?u`<section class=\"card hm-kartu\">\n        <h2 class=\"card__judul\">Ranking akurasi</h2>\n        <div class=\"seg\" style=\"margin:10px 0 4px\"><button type=\"button\" class=\"${\"petugas\"===a.lb?\"is-on\":\"\"}\" data-aksi=\"report-lb\" data-v=\"petugas\">Petugas</button><button type=\"button\" class=\"${\"validator\"===a.lb?\"is-on\":\"\"}\" data-aksi=\"report-lb\" data-v=\"validator\">Validator</button></div>\n        <p class=\"card__ket\">${\"petugas\"===a.lb?\"Dari item yang hasilnya sudah final. Ketuk nama untuk ringkasannya.\":\"Dari item yang verifikasinya sudah selesai. Validator dianggap salah bila hitung ulangnya yang keliru.\"}</p>\n        ${d&&d.length?u`<div class=\"rp-daftar\">${d.map((t,i)=>this.barisRank(t,i,\"petugas\"===a.lb))}</div>`:u`<p class=\"redup\" style=\"padding:14px 0 4px\">Belum ada data untuk ranking ini.</p>`}\n      </section>`:\"\"}\n      ${t.errorAnalysis&&t.errorAnalysis.length?u`<section class=\"card hm-kartu\"><h2 class=\"card__judul\">Analisis kesalahan</h2><p class=\"card__ket\">Selisih per kategori penyebab. Ketuk untuk melihat itemnya.</p>\n        <div class=\"rp-daftar\">${t.errorAnalysis.map(a=>u`<button type=\"button\" class=\"rp-batang\" data-aksi=\"report-kategori\" data-alasan=\"${a.alasan}\"><span class=\"rp-batang__atas\"><b>${a.alasan}</b><span>${c(a.jumlah)} item, ${m(a.persen)}%</span></span><span class=\"bar bar--tipis bar--merah\"><i style=\"width:${a.persen}%\"></i></span></button>`)}</div></section>`:\"\"}\n      ${s.total?u`<section class=\"card hm-kartu\"><h2 class=\"card__judul\">Item bermasalah</h2><p class=\"card__ket\">Item yang selisih, menunggu validasi, atau sedang diverifikasi pada periode ini.</p>\n        <div id=\"rpMasalah\">${this.masalahHtml()}</div></section>\n      <button type=\"button\" class=\"btn btn--teks\" data-aksi=\"report-ekspor\" style=\"margin:10px auto 0\">${I(\"unduh\",\"ic--kecil\")}Ekspor ke Excel</button>`:\"\"}`},barisRank(a,t,i){const s=u`<span class=\"rp-rank__no\">${t+1}</span><span class=\"rp-rank__isi\"><b>${a.nama}</b><span>${c(a.hit)} benar dari ${c(a.total)} item</span></span>\n      <span class=\"rp-rank__persen\" style=\"color:${a.akurasi>=99?\"var(--hijau-teks)\":\"var(--merah-teks)\"}\">${m(a.akurasi)}%</span>`;return i?u`<button type=\"button\" class=\"rp-rank\" data-aksi=\"report-user\" data-nama=\"${a.nama}\">${s}</button>`:u`<div class=\"rp-rank\">${s}</div>`},masalahHtml(){const a=this.dash;if(a.detailSibuk)return ga(\"Memuat item…\");if(!a.detail)return u`<button type=\"button\" class=\"btn btn--tenang\" data-aksi=\"report-masalah\" style=\"margin-top:12px\">Tampilkan item bermasalah</button>`;if(!a.detail.length)return u`<p class=\"redup\" style=\"padding-top:12px\">Tidak ada item bermasalah pada periode ini.</p>`;const t=(a.cari||\"\").trim().toLowerCase(),i=a.detail.filter(a=>!t||B([a.lokasi,a.article,a.description,a.namaPetugas,a.status].join(\" \"),t));return u`<div class=\"cari\" style=\"margin:12px 0 10px\">${I(\"cari\",\"ic--kecil\")}<input class=\"field\" type=\"search\" id=\"rpCari\" placeholder=\"Cari lokasi, SKU, atau nama\" value=\"${a.cari||\"\"}\" autocomplete=\"off\"></div>\n      <div class=\"rp-daftar\">${i.slice(0,60).map(a=>{return u`<div class=\"rp-item\"><div class=\"vf-row__atas\">${C(a.lokasi)}<span class=\"row__t\">${a.article}</span><span class=\"pill vf-selisih\">${\"\"===a.selisih||null==a.selisih?\"-\":p(a.selisih)}</span></div>\n        <div class=\"row__s\">${a.description}</div><div class=\"rp-item__bawah\"><span class=\"redup\">${a.namaPetugas}</span><span class=\"pill pill--${t=a.status,/Selesai/.test(t)?\"hijau\":/Verifikasi/.test(t)?\"merah\":/Menunggu/.test(t)?\"ungu\":\"jingga\"}\">${a.status}</span></div></div>`;var t})}</div>\n      ${i.length>60?u`<p class=\"redup\" style=\"text-align:center;margin-top:10px\">Dan ${c(i.length-60)} item lagi. Persempit dengan pencarian atau ekspor ke Excel.</p>`:\"\"}${i.length?\"\":u`<p class=\"redup\" style=\"padding:10px 0\">Tidak ada yang cocok.</p>`}`},async muatMasalah(a){const t=this.dash,s=t.req;a||(t.detailSibuk=!0,o(i(\"rpMasalah\"),this.masalahHtml()));let e=null,n=null;try{e=await Y(\"getProblemItemsDetail\",t.mode,\"daily\"===t.mode?t.tanggal:t.bulan,Q.user.username)||[]}catch(a){n=a}if(this.dash===t&&t.req===s&&(t.detailSibuk=!1,n?a||sa.galat(n):t.detail=e,\"report\"===ma.aktif()&&\"dashboard\"===this.tampilan)){const a=window.scrollY;this.gambar(),window.scrollTo(0,a)}},async bukaBacklog(a){const t=ea.buka({judul:\"Belum di cycle\",isi:ga()});try{const i=await Y(\"getBacklogDetailByDate\",a,Q.user.username),s=Math.max(1,Math.max.apply(null,i.map(a=>a.pending).concat([1])));o(t.isi,i.length?u`<p class=\"redup\" style=\"margin-bottom:10px\">Tugas dari upload ${w(String(a).slice(0,10))}.</p>${i.map(a=>u`<div class=\"orang\"><div class=\"orang__atas\"><b>${a.namaPetugas}</b><span class=\"angka\">${c(a.pending)}</span></div><div class=\"bar bar--tipis bar--jingga\"><i style=\"width:${a.pending/s*100}%\"></i></div><div class=\"orang__ket\">dari ${c(a.total)} tugas</div></div>`)}`:ka(\"cek\",\"Semua sudah selesai\",\"\"))}catch(a){o(t.isi,u`<p class=\"pesan pesan--galat\">${a.message}</p>`)}},async bukaUser(a){const t=this.dash,i=ea.buka({judul:a,isi:ga()});try{const s=await Y(\"getUserDashboardDetail\",a,t.mode,\"daily\"===t.mode?t.tanggal:t.bulan,Q.user.username),e=s.summary,n=s.investigasi;o(i.isi,u`<div class=\"rp-empat\"><div><b class=\"angka\">${c(e.total)}</b><span>Total item</span></div><div><b class=\"angka\" style=\"color:var(--merah-teks)\">${c(e.kesalahanHitung)}</b><span>Salah hitung</span></div>\n        <div><b class=\"angka\" style=\"color:var(--hijau-teks)\">${m(e.akurasi)}%</b><span>Akurasi</span></div><div><b class=\"angka\">${m(e.persenDiscrepancy)}%</b><span>Discrepancy</span></div></div>\n        ${n&&n.totalSelesai?u`<h3 class=\"bag\" style=\"margin-top:18px\"><span>Verifikasi diselesaikan</span></h3><div class=\"rp-empat rp-empat--dua\"><div><b class=\"angka\">${c(n.totalSelesai)}</b><span>Task selesai</span></div><div><b class=\"angka\">${m(n.avgDurasiJam)} jam</b><span>Rata-rata buka sampai tutup</span></div></div>\n          <p class=\"ket\">Durasi dihitung dari task dibuka sampai ditutup, termasuk waktu antre. Jangan dipakai sebagai satu-satunya ukuran kecepatan.</p>`:\"\"}`)}catch(a){o(i.isi,u`<p class=\"pesan pesan--galat\">${a.message}</p>`)}},async bukaKategori(a){const t=this.dash,i=ea.buka({judul:a,isi:ga()});try{const s=await Y(\"getErrorAnalysisDetail\",a,t.mode,\"daily\"===t.mode?t.tanggal:t.bulan,Q.user.username);o(i.isi,s.length?u`<div class=\"list\">${s.map(a=>u`<div class=\"row\">${C(a.lokasi)}<div class=\"row__isi\"><div class=\"row__t\">${a.article}</div><div class=\"row__s\">${a.description}</div><div class=\"row__s\">Sistem ${c(a.qtySystem)}, hitung ${c(a.qtyCount)}</div><div class=\"row__s\">Dihitung oleh ${a.namaPetugas}</div></div></div>`)}</div>`:ka(\"daftar\",\"Tidak ada item\",\"\"))}catch(a){o(i.isi,u`<p class=\"pesan pesan--galat\">${a.message}</p>`)}},eksporDash(){const a=this.dash,t=a.data;if(!t)return;const i=a=>(a||[]).map(a=>({Nama:a.nama,\"Total Item\":a.total,\"Hit (Benar)\":a.hit,\"Akurasi (%)\":a.akurasi,\"Discrepancy (%)\":a.persenDiscrepancy})),s=[{nama:\"Ranking Petugas\",baris:i(t.summary.leaderboard)},{nama:\"Ranking Validator\",baris:i(t.summary.leaderboardValidator)},{nama:\"Analisis Kesalahan\",baris:(t.errorAnalysis||[]).map(a=>({Kategori:a.alasan,Jumlah:a.jumlah,\"Persen (%)\":a.persen}))}];a.detail&&s.push({nama:\"Item Bermasalah\",baris:a.detail.map(a=>({Tanggal:a.tanggal,Lokasi:a.lokasi,SKU:a.article,Deskripsi:a.description,Selisih:a.selisih,Petugas:a.namaPetugas,Status:a.status}))}),wa(\"Dashboard_CycleCount_\"+(\"monthly\"===a.mode?\"Bulanan_\"+a.bulan:\"Harian_\"+a.tanggal)+\".xlsx\",s)},muatProd(){const a=this.prod,t=a.tanggal;return this.ambil(\"prod\",t,()=>Y(\"getProductivity\",t,Q.user.username,-(new Date).getTimezoneOffset()),t=>{a.data=t})},prodHtml(){const a=this.prod,t=a.data,i=u`<div class=\"rp-periode\"><input type=\"date\" class=\"field field--kecil\" id=\"rpProdTanggal\" value=\"${a.tanggal}\" max=\"${v()}\" aria-label=\"Tanggal\"></div>`;if(!t)return u`${i}${a.galat?\"fungsi-tak-ada\"===a.galat.jenis?u`<section class=\"card\">${ka(\"kilat\",\"Belum tersedia\",\"Laporan ini tersedia setelah admin aplikasi memperbarui sistem.\")}</section>`:fa(a.galat,\"report-segar\"):ba(3,120)}`;const s=t.ringkasan,e=t.petugas.filter(a=>a.total>0),n=Math.max(1,Math.max.apply(null,t.petugas.map(a=>a.total).concat([1])));return s.totalItem||s.sisa?u`${i}\n      ${this.belumSegar(a)}\n      <div class=\"rp-tiga rp-tiga--kartu\"><div><b class=\"angka\">${c(s.totalItem)}</b><span>Item dihitung</span></div><div><b class=\"angka\">${c(s.petugasAktif)}</b><span>Petugas aktif</span></div><div><b class=\"angka\">${m(s.perJam)}</b><span>Item per jam per orang</span></div></div>\n      ${t.perJam.length?u`<section class=\"card hm-kartu\"><h2 class=\"card__judul\">Hitungan per jam</h2><p class=\"card__ket\">Jumlah item yang dihitung tiap jam, semua petugas.</p><svg class=\"grafik\" id=\"rpJam\" style=\"margin-top:10px\"></svg></section>`:\"\"}\n      <section class=\"card hm-kartu\"><h2 class=\"card__judul\">Per petugas</h2><p class=\"card__ket\">Per jam dihitung dari jam yang benar-benar berisi hitungan, jadi istirahat tidak menurunkan angka.</p>\n        <div class=\"rp-daftar\">${t.petugas.map(a=>u`<div class=\"orang\"><div class=\"orang__atas\"><b>${a.nama}${a.role?u` <span class=\"pill\">${j(a.role)}</span>`:\"\"}</b><span class=\"angka\">${c(a.total)}</span></div>\n          <div class=\"bar bar--tipis bar--biru\"><i style=\"width:${a.total/n*100}%\"></i></div>\n          <div class=\"orang__ket\">${a.total?u`${$(a.mulai)} sampai ${$(a.akhir)}, ${m(a.perJam)} per jam${a.selisih?u`, ${c(a.selisih)} selisih awal`:\"\"}`:\"Belum mulai menghitung\"}${a.sisa?u` <span class=\"pill pill--jingga\">${c(a.sisa)} belum</span>`:\"\"}</div></div>`)}</div></section>\n      ${t.validator.length?u`<section class=\"card hm-kartu\"><h2 class=\"card__judul\">Validasi</h2><p class=\"card__ket\">${c(s.totalValidasi)} item divalidasi pada tanggal ini.</p>\n        <div class=\"rp-daftar\">${t.validator.map(a=>u`<div class=\"rp-rank\"><span class=\"rp-rank__isi\"><b>${a.nama}</b><span>${$(a.mulai)} sampai ${$(a.akhir)}, ${c(a.hit)} ternyata cocok</span></span><span class=\"rp-rank__persen\">${c(a.total)}</span></div>`)}</div></section>`:\"\"}\n      ${e.length?u`<button type=\"button\" class=\"btn btn--teks\" data-aksi=\"report-ekspor\" style=\"margin:10px auto 0\">${I(\"unduh\",\"ic--kecil\")}Ekspor ke Excel</button>`:\"\"}`:u`${i}${this.belumSegar(a)}<section class=\"card\">${ka(\"kilat\",\"Belum ada hitungan\",\"Tidak ada item yang dihitung pada \"+w(a.tanggal)+\".\")}</section>`},eksporProd(){const a=this.prod.data;a&&wa(\"Productivity_\"+a.tanggal+\".xlsx\",[{nama:\"Petugas\",baris:a.petugas.map(a=>({Nama:a.nama,Peran:j(a.role),\"Item Dihitung\":a.total,\"Cocok (awal)\":a.hit,\"Selisih (awal)\":a.selisih,Mulai:a.mulai?$(a.mulai):\"\",Terakhir:a.akhir?$(a.akhir):\"\",\"Jam Aktif\":a.jamAktif,\"Item per Jam\":a.perJam,\"Belum Dihitung\":a.sisa}))},{nama:\"Validator\",baris:a.validator.map(a=>({Nama:a.nama,\"Item Divalidasi\":a.total,\"Ternyata Cocok\":a.hit,Mulai:$(a.mulai),Terakhir:$(a.akhir)}))},{nama:\"Per Jam\",baris:a.perJam.map(a=>({Jam:k(a.jam)+\".00\",\"Item Dihitung\":a.total}))}])},muatAn(){const a=this.an;if(a.dari>a.sampai){const t=a.dari;a.dari=a.sampai,a.sampai=t}const t=a.dari,i=a.sampai;return this.ambil(\"an\",t+\":\"+i,()=>Y(\"getAnalyticsRootCauseData\",t,i,Q.user.username),t=>{a.data=t})},anHtml(){const a=this.an,t=a.data,i=u`<div class=\"rentang\"><label><span>Dari</span><input type=\"date\" class=\"field field--kecil\" id=\"rpDari\" value=\"${a.dari}\" max=\"${v()}\"></label><label><span>Sampai</span><input type=\"date\" class=\"field field--kecil\" id=\"rpSampai\" value=\"${a.sampai}\" max=\"${v()}\"></label></div>`;if(!t)return u`${i}${a.galat?fa(a.galat,\"report-segar\"):ba(3,130)}`;if(!t.kpi.total)return u`${i}${this.belumSegar(a)}<section class=\"card\">${ka(\"tren\",\"Belum ada data\",\"Tidak ada cycle count pada rentang tanggal ini.\")}</section>`;const s=t.rootCause||[],e=s.slice().sort((a,t)=>t.jumlah-a.jumlah).filter(a=>a.jumlah>0),n=t.contributors&&t.contributors[a.kontrib]||[],l=Math.max(1,Math.max.apply(null,n.map(a=>a.jumlah).concat([1])));return u`${i}\n      ${this.belumSegar(a)}\n      <div class=\"rp-tiga rp-tiga--kartu rp-tiga--dua\"><div><b class=\"angka\" style=\"color:var(--hijau-teks)\">${m(t.kpi.akurasi)}%</b><span>Akurasi</span></div><div><b class=\"angka\" style=\"color:var(--merah-teks)\">${m(t.kpi.errorPersen)}%</b><span>${c(t.kpi.discrepancy)} item discrepancy</span></div></div>\n      ${t.trend.length>1?u`<section class=\"card hm-kartu\"><h2 class=\"card__judul\">Tren akurasi</h2><p class=\"card__ket\">${c(t.trend.length)} hari dalam rentang terpilih.</p><svg class=\"grafik\" id=\"rpAnTren\" style=\"margin-top:10px\"></svg></section>`:\"\"}\n      <section class=\"card hm-kartu\"><h2 class=\"card__judul\">Akar masalah</h2><p class=\"card__ket\">Selisih dikelompokkan menurut penyebabnya. Ketuk untuk melihat itemnya.</p>\n        <div class=\"rp-donat\">${Sa.donat(s.map(a=>({nilai:a.jumlah,warna:Wa[a.label]||\"var(--abu-2)\"})),c(t.kpi.discrepancy),\"discrepancy\")}\n          <div class=\"rp-daftar rp-daftar--rapat\">${s.map(a=>u`<button type=\"button\" class=\"rp-akar\" data-aksi=\"report-akar\" data-label=\"${a.label}\" ${a.jumlah?\"\":\"disabled\"}><i style=\"background:${Wa[a.label]||\"var(--abu-2)\"}\"></i><span>${a.label}</span><b class=\"angka\">${c(a.jumlah)}</b></button>`)}</div></div>\n      </section>\n      <section class=\"card hm-kartu\"><h2 class=\"card__judul\">Penyumbang terbesar</h2>\n        <div class=\"seg\" style=\"margin:10px 0 10px\">${[[\"userCycle\",\"Petugas\"],[\"validator\",\"Validator\"],[\"operasional\",\"Masalah\"]].map(t=>u`<button type=\"button\" class=\"${a.kontrib===t[0]?\"is-on\":\"\"}\" data-aksi=\"report-kontrib\" data-v=\"${t[0]}\">${t[1]}</button>`)}</div>\n        ${n.length?u`<div class=\"rp-daftar\">${n.map(a=>u`<div class=\"orang\"><div class=\"orang__atas\"><b>${a.nama}</b><span class=\"angka\">${c(a.jumlah)}</span></div><div class=\"bar bar--tipis bar--merah\"><i style=\"width:${a.jumlah/l*100}%\"></i></div><div class=\"orang__ket\">${m(a.persen)}% dari semua discrepancy</div></div>`)}</div>`:u`<p class=\"redup\">Belum ada data pada rentang ini.</p>`}\n      </section>\n      <section class=\"card hm-kartu\"><h2 class=\"card__judul\">Saran tindakan</h2>\n        ${e.length?u`<p class=\"card__ket\">${m(e[0].persen)}% selisih berasal dari ${e[0].label}. Mulai dari sana.</p>`:u`<p class=\"card__ket\">Tidak ada discrepancy pada rentang ini.</p>`}\n        <div class=\"rp-daftar\">${e.slice(0,2).map(a=>Ga[a.label]).filter(Boolean).concat([[\"grafik\",\"Pantau lokasi yang berulang\",\"Perhatikan lokasi dan SKU yang sering muncul di daftar selisih.\"]]).map(a=>u`<div class=\"rp-saran\">${I(a[0])}<div><b>${a[1]}</b><span>${a[2]}</span></div></div>`)}</div>\n      </section>`},async bukaAkar(a){const t=this.an,s=ea.buka({judul:a,penuh:!0,isi:ga()});try{const e=await Y(\"getAnalyticsRootCauseDetail\",a,t.dari,t.sampai,Q.user.username);let n=12;const l=()=>{o(s.isi,e.length?u`<p class=\"redup\" style=\"margin-bottom:10px\">${c(e.length)} item, terbaru di atas.</p><div class=\"rp-daftar\">${e.slice(0,n).map(a=>u`<div class=\"rp-item\"><div class=\"vf-row__atas\">${C(a.lokasi)}<span class=\"row__t\">${a.sku}</span><span class=\"pill pill--merah vf-selisih\">${c(a.qtyError)}</span></div>\n          <div class=\"row__s row__s--lipat\">${a.description}</div>\n          <div class=\"rinci rinci--rapat\"><div class=\"rinci__baris\"><span>Tanggal</span><b>${_(a.tanggal)}</b></div><div class=\"rinci__baris\"><span>Petugas</span><b>${a.userCycle}</b></div><div class=\"rinci__baris\"><span>Validator</span><b>${a.validator}</b></div><div class=\"rinci__baris\"><span>User WMS</span><b>${a.wmsUser} (${a.transaksi})</b></div>${a.kategori&&\"-\"!==a.kategori?u`<div class=\"rinci__baris\"><span>Kategori</span><b>${a.kategori}</b></div>`:\"\"}</div>\n          ${a.hasilInvestigasi?u`<div class=\"vf-catatan\">${a.hasilInvestigasi}</div>`:\"\"}</div>`)}</div>\n          ${e.length>n?u`<button type=\"button\" class=\"btn btn--tenang\" id=\"rpAkarLagi\" style=\"margin-top:12px\">Tampilkan ${c(Math.min(12,e.length-n))} lagi</button>`:\"\"}`:ka(\"daftar\",\"Tidak ada item\",\"Tidak ada selisih dengan penyebab ini pada rentang tanggal terpilih.\"));const a=i(\"rpAkarLagi\");a&&a.addEventListener(\"click\",()=>{n+=12,l()})};l()}catch(a){o(s.isi,u`<p class=\"pesan pesan--galat\">${a.message}</p>`)}}};Ya.reset(),ca.report={masuk(){Ya.gambar(),Ya.buka()}},ia[\"report-segar\"]=()=>Ya.segar(),ia[\"report-tampilan\"]=a=>{Ya.tampilan=a.dataset.v,Ya.gambar(),Ya.buka()},ia[\"report-mode\"]=a=>{Ya.dash.mode=a.dataset.v,Ya.muatDash()},ia[\"report-lb\"]=a=>{Ya.dash.lb=a.dataset.v;const t=window.scrollY;Ya.gambar(),window.scrollTo(0,t)},ia[\"report-backlog\"]=a=>Ya.bukaBacklog(a.dataset.tgl),ia[\"report-user\"]=a=>Ya.bukaUser(a.dataset.nama),ia[\"report-kategori\"]=a=>Ya.bukaKategori(a.dataset.alasan),ia[\"report-masalah\"]=()=>Ya.muatMasalah(!1),ia[\"report-ekspor\"]=()=>\"productivity\"===Ya.tampilan?Ya.eksporProd():Ya.eksporDash(),ia[\"report-kontrib\"]=a=>{Ya.an.kontrib=a.dataset.v;const t=window.scrollY;Ya.gambar(),window.scrollTo(0,t)},ia[\"report-akar\"]=a=>Ya.bukaAkar(a.dataset.label);const Qa=a=>/^(developer|dev|dewa)$/.test(String(a||\"\").trim().toLowerCase()),Xa={tab:\"user\",users:null,roleValid:[],cariUser:\"\",tugas:null,log:null,logSemua:!1,akses:null,facility:null,jalan:{},galat:{},req:0,reset(){Object.assign(this,{tab:\"user\",users:null,cariUser:\"\",tugas:null,log:null,logSemua:!1,akses:null,facility:null,jalan:{},galat:{}})},saya:()=>Q.user.username,gambar(){const a=i(\"scr-config\");o(a,u`${(U.panel?va:ya)(\"Config\",Q.user.facilityName||\"\",u`<button type=\"button\" class=\"iconbtn${this.jalan[this.tab]?\" is-putar\":\"\"}\" data-aksi=\"config-segar\" aria-label=\"Perbarui\">${I(\"segar\")}</button>`)}\n      <div class=\"chips\" style=\"margin-bottom:14px\" role=\"tablist\">${[[\"user\",\"User\"],[\"tugas\",\"Pembagian tugas\"],[\"akses\",\"Akses\"],[\"facility\",\"Facility\"]].map(a=>u`<button type=\"button\" role=\"tab\" class=\"chip${this.tab===a[0]?\" is-on\":\"\"}\" data-aksi=\"config-tab\" data-v=\"${a[0]}\">${a[1]}</button>`)}</div>\n      <div id=\"cfIsi\">${this.galat[this.tab]?fa(this.galat[this.tab],\"config-segar\"):this[this.tab+\"Html\"]()}</div>`);const s=i(\"cfCari\");s&&s.addEventListener(\"input\",E(()=>{this.cariUser=s.value,o(i(\"cfUsers\"),this.daftarUser())},160)),t(\"input[data-m-role]\",a).forEach(a=>a.addEventListener(\"change\",()=>{const t=this.grup(a.dataset.mGrup),s=t.roles.indexOf(a.dataset.mRole);a.checked&&s<0?t.roles.push(a.dataset.mRole):!a.checked&&s>=0&&t.roles.splice(s,1),o(i(\"cfCara\"),this.caraHtml())}));const e=i(\"cfMaks\");e&&e.addEventListener(\"input\",()=>{this.tugas.maksGrupAlat=e.value})},gambarTetap(){const a=window.scrollY;this.gambar(),window.scrollTo(0,a)},ada(a){return!!(\"user\"===a?this.users:\"tugas\"===a?this.tugas:\"akses\"===a?this.akses:this.facility)},async muat(a,t){const i=t||this.tab;if(!Q.user||!a&&(this.ada(i)||this.jalan[i]))return;const s=this.jalan[i]=++this.req,e=this.saya();delete this.galat[i],\"config\"===ma.aktif()&&this.tab===i&&this.gambar();const n=a=>{if(a&&!1===a.success)throw new Error(a.message);return a};let l=null,r=null;try{if(\"user\"===i){const a=n(await Y(\"getDaftarUserMaster\",e));l={users:a.users,roleValid:a.roleValid||[]}}else if(\"tugas\"===i){const a=await Promise.all([Y(\"getLevelAssignmentConfig\",e),Y(\"getLogPerubahanConfig\",e)]);l={tugas:n(a[0]),log:a[1]&&a[1].log||[]}}else l=\"akses\"===i?{akses:n(await Y(\"getDaftarAksesSetting\",e)).daftar}:{facility:n(await Y(\"getDaftarFacility\",e)).facilities}}catch(a){r=a}this.jalan[i]===s&&(delete this.jalan[i],r?this.galat[i]=r:Object.assign(this,l),\"config\"===ma.aktif()&&this.tab===i&&this.gambar())},async ubah(a,t,i){a&&a.classList.add(\"is-sibuk\");try{const a=await t();return a&&!1===a.success?(sa.galat(new Error(a.message)),!1):(a&&a.message&&sa.tampil(a.message,{lama:a.warning?6e3:3200}),i&&i(a),!0)}catch(a){return sa.galat(a),!1}finally{a&&a.classList.remove(\"is-sibuk\")}},peranBoleh(){return this.roleValid.filter(a=>\"developer\"!==a||Z.developer())},userHtml(){return this.users?u`<section class=\"card\"><h2 class=\"card__judul\">Tambah user</h2>\n        <label class=\"lbl\" for=\"cfNik\">NIK</label><input class=\"field\" id=\"cfNik\" type=\"text\" autocapitalize=\"off\" autocorrect=\"off\" spellcheck=\"false\" placeholder=\"contoh: 123456.nama\" autocomplete=\"off\">\n        <label class=\"lbl\" for=\"cfPeran\">Peran</label><select class=\"field\" id=\"cfPeran\"><option value=\"\">Pilih peran</option>${this.peranBoleh().map(a=>u`<option value=\"${a}\">${j(a)}</option>`)}</select>\n        <p class=\"ket\">User baru otomatis masuk ke facility Anda.</p>\n        <button type=\"button\" class=\"btn\" data-aksi=\"config-user-tambah\" style=\"margin-top:14px\">Tambah user</button></section>\n      <h2 class=\"bag\"><span>Daftar user</span><small>${c(this.users.length)} orang</small></h2>\n      ${this.users.length>8?u`<div class=\"cari\" style=\"margin-bottom:10px\">${I(\"cari\",\"ic--kecil\")}<input class=\"field\" type=\"search\" id=\"cfCari\" placeholder=\"Cari NIK atau peran\" value=\"${this.cariUser}\" autocomplete=\"off\"></div>`:\"\"}\n      <div id=\"cfUsers\">${this.daftarUser()}</div>`:ba(4,62)},daftarUser(){const a=this.cariUser.trim().toLowerCase(),t=this.users.filter(t=>!a||B(t.username+\" \"+t.role+\" \"+t.status,a));return t.length?u`<div class=\"list\">${t.map(a=>u`<button type=\"button\" class=\"row\" data-aksi=\"config-user-buka\" data-row=\"${a.rowIndex}\"><span class=\"row__isi\"><span class=\"row__t\">${a.username}</span><span class=\"row__s\">${a.facilityNama||\"Belum punya facility\"}</span></span>\n      <span class=\"row__ekor\"><span class=\"pill${\"Nonaktif\"===a.status?\" pill--merah\":\"\"}\">${\"Nonaktif\"===a.status?\"Nonaktif\":j(Qa(a.role)?\"developer\":a.role)}</span>${I(\"kanan\",\"ic--kecil\")}</span></button>`)}</div>`:u`<p class=\"redup\" style=\"text-align:center;padding:18px 0\">Tidak ada user yang cocok.</p>`},bukaUser(a){const t=this.users.find(t=>String(t.rowIndex)===String(a));if(!t)return;const s=\"Nonaktif\"===t.status,e=Qa(t.role)?\"developer\":String(t.role||\"\").trim().toLowerCase(),n=String(t.username).toLowerCase()===this.saya().toLowerCase();if(\"developer\"===e&&!Z.developer())return void ea.buka({judul:t.username,isi:u`<div class=\"rinci\"><div class=\"rinci__baris\"><span>Peran</span><b>${j(e)}</b></div><div class=\"rinci__baris\"><span>Status</span><b>${s?\"Nonaktif\":\"Aktif\"}</b></div></div>\n        <p class=\"ket\">Akun Developer hanya bisa diubah oleh Developer.</p>`});const l=this.peranBoleh().indexOf(e)>=0?this.peranBoleh():this.peranBoleh().concat([e]),r=ea.buka({judul:t.username,isi:u`<label class=\"lbl\" for=\"cfPeranUbah\">Peran</label><select class=\"field\" id=\"cfPeranUbah\">${l.map(a=>u`<option value=\"${a}\" ${a===e?\"selected\":\"\"}>${j(a)}</option>`)}</select>\n        <button type=\"button\" class=\"btn\" id=\"cfPeranSimpan\" style=\"margin-top:14px\">Simpan peran</button>\n        ${n?u`<p class=\"ket\">Ini akun Anda. Akun sendiri tidak bisa dinonaktifkan.</p>`:u`<button type=\"button\" class=\"btn btn--garis\" id=\"cfStatusUbah\" style=\"margin-top:10px;${s?\"\":\"color:var(--merah-teks)\"}\">${s?\"Aktifkan kembali\":\"Nonaktifkan user\"}</button>\n        <p class=\"ket\">${s?\"User ini sedang tidak bisa masuk.\":\"User nonaktif tidak bisa masuk dan tidak menerima tugas baru. Riwayatnya tetap tersimpan.\"}</p>`}`});i(\"cfPeranSimpan\").addEventListener(\"click\",a=>this.ubah(a.currentTarget,()=>Y(\"updateRoleUserMaster\",this.saya(),t.rowIndex,t.username,i(\"cfPeranUbah\").value),()=>{r.tutup(),this.users=null,this.muat(!0,\"user\")})),n||i(\"cfStatusUbah\").addEventListener(\"click\",async a=>{const i=a.currentTarget;(s||await na({judul:\"Nonaktifkan \"+t.username+\"?\",pesan:\"User ini tidak akan bisa masuk lagi sampai diaktifkan kembali.\",ya:\"Nonaktifkan\",bahaya:!0}))&&this.ubah(i,()=>Y(\"setStatusUserMaster\",this.saya(),t.rowIndex,t.username,s?\"Aktif\":\"Nonaktif\"),()=>{r.tutup(),this.users=null,this.muat(!0,\"user\")})})},PENDEK:{bawah:[\"Bawah\",\"Level 1-2\"],tangga:[\"Tangga\",\"Level 3-4\"],reach_truck:[\"Reach truck\",\"Level 5-6\"]},grup(a){const t=this.tugas;let i=t.matrix.find(t=>t.grupAlat===a);return i||(i={grupAlat:a,roles:[],mode:\"pemerataan\"},t.matrix.push(i)),i},tugasHtml(){const a=this.tugas;if(!a)return ba(3,110);const t=\"matrix\"===a.modeAssignment,i=this.PENDEK,s=this.log||[];return u`<section class=\"card\"><h2 class=\"card__judul\">${t?\"Mode Matrix aktif\":\"Mode lama aktif\"}</h2>\n        <p class=\"card__ket\">${t?\"Tugas harian dibagi mengikuti tabel peran di bawah.\":\"Storing mengerjakan Level 5-6, peran lain mengerjakan Level 1-4. Tabel berikut belum berlaku sampai Mode Matrix diaktifkan.\"}</p>\n        <button type=\"button\" class=\"btn ${t?\"btn--garis\":\"\"}\" data-aksi=\"config-mode\" data-v=\"${t?\"legacy\":\"matrix\"}\" style=\"margin-top:14px\">${t?\"Kembali ke mode lama\":\"Aktifkan Mode Matrix\"}</button></section>\n      <section class=\"card hm-kartu\"><h2 class=\"card__judul\">Peran per grup alat</h2><p class=\"card__ket\">Centang peran yang boleh mengerjakan tiap grup alat.</p>\n        <div class=\"cf-matriks\"><table><thead><tr><th>Peran</th>${a.grupAlatList.map(a=>u`<th>${(i[a.key]||[a.label])[0]}<small>${(i[a.key]||[\"\",\"\"])[1]}</small></th>`)}</tr></thead>\n          <tbody>${a.roleValid.map(t=>u`<tr><th scope=\"row\">${j(t)}</th>${a.grupAlatList.map(a=>u`<td><input type=\"checkbox\" class=\"cek\" data-m-role=\"${t}\" data-m-grup=\"${a.key}\" ${this.grup(a.key).roles.indexOf(t)>=0?\"checked\":\"\"} aria-label=\"${j(t)} di ${a.label}\"></td>`)}</tr>`)}</tbody></table></div>\n        <h3 class=\"lbl\" style=\"margin-top:18px\">Cara bagi</h3>\n        <div class=\"cf-cara\" id=\"cfCara\">${this.caraHtml()}</div>\n        <p class=\"ket\">Merata: semua peran yang dicentang mendapat bagian yang sama. Prioritas: peran dipakai berurutan, peran berikutnya baru ikut bila yang sebelumnya tidak cukup.</p>\n        <label class=\"lbl\" for=\"cfMaks\">Maksimal grup alat per peran</label><input class=\"field\" id=\"cfMaks\" type=\"number\" inputmode=\"numeric\" min=\"1\" max=\"3\" value=\"${a.maksGrupAlat}\" style=\"max-width:120px\">\n        <p class=\"ket\">Hanya peringatan saat dilampaui, tidak menghalangi penyimpanan.</p>\n        <button type=\"button\" class=\"btn\" data-aksi=\"config-matriks-simpan\" style=\"margin-top:14px\">Simpan pembagian tugas</button></section>\n      <h2 class=\"bag\"><span>Riwayat perubahan</span>${s.length?u`<small>${c(s.length)} terakhir</small>`:\"\"}</h2>\n      ${s.length?u`<div class=\"list\">${s.slice(0,this.logSemua?50:5).map(a=>u`<div class=\"row\"><div class=\"row__isi\"><div class=\"row__t row__t--lipat\" style=\"font-weight:500\">${a.detail}</div><div class=\"row__s\">${a.nik}, ${this.waktu(a.waktu)}</div></div></div>`)}</div>\n        ${s.length>5?u`<button type=\"button\" class=\"btn btn--teks\" data-aksi=\"config-log\" style=\"margin:6px auto 0\">${this.logSemua?\"Tampilkan lebih sedikit\":\"Tampilkan semua\"}</button>`:\"\"}`:u`<p class=\"redup\">Belum ada perubahan.</p>`}`},caraHtml(){const a=this.tugas,t=this.PENDEK;return a.grupAlatList.map(a=>{const i=this.grup(a.key),s=\"prioritas\"===i.mode;return u`<div class=\"cf-cara__baris\"><div class=\"cf-cara__atas\"><span>${(t[a.key]||[a.label])[0]} <small>${(t[a.key]||[\"\",\"\"])[1]}</small></span>\n          <div class=\"seg seg--kecil\" role=\"group\" aria-label=\"Cara bagi ${a.label}\"><button type=\"button\" class=\"${s?\"\":\"is-on\"}\" aria-pressed=\"${s?\"false\":\"true\"}\" data-aksi=\"config-cara\" data-grup=\"${a.key}\" data-v=\"pemerataan\">Merata</button><button type=\"button\" class=\"${s?\"is-on\":\"\"}\" aria-pressed=\"${s?\"true\":\"false\"}\" data-aksi=\"config-cara\" data-grup=\"${a.key}\" data-v=\"prioritas\">Prioritas</button></div></div>\n        ${s?u`<p class=\"ket\">${i.roles.length?u`Urutan: <b>${i.roles.map(j).join(\", lalu \")}</b>. Untuk memindahkan satu peran ke urutan terakhir, hapus centangnya lalu centang lagi.`:\"Belum ada peran yang dicentang.\"}</p>`:\"\"}</div>`})},waktu(a){const t=new Date(a);return isNaN(t)?String(a||\"\"):t.getDate()+\" \"+b[t.getMonth()]+\" \"+t.getFullYear()+\", \"+$(t.getTime())},simpanMatriks(a){const t=this.tugas.grupAlatList.map(a=>{const t=this.grup(a.key);return{grupAlat:a.key,roles:t.roles.slice(),mode:\"prioritas\"===t.mode?\"prioritas\":\"pemerataan\"}});this.ubah(a,()=>Y(\"saveLevelAssignmentMatrix\",this.saya(),t,Number(i(\"cfMaks\").value)||2),a=>{a.warning&&sa.tampil(a.warning,{lama:7e3}),this.tugas=null,this.muat(!0,\"tugas\")})},async gantiMode(a,t){await na(\"matrix\"===t?{judul:\"Aktifkan Mode Matrix?\",pesan:\"Mulai upload berikutnya, tugas dibagi mengikuti tabel peran per grup alat.\",ya:\"Aktifkan\"}:{judul:\"Kembali ke mode lama?\",pesan:\"Tugas kembali dibagi dengan aturan Storing dan non-Storing. Isi tabel tidak hilang dan bisa diaktifkan lagi.\",ya:\"Kembali ke mode lama\"})&&this.ubah(a,()=>Y(\"setModeAssignment\",this.saya(),t),()=>{this.tugas=null,Ja.form=null,this.muat(!0,\"tugas\")})},aksesHtml(){return this.akses?u`<section class=\"card\"><h2 class=\"card__judul\">Beri akses Config</h2><p class=\"card__ket\">NIK di daftar ini bisa membuka halaman Config.</p>\n        <label class=\"lbl\" for=\"cfAksesNik\">NIK</label><input class=\"field\" id=\"cfAksesNik\" type=\"text\" autocapitalize=\"off\" autocorrect=\"off\" spellcheck=\"false\" autocomplete=\"off\" placeholder=\"contoh: 123456.nama\">\n        <label class=\"lbl\" for=\"cfAksesNama\">Nama <small>(opsional)</small></label><input class=\"field\" id=\"cfAksesNama\" type=\"text\" autocomplete=\"off\">\n        <button type=\"button\" class=\"btn\" data-aksi=\"config-akses-tambah\" style=\"margin-top:14px\">Beri akses</button></section>\n      <h2 class=\"bag\"><span>Pemegang akses</span><small>${c(this.akses.length)} orang</small></h2>\n      ${this.akses.length?u`<div class=\"list\">${this.akses.map(a=>{const t=String(a.nik).toLowerCase()===this.saya().toLowerCase();return u`<div class=\"row\"><div class=\"row__isi\"><div class=\"row__t\">${a.nik}</div><div class=\"row__s\">${a.nama||\"Tanpa nama\"}${a.tanggal?u`, sejak ${S(a.tanggal)}`:\"\"}</div></div>\n        <div class=\"row__ekor\">${t?u`<span class=\"pill\">Anda</span>`:u`<button type=\"button\" class=\"iconbtn iconbtn--polos\" data-aksi=\"config-akses-cabut\" data-row=\"${a.rowIndex}\" data-nik=\"${a.nik}\" aria-label=\"Cabut akses ${a.nik}\" style=\"color:var(--merah-teks)\">${I(\"sampah\")}</button>`}</div></div>`})}</div>`:u`<p class=\"redup\">Belum ada NIK terdaftar.</p>`}`:ba(3,62)},facilityHtml(){return this.facility?u`${Z.developer()?u`<button type=\"button\" class=\"btn btn--tenang\" data-aksi=\"config-fac-baru\" style=\"margin-bottom:12px\">${I(\"tambah\")}Tambah facility baru</button>`:\"\"}\n      ${this.facility.length?u`<div class=\"tumpuk\">${this.facility.map(a=>{const t=\"Aktif\"===a.status,i=a.spreadsheetId?\"https://docs.google.com/spreadsheets/d/\"+encodeURIComponent(a.spreadsheetId)+\"/edit\":a.url;return u`<section class=\"card\">\n        <div class=\"cf-fac__atas\"><div><h2 class=\"card__judul\">${a.nama}</h2><div class=\"card__ket\">${a.kode}${a.tanggalDibuat?u`, dibuat ${S(a.tanggalDibuat)}`:\"\"}</div></div><span class=\"pill ${t?\"pill--hijau\":\"pill--merah\"}\">${a.status}</span></div>\n        <div class=\"cf-fac__aksi\">\n          <button type=\"button\" class=\"btn btn--kecil btn--garis\" data-aksi=\"config-fac-lokasi\" data-id=\"${a.id}\">${I(\"pin\",\"ic--kecil\")}Lokasi aktif</button>\n          <button type=\"button\" class=\"btn btn--kecil btn--garis\" data-aksi=\"config-fac-user\" data-id=\"${a.id}\">${I(\"orang\",\"ic--kecil\")}User</button>\n          <button type=\"button\" class=\"btn btn--kecil btn--garis\" data-aksi=\"config-fac-ubah\" data-id=\"${a.id}\">${I(\"pensil\",\"ic--kecil\")}Ubah</button>\n          <button type=\"button\" class=\"btn btn--kecil btn--garis\" data-aksi=\"config-fac-status\" data-id=\"${a.id}\">${t?\"Nonaktifkan\":\"Aktifkan\"}</button>\n        </div>\n        ${i?u`<a class=\"tautan cf-fac__db\" href=\"${i}\" target=\"_blank\" rel=\"noopener\">Buka database di Google Sheets</a>`:\"\"}</section>`})}</div>`:u`<section class=\"card\">${ka(\"gedung\",\"Belum ada facility\",\"Facility dibuat oleh Developer.\")}</section>`}`:ba(2,150)},fac(a){return(this.facility||[]).find(t=>t.id===a)},ubahFacility(a){const t=this.fac(a);if(!t)return;const s=ea.buka({judul:\"Ubah facility\",isi:u`<label class=\"lbl\" for=\"cfFacNama\">Nama facility</label><input class=\"field\" id=\"cfFacNama\" type=\"text\" value=\"${t.nama}\">\n        <label class=\"lbl\" for=\"cfFacKode\">Kode</label><input class=\"field\" id=\"cfFacKode\" type=\"text\" value=\"${t.kode}\">\n        <button type=\"button\" class=\"btn\" id=\"cfFacSimpan\" style=\"margin-top:16px\">Simpan</button>`});i(\"cfFacSimpan\").addEventListener(\"click\",t=>{const e=i(\"cfFacNama\").value.trim();e?this.ubah(t.currentTarget,()=>Y(\"updateNamaFacility\",this.saya(),a,e,i(\"cfFacKode\").value.trim()),()=>{s.tutup(),this.facility=null,this.muat(!0,\"facility\")}):sa.galat(new Error(\"Nama facility belum diisi.\"))})},async statusFacility(a,t){const i=this.fac(t);if(!i)return;const s=\"Aktif\"===i.status;s&&!await na({judul:\"Nonaktifkan \"+i.nama+\"?\",pesan:\"User di facility ini tidak bisa mengirim hasil sampai facility diaktifkan kembali.\",ya:\"Nonaktifkan\",bahaya:!0})||this.ubah(a,()=>Y(\"setStatusFacility\",this.saya(),t,s?\"Nonaktif\":\"Aktif\"),()=>{this.facility=null,this.muat(!0,\"facility\")})},async userFacility(a){const t=this.fac(a);if(!t)return;const s=ea.buka({judul:\"User di \"+t.nama,isi:ga()}),e=async()=>{try{const t=await Promise.all([Y(\"getDaftarUserMaster\",this.saya()),Y(\"getDaftarUserFacilityAssignment\",this.saya())]);if(!1===t[0].success)throw new Error(t[0].message);const n=t[0].users.filter(a=>\"Nonaktif\"!==a.status),l=n.filter(t=>t.facilityId===a),r={};(t[1]&&t[1].assignments||[]).forEach(t=>{t.facilityId===a&&(r[t.username]=t.tanggalDiassign)}),o(s.isi,u`<label class=\"lbl\" for=\"cfAssign\">Masukkan user ke facility ini</label>\n          <select class=\"field\" id=\"cfAssign\"><option value=\"\">Pilih user</option>${n.filter(t=>t.facilityId!==a).map(a=>u`<option value=\"${a.username}\">${a.username} (${j(a.role)}${a.facilityNama?\", kini di \"+a.facilityNama:\"\"})</option>`)}</select>\n          <button type=\"button\" class=\"btn\" id=\"cfAssignBtn\" style=\"margin-top:12px\">Masukkan</button>\n          <h3 class=\"bag\"><span>Sudah di facility ini</span><small>${c(l.length)} orang</small></h3>\n          ${l.length?u`<div class=\"list\">${l.map(a=>u`<div class=\"row\"><div class=\"row__isi\"><div class=\"row__t\">${a.username}</div>${r[a.username]?u`<div class=\"row__s\">Sejak ${S(r[a.username])}</div>`:\"\"}</div><span class=\"pill\">${j(a.role)}</span></div>`)}</div>`:u`<p class=\"redup\">Belum ada user.</p>`}`),i(\"cfAssignBtn\").addEventListener(\"click\",t=>{const s=i(\"cfAssign\").value;s?this.ubah(t.currentTarget,()=>Y(\"assignUserKeFacility\",this.saya(),s,a),()=>{this.users=null,e()}):sa.galat(new Error(\"Pilih user dulu.\"))})}catch(a){o(s.isi,u`<p class=\"pesan pesan--galat\">${a.message}</p>`)}};e()},sumberLokasiHtml:(a,t)=>u`<div class=\"seg\" id=\"${a}Sumber\"><button type=\"button\" class=\"is-on\" data-v=\"file\">File</button><button type=\"button\" data-v=\"tempel\">Tempel teks</button>${t.length?u`<button type=\"button\" data-v=\"salin\">Salin</button>`:\"\"}</div>\n      <div id=\"${a}File\" style=\"margin-top:10px\"><label class=\"up-berkas\" id=\"${a}Ubin\"><input type=\"file\" class=\"sr\" id=\"${a}Berkas\" accept=\".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv\">\n        <span class=\"up-berkas__ic\">${I(\"berkas\")}</span><span class=\"up-berkas__isi\"><span class=\"up-berkas__judul\">File daftar lokasi</span><span class=\"up-berkas__ket\">.xlsx atau .csv, kode lokasi di kolom pertama. Baris judul boleh ada.</span></span><span class=\"up-berkas__aksi\">Pilih file</span></label></div>\n      <div id=\"${a}Tempel\" style=\"margin-top:10px\" hidden><textarea class=\"field field--tempel\" id=\"${a}Teks\" rows=\"5\" placeholder=\"A01.001.1&#10;A01.001.2&#10;A01.002.1\" spellcheck=\"false\" autocapitalize=\"characters\"></textarea><p class=\"ket\">Satu lokasi per baris.</p></div>\n      <div id=\"${a}Salin\" style=\"margin-top:10px\" hidden><select class=\"field\" id=\"${a}Dari\"><option value=\"\">Pilih facility sumber</option>${t.map(a=>u`<option value=\"${a.id}\">${a.nama} (${a.kode})</option>`)}</select><p class=\"ket\">Daftar lokasi aktif facility itu disalin ke sini.</p></div>`,pasangSumber(s){const e=t(\"#\"+s+\"Sumber button\");e.forEach(a=>a.addEventListener(\"click\",()=>{e.forEach(t=>t.classList.toggle(\"is-on\",t===a)),[\"File\",\"Tempel\",\"Salin\"].forEach(t=>{const e=i(s+t);e&&(e.hidden=t.toLowerCase()!==a.dataset.v)})}));const n=i(s+\"Berkas\"),l=i(s+\"Ubin\");return this._lokasiBerkas=this._lokasiBerkas||{},delete this._lokasiBerkas[s],n.addEventListener(\"change\",async()=>{const t=n.files&&n.files[0];if(!t)return;const i=a(\".up-berkas__ket\",l),e=a(\".up-berkas__ic\",l),r=(t,s,n)=>{l.className=\"up-berkas\"+(t?\" \"+t:\"\"),i.textContent=s,o(e,n),a(\".up-berkas__aksi\",l).textContent=\"Ganti\"};delete this._lokasiBerkas[s],r(\"\",\"Membaca \"+t.name+\"…\",u`<span class=\"putar\"></span>`);try{const a=await this.bacaBerkasLokasi(t);if(!a.length)throw new Error(\"Tidak ada kode lokasi di kolom pertama file ini.\");this._lokasiBerkas[s]=a,r(\"is-ok\",t.name+\": \"+c(a.length)+\" lokasi terbaca.\",I(\"cek\"))}catch(a){r(\"is-galat\",a.message,I(\"berkas\"))}}),()=>{const a=e.find(a=>a.classList.contains(\"is-on\"));return a?a.dataset.v:\"file\"}},async bacaBerkasLokasi(a){const t=await _a(),i=await new Promise((t,i)=>{const s=new FileReader;s.onload=()=>t(s.result),s.onerror=()=>i(new Error(\"File tidak bisa dibaca.\")),s.readAsArrayBuffer(a)}),s=t.read(new Uint8Array(i),{type:\"array\",raw:!0}),e=[],n={};return t.utils.sheet_to_json(s.Sheets[s.SheetNames[0]],{header:1}).forEach((a,t)=>{const i=Va(a&&a[0]);!i||0===t&&/^(lokasi|location|loc)\\b/i.test(i)||n[i.toUpperCase()]||(n[i.toUpperCase()]=!0,e.push(i))}),e},async bacaLokasi(a,t){if(\"tempel\"===t)return i(a+\"Teks\").value.split(/\\r?\\n/).map(a=>a.trim()).filter(Boolean);const s=this._lokasiBerkas&&this._lokasiBerkas[a];if(s)return s;const e=i(a+\"Berkas\");if(!e.files||!e.files[0])throw new Error(\"Pilih file daftar lokasi dulu.\");return this.bacaBerkasLokasi(e.files[0])},async bukaLokasi(a,s,e){const n=ea.buka({judul:\"Lokasi aktif\",penuh:!0,isi:ga()});let l=null;try{const t=await Y(\"getDaftarLokasiAktif\",this.saya(),a);if(!1===t.success)throw new Error(t.message);l=t.total||0}catch(a){return void o(n.isi,u`<p class=\"pesan pesan--galat\">${a.message}</p>`)}if(!this.facility)try{const a=await Y(\"getDaftarFacility\",this.saya());a.success&&(this.facility=a.facilities)}catch(a){}const r=(this.facility||[]).filter(t=>t.id!==a&&\"Aktif\"===t.status);o(n.isi,u`<div class=\"up-hasil\"><b class=\"angka\">${c(l)}</b><span>lokasi aktif di ${s}</span></div>\n      <p class=\"redup\" style=\"margin-bottom:14px\">Hanya transaksi di lokasi aktif yang dijadikan tugas. Lokasi virtual dan lantai tidak perlu dimasukkan.</p>\n      <span class=\"lbl\">Sumber daftar lokasi</span>${this.sumberLokasiHtml(\"cfLok\",r)}\n      <span class=\"lbl\">Cara memasukkan</span>\n      <div class=\"seg\" id=\"cfLokCara\"><button type=\"button\" class=\"is-on\" data-v=\"tambah\">Tambahkan</button><button type=\"button\" data-v=\"ganti\">Ganti semua</button></div>\n      <p class=\"ket\">Tambahkan: lokasi yang sudah ada dibiarkan. Ganti semua: daftar lama dihapus dulu.</p>\n      <div class=\"pesan pesan--galat\" id=\"cfLokPesan\" role=\"alert\"></div>\n      <button type=\"button\" class=\"btn\" id=\"cfLokSimpan\" style=\"margin-top:14px\">Impor lokasi</button>`);const d=this.pasangSumber(\"cfLok\"),p=t(\"#cfLokCara button\");p.forEach(a=>a.addEventListener(\"click\",()=>p.forEach(t=>t.classList.toggle(\"is-on\",t===a)))),i(\"cfLokSimpan\").addEventListener(\"click\",async t=>{const s=t.currentTarget,r=p.find(a=>a.classList.contains(\"is-on\")),u=!!r&&\"ganti\"===r.dataset.v,o=a=>{i(\"cfLokPesan\").textContent=a};o(\"\");try{let t;if(\"salin\"===d()){const s=i(\"cfLokDari\").value;if(!s)return o(\"Pilih facility sumber.\");t=()=>Y(\"copyLokasiDariFacility\",this.saya(),s,a,u)}else{const i=await this.bacaLokasi(\"cfLok\",d());if(!i.length)return o(\"Daftar lokasi masih kosong.\");if(u&&!await na({judul:\"Ganti semua lokasi aktif?\",pesan:c(l)+\" lokasi lama dihapus dan diganti \"+c(i.length)+\" lokasi baru.\",ya:\"Ganti semua\",bahaya:!0}))return;t=()=>Y(\"importLokasiAktif\",this.saya(),a,i,u)}n.terkunci=!0;await this.ubah(s,t,()=>{n.terkunci=!1,n.tutup(),e&&e()})||(n.terkunci=!1)}catch(a){o(a.message),n.terkunci=!1}})},bukaFacilityBaru(){const a=(this.facility||[]).filter(a=>\"Aktif\"===a.status),t=ea.buka({judul:\"Tambah facility baru\",penuh:!0,isi:u`<p class=\"redup\" style=\"margin-bottom:6px\">Tiap facility punya spreadsheet sendiri. Spreadsheet baru dibuat otomatis di Google Drive akun yang memasang aplikasi ini.</p>\n        <label class=\"lbl\" for=\"cfBaruNama\">Nama facility</label><input class=\"field\" id=\"cfBaruNama\" type=\"text\" placeholder=\"contoh: RDC Manado\">\n        <label class=\"lbl\" for=\"cfBaruKode\">Kode</label><input class=\"field\" id=\"cfBaruKode\" type=\"text\" placeholder=\"contoh: RDC-MDC\">\n        <label class=\"lbl\" for=\"cfBaruSs\">Nama spreadsheet</label><input class=\"field\" id=\"cfBaruSs\" type=\"text\" placeholder=\"contoh: Cycle Count RDC Manado\">\n        <span class=\"lbl\">Daftar lokasi aktif</span>${this.sumberLokasiHtml(\"cfBaru\",a)}\n        <div class=\"pesan pesan--galat\" id=\"cfBaruPesan\" role=\"alert\"></div>\n        <button type=\"button\" class=\"btn\" id=\"cfBaruSimpan\" style=\"margin-top:16px\">Buat facility</button>`}),s=this.pasangSumber(\"cfBaru\");let e=!1;i(\"cfBaruSimpan\").addEventListener(\"click\",async a=>{const n=a.currentTarget,l=a=>{i(\"cfBaruPesan\").textContent=a},r=i(\"cfBaruNama\").value.trim(),u=i(\"cfBaruKode\").value.trim(),o=i(\"cfBaruSs\").value.trim();if(!r||!u||!o)return l(\"Nama facility, kode, dan nama spreadsheet wajib diisi.\");if(!e){e=!0,t.terkunci=!0,n.classList.add(\"is-sibuk\"),l(\"\");try{let a;if(\"salin\"===s()){const t=i(\"cfBaruDari\").value;if(!t)return l(\"Pilih facility sumber.\");const s=await Y(\"getDaftarLokasiAktif\",this.saya(),t);if(!1===s.success)throw new Error(s.message);a=s.lokasi||[]}else a=await this.bacaLokasi(\"cfBaru\",s());if(!a.length)return l(\"Daftar lokasi aktif masih kosong.\");const e=\"job_\"+Date.now()+\"_\"+Math.floor(1e5*Math.random());la.tampil(\"Membuat facility baru\",c(a.length)+\" lokasi\"),la.pantau(e);const n=await Y(\"tambahFacility\",this.saya(),r,u,o,a,e);if(la.sembunyi(),!1===n.success)return l(n.message);t.terkunci=!1,t.tutup(),sa.tampil(n.message||\"Facility dibuat.\"),this.facility=null,this.muat(!0,\"facility\")}catch(a){la.sembunyi(),l(a.message)}finally{e=!1,t.terkunci=!1,n.classList.remove(\"is-sibuk\")}}})}};ca.config={masuk(){Z.config()?(Xa.gambar(),Xa.muat(!1)):ma.ke(ma.awal())}},ia[\"config-segar\"]=()=>Xa.muat(!0),ia[\"config-tab\"]=a=>{Xa.tab=a.dataset.v,delete Xa.galat[Xa.tab],Xa.gambar(),Xa.muat(!1)},ia[\"config-user-tambah\"]=a=>{const t=i(\"cfNik\").value.trim();return t?i(\"cfPeran\").value?void Xa.ubah(a,()=>Y(\"tambahUserMaster\",Xa.saya(),t,i(\"cfPeran\").value),()=>{Xa.users=null,Ja.form=null,Xa.muat(!0,\"user\")}):(sa.galat(new Error(\"Pilih peran untuk user ini.\")),void i(\"cfPeran\").focus()):(sa.galat(new Error(\"NIK belum diisi.\")),void i(\"cfNik\").focus())},ia[\"config-user-buka\"]=a=>Xa.bukaUser(a.dataset.row),ia[\"config-mode\"]=a=>Xa.gantiMode(a,a.dataset.v),ia[\"config-matriks-simpan\"]=a=>Xa.simpanMatriks(a),ia[\"config-cara\"]=a=>{Xa.grup(a.dataset.grup).mode=a.dataset.v,o(i(\"cfCara\"),Xa.caraHtml())},ia[\"config-log\"]=()=>{Xa.logSemua=!Xa.logSemua,Xa.gambarTetap()},ia[\"config-akses-tambah\"]=a=>{const t=i(\"cfAksesNik\").value.trim();if(!t)return sa.galat(new Error(\"NIK belum diisi.\")),void i(\"cfAksesNik\").focus();Xa.ubah(a,()=>Y(\"tambahAksesSetting\",Xa.saya(),t,i(\"cfAksesNama\").value.trim()),()=>{Xa.akses=null,Xa.muat(!0,\"akses\")})},ia[\"config-akses-cabut\"]=async a=>{await na({judul:\"Cabut akses \"+a.dataset.nik+\"?\",pesan:\"NIK ini tidak bisa lagi membuka Config.\",ya:\"Cabut akses\",bahaya:!0})&&Xa.ubah(null,()=>Y(\"hapusAksesSetting\",Xa.saya(),Number(a.dataset.row),a.dataset.nik),()=>{Xa.akses=null,Xa.muat(!0,\"akses\")})},ia[\"config-fac-baru\"]=()=>Xa.bukaFacilityBaru(),ia[\"config-fac-ubah\"]=a=>Xa.ubahFacility(a.dataset.id),ia[\"config-fac-status\"]=a=>Xa.statusFacility(a,a.dataset.id),ia[\"config-fac-user\"]=a=>Xa.userFacility(a.dataset.id),ia[\"config-fac-lokasi\"]=a=>{const t=Xa.fac(a.dataset.id);t&&Xa.bukaLokasi(t.id,t.nama,null)};const Za={gelap:()=>\"gelap\"===document.documentElement.getAttribute(\"data-tema\"),pasang(){this._bilah()},atur(a){a?document.documentElement.setAttribute(\"data-tema\",\"gelap\"):document.documentElement.removeAttribute(\"data-tema\"),d.set(\"ct.tema\",a?\"gelap\":\"terang\"),this._bilah()},_bilah(){try{O&&K.SystemBars&&K.SystemBars.setStyle&&K.SystemBars.setStyle({style:this.gelap()?\"DARK\":\"LIGHT\"}).catch(()=>{})}catch(a){}}},at={rilis:null,sibuk:!1,berkas:\"\",dicekAt:0,lembar:null,mengecek:!1,_manual:!1,_bagian:a=>String(a||\"\").replace(/^v/i,\"\").split(\".\").map(a=>parseInt(a,10)||0),lebihBaru(a,t){const i=this._bagian(a),s=this._bagian(t);for(let a=0;a<Math.max(i.length,s.length);a++)if((i[a]||0)!==(s[a]||0))return(i[a]||0)>(s[a]||0);return!1},plugin(){try{return O&&H.isPluginAvailable&&H.isPluginAvailable(\"CapacitorUpdater\")?K.CapacitorUpdater:null}catch(a){return null}},bisaKilat(a){const t=a&&a.live;return!!(t&&this.plugin()&&U.nativeBase&&t.native===U.nativeBase)&&(d.get(\"ct.live.failed\")||[]).indexOf(a.version)<0},async ambil(){const a=\"https://github.com/\"+U.updateRepo+\"/releases/latest/download/\",t=await V(a+\"latest.json\"),i={version:String(t.version||\"\"),notes:String(t.notes||\"\"),url:a+t.apk,size:Number(t.apkSize)||0,live:null};return t.live&&/^[0-9a-f]{64}$/.test(t.live.sha256||\"\")&&(i.live={native:String(t.live.native),sha256:t.live.sha256,size:Number(t.live.size)||0,url:a+t.live.asset}),i},async cek(a){if(U.panel)return;if(!U.updateRepo||\"dev\"===U.version||U.demoSaja)return void(a&&sa.tampil(\"Pembaruan hanya tersedia di aplikasi Android.\"));const t=()=>i(\"akVersi\");if(a&&(this._manual=!0),this.mengecek)return;if(!a&&Date.now()-this.dicekAt<9e5)return;this.mengecek=!0,this.dicekAt=Date.now(),t()&&(t().textContent=tt.versiTeks());let s=null;try{const a=await this.ambil();d.set(\"ct.upd.latest\",a),this.rilis=a.version&&this.lebihBaru(a.version,U.version)?a:null}catch(a){s=a;const t=d.get(\"ct.upd.latest\");t&&this.lebihBaru(t.version,U.version)&&(this.rilis=t)}this.mengecek=!1;const e=this._manual;if(this._manual=!1,this.gambarKartu(),t()&&(t().textContent=tt.versiTeks()),e)if(this.rilis){const a=ea.tumpukan[ea.tumpukan.length-1];a?tt.lembar&&a===tt.lembar?(a.tutup(),this.buka()):sa.tampil(\"Versi \"+this.rilis.version+\" tersedia.\",{aksi:\"Lihat\",onAksi:()=>this.buka(),lama:8e3}):this.buka()}else s?sa.galat(new Error(\"Tidak bisa memeriksa pembaruan: \"+F(s).message)):sa.tampil(\"Sudah versi terbaru (\"+U.version+\").\")},gambarKartu(){const a=i(\"hmPembaruan\");a&&o(a,this.rilis?u`<button type=\"button\" class=\"lanjut lanjut--biru\" data-aksi=\"pembaruan-buka\" style=\"margin-bottom:12px\"><span class=\"lanjut__ic\">${I(\"unduh\")}</span><span class=\"lanjut__teks\"><b>Versi ${this.rilis.version} tersedia</b><span>Ketuk untuk memperbarui aplikasi.</span></span>${I(\"kanan\")}</button>`:\"\")},catatan(a){const t=String(a||\"\").replace(/\\r/g,\"\").replace(/<\\!--[\\s\\S]*?-->/g,\"\"),i=t.search(/^## Catatan developer/m),s=(i>=0?t.slice(0,i):t).split(\"\\n\").map(a=>a.trim()).filter(a=>a&&!/^## /.test(a)&&!/Unduh file \\.apk/i.test(a)).slice(0,12);return s.length?s.map(a=>u`<li>${a.replace(/^[-*]\\s*/,\"\")}</li>`):u`<li>Perbaikan dan peningkatan.</li>`},buka(){const a=this.rilis;if(!a||this.lembar)return;const t=this.bisaKilat(a);if(this.lembar=ea.buka({judul:\"Versi \"+a.version+\" tersedia\",isi:u`<ul class=\"upd__catatan\">${this.catatan(a.notes)}</ul>\n        <div class=\"bar bar--tipis bar--biru\" id=\"updBar\" hidden><i></i></div><div class=\"pesan\" id=\"updStatus\" aria-live=\"polite\"></div>\n        <button type=\"button\" class=\"btn\" id=\"updTombol\" data-aksi=\"pembaruan-jalan\" style=\"margin-top:14px\">${t?\"Perbarui sekarang\":\"Unduh dan pasang\"}</button>\n        <p class=\"ket\" style=\"text-align:center\">${t?\"Langsung terpasang, tanpa mengunduh APK.\":\"Data tetap aman.\"}</p>`,onTutup:()=>{this.lembar=null}}),this.sibuk){const a=i(\"updTombol\");a.disabled=!0,a.textContent=\"Mengunduh…\"}},_status(t,s){const e=i(\"updBar\"),n=i(\"updStatus\");e&&(e.hidden=null==t,null!=t&&(a(\"i\",e).style.width=Math.max(3,Math.min(100,t))+\"%\")),n&&o(n,s||\"\")},async jalan(){const a=this.rilis;if(!a||this.sibuk)return;const t=i(\"updTombol\");let s=null;if(a.live&&this.bisaKilat(a)){this.sibuk=!0,t&&(t.disabled=!0);const i=this.plugin();try{s=await i.addListener(\"download\",a=>{const t=Number(a&&a.percent)||0;this._status(t,t<70?\"Mengunduh… \"+t+\"%\":\"Memasang…\")}),this._status(2,\"Mengunduh…\");const t=await i.download({url:a.live.url,version:a.version,checksum:a.live.sha256});this._status(100,\"Memuat ulang aplikasi…\"),d.set(\"ct.live.pending\",{v:a.version,from:U.version}),await i.set({id:t.id})}catch(i){d.del(\"ct.live.pending\"),this.sibuk=!1,a.live=null,t&&(t.disabled=!1,t.textContent=\"Unduh dan pasang\"),this._status(null,u`Pembaruan langsung gagal (${i&&i.message||i}). Ketuk tombol untuk mengunduh APK.`)}finally{try{s&&s.remove()}catch(a){}}return}const e=K.FileTransfer,n=K.Filesystem,l=K.FileOpener;if(O&&e&&n&&l){if(this.berkas)return this.pasang();this.sibuk=!0,t&&(t.disabled=!0,t.textContent=\"Mengunduh…\");try{const i=\"CycleTransaksi-\"+a.version+\".apk\";try{await n.deleteFile({path:i,directory:\"CACHE\"})}catch(a){}const l=(await n.getUri({path:i,directory:\"CACHE\"})).uri,r=String(l).replace(/^file:\\/\\//,\"\"),o=a.size||0;s=await e.addListener(\"progress\",a=>{const t=a.lengthComputable&&a.contentLength?a.contentLength:o;t&&this._status(a.bytes/t*100,\"Mengunduh \"+(a.bytes/1048576).toFixed(1)+\" dari \"+(t/1048576).toFixed(1)+\" MB\")}),this._status(1,\"Mengunduh…\"),await e.downloadFile({url:a.url,path:r,progress:!0,connectTimeout:2e4,readTimeout:6e4}),this.berkas=l,this.sibuk=!1,t&&(t.disabled=!1,t.textContent=\"Pasang versi \"+a.version),this._status(100,u`Unduhan selesai. Pilih <b>Update</b> di layar berikutnya.`),await this.pasang()}catch(a){this.sibuk=!1,this.berkas=\"\",t&&(t.disabled=!1,t.textContent=\"Coba lagi\"),this._status(null,\"Unduhan gagal: \"+F(a).message)}finally{try{s&&s.remove()}catch(a){}}}else window.open(a.url,\"_blank\")},async pasang(){try{await K.FileOpener.open({filePath:this.berkas,contentType:\"application/vnd.android.package-archive\",openWithDefault:!0}),this._status(100,u`Penginstal Android terbuka. Kalau diminta, izinkan <b>Instal aplikasi tidak dikenal</b> untuk Cycle Transaksi, lalu kembali dan ketuk tombol lagi.`)}catch(a){this._status(100,\"Tidak bisa membuka penginstal (\"+(a&&a.message||a)+\").\")}},periksaTertunda(){const a=d.get(\"ct.live.pending\");if(!a||!a.v)return;if(d.del(\"ct.live.pending\"),a.v===U.version)return;const t=d.get(\"ct.live.failed\")||[];t.indexOf(a.v)<0&&t.push(a.v),d.set(\"ct.live.failed\",t.slice(-10))}};ia[\"cek-versi\"]=()=>at.cek(!0),ia[\"pembaruan-buka\"]=()=>{ea.tutupSemua(),at.buka()},ia[\"pembaruan-jalan\"]=()=>at.jalan();const tt={lembar:null,versiTeks:()=>at.mengecek?\"Memeriksa…\":\"Versi \"+U.version+(ja.data&&ja.data.version?\", sistem \"+ja.data.version:\"\"),buka(){const a=Q.user;if(!a)return;const t=ta.jumlah();this.lembar=ea.buka({judul:\"Akun\",isi:u`<div class=\"ak__siapa\"><div class=\"hm-avatar\">${T(a.username)}</div><div><div class=\"ak__nama\">${a.username}</div><div class=\"redup\">${j(a.role)}${a.facilityName?u`, ${a.facilityName}`:\"\"}</div></div></div>\n        ${t?u`<button type=\"button\" class=\"info info--jingga\" data-aksi=\"akun-kirim\" style=\"margin-bottom:12px\">${I(\"awan_kirim\")}<span><b>${c(t)} hasil belum terkirim.</b> Ketuk untuk mengirim sekarang.</span></button>`:\"\"}\n        <div class=\"list\">\n          <div class=\"row\"><span class=\"row__isi\"><span class=\"row__t\">Tema gelap</span><span class=\"row__s\">Untuk area gudang yang redup</span></span><button type=\"button\" class=\"saklar\" role=\"switch\" aria-checked=\"${Za.gelap()}\" aria-label=\"Tema gelap\" data-aksi=\"akun-tema\"></button></div>\n          ${Z.developer()?u`<div class=\"row\"><span class=\"row__isi\"><span class=\"row__t\">Facility aktif</span><select class=\"field field--kecil\" id=\"akFacility\" style=\"margin-top:8px\" aria-label=\"Facility aktif\"><option value=\"\">Memuat daftar facility…</option></select></span></div>`:\"\"}\n          ${Z.config()&&!U.panel?u`<button type=\"button\" class=\"row\" data-aksi=\"akun-config\">${I(\"atur\")}<span class=\"row__isi\"><span class=\"row__t\">Config</span><span class=\"row__s\">User, pembagian tugas, akses, facility</span></span><span class=\"row__ekor\">${I(\"kanan\",\"ic--kecil\")}</span></button>`:\"\"}\n          ${U.panel?\"\":u`<button type=\"button\" class=\"row\" data-aksi=\"cek-versi\">${I(\"unduh\")}<span class=\"row__isi\"><span class=\"row__t\">Periksa pembaruan</span><span class=\"row__s\" id=\"akVersi\">${this.versiTeks()}</span></span>${at.rilis?u`<span class=\"pill pill--biru\">Versi baru</span>`:\"\"}</button>`}\n          <button type=\"button\" class=\"row\" data-aksi=\"akun-keluar\" style=\"color:var(--merah-teks)\">${I(\"keluar\")}<span class=\"row__isi\"><span class=\"row__t\">${J.aktif&&!U.demoSaja?\"Keluar dari mode demo\":\"Keluar\"}</span></span></button>\n        </div>`,onTutup:()=>{this.lembar=null}}),Z.developer()&&this.muatFacility()},async muatFacility(){const a=i(\"akFacility\");if(a)try{const t=await Y(\"getDaftarFacility\",Q.user.username);if(!t.success)throw new Error(t.message);o(a,u`<option value=\"\">Assignment normal saya</option>${t.facilities.filter(a=>\"Aktif\"===a.status).map(a=>u`<option value=\"${a.id}\" ${Q.user.facilityId===a.id?\"selected\":\"\"}>${a.nama} (${a.kode})</option>`)}`),a.addEventListener(\"change\",()=>this.gantiFacility(a.value))}catch(t){o(a,u`<option value=\"\">Daftar facility tidak bisa dimuat</option>`)}},async gantiFacility(a){const t=i(\"akFacility\");t&&(t.disabled=!0);try{const t=await Y(\"setDeveloperActiveFacility\",Q.user.username,a);if(!t.success)throw new Error(t.message);const i=await Y(\"getUserRole\",Q.user.username);i&&(Q.simpan(Q.dari(i)),aa.bersih(),ja.mulai(),Aa.reset(),Oa.reset(),Ja.reset(),Ya.reset(),Xa.reset()),ea.tutupSemua(),sa.tampil(t.message),ma.ke(ma.awal()),U.panel||ja.muat(!0)}catch(a){sa.galat(a),t&&(t.disabled=!1)}},async keluar(){const a=J.aktif?0:ta.jumlah();a&&!await na({judul:\"Keluar sekarang?\",pesan:c(a)+\" hasil belum terkirim. Hasil itu tetap dikirim atas nama \"+Q.user.username+\" begitu jaringan tersedia.\",ya:\"Keluar\"})||et(\"\")}};function it(){ea.tutupSemua();const a=ma.aktif();a&&ca[a]&&ca[a].keluar&&ca[a].keluar(),ma.tab=\"login\",ma.halaman=[],i(\"nav\").hidden=!0,document.body.classList.add(\"tanpa-nav\"),ma._tampil(\"login\"),ca.login.masuk()}function st(a){Q.simpan(a),ja.mulai(),Aa.reset(),Oa.reset(),Ja.reset(),Ya.reset(),Xa.reset(),ma.lencana={},i(\"nav\").hidden=!1,ma.bangun(),J.aktif?ra.pasang(\"demo\",u`Mode demo. Data contoh, perubahan tidak tersimpan.`):ra.lepas(\"demo\"),!U.panel||ma.awal()?ma.ke(ma.awal()):et(\"Halaman ini khusus admin dan pemegang akses Config. Petugas memakai aplikasi di HP.\")}function et(a){const t=J.aktif;Q.hapus(),aa.bersih(),ra.lepas(\"demo\"),t&&(ta.kosongkan(),U.demoSaja||(J.berhenti(),ta.muat(),ta.jalan())),it(),a&&Ma.pesan(a)}ia.akun=()=>tt.buka(),ia[\"akun-tema\"]=a=>{const t=!Za.gelap();Za.atur(t),a.setAttribute(\"aria-checked\",String(t))},ia[\"akun-config\"]=()=>{ea.tutupSemua(),ma.dorong(\"config\")},ia[\"akun-keluar\"]=()=>tt.keluar(),ia[\"akun-kirim\"]=()=>{ta.cobaLagi(),ea.tutupSemua(),sa.tampil(\"Mengirim hasil yang tertunda…\")};let nt=0;async function lt(){const a=Q.user;if(a&&!J.aktif){nt=Date.now();try{const t=await Y(\"getUserRole\",a.username);if(Q.user!==a)return;if(!t)return void et(\"Akun ini tidak lagi terdaftar. Hubungi admin.\");const i=Q.dari(t);if(!(i.role!==a.role||i.akses!==a.akses||i.facilityId!==a.facilityId||i.facilityName!==a.facilityName))return void Q.simpan(i);ea.tutupSemua(),aa.bersih();const s=i.facilityId!==a.facilityId?ta.buangFacilityLain(i.username,i.facilityId||\"-\"):0;st(i),sa.tampil(s?\"Akun Anda dipindahkan admin ke facility lain. \"+c(s)+\" hasil yang belum terkirim tidak berlaku lagi.\":\"Akun Anda diperbarui admin. Data dimuat ulang.\",{lama:s?8e3:4500})}catch(a){}}}function rt(){if(da(),z.sinkron(!1),at.cek(!1),!Q.user)return ta.cobaLagi(),void(\"login\"===ma.tab&&\"siap\"!==Ma.status&&\"demo\"!==Ma.status&&Ma.cekServer());Date.now()-nt>3e5&&ta.tahanSampai(lt(),4e3),ta.cobaLagi();const a=ma.aktif();\"home\"===a?ja.muat(!1):\"cycle\"!==a&&\"validasi\"!==a||Aa.muat(a,!1)}function ut(){Za.pasang(),oa(),setTimeout(oa,400),setTimeout(oa,1500),at.periksaTertunda(),ta.muat(),Ma.gambar(),da();const a=at.plugin();a&&a.notifyAppReady().catch(()=>{}),O&&K.App?(K.App.addListener(\"backButton\",()=>{!ma.kembali()&&K.App.minimizeApp&&K.App.minimizeApp()}),K.App.addListener(\"appStateChange\",a=>{a&&a.isActive&&rt()})):document.addEventListener(\"visibilitychange\",()=>{document.hidden||rt()});const t=U.demoSaja?null:Q.muat();if(t)st(t),ta.tahanSampai(lt(),4e3);else{it();const a=U.demoSaja?null:d.get(\"ct.user\");a&&\"string\"==typeof a&&(d.del(\"ct.user\"),i(\"lgNik\").value=a,z.url()&&Ma.kirim(a))}z.sinkron(!1),ta.jalan(),setTimeout(()=>at.cek(!1),1500)}window.CT={panggil:Y,Sesi:Q,Antrean:ta,Nav:ma,Home:ja,Tugas:Aa,Fokus:La,Demo:J,Server:z,Pembaruan:at,versi:U.version},\"loading\"!==document.readyState?ut():document.addEventListener(\"DOMContentLoaded\",ut)}();</script></body></html>";
