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
  return HtmlService.createHtmlOutput('<div style="font-family:sans-serif;padding:24px;line-height:1.5">' + '<h3>Cycle Transaksi — server aktif (' + APP_VERSION + ')</h3>' + '<p>Halaman ini adalah server aplikasi Android Cycle Transaksi. Alamat halaman ini (yang berakhiran <b>/exec</b>) ' + 'ditanam di aplikasi oleh pengembang; pengguna cukup masuk dengan NIK.</p></div>').setTitle('Cycle Transaksi').addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
  const fromDate = dateFrom ? new Date(dateFrom + 'T00:00:00') : null;
  const toDate = dateTo ? new Date(dateTo + 'T23:59:59') : null;
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
