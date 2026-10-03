/*
 * Report (admin): Dashboard (akurasi & ranking per hari/bulan), Productivity (hitungan per orang & per jam),
 * Analytics (akar masalah selisih pada rentang tanggal).
 */
const WARNA_AKAR = { 'Salah Picking / Move': 'var(--merah)', 'Salah Putaway': 'var(--jingga)', 'Human Error Cycle': 'var(--kuning-tepi)', 'Master Data / Location': 'var(--biru)', 'Tidak diketahui': 'var(--abu-2)' };
const SARAN_AKAR = {
  'Salah Picking / Move': ['orang', 'Coaching petugas picking dan move', 'Dampingi petugas dengan kesalahan picking atau move terbanyak.'],
  'Salah Putaway': ['kotak', 'Tinjau proses putaway', 'Periksa penempatan barang di lokasi yang sering bermasalah.'],
  'Human Error Cycle': ['daftar', 'Tinjau cara menghitung', 'Ulangi pelatihan ketelitian hitung untuk petugas cycle.'],
  'Master Data / Location': ['atur', 'Perbaiki master data', 'Periksa data lokasi dan stok yang tidak sesuai.'],
  'Tidak diketahui': ['cari', 'Selesaikan verifikasi', 'Tindak lanjuti selisih yang belum punya kategori penyebab.']
};

const Report = {
  tampilan: 'dashboard',
  dash: null, prod: null, an: null,

  // Objek keadaan dibuat baru, sehingga permintaan milik pengguna sebelumnya yang baru selesai tidak punya tempat mendarat.
  reset() {
    this.tampilan = 'dashboard';
    this.dash = { mode: 'daily', tanggal: '', bulan: '', data: null, untuk: '', sibuk: false, galat: null, req: 0, lb: 'petugas', backlog: null, backlogBasi: false, detail: null, detailSibuk: false, cari: '' };
    this.prod = { tanggal: '', data: null, untuk: '', sibuk: false, galat: null, req: 0 };
    this.an = { dari: '', sampai: '', data: null, untuk: '', sibuk: false, galat: null, req: 0, kontrib: 'userCycle' };
  },
  /*
   * Satu pola untuk ketiga laporan. Pilihan (tanggal, bulan, rentang) boleh diganti selagi data masih dimuat:
   * permintaan terbaru yang berlaku, hasil permintaan lama dibuang. Data periode lain tidak pernah ditampilkan
   * di bawah pilihan yang baru; memperbarui periode yang sama tetap menampilkan data lama sampai yang baru tiba.
   */
  async ambil(nama, kunci, minta, terima) {
    const s = this[nama], no = ++s.req;
    if (s.untuk !== kunci) { s.data = null; s.untuk = kunci; }
    s.sibuk = true; s.galat = null;
    if (Nav.aktif() === 'report') this.gambar();
    let hasil = null, salah = null;
    try { hasil = await minta(); } catch (e) { salah = e; }
    if (this[nama] !== s || s.req !== no) return;
    s.sibuk = false;
    if (salah) s.galat = salah; else terima(hasil);
    if (Nav.aktif() === 'report' && this.keadaan() === s) this.gambar();
  },
  keadaan() { const v = this.tampilan; return v === 'dashboard' ? this.dash : v === 'productivity' ? this.prod : this.an; },
  belumSegar(s) { return s.galat && s.data ? h`<div class="info info--jingga" style="margin-bottom:12px">${ikon('awan_mati')}<div>Belum bisa diperbarui: ${s.galat.message}</div></div>` : ''; },
  awal() {
    const t = hariIni();
    if (!this.dash.tanggal) { this.dash.tanggal = t; this.dash.bulan = t.slice(0, 7); }
    if (!this.prod.tanggal) this.prod.tanggal = t;
    if (!this.an.sampai) { this.an.sampai = t; this.an.dari = geserHari(t, -6); }
  },

  gambar() {
    this.awal();
    const node = el('scr-report'), v = this.tampilan, s = this.keadaan();
    const kanan = h`<button type="button" class="iconbtn${s.sibuk ? ' is-putar' : ''}" data-aksi="report-segar" aria-label="Perbarui">${ikon('segar')}</button>`;
    isi(node, h`${kepala('Report', '', kanan)}
      <div class="seg rp-seg" role="tablist">${[['dashboard', 'Dashboard'], ['productivity', 'Productivity'], ['analytics', 'Analytics']].map((x) => h`<button type="button" role="tab" class="${v === x[0] ? 'is-on' : ''}" data-aksi="report-tampilan" data-v="${x[0]}">${x[1]}</button>`)}</div>
      <div id="rpIsi">${v === 'dashboard' ? this.dashHtml() : v === 'productivity' ? this.prodHtml() : this.anHtml()}</div>`);
    this.pasang();
  },
  pasang() {
    const v = this.tampilan;
    const ubah = (id, fn) => { const n = el(id); if (n) n.addEventListener('change', () => { if (n.value) fn(n.value); }); };
    if (v === 'dashboard') {
      ubah('rpTanggal', (x) => { this.dash.tanggal = x; this.muatDash(); });
      ubah('rpBulan', (x) => { this.dash.bulan = x; this.muatDash(); });
      const d = this.dash.data;
      if (d && el('rpTren')) Grafik.tren(el('rpTren'), d.trend.map((t) => Object.assign({}, t, { label: this.dash.mode === 'monthly' ? bulanPendek(t.period) : tanggalPendek(t.period), judul: this.dash.mode === 'monthly' ? bulanPendek(t.period) : tanggalPanjang(t.period) })), el('rpTrenInfo'));
      const c = el('rpCari');
      if (c) c.addEventListener('input', tunda(() => { this.dash.cari = c.value; isi(el('rpMasalah'), this.masalahHtml()); }, 160));
    } else if (v === 'productivity') {
      ubah('rpProdTanggal', (x) => { this.prod.tanggal = x; this.muatProd(); });
      const d = this.prod.data;
      if (d && d.perJam.length && el('rpJam')) Grafik.batang(el('rpJam'), d.perJam.map((j) => ({ label: dua(j.jam), nilai: j.total })));
    } else {
      ubah('rpDari', (x) => { this.an.dari = x; this.muatAn(); });
      ubah('rpSampai', (x) => { this.an.sampai = x; this.muatAn(); });
      const d = this.an.data;
      if (d && d.trend.length > 1 && el('rpAnTren')) Grafik.garis(el('rpAnTren'), d.trend.map((t) => ({ label: tanggalPendek(t.tanggal), akurasi: t.akurasi })));
    }
  },
  muatIni() { const v = this.tampilan; return v === 'dashboard' ? this.muatDash() : v === 'productivity' ? this.muatProd() : this.muatAn(); },
  segar() { const s = this.keadaan(); if (s.sibuk) return; if (s === this.dash) s.backlogBasi = true; this.muatIni(); },
  buka() { const s = this.keadaan(); if (!s.data && !s.sibuk) this.muatIni(); },

  /* ================= Dashboard ================= */
  muatDash() {
    const s = this.dash, nilai = s.mode === 'daily' ? s.tanggal : s.bulan, siapa = Sesi.user.username, kunci = s.mode + ':' + nilai;
    // Ganti periode: daftar item bermasalah milik periode lama dibuang. Memperbarui periode yang sama: kartu
    // Outstanding Cycle dan daftar itu tetap tampil sampai data baru tiba (dan tetap ada bila pembaruan gagal).
    const ganti = s.untuk !== kunci, adaDaftar = !ganti && !!s.detail;
    if (ganti) { s.detail = null; s.cari = ''; }
    s.detailSibuk = false;
    return this.ambil('dash', kunci,
      () => Promise.all([panggil('getDashboardData', s.mode, nilai, s.mode === 'monthly' ? 6 : 7, siapa), s.backlog && !s.backlogBasi ? s.backlog : panggil('getPendingBacklog', siapa)]),
      (r) => { s.data = r[0]; s.backlog = r[1] || []; s.backlogBasi = false; if (adaDaftar) this.muatMasalah(true); });
  },
  dashHtml() {
    const s = this.dash, d = s.data;
    const atas = h`<div class="rp-periode"><div class="seg"><button type="button" class="${s.mode === 'daily' ? 'is-on' : ''}" data-aksi="report-mode" data-v="daily">Harian</button><button type="button" class="${s.mode === 'monthly' ? 'is-on' : ''}" data-aksi="report-mode" data-v="monthly">Bulanan</button></div>
      ${s.mode === 'daily' ? h`<input type="date" class="field field--kecil" id="rpTanggal" value="${s.tanggal}" max="${hariIni()}" aria-label="Tanggal">` : h`<input type="month" class="field field--kecil" id="rpBulan" value="${s.bulan}" aria-label="Bulan">`}</div>`;
    if (!d) return h`${atas}${s.galat ? gagalMuat(s.galat, 'report-segar') : kerangka(3, 130)}`;
    const k = d.summary.kpi, tunggu0 = Math.max(0, k.total - k.hit - k.discrepancy);
    const akurasi = k.total ? Math.round((k.hit / k.total) * 1000) / 10 : 0;
    const valid = d.trend.filter((t) => t.total > 0);
    const rata = valid.length ? Math.round((valid.reduce((a, t) => a + t.akurasi, 0) / valid.length) * 10) / 10 : null;
    const delta = valid.length >= 2 ? Math.round((valid[valid.length - 1].akurasi - valid[valid.length - 2].akurasi) * 10) / 10 : null;
    const lb = s.lb === 'validator' ? d.summary.leaderboardValidator : d.summary.leaderboard;
    return h`${atas}
      ${this.belumSegar(s)}
      ${!k.total ? h`<section class="card">${kosong('grafik', 'Belum ada hitungan', s.mode === 'daily' ? 'Tidak ada data cycle count pada tanggal ini.' : 'Tidak ada data cycle count pada bulan ini.')}</section>` : h`
      <section class="card rp-akurasi">
        <div class="rp-akurasi__utama"><b class="angka" style="color:${akurasi >= 99 ? 'var(--hijau-teks)' : 'var(--tinta)'}">${desimal(akurasi)}%</b><span>Akurasi ${s.mode === 'daily' ? tanggalPendek(s.tanggal) : bulanPendek(s.bulan)}</span></div>
        <div class="rp-tiga"><div><b class="angka">${angka(k.total)}</b><span>Total hitung</span></div><div><b class="angka" style="color:var(--hijau-teks)">${angka(k.hit)}</b><span>Hit</span></div><div><b class="angka" style="color:var(--merah-teks)">${angka(k.discrepancy)}</b><span>Discrepancy</span></div></div>
        ${tunggu0 ? h`<div class="ket">${angka(tunggu0)} item masih menunggu validasi dan belum dihitung sebagai hit atau discrepancy.</div>` : ''}
      </section>`}
      <section class="card hm-kartu">
        <h2 class="card__judul">Tren akurasi</h2><p class="card__ket">${s.mode === 'monthly' ? '6 bulan terakhir.' : '7 hari terakhir.'} Ketuk batang untuk melihat angkanya.</p>
        <svg class="grafik" id="rpTren" style="margin-top:8px"></svg>
        <div class="rp-info" id="rpTrenInfo"></div>
        <div class="legenda" style="margin-top:10px"><span><i style="background:var(--hijau)"></i>Hit</span><span><i style="background:var(--merah)"></i>Discrepancy</span><span><i style="background:var(--tinta);border-radius:50%"></i>Akurasi</span></div>
        ${rata != null ? h`<div class="rp-tiga rp-tiga--garis"><div><b class="angka">${desimal(rata)}%</b><span>Rata-rata akurasi</span></div><div><b class="angka">${angka(d.trend.reduce((a, t) => a + t.discrepancy, 0))}</b><span>Total discrepancy</span></div>
          <div><b class="angka" style="color:${delta > 0 ? 'var(--hijau-teks)' : delta < 0 ? 'var(--merah-teks)' : 'inherit'}">${delta == null ? '-' : (delta > 0 ? '+' : '') + desimal(delta)}</b><span>Poin dari periode lalu</span></div></div>` : ''}
      </section>
      ${s.backlog && s.backlog.length ? h`<section class="card hm-kartu"><h2 class="card__judul">Outstanding Cycle</h2><p class="card__ket">Tugas yang belum dihitung, per tanggal upload. Ketuk untuk melihat petugasnya.</p>
        <div class="rp-daftar">${s.backlog.map((b) => h`<button type="button" class="rp-baris" data-aksi="report-backlog" data-tgl="${b.tanggal}"><span>${tanggalPanjang(String(b.tanggal).slice(0, 10))}</span><span class="pill pill--jingga">${angka(b.pending)} belum</span>${ikon('kanan', 'ic--kecil')}</button>`)}</div></section>` : ''}
      ${k.total ? h`<section class="card hm-kartu">
        <h2 class="card__judul">Ranking akurasi</h2>
        <div class="seg" style="margin:10px 0 4px"><button type="button" class="${s.lb === 'petugas' ? 'is-on' : ''}" data-aksi="report-lb" data-v="petugas">Petugas</button><button type="button" class="${s.lb === 'validator' ? 'is-on' : ''}" data-aksi="report-lb" data-v="validator">Validator</button></div>
        <p class="card__ket">${s.lb === 'petugas' ? 'Dari item yang hasilnya sudah final. Ketuk nama untuk ringkasannya.' : 'Dari item yang verifikasinya sudah selesai. Validator dianggap salah bila hitung ulangnya yang keliru.'}</p>
        ${lb && lb.length ? h`<div class="rp-daftar">${lb.map((p, i) => this.barisRank(p, i, s.lb === 'petugas'))}</div>` : h`<p class="redup" style="padding:14px 0 4px">Belum ada data untuk ranking ini.</p>`}
      </section>` : ''}
      ${d.errorAnalysis && d.errorAnalysis.length ? h`<section class="card hm-kartu"><h2 class="card__judul">Analisis kesalahan</h2><p class="card__ket">Selisih per kategori penyebab. Ketuk untuk melihat itemnya.</p>
        <div class="rp-daftar">${d.errorAnalysis.map((a) => h`<button type="button" class="rp-batang" data-aksi="report-kategori" data-alasan="${a.alasan}"><span class="rp-batang__atas"><b>${a.alasan}</b><span>${angka(a.jumlah)} item, ${desimal(a.persen)}%</span></span><span class="bar bar--tipis bar--merah"><i style="width:${a.persen}%"></i></span></button>`)}</div></section>` : ''}
      ${k.total ? h`<section class="card hm-kartu"><h2 class="card__judul">Item bermasalah</h2><p class="card__ket">Item yang selisih, menunggu validasi, atau sedang diverifikasi pada periode ini.</p>
        <div id="rpMasalah">${this.masalahHtml()}</div></section>
      <button type="button" class="btn btn--teks" data-aksi="report-ekspor" style="margin:10px auto 0">${ikon('unduh', 'ic--kecil')}Ekspor ke Excel</button>` : ''}`;
  },
  barisRank(p, i, bisaDiketuk) {
    const dalam = h`<span class="rp-rank__no">${i + 1}</span><span class="rp-rank__isi"><b>${p.nama}</b><span>${angka(p.hit)} benar dari ${angka(p.total)} item</span></span>
      <span class="rp-rank__persen" style="color:${p.akurasi >= 99 ? 'var(--hijau-teks)' : 'var(--merah-teks)'}">${desimal(p.akurasi)}%</span>`;
    return bisaDiketuk ? h`<button type="button" class="rp-rank" data-aksi="report-user" data-nama="${p.nama}">${dalam}</button>` : h`<div class="rp-rank">${dalam}</div>`;
  },
  masalahHtml() {
    const s = this.dash;
    if (s.detailSibuk) return memuat('Memuat item…');
    if (!s.detail) return h`<button type="button" class="btn btn--tenang" data-aksi="report-masalah" style="margin-top:12px">Tampilkan item bermasalah</button>`;
    if (!s.detail.length) return h`<p class="redup" style="padding-top:12px">Tidak ada item bermasalah pada periode ini.</p>`;
    const q = (s.cari || '').trim().toLowerCase();
    const daftar = s.detail.filter((x) => !q || cocok([x.lokasi, x.article, x.description, x.namaPetugas, x.status].join(' '), q));
    const pil = (st) => (/Selesai/.test(st) ? 'hijau' : /Verifikasi/.test(st) ? 'merah' : /Menunggu/.test(st) ? 'ungu' : 'jingga');
    return h`<div class="cari" style="margin:12px 0 10px">${ikon('cari', 'ic--kecil')}<input class="field" type="search" id="rpCari" placeholder="Cari lokasi, SKU, atau nama" value="${s.cari || ''}" autocomplete="off"></div>
      <div class="rp-daftar">${daftar.slice(0, 60).map((x) => h`<div class="rp-item"><div class="vf-row__atas">${labelRak(x.lokasi)}<span class="row__t">${x.article}</span><span class="pill vf-selisih">${x.selisih === '' || x.selisih == null ? '-' : bertanda(x.selisih)}</span></div>
        <div class="row__s">${x.description}</div><div class="rp-item__bawah"><span class="redup">${x.namaPetugas}</span><span class="pill pill--${pil(x.status)}">${x.status}</span></div></div>`)}</div>
      ${daftar.length > 60 ? h`<p class="redup" style="text-align:center;margin-top:10px">Dan ${angka(daftar.length - 60)} item lagi. Persempit dengan pencarian atau ekspor ke Excel.</p>` : ''}${!daftar.length ? h`<p class="redup" style="padding:10px 0">Tidak ada yang cocok.</p>` : ''}`;
  },
  // diam: daftar yang sudah terbuka diperbarui di belakang (setelah Dashboard diperbarui) tanpa mengosongkannya dulu.
  async muatMasalah(diam) {
    const s = this.dash, no = s.req;
    if (!diam) { s.detailSibuk = true; isi(el('rpMasalah'), this.masalahHtml()); }
    let hasil = null, salah = null;
    try { hasil = await panggil('getProblemItemsDetail', s.mode, s.mode === 'daily' ? s.tanggal : s.bulan, Sesi.user.username) || []; } catch (e) { salah = e; }
    if (this.dash !== s || s.req !== no) return; // periode sudah diganti: daftar ini milik periode lama
    s.detailSibuk = false;
    if (!salah) s.detail = hasil; else if (!diam) Toast.galat(salah);
    if (Nav.aktif() === 'report' && this.tampilan === 'dashboard') { const y = window.scrollY; this.gambar(); window.scrollTo(0, y); }
  },
  async bukaBacklog(tgl) {
    const l = Lembar.buka({ judul: 'Belum di cycle', isi: memuat() });
    try {
      const u = await panggil('getBacklogDetailByDate', tgl, Sesi.user.username);
      const maks = Math.max(1, Math.max.apply(null, u.map((x) => x.pending).concat([1])));
      isi(l.isi, u.length ? h`<p class="redup" style="margin-bottom:10px">Tugas dari upload ${tanggalPanjang(String(tgl).slice(0, 10))}.</p>${u.map((x) => h`<div class="orang"><div class="orang__atas"><b>${x.namaPetugas}</b><span class="angka">${angka(x.pending)}</span></div><div class="bar bar--tipis bar--jingga"><i style="width:${(x.pending / maks) * 100}%"></i></div><div class="orang__ket">dari ${angka(x.total)} tugas</div></div>`)}` : kosong('cek', 'Semua sudah selesai', ''));
    } catch (e) { isi(l.isi, h`<p class="pesan pesan--galat">${e.message}</p>`); }
  },
  async bukaUser(nama) {
    const s = this.dash, l = Lembar.buka({ judul: nama, isi: memuat() });
    try {
      const r = await panggil('getUserDashboardDetail', nama, s.mode, s.mode === 'daily' ? s.tanggal : s.bulan, Sesi.user.username), x = r.summary, inv = r.investigasi;
      isi(l.isi, h`<div class="rp-empat"><div><b class="angka">${angka(x.total)}</b><span>Total item</span></div><div><b class="angka" style="color:var(--merah-teks)">${angka(x.kesalahanHitung)}</b><span>Salah hitung</span></div>
        <div><b class="angka" style="color:var(--hijau-teks)">${desimal(x.akurasi)}%</b><span>Akurasi</span></div><div><b class="angka">${desimal(x.persenDiscrepancy)}%</b><span>Discrepancy</span></div></div>
        ${inv && inv.totalSelesai ? h`<h3 class="bag" style="margin-top:18px"><span>Verifikasi diselesaikan</span></h3><div class="rp-empat rp-empat--dua"><div><b class="angka">${angka(inv.totalSelesai)}</b><span>Task selesai</span></div><div><b class="angka">${desimal(inv.avgDurasiJam)} jam</b><span>Rata-rata buka sampai tutup</span></div></div>
          <p class="ket">Durasi dihitung dari task dibuka sampai ditutup, termasuk waktu antre. Jangan dipakai sebagai satu-satunya ukuran kecepatan.</p>` : ''}`);
    } catch (e) { isi(l.isi, h`<p class="pesan pesan--galat">${e.message}</p>`); }
  },
  async bukaKategori(alasan) {
    const s = this.dash, l = Lembar.buka({ judul: alasan, isi: memuat() });
    try {
      const it = await panggil('getErrorAnalysisDetail', alasan, s.mode, s.mode === 'daily' ? s.tanggal : s.bulan, Sesi.user.username);
      isi(l.isi, it.length ? h`<div class="list">${it.map((x) => h`<div class="row">${labelRak(x.lokasi)}<div class="row__isi"><div class="row__t">${x.article}</div><div class="row__s">${x.description}</div><div class="row__s">Sistem ${angka(x.qtySystem)}, hitung ${angka(x.qtyCount)}</div><div class="row__s">Dihitung oleh ${x.namaPetugas}</div></div></div>`)}</div>` : kosong('daftar', 'Tidak ada item', ''));
    } catch (e) { isi(l.isi, h`<p class="pesan pesan--galat">${e.message}</p>`); }
  },
  eksporDash() {
    const s = this.dash, d = s.data; if (!d) return;
    const rank = (x) => (x || []).map((p) => ({ Nama: p.nama, 'Total Item': p.total, 'Hit (Benar)': p.hit, 'Akurasi (%)': p.akurasi, 'Discrepancy (%)': p.persenDiscrepancy }));
    const lembar = [{ nama: 'Ranking Petugas', baris: rank(d.summary.leaderboard) }, { nama: 'Ranking Validator', baris: rank(d.summary.leaderboardValidator) },
      { nama: 'Analisis Kesalahan', baris: (d.errorAnalysis || []).map((a) => ({ Kategori: a.alasan, Jumlah: a.jumlah, 'Persen (%)': a.persen })) }];
    if (s.detail) lembar.push({ nama: 'Item Bermasalah', baris: s.detail.map((x) => ({ Tanggal: x.tanggal, Lokasi: x.lokasi, SKU: x.article, Deskripsi: x.description, Selisih: x.selisih, Petugas: x.namaPetugas, Status: x.status })) });
    eksporExcel('Dashboard_CycleCount_' + (s.mode === 'monthly' ? 'Bulanan_' + s.bulan : 'Harian_' + s.tanggal) + '.xlsx', lembar);
  },

  /* ================= Productivity ================= */
  muatProd() {
    const s = this.prod, tanggal = s.tanggal;
    return this.ambil('prod', tanggal, () => panggil('getProductivity', tanggal, Sesi.user.username, -new Date().getTimezoneOffset()), (r) => { s.data = r; });
  },
  prodHtml() {
    const s = this.prod, d = s.data;
    const atas = h`<div class="rp-periode"><input type="date" class="field field--kecil" id="rpProdTanggal" value="${s.tanggal}" max="${hariIni()}" aria-label="Tanggal"></div>`;
    if (!d) return h`${atas}${s.galat ? (s.galat.jenis === 'fungsi-tak-ada' ? h`<section class="card">${kosong('kilat', 'Belum tersedia', 'Laporan ini tersedia setelah admin aplikasi memperbarui sistem.')}</section>` : gagalMuat(s.galat, 'report-segar')) : kerangka(3, 120)}`;
    const r = d.ringkasan, aktif = d.petugas.filter((p) => p.total > 0), maks = Math.max(1, Math.max.apply(null, d.petugas.map((p) => p.total).concat([1])));
    if (!r.totalItem && !r.sisa) return h`${atas}${this.belumSegar(s)}<section class="card">${kosong('kilat', 'Belum ada hitungan', 'Tidak ada item yang dihitung pada ' + tanggalPanjang(s.tanggal) + '.')}</section>`;
    return h`${atas}
      ${this.belumSegar(s)}
      <div class="rp-tiga rp-tiga--kartu"><div><b class="angka">${angka(r.totalItem)}</b><span>Item dihitung</span></div><div><b class="angka">${angka(r.petugasAktif)}</b><span>Petugas aktif</span></div><div><b class="angka">${desimal(r.perJam)}</b><span>Item per jam per orang</span></div></div>
      ${d.perJam.length ? h`<section class="card hm-kartu"><h2 class="card__judul">Hitungan per jam</h2><p class="card__ket">Jumlah item yang dihitung tiap jam, semua petugas.</p><svg class="grafik" id="rpJam" style="margin-top:10px"></svg></section>` : ''}
      <section class="card hm-kartu"><h2 class="card__judul">Per petugas</h2><p class="card__ket">Per jam dihitung dari jam yang benar-benar berisi hitungan, jadi istirahat tidak menurunkan angka.</p>
        <div class="rp-daftar">${d.petugas.map((p) => h`<div class="orang"><div class="orang__atas"><b>${p.nama}${p.role ? h` <span class="pill">${namaPeran(p.role)}</span>` : ''}</b><span class="angka">${angka(p.total)}</span></div>
          <div class="bar bar--tipis bar--biru"><i style="width:${(p.total / maks) * 100}%"></i></div>
          <div class="orang__ket">${p.total ? h`${jam(p.mulai)} sampai ${jam(p.akhir)}, ${desimal(p.perJam)} per jam${p.selisih ? h`, ${angka(p.selisih)} selisih awal` : ''}` : 'Belum mulai menghitung'}${p.sisa ? h` <span class="pill pill--jingga">${angka(p.sisa)} belum</span>` : ''}</div></div>`)}</div></section>
      ${d.validator.length ? h`<section class="card hm-kartu"><h2 class="card__judul">Validasi</h2><p class="card__ket">${angka(r.totalValidasi)} item divalidasi pada tanggal ini.</p>
        <div class="rp-daftar">${d.validator.map((x) => h`<div class="rp-rank"><span class="rp-rank__isi"><b>${x.nama}</b><span>${jam(x.mulai)} sampai ${jam(x.akhir)}, ${angka(x.hit)} ternyata cocok</span></span><span class="rp-rank__persen">${angka(x.total)}</span></div>`)}</div></section>` : ''}
      ${aktif.length ? h`<button type="button" class="btn btn--teks" data-aksi="report-ekspor" style="margin:10px auto 0">${ikon('unduh', 'ic--kecil')}Ekspor ke Excel</button>` : ''}`;
  },
  eksporProd() {
    const d = this.prod.data; if (!d) return;
    eksporExcel('Productivity_' + d.tanggal + '.xlsx', [
      { nama: 'Petugas', baris: d.petugas.map((p) => ({ Nama: p.nama, Peran: namaPeran(p.role), 'Item Dihitung': p.total, 'Cocok (awal)': p.hit, 'Selisih (awal)': p.selisih, Mulai: p.mulai ? jam(p.mulai) : '', Terakhir: p.akhir ? jam(p.akhir) : '', 'Jam Aktif': p.jamAktif, 'Item per Jam': p.perJam, 'Belum Dihitung': p.sisa })) },
      { nama: 'Validator', baris: d.validator.map((x) => ({ Nama: x.nama, 'Item Divalidasi': x.total, 'Ternyata Cocok': x.hit, Mulai: jam(x.mulai), Terakhir: jam(x.akhir) })) },
      { nama: 'Per Jam', baris: d.perJam.map((j) => ({ Jam: dua(j.jam) + '.00', 'Item Dihitung': j.total })) }]);
  },

  /* ================= Analytics ================= */
  muatAn() {
    const s = this.an;
    if (s.dari > s.sampai) { const t = s.dari; s.dari = s.sampai; s.sampai = t; }
    const dari = s.dari, sampai = s.sampai;
    return this.ambil('an', dari + ':' + sampai, () => panggil('getAnalyticsRootCauseData', dari, sampai, Sesi.user.username), (r) => { s.data = r; });
  },
  anHtml() {
    const s = this.an, d = s.data;
    const atas = h`<div class="rentang"><label><span>Dari</span><input type="date" class="field field--kecil" id="rpDari" value="${s.dari}" max="${hariIni()}"></label><label><span>Sampai</span><input type="date" class="field field--kecil" id="rpSampai" value="${s.sampai}" max="${hariIni()}"></label></div>`;
    if (!d) return h`${atas}${s.galat ? gagalMuat(s.galat, 'report-segar') : kerangka(3, 130)}`;
    if (!d.kpi.total) return h`${atas}${this.belumSegar(s)}<section class="card">${kosong('tren', 'Belum ada data', 'Tidak ada cycle count pada rentang tanggal ini.')}</section>`;
    const akar = d.rootCause || [], top = akar.slice().sort((a, b) => b.jumlah - a.jumlah).filter((x) => x.jumlah > 0);
    const kontrib = (d.contributors && d.contributors[s.kontrib]) || [], maks = Math.max(1, Math.max.apply(null, kontrib.map((x) => x.jumlah).concat([1])));
    return h`${atas}
      ${this.belumSegar(s)}
      <div class="rp-tiga rp-tiga--kartu rp-tiga--dua"><div><b class="angka" style="color:var(--hijau-teks)">${desimal(d.kpi.akurasi)}%</b><span>Akurasi</span></div><div><b class="angka" style="color:var(--merah-teks)">${desimal(d.kpi.errorPersen)}%</b><span>${angka(d.kpi.discrepancy)} item discrepancy</span></div></div>
      ${d.trend.length > 1 ? h`<section class="card hm-kartu"><h2 class="card__judul">Tren akurasi</h2><p class="card__ket">${angka(d.trend.length)} hari dalam rentang terpilih.</p><svg class="grafik" id="rpAnTren" style="margin-top:10px"></svg></section>` : ''}
      <section class="card hm-kartu"><h2 class="card__judul">Akar masalah</h2><p class="card__ket">Selisih dikelompokkan menurut penyebabnya. Ketuk untuk melihat itemnya.</p>
        <div class="rp-donat">${Grafik.donat(akar.map((x) => ({ nilai: x.jumlah, warna: WARNA_AKAR[x.label] || 'var(--abu-2)' })), angka(d.kpi.discrepancy), 'discrepancy')}
          <div class="rp-daftar rp-daftar--rapat">${akar.map((x) => h`<button type="button" class="rp-akar" data-aksi="report-akar" data-label="${x.label}" ${x.jumlah ? '' : 'disabled'}><i style="background:${WARNA_AKAR[x.label] || 'var(--abu-2)'}"></i><span>${x.label}</span><b class="angka">${angka(x.jumlah)}</b></button>`)}</div></div>
      </section>
      <section class="card hm-kartu"><h2 class="card__judul">Penyumbang terbesar</h2>
        <div class="seg" style="margin:10px 0 10px">${[['userCycle', 'Petugas'], ['validator', 'Validator'], ['operasional', 'Masalah']].map((x) => h`<button type="button" class="${s.kontrib === x[0] ? 'is-on' : ''}" data-aksi="report-kontrib" data-v="${x[0]}">${x[1]}</button>`)}</div>
        ${kontrib.length ? h`<div class="rp-daftar">${kontrib.map((c) => h`<div class="orang"><div class="orang__atas"><b>${c.nama}</b><span class="angka">${angka(c.jumlah)}</span></div><div class="bar bar--tipis bar--merah"><i style="width:${(c.jumlah / maks) * 100}%"></i></div><div class="orang__ket">${desimal(c.persen)}% dari semua discrepancy</div></div>`)}</div>` : h`<p class="redup">Belum ada data pada rentang ini.</p>`}
      </section>
      <section class="card hm-kartu"><h2 class="card__judul">Saran tindakan</h2>
        ${top.length ? h`<p class="card__ket">${desimal(top[0].persen)}% selisih berasal dari ${top[0].label}. Mulai dari sana.</p>` : h`<p class="card__ket">Tidak ada discrepancy pada rentang ini.</p>`}
        <div class="rp-daftar">${top.slice(0, 2).map((x) => SARAN_AKAR[x.label]).filter(Boolean).concat([['grafik', 'Pantau lokasi yang berulang', 'Perhatikan lokasi dan SKU yang sering muncul di daftar selisih.']]).map((a) => h`<div class="rp-saran">${ikon(a[0])}<div><b>${a[1]}</b><span>${a[2]}</span></div></div>`)}</div>
      </section>`;
  },
  async bukaAkar(label) {
    const s = this.an, l = Lembar.buka({ judul: label, penuh: true, isi: memuat() });
    try {
      const rows = await panggil('getAnalyticsRootCauseDetail', label, s.dari, s.sampai, Sesi.user.username);
      let tampil = 12;
      const gambarIsi = () => {
        isi(l.isi, rows.length ? h`<p class="redup" style="margin-bottom:10px">${angka(rows.length)} item, terbaru di atas.</p><div class="rp-daftar">${rows.slice(0, tampil).map((x) => h`<div class="rp-item"><div class="vf-row__atas">${labelRak(x.lokasi)}<span class="row__t">${x.sku}</span><span class="pill pill--merah vf-selisih">${angka(x.qtyError)}</span></div>
          <div class="row__s row__s--lipat">${x.description}</div>
          <div class="rinci rinci--rapat"><div class="rinci__baris"><span>Tanggal</span><b>${tanggalPendek(x.tanggal)}</b></div><div class="rinci__baris"><span>Petugas</span><b>${x.userCycle}</b></div><div class="rinci__baris"><span>Validator</span><b>${x.validator}</b></div><div class="rinci__baris"><span>User WMS</span><b>${x.wmsUser} (${x.transaksi})</b></div>${x.kategori && x.kategori !== '-' ? h`<div class="rinci__baris"><span>Kategori</span><b>${x.kategori}</b></div>` : ''}</div>
          ${x.hasilInvestigasi ? h`<div class="vf-catatan">${x.hasilInvestigasi}</div>` : ''}</div>`)}</div>
          ${rows.length > tampil ? h`<button type="button" class="btn btn--tenang" id="rpAkarLagi" style="margin-top:12px">Tampilkan ${angka(Math.min(12, rows.length - tampil))} lagi</button>` : ''}` : kosong('daftar', 'Tidak ada item', 'Tidak ada selisih dengan penyebab ini pada rentang tanggal terpilih.'));
        const b = el('rpAkarLagi'); if (b) b.addEventListener('click', () => { tampil += 12; gambarIsi(); });
      };
      gambarIsi();
    } catch (e) { isi(l.isi, h`<p class="pesan pesan--galat">${e.message}</p>`); }
  }
};

Report.reset();
Layar.report = { masuk() { Report.gambar(); Report.buka(); } };
Aksi['report-segar'] = () => Report.segar();
Aksi['report-tampilan'] = (t) => { Report.tampilan = t.dataset.v; Report.gambar(); Report.buka(); };
Aksi['report-mode'] = (t) => { Report.dash.mode = t.dataset.v; Report.muatDash(); };
Aksi['report-lb'] = (t) => { Report.dash.lb = t.dataset.v; const y = window.scrollY; Report.gambar(); window.scrollTo(0, y); };
Aksi['report-backlog'] = (t) => Report.bukaBacklog(t.dataset.tgl);
Aksi['report-user'] = (t) => Report.bukaUser(t.dataset.nama);
Aksi['report-kategori'] = (t) => Report.bukaKategori(t.dataset.alasan);
Aksi['report-masalah'] = () => Report.muatMasalah(false);
Aksi['report-ekspor'] = () => (Report.tampilan === 'productivity' ? Report.eksporProd() : Report.eksporDash());
Aksi['report-kontrib'] = (t) => { Report.an.kontrib = t.dataset.v; const y = window.scrollY; Report.gambar(); window.scrollTo(0, y); };
Aksi['report-akar'] = (t) => Report.bukaAkar(t.dataset.label);
