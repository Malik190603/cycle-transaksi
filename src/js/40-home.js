/*
 * Home. Satu panggilan server (getHomeBundle) mengisi seluruh layar; data terakhir disimpan di HP
 * supaya Home langsung tampil saat aplikasi dibuka, lalu diperbarui di belakang.
 */
const Home = {
  data: null, at: 0, sibuk: false, galat: null, req: 0,

  // Dipanggil tiap pengguna atau facility berganti: jawaban dari permintaan yang masih berjalan
  // (milik pengguna/facility sebelumnya) diabaikan lewat nomor permintaan.
  mulai() { const s = Simpanan.get('home'); this.data = s ? s.d : null; this.at = s ? s.at : 0; this.galat = null; this.sibuk = false; this.req++; },

  async muat(paksa) {
    if (this.sibuk || !Sesi.user) return;
    if (!paksa && this.data && Date.now() - this.at < 45000) return;
    const id = ++this.req, siapa = Sesi.user.username;
    this.sibuk = true; this.tandaSegar();
    try {
      let d;
      try { d = await panggil('getHomeBundle', siapa); }
      catch (e) { if (e.jenis !== 'fungsi-tak-ada') throw e; d = await this.muatLama(siapa); }
      if (id !== this.req) return;
      this.data = d; this.at = Date.now(); this.galat = null;
      Simpanan.set('home', { d, at: this.at });
    } catch (e) {
      if (id !== this.req) return;
      if (e.jenis === 'sesi') { keluarAplikasi('Akun ini tidak lagi terdaftar. Hubungi admin.'); return; }
      this.galat = e;
      if (this.data && Nav.aktif() === 'home') Toast.galat(e);
    } finally {
      if (id === this.req) {
        this.sibuk = false;
        this.lencana();
        if (Nav.aktif() === 'home') this.gambar();
      }
    }
  },
  // Backend lama (sebelum getHomeBundle): rakit dari fungsi-fungsi terpisah.
  async muatLama(siapa) {
    const tugas = [panggil('getHomeSummary', siapa), panggil('getMyPendingCount', siapa)];
    if (Boleh.validasi()) tugas.push(panggil('getPlusMinusSummary', siapa));
    if (Boleh.admin()) tugas.push(panggil('getPendingBacklog', siapa));
    const r = await Promise.all(tugas);
    const sisaSemua = Boleh.admin() ? (r[3] || []).reduce((s, b) => s + b.pending, 0) : r[1];
    return Object.assign({}, r[0], { role: peran(), serverTime: Date.now(), myPending: r[1], outstanding: sisaSemua, pendingValidasi: r[0].belumValidasi, openTasks: 0, plusMinus: r[2] || null, perPetugas: [], lama: true });
  },
  lencana() {
    const d = this.data; if (!d) return;
    const c = Tugas.sisa('cycle'), v = Tugas.sisa('validasi');
    Nav.setLencana('cycle', c != null ? c : d.myPending);
    Nav.setLencana('validasi', v != null ? v : d.pendingValidasi);
    Nav.setLencana('verifikasi', d.openTasks);
  },
  // Tugas cycle saya yang belum dihitung. Daftar tugas di HP lebih baru daripada angka server
  // (server baru mencatat hasil sekitar semenit setelah dikirim), jadi dipakai bila sudah dimuat.
  sisaSaya(d) {
    const lokal = Tugas.sisa('cycle');
    return lokal != null ? lokal : Math.max(0, d.myPending - Antrean.jumlah('cycle'));
  },
  tandaSegar() { const b = el('hmSegar'); if (b) b.classList.toggle('is-putar', this.sibuk); },

  /* ---------- gambar ---------- */
  gambar() {
    const u = Sesi.user; if (!u) return;
    const d = this.data;
    isi(el('scr-home'), h`
      <header class="hm-kepala">
        <button type="button" class="hm-avatar" data-aksi="akun" aria-label="Akun dan pengaturan">${inisial(u.username)}</button>
        <div class="hm-siapa">
          <div class="hm-sapa">${sapaan()},</div>
          <div class="hm-nama">${u.username}</div>
          <div class="hm-tempat">${u.facilityName || 'Belum punya facility'}</div>
          <span class="pill pill--biru">${namaPeran(u.role)}</span>
        </div>
        <button type="button" class="hm-segar${this.sibuk ? ' is-putar' : ''}" id="hmSegar" data-aksi="home-segar" aria-label="Perbarui data">${ikon('segar', 'ic--kecil')}<span>${this.at ? jam(this.at) : 'Perbarui'}</span></button>
      </header>
      <div id="hmPembaruan"></div>
      ${!u.facilityName ? h`<div class="info info--jingga" style="margin-bottom:12px">${ikon('awas')}<div>Akun ini belum dimasukkan ke facility mana pun, jadi belum bisa menerima tugas atau mengirim hasil. Minta admin mengaturnya di Config.</div></div>` : ''}
      ${d ? (d.tanpaFacility ? '' : this.badan(d)) : (this.galat ? gagalMuat(this.galat, 'home-segar') : h`<div class="kerangka" style="height:74px;border-radius:var(--r-kartu);margin-bottom:12px"></div><div class="kpi-kisi">${[1, 2, 3, 4].map(() => h`<div class="kerangka" style="height:112px;border-radius:var(--r-kartu)"></div>`)}</div>`)}
    `);
    Pembaruan.gambarKartu();
  },
  badan(d) {
    if (Boleh.admin()) return h`${this.berikutnya(d)}${this.kpiAdmin(d)}${this.progres(d)}${this.plusMinus(d)}`;
    if (Boleh.validasi()) return h`${this.berikutnya(d)}${this.kpiInventory(d)}${this.plusMinus(d)}`;
    return this.tugasSaya(d);
  },

  /* Satu saran tindakan, dipilih dari keadaan data: aplikasi yang menunjukkan apa yang perlu dikerjakan. */
  berikutnya(d) {
    const total = d.totalCycleHariIni + d.outstanding, punyaSaya = this.sisaSaya(d);
    let k;
    if (punyaSaya > 0) k = ['biru', 'kotak', angka(punyaSaya) + ' item menunggu Anda hitung', 'Lanjutkan cycle dari lokasi berikutnya.', 'home-mulai-cycle'];
    else if (Boleh.admin() && total === 0) k = ['biru', 'unggah', 'Belum ada tugas hari ini', 'Upload data transaksi dan stok untuk membagi tugas.', 'home-ke-upload'];
    else if (Boleh.admin() && d.outstanding > 0) k = ['jingga', 'jam', angka(d.outstanding) + ' item belum di cycle', 'Lihat siapa yang masih punya tugas.', 'home-rinci-sisa'];
    else if (d.pendingValidasi > 0) k = ['ungu', 'perisai', angka(d.pendingValidasi) + ' item menunggu validasi', Boleh.admin() ? 'Selisih hitung perlu dihitung ulang oleh Inventory.' : 'Hitung ulang item yang selisih.', 'home-ke-validasi'];
    else if (d.openTasks > 0) k = ['merah', 'selisih', angka(d.openTasks) + ' selisih belum selesai', 'Tindak lanjuti di Verifikasi.', 'home-ke-verifikasi'];
    else if (total > 0 || d.selesaiHariIni > 0) k = ['hijau', 'cek', 'Semua beres untuk saat ini', 'Tidak ada tugas, validasi, atau selisih yang menunggu.', ''];
    else return '';
    const dalam = h`<span class="lanjut__ic">${ikon(k[1])}</span><span class="lanjut__teks"><b>${k[2]}</b><span>${k[3]}</span></span>${k[4] ? ikon('kanan') : ''}`;
    return k[4] ? h`<button type="button" class="lanjut lanjut--${k[0]}" data-aksi="${k[4]}">${dalam}</button>` : h`<div class="lanjut lanjut--${k[0]}">${dalam}</div>`;
  },

  kpi(warna, ik, judul, nilai, ket, aksi) {
    return h`<button type="button" class="kpi kpi--${warna}" data-aksi="${aksi}"><span class="kpi__atas"><span class="kpi__ic">${ikon(ik)}</span><span class="kpi__judul">${judul}</span></span><span class="kpi__nilai angka">${angka(nilai)}</span><span class="kpi__ket">${ket}</span></button>`;
  },
  kpiAdmin(d) {
    return h`<div class="kpi-kisi">
      ${this.kpi('biru', 'kotak', 'Total Cycle', d.totalCycleHariIni, 'Item sudah di cycle', 'home-rinci-cycle')}
      ${this.kpi('jingga', 'jam', 'Outstanding Cycle', d.outstanding, 'Item belum di cycle', 'home-rinci-sisa')}
      ${this.kpi('ungu', 'perisai', 'Belum Validasi', d.belumValidasi, 'Item menunggu validasi', 'home-ke-validasi')}
      ${this.kpi('hijau', 'cek', 'Selesai', d.selesaiHariIni, 'Item selesai', 'home-ke-report')}
    </div>`;
  },
  kpiInventory(d) {
    return h`<div class="kpi-kisi">
      ${this.kpi('ungu', 'perisai', 'Belum Validasi', d.pendingValidasi, 'Item menunggu Anda validasi', 'home-ke-validasi')}
      ${this.kpi('merah', 'selisih', 'Verifikasi', d.openTasks, 'Selisih belum selesai', 'home-ke-verifikasi')}
      ${this.kpi('jingga', 'jam', 'Tugas Cycle', this.sisaSaya(d), 'Item belum Anda hitung', 'home-ke-cycle')}
      ${this.kpi('hijau', 'cek', 'Sudah Dihitung', d.totalCycleHariIni, 'Item Anda hitung hari ini', 'home-ke-cycle')}
    </div>`;
  },

  progres(d) {
    const sudah = d.totalCycleHariIni, sisa = d.outstanding, total = sudah + sisa;
    if (!total) return h`<section class="card hm-kartu"><h2 class="card__judul">Progress cycle hari ini</h2><p class="card__ket">Belum ada item dalam antrean. Angka di sini muncul setelah data di-upload.</p></section>`;
    return h`<section class="card hm-kartu">
      <h2 class="card__judul">Progress cycle hari ini</h2>
      <p class="card__ket">Item yang sudah dan belum di cycle hari ini.</p>
      <div class="hm-sisa ${sisa ? '' : 'hm-sisa--beres'}"><b class="angka">${sisa ? angka(sisa) : ''}</b><span>${sisa ? 'item belum di cycle' : 'Semua item sudah di cycle'}</span></div>
      <div class="bar" role="progressbar" aria-valuemin="0" aria-valuemax="${total}" aria-valuenow="${sudah}"><i style="width:${Math.min(100, (sudah / total) * 100)}%"></i></div>
      <div class="hm-dua">
        <div><span class="hm-persen" style="color:var(--hijau-teks)">${persen(sudah, total)}</span><b class="angka">${angka(sudah)}</b><span>Item sudah di cycle</span></div>
        <div><span class="hm-persen" style="color:var(--jingga-teks)">${persen(sisa, total)} tersisa</span><b class="angka">${angka(total)}</b><span>Total item antrean</span></div>
      </div>
      ${sisa && d.perPetugas && d.perPetugas.length ? h`<button type="button" class="info info--jingga hm-catatan" data-aksi="home-rinci-sisa">${ikon('orang2')}<span>${this.ringkasSisa(d)}</span>${ikon('kanan', 'ic--kecil')}</button>` : ''}
    </section>`;
  },
  ringkasSisa(d) {
    const p = d.perPetugas.filter((x) => x.sisa > 0);
    const tanpa = d.belumDitugaskan ? h` ${angka(d.belumDitugaskan)} item belum ditugaskan ke siapa pun.` : '';
    if (!p.length) return tanpa || 'Lihat rincian per petugas.';
    const besar = p[0];
    return p.length === 1 ? h`Hanya <b>${besar.nama}</b> yang masih punya tugas, ${angka(besar.sisa)} item.${tanpa}` : h`<b>${p.length} petugas</b> masih punya tugas. Terbanyak <b>${besar.nama}</b>, ${angka(besar.sisa)} item.${tanpa}`;
  },

  plusMinus(d) {
    const pm = d.plusMinus; if (!pm) return '';
    const kosongPm = !pm.plus.totalSku && !pm.minus.totalSku;
    const ubin = (sisi, ik, nama, x) => h`<button type="button" class="pm pm--${sisi}" data-aksi="home-rinci-pm" data-sisi="${sisi}" ${x.totalSku ? '' : 'disabled'}>
      <span class="pm__nama">${nama}</span><span class="pm__nilai">${ikon(ik)}<b class="angka">${angka(x.totalQty)}</b><small>qty</small></span><span class="pm__sku">${angka(x.totalSku)} SKU</span></button>`;
    return h`<section class="card hm-kartu">
      <h2 class="card__judul">Summary plus minus</h2>
      <p class="card__ket">${kosongPm ? 'Semua selisih sudah selesai.' : 'Selisih yang belum selesai di Verifikasi. Ketuk untuk melihat per SKU.'}</p>
      <div class="pm-kisi">${ubin('plus', 'naik', 'Plus', pm.plus)}${ubin('minus', 'turun', 'Minus', pm.minus)}</div>
    </section>`;
  },

  /* Home petugas: satu hal saja, tugasnya sendiri. */
  tugasSaya(d) {
    const sisa = this.sisaSaya(d), total = Math.max(d.totalCycleHariIni + d.myPending, sisa), sudah = total - sisa;
    if (!total) return h`<section class="card">${kosong('daftar', 'Belum ada tugas untuk Anda', 'Tugas muncul di sini setelah admin meng-upload dan membagi data.')}</section>`;
    return h`<section class="card hm-tugas">
      <h2 class="card__judul">Tugas Anda hari ini</h2>
      <div class="hm-sisa ${sisa ? '' : 'hm-sisa--beres'}"><b class="angka">${sisa ? angka(sisa) : ''}</b><span>${sisa ? 'item belum dihitung' : 'Semua item sudah dihitung'}</span></div>
      <div class="bar bar--biru" role="progressbar" aria-valuemin="0" aria-valuemax="${total}" aria-valuenow="${sudah}"><i style="width:${(sudah / total) * 100}%"></i></div>
      <div class="hm-dari"><span><b class="angka">${angka(sudah)}</b> dari ${angka(total)} item sudah dihitung</span><span>${persen(sudah, total)}</span></div>
      ${sisa ? h`<button type="button" class="btn" data-aksi="home-mulai-cycle">${sudah ? 'Lanjutkan cycle' : 'Mulai cycle'}</button>` : ''}
    </section>
    ${d.selesaiHariIni < d.totalCycleHariIni ? h`<div class="info" style="margin-top:12px">${ikon('perisai')}<div><b>${angka(d.totalCycleHariIni - d.selesaiHariIni)} item</b> hasil hitung Anda menunggu validasi oleh Inventory.</div></div>` : ''}`;
  },

  /* ---------- rincian (lembar) ---------- */
  rinciPetugas(mode) {
    const d = this.data; if (!d) return;
    const daftar = (d.perPetugas || []).filter((p) => (mode === 'sisa' ? p.sisa > 0 : p.dihitung > 0))
      .sort((a, b) => (mode === 'sisa' ? b.sisa - a.sisa : b.dihitung - a.dihitung));
    const maks = Math.max(1, Math.max.apply(null, daftar.map((p) => (mode === 'sisa' ? p.sisa : p.dihitung)).concat([1])));
    const baris = daftar.map((p) => {
      const v = mode === 'sisa' ? p.sisa : p.dihitung;
      return h`<div class="orang"><div class="orang__atas"><b>${p.nama}</b><span class="angka">${angka(v)}</span></div>
        <div class="bar bar--tipis ${mode === 'sisa' ? 'bar--jingga' : 'bar--biru'}"><i style="width:${(v / maks) * 100}%"></i></div>
        <div class="orang__ket">${mode === 'sisa' ? h`Sudah dihitung ${angka(p.dihitung)} item` : h`${p.sisa ? h`Sisa ${angka(p.sisa)} item` : 'Tugas tuntas'}${p.final < p.dihitung ? h`, ${angka(p.dihitung - p.final)} menunggu validasi` : ''}`}</div></div>`;
    });
    Lembar.buka({
      judul: mode === 'sisa' ? 'Belum di cycle' : 'Cycle hari ini',
      isi: daftar.length
        ? h`<p class="redup" style="margin-bottom:14px">${mode === 'sisa' ? h`${angka(d.outstanding)} item masih menunggu, termasuk sisa hari sebelumnya.${d.belumDitugaskan ? h` ${angka(d.belumDitugaskan)} di antaranya belum ditugaskan ke siapa pun.` : ''}` : h`${angka(d.totalCycleHariIni)} item sudah dihitung oleh ${daftar.length} petugas.`}</p>${baris}`
        : (d.lama ? h`<p class="redup">Rincian per petugas tersedia setelah admin aplikasi memperbarui sistem.</p>` : kosong('cek', mode === 'sisa' ? 'Tidak ada sisa tugas' : 'Belum ada yang dihitung', mode === 'sisa' ? 'Semua item sudah di cycle.' : 'Angka muncul setelah petugas mulai menghitung.'))
    });
  },
  rinciPlusMinus(sisi) {
    const pm = this.data && this.data.plusMinus; if (!pm) return;
    const x = pm[sisi], plus = sisi === 'plus';
    Lembar.buka({
      judul: plus ? 'Selisih plus' : 'Selisih minus',
      isi: h`<p class="redup" style="margin-bottom:12px">${angka(x.totalQty)} qty pada ${angka(x.totalSku)} SKU, belum selesai di Verifikasi.</p>
        <div class="list">${x.items.map((it) => h`<div class="row"><div class="row__isi"><div class="row__t">${it.article}</div><div class="row__s row__s--lipat">${it.description}</div>
          <div class="pm-lokasi">${String(it.lokasi || '').split(', ').slice(0, 4).map((l) => labelRak(l))}${it.jumlahLokasi > 4 ? h`<span class="redup">+${it.jumlahLokasi - 4} lokasi</span>` : ''}</div></div>
          <div class="row__ekor"><span class="pill pill--${plus ? 'hijau' : 'merah'}">${plus ? '+' : '−'}${angka(it.qty)}</span></div></div>`)}</div>`,
      kaki: h`<button type="button" class="btn btn--tenang" data-aksi="ke" data-tab="verifikasi">Buka Verifikasi</button>`
    });
  }
};

Layar.home = { masuk() { Home.gambar(); Home.muat(false); } };
Aksi['home-segar'] = () => Home.muat(true);
Aksi['home-mulai-cycle'] = () => Nav.ke('cycle', { mulai: true });
Aksi['home-ke-cycle'] = () => Nav.ke('cycle');
Aksi['home-ke-upload'] = () => Nav.ke('upload');
Aksi['home-ke-validasi'] = () => Nav.ke('validasi');
Aksi['home-ke-verifikasi'] = () => Nav.ke('verifikasi');
Aksi['home-ke-report'] = () => Nav.ke('report');
Aksi['home-rinci-sisa'] = () => Home.rinciPetugas('sisa');
Aksi['home-rinci-cycle'] = () => Home.rinciPetugas('cycle');
Aksi['home-rinci-pm'] = (t) => Home.rinciPlusMinus(t.dataset.sisi);
