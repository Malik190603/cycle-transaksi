/*
 * Cycle dan Validasi memakai mesin yang sama: daftar tugas → mode fokus (satu lokasi per layar,
 * keypad bawaan). Hitungan bersifat buta: qty sistem tidak pernah ditampilkan, dan hasil cocok/selisih
 * baru muncul di riwayat sesi setelah hasil terkirim dan tidak bisa diurungkan lagi.
 */
const TAHAN_MS = 5000; // jendela "Urungkan" sebelum hasil dikirim ke server

const MODE = {
  cycle: {
    judul: 'Cycle', fungsi: 'getMyPendingTasks', warna: 'biru', ikon: 'kotak', satuan: 'Qty hasil hitung',
    kunci: (it) => it.no, args: (it, q) => [it.no, Sesi.user.username, q],
    ket: (n) => angka(n) + ' item menunggu dihitung',
    kosong: ['daftar', 'Belum ada tugas untuk Anda', 'Cek lagi setelah admin meng-upload dan membagi data baru.'],
    tuntas: ['Semua item selesai dihitung', 'Hasil dikirim otomatis. Item yang selisih masuk ke Validasi.']
  },
  validasi: {
    judul: 'Validasi', fungsi: 'getPendingValidasi', warna: 'ungu', ikon: 'perisai', satuan: 'Qty hasil validasi',
    kunci: (it) => it.id, args: (it, q) => [it.id, Sesi.user.username, q],
    ket: (n) => angka(n) + ' item selisih menunggu dihitung ulang',
    kosong: ['perisai', 'Tidak ada item menunggu validasi', 'Semua selisih sudah divalidasi.'],
    tuntas: ['Semua item sudah divalidasi', 'Item yang tetap selisih masuk ke Verifikasi.']
  }
};

const Tugas = {
  cycle: { items: null, at: 0, sibuk: false, galat: null, area: '', alat: '', cari: '', req: 0 },
  validasi: { items: null, at: 0, sibuk: false, galat: null, area: '', alat: '', cari: '', req: 0 },

  // Dipanggil tiap pengguna atau facility berganti; jawaban permintaan yang masih berjalan diabaikan.
  reset() { ['cycle', 'validasi'].forEach((m) => Object.assign(this[m], { items: null, at: 0, sibuk: false, galat: null, area: '', alat: '', cari: '', req: this[m].req + 1 })); },

  // Urutan kerja: selesaikan satu area dulu, dan di dalam area satu jenis alat dulu, supaya tidak
  // bolak-balik gudang. Urutan lokasi dari server (rute per alat) dipertahankan di dalam kelompok.
  susun(mode, daftar) {
    const urutAlat = {};
    const out = daftar.map((it, i) => {
      const x = Object.assign({}, it);
      x.area = x.area || areaLokasi(x.lokasi);
      x.alat = x.levelGroupKey || alatLokasi(x.lokasi);
      if (urutAlat[x.alat] === undefined) urutAlat[x.alat] = Object.keys(urutAlat).length;
      x._i = i;
      return x;
    });
    out.sort((a, b) => (a.area < b.area ? -1 : a.area > b.area ? 1 : 0) || (urutAlat[a.alat] - urutAlat[b.alat]) || (mode === 'validasi' ? bandingLokasi(a.lokasi, b.lokasi) : 0) || (a._i - b._i));
    return out;
  },
  tampak(mode) {
    const s = this[mode], M = MODE[mode];
    return (s.items || []).filter((it) => !Antrean.sembunyi(mode, M.kunci(it)));
  },
  tersaring(mode) {
    const s = this[mode], q = s.cari.trim().toLowerCase();
    return this.tampak(mode).filter((it) => (!s.area || it.area === s.area) && (!s.alat || it.alat === s.alat) && (!q || cocok(it.lokasi + ' ' + it.article + ' ' + it.description, q)));
  },
  sisa(mode) { return this[mode].items ? this.tampak(mode).length : null; },

  async muat(mode, paksa) {
    const s = this[mode], M = MODE[mode];
    if (s.sibuk || !Sesi.user) return;
    if (!paksa && s.items && Date.now() - s.at < 60000) return;
    const id = ++s.req, siapa = Sesi.user.username;
    s.sibuk = true; this.gambar(mode);
    try {
      const data = await panggil(M.fungsi, siapa);
      if (id !== s.req) return;
      s.items = this.susun(mode, data || []); s.at = Date.now(); s.galat = null;
      Simpanan.set('tugas-' + mode, { items: data || [], at: s.at });
    } catch (e) {
      if (id !== s.req) return;
      s.galat = e;
      if (!s.items) { const c = Simpanan.get('tugas-' + mode); if (c && c.items) { s.items = this.susun(mode, c.items); s.at = c.at; } }
    } finally {
      if (id === s.req) {
        s.sibuk = false;
        this.lencana();
        if (Nav.aktif() === mode) this.gambar(mode);
      }
    }
  },
  lencana() {
    const c = this.sisa('cycle'), v = this.sisa('validasi');
    if (c != null) Nav.setLencana('cycle', c);
    if (v != null) Nav.setLencana('validasi', v);
  },

  /* ---------- layar daftar ---------- */
  gambar(mode) {
    const s = this[mode], M = MODE[mode], node = el('scr-' + mode);
    const semua = this.tampak(mode);
    const segar = h`<button type="button" class="iconbtn${s.sibuk ? ' is-putar' : ''}" data-aksi="tugas-segar" data-mode="${mode}" aria-label="Perbarui daftar">${ikon('segar')}</button>`;
    if (!s.items) {
      isi(node, h`${kepala(M.judul, '', segar)}${s.galat ? gagalMuat(s.galat, 'tugas-segar-' + mode) : kerangka(5, 62)}`);
      return;
    }
    if (!semua.length) {
      const antre = Antrean.jumlah(mode);
      const admin = mode === 'cycle' && Boleh.admin();
      isi(node, h`${kepala(M.judul, '', segar)}<section class="card">${kosong(M.kosong[0], M.kosong[1], antre ? angka(antre) + ' hasil terakhir sedang dikirim.' : admin ? 'Tugas dibagi ke petugas saat upload data. Anda ikut mendapat tugas bila memilih diri sendiri sebagai petugas.' : M.kosong[2],
        admin ? h`<button type="button" class="btn btn--tenang" data-aksi="ke" data-tab="upload">Upload data</button>` : '')}</section>`);
      return;
    }
    const hitung = (kunci) => { const m = {}; semua.forEach((it) => { m[it[kunci]] = (m[it[kunci]] || 0) + 1; }); return m; };
    const perArea = hitung('area'), perAlat = hitung('alat');
    if (s.area && !perArea[s.area]) s.area = '';
    if (s.alat && !perAlat[s.alat]) s.alat = '';
    const daftar = this.tersaring(mode);
    const cip = (jenis, nilai, teks, n) => h`<button type="button" class="chip${s[jenis] === nilai ? ' is-on' : ''}" data-aksi="tugas-saring" data-mode="${mode}" data-jenis="${jenis}" data-nilai="${nilai}">${teks}${n != null ? h` <b>${angka(n)}</b>` : ''}</button>`;
    const areaKeys = Object.keys(perArea).sort(), alatKeys = Object.keys(perAlat);
    const BATAS = 120;
    isi(node, h`
      ${kepala(M.judul, M.ket(semua.length), segar)}
      ${s.galat ? h`<div class="info info--jingga" style="margin-bottom:12px">${ikon('awan_mati')}<div>Daftar belum bisa diperbarui: ${s.galat.message} Menampilkan daftar terakhir.</div></div>` : ''}
      <button type="button" class="btn tg-mulai" data-aksi="tugas-mulai" data-mode="${mode}" ${daftar.length ? '' : 'disabled'}>${ikon(M.ikon)}<span id="tgMulai-${mode}">${this.teksMulai(daftar.length, semua.length)}</span></button>
      ${areaKeys.length > 1 ? h`<div class="chips tg-cip" role="group" aria-label="Area gudang">${cip('area', '', 'Semua area')}${areaKeys.map((a) => cip('area', a, 'Area ' + a, perArea[a]))}</div>` : ''}
      ${alatKeys.length > 1 ? h`<div class="chips tg-cip" role="group" aria-label="Alat bantu">${cip('alat', '', 'Semua alat')}${alatKeys.map((a) => cip('alat', a, (ALAT[a] || {}).nama || a, perAlat[a]))}</div>` : ''}
      ${semua.length > 8 ? h`<div class="cari tg-cari">${ikon('cari', 'ic--kecil')}<input class="field" type="search" id="tgCari-${mode}" placeholder="Cari lokasi atau SKU" value="${s.cari}" autocomplete="off"></div>` : ''}
      <div id="tgDaftar-${mode}">${this.daftarHtml(mode, daftar, BATAS)}</div>
    `);
    const c = el('tgCari-' + mode);
    // Mode fokus memakai daftar yang sedang tersaring, jadi tombol mulai ikut menyebut jumlahnya saat mencari.
    if (c) c.addEventListener('input', tunda(() => {
      s.cari = c.value;
      const d = this.tersaring(mode), tombol = $('.tg-mulai', node), teks = el('tgMulai-' + mode);
      if (!el('tgDaftar-' + mode)) return; // daftar sudah digambar ulang (kosong atau galat) selagi menunggu jeda ketik
      isi(el('tgDaftar-' + mode), this.daftarHtml(mode, d, BATAS));
      if (teks) teks.textContent = this.teksMulai(d.length, this.tampak(mode).length);
      if (tombol) tombol.disabled = !d.length;
    }, 160));
  },
  teksMulai(terpilih, semua) { return !terpilih ? 'Tidak ada item yang cocok' : terpilih === semua ? 'Mulai dari lokasi pertama' : 'Mulai ' + angka(terpilih) + ' item terpilih'; },
  // Daftar mengikuti urutan kerja (area → alat bantu). Judul kelompok membuat urutan itu terbaca.
  daftarHtml(mode, daftar, batas) {
    if (!daftar.length) return h`<p class="redup" style="text-align:center;padding:24px 0">Tidak ada item yang cocok.</p>`;
    const M = MODE[mode], tampil = daftar.slice(0, batas);
    const kel = (it) => it.area + '|' + it.alat, jumlah = {};
    daftar.forEach((it) => { jumlah[kel(it)] = (jumlah[kel(it)] || 0) + 1; });
    const berkelompok = Object.keys(jumlah).length > 1;
    let lalu = null;
    return h`<div class="list">${tampil.map((it) => {
      const k = kel(it), judul = berkelompok && k !== lalu ? h`<div class="list__grup"><span>${it.area ? 'Area ' + it.area + ', ' : ''}${((ALAT[it.alat] || {}).nama || it.alat).toLowerCase()}</span><b>${angka(jumlah[k])}</b></div>` : '';
      lalu = k;
      return h`${judul}<button type="button" class="row" data-aksi="tugas-buka" data-mode="${mode}" data-kunci="${M.kunci(it)}">
      ${labelRak(it.lokasi)}<span class="row__isi"><span class="row__t">${it.article}</span><span class="row__s">${it.description}</span></span>
      <span class="row__ekor" title="${(ALAT[it.alat] || {}).nama || ''}">${ikon((ALAT[it.alat] || {}).ikon || 'jalan', 'ic--kecil')}</span></button>`;
    })}</div>
      ${daftar.length > batas ? h`<p class="redup" style="text-align:center;margin-top:12px">Dan ${angka(daftar.length - batas)} item lagi. Gunakan pencarian atau mulai menghitung.</p>` : ''}`;
  }
};

/* ---------- mode fokus ---------- */
const Fokus = {
  mode: 'cycle', daftar: [], pos: 0, awal: 0, selesai: 0, qty: '', mulaiAt: 0, terakhir: null, log: [], areaLalu: '', alatLalu: '',

  buka(mode, kunciAwal) {
    const M = MODE[mode];
    this.mode = mode; this.daftar = Tugas.tersaring(mode).slice();
    this.pos = Math.max(0, this.daftar.findIndex((it) => String(M.kunci(it)) === String(kunciAwal)));
    this.awal = this.daftar.length; this.selesai = 0; this.qty = ''; this.mulaiAt = Date.now(); this.terakhir = null; this.log = [];
    this.areaLalu = ''; this.alatLalu = ''; this.lokasiLalu = '';
    if (!this.daftar.length) { Toast.tampil('Tidak ada item untuk dihitung.'); return; }
    Nav.dorong('hitung');
  },
  item() { return this.daftar[this.pos] || null; },

  gambar() {
    const M = MODE[this.mode], it = this.item(), node = el('scr-hitung');
    node.className = 'scr is-on fk fk--' + M.warna;
    if (!it) { this.gambarTuntas(); return; }
    const baru = (this.areaLalu && this.areaLalu !== it.area) || (this.alatLalu && this.alatLalu !== it.alat);
    // Beberapa SKU bisa berada di satu lokasi: beri tahu supaya tidak pindah rak dan tidak salah SKU.
    const lokasiSama = !!this.lokasiLalu && this.lokasiLalu === it.lokasi;
    this.areaLalu = it.area; this.alatLalu = it.alat;
    const alat = ALAT[it.alat] || ALAT.tanpa_alat;
    isi(node, h`
      <header class="fk__kepala">
        <button type="button" class="iconbtn" data-aksi="kembali" aria-label="Kembali ke daftar">${ikon('kiri')}</button>
        <div class="fk__judul"><b>${M.judul}</b><span>Item ${angka(Math.min(this.selesai + 1, this.awal))} dari ${angka(this.awal)}</span></div>
        <button type="button" class="fk__antre" id="fkAntre" data-aksi="fokus-log" aria-label="Riwayat sesi ini"></button>
      </header>
      <div class="bar bar--tipis bar--${M.warna}"><i style="width:${this.awal ? (this.selesai / this.awal) * 100 : 0}%"></i></div>
      <div class="fk__kartu" id="fkKartu">
        ${labelRakBesar(it.lokasi)}
        <div class="fk__barang">
          <div class="fk__teks"><div class="fk__sku">${it.article}</div><div class="fk__nama">${it.description || ''}</div></div>
          ${this.daftar.length > 1 ? h`<button type="button" class="fk__lewati" data-aksi="fokus-lewati" aria-label="Lewati item ini dulu">${ikon('lewati', 'ic--kecil')}Lewati</button>` : ''}
        </div>
        <div class="fk__cip${baru ? ' is-baru' : ''}${lokasiSama ? ' fk__cip--sama' : ''}">${lokasiSama ? h`<span class="pill pill--kuning">${ikon('pin', 'ic--kecil')}Masih di lokasi ini, SKU berbeda</span>` : h`<span class="pill">${ikon('pin', 'ic--kecil')}Area ${it.area || '-'}</span><span class="pill">${ikon(alat.ikon, 'ic--kecil')}${alat.nama}</span>`}</div>
      </div>
      <div class="fk__terakhir" id="fkTerakhir" aria-live="polite"></div>
      <div class="fk__qty" aria-live="polite"><output class="fk__angka" id="fkAngka"></output><span class="fk__satuan">${M.satuan}</span></div>
      <div class="pad" id="fkPad">
        ${[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => h`<button type="button" data-aksi="pad" data-k="${n}">${n}</button>`)}
        <button type="button" data-aksi="pad" data-k="hapus" aria-label="Hapus angka terakhir">${ikon('hapus_kiri')}</button>
        <button type="button" data-aksi="pad" data-k="0">0</button>
        <button type="button" class="pad__simpan" id="fkSimpan" data-aksi="fokus-simpan">Simpan</button>
      </div>
    `);
    this.gambarQty(); this.gambarAntre(); this.gambarTerakhir();
  },
  // Hasil terakhir tampil di ruang kosong di atas angka (bukan toast yang menutupi layar), lengkap
  // dengan tombol Urungkan selama hasil itu belum dikirim.
  gambarTerakhir() {
    const n = el('fkTerakhir'); if (!n) return;
    const t = this.terakhir, sisaMs = t ? t.entri.kirimAt - Date.now() : 0;
    clearTimeout(this._tTerakhir);
    if (!t || sisaMs <= 0 || t.entri.status !== 'tahan') { isi(n, ''); return; }
    isi(n, h`<div class="simpanan"><span class="simpanan__ic">${ikon('cek', 'ic--kecil')}</span><span class="simpanan__teks"><b>${t.it.lokasi}</b> tersimpan, qty ${angka(t.entri.info.qty)}</span>
      <button type="button" class="simpanan__urung" data-aksi="fokus-urung">Urungkan</button><i class="simpanan__waktu" style="animation-duration:${sisaMs}ms"></i></div>`);
    this._tTerakhir = setTimeout(() => this.gambarTerakhir(), sisaMs + 30);
  },
  gambarQty() {
    const a = el('fkAngka'); if (!a) return;
    a.textContent = this.qty === '' ? '0' : angka(Number(this.qty));
    a.classList.toggle('is-kosong', this.qty === '');
    el('fkSimpan').disabled = this.qty === '';
  },
  gambarAntre() {
    const b = el('fkAntre'); if (!b) return;
    const n = Antrean.jumlah();
    b.classList.toggle('is-tunggu', n > 1 || (n > 0 && !navigator.onLine));
    isi(b, h`${ikon(n > 0 ? 'awan_kirim' : 'daftar', 'ic--kecil')}<span>${n > 0 ? angka(n) : angka(this.log.length)}</span>`);
  },
  tekan(k) {
    if (!this.item()) return;
    if (k === 'hapus') this.qty = this.qty.slice(0, -1);
    else if (this.qty.length < 5) this.qty = (this.qty === '0' ? '' : this.qty) + k;
    this.gambarQty();
  },
  simpan() {
    const it = this.item(), M = MODE[this.mode];
    if (!it || this.qty === '') return;
    const q = Number(this.qty);
    const entri = Antrean.tambah({ jenis: this.mode, kunci: M.kunci(it), args: M.args(it, q), info: { lokasi: it.lokasi, article: it.article, description: it.description, qty: q, sys: it.qtySystem } }, TAHAN_MS);
    this.terakhir = { it, entri, pos: this.pos }; this.lokasiLalu = it.lokasi;
    this.log.unshift({ id: entri.id, lokasi: it.lokasi, article: it.article, qty: q, sys: it.qtySystem, status: 'antre' });
    this.daftar.splice(this.pos, 1);
    if (this.pos >= this.daftar.length) this.pos = 0;
    this.selesai++; this.qty = '';
    Tugas.lencana();
    this.gambar();
    const k = el('fkKartu'); if (k) k.classList.add('is-masuk');
  },
  urung() {
    const t = this.terakhir; if (!t) return;
    const e = Antrean.urung(t.entri.id);
    if (!e) { this.terakhir = null; this.gambarTerakhir(); Toast.tampil('Hasil ini sudah terkirim, tidak bisa diurungkan.'); return; }
    this.terakhir = null;
    this.log = this.log.filter((l) => l.id !== e.id);
    this.daftar.splice(Math.min(t.pos, this.daftar.length), 0, t.it);
    this.pos = Math.min(t.pos, this.daftar.length - 1);
    this.selesai = Math.max(0, this.selesai - 1); this.qty = String(e.info.qty); this.lokasiLalu = '';
    Tugas.lencana();
    this.gambar();
  },
  lewati() {
    if (this.daftar.length < 2) return;
    this.pos = (this.pos + 1) % this.daftar.length; this.qty = ''; this.lokasiLalu = '';
    this.gambar();
    const k = el('fkKartu'); if (k) k.classList.add('is-masuk');
  },
  gambarTuntas() {
    const M = MODE[this.mode], sisaLain = Tugas.tampak(this.mode).length;
    const menit = Math.max(1, Math.round((Date.now() - this.mulaiAt) / 60000));
    isi(el('scr-hitung'), h`
      <header class="fk__kepala"><button type="button" class="iconbtn" data-aksi="kembali" aria-label="Kembali ke daftar">${ikon('kiri')}</button>
        <div class="fk__judul"><b>${M.judul}</b><span>${angka(this.selesai)} item selesai</span></div>
        <button type="button" class="fk__antre" id="fkAntre" data-aksi="fokus-log" aria-label="Riwayat sesi ini"></button></header>
      <div class="bar bar--tipis bar--${M.warna}"><i style="width:100%"></i></div>
      <div class="fk__terakhir fk__terakhir--atas" id="fkTerakhir" aria-live="polite"></div>
      <div class="fk__tuntas">
        <div class="fk__tuntas-ic">${ikon('cek', 'ic--besar')}</div>
        <h2>${sisaLain ? 'Bagian ini selesai' : M.tuntas[0]}</h2>
        <p>${angka(this.selesai)} item dalam ${angka(menit)} menit. ${sisaLain ? h`Masih ada ${angka(sisaLain)} item di bagian lain.` : M.tuntas[1]}</p>
        ${sisaLain ? h`<button type="button" class="btn" data-aksi="fokus-lanjut-semua">Lanjutkan ${angka(sisaLain)} item lainnya</button>` : ''}
        <button type="button" class="btn ${sisaLain ? 'btn--garis' : ''}" data-aksi="ke" data-tab="home">Kembali ke Home</button>
      </div>`);
    this.gambarAntre(); this.gambarTerakhir(); // item terakhir pun masih bisa diurungkan
  },
  bukaLog() {
    const baris = this.log.map((l) => {
      let status;
      if (l.status === 'terkirim') {
        const selisih = l.qty - Number(l.sys);
        status = isNaN(selisih) ? h`<span class="pill pill--hijau">Terkirim</span>`
          : selisih === 0 ? h`<span class="pill pill--hijau">Cocok</span>`
            : h`<span class="pill pill--jingga">Selisih ${bertanda(selisih)}</span>`;
      } else if (l.status === 'ditolak') status = h`<span class="pill pill--merah">Ditolak</span>`;
      else status = h`<span class="pill">Menunggu kirim</span>`;
      return h`<div class="row">${labelRak(l.lokasi)}<div class="row__isi"><div class="row__t">${l.article}</div><div class="row__s">Qty ${angka(l.qty)}${l.pesan ? h`, ${l.pesan}` : ''}</div></div><div class="row__ekor">${status}</div></div>`;
    });
    Lembar.buka({
      judul: 'Riwayat sesi ini',
      isi: this.log.length
        ? h`<p class="redup" style="margin-bottom:12px">${this.mode === 'cycle' ? 'Item yang selisih otomatis masuk antrean Validasi.' : 'Item yang tetap selisih masuk ke Verifikasi.'} Hasil muncul setelah terkirim.</p><div class="list">${baris}</div>`
        : kosong('daftar', 'Belum ada hasil', 'Hasil ' + (this.mode === 'cycle' ? 'hitung' : 'validasi') + ' pada sesi ini muncul di sini.')
    });
  },
  // Hasil yang ditolak server kembali menjadi tugas di sesi ini (di akhir urutan).
  kembalikan(entri) {
    if (!this.log.some((l) => l.id === entri.id)) return; // hasil dari sesi sebelumnya: cukup muncul lagi di daftar tugas
    const M = MODE[this.mode], it = (Tugas[this.mode].items || []).find((x) => String(M.kunci(x)) === String(entri.kunci));
    if (!it || this.daftar.indexOf(it) >= 0) return;
    this.daftar.push(it); this.selesai = Math.max(0, this.selesai - 1);
    this.gambar();
  },
  // kabar dari antrean kirim
  kabar(ev) {
    const l = this.log.find((x) => x.id === ev.entri.id);
    if (l) {
      if (ev.tipe === 'terkirim') l.status = 'terkirim';
      else if (ev.tipe === 'ditolak') { l.status = 'ditolak'; l.pesan = ev.pesan; }
    }
    if (ev.tipe === 'ditolak') {
      Toast.tampil(h`${(ev.entri.info || {}).lokasi || 'Item'} belum tersimpan: ${ev.pesan}`, { galat: true, atas: Nav.aktif() === 'hitung', lama: 6000 });
      if (Nav.aktif() === 'hitung' && ev.entri.jenis === Fokus.mode) Fokus.kembalikan(ev.entri);
      Tugas[ev.entri.jenis].at = 0; // daftar perlu diambil ulang: item itu kembali jadi tugas
    }
    if (Nav.aktif() === 'hitung') this.gambarAntre();
    Tugas.lencana();
  }
};

/* ---------- layar & aksi ---------- */
['cycle', 'validasi'].forEach((mode) => {
  Layar[mode] = {
    masuk(p) {
      Tugas.gambar(mode);
      const muat = Tugas.muat(mode, false);
      if (p && p.mulai) {
        if (Tugas[mode].items) Fokus.buka(mode);
        else Promise.resolve(muat).then(() => { if (Nav.aktif() === mode && Tugas.tampak(mode).length) Fokus.buka(mode); });
      }
    }
  };
  Aksi['tugas-segar-' + mode] = () => Tugas.muat(mode, true);
});
Layar.hitung = {
  masuk() { document.body.classList.add('mode-fokus'); ukurLayar(); Fokus.gambar(); document.addEventListener('keydown', Fokus._tombol); },
  keluar() { document.body.classList.remove('mode-fokus'); document.removeEventListener('keydown', Fokus._tombol); clearTimeout(Fokus._tTerakhir); Antrean.lepas(); Tugas[Fokus.mode].cari = ''; }
};
// Papan ketik fisik atau pemindai: angka, Backspace, Enter
Fokus._tombol = (ev) => {
  if (Lembar.tumpukan.length || ev.ctrlKey || ev.metaKey || ev.altKey) return;
  if (/^[0-9]$/.test(ev.key)) { Fokus.tekan(ev.key); ev.preventDefault(); }
  else if (ev.key === 'Backspace') { Fokus.tekan('hapus'); ev.preventDefault(); }
  else if (ev.key === 'Enter') { Fokus.simpan(); ev.preventDefault(); }
};
Aksi['tugas-segar'] = (t) => Tugas.muat(t.dataset.mode, true);
Aksi['tugas-saring'] = (t) => { const s = Tugas[t.dataset.mode]; s[t.dataset.jenis] = t.dataset.nilai; Tugas.gambar(t.dataset.mode); };
Aksi['tugas-mulai'] = (t) => Fokus.buka(t.dataset.mode);
Aksi['tugas-buka'] = (t) => Fokus.buka(t.dataset.mode, t.dataset.kunci);
Aksi['pad'] = (t) => Fokus.tekan(t.dataset.k);
Aksi['fokus-simpan'] = () => Fokus.simpan();
Aksi['fokus-lewati'] = () => Fokus.lewati();
Aksi['fokus-urung'] = () => Fokus.urung();
Aksi['fokus-log'] = () => Fokus.bukaLog();
Aksi['fokus-lanjut-semua'] = () => { const s = Tugas[Fokus.mode]; s.area = ''; s.alat = ''; s.cari = ''; Nav.kembali(); Fokus.buka(Fokus.mode); };
Antrean.dengar((ev) => Fokus.kabar(ev));
