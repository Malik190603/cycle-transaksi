/* Antarmuka bersama: aksi (data-aksi), toast, lembar bawah, konfirmasi, tabir sibuk, navigasi. */

/* ---------- aksi: satu pendengar klik untuk seluruh aplikasi ---------- */
const Aksi = {};
document.addEventListener('click', (ev) => {
  const t = ev.target.closest('[data-aksi]');
  if (!t || t.disabled) return;
  const fn = Aksi[t.dataset.aksi];
  if (!fn) { console.error('Aksi tidak dikenal: ' + t.dataset.aksi); return; }
  ev.preventDefault();
  fn(t, ev);
});

/* ---------- toast ---------- */
const Toast = {
  _t: 0, _el: null,
  tampil(teks, opsi) {
    const o = opsi || {};
    if (!this._el) { this._el = document.createElement('div'); document.body.appendChild(this._el); }
    const n = this._el;
    n.className = 'toast' + (o.atas ? ' toast--atas' : '') + (o.galat ? ' toast--galat' : '');
    n.setAttribute('role', o.galat ? 'alert' : 'status');
    isi(n, h`<span class="toast__teks">${teks}</span>${o.aksi ? h`<button type="button" class="toast__aksi">${o.aksi}</button>` : ''}`);
    const b = $('.toast__aksi', n);
    if (b) b.addEventListener('click', () => { this.sembunyi(); if (o.onAksi) o.onAksi(); });
    void n.offsetWidth; n.classList.add('is-on');
    clearTimeout(this._t);
    this._t = setTimeout(() => this.sembunyi(), o.lama || (o.aksi ? 5200 : 3200));
  },
  sembunyi() { clearTimeout(this._t); if (this._el) this._el.classList.remove('is-on'); },
  galat(e) { this.tampil((e && e.message) || String(e), { galat: true, lama: 4500 }); }
};

/* ---------- lembar bawah ---------- */
const Lembar = {
  tumpukan: [],
  buka(opsi) {
    const o = opsi || {};
    const wadah = document.createElement('div');
    wadah.className = 'lembar-wadah';
    isi(wadah, h`<div class="lembar-tabir"></div>
      <div class="lembar${o.penuh ? ' lembar--penuh' : ''}" role="dialog" aria-modal="true" aria-label="${o.judul || ''}">
        <div class="lembar__pegangan"></div>
        <div class="lembar__kepala"><h2 class="lembar__judul">${o.judul || ''}</h2><button type="button" class="iconbtn iconbtn--polos" data-tutup aria-label="Tutup">${ikon('tutup')}</button></div>
        <div class="lembar__isi"></div>
        ${o.kaki ? h`<div class="lembar__kaki">${o.kaki}</div>` : ''}
      </div>`);
    document.body.appendChild(wadah);
    const ctl = {
      el: wadah, isi: $('.lembar__isi', wadah), kaki: $('.lembar__kaki', wadah), onTutup: o.onTutup, terkunci: false,
      judul(t) { $('.lembar__judul', wadah).textContent = t; },
      tutup: () => this.tutup(ctl)
    };
    if (o.isi) isi(ctl.isi, o.isi);
    $('.lembar-tabir', wadah).addEventListener('click', () => { if (!ctl.terkunci) ctl.tutup(); });
    $('[data-tutup]', wadah).addEventListener('click', () => { if (!ctl.terkunci) ctl.tutup(); });
    this.tumpukan.push(ctl);
    void wadah.offsetWidth; wadah.classList.add('is-on');
    return ctl;
  },
  tutup(ctl) {
    const c = ctl || this.tumpukan[this.tumpukan.length - 1];
    if (!c) return false;
    const i = this.tumpukan.indexOf(c);
    if (i >= 0) this.tumpukan.splice(i, 1);
    c.el.classList.remove('is-on');
    setTimeout(() => c.el.remove(), 260);
    if (c.onTutup) { const f = c.onTutup; c.onTutup = null; f(); }
    return true;
  },
  tutupSemua() { while (this.tumpukan.length) this.tutup(); }
};

/* ---------- konfirmasi (pengganti confirm() bawaan) ---------- */
function tanya(opsi) {
  const o = opsi || {};
  return new Promise((jawab) => {
    let hasil = false;
    const l = Lembar.buka({
      judul: o.judul || 'Lanjutkan?',
      isi: h`${o.pesan ? h`<p class="redup" style="margin-bottom:18px">${o.pesan}</p>` : ''}
        <div class="baris-tombol"><button type="button" class="btn btn--garis" data-j="0">${o.batal || 'Batal'}</button>
        <button type="button" class="btn${o.bahaya ? ' btn--bahaya' : ''}" data-j="1">${o.ya || 'Ya, lanjutkan'}</button></div>`,
      onTutup: () => jawab(hasil)
    });
    $$('[data-j]', l.isi).forEach((b) => b.addEventListener('click', () => { hasil = b.dataset.j === '1'; l.tutup(); }));
  });
}

/* ---------- tabir sibuk dengan cincin persen (impor data, buat facility) ---------- */
const Sibuk = {
  _el: null, _poll: 0, K: 2 * Math.PI * 34,
  tampil(teks, ket) {
    this.sembunyi();
    const n = document.createElement('div'); n.className = 'sibuk'; n.setAttribute('role', 'alertdialog'); n.setAttribute('aria-busy', 'true');
    isi(n, h`<div class="sibuk__kotak"><div class="sibuk__cincin"><svg viewBox="0 0 84 84"><circle class="jalur" cx="42" cy="42" r="34"/><circle class="isi" cx="42" cy="42" r="34" stroke-dasharray="${this.K}" stroke-dashoffset="${this.K}"/></svg><div class="sibuk__persen">0%</div></div>
      <div class="sibuk__teks">${teks}</div><div class="sibuk__ket">${ket || 'Jangan tutup aplikasi.'}</div></div>`);
    document.body.appendChild(n); this._el = n;
  },
  persen(p, ket) {
    if (!this._el) return;
    const v = Math.max(0, Math.min(100, Number(p) || 0));
    $('.isi', this._el).style.strokeDashoffset = this.K * (1 - v / 100);
    $('.sibuk__persen', this._el).textContent = Math.round(v) + '%';
    if (ket) $('.sibuk__ket', this._el).textContent = ket;
  },
  // Backend melaporkan progres lewat CacheService; diambil tiap 1,5 detik (tiap panggilan = satu eksekusi Apps Script).
  pantau(jobId) {
    clearInterval(this._poll);
    this._poll = setInterval(() => { panggil('getImportProgress', jobId).then((r) => { if (r && r.percent) this.persen(r.percent, r.label); }).catch(() => {}); }, 1500);
  },
  sembunyi() { clearInterval(this._poll); if (this._el) { this._el.remove(); this._el = null; } }
};

/* ---------- pita atas (offline, demo) ---------- */
const Pita = {
  _isi: {},
  pasang(kunci, html, kelas) { this._isi[kunci] = { html, kelas }; this._gambar(); },
  lepas(kunci) { delete this._isi[kunci]; this._gambar(); },
  _gambar() {
    let n = el('pita');
    const kunci = Object.keys(this._isi);
    if (!kunci.length) { if (n) n.remove(); document.body.classList.remove('ada-pita'); ukurLayar(); return; }
    if (!n) { n = document.createElement('div'); n.id = 'pita'; document.body.appendChild(n); }
    const k = this._isi.offline ? 'offline' : kunci[0];
    n.className = 'pita' + (this._isi[k].kelas ? ' ' + this._isi[k].kelas : '');
    isi(n, this._isi[k].html);
    document.body.classList.add('ada-pita');
    document.documentElement.style.setProperty('--pita-t', (n.offsetHeight - (parseFloat(getComputedStyle(n).paddingTop) || 0) + 5) + 'px');
    ukurLayar();
  }
};

/*
 * Kerapatan mode fokus mengikuti tinggi layar yang benar-benar bisa dipakai: tinggi tampilan dikurangi
 * bilah sistem Android dan pita atas. Media query (max-height) tidak cukup, karena tampilan aplikasi
 * membentang di bawah bilah sistem sehingga tinggi viewport lebih besar dari ruang yang tersedia.
 * Hasilnya ditulis ke <html data-rapat="1 2 ..."> dan dipakai 30-layar.css.
 */
const BATAS_RAPAT = [728, 668, 578, 558]; // tinggi alami tingkat sebelumnya (diukur): di bawah angka ini tingkat 1, 2, 3, 4 berlaku
function ukurLayar() {
  let p = el('ukurAman');
  if (!p) {
    p = document.createElement('div'); p.id = 'ukurAman';
    p.style.cssText = 'position:fixed;visibility:hidden;pointer-events:none;top:0;left:0;width:0;height:0;padding-top:var(--sat);padding-bottom:var(--sab)';
    document.body.appendChild(p);
  }
  const cs = getComputedStyle(p), atas = parseFloat(cs.paddingTop) || 0, bawah = parseFloat(cs.paddingBottom) || 0;
  const pita = el('pita') ? Math.max(0, el('pita').offsetHeight - atas) : 0;
  const tinggi = window.innerHeight - atas - bawah - pita;
  const tingkat = BATAS_RAPAT.map((b, i) => (tinggi < b ? String(i + 1) : '')).filter(Boolean).join(' ');
  if (document.documentElement.getAttribute('data-rapat') !== tingkat) document.documentElement.setAttribute('data-rapat', tingkat);
}
window.addEventListener('resize', ukurLayar);
function periksaJaringan() {
  if (navigator.onLine) Pita.lepas('offline');
  else Pita.pasang('offline', h`Tidak ada internet. Hasil tetap tersimpan di HP.`, 'pita--jingga');
}
window.addEventListener('online', periksaJaringan);
window.addEventListener('offline', periksaJaringan);

/* ---------- navigasi: tab bawah + halaman tumpuk ---------- */
const Layar = {}; // id → { tab, masuk(params), keluar(), segar() }
const TAB = [
  { id: 'home', nama: 'Home', ikon: 'home', boleh: () => true },
  { id: 'cycle', nama: 'Cycle', ikon: 'kotak', boleh: () => true, lencana: 'jingga' },
  { id: 'validasi', nama: 'Validasi', ikon: 'perisai', boleh: () => Boleh.validasi(), lencana: 'ungu' },
  { id: 'verifikasi', nama: 'Verifikasi', ikon: 'selisih', boleh: () => Boleh.validasi(), lencana: 'merah' },
  { id: 'upload', nama: 'Upload', ikon: 'unggah', boleh: () => Boleh.admin() },
  { id: 'report', nama: 'Report', ikon: 'grafik', boleh: () => Boleh.admin() }
];
const Nav = {
  tab: '', halaman: [], lencana: {},
  bangun() {
    const n = el('nav');
    isi(n, TAB.filter((t) => t.boleh()).map((t) => h`<button type="button" class="nav__item" data-aksi="tab" data-tab="${t.id}" id="tab-${t.id}">${ikon(t.ikon)}<span>${t.nama}</span></button>`));
    this._lencana();
  },
  setLencana(id, n) { this.lencana[id] = Number(n) || 0; this._lencana(); },
  _lencana() {
    TAB.forEach((t) => {
      const b = el('tab-' + t.id); if (!b) return;
      const lama = $('.nav__lencana', b); if (lama) lama.remove();
      const v = this.lencana[t.id];
      if (t.lencana && v > 0) b.insertAdjacentHTML('beforeend', '<span class="nav__lencana nav__lencana--' + t.lencana + '">' + (v > 99 ? '99+' : v) + '</span>');
      b.setAttribute('aria-label', t.nama + (v > 0 ? ', ' + v : ''));
    });
  },
  _tampil(id) {
    $$('.scr').forEach((s) => s.classList.toggle('is-on', s.id === 'scr-' + id));
    window.scrollTo(0, 0);
  },
  aktif() { return this.halaman.length ? this.halaman[this.halaman.length - 1].id : this.tab; },
  ke(id, params) {
    const t = TAB.find((x) => x.id === id);
    if (!t || !t.boleh()) id = 'home';
    const lama = this.aktif();
    if (lama && Layar[lama] && Layar[lama].keluar) Layar[lama].keluar();
    this.halaman = []; this.tab = id;
    document.body.classList.remove('tanpa-nav');
    $$('.nav__item').forEach((b) => { const on = b.dataset.tab === id; b.classList.toggle('is-on', on); if (on) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current'); });
    this._tampil(id);
    if (Layar[id] && Layar[id].masuk) Layar[id].masuk(params || {});
  },
  dorong(id, params) {
    const lama = this.aktif();
    if (lama && Layar[lama] && Layar[lama].keluar) Layar[lama].keluar();
    this.halaman.push({ id });
    document.body.classList.add('tanpa-nav');
    this._tampil(id);
    if (Layar[id] && Layar[id].masuk) Layar[id].masuk(params || {});
  },
  // true bila ada yang ditutup (lembar, halaman, atau pindah ke Home)
  kembali() {
    if (Lembar.tumpukan.length) { const l = Lembar.tumpukan[Lembar.tumpukan.length - 1]; if (!l.terkunci) Lembar.tutup(); return true; }
    if (this.halaman.length) {
      const h0 = this.halaman.pop();
      if (Layar[h0.id] && Layar[h0.id].keluar) Layar[h0.id].keluar();
      const id = this.aktif();
      if (!this.halaman.length) document.body.classList.remove('tanpa-nav');
      this._tampil(id);
      if (Layar[id] && Layar[id].masuk) Layar[id].masuk({ kembali: true });
      return true;
    }
    if (this.tab && this.tab !== 'home' && this.tab !== 'login') { this.ke('home'); return true; }
    return false;
  }
};
Aksi['tab'] = (t) => Nav.ke(t.dataset.tab);
Aksi['kembali'] = () => Nav.kembali();
Aksi['ke'] = (t) => { Lembar.tutupSemua(); Nav.ke(t.dataset.tab); };

/* ---------- potongan tampilan yang sering dipakai ---------- */
const kosong = (ik, judul, ket, tombol) => h`<div class="kosong"><div class="kosong__ic">${ikon(ik, 'ic--besar')}</div><div class="kosong__t">${judul}</div>${ket ? h`<div class="kosong__s">${ket}</div>` : ''}${tombol || ''}</div>`;
const memuat = (teks) => h`<div class="muat"><span class="putar"></span>${teks || 'Memuat…'}</div>`;
const kerangka = (baris, tinggi) => h`<div class="tumpuk">${Array.from({ length: baris || 3 }, () => h`<div class="kerangka" style="height:${tinggi || 64}px;border-radius:var(--r-kartu)"></div>`)}</div>`;
const gagalMuat = (e, aksi) => kosong(e && e.jenis === 'offline' ? 'awan_mati' : 'awas', e && e.jenis === 'offline' ? 'Tidak ada koneksi' : 'Data tidak bisa dimuat', (e && e.message) || '', h`<button type="button" class="btn btn--tenang" data-aksi="${aksi}">Coba lagi</button>`);
const kepala = (judul, ket, kanan) => h`<header class="hd"><div class="hd__teks"><h1 class="hd__judul">${judul}</h1>${ket ? h`<div class="hd__ket">${ket}</div>` : ''}</div>${kanan || ''}</header>`;
const kepalaHalaman = (judul, ket, kanan) => h`<header class="hd"><button type="button" class="iconbtn" data-aksi="kembali" aria-label="Kembali">${ikon('kiri')}</button><div class="hd__teks"><h1 class="hd__judul" style="font-size:21px">${judul}</h1>${ket ? h`<div class="hd__ket">${ket}</div>` : ''}</div>${kanan || ''}</header>`;

/* ---------- SheetJS dimuat saat dibutuhkan saja (unggah & ekspor): 900 KB tidak membebani pembukaan aplikasi ---------- */
let _xlsxJanji = null;
function muatXlsx() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  if (!_xlsxJanji) {
    _xlsxJanji = new Promise((ok, gagal) => {
      const s = document.createElement('script'); s.src = CFG.xlsxUrl || 'vendor/xlsx.full.min.js'; // xlsxUrl: hanya untuk pratinjau di browser
      s.onload = () => ok(window.XLSX); s.onerror = () => { _xlsxJanji = null; gagal(new Error('Pembaca Excel tidak bisa dimuat.')); };
      document.head.appendChild(s);
    });
  }
  return _xlsxJanji;
}
// Di WebView Android unduhan lewat tautan blob tidak menghasilkan file: tulis ke folder cache lalu buka
// dengan aplikasi spreadsheet di HP.
async function eksporExcel(namaFile, lembar) {
  // Pratinjau di browser tidak bisa menyimpan file ke HP; jangan berpura-pura berhasil.
  if (CFG.demoSaja) { Toast.tampil('Ekspor ke Excel tersedia di aplikasi Android.'); return; }
  try {
    const X = await muatXlsx();
    const wb = X.utils.book_new();
    lembar.forEach((l) => X.utils.book_append_sheet(wb, X.utils.json_to_sheet(l.baris.length ? l.baris : [{ Info: 'Tidak ada data' }]), String(l.nama).slice(0, 31)));
    const file = String(namaFile).replace(/[\\/:*?"<>|]+/g, '_');
    if (NATIVE && PL.Filesystem && PL.FileOpener) {
      const data = X.write(wb, { bookType: 'xlsx', type: 'base64' });
      await PL.Filesystem.writeFile({ path: file, data, directory: 'CACHE' });
      const uri = (await PL.Filesystem.getUri({ path: file, directory: 'CACHE' })).uri;
      await PL.FileOpener.open({ filePath: uri, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', openWithDefault: false });
    } else {
      X.writeFile(wb, file);
    }
    Toast.tampil('File ' + file + ' dibuat.');
  } catch (e) { Toast.galat(new Error('Ekspor gagal: ' + ((e && e.message) || e))); }
}
