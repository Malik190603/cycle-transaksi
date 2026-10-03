/*
 * Verifikasi: tindak lanjut selisih yang sudah dikonfirmasi validasi.
 *   Task      → daftar selisih aktif; ketuk untuk memperbarui status atau menutup dengan bukti WMS
 *   Pasangan  → minus di satu lokasi + plus di lokasi lain (SKU & qty sama) ditutup bersama satu bukti MOVE.
 *               Pasangan yang cocok dicari otomatis, bukan dipilih manual dari dua daftar.
 *   Riwayat   → task yang sudah ditutup
 */
const KATEGORI_TANPA_BUKTI = ['Salah Hitung', 'Barang Sudah di Picking'];
const KATEGORI_ADJUSTMENT = ['Adjustment Plus', 'Adjustment Minus'];
const KOLOM_BUKTI = ['No', 'STORERKEY', 'TRANTYPE', 'SKU', 'Description', 'SKUGROUP', 'LOT', 'FROMLOC', 'FROMID', 'TOLOC', 'TOID', 'SOURCEKEY', 'QTY', 'ADDDATE', 'ADDWHO'];
const TAHAP = ['Open', 'Sedang Dicari', 'Menunggu Konfirmasi', 'Selesai'];
const namaTahap = (s) => (s === 'Open' ? 'Verifikasi' : s === 'Sedang Dicari' ? 'Sedang dicari' : s === 'Menunggu Konfirmasi' ? 'Menunggu konfirmasi' : s);

// Bukti ditempel dari WMS/Excel: satu baris per transaksi, kolom dipisah tab, tanpa baris judul.
function bacaBukti(teks) {
  const baris = String(teks || '').split(/\r?\n/).map((l) => l.replace(/\s+$/, '')).filter((l) => l.trim());
  if (!baris.length) throw new Error('Bukti masih kosong.');
  const rows = baris.map((l) => { const sel = l.split('\t'); return KOLOM_BUKTI.map((_, i) => (sel[i] !== undefined ? sel[i].trim() : '')); });
  const cacat = rows.findIndex((r) => !r[2] || !r[3] || !r[12]) + 1;
  if (cacat) throw new Error('Baris ke-' + cacat + ' tidak lengkap. TRANTYPE, SKU, dan QTY wajib terisi; urutan kolom: ' + KOLOM_BUKTI.join(', ') + '.');
  return rows;
}
// Cermin validateBuktiTransaksi_ di backend, supaya kesalahan terlihat sebelum dikirim. Server tetap memeriksa ulang.
function periksaBukti(rows, article, selisihAbs, kategori, lokasi) {
  const sku = String(article || '').trim().toUpperCase(), up = (v) => String(v || '').trim().toUpperCase();
  if (KATEGORI_ADJUSTMENT.indexOf(kategori) >= 0) {
    const rel = rows.filter((r) => up(r[3]) === sku && up(r[2]) === 'ADJUSTMENT');
    if (!rel.length) return 'Belum ada baris ADJUSTMENT untuk SKU ' + article + '.';
    const total = rel.reduce((s, r) => s + (Number(r[12]) || 0), 0), perlu = kategori === 'Adjustment Plus' ? selisihAbs : -selisihAbs;
    if (Math.round(total) !== Math.round(perlu)) return 'Total qty adjustment ' + total + ', seharusnya ' + perlu + ' (tanda plus/minus harus sesuai).';
    if (!rel.some((r) => up(r[9]) === up(lokasi))) return 'TOLOC harus lokasi task ini (' + lokasi + ').';
    return '';
  }
  const rel = rows.filter((r) => up(r[3]) === sku && (up(r[2]) === 'MOVE' || up(r[2]) === 'PICKING'));
  if (!rel.length) return 'Belum ada baris MOVE atau PICKING untuk SKU ' + article + '.';
  const total = rel.reduce((s, r) => s + Math.abs(Number(r[12]) || 0), 0);
  if (Math.round(total) !== Math.round(selisihAbs)) return 'Total qty di bukti ' + total + ', harus sama dengan selisih (' + selisihAbs + ').';
  if (!rel.some((r) => r[7] || r[9])) return 'Bukti harus memuat FROMLOC atau TOLOC.';
  return '';
}
function periksaBuktiPasangan(rows, article, qty, lokMinus, lokPlus) {
  const sku = String(article || '').trim().toUpperCase(), up = (v) => String(v || '').trim().toUpperCase();
  const rel = rows.filter((r) => up(r[3]) === sku && up(r[2]) === 'MOVE');
  if (!rel.length) return 'Belum ada baris MOVE untuk SKU ' + article + '.';
  const total = rel.reduce((s, r) => s + Math.abs(Number(r[12]) || 0), 0);
  if (Math.round(total) !== Math.round(qty)) return 'Total qty MOVE ' + total + ', harus ' + qty + '.';
  if (!rel.some((r) => up(r[7]) === up(lokMinus) && up(r[9]) === up(lokPlus))) return 'Harus ada baris dengan FROMLOC ' + lokMinus + ' dan TOLOC ' + lokPlus + '.';
  return '';
}
// "MOVE:a,b;PICKING:c" → [['Move', 'a, b'], ['Picking', 'c']] (format lama tanpa jenis masuk "User WMS")
function pecahAddWho(str) {
  const s = String(str || '').trim(); if (!s) return [];
  if (!/(MOVE|PICKING|LAINNYA):/i.test(s)) return [['User WMS', s.split(',').map((x) => x.trim()).filter(Boolean).join(', ')]];
  const NAMA = { move: 'Transaksi Move terakhir', picking: 'Transaksi Picking terakhir', lainnya: 'User WMS' };
  return s.split(';').map((p) => { const i = p.indexOf(':'); if (i < 0) return null; const k = p.slice(0, i).trim().toLowerCase(); const n = p.slice(i + 1).split(',').map((x) => x.trim()).filter(Boolean).join(', '); return n ? [NAMA[k] || k, n] : null; }).filter(Boolean);
}

const Verifikasi = {
  tasks: null, form: null, at: 0, sibuk: false, galat: null, tampilan: 'task', saring: 'semua', cari: '', tertutup: {}, log: null, logDari: '', logSampai: '', logSibuk: false, req: 0, logReq: 0,

  // Dipanggil tiap pengguna atau facility berganti; jawaban permintaan yang masih berjalan diabaikan.
  reset() { Object.assign(this, { tasks: null, form: null, at: 0, sibuk: false, galat: null, tampilan: 'task', saring: 'semua', cari: '', tertutup: {}, log: null, logGalat: null, logDari: '', logSampai: '', logSibuk: false, req: this.req + 1, logReq: this.logReq + 1 }); },

  async muat(paksa) {
    if (this.sibuk || !Sesi.user) return;
    if (!paksa && this.tasks && Date.now() - this.at < 60000) return;
    const id = ++this.req, siapa = Sesi.user.username;
    this.sibuk = true; this.gambar();
    try {
      const r = await Promise.all([panggil('getOpenTasks', siapa), this.form ? this.form : panggil('getInvestigasiFormData', siapa)]);
      if (id !== this.req) return;
      this.tasks = (r[0] || []).map((t) => Object.assign(t, { namaValidator: t.namaValidator || t.nameValidator || '', selisih: Number(t.selisih) || 0 }));
      this.form = r[1]; this.at = Date.now(); this.galat = null;
    } catch (e) { if (id !== this.req) return; this.galat = e; }
    finally { if (id === this.req) { this.sibuk = false; if (Nav.aktif() === 'verifikasi') this.gambar(); this.lencana(); } }
  },
  aktif() { const batas = Date.now() - 4 * 60e3; return (this.tasks || []).filter((t) => !(this.tertutup[t.id] > batas)); },
  lencana() { if (this.tasks) { Nav.setLencana('verifikasi', this.aktif().length); if (Home.data) Home.data.openTasks = this.aktif().length; } },
  tersaring() {
    const q = this.cari.trim().toLowerCase(), f = this.saring, saya = Sesi.user.username;
    return this.aktif().filter((t) => (f === 'semua' || (f === 'kritis' && t.isCritical) || (f === 'minus' && t.selisih < 0) || (f === 'plus' && t.selisih > 0) || (f === 'saya' && t.picInvestigasi === saya)) &&
      (!q || cocok([t.lokasi, t.article, t.description, t.namaPetugas, t.picInvestigasi, t.namaValidator, t.addWhoTransaksi, namaTahap(t.statusTask)].join(' '), q)));
  },
  // Calon pasangan: SKU sama, besar selisih sama, satu minus dan satu plus (syarat yang sama dengan server).
  // Bila satu selisih punya beberapa calon, semuanya ditampilkan; bukti MOVE yang menentukan mana yang benar.
  pasangan() {
    const kunci = (t) => String(t.article).trim().toUpperCase() + '|' + Math.abs(t.selisih);
    const plus = this.aktif().filter((t) => t.selisih > 0), out = [];
    this.aktif().filter((t) => t.selisih < 0).forEach((m) => plus.forEach((p) => { if (kunci(p) === kunci(m)) out.push({ minus: m, plus: p }); }));
    return out;
  },
  // Berapa pasangan yang benar-benar bisa ditutup (tiap task hanya bisa dipakai sekali).
  jumlahPasangan(pasangan) {
    const m = {}, p = {};
    pasangan.forEach((x) => { const k = String(x.minus.article).trim().toUpperCase() + '|' + Math.abs(x.minus.selisih); (m[k] = m[k] || {})[x.minus.id] = 1; (p[k] = p[k] || {})[x.plus.id] = 1; });
    return Object.keys(m).reduce((n, k) => n + Math.min(Object.keys(m[k]).length, Object.keys(p[k]).length), 0);
  },

  /* ---------- daftar ---------- */
  gambar() {
    const node = el('scr-verifikasi');
    const kanan = h`<button type="button" class="iconbtn${this.sibuk ? ' is-putar' : ''}" data-aksi="verif-segar" aria-label="Perbarui">${ikon('segar')}</button>`;
    if (!this.tasks) { isi(node, h`${kepala('Verifikasi', '', kanan)}${this.galat ? gagalMuat(this.galat, 'verif-segar') : kerangka(5, 84)}`); return; }
    const semua = this.aktif(), pasangan = this.pasangan(), nPasang = this.jumlahPasangan(pasangan);
    const seg = h`<div class="seg vf-seg" role="tablist">
      <button type="button" role="tab" class="${this.tampilan === 'task' ? 'is-on' : ''}" data-aksi="verif-tampilan" data-v="task">Task ${semua.length ? angka(semua.length) : ''}</button>
      <button type="button" role="tab" class="${this.tampilan === 'pasangan' ? 'is-on' : ''}" data-aksi="verif-tampilan" data-v="pasangan">Pasangan ${nPasang ? angka(nPasang) : ''}</button>
      <button type="button" role="tab" class="${this.tampilan === 'riwayat' ? 'is-on' : ''}" data-aksi="verif-tampilan" data-v="riwayat">Riwayat</button></div>`;
    let badan;
    if (this.tampilan === 'pasangan') badan = this.gambarPasangan(pasangan);
    else if (this.tampilan === 'riwayat') badan = this.gambarRiwayat();
    else badan = this.gambarTask(semua);
    isi(node, h`${kepala('Verifikasi', semua.length ? angka(semua.length) + ' selisih belum selesai' : 'Semua selisih sudah selesai', kanan)}
      ${this.galat ? h`<div class="info info--jingga" style="margin-bottom:12px">${ikon('awan_mati')}<div>Belum bisa diperbarui: ${this.galat.message}</div></div>` : ''}
      ${seg}${badan}`);
    const c = el('vfCari');
    if (c) c.addEventListener('input', tunda(() => { this.cari = c.value; isi(el('vfDaftar'), this.daftarTask(this.tersaring())); }, 160));
    if (this.tampilan === 'riwayat') {
      ['vfLogDari', 'vfLogSampai'].forEach((id) => el(id).addEventListener('change', () => { this.logDari = el('vfLogDari').value; this.logSampai = el('vfLogSampai').value; this.muatLog(); }));
      if (!this.log && !this.logSibuk && !this.logGalat) this.muatLog();
    }
  },
  gambarTask(semua) {
    if (!semua.length) return h`<section class="card">${kosong('cek', 'Semua selisih sudah selesai', 'Selisih baru muncul di sini setelah validasi memastikan hasil hitung berbeda dari stok.')}</section>`;
    const n = (fn) => semua.filter(fn).length, saya = Sesi.user.username;
    const cip = (v, teks, jml) => h`<button type="button" class="chip${this.saring === v ? ' is-on' : ''}" data-aksi="verif-saring" data-v="${v}">${teks} <b>${angka(jml)}</b></button>`;
    return h`<div class="cari" style="margin-bottom:10px">${ikon('cari', 'ic--kecil')}<input class="field" type="search" id="vfCari" placeholder="Cari lokasi, SKU, atau nama" value="${this.cari}" autocomplete="off"></div>
      <div class="chips" style="margin-bottom:12px" role="group" aria-label="Saring task">${cip('semua', 'Semua', semua.length)}${cip('kritis', 'Kritis', n((t) => t.isCritical))}${cip('minus', 'Minus', n((t) => t.selisih < 0))}${cip('plus', 'Plus', n((t) => t.selisih > 0))}${cip('saya', 'PIC saya', n((t) => t.picInvestigasi === saya))}</div>
      <div id="vfDaftar">${this.daftarTask(this.tersaring())}</div>
      <button type="button" class="btn btn--teks" data-aksi="verif-ekspor" style="margin:10px auto 0">${ikon('unduh', 'ic--kecil')}Ekspor daftar ini ke Excel</button>`;
  },
  tahap(status) {
    const i = Math.max(0, TAHAP.indexOf(status));
    return h`<span class="tahap" aria-hidden="true">${[0, 1, 2].map((k) => h`<i class="${k <= i ? 'is-on' : ''}"></i>`)}</span>`;
  },
  daftarTask(daftar) {
    if (!daftar.length) return h`<p class="redup" style="text-align:center;padding:24px 0">Tidak ada task yang cocok.</p>`;
    const BATAS = 80;
    return h`<div class="list">${daftar.slice(0, BATAS).map((t) => h`<button type="button" class="row vf-row" data-aksi="verif-buka" data-id="${t.id}">
        <span class="row__isi"><span class="vf-row__atas">${labelRak(t.lokasi)}<span class="row__t">${t.article}</span></span>
          <span class="row__s">${t.description}</span>
          <span class="vf-row__status">${this.tahap(t.statusTask)}<span>${namaTahap(t.statusTask)}</span>${t.isCritical ? h`<span class="pill pill--merah">Kritis, ${angka(t.umurHari)} hari</span>` : h`<span class="redup">${t.umurHari ? angka(t.umurHari) + ' hari' : 'Hari ini'}</span>`}</span></span>
        <span class="row__ekor"><span class="pill pill--${t.selisih > 0 ? 'hijau' : 'merah'} vf-selisih">${t.selisih > 0 ? '+' : '−'}${angka(Math.abs(t.selisih))}</span></span></button>`)}</div>
      ${daftar.length > BATAS ? h`<p class="redup" style="text-align:center;margin-top:12px">Dan ${angka(daftar.length - BATAS)} task lagi. Persempit dengan pencarian atau saringan.</p>` : ''}`;
  },
  gambarPasangan(pasangan) {
    if (!pasangan.length) return h`<section class="card">${kosong('kait', 'Belum ada pasangan yang cocok', 'Pasangan muncul bila ada selisih minus dan plus untuk SKU yang sama dengan qty yang sama persis.')}</section>`;
    return h`<p class="redup" style="margin:0 2px 12px;font-size:14px">Barang kurang di satu lokasi dan lebih di lokasi lain. Selesaikan keduanya sekaligus dengan satu bukti Move.${pasangan.length > this.jumlahPasangan(pasangan) ? ' Ada selisih dengan beberapa calon pasangan: pilih yang lokasinya sesuai bukti Move.' : ''}</p>
      <div class="tumpuk">${pasangan.map((p) => h`<button type="button" class="card vf-pasang" data-aksi="verif-pasang" data-minus="${p.minus.id}" data-plus="${p.plus.id}">
        <span class="row__t">${p.minus.article}</span><span class="row__s">${p.minus.description}</span>
        <span class="vf-pasang__rute"><span>${labelRak(p.minus.lokasi)}<b style="color:var(--merah-teks)">−${angka(Math.abs(p.minus.selisih))}</b></span>${ikon('kanan')}<span>${labelRak(p.plus.lokasi)}<b style="color:var(--hijau-teks)">+${angka(Math.abs(p.plus.selisih))}</b></span></span></button>`)}</div>`;
  },

  /* ---------- lembar task ---------- */
  bukaTask(id) {
    const t = (this.tasks || []).find((x) => String(x.id) === String(id)); if (!t) return;
    const f = this.form || { kategoriList: [], mapping: [], allUsers: [], statusFlow: TAHAP };
    const alur = (f.statusFlow && f.statusFlow.length ? f.statusFlow : TAHAP), lanjut = alur.slice(alur.indexOf(t.statusTask) + 1);
    const baris = (k, v) => (v ? h`<div class="rinci__baris"><span>${k}</span><b>${v}</b></div>` : '');
    const l = Lembar.buka({
      judul: 'Tindak lanjut selisih', penuh: true,
      isi: h`<div class="vf-kepala">${labelRak(t.lokasi)}<span class="pill pill--${t.selisih > 0 ? 'hijau' : 'merah'} vf-selisih">${t.selisih > 0 ? 'Plus ' : 'Minus '}${angka(Math.abs(t.selisih))}</span>${t.isCritical ? h`<span class="pill pill--merah">Kritis</span>` : ''}</div>
        <div class="vf-sku">${t.article}</div><div class="redup">${t.description}</div>
        <div class="rinci">${baris('Dihitung oleh', t.namaPetugas)}${baris('Tanggal cycle', tanggalPanjang(String(t.tanggal).slice(0, 10)))}${baris('Validator', t.namaValidator)}${baris('PIC saat ini', t.picInvestigasi)}${pecahAddWho(t.addWhoTransaksi).map((x) => baris(x[0], x[1]))}${baris('Umur task', t.umurHari ? angka(t.umurHari) + ' hari' : 'Hari ini')}</div>
        <form id="vfForm" novalidate>
          <span class="lbl">Status berikutnya</span>
          <div class="seg" id="vfStatus">${lanjut.map((s) => h`<button type="button" data-s="${s}" class="${s === 'Selesai' ? 'is-on' : ''}">${s === 'Selesai' ? 'Selesai' : namaTahap(s)}</button>`)}</div>
          <label class="lbl" for="vfKategori">Kategori selisih <small id="vfWajib1"></small></label>
          <select class="field" id="vfKategori"><option value="">Belum diketahui</option>${f.kategoriList.map((k) => h`<option value="${k}" ${t.kategori === k ? 'selected' : ''}>${k}</option>`)}</select>
          <label class="lbl" for="vfPic">PIC <small id="vfWajib2"></small></label>
          <select class="field" id="vfPic"></select>
          <label class="lbl" for="vfCatatan">Catatan</label>
          <textarea class="field" id="vfCatatan" rows="3" placeholder="Apa yang sudah dicek atau ditemukan. Contoh: sudah cek lokasi B02.031.4, belum ketemu."></textarea>
          <div class="blok" id="vfBuktiWadah" hidden>
            <label class="lbl" for="vfBukti" id="vfBuktiLabel">Bukti transaksi WMS</label>
            <textarea class="field field--tempel" id="vfBukti" rows="4" placeholder="Tempel baris dari WMS di sini" spellcheck="false"></textarea>
            <div class="ket" id="vfBuktiKet"></div>
            <div class="pesan" id="vfBuktiCek" aria-live="polite"></div>
          </div>
          <div class="pesan pesan--galat" id="vfPesan" role="alert"></div>
        </form>`,
      kaki: h`<button type="button" class="btn" id="vfSimpan">Tandai selesai</button>`
    });
    const status = () => { const on = $('#vfStatus .is-on', l.el); return on ? on.dataset.s : lanjut[lanjut.length - 1]; };
    const selK = el('vfKategori'), selP = el('vfPic'), taB = el('vfBukti');
    const isiPic = () => {
      const k = selK.value;
      if (k === 'Salah Hitung' && t.namaValidator) { isi(selP, h`<option value="${t.namaValidator}" selected>${t.namaValidator} (validator)</option>`); selP.disabled = true; return; }
      selP.disabled = false;
      const cocokK = f.mapping.filter((m) => m.kategori === k);
      const sumber = k && cocokK.length ? cocokK : f.allUsers;
      const lihat = {}, opsi = sumber.filter((u) => (lihat[u.username] ? false : (lihat[u.username] = true)));
      const ada = opsi.some((u) => u.username === t.picInvestigasi);
      isi(selP, h`<option value="">Pilih PIC</option>${opsi.map((u) => h`<option value="${u.username}" ${u.username === t.picInvestigasi ? 'selected' : ''}>${u.username}${u.role ? ' (' + namaPeran(u.role) + ')' : ''}</option>`)}${t.picInvestigasi && !ada ? h`<option value="${t.picInvestigasi}" selected>${t.picInvestigasi} (saat ini)</option>` : ''}`);
    };
    const perluBukti = () => status() === 'Selesai' && selK.value && KATEGORI_TANPA_BUKTI.indexOf(selK.value) < 0;
    const segarkan = () => {
      const final = status() === 'Selesai';
      el('vfSimpan').textContent = final ? 'Tandai selesai' : 'Simpan status';
      el('vfWajib1').textContent = final ? '' : '(boleh dikosongkan dulu)';
      el('vfWajib2').textContent = final ? '' : '(boleh dikosongkan dulu)';
      el('vfBuktiWadah').hidden = !perluBukti();
      if (perluBukti()) {
        const adj = KATEGORI_ADJUSTMENT.indexOf(selK.value) >= 0, abs = Math.abs(t.selisih);
        el('vfBuktiLabel').textContent = adj ? 'Bukti transaksi WMS (Adjustment)' : 'Bukti transaksi WMS (Move atau Picking)';
        isi(el('vfBuktiKet'), adj
          ? h`Salin barisnya dari WMS tanpa judul kolom. Jenis transaksi ADJUSTMENT, total qty <b>${selK.value === 'Adjustment Minus' ? '−' : '+'}${angka(abs)}</b>, dan TOLOC <b>${t.lokasi}</b>.`
          : h`Salin barisnya dari WMS tanpa judul kolom. Baris MOVE atau PICKING untuk SKU ini dengan total qty <b>${angka(abs)}</b> dan lokasi asal atau tujuan.`);
        cekBukti();
      }
    };
    const cekBukti = () => {
      const n = el('vfBuktiCek'); if (!taB.value.trim()) { n.textContent = ''; n.className = 'pesan'; return; }
      try {
        const rows = bacaBukti(taB.value), salah = periksaBukti(rows, t.article, Math.abs(t.selisih), selK.value, t.lokasi);
        n.className = 'pesan ' + (salah ? 'pesan--galat' : 'pesan--ok');
        n.textContent = salah || (angka(rows.length) + ' baris terbaca. Bukti sesuai.');
      } catch (e) { n.className = 'pesan pesan--galat'; n.textContent = e.message; }
    };
    $$('#vfStatus button', l.el).forEach((b) => b.addEventListener('click', () => { $$('#vfStatus button', l.el).forEach((x) => x.classList.toggle('is-on', x === b)); segarkan(); }));
    selK.addEventListener('change', () => {
      isiPic();
      if (selK.value === 'Barang Sudah di Picking') { // jalan pintas: tinggal ketuk "Tandai selesai"
        $$('#vfStatus button', l.el).forEach((x) => x.classList.toggle('is-on', x.dataset.s === 'Selesai'));
        if (!el('vfCatatan').value.trim()) el('vfCatatan').value = 'Barang sudah di picking, selisih terkonfirmasi.';
        if (!selP.value) { const m = f.mapping.find((x) => x.kategori === selK.value); selP.value = m ? m.username : Sesi.user.username; }
      }
      segarkan();
    });
    taB.addEventListener('input', tunda(cekBukti, 250));
    isiPic(); segarkan();
    el('vfSimpan').addEventListener('click', () => this.simpanTask(t, l, status()));
  },
  async simpanTask(t, l, target) {
    const kategori = el('vfKategori').value, pic = el('vfPic').value, catatan = el('vfCatatan').value.trim(), final = target === 'Selesai';
    const pesan = (m) => { el('vfPesan').textContent = m; if (m) el('vfPesan').scrollIntoView({ block: 'nearest' }); };
    if (!catatan) return pesan('Catatan belum diisi.');
    if (final && !kategori) return pesan('Pilih kategori selisih untuk menandai selesai.');
    if (final && !pic) return pesan('Pilih PIC untuk menandai selesai.');
    let bukti = null;
    if (final && KATEGORI_TANPA_BUKTI.indexOf(kategori) < 0) {
      if (!el('vfBukti').value.trim()) return pesan('Tempel bukti transaksi WMS untuk menandai selesai.');
      try { bukti = bacaBukti(el('vfBukti').value); } catch (e) { return pesan(e.message); }
      const salah = periksaBukti(bukti, t.article, Math.abs(t.selisih), kategori, t.lokasi);
      if (salah) return pesan(salah);
    }
    pesan('');
    const b = el('vfSimpan'); b.classList.add('is-sibuk'); l.terkunci = true;
    try {
      const res = await panggil('updateTaskStatus', t.id, Sesi.user.username, target, catatan, kategori, pic, bukti, t.lokasi, t.article, t.selisih);
      if (!res.success) { pesan(res.message); return; }
      if (final) { this.tertutup[t.id] = Date.now(); this.log = null; } else { t.statusTask = target; if (kategori) t.kategori = kategori; if (pic) t.picInvestigasi = pic; }
      l.terkunci = false; l.tutup();
      Toast.tampil(final ? 'Task ditandai selesai.' : 'Status diperbarui: ' + namaTahap(target) + '.');
      this.gambar(); this.lencana();
    } catch (e) { pesan(e.message); }
    finally { b.classList.remove('is-sibuk'); l.terkunci = false; }
  },

  /* ---------- lembar pasangan ---------- */
  bukaPasangan(idMinus, idPlus) {
    const p = this.pasangan().find((x) => String(x.minus.id) === String(idMinus) && String(x.plus.id) === String(idPlus)); if (!p) return;
    const qty = Math.abs(p.minus.selisih), f = this.form || { allUsers: [] };
    const l = Lembar.buka({
      judul: 'Selesaikan pasangan plus minus', penuh: true,
      isi: h`<div class="vf-sku">${p.minus.article}</div><div class="redup">${p.minus.description}</div>
        <div class="vf-pasang__rute vf-pasang__rute--besar"><span>${labelRak(p.minus.lokasi)}<b style="color:var(--merah-teks)">−${angka(qty)}</b></span>${ikon('kanan')}<span>${labelRak(p.plus.lokasi)}<b style="color:var(--hijau-teks)">+${angka(qty)}</b></span></div>
        <label class="lbl" for="vpPic">PIC</label>
        <select class="field" id="vpPic"><option value="">Pilih PIC</option>${f.allUsers.map((u) => h`<option value="${u.username}">${u.username}${u.role ? ' (' + namaPeran(u.role) + ')' : ''}</option>`)}</select>
        <label class="lbl" for="vpCatatan">Catatan</label>
        <textarea class="field" id="vpCatatan" rows="2" placeholder="Contoh: barang ketemu di lokasi tujuan, sudah dipindahkan sesuai bukti Move."></textarea>
        <label class="lbl" for="vpBukti">Bukti transaksi WMS (Move)</label>
        <textarea class="field field--tempel" id="vpBukti" rows="4" placeholder="Tempel baris dari WMS di sini" spellcheck="false"></textarea>
        <div class="ket">Salin barisnya dari WMS tanpa judul kolom. Baris MOVE untuk SKU ini dengan FROMLOC <b>${p.minus.lokasi}</b>, TOLOC <b>${p.plus.lokasi}</b>, dan total qty <b>${angka(qty)}</b>.</div>
        <div class="pesan" id="vpCek" aria-live="polite"></div>
        <div class="pesan pesan--galat" id="vpPesan" role="alert"></div>`,
      kaki: h`<button type="button" class="btn" id="vpSimpan">Tandai keduanya selesai</button>`
    });
    const cek = () => {
      const n = el('vpCek'), v = el('vpBukti').value; if (!v.trim()) { n.textContent = ''; return; }
      try { const rows = bacaBukti(v), salah = periksaBuktiPasangan(rows, p.minus.article, qty, p.minus.lokasi, p.plus.lokasi); n.className = 'pesan ' + (salah ? 'pesan--galat' : 'pesan--ok'); n.textContent = salah || (angka(rows.length) + ' baris terbaca. Bukti sesuai.'); }
      catch (e) { n.className = 'pesan pesan--galat'; n.textContent = e.message; }
    };
    el('vpBukti').addEventListener('input', tunda(cek, 250));
    el('vpSimpan').addEventListener('click', async () => {
      const pesan = (m) => { el('vpPesan').textContent = m; };
      const pic = el('vpPic').value, catatan = el('vpCatatan').value.trim();
      if (!pic) return pesan('Pilih PIC.');
      if (!catatan) return pesan('Catatan belum diisi.');
      let bukti;
      try { bukti = bacaBukti(el('vpBukti').value); } catch (e) { return pesan(e.message === 'Bukti masih kosong.' ? 'Tempel bukti Move dari WMS.' : e.message); }
      const salah = periksaBuktiPasangan(bukti, p.minus.article, qty, p.minus.lokasi, p.plus.lokasi);
      if (salah) return pesan(salah);
      pesan('');
      const b = el('vpSimpan'); b.classList.add('is-sibuk'); l.terkunci = true;
      try {
        const res = await panggil('closePlusMinusPair', p.minus.id, p.plus.id, Sesi.user.username, catatan, pic, bukti);
        if (!res.success) { pesan(res.message); return; }
        this.tasks = this.tasks.filter((t) => t.id !== p.minus.id && t.id !== p.plus.id); this.log = null;
        l.terkunci = false; l.tutup(); Toast.tampil('Kedua task ditandai selesai.');
        this.gambar(); this.lencana(); Home.at = 0;
      } catch (e) { pesan(e.message); }
      finally { b.classList.remove('is-sibuk'); l.terkunci = false; }
    });
  },

  /* ---------- riwayat ---------- */
  gambarRiwayat() {
    if (!this.logDari) { this.logSampai = hariIni(); this.logDari = geserHari(this.logSampai, -30); }
    const atas = h`<div class="rentang"><label><span>Dari</span><input type="date" class="field field--kecil" id="vfLogDari" value="${this.logDari}"></label><label><span>Sampai</span><input type="date" class="field field--kecil" id="vfLogSampai" value="${this.logSampai}"></label></div>`;
    if (this.logSibuk || !this.log) return h`${atas}${this.logGalat ? gagalMuat(this.logGalat, 'verif-log-lagi') : memuat('Memuat riwayat…')}`;
    if (!this.log.length) return h`${atas}<section class="card">${kosong('daftar', 'Belum ada task selesai', 'Tidak ada task yang selesai pada rentang tanggal ini.')}</section>`;
    return h`${atas}<div class="list">${this.log.map((x, i) => h`<button type="button" class="row vf-row" data-aksi="verif-log-buka" data-i="${i}">
      <span class="row__isi"><span class="vf-row__atas">${labelRak(x.lokasi)}<span class="row__t">${x.article}</span></span>
        <span class="row__s">${x.kategori || x.alasan}</span><span class="row__s">Diselesaikan ${x.diselesaikanOleh}, ${x.waktuSelesai}</span></span>
      <span class="row__ekor"><span class="pill vf-selisih">${Number(x.selisih) > 0 ? '+' : '−'}${angka(Math.abs(Number(x.selisih) || 0))}</span></span></button>`)}</div>`;
  },
  // Rentang tanggal bisa diganti selagi permintaan sebelumnya masih berjalan: hanya jawaban terakhir yang dipakai.
  async muatLog() {
    const id = ++this.logReq;
    this.logSibuk = true; this.logGalat = null; if (Nav.aktif() === 'verifikasi' && this.tampilan === 'riwayat') this.gambar();
    try { const r = await panggil('getTaskLog', 300, Sesi.user.username, this.logDari, this.logSampai) || []; if (id !== this.logReq) return; this.log = r; }
    catch (e) { if (id !== this.logReq) return; this.logGalat = e; this.log = null; }
    finally { if (id === this.logReq) { this.logSibuk = false; if (Nav.aktif() === 'verifikasi' && this.tampilan === 'riwayat') this.gambar(); } }
  },
  bukaLog(i) {
    const x = (this.log || [])[i]; if (!x) return;
    const baris = (k, v) => (v ? h`<div class="rinci__baris"><span>${k}</span><b>${v}</b></div>` : '');
    Lembar.buka({
      judul: 'Task selesai',
      isi: h`<div class="vf-kepala">${labelRak(x.lokasi)}<span class="pill vf-selisih">${Number(x.selisih) > 0 ? 'Plus ' : 'Minus '}${angka(Math.abs(Number(x.selisih) || 0))}</span></div>
        <div class="vf-sku">${x.article}</div><div class="redup">${x.description || ''}</div>
        <div class="rinci">${baris('Kategori', x.kategori || x.alasan)}${baris('PIC', x.picUsername)}${baris('Diselesaikan oleh', x.diselesaikanOleh)}${baris('Waktu selesai', x.waktuSelesai)}${baris('Dihitung oleh', x.namaPetugas)}${baris('Validator', x.namaValidator)}${pecahAddWho(x.addWhoTransaksi).map((y) => baris(y[0], y[1]))}</div>
        ${x.catatan ? h`<span class="lbl">Catatan</span><div class="vf-catatan">${x.catatan}</div>` : ''}`
    });
  },
  ekspor() {
    const baris = this.tersaring().map((t) => ({ Lokasi: t.lokasi, SKU: t.article, Deskripsi: t.description, 'Qty Selisih': t.selisih, Tanggal: String(t.tanggal).slice(0, 10), Petugas: t.namaPetugas, Status: namaTahap(t.statusTask), 'Umur (hari)': t.umurHari, 'PIC': t.picInvestigasi || '', Validator: t.namaValidator || '', 'User Transaksi WMS': t.addWhoTransaksi || '' }));
    eksporExcel('Verifikasi_' + hariIni() + '_' + baris.length + '_task.xlsx', [{ nama: 'Task', baris }]);
  }
};

Layar.verifikasi = { masuk() { Verifikasi.gambar(); Verifikasi.muat(false); } };
Aksi['verif-segar'] = () => Verifikasi.muat(true);
Aksi['verif-tampilan'] = (t) => { Verifikasi.tampilan = t.dataset.v; Verifikasi.gambar(); };
Aksi['verif-saring'] = (t) => { Verifikasi.saring = t.dataset.v; Verifikasi.gambar(); };
Aksi['verif-buka'] = (t) => Verifikasi.bukaTask(t.dataset.id);
Aksi['verif-pasang'] = (t) => Verifikasi.bukaPasangan(t.dataset.minus, t.dataset.plus);
Aksi['verif-ekspor'] = () => Verifikasi.ekspor();
Aksi['verif-log-lagi'] = () => Verifikasi.muatLog();
Aksi['verif-log-buka'] = (t) => Verifikasi.bukaLog(Number(t.dataset.i));
