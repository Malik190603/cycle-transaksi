/* Alat kecil yang dipakai semua modul: DOM, templat aman, format angka & tanggal, ikon. */

const $ = (sel, akar) => (akar || document).querySelector(sel);
const $$ = (sel, akar) => Array.from((akar || document).querySelectorAll(sel));
const el = (id) => document.getElementById(id);

/* ---------- templat: semua nilai di-escape kecuali hasil h`` / mentah() ---------- */
const ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
const esc = (v) => String(v == null ? '' : v).replace(/[&<>"']/g, (m) => ESC[m]);
// HTML mentah hanya dikenali lewat kelas ini. Objek dari server (hasil JSON.parse) tidak mungkin
// menjadi instance-nya, jadi isi sheet yang berbentuk {"__html": "..."} tetap di-escape.
class Html { constructor(s) { this.html = s; } }
const mentah = (s) => new Html(String(s == null ? '' : s));
const nilaiHtml = (v) => (v == null || v === false || v === true) ? '' : Array.isArray(v) ? v.map(nilaiHtml).join('') : (v instanceof Html ? v.html : esc(v));
function h(str, ...nilai) {
  let out = str[0];
  for (let i = 0; i < nilai.length; i++) out += nilaiHtml(nilai[i]) + str[i + 1];
  return mentah(out);
}
const isi = (node, html) => { if (node) node.innerHTML = nilaiHtml(html); return node; };

/* ---------- penyimpanan lokal (aman bila diblokir) ---------- */
const LS = {
  get(k) { try { const v = localStorage.getItem(k); return v == null ? null : JSON.parse(v); } catch (e) { return null; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* penuh / diblokir */ } },
  del(k) { try { localStorage.removeItem(k); } catch (e) { /* abaikan */ } }
};

/* ---------- angka ---------- */
const angka = (n) => (Number(n) || 0).toLocaleString('id-ID');
const bertanda = (n) => (Number(n) > 0 ? '+' : '') + angka(n);
// Persen yang jujur: tidak pernah menulis 100% selama masih ada sisa, dan tidak 0% selama sudah ada yang selesai.
function persen(bagian, total) {
  if (!total) return '0%';
  if (bagian >= total) return '100%';
  if (bagian <= 0) return '0%';
  const p = (bagian / total) * 100;
  if (p >= 99) return String(Math.min(99.9, Math.floor(p * 10) / 10)).replace('.', ',') + '%';
  if (p < 1) return String(Math.max(0.1, Math.ceil(p * 10) / 10)).replace('.', ',') + '%';
  return Math.round(p) + '%';
}
const desimal = (n) => String(Math.round((Number(n) || 0) * 10) / 10).replace('.', ',');

/* ---------- tanggal & jam ----------
 * Jam ditampilkan menurut zona waktu HP. "Hari ini" mengikuti hari kerja server (Asia/Jakarta, UTC+7):
 * semua tanggal di Google Sheets dicatat dengan zona itu, jadi di WITA/WIT hari berganti pukul 01.00/02.00.
 */
const dua = (n) => String(n).padStart(2, '0');
const BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
const BULAN3 = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
const HARI = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
const isoTanggal = (d) => d.getFullYear() + '-' + dua(d.getMonth() + 1) + '-' + dua(d.getDate());
const hariIni = () => new Date(Date.now() + 7 * 3600000).toISOString().slice(0, 10);
const geserHari = (iso, n) => { const d = new Date(iso + 'T12:00:00'); d.setDate(d.getDate() + n); return isoTanggal(d); };
const jam = (ms) => { const d = new Date(ms); return dua(d.getHours()) + '.' + dua(d.getMinutes()); };
const tanggalPendek = (iso) => { const p = String(iso || '').split('-'); return p.length < 3 ? String(iso || '') : Number(p[2]) + ' ' + BULAN3[Number(p[1]) - 1]; };
const tanggalPanjang = (iso) => { const d = new Date(iso + 'T12:00:00'); return isNaN(d) ? String(iso || '') : HARI[d.getDay()] + ', ' + d.getDate() + ' ' + BULAN[d.getMonth()] + ' ' + d.getFullYear(); };
// Tanggal dari server bisa 'yyyy-MM-dd', 'yyyy-MM-dd HH:mm', ISO, atau teks Date bawaan → '2 Okt 2026'
function tanggalBebas(v) {
  const s = String(v == null ? '' : v).trim(); if (!s) return '';
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) return Number(m[3]) + ' ' + BULAN3[Number(m[2]) - 1] + ' ' + m[1];
  const d = new Date(s);
  return isNaN(d) ? s : d.getDate() + ' ' + BULAN3[d.getMonth()] + ' ' + d.getFullYear();
}
const bulanPendek = (ym) => { const p = String(ym || '').split('-'); return p.length < 2 ? String(ym || '') : BULAN3[Number(p[1]) - 1] + ' ' + p[0].slice(2); };
function sapaan() {
  const j = new Date().getHours();
  return j < 11 ? 'Selamat pagi' : j < 15 ? 'Selamat siang' : j < 18 ? 'Selamat sore' : 'Selamat malam';
}

/* ---------- pengguna & peran ---------- */
const PERAN = { admin: 'Admin', administrator: 'Admin', inventory: 'Inventory', outbound: 'Outbound', storing: 'Storing', inbound: 'Inbound', lp: 'LP', maintenance: 'Maintenance', developer: 'Developer', cycle: 'Cycle' };
const namaPeran = (r) => PERAN[String(r || '').toLowerCase()] || (r ? String(r).charAt(0).toUpperCase() + String(r).slice(1) : '');
// "107183.SUMARNI" → "S" (huruf pertama nama, bukan angka pertama NIK)
function inisial(username) {
  const s = String(username || '');
  const nama = s.indexOf('.') >= 0 ? s.slice(s.indexOf('.') + 1) : s;
  const m = nama.match(/[A-Za-z]/) || s.match(/[A-Za-z0-9]/);
  return m ? m[0].toUpperCase() : '?';
}

/* ---------- lokasi rak ---------- */
// "B03.016.5" → lorong B03, section 016, level 5. Rollcage ("A03.23.K09") tidak memakai alat bantu.
function pecahLokasi(lokasi) {
  const raw = String(lokasi || '').trim().toUpperCase();
  const bagian = raw.split('.').filter((b) => b !== '');
  const akhir = bagian[bagian.length - 1] || '';
  const rollcage = /K/.test(akhir) && bagian.length > 1;
  const m = akhir.match(/\d+/);
  const level = !rollcage && m && bagian.length > 1 ? parseInt(m[0], 10) : null;
  return { raw, bagian, level, rollcage };
}
const ALAT = {
  tanpa_alat: { nama: 'Tanpa alat bantu', ikon: 'jalan' },
  tangga: { nama: 'Tangga pesawat', ikon: 'tangga' },
  reach_truck: { nama: 'Reach truck', ikon: 'forklift' }
};
function alatLokasi(lokasi) {
  const p = pecahLokasi(lokasi);
  if (p.rollcage || p.level == null || p.level <= 2) return 'tanpa_alat';
  return p.level <= 4 ? 'tangga' : 'reach_truck';
}
// Area gudang dari huruf awal lokasi; "WA"/"WB" (rak dinding) ikut gudang A/B. Sama dengan extractArea_ di backend.
function areaLokasi(lokasi) {
  const m = String(lokasi || '').trim().toUpperCase().match(/^([A-Z]+)/);
  if (!m) return '';
  return m[1].length >= 2 && m[1].charAt(0) === 'W' ? m[1].charAt(1) : m[1].charAt(0);
}
// Urutan jalan: gudang → lorong → section → level (sama dengan compareLocation_ di backend).
function kunciLokasi(lokasi) {
  const p = String(lokasi || '').trim().toUpperCase().split('.');
  const a = (p[0] || '').match(/^([A-Z]+)?(\d+)?/) || [];
  const sec = (p[1] || '').match(/\d+/), lv = (p[p.length - 1] || '').match(/\d+/);
  return [a[1] || '', a[2] ? parseInt(a[2], 10) : 999999, sec ? parseInt(sec[0], 10) : 999999, lv ? parseInt(lv[0], 10) : 999999];
}
function bandingLokasi(x, y) {
  const a = kunciLokasi(x), b = kunciLokasi(y);
  for (let i = 0; i < 4; i++) { if (a[i] !== b[i]) return a[i] < b[i] ? -1 : 1; }
  return 0;
}
function labelRak(lokasi) { return h`<span class="rak">${lokasi}</span>`; }
function labelRakBesar(lokasi) {
  const p = pecahLokasi(lokasi);
  if (p.bagian.length !== 3) return h`<div class="rak-besar rak-besar--utuh"><div class="rak-besar__bagian"><span class="rak-besar__nilai">${p.raw}</span><span class="rak-besar__nama">Lokasi</span></div></div>`;
  const nama = ['Lorong', 'Section', p.rollcage ? 'Rollcage' : 'Level'];
  const tangga = p.level && p.level <= 6
    ? h`<div class="tingkat" aria-hidden="true">${[6, 5, 4, 3, 2, 1].map((n) => h`<i class="${n === p.level ? 'is-on' : ''}"></i>`)}</div>` : '';
  return h`<div class="rak-besar" aria-label="Lokasi ${p.raw}">${p.bagian.map((b, i) => h`<div class="rak-besar__bagian"><span class="rak-besar__nilai">${b}</span><span class="rak-besar__nama">${nama[i]}</span></div>`)}${tangga}</div>`;
}

/* ---------- lain-lain ---------- */
const tunggu = (ms) => new Promise((r) => setTimeout(r, ms));
function tunda(fn, ms) { let t; return function () { const a = arguments, s = this; clearTimeout(t); t = setTimeout(() => fn.apply(s, a), ms); }; }
const cocok = (teks, q) => String(teks || '').toLowerCase().indexOf(q) !== -1;

/* ---------- ikon garis (24x24, satu gaya untuk seluruh aplikasi) ---------- */
const IKON = {
  home: '<path d="M3.5 10.5 12 3.5l8.5 7"/><path d="M5.5 9.5V20h4.6v-5.6h3.8V20h4.6V9.5"/>',
  kotak: '<path d="m12 3 8 4v10l-8 4-8-4V7l8-4z"/><path d="m4 7 8 4 8-4"/><path d="M12 11v10"/>',
  perisai: '<path d="M12 3.2 19 6v5.2c0 4.3-2.8 8-7 9.6-4.2-1.6-7-5.3-7-9.6V6l7-2.8z"/><path d="m8.8 12 2.3 2.3 4.2-4.4"/>',
  selisih: '<circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/><path d="M8.3 11h5.4M11 8.3v5.4"/>',
  unggah: '<path d="M12 16V4.5"/><path d="m7 9 5-5 5 5"/><path d="M4 15.5V19a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3.5"/>',
  grafik: '<path d="M4 20h16"/><path d="M7.5 20v-6"/><path d="M12 20V6"/><path d="M16.5 20v-9.5"/>',
  segar: '<path d="M19.5 11A7.6 7.6 0 0 0 6 7.2L4.5 9"/><path d="M4.5 4.5V9H9"/><path d="M4.5 13A7.6 7.6 0 0 0 18 16.8l1.5-1.8"/><path d="M19.5 19.5V15H15"/>',
  kanan: '<path d="m9 6 6 6-6 6"/>',
  kiri: '<path d="m15 6-6 6 6 6"/>',
  bawah: '<path d="m6 9 6 6 6-6"/>',
  atas: '<path d="m6 15 6-6 6 6"/>',
  tutup: '<path d="M6 6l12 12M18 6 6 18"/>',
  cek: '<path d="m5 12.5 4.5 4.5L19 7.5"/>',
  jam: '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/>',
  awas: '<path d="M12 4 2.8 19.5h18.4L12 4z"/><path d="M12 10v4.3"/><path d="M12 17v.2"/>',
  info: '<circle cx="12" cy="12" r="8.5"/><path d="M12 11v5"/><path d="M12 7.7v.2"/>',
  naik: '<path d="M12 19V5"/><path d="m6 11 6-6 6 6"/>',
  turun: '<path d="M12 5v14"/><path d="m6 13 6 6 6-6"/>',
  urung: '<path d="M8 5 3 10l5 5"/><path d="M3 10h11a6 6 0 0 1 0 12h-3"/>',
  hapus_kiri: '<path d="M21 5H9l-6 7 6 7h12a1 1 0 0 0 1-1V6a1 1 0 0 0-1-1z"/><path d="m12 9.5 5 5m0-5-5 5"/>',
  orang: '<circle cx="12" cy="8" r="4"/><path d="M4.5 20c1-3.5 4-5 7.5-5s6.5 1.5 7.5 5"/>',
  orang2: '<circle cx="9" cy="8" r="3.5"/><path d="M2.5 19.5c.8-3 3.3-4.5 6.5-4.5s5.7 1.5 6.5 4.5"/><path d="M16 4.8a3.5 3.5 0 0 1 0 6.4"/><path d="M18.5 15.3c1.5.6 2.6 1.9 3 4.2"/>',
  bulan: '<path d="M20 14.5A8.5 8.5 0 0 1 9.5 4a7.2 7.2 0 1 0 10.5 10.5z"/>',
  keluar: '<path d="M10 4H6a1 1 0 0 0-1 1v14a1 1 0 0 0 1 1h4"/><path d="m15 8 4 4-4 4"/><path d="M19 12H9"/>',
  atur: '<path d="M4 7h9M19 7h1M4 17h3M13 17h7"/><circle cx="16" cy="7" r="2.5"/><circle cx="10" cy="17" r="2.5"/>',
  unduh: '<path d="M12 4v11.5"/><path d="m7 11 5 5 5-5"/><path d="M4 20h16"/>',
  saring: '<path d="M4 6h16M7 12h10M10 18h4"/>',
  cari: '<circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/>',
  berkas: '<path d="M7 3h7l5 5v12a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"/><path d="M14 3v5h5"/>',
  tangga: '<path d="M8 3v18M16 3v18M8 7.5h8M8 12h8M8 16.5h8"/>',
  forklift: '<path d="M3 16.5h9.5V8H8.5L5 12.5v4z"/><path d="M15.5 4v12.5H21"/><circle cx="6.5" cy="18.5" r="1.8"/><circle cx="11" cy="18.5" r="1.8"/>',
  jalan: '<circle cx="12.5" cy="4.8" r="1.8"/><path d="m9 20 2.5-6-2-2.5 1-4 3 2 3 1"/><path d="m13.5 14 2 6"/><path d="m10.5 7.5-3 1.5-1 3"/>',
  tambah: '<path d="M12 5v14M5 12h14"/>',
  kurang: '<path d="M5 12h14"/>',
  kait: '<path d="m9 15 6-6"/><path d="m10.5 6.5 1-1a4.2 4.2 0 0 1 6 6l-1 1"/><path d="m13.5 17.5-1 1a4.2 4.2 0 0 1-6-6l1-1"/>',
  gedung: '<path d="M5 20V5h9v15"/><path d="M14 10h5v10"/><path d="M3 20h18"/><path d="M8 9h3M8 13h3"/>',
  kunci: '<circle cx="8" cy="15.5" r="4"/><path d="m11 12.5 8.5-8.5"/><path d="m16 7.5 3 3"/>',
  sampah: '<path d="M5 7h14"/><path d="M10 7V4h4v3"/><path d="m7 7 1 13h8l1-13"/>',
  pensil: '<path d="m4 20 4-1L19 8l-3-3L5 16l-1 4z"/>',
  pin: '<path d="M12 21s7-6.2 7-11a7 7 0 1 0-14 0c0 4.8 7 11 7 11z"/><circle cx="12" cy="10" r="2.5"/>',
  kilat: '<path d="M13 3 5 13.5h6L10 21l8-10.5h-6L13 3z"/>',
  tren: '<path d="m4 17.5 5-6 4 3 7-8.5"/><path d="M15.5 6H20v4.5"/>',
  awan_mati: '<path d="M3 4l18 17"/><path d="M8.5 8.2A5.5 5.5 0 0 1 17.8 10 4.2 4.2 0 0 1 20 17.5"/><path d="M17 19H7.5a4.5 4.5 0 0 1-1.5-8.7"/>',
  awan_kirim: '<path d="M7.5 18.5a4.5 4.5 0 0 1-.5-9A5.5 5.5 0 0 1 17.8 10a4.3 4.3 0 0 1-.8 8.5"/><path d="M12 20v-8"/><path d="m9 14.5 3-3 3 3"/>',
  daftar: '<path d="M8 6.5h12M8 12h12M8 17.5h12"/><path d="M4 6.5h.1M4 12h.1M4 17.5h.1"/>',
  lewati: '<path d="m5 5 8 7-8 7V5z"/><path d="M18 5v14"/>',
  mata: '<path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z"/><circle cx="12" cy="12" r="3"/>'
};
function ikon(nama, kelas) {
  return mentah('<svg class="ic' + (kelas ? ' ' + kelas : '') + '" viewBox="0 0 24 24" aria-hidden="true">' + (IKON[nama] || '') + '</svg>');
}
