/* Grafik SVG ringan (tanpa pustaka). Warna mengikuti token status: hijau = hit, merah = discrepancy. */

const Grafik = {
  lebar(node) { return Math.max(280, Math.round(node.clientWidth || node.parentNode.clientWidth || 320)); },

  /**
   * Batang bertumpuk (hit + discrepancy) per periode dengan garis akurasi di atasnya.
   * data: [{ label, total, hit, discrepancy, akurasi }]. Ketuk satu kolom untuk melihat angkanya di `info`.
   */
  tren(svg, data, info, pilihAwal) {
    const W = this.lebar(svg), H = 196, x0 = 6, x1 = W - 6, yAtas = 26, yBawah = H - 30;
    const n = data.length, kolom = (x1 - x0) / Math.max(1, n);
    const maks = Math.max(1, Math.max.apply(null, data.map((d) => d.total)));
    const cx = (i) => x0 + kolom * (i + 0.5);
    const tinggi = (v) => (v / maks) * (yBawah - yAtas - 34); // sisakan ruang di atas untuk garis akurasi
    const yAkurasi = (a) => yAtas + (1 - Math.max(0, Math.min(100, a) - 80) / 20) * 30; // 80–100% dipetakan ke 30px teratas
    const lb = Math.max(8, Math.min(34, kolom * 0.56));
    let out = '<line x1="' + x0 + '" y1="' + yBawah + '" x2="' + x1 + '" y2="' + yBawah + '" stroke="var(--garis)" stroke-width="1"/>';
    let jalur = '', titik = '', menggambar = false;
    data.forEach((d, i) => {
      const x = cx(i) - lb / 2;
      if (d.total > 0) {
        const hHit = tinggi(d.hit), hLain = tinggi(d.total - d.hit - d.discrepancy), hDisc = d.discrepancy > 0 ? Math.max(3, tinggi(d.discrepancy)) : 0;
        let y = yBawah;
        if (hHit > 0) { y -= hHit; out += '<rect x="' + x + '" y="' + y + '" width="' + lb + '" height="' + hHit + '" rx="3" fill="var(--hijau)"/>'; }
        if (hLain > 0.5) { y -= hLain; out += '<rect x="' + x + '" y="' + y + '" width="' + lb + '" height="' + hLain + '" fill="var(--ungu)" opacity="0.75"/>'; }
        if (hDisc > 0) { y -= hDisc; out += '<rect x="' + x + '" y="' + y + '" width="' + lb + '" height="' + hDisc + '" rx="1.5" fill="var(--merah)"/>'; }
        const ya = yAkurasi(d.akurasi);
        jalur += (menggambar ? ' L ' : ' M ') + cx(i) + ' ' + ya; menggambar = true;
        titik += '<circle cx="' + cx(i) + '" cy="' + ya + '" r="3.6" fill="var(--kertas)" stroke="var(--tinta)" stroke-width="2"/>';
      } else {
        menggambar = false;
        out += '<line x1="' + (cx(i) - 5) + '" y1="' + (yBawah - 8) + '" x2="' + (cx(i) + 5) + '" y2="' + (yBawah - 8) + '" stroke="var(--garis-kuat)" stroke-width="2" stroke-linecap="round"/>';
      }
      out += '<text x="' + cx(i) + '" y="' + (H - 10) + '" font-size="11.5" text-anchor="middle" data-lbl="' + i + '">' + esc(d.label) + '</text>';
    });
    if (jalur) out += '<path d="' + jalur.trim() + '" fill="none" stroke="var(--tinta)" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>' + titik;
    data.forEach((d, i) => { out += '<rect data-i="' + i + '" x="' + (x0 + kolom * i) + '" y="0" width="' + kolom + '" height="' + H + '" fill="transparent"/>'; });
    svg.setAttribute('viewBox', '0 0 ' + W + ' ' + H); svg.setAttribute('height', H);
    svg.innerHTML = '<rect data-sorot x="0" y="6" width="' + kolom + '" height="' + (H - 12) + '" rx="10" fill="var(--lantai)" opacity="0"/>' + out;
    const pilih = (i) => {
      const d = data[i]; if (!d) return;
      const s = $('[data-sorot]', svg); s.setAttribute('x', x0 + kolom * i); s.setAttribute('opacity', '1');
      $$('[data-lbl]', svg).forEach((t) => { const on = Number(t.dataset.lbl) === i; t.setAttribute('font-weight', on ? '700' : '400'); t.style.fill = on ? 'var(--tinta)' : ''; });
      if (info) isi(info, d.total > 0
        ? h`<b>${d.judul || d.label}</b><span>Akurasi <b>${desimal(d.akurasi)}%</b></span><span>Hit <b>${angka(d.hit)}</b></span><span>Discrepancy <b>${angka(d.discrepancy)}</b></span>`
        : h`<b>${d.judul || d.label}</b><span>Tidak ada data</span>`);
    };
    $$('rect[data-i]', svg).forEach((r) => r.addEventListener('click', () => pilih(Number(r.dataset.i))));
    let awal = pilihAwal;
    if (awal == null) { awal = -1; data.forEach((d, i) => { if (d.total > 0) awal = i; }); }
    if (awal >= 0) pilih(awal); else if (info) isi(info, h`<span>Belum ada data pada periode ini.</span>`);
  },

  /** Batang sederhana: [{ label, nilai }] — dipakai sebaran hitungan per jam. */
  batang(svg, data, warna) {
    const W = this.lebar(svg), H = 150, x0 = 4, x1 = W - 4, yAtas = 20, yBawah = H - 24;
    const n = data.length, kolom = (x1 - x0) / Math.max(1, n), lb = Math.max(8, Math.min(30, kolom * 0.62));
    const maks = Math.max(1, Math.max.apply(null, data.map((d) => d.nilai)));
    let out = '<line x1="' + x0 + '" y1="' + yBawah + '" x2="' + x1 + '" y2="' + yBawah + '" stroke="var(--garis)" stroke-width="1"/>';
    const langkah = Math.ceil(n / 9);
    data.forEach((d, i) => {
      const t = (d.nilai / maks) * (yBawah - yAtas), x = x0 + kolom * (i + 0.5) - lb / 2;
      if (d.nilai > 0) {
        out += '<rect x="' + x + '" y="' + (yBawah - Math.max(2, t)) + '" width="' + lb + '" height="' + Math.max(2, t) + '" rx="3" fill="' + (warna || 'var(--biru)') + '"/>';
        if (kolom >= 26 || d.nilai === maks) out += '<text x="' + (x + lb / 2) + '" y="' + (yBawah - Math.max(2, t) - 5) + '" font-size="11" font-weight="600" text-anchor="middle" style="fill:var(--tinta)">' + angka(d.nilai) + '</text>';
      }
      if (i % langkah === 0 || n <= 9) out += '<text x="' + (x + lb / 2) + '" y="' + (H - 7) + '" font-size="11.5" text-anchor="middle">' + esc(d.label) + '</text>';
    });
    svg.setAttribute('viewBox', '0 0 ' + W + ' ' + H); svg.setAttribute('height', H);
    svg.innerHTML = out;
  },

  /** Garis akurasi (Analytics): [{ label, akurasi }] */
  garis(svg, data) {
    const W = this.lebar(svg), H = 150, x0 = 34, x1 = W - 20, yAtas = 14, yBawah = H - 24, n = data.length;
    const min = Math.max(0, Math.min(90, Math.floor((Math.min.apply(null, data.map((d) => d.akurasi)) - 2) / 5) * 5));
    const cx = (i) => (n > 1 ? x0 + ((x1 - x0) * i) / (n - 1) : (x0 + x1) / 2);
    const cy = (a) => yBawah - ((a - min) / (100 - min)) * (yBawah - yAtas);
    let out = '';
    [min, (min + 100) / 2, 100].forEach((p) => {
      out += '<line x1="' + x0 + '" y1="' + cy(p) + '" x2="' + x1 + '" y2="' + cy(p) + '" stroke="var(--garis)" stroke-width="1" stroke-dasharray="3 4"/>' +
        '<text x="' + (x0 - 6) + '" y="' + (cy(p) + 4) + '" font-size="11" text-anchor="end">' + Math.round(p) + '</text>';
    });
    let jalur = '', titik = '';
    const langkah = Math.ceil(n / Math.max(2, Math.floor((x1 - x0) / 46))); // label ±40px; lewati sebagian bila terlalu rapat
    data.forEach((d, i) => {
      jalur += (i ? ' L ' : 'M ') + cx(i) + ' ' + cy(d.akurasi);
      titik += '<circle cx="' + cx(i) + '" cy="' + cy(d.akurasi) + '" r="' + (n > 20 ? 2 : 3.5) + '" fill="var(--hijau)"/>';
      // label terakhir selalu ditulis; label sebelum yang terakhir dilewati bila akan bertabrakan dengannya
      if (i === n - 1 || (i % langkah === 0 && n - 1 - i >= langkah * 0.75)) out += '<text x="' + cx(i) + '" y="' + (H - 7) + '" font-size="11.5" text-anchor="middle">' + esc(d.label) + '</text>';
    });
    svg.setAttribute('viewBox', '0 0 ' + W + ' ' + H); svg.setAttribute('height', H);
    svg.innerHTML = out + '<path d="' + jalur + '" fill="none" stroke="var(--hijau)" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>' + titik;
  },

  /** Donat: bagian = [{ nilai, warna }] → markup SVG (string mentah) */
  donat(bagian, tengah, ket) {
    const R = 46, K = 2 * Math.PI * R, total = bagian.reduce((s, b) => s + b.nilai, 0);
    let off = 0, out = '<circle cx="60" cy="60" r="' + R + '" fill="none" stroke="var(--lantai-2)" stroke-width="16"/>';
    if (total > 0) bagian.forEach((b) => {
      if (!b.nilai) return;
      const p = (b.nilai / total) * K;
      out += '<circle cx="60" cy="60" r="' + R + '" fill="none" stroke="' + b.warna + '" stroke-width="16" stroke-dasharray="' + Math.max(0, p - 1.5) + ' ' + (K - Math.max(0, p - 1.5)) + '" stroke-dashoffset="' + (-off) + '" transform="rotate(-90 60 60)"/>';
      off += p;
    });
    return mentah('<svg viewBox="0 0 120 120" width="120" height="120" role="img" aria-label="' + esc(ket || '') + '">' + out +
      '<text x="60" y="60" text-anchor="middle" font-size="26" font-weight="700" style="fill:var(--tinta);font-family:var(--huruf-rapat)">' + esc(tengah) + '</text>' +
      '<text x="60" y="77" text-anchor="middle" font-size="11.5" style="fill:var(--abu);font-family:var(--huruf)">' + esc(ket || '') + '</text></svg>');
  }
};
