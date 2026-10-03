# Desain Cycle Transaksi

Catatan keputusan desain aplikasi, supaya perubahan berikutnya tetap satu suara.

## Untuk siapa

Petugas gudang yang menghitung barang sambil berdiri di antara rak: satu tangan memegang HP, sinyal sering putus, cahaya kadang redup, dan pekerjaannya berulang ratusan kali sehari. Inventory dan admin memakai aplikasi yang sama untuk memvalidasi, menutup selisih, membagi tugas, dan melihat laporan.

Tugas utama aplikasi: **membawa orang ke lokasi yang benar dan mencatat qty secepat mungkin, tanpa salah lokasi.**

## Lima prinsip

1. **Label rak adalah tokohnya.** Kode lokasi (`B03.016.5`) selalu tampil sebagai label kuning seperti label di rak. Di mode fokus label itu dipecah menjadi Lorong, Section, Level, lengkap dengan penanda tingkat. Kuning tidak dipakai untuk hal lain.
2. **Satu layar, satu pekerjaan.** Home membuka dengan satu saran tindakan berikutnya. Mode fokus hanya menampilkan satu lokasi. Rincian dibuka dari bawah (lembar), bukan pindah halaman.
3. **Warna adalah status,** dan artinya sama di KPI, lencana menu, grafik, dan pil:
   biru = cycle dan tindakan utama, jingga = belum dikerjakan, ungu = validasi, hijau = selesai/cocok/plus, merah = minus/kritis/galat.
4. **Jempol dulu.** Keypad dan tombol utama ada di bawah, informasi di atas. Mode fokus harus muat tanpa menggulir di layar 360×568 (handheld gudang). Sasaran sentuh minimal 44 px.
5. **Data yang jujur.** Tidak menulis 100% selama masih ada sisa. Qty sistem tidak tampil saat menghitung. Yang belum terkirim ke server ditandai belum terkirim.

## Warna

| Token | Terang | Arti |
|---|---|---|
| `--lantai` | `#F3F4F7` | latar halaman |
| `--kertas` | `#FFFFFF` | kartu dan lembar |
| `--tinta` | `#191B20` | teks utama |
| `--abu` | `#646B7A` | teks pendukung |
| `--biru` | `#0A66E4` | tombol utama, tautan, Cycle |
| `--kuning` | `#FFD326` | label rak saja |
| `--jingga` | `#E8830C` | outstanding, peringatan |
| `--ungu` | `#7447E0` | validasi |
| `--hijau` | `#1A9D53` | selesai, cocok, plus |
| `--merah` | `#DC3A3A` | minus, kritis, galat |

Tiap warna status punya tiga turunan: isian (`--jingga`), teks di atas latar terang (`--jingga-teks`), dan latar lembut (`--jingga-muda`). Tema gelap mengganti nilainya lewat `html[data-tema='gelap']`; teks di atas isian berwarna memakai `--di-warna` supaya kontrasnya tetap cukup di kedua tema.

## Huruf

- **Barlow** (400, 500, 600, 700) untuk antarmuka.
- **Barlow Condensed** (600, 700) untuk kode lokasi dan angka besar: rapat, tegas, dan angkanya sama lebar, seperti cetakan label gudang.
- Keduanya dibundel di dalam aplikasi (`www/fonts`), tidak diambil dari internet.
- Huruf kecil-besar biasa (bukan kapital semua), tanpa label kecil di atas judul.

## Bentuk dan gerak

- Radius mengikuti hierarki: kartu 20, lembar 26, tombol 16, isian 14, ubin ikon 13, label rak 8.
- Navigasi bawah berbentuk pil mengambang; lencana angkanya memakai warna status.
- Gerak hanya menjawab tindakan pengguna: lembar naik, kartu lokasi berikutnya bergeser masuk, garis waktu "Urungkan" menyusut. `prefers-reduced-motion` dihormati.

## Bahasa

- Istilah mengikuti yang sudah dipakai tim: Home, Cycle, Validasi, Verifikasi, Upload, Report, Outstanding Cycle, Summary plus minus.
- Tombol menyebut hasilnya ("Proses dan bagi tugas", "Tandai selesai"), dan nama itu dipakai lagi di pesan sesudahnya.
- Pesan galat menyebut apa yang terjadi dan apa yang harus dilakukan. Layar kosong mengajak bertindak.

## Pola interaksi

- **Simpan lalu urungkan.** Hasil hitung disimpan tanpa dialog konfirmasi dan bisa diurungkan 5 detik. Selama itu hasil belum dikirim, dan cocok-tidaknya belum ditampilkan (menjaga hitungan buta).
- **Antrean kirim.** Semua hasil lewat antrean di HP (`15-antrean.js`): mencoba ulang dengan jeda makin panjang, bertahan saat aplikasi ditutup.
- **Tampil dulu, perbarui kemudian.** Data terakhir disimpan di HP supaya layar langsung terisi, lalu disegarkan di belakang.
- **Periksa di tempat.** File unggahan, bukti WMS, dan daftar lokasi diperiksa begitu dimasukkan, bukan setelah tombol kirim ditekan.
