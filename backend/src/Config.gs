/**
 * Config.gs
 * Konstanta global & header versi (APP_VERSION, nama sheet, HEADERS, role, kategori, dsb)
 * (Hasil pemecahan dari Code.gs -- semua file .gs berbagi global scope yang sama di Apps Script,
 *  jadi urutan file tidak masalah, fungsi di file lain tetap bisa saling panggil.)
 */

/**
 * CYCLE COUNT TRANSAKSI - Backend Apps Script (v8)
 * ---------------------------------------------------
 * PERUBAHAN UTAMA DARI v7 (Fase 2 - Dashboard action-oriented + analisis lokasi/SKU/user):
 * 1. getPrioritasHariIni() BARU: ringkasan real-time untuk Home (khusus Inventory/Admin) --
 *    total task Open, total Critical, Pending Validasi, dan 5 task paling urgent. Reuse
 *    penuh getOpenTasksData_/getPendingValidasiCountData_ (helper hasil refactor dari v7,
 *    TIDAK ada scan sheet tambahan di luar yang sudah ada).
 * 2. getTopProblemLocations() & getTopProblemSKU() BARU: agregasi discrepancy final by
 *    Lokasi / Article, mirip pola getErrorAnalysis (cuma ganti dimensi group-by). Murni
 *    additive, tidak sentuh skema/workflow.
 * 3. getUserSummary() & leaderboard di getSummaryByPeriod() sekarang punya field
 *    persenDiscrepancy eksplisit (selain akurasi/persenHit yang sudah ada) sesuai
 *    permintaan -- nilainya diturunkan dari data yang sama, bukan hitungan baru.
 * NOTIFIKASI & ZONA KERJA (dari analisis 11 poin sebelumnya): SENGAJA TIDAK dikerjakan
 * sesuai keputusan -- lihat riwayat diskusi kalau mau dipertimbangkan lagi nanti.
 *
 * PERUBAHAN DARI v8.0.0 (Fase 3 - Performa User diperluas):
 * getUserInvestigasiSummary() BARU: jumlah task investigasi yang diselesaikan user (kolom
 * Diselesaikan_Oleh) + rata-rata durasi Waktu_Cycle->Waktu_Selesai_Task dalam jam. Murni
 * additive, tidak ada kolom/skema baru. Recount Rate TIDAK dapat fungsi terpisah -- reuse
 * kesalahanHitung/persenDiscrepancy yang sudah ada di getUserSummary (cuma beda label di UI).
 *
 * PERUBAHAN DARI v8.1.0 (Fase 4 - Skor Prioritas Discrepancy):
 * getOpenTasksData_() sekarang hitung skorPrioritas = (umurHari * PRIORITAS_BOBOT_UMUR) +
 * |selisih|, dipakai sebagai tie-breaker sort DI DALAM tier Critical/Non-Critical yang sama
 * (Critical tetap prioritas #1 secara SLA, tidak diubah). Otomatis dari data existing, TANPA
 * data referensi manual (Fast Moving Item/Critical Location -- itu butuh maintenance admin
 * terus-menerus, sengaja ditunda sesuai rekomendasi awal).
 *
 * PERUBAHAN DARI v8.2.0 (Fase 5 - Dashboard Trend):
 * getTrendData() BARU: data beberapa periode terakhir sekaligus (default 7 hari/6 bulan)
 * untuk grafik tren di Dashboard. SATU KALI scan Riwayat yang langsung mengelompokkan ke
 * semua periode dalam satu pass (object lookup per baris) -- BUKAN loop yang manggil
 * getSummaryByPeriod N kali, jadi biayanya tetap 1x scan berapa pun jumlah periodenya.
 * Chart-nya di-render pakai div/CSS murni (stacked bar hit/discrepancy), TANPA library
 * eksternal (Chart.js/Recharts dari CDN) -- supaya tidak tergantung koneksi internet saat
 * WebApp diakses.
 *
 * PERUBAHAN DARI v8.3.0 (Fase 6a - Export Excel Dashboard, murni UI/Index.html):
 * Tombol "Export Excel" di Dashboard -- gabungkan Leaderboard + Lokasi Bermasalah + SKU
 * Bermasalah jadi 1 file .xlsx (3 sheet), pakai data yang sudah di-render di layar (tidak
 * query ulang ke server) + SheetJS yang sudah dimuat di app ini. TIDAK ADA perubahan di
 * Code.gs untuk bagian ini -- version di-bump biar dua file tetap sinkron penomorannya.
 * Freeze header & search/filter kolom (sisa Fase 6) BELUM dikerjakan -- ditunda ke sesi
 * berikutnya karena budget sesi ini terbatas.
 *
 * PERUBAHAN DARI v8.4.0 (Fase 6b - Freeze Header + Search, murni UI/Index.html):
 * "Detail Item Bermasalah" (Dashboard) sekarang punya search box (filter lokasi/SKU/
 * deskripsi/user/status) + tabel scroll sendiri dengan header nempel di atas (desktop,
 * >=860px) -- bukan scroll seluruh halaman lagi. Task Investigasi juga dapat search box,
 * tapi filternya TOGGLE DISPLAY (bukan re-render) supaya form isian yang sedang dikerjakan
 * user tidak kebongkar/kereset cuma karena mengetik di kotak cari. TIDAK ADA perubahan di
 * Code.gs untuk bagian ini -- version di-bump biar dua file tetap sinkron penomorannya.
 * Fase 6 SELESAI dengan ini (freeze header + search sudah ada di 2 tabel utama; filter
 * per-kolom terpisah tidak dibuat karena search box global sudah cover kebutuhan yang sama
 * dengan effort jauh lebih kecil).
 *
 * PERUBAHAN UTAMA DARI v6 (Fase 1 - Task Investigasi end-to-end):
 * 1. STATUS TASK BERTAHAP: Status_Task sekarang punya 4 tahap (TASK_STATUS_FLOW): Open ->
 *    Sedang Dicari -> Menunggu Konfirmasi -> Selesai. Tidak boleh mundur, boleh loncat maju
 *    (mis. Open langsung Selesai kalau cepat ketemu). closeTask (v6) DIHAPUS, diganti
 *    updateTaskStatus() yang menangani semua transisi + catatan progres AKUMULATIF (bukan
 *    menimpa) di kolom Catatan_Penyelesaian.
 * 2. PIC DI-ASSIGN SAAT OPEN: begitu discrepancy dikonfirmasi lewat validasi, PIC_Investigasi
 *    langsung ditunjuk otomatis (round-robin Inventory+Admin, assignInitialPIC_) -- bukan
 *    ditunggu sampai closeTask seperti v6. Kategori & PIC final (sesuai mapping Master I/J/K)
 *    tetap bisa dikoreksi kapan saja lewat updateTaskStatus begitu root-cause ketahuan.
 * 3. AGING/CRITICAL OTOMATIS: getOpenTasks menghitung umur task dari Waktu_Cycle (TANPA kolom
 *    Due_Date baru) -- task masih aktif & umurnya > TASK_SLA_HARI (default 1 hari) ditandai
 *    isCritical=true, diprioritaskan tampil di atas.
 * 4. Kolom baru: Waktu_Update_Status (kapan status task terakhir berubah, beda dari Waktu_Cycle
 *    yang menandai kapan discrepancy PERTAMA terjadi).
 *
 * PERUBAHAN v6 (masih berlaku): role Outbound/Storing/Inbound (setara "Cycle" lama), validasi
 * dibagi round-robin antar Inventory, kategori selisih 8 pilihan menggantikan alasan lama.
 *
 * SHEET YANG DIPAKAI: Master, Data Count, Riwayat.
 * - Master kolom I/J/K: KATEGORI SELISIH / USER PIC / ROLE PIC -- database mapping,
 *   diisi manual oleh admin (lihat menu "Setup Mapping PIC Kategori Selisih").
 */

/**
 * PERUBAHAN DARI v8.5.0 (Fase 7 - Validator + breakdown user transaksi per jenis di Task
 * Investigasi):
 * 1. Task Investigasi (open tasks) & Riwayat Task Log sekarang menampilkan Nama_Validator --
 *    kolom ini SUDAH ADA di Riwayat (kol.13, terisi otomatis saat submitValidasi), cuma belum
 *    pernah di-return ke UI. Murni "membaca" kolom existing, tidak ada perubahan logic
 *    validasi.
 * 2. AddWho_Transaksi sekarang dipecah per jenis transaksi (Move/Picking), bukan cuma daftar
 *    nama flat campur aduk seperti sebelumnya -- supaya investigator langsung tahu "siapa yang
 *    Move" vs "siapa yang Picking". Kolom & alur datanya TETAP SAMA (masih 1 kolom yang sama,
 *    masih dibawa lewat parameter yang sama ke Riwayat), cuma FORMAT isinya berubah dari
 *    "nama1, nama2" jadi "MOVE:nama1,nama2;PICKING:nama3" (lihat encodeAddWhoByType_/
 *    decodeAddWhoByType_). Data lama (format flat, sudah kepalang tersimpan di baris existing)
 *    tetap kebaca -- otomatis masuk kelompok LAINNYA saat ditampilkan/di-merge, tidak hilang.
 *    Tidak menyentuh Qty_System/Qty_Count/Selisih/status validasi sama sekali.
 */

/**
 * PERUBAHAN DARI v8.6.0 (Fase 8 - Pembagian task equipment-aware, Reach Truck & Tangga
 * Pesawat terbatas jumlah alat):
 * 1. Layar Upload Data sekarang punya 2 field baru: "Reach Truck Ready" & "Tangga Pesawat
 *    Ready" (default diambil dari sheet Master kolom F/G yang labelnya mengandung "reach
 *    truck"/"tangga", bisa diubah manual tiap hari sesuai alat yang benar-benar jalan).
 * 2. Dari daftar petugas yang dicentang, N orang pertama otomatis jadi slot Reach Truck,
 *    M orang berikutnya jadi slot Tangga Pesawat, sisanya jadi slot Bawah -- urutan siapa
 *    masuk slot mana tidak penting, yang penting JUMLAHNYA sesuai alat yang ready (lihat
 *    importRawData).
 * 3. buildAssignedRows_ (mesin chunking effort + urutan lokasi lorong/section) TIDAK
 *    diubah sama sekali -- cuma sekarang dipanggil 3x bertahap lewat
 *    buildAssignedRowsEquipmentAware_:
 *      Tahap 1: Level 5-6 -> HANYA slot Reach Truck.
 *      Tahap 2: Level 3-4 -> slot Tangga + Reach Truck (Reach Truck bawa sisa effort dari
 *               Tahap 1, jadi otomatis lebih sedikit kebagian Tangga kalau sudah penuh di
 *               Tahap 1 -- dan otomatis menyerap kalau beban Tangga berat).
 *      Tahap 3: Level 1-2 -> slot Bawah (petugas sisa di luar Reach Truck & Tangga).
 *    Kalau suatu tahap kehabisan petugas padahal itemnya ada, item itu TIDAK dipaksakan ke
 *    orang yang tidak punya alat -- ditampilkan sebagai warning ke admin di layar Upload.
 */

/**
 * PERUBAHAN DARI v8.7.0 (Fase 9 - Verifikasi tulis-balik Waktu_Input + Log_Anomali):
 * Ditemukan kasus Waktu_Input kosong di beberapa baris Data Count padahal Status/Qty_Count/
 * Selisih/Hasil_Awal sudah terisi (submitCount jelas sudah jalan) -- root cause belum
 * ketemu (bukan edit manual, bukan onEdit, bukan importRawData, trigger time-driven juga
 * kosong). Sambil investigasi lanjut, submitCount() sekarang self-verifying:
 * 1. Setelah nulis Waktu_Input, langsung baca-ulang cell-nya. Kalau ternyata bukan Date
 *    yang valid, retry SEKALI nulis ulang cell itu saja.
 * 2. Baik retry berhasil maupun tidak, dicatat ke sheet BARU "Log_Anomali" (Waktu_Deteksi,
 *    No, Row_Index, Nama_Petugas, Waktu_Seharusnya, Nilai_Terbaca_Saat_Cek, Status_Retry) --
 *    supaya kalau kejadian lagi ada bukti konkret, bukan cuma dugaan dari Version History.
 * Murni additive: tidak mengubah skema Data Count/Riwayat, tidak mengubah alur normal kalau
 * tidak ada anomali (cuma nambah 1x getValue() read-back per submit, biaya kecil).
 */
/**
 * PERUBAHAN DARI v8.8.0 (Fase 10 - Urutan section Reach Truck genap dulu baru ganjil):
 * Khusus Level 5-6 (Tahap 1 di buildAssignedRowsEquipmentAware_), urutan section saat
 * chunking sekarang GENAP dulu naik (2,4,6,...) baru GANJIL naik (1,3,5,...) -- lihat
 * compareLocationReachTruck_ di Utils.gs -- supaya reach truck bisa jalan LURUS di satu
 * sisi lorong dulu, baru balik ke sisi satunya, daripada zig-zag kiri-kanan tiap section.
 * buildAssignedRows_ dikasih parameter opsional ke-6 (compareFn, default compareLocation_)
 * untuk ini -- Tahap 2 (Tangga) & Tahap 3 (Bawah) TIDAK mengirim compareFn, jadi urutannya
 * tetap seperti semula (naik lurus, tidak dipisah ganjil/genap). Murni additive.
 */
/**
 * PERUBAHAN DARI v8.9.0 (Fase 11 - Rollcage exception + Filter Area Gudang & Level Alat):
 * 1) Rollcage (lokasi format "A03.23.K09", ditandai huruf K di segmen level): sebelumnya
 *    ke-parse level=9 oleh extractLevel_() lalu jatuh ke fallback diam-diam getLevelGroup_()
 *    (default grup TERAKHIR = Reach Truck kalau gak match range manapun) -- salah, karena
 *    Rollcage sebenarnya Tanpa Alat Bantu. Ditambah isRollcageLocation_() +
 *    getLevelGroupByLocation_() di Utils.gs, dipakai gantiin getLevelGroup_(extractLevel_())
 *    di 4 titik ImportData.gs supaya konsisten di seluruh alur assignment.
 * 2) Filter Area Gudang + Level Alat di layar Counting (client-side murni, TIDAK mengubah
 *    urutan/logic assignment server): root cause "user lompat gudang C->W->C->W" adalah
 *    Reach Truck & Tangga itu 2 proses sorting TERPISAH (Tahap 1 & 2 di
 *    buildAssignedRowsEquipmentAware_), jadi begitu user pindah batch/tahap dia bisa balik
 *    ke gudang awal. extractArea_() (Utils.gs) baca huruf awal lokasi jadi Area, dengan
 *    pengecualian WA/WB (Wall nempel ke gudang A/B, bukan area sendiri). getMyPendingTasks
 *    sekarang kirim area+levelGroupKey+levelGroupLabel per item; client (JsCycleValidasi.html)
 *    filter+kelompokkan jadi 2 dropdown, opsi yang sudah tuntas otomatis ke-disable ("Area
 *    ini habis -> tidak bisa diklik lagi"), dropdown otomatis sembunyi kalau cuma 1 opsi.
 */
/**
 * PERUBAHAN DARI v8.10.0 (Fase 12 - Cegah task hilang tanpa jejak + Storing murni Level atas):
 * 1) SEMUA lokasi Move/Picking WAJIB jadi tugas, tanpa terkecuali -- filter "lokasi harus
 *    ada di daftar aktif Master" DIHAPUS sebagai penyebab exclude (importRawData). Lokasi
 *    di luar daftar Master TETAP diproses, cuma dicatat sebagai info (notInMasterList).
 *    Satu-satunya baris yang genuinely gak bisa diproses adalah lokasi kosong di file
 *    mentah -- itupun sekarang dihitung jelas (skippedBlankLokasi), tidak hilang diam-diam
 *    seperti sebelumnya. Baris tipe selain Move/Picking tetap diabaikan (memang di luar
 *    scope cycle count) tapi sekarang juga dihitung (skippedOtherType) buat visibilitas.
 * 2) Warning "item Level 5-6/3-4/1-2 gak ke-assign karena gak ada petugas di slot itu"
 *    sekarang DICATAT PERMANEN ke sheet baru Log_Unassigned (bukan cuma toast di Admin
 *    yang sebelumnya auto-hilang 2.5 detik). Layar Admin juga TIDAK LAGI auto-redirect
 *    kalau ada warning -- admin wajib klik "Mengerti" dulu, supaya gak kelewat baca.
 * 3) Storing/Reach Truck sekarang MURNI Level atas (5-6) saja -- overflow otomatis ke
 *    Tangga (Level 3-4) saat beban Tangga berat DIHAPUS. Level 3-4 100% jadi tanggung
 *    jawab tim Tangga sendiri, seberapa pun beban-nya (backlog per user boleh numpuk,
 *    ini konsekuensi yang diminta, bukan bug). sumEffortByUser_ (dead code akibat
 *    perubahan ini) dihapus dari ImportData.gs.
 */
/**
 * PERUBAHAN DARI v8.11.0 (Fase 13 - Koreksi: lokasi di luar Master TETAP di-skip):
 * REVERT sebagian dari Fase 12 -- ternyata lokasi yang tidak ada di daftar aktif Master
 * (kolom A) itu SENGAJA dikecualikan (lokasi virtual & floor yang memang tidak perlu
 * di-cycle), BUKAN bug seperti dugaan awal. skippedInactive dikembalikan jadi filter
 * exclude seperti semula. Yang TETAP dipertahankan dari v8.11.0 (bagian yang memang
 * perbaikan valid): skip ini tetap dihitung & ditampilkan jelas di pesan hasil upload
 * (bukan cuma tersirat), dan baris lokasi kosong (skippedBlankLokasi) & tipe selain
 * Move/Picking (skippedOtherType) tetap punya counter terpisah sendiri -- supaya 3 jenis
 * skip yang beda alasan tidak tercampur jadi satu angka yang membingungkan.
 */
/**
 * PERUBAHAN DARI v8.12.0 (Fase 14 - Urutan section genap-ganjil juga untuk Tangga):
 * Pola urutan section "genap dulu (2,4,6,...), baru ganjil (1,3,5,...)" yang sebelumnya
 * cuma dipakai Level 5-6 (Reach Truck) sekarang dipakai juga untuk Level 3-4 (Tangga
 * Pesawat) -- alasan sama: alat lebih efisien jalan lurus daripada zig-zag kiri-kanan.
 * compareLocationReachTruck_ di-rename jadi compareLocationEvenOddSection_ (Utils.gs)
 * supaya namanya tidak menyesatkan sekarang dipakai 2 tier alat, bukan cuma Reach Truck.
 * Level 1-2 (Bawah) TETAP compareLocation_ biasa (urutan lurus naik, tidak dipisah
 * genap/ganjil) -- petugas jalan kaki tanpa alat, tidak ada isu manuver alat berat.
 */
/**
 * PERUBAHAN DARI v8.13.0 (Fase 15 - UI: versi pindah ke titlebar + rename label Stock):
 * 1) Versi app pindah dari footer ("Cycle Count vX" di bawah, cuma kelihatan di Home) ke
 *    titlebar ("📋 Cycle Transaksi - vX", kelihatan di SEMUA layar) -- warna teks versi
 *    dibikin lebih redup (opacity 0.55) daripada judulnya. Murni UI, sumber data
 *    (getHomeSummary().version) tidak berubah.
 * 2) Label "File Stock by Date" di layar Upload Data diganti jadi "File Stock by
 *    Location" (juga di semua pesan error terkait) -- cuma rename teks, TIDAK ada
 *    perubahan kolom/parsing (STOCK_HEADER_ALIASES sudah lama menerima "location"/"loc"/
 *    "lokasi" sebagai alias header, jadi ini murni koreksi penamaan di UI).
 */
/**
 * PERUBAHAN DARI v8.14.0 (Fase 16 - Fix performa: proses import bisa nabrak limit eksekusi):
 * Root cause "proses terhenti karena limit": loop dedup-merge di importRawData() sebelumnya
 * memanggil sheet.getRange(...).getValue()/.setValue() SATU-SATU per baris yang di-merge
 * (sampai 4 panggilan API per baris). Kalau yang ke-merge ratusan/ribuan baris -- makin
 * mungkin sekarang karena backlog Tangga sengaja dibiarkan numpuk kalau overload (v8.11.0)
 * -- ini bisa nabrak batas waktu eksekusi Apps Script. Diperbaiki jadi: getExistingPendingMap_
 * sekalian cache qtyTransaksi & addWho (gratis, sudah kebaca di 1 pembacaan awal), merge
 * dihitung MURNI di memori (tanpa panggilan sheet sama sekali di dalam loop), lalu ditulis
 * balik dalam MAKSIMAL 4 panggilan API total (2 baca + 2 tulis, rentang gabungan baris
 * ter-kecil s/d ter-besar yang kena merge) -- BUKAN lagi per baris. Aman dari race condition
 * karena importRawData() sudah dari awal dibungkus LockService lock yang sama.
 *
 * PERUBAHAN DARI v8.15.0 (Fase 17 - Dashboard cleanup & drill-down Analisis Kesalahan):
 * 1. Card "Lokasi Bermasalah" & "SKU Bermasalah" DIHAPUS (backend getTopProblemLocations/
 *    getTopProblemSKU ikut dihapus) -- sudah redundan dengan card "Detail Item Bermasalah"
 *    yang lebih lengkap (per-item, bisa dicari & di-export).
 * 2. getErrorAnalysisDetail() BARU: drill-down per kategori pada card "Analisis Kesalahan",
 *    lazy-load sama seperti pola getBacklogDetailByDate -- klik kategori untuk lihat detail
 *    item (Lokasi, Item, Deskripsi, Qty_System, Qty_Count, Nama_Petugas).
 * 3. Standar akurasi ranking (leaderboard petugas & validator) sekarang 99% -- di bawah itu
 *    ditandai merah (sebelumnya 2 tier 95%/80%).
 * 4. Teks Outstanding Cycle disederhanakan dari "X dari Y pending/belum" jadi "tersisa X Item
 *    belum di cycle" (card utama & detail per user).
 * 5. Skala bar Hit/Discrepancy pada grafik Tren dinaikkan plafonnya (2500) supaya batang tidak
 *    mentok penuh ke atas dan berbenturan dengan garis %Accuracy -- otomatis pakai nilai riil
 *    kalau data suatu saat melebihi plafon ini.
 *
 * PERUBAHAN DARI v8.16.0 (Fase 18 - Performa: Dashboard baca sheet berkali-kali jadi 1x):
 * Root cause "buka Dashboard lambat": setiap kali Dashboard dibuka/ganti periode, client
 * manggil 3 fungsi server terpisah (getSummaryByPeriod, getErrorAnalysis, getTrendData) --
 * TIAP fungsi baca ulang SELURUH sheet Riwayat dari awal walau datanya sama persis, jadi 3x
 * scan penuh untuk 1x buka. Klik nama di leaderboard nambah 2x scan lagi (getUserSummary +
 * getUserInvestigasiSummary). Diperbaiki dengan memecah logika hitung jadi compute*FromData_
 * helper yang menerima data yang SUDAH dibaca, lalu menambahkan 2 fungsi BUNDLE baru:
 * getDashboardData() (summary+errorAnalysis+trend, 1x baca) dan getUserDashboardDetail()
 * (userSummary+userInvestigasi, 1x baca). JsDashboard.html diarahkan pakai bundle ini --
 * dari 3x+2x scan per interaksi jadi cuma 1x+1x. Fungsi individual lama (getSummaryByPeriod,
 * dst) TETAP ada apa adanya (kompatibilitas), cuma isinya sekarang reuse helper yang sama.
 *
 * PERUBAHAN DARI v8.17.0 (Fase 19 - Fix akar masalah Waktu_Input kosong di Log_Anomali):
 * Root cause dari Fase 9 (v8.7.0) akhirnya ketemu: submitCount() menulis Waktu_Input lalu
 * LANGSUNG baca-ulang cell yang sama tanpa SpreadsheetApp.flush() -- Apps Script kadang
 * tidak langsung commit perubahan ke server Sheets (di-batch, baru benar2 tersimpan saat
 * script selesai atau saat di-flush manual), jadi baca-ulang secepat itu kadang masih dapat
 * versi lama/kosong walau tulisannya sendiri sukses. Ditambahkan SpreadsheetApp.flush() tepat
 * setelah tiap penulisan Waktu_Input (baik penulisan awal maupun retry), sebelum baca-ulang
 * verifikasi jalan -- supaya verifikasinya membaca kondisi yang sudah benar2 ke-commit.
 * CATATAN: 563 baris yang sudah kejadian SEBELUM fix ini kemungkinan besar Waktu_Input di
 * Data Count/Riwayat-nya SUDAH benar (Apps Script tetap auto-commit semua perubahan saat
 * script selesai) -- Log_Anomali yang lama kemungkinan false alarm dari cara verifikasinya,
 * bukan bukti data beneran hilang. Tetap disarankan spot-check beberapa baris (mis. No 9064)
 * di Data Count untuk pastikan sebelum dipakai hitung produktivitas.
 *
 * PERUBAHAN DARI v8.17.1 (Fase 20 - Tool perbaikan Waktu_Input kosong yang sudah kejadian):
 * Setelah dicek langsung di sheet, ternyata baris Waktu_Input kosong JAUH lebih banyak dari
 * yang sempat ke-log di Log_Anomali -- mekanisme pengecekan lama (sebelum flush() di v8.17.1)
 * ternyata tidak reliable, jadi 563 baris yang ke-log cuma sebagian dari yang beneran kosong.
 * flush() mencegah kasus BARU ke depan, tapi tidak memperbaiki data lama yang sudah kadung
 * kosong. Ditambahkan repairWaktuInputDariRiwayat_() + menu "Cycle Count > Cek & Perbaiki
 * Waktu_Input Kosong" -- backfill Waktu_Input dari Riwayat.Waktu_Cycle (sumber independen,
 * dicatat lewat cara nulis yang beda / append baris baru, jadi kemungkinan besar tidak kena
 * bug yang sama), dicocokkan lewat kombinasi Lokasi+Article+Nama_Petugas+Qty_Count+Selisih+
 * Hasil_Awal. Ada mode dry-run (default) yang cuma melaporkan tanpa menulis apa pun dulu.
 *
 * PERUBAHAN DARI v8.17.2 (Fase 21 - 1 Lokasi wajib 1 orang, tidak boleh kepotong 2 user):
 * buildAssignedRows_() (dipakai semua tahap: Reach Truck, Tangga, Bawah) SEBELUMNYA membagi
 * tugas per-ITEM (per Lokasi+Article) murni berdasar kuota effort per user -- akibatnya kalau
 * 1 lokasi kebetulan punya banyak article/transaksi dan kuota user pas habis di tengah-tengah
 * lokasi itu, sisa article-nya kelanjut ke user berikutnya walau lokasinya sama persis. Sekarang
 * item2 dikelompokkan dulu per Lokasi (grup, bukan per item) sebelum dibagi -- keputusan pindah
 * ke user berikutnya cuma dievaluasi SETELAH 1 lokasi selesai semua, TIDAK PERNAH di tengah 1
 * lokasi. Konsekuensi: pembagian beban antar-user jadi sedikit kurang presisi dibanding sebelumnya
 * (bisa lebih njomplang kalau ada lokasi yang article-nya jauh lebih banyak dari lokasi lain) --
 * trade-off yang memang diminta demi "1 lokasi = 1 orang". CATATAN: fix ini khusus utk kasus
 * dalam 1x proses upload yang sama; kasus lokasi yang sama ke-assign ke orang beda di HARI/upload
 * yang berbeda beda root cause-nya (belum ditangani di sini -- perlu obrolan terpisah kalau
 * dibutuhkan, karena butuh keputusan desain soal "sticky assignment" & fallback kalau orangnya
 * lagi off).
 *
 * PERUBAHAN DARI v8.18.0 (Fase 22 - Rename Task Investigasi -> Verifikasi + bukti WMS wajib
 * + dashboard Plus/Minus):
 * 1. Menu/label "Task Investigasi" diganti jadi "Verifikasi" di semua tempat user-facing
 *    (sidebar, tombol Home, judul layar, log task selesai). Ini CUMA ganti nama tampilan --
 *    nilai status yang tersimpan di kolom Status_Task TETAP 'Open' (tidak ada migrasi data,
 *    semua perbandingan status di kode lain tetap jalan apa adanya). statusDisplayLabel_() di
 *    JsTaskInvestigasi.html yang menerjemahkan 'Open' -> "Verifikasi" saat ditampilkan ke user.
 * 2. Bukti transaksi WMS (Move/Picking) sekarang WAJIB di-paste sebelum task bisa ditutup,
 *    KECUALI kategori 'Salah Hitung' (yang memang bukan discrepancy nyata). Divalidasi 3 hal:
 *    SKU harus cocok Article task + jenis transaksi MOVE/PICKING, total Qty harus SAMA PERSIS
 *    dengan besar selisih task, dan harus ada nomor lokasi (FROMLOC/TOLOC). Data yang di-paste
 *    disimpan lengkap ke sheet BARU "Log_Bukti_Investigasi" untuk audit trail.
 * 3. Widget Home "🎯 Prioritas Hari Ini" DIGANTI TOTAL jadi "📊 Summary Plus Minus" -- 2 kotak
 *    (Plus/Minus) menampilkan total SKU & Qty dari discrepancy yang masih aktif, bisa diklik
 *    untuk expand detail per SKU. getPrioritasHariIni() dihapus, diganti getPlusMinusSummary().
 */
const APP_VERSION = 'v8.29.3';

const MASTER_SHEET_NAME = 'Master';
const SHEET_NAME = 'Data Count';
const RIWAYAT_SHEET_NAME = 'Riwayat';
const HEADERS = ['No', 'Tanggal_Upload', 'Lokasi', 'Level', 'Article', 'Description', 'Qty_Transaksi', 'Qty_System', 'Batch', 'Nama_Petugas', 'Status', 'Qty_Count', 'Selisih', 'Hasil_Awal', 'Waktu_Input', 'AddWho_Transaksi'];

// v8.7.1: Log akses -- 1 baris per user per hari, dipakai untuk mendeteksi user yang masih
// menjalankan versi/deployment lama (lihat catatAksesLog_ di HomeSummary.gs). Sengaja terpisah
// dari Data Count/Riwayat supaya tidak ikut ter-reset oleh menu "Reset Semua Hasil Count Hari Ini".
const LOG_AKSES_SHEET_NAME = 'Log_Akses';
const LOG_AKSES_HEADERS = ['Tanggal', 'Username', 'Nama_Petugas', 'Role', 'Versi_App', 'Waktu_Pertama', 'Waktu_Terakhir', 'Jumlah_Buka'];

// v8.8.0: Log anomali tulis-balik Waktu_Input (lihat submitCount & catatLogAnomali_ di
// CycleCount.gs). Sengaja terpisah dari Data Count/Riwayat -- 1 baris cuma dibuat KALAU
// terdeteksi anomali (bukan tiap submit), jadi sheet ini seharusnya tetap kosong/kecil
// selama semuanya normal.
const LOG_ANOMALI_SHEET_NAME = 'Log_Anomali';
const LOG_ANOMALI_HEADERS = ['Waktu_Deteksi', 'No', 'Row_Index', 'Nama_Petugas', 'Waktu_Seharusnya', 'Nilai_Terbaca_Saat_Cek', 'Status_Retry'];

// v8.11.0: Log permanen buat warning "item tidak ke-assign ke siapapun" saat import (mis.
// Level atas ada tapi role Storing gak dicentang hari itu). Sebelumnya cuma toast di layar
// Admin yang auto-hilang 2.5 detik -- sekarang selalu tercatat di sini juga, gak hilang.
const LOG_UNASSIGNED_SHEET_NAME = 'Log_Unassigned';
const LOG_UNASSIGNED_HEADERS = ['Waktu', 'Admin', 'Pesan_Warning'];

// v8.18.0: bukti transaksi WMS (Move/Picking) yang di-paste user saat menutup Task Investigasi
// (Verifikasi) dengan kategori BUKAN 'Salah Hitung' -- 1 baris pasted data = 1 baris di sini,
// plus kolom konteks (Waktu, ID_Riwayat, dst) di depan. Kolom TRANTYPE..ADDWHO urutannya
// SENGAJA disamakan persis dengan contoh format export WMS yang dikasih user, supaya paste
// mentah dari WMS/Excel bisa langsung dipetakan tanpa perlu mengubah urutan kolom.
const LOG_BUKTI_SHEET_NAME = 'Log_Bukti_Investigasi';
const LOG_BUKTI_HEADERS = ['Waktu_Submit', 'ID_Riwayat', 'Lokasi_Task', 'Article_Task', 'Nama_User',
  'No', 'STORERKEY', 'TRANTYPE', 'SKU', 'Description', 'SKUGROUP', 'LOT', 'FROMLOC', 'FROMID', 'TOLOC', 'TOID', 'SOURCEKEY', 'QTY', 'ADDDATE', 'ADDWHO'];

const LEVEL_GROUPS = [
  { key: 'tanpa_alat', label: 'Tanpa Alat Bantu', min: 1, max: 2, defaultProd: 150 },
  { key: 'tangga', label: 'Tangga Pesawat', min: 3, max: 4, defaultProd: 100 },
  { key: 'reach_truck', label: 'Reach Truck', min: 5, max: 6, defaultProd: 60 }
];

// v8.24.0 (FASE OVERFLOW): "Overflow" -- saat demand Level 5-6 (Reach Truck) RENDAH
// dibanding kapasitas user Storing yang standby hari itu, kelebihan orang Storing
// dialihkan bantu Level 3-4 (Tangga) SEBELUM Tahap 2 (buildAssignedRowsEquipmentAware_)
// jalan. Reach Truck tetap prioritas #1 -- overflow baru terjadi setelah kapasitas
// Reach Truck sendiri "terpenuhi" (dihitung dari produktivitas, bukan dibagi rata dari
// awal). Beda dari behavior lama sebelum v8.11.0 (yang dihapus karena overcorrection)
// karena sekarang terukur pakai kapasitas & bisa dimatikan/diubah dari UI Setting
// (Config_Sistem sheet), TANPA edit kode.
// Disimpan sebagai sheet key-value "Config_Sistem" supaya dibaca/ditulis lewat UI, bukan
// diedit manual seperti kolom F-G Master. Default value dipakai kalau baris belum ada di
// sheet (mis. project lama yang belum pernah buka menu Setting sama sekali).
const CONFIG_SISTEM_SHEET_NAME = 'Config_Sistem';
const CONFIG_SISTEM_HEADERS = ['Key', 'Value', 'Keterangan', 'Waktu_Update', 'Diubah_Oleh'];
const CONFIG_SISTEM_DEFAULTS_ = {
  overflow_enabled: 'FALSE',
  overflow_min_sisa_reach_truck: '1',
  overflow_maks_orang: '', // kosong = tanpa batas
  overflow_force_off_today: '', // diisi tanggal (yyyy-MM-dd) oleh tombol rem darurat, auto-expired besoknya
};
const CONFIG_SISTEM_KETERANGAN_ = {
  overflow_enabled: 'Jika diaktifkan, saat demand Level 5 & 6 rendah maka tugas Level 3 dan 4 akan di-assign ke Storing.',
  overflow_min_sisa_reach_truck: 'Minimal user Storing yang wajib tetap di Level 5-6, tidak ikut dialihkan.',
  overflow_maks_orang: 'Maks user Storing yang boleh dialihkan per import. Kosongkan = tanpa batas.',
  overflow_force_off_today: '(otomatis) Tanggal rem darurat overflow dari tombol "Nonaktifkan Overflow Hari Ini".',
};

// v8.24.0: daftar NIK yang boleh membuka menu "Setting Overflow" di app -- dicek dari
// username login (kolom C Master = NIK). Sengaja TERPISAH dari role Master (kolom D),
// karena akses setting ini urusan "siapa boleh utak-atik konfigurasi sistem", bukan
// urusan role operasional (Storing/Outbound/dst). NIK yang tidak terdaftar di sini TIDAK
// akan melihat menu ini sama sekali (bukan cuma disabled).
const CONFIG_AKSES_SETTING_SHEET_NAME = 'Config_Akses_Setting';
const CONFIG_AKSES_SETTING_HEADERS = ['NIK', 'Nama', 'Tanggal_Ditambahkan'];

// Role yang setara dengan "Cycle" lama -- boleh mengerjakan Cycle Transaksi, TIDAK boleh
// Validasi/Task Investigasi/Admin/Dashboard. Dipisah jadi 3 supaya bisa jadi acuan departemen
// di mapping PIC kategori selisih (kolom K Master).
const CYCLE_ROLES = ['outbound', 'storing', 'inbound'];
// termasuk 'cycle' (nilai lama) supaya user yang di Master masih tertulis "Cycle" (belum
// sempat diganti admin) tetap bisa akses menu Cycle Transaksi tanpa error.
const CYCLE_ROLES_WITH_LEGACY = CYCLE_ROLES.concat(['cycle']);
// v8.26.0: Role 'developer' = super user, bisa akses SEMUA menu seperti admin,
// tapi TIDAK PERNAH masuk daftar assignable (tidak pernah diberi tugas cycle/validasi).
// Cocok untuk developer yang perlu cek semua keadaan sistem tanpa ikut operasional.
const ALL_CYCLE_LIKE_ROLES = CYCLE_ROLES_WITH_LEGACY.concat(['inventory', 'admin', 'developer']);
const DEVELOPER_ROLE = 'developer';

// Kategori selisih yang WAJIB dipilih saat menutup Task Investigasi (menggantikan
// Alasan_Discrepancy versi lama "Salah Lokasi/Barang Tidak Ditemukan/Salah Label").
// v8.7: disederhanakan jadi 4 kategori transaksi + "Salah Hitung" (dipilih manual saat
// investigasi ketahuan bahwa VALIDATOR yang salah hitung saat blind recount, bukan discrepancy
// transaksi asli). PENTING: "Salah Hitung" JUGA di-set OTOMATIS oleh sistem (lihat submitValidasi)
// saat hasil validasi ternyata HIT (petugas awal yang salah hitung) -- kedua kasus sama-sama
// bernilai "Salah Hitung" di kolom Kategori_Selisih, DIBEDAKAN lewat Hasil_Final saat dipakai
// untuk ranking: Hasil_Final=HIT -> petugas salah hitung (auto, tidak lewat task beneran),
// Hasil_Final=DISCREPANCY -> validator salah hitung (manual, lewat Task Investigasi beneran).
// v8.21: ditambah 'Adjustment Plus' & 'Adjustment Minus' -- dipakai saat discrepancy DIKONFIRMASI
// nyata (barang plus/minus beneran, TIDAK ketemu fisik di lokasi lain), ditutup dengan bukti
// transaksi ADJUSTMENT dari WMS (lihat KATEGORI_ADJUSTMENT_ & validateBuktiTransaksi_).
// v8.21.1: ditambah 'Barang ketemu di lokasi lain' -- untuk barang yang KETEMU fisiknya di
// lokasi lain (beda dari Lebih/Kurang Move yang sudah ada), ditutup dengan bukti histori MOVE
// dari WMS, qty-nya harus sama persis dengan besar selisih task ini (lihat validateBuktiTransaksi_,
// masuk cabang default/non-Adjustment yang sama seperti Lebih Move/Kurang Move).
const TASK_KATEGORI_LIST = ['Lebih Picking', 'Kurang Picking', 'Lebih Move', 'Kurang Move', 'Barang ketemu di lokasi lain', 'Adjustment Plus', 'Adjustment Minus', 'Salah Hitung', 'Barang Sudah di Picking'];
// v8.20: kategori 'Barang Sudah di Picking' dipakai sebagai shortcut penutupan cepat di
// Task Investigasi -- begitu dipilih di dropdown, task langsung auto-close (lihat JsTaskInvestigasi.html).
// Sama seperti 'Salah Hitung', kategori ini TIDAK butuh bukti transaksi WMS (lihat KATEGORI_TANPA_BUKTI_).
const KATEGORI_TANPA_BUKTI_ = ['Salah Hitung', 'Barang Sudah di Picking'];
// v8.21: kategori "Adjustment" -- bukti WMS yang wajib di-paste HARUS bertipe transaksi
// ADJUSTMENT (bukan MOVE/PICKING), dan qty-nya harus SESUAI TANDA (bukan cuma sama besarannya
// pakai abs seperti Move/Picking) -- task Minus wajib bukti qty negatif, task Plus wajib qty
// positif. Ini mengonfirmasi barangnya BENERAN plus/minus (bukan cuma salah lokasi/pindah).
const KATEGORI_ADJUSTMENT_ = ['Adjustment Plus', 'Adjustment Minus'];
// v8.21: kategori KHUSUS yang HANYA di-set oleh alur "Penyelesaian Plus Minus" (lihat
// closePlusMinusPair di TaskInvestigasi.gs) -- SENGAJA tidak dimasukkan ke TASK_KATEGORI_LIST
// supaya tidak muncul/tidak bisa dipilih manual dari dropdown Kategori Selisih biasa di layar
// Verifikasi. Dipakai saat 2 task (1 Minus + 1 Plus, lokasi berbeda, SKU & qty sama persis)
// dipasangkan & ditutup BERSAMAAN dengan 1 bukti MOVE yang menghubungkan kedua lokasi tsb.
const KATEGORI_PLUS_MINUS_PAIR_ = 'Pindah Lokasi (Plus-Minus)';

// v7: Task Investigasi sekarang bertahap (bukan cuma Open/Selesai), supaya progres
// pencarian barang terlihat -- bukan cuma "ada task" tapi "lagi di tahap mana".
// Urutan TIDAK BOLEH mundur (lihat updateTaskStatus_) tapi BOLEH loncat maju
// (mis. dari Open langsung Selesai kalau memang cepat ketemu).
const TASK_STATUS_FLOW = ['Open', 'Sedang Dicari', 'Menunggu Konfirmasi', 'Selesai'];
// Task dianggap CRITICAL kalau masih aktif (belum Selesai) dan sudah lewat SLA ini,
// dihitung otomatis dari Waktu_Cycle -- TIDAK ada kolom Due_Date tersendiri di sheet,
// supaya tidak perlu maintenance data tambahan (SLA berubah = tinggal ganti angka ini).
const TASK_SLA_HARI = 1;

// FASE 4 (v8.2): bobot skor prioritas -- skorPrioritas = (umurHari * BOBOT_UMUR) + |selisih|.
// Otomatis dari data yang sudah ada (umur & qty selisih), TANPA data referensi manual
// (Fast Moving/Critical Location) sesuai keputusan: mulai dari yang bisa dihitung otomatis
// dulu. 1 hari umur setara 5 unit qty selisih -- tinggal ganti angka ini kalau mau
// diseimbangkan ulang (mis. kalau ternyata qty selalu mendominasi/terlalu kecil pengaruhnya).
const PRIORITAS_BOBOT_UMUR = 5;

// Riwayat: kolom 18 (index 17) dipakai untuk Kategori_Selisih (dulu bernama Alasan_Discrepancy;
// isinya sekarang salah satu dari TASK_KATEGORI_LIST, atau "Salah Hitung" otomatis).
// Kolom 23 (Assigned_Validator) & 24 (PIC_Investigasi) ditambahkan di v6.
// Kolom 25 (Waktu_Update_Status) BARU di v7 -- dipakai untuk tahu kapan status task
// terakhir berubah (beda dengan Waktu_Cycle yang menandai kapan discrepancy PERTAMA
// terjadi, dipakai untuk hitung aging/critical).
// Kolom 26 (AddWho_Transaksi) BARU -- daftar nama unik user WMS (kolom ADDWHO di file Data
// Transaksi) yang tercatat melakukan move/picking ke Lokasi+Article ini sebelum jadi task
// investigasi. Kalau lebih dari 1 user berbeda, nama-nama digabung dipisah koma.
const RIWAYAT_HEADERS = ['ID', 'Tanggal', 'Lokasi', 'Article', 'Description', 'Qty_Transaksi', 'Qty_System', 'Nama_Petugas', 'Qty_Count', 'Selisih', 'Hasil_Awal', 'Status_Validasi', 'Nama_Validator', 'Qty_Validasi', 'Hasil_Final', 'Waktu_Cycle', 'Waktu_Validasi', 'Kategori_Selisih', 'Status_Task', 'Catatan_Penyelesaian', 'Diselesaikan_Oleh', 'Waktu_Selesai_Task', 'Assigned_Validator', 'PIC_Investigasi', 'Waktu_Update_Status', 'AddWho_Transaksi', 'Pasangan_Task_ID'];
// index (0-based) kolom Pasangan_Task_ID, dihitung dinamis dari RIWAYAT_HEADERS supaya tidak
// perlu diketik ulang manual kalau urutan kolom berubah lagi -- lihat closePlusMinusPair().
const RIWAYAT_COL_PASANGAN_TASK_ID_ = RIWAYAT_HEADERS.indexOf('Pasangan_Task_ID') + 1; // 1-based utk getRange

// =====================================================================
// ARCHIVING (Fase 22 - performa: pisahkan histori lama dari data aktif)
// =====================================================================
// Sheet Riwayat dibaca PENUH (full-scan) oleh 12+ fungsi (Dashboard, Task Investigasi,
// Home widget Plus/Minus, Validasi) -- makin lama sheet ini membesar, SEMUA layar itu makin
// lambat, bukan cuma Dashboard. Solusinya: baris yang Status_Task='Selesai' DAN sudah lebih
// tua dari ARCHIVE_AGE_MONTHS dipindah ke sheet Riwayat_Archive lewat menu "Cycle Count >
// Cek & Jalankan Archive Riwayat Lama" (lihat Archiving.gs) -- Riwayat utama tetap ramping,
// cuma berisi task aktif + histori beberapa bulan terakhir. Riwayat_Archive TIDAK pernah
// dibaca otomatis oleh sistem (murni cold storage untuk kebutuhan audit/laporan manual).
const RIWAYAT_ARCHIVE_SHEET_NAME = 'Riwayat_Archive';
const ARCHIVE_AGE_MONTHS = 3;

// =====================================================================
// RINGKASAN HARIAN (Fase 22 - performa: precompute KPI/leaderboard per hari)
// =====================================================================
// Sheet kecil (1 baris = 1 tanggal cycle count) yang di-UPDATE INCREMENTAL (delta, bukan
// re-scan) di 4 titik tulis: submitCount (HIT auto), submitValidasi, updateTaskStatus,
// closePlusMinusPair (lihat RingkasanHarian.gs). Dashboard (getDashboardData,
// getUserDashboardDetail) baca sheet KECIL ini, bukan scan penuh Riwayat -- jadi kecepatan
// buka Dashboard TIDAK lagi tergantung berapa banyak baris Riwayat. KPI (Total/Hit/Disc)
// disimpan sebagai angka biasa; leaderboard petugas/validator & breakdown kategori selisih
// disimpan sebagai JSON string per tanggal (ukurannya kecil, jauh di bawah batas 50rb
// karakter per sel Google Sheets).
const RINGKASAN_HARIAN_SHEET_NAME = 'Ringkasan_Harian';
const RINGKASAN_HARIAN_HEADERS = ['Tanggal', 'KPI_Total', 'KPI_Hit', 'KPI_Disc', 'Petugas_JSON', 'Validator_JSON', 'Kategori_JSON', 'Petugas_Raw_JSON'];
// Petugas_Raw_JSON (Fase 23): BEDA dari Petugas_JSON (yang cuma hitung item yang SUDAH FINAL,
// dipakai leaderboard) -- kolom ini hitung SEMUA item yang dihitung petugas hari itu TERMASUK
// yang masih Pending, dipakai Home "Total Cycle Hari Ini" / "Selesai Hari Ini".

// =====================================================================
// ANTRIAN AKTIF (Fase 23 - performa: counter live utk Home & Summary Plus Minus)
// =====================================================================
// Beda dari Ringkasan_Harian (agregat per TANGGAL, historis, tidak berubah lagi setelah hari
// itu lewat), sheet ini nyimpen counter yang mencerminkan KONDISI SEKARANG (antrian belum
// divalidasi, SKU yang masih Plus/Minus aktif) -- 1 item bisa saja dihitung berhari-hari lalu
// tapi statusnya masih "aktif" sampai sekarang, jadi tidak bisa dikelompokkan per tanggal.
// Struktur: key-value sederhana (1 baris = 1 key), di-update +1/-1 di titik tulis yang sama
// (appendRiwayatCycle_, submitValidasi, updateTaskStatus, closePlusMinusPair) -- BUKAN scan.
const ANTRIAN_AKTIF_SHEET_NAME = 'Antrian_Aktif';
const ANTRIAN_AKTIF_HEADERS = ['Key', 'Value_JSON'];

// =====================================================================
// QUEUE DEFERRED UPDATES (v8.29.2 - performa submitCount)
// =====================================================================
// submitCount() SEBELUMNYA memanggil applyRingkasanDelta_() & incrementCounterKey_() SECARA
// LANGSUNG di jalur kritis (masing-masing = baca+tulis ke spreadsheet facility EKSTERNAL) --
// total bisa 6-8 round-trip tambahan cuma untuk update KPI dashboard, padahal user cuma butuh
// tahu submit-nya BERHASIL (tulis ke Data Count + Riwayat). Sekarang submitCount() CUKUP
// menitipkan "delta" ke 2 sheet antrian kecil ini (append cepat, TANPA baca dulu), lalu
// processDeferredQueueAllFacilities_() (lihat QueueDelta.gs, dijadwalkan tiap 1 menit lewat
// trigger installDeferredQueueTrigger_()) yang benar-benar menerapkan delta ke Ringkasan_Harian
// & Antrian_Aktif. Efeknya: submitCount() jadi jauh lebih cepat, KPI dashboard "telat" maksimal
// ~1 menit (bukan real-time lagi, tapi cukup untuk kebutuhan monitoring).
const QUEUE_RINGKASAN_SHEET_NAME = 'Queue_Ringkasan_Delta';
const QUEUE_RINGKASAN_HEADERS = ['Timestamp', 'Tanggal', 'RowJSON'];

const QUEUE_COUNTER_SHEET_NAME = 'Queue_Antrian_Counter';
const QUEUE_COUNTER_HEADERS = ['Timestamp', 'Key', 'Delta'];

// =====================================================================
// SUBMIT QUEUE (v8.30.0 - Asynchronous Queue Processing)
// =====================================================================
// Pola lama (v8.29.x): submitCount() BUKA spreadsheet facility (SpreadsheetApp.openById,
// mahal & TIDAK bisa di-cache antar eksekusi) dan tulis LANGSUNG ke Data Count + Riwayat --
// jadi 20 detik/submit walau lock sudah 16-slot, karena bottleneck-nya openById itu sendiri,
// bukan konkurensi.
//
// Pola baru: submitCount() HANYA nulis payload ke sheet "Queue_Submit_Cycle" ini, yang
// hidup di spreadsheet BOUND/aktif punya SCRIPT INI SENDIRI (SpreadsheetApp.getActiveSpreadsheet(),
// BUKAN openById -- makanya cepat), lalu langsung return ke user. Worker
// (processSubmitQueue_() di SubmitQueue.gs, dijadwalkan tiap 1 menit) yang nanti BENAR-BENAR
// buka spreadsheet facility -- tapi CUMA SEKALI per facility per siklus, dipakai utk proses
// SEMUA item yang menumpuk sekaligus (bukan 1x buka per item seperti sebelumnya).
const SUBMIT_QUEUE_SHEET_NAME = 'Queue_Submit_Cycle';
const SUBMIT_QUEUE_HEADERS = ['Timestamp', 'No', 'NamaPetugas', 'QtyCount', 'FacilityId'];

// v8.31.0: pola queue yang sama diterapkan ke submitValidasi() & updateTaskStatus() (lihat
// diskusi di ValidasiQueue.gs & TaskQueue.gs). closePlusMinusPair() SENGAJA TIDAK ikut --
// validasi pemasangannya wajib baca data TERBARU kedua task (bukan operasi frekuensi tinggi
// spt submit/validasi biasa), jadi tetap sinkron seperti sebelumnya.
const VALIDASI_QUEUE_SHEET_NAME = 'Queue_Validasi';
const VALIDASI_QUEUE_HEADERS = ['Timestamp', 'Id', 'NamaValidator', 'QtyValidasi', 'FacilityId'];

const TASK_QUEUE_SHEET_NAME = 'Queue_Task_Investigasi';
const TASK_QUEUE_HEADERS = ['Timestamp', 'Id', 'NamaUser', 'NewStatus', 'Catatan', 'Kategori', 'PicUsername', 'BuktiRowsJSON', 'FacilityId'];

