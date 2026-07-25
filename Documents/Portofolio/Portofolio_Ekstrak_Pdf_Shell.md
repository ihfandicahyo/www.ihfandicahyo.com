# [📑 Case Study Portofolio: Otomasi Ekstrak PDF & Rekap Laporan Gudang Shell](https://github.com/ACC-TAX-REIGHTEEN/Ekstrak-PDF-Gudang-Shell)

> **Kategori:** Automation & Data Engineering / Python Tooling  
> **Waktu Pengerjaan Manual:** ~3 - 10 Menit per siklus / berkas  
> **Waktu Pengerjaan Otomatis:** < 30 Detik (Full Pipeline, full data)  

---

## 🎯 1. Problem (Masalah)
Setiap bulan, tim divisi keuangan dan perpajakan harus memproses faktur penjualan dari Shell yang dikirim melalui email dalam format PDF. Proses manual ini membutuhkan waktu **3 hingga 10 menit** per transaksi/batch dan meliputi langkah-langkah yang rentan kesalahan (*human error*):
- Mengunduh dan membuka berkas PDF faktur satu per satu dari Gmail.
- Mengetik ulang data faktur (No. Invoice, No. Surat Jalan, No. PO, Tanggal, DPP, Total, Jatuh Tempo) ke dalam spreadsheet Excel.
- Memetakan kode gudang secara manual ke cabang tujuan (seperti SMG, PATI, PWT, MKS, dll).
- Pencocokan manual Nomor Faktur Pajak dari data ekspor pajak (CTX) berdasarkan Tanggal & DPP, yang sering kali rentan tertukar saat terdapat transaksi ber-DPP sama pada tanggal yang sama.
- Menyusun rekapitulasi bulanan ke dalam template Excel dengan penataan border, merge cell, dan rumus `SUM` manual.

---

## 👤 2. User (Pengguna Sasaran)
- **Tim Tax & Accounting / Admin Gudang:** Staf yang bertanggung jawab merekap laporan bulanan dan mencocokkan Faktur Pajak dengan data penjualan.
- **Finance Operations:** Manajemen yang membutuhkan laporan rekapitulasi gudang antar-cabang secara cepat, akurat, dan terstruktur.

---

## 💡 3. Solution (Solusi)
Mengembangkan **Skrip Otomasi Python End-to-End** (`Ekstrak PDF Gudang Shell`) yang mengorkestrasi seluruh alur kerja secara mandiri tanpa perlu intervensi manual:
- Mengambil lampiran PDF secara otomatis dari Gmail API berdasarkan query & rentang tanggal.
- Mengekstrak data terstruktur dari dokumen PDF menggunakan parser tingkat lanjut.
- Melakukan pemetaan kode gudang otomatis berbasis konfigurasi terpusat (`gudang.conf`).
- Mengimplementasikan logika antrean FIFO (*First-In, First-Out*) untuk pencocokan Nomor Faktur Pajak secara presisi.
- Menyuntikkan data baru secara *incremental* ke dalam template Excel laporan rekapitulasi bulanan lengkap dengan formula dan styling otomatis.

---

## ✨ 4. Key Features (Fitur Utama)
- 📥 **Auto-Download Gmail API:** Penarikan lampiran PDF faktur otomatis berdasar kriteria pencarian query dan filter tanggal ketat.
- 🔍 **Pengekstrak PDF Cerdas (`pdfplumber`):** Mampu membaca info faktur di halaman pertama dan referensi gudang/SJ/PO di halaman berikutnya.
- 🗺️ **Dynamic Warehouse Mapping:** Pemetaan fleksibel banyak-ke-satu (contoh: `MAKASSAR`, `MAKSSAR`, `MKSR` → `MKS`).
- ⚡ **Pencocokan Faktur Pajak FIFO:** Otomasi algoritma pencocokan `(Tanggal + DPP)` berbasis antrean untuk mencegah duplikasi atau salah pasang No. FP.
- 📊 **Pengolahan Template Excel Otomatis:** Penambahan baris *incremental*, pembentukan blok bulan baru, auto-formatting border, merge cell, dan pembaruan rumus `=SUM()`.
- 🔄 **Dua Mode Eksekusi:**
  1. *Menu 1:* Full Pipeline Orchestration (Gmail → Ekstraksi PDF → Matching → Rekap Excel).
  2. *Menu 2:* Standalone Lookup (Hanya mencocokkan No. FP pada file laporan yang sudah ada).

---

## 🧠 5. Challenge (Tantangan & Hambatan)
1. **Penanganan Transaksi Duplikat (Tanggal & DPP Sama):**
   - *Tantangan:* Ketika ada dua transaksi ber-DPP senilai Rp 10.000.000 pada tanggal yang sama, fungsi `XLOOKUP` biasa di Excel selalu mengambil baris pertama.
   - *Solusi:* Membangun sistem pencocokan antrean FIFO (Queue) di Python, di mana nilai yang dicocokkan langsung di-*pop* dari antrean.
2. **Struktur PDF Multi-Halaman & Variasi Format:**
   - *Tantangan:* Teks referensi gudang dan nomor PO terletak di lokasi atau halaman yang berbeda dari nominal faktur.
   - *Solusi:* Membagi logika ekstraksi per halaman dengan penanganan penanda (*boundary markers*) yang kuat.
3. **Konsistensi Layout Excel & Merge Cell:**
   - *Tantangan:* Menyisipkan data baru di tengah-tengah sheet Excel tanpa merusak formula `SUM` eksis atau memutus border tabel.
   - *Solusi:* Manipulasi tingkat rendah DOM OpenPyXL untuk menghitung ulang koordinat range formula `SUM` dan menyatukan kembali merge cells secara dinamis.

---

## 🚀 6. Impact (Dampak & Hasil)
- ⏱️ **Efisiensi Waktu Signifikan:** Memangkas waktu proses dari **3–10 menit kerja manual** per berkas/batch menjadi hanya **kurang dari 30 detik** secara otomatis.
- 🎯 **Akurasi 100%:** Mengeliminasi kesalahan pengetikan ulang data dan salah pasang Nomor Faktur Pajak.
- 📈 **Standardisasi Laporan:** Rekapitulasi bulanan seluruh cabang (SMG, PATI, PWT, MKS, dll.) kini terstruktur seragam secara otomatis.

---

## 🛠️ 7. Tech Choices (Pilihan Teknologi)
- **Python 3.8+:** Bahasa pemrograman utama untuk orchestrator & pemrosesan data.
- **`pandas`:** Manipulasi, pembersihan, dan deduplikasi data tabel.
- **`openpyxl`:** Pembentukan file Excel `.xlsx`, manipulasi *cell styles*, *borders*, *merge cells*, dan pembuatan formula dynamic `=SUM()`.
- **`pdfplumber`:** Ekstraksi data teks dan tabel presisi tinggi dari PDF.
- **Google Gmail API (`google-api-python-client` & OAuth 2.0):** Interaksi aman dan otomatisasi pengunduhan email.

---

## 🖼️ 8. Screenshot & Visuals
*(Asset visual pendukung portofolio)*
- **Tampilan Interaktif CLI Orchestrator:** Menu pilihan proses (Menu 1 vs Menu 2).
- **Arsitektur Pipeline Data:** Diagram alir pemrosesan dari Gmail → PDF → FIFO Matching → Excel Template.
- **Hasil Akhir Laporan Rekapitulasi Excel:** Sheet gudang terisi rapi per blok bulan dengan format border, header, dan total `SUM` terstruktur.
