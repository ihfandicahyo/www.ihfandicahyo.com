# [📊 Portofolio Project: Rekonsiliasi Otomatis PPN Masukan (Accurate vs Coretax DJP)](https://github.com/ACC-TAX-REIGHTEEN/PPN-Masukan-Barang-Jasa)

> **Otomatisasi pencocokan data PPN Masukan Barang & Jasa secara presisi, cepat, dan siap audit.**

---

## 🎯 1. Problem
Pelaporan dan rekonsiliasi PPN Masukan antara sistem pembukuan internal (**Accurate**) dan portal pajak resmi (**Coretax DJP**) merupakan tantangan rutin bagi tim keuangan:
* **Perbedaan Penulisan & Format:** Variasi nama vendor (misal `PT. Shell` vs `Shell Indonesia`), format nomor faktur pajak, serta penanganan notasi angka.
* **Tidak Ada Nomor Faktur pada Transaksi Jasa (JV):** Transaksi *Journal Voucher* / Jasa di Accurate tidak mencatat nomor faktur pajak, sehingga pencocokan manual 1-per-1 memakan waktu lama.
* **Volume Data & Human Error:** Pencocokan ribuan baris data secara manual sangat rentan terhadap kekeliruan input, pembulatan, dan penetapan selisih.
* **Inkonsistensi Kategori:** Data Barang dan Jasa sering kali tercampur tanpa mekanisme pemisahan otomatis yang terstruktur.

---

## 👤 2. User
* **Tim Akuntansi & Pajak (Tax & Accounting Dept.):** Menyiapkan rekonsiliasi PPN setiap masa pajak sebelum pelaporan resmi ke DJP.
* **Finance Officer:** Melakukan verifikasi faktur pajak masukan dari pemasok/vendor.
* **Auditor Internal & Eksternal:** Membutuhkan kertas kerja rekonsiliasi yang rapi, transparan, dan dapat ditelusuri (*audit trail*).

---

## 💡 3. Solution
Mengembangkan pipeline otomatisasi Python 5-langkah (`Masukan_Barang_Jasa.py`) yang menghubungkan data ekspor Accurate (`Accuratem.xls`) dan Coretax DJP (`Coretaxm.xlsx`):
* **Pemisahan Jalur Transaksi:** Memisahkan transaksi Barang dan Jasa/JV secara otomatis berbasis penanda seksi dan filter daftar vendor (`hbrg.txt` & `hjv.txt`).
* **Algoritma Match Bertingkat:**
  * **Barang:** Pencocokan presisi berbasis *Clean Invoice Number* + *Daily Aggregate Match*.
  * **Jasa/JV:** *4-Level Smart Matching* (Offset Pair, Vendor Group Sum, Single 1:1, dan Kombinatorial 2–4 entri).
* **Laporan Terpadu:** Penggabungan otomatis seluruh analisis ke dalam satu file Excel multi-sheet terformat lengkap dengan penanda selisih otomatis.

---

## ✨ 4. Key Features
1. **Normalisasi Data Cerdas:** Penanganan alias nama vendor via `config.conf`, pembersihan prefix (PT/CV/UD), dan normalisasi string faktur pajak (menghapus tanda baca & notasi ilmiah).
2. **Daily Aggregate Matching (Barang):** Toleransi selisih harian per vendor (threshold $\le 50$) untuk memaafkan perbedaan perincian faktur individual di hari yang sama.
3. **4-Level Smart Matching Engine (JV/Jasa):**
   * *Level 1:* Deteksi pasangan entri saling hapus (*Offset Pair $+X / -X$*).
   * *Level 2:* *Vendor Group Sum* (Total Coretax per vendor vs entri Accurate).
   * *Level 3:* *Single Match 1:1* dengan batas toleransi selisih.
   * *Level 4:* *Combinatorial Search* (mencocokkan 1 Coretax dengan kombinasi 2–4 entri Accurate dalam window 60 hari).
4. **Klasifikasi Selisih Otomatis:** Memberikan label otomatis seperti `Belum Input di Accurate / Beda Masa`, `Input di Accurate Tidak Dikenal Coretax`, `Selisih Pembulatan`, atau `Selisih Nominal`.
5. **Output Siap Audit:** Laporan Excel multi-sheet (`Summary Total`, `Rincian Selisih`, `Detail Data`, `Analisis JV`) lengkap dengan rumus agregat native Excel.

---

## ⚡ 5. Challenge
* **Ekspor Accurate Legacy (`.xls`):** Struktur file tanpa header baku yang bercampur antara metadata, judul, dan sub-total tabel, membutuhkan ekstraksi tingkat rendah (*low-level parsing*).
* **Format Kolom Coretax Dinamis:** Adanya pembaruan nama kolom pada sistem Coretax DJP (versi lama vs baru) diatasi dengan mekanisme *dynamic column candidate fallback*.
* **Eksplorasi Kombinatorial JV:** Mencocokkan 1 data Coretax dengan kombinasi entri Accurate tanpa memicu masalah *performance bottleneck* (dioptimalkan dengan pembatasan window hari dan ukuran pool dataset).

---

## 📈 6. Impact
* **Efisiensi Waktu Signifikan:** Proses rekonsiliasi yang biasanya membutuhkan **3 hingga 10 menit** (bahkan lebih lama jika dilakukan secara manual) kini selesai secara otomatis dalam hitungan detik.
* **Akurasi 100% Berbasis Algoritma:** Menghilangkan *human error* dalam pencocokan dan kalkulasi angka PPN.
* **Kemudahan Audit:** Mempercepat penelusuran selisih masa pajak sehingga tim pajak dapat langsung fokus melakukan tindakan koreksi pada transaksi bermasalah.

---

## 🛠️ 7. Tech Choices
* **Python 3.8+:** Bahasa pemrograman utama untuk pipeline dan orkestrasi skrip.
* **Pandas:** Digunakan untuk pemrosesan dataset, penggabungan (*merge/join*), agregasi, dan pembersihan data.
* **OpenPyXL & XlsxWriter:** Digunakan untuk pembacaan, penyusunan layout, penyalinan styling, dan pembuatan format accounting pada file `.xlsx`.
* **Xlrd:** Membaca format legacy `.xls` dari ekspor Accurate.
* **Standard Library (`itertools`, `re`, `configparser`):** Modul kombinatorik pencocokan JV, *regular expressions* perbersihan faktur, dan manajemen file konfigurasi.

---

## 📸 8. Screenshot & Layout
```
+-----------------------------------------------------------------------------------------+
|                                HABIL_ANALISIS_BARANG_DAN_JASA.XLSX                      |
+-----------------------------------------------------------------------------------------+
| [Sheet 1: Summary Total]    | [Sheet 2: Rincian Selisih] | [Sheet 3: Detail Data] | ... |
| - Ringkasan PPN per Vendor | - Daftar transaksi tidak   | - Data mentah Barang   |     |
| - Total PPN Barang & JV     |   match (|selisih| > 50) |   lengkap + match      |     |
| - Grand Total Rumus Excel   | - Keterangan penyebab      |   status               |     |
+-----------------------------+----------------------------+------------------------+-----+

[Alur Eksekusi Terminal / CLI Pipeline]:
 ├── 1_AccCleaner&PshBrgJs.py   ---> [Ekstraksi & Pisah Accurate]
 ├── 2_CtxPshBrgJs.py           ---> [Filter & Pisah Coretax]
 ├── 3_AnalyticsBrgAccCtx.py    ---> [Rekonsiliasi Barang via No. Faktur]
 ├── 4_AnalyticsJsAccCtx.py     ---> [Smart Matching JV/Jasa 4-Level]
 └── 5_MergeHasil.py            ---> [Gabung ke Workbook Terpadu Final]
```
