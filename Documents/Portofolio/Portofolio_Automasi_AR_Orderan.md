# [📑 Portofolio: Automasi AR Orderan](https://github.com/ACC-TAX-REIGHTEEN/Automasi-AR-Orderan)

## 🎯 Problem
- **Proses Manual yang Lambat:** Pengecekan data piutang (*Account Receivable*) dan rekap giro dari Accurate ke *Google Sheets order tracker* harian sebelumnya dilakukan secara manual satu per satu.
- **Memakan Waktu:** Membutuhkan waktu sekitar **3–5 menit per transaksi/pengecekan** jika dilakukan secara manual.
- **Risiko *Human Error*:** Rentan terjadi kesalahan catat nilai piutang, status giro, atau terlewatnya informasi pelanggan berisiko (over-limit), FRAUD dan memiliki koin bersama (owing).

---

## 👥 User
- **Tim Admin Sales & Sales Field:** Menggunakan Google Sheets sebagai *order tracker* harian dan membutuhkan informasi kondisi kredit pelanggan secara instan saat input pesanan baru.
- **Tim Finance & Accounting / Manajemen:** Membutuhkan kontrol atas plafon kredit, transparansi status piutang aktif, dan pemantauan riwayat pembayaran pelanggan.

---

## 💡 Solution
- **Pipeline Automasi Python (8-Langkah):** Sistem yang menyinkronkan data piutang (`Piutang.xls`) dan giro (`Giro.xls`) dari Accurate langsung ke Google Sheets secara real-time dan berkelanjutan.
- **Injeksi Data Ganda:** Secara otomatis mendeteksi pesanan baru dan mengisi **nilai total sisa piutang** pada sel, serta menyuntikkan **profil kredit lengkap sebagai *Cell Note*** tanpa menimpa data yang sudah ada (*only-empty fill*).

---

## ⭐ Key Features
- **Loop Sinkronisasi Real-Time:** Otomatis berjalan dalam *background loop* (interval bawaan per 15 menit) untuk menangkap pesanan baru.
- **Cell Note Terstruktur (20+ Flag Konfigurasi):** Menampilkan rincian plafon, rata-rata bayar, riwayat hari bayar, rincian faktur aktif, nilai titip bayar, flag `OWING`, hingga tanggal giro.
- **Smart Multi-Kode & Standarisasi:** Otomatis menormalisasi format penulisan kode pelanggan (misal: `SL001` → `SL-001`) serta mendukung multi-kode per sel (`MGL-001 & MGL-002`).
- **Filter Giro Kadaluarsa Auto-Clean:** Otomatis membuang giro yang tanggal cairnya sudah terlewat agar *note* selalu relevan.
- **Fallback Pelanggan Cash & Mode Demo:** Tetap mampu menampilkan profil pelanggan *cash* tanpa piutang serta mendukung pengujian tanpa URL via sampel data GitHub.

---

## 🧩 Challenge
- **Format File Legacy Accurate:** Memproses dan membersihkan format tabel `.xls` bawaan Accurate yang memiliki header dinamis, spasi *spacer*, dan format tanggal/angka yang tidak standar.
- **Variasi Input Admin:** Menangani inkonsistensi gaya penulisan kode pelanggan dari admin sales menggunakan pemecahan pola *regex*.
- **Manajemen Filter Tanggal Dinamis:** Memastikan jadwal pencairan giro yang ditampilkan pada *cell note* hanya yang benar-benar aktif/mendatang (*real-time date filtering*).

---

## 📈 Impact
- **Efisiensi Waktu Signifikan:** Memotong proses manual dari **3–5 menit menjadi hitungan detik** secara otomatis saat baris baru dibuat.
- **Pencegahan Risiko Kredit:** Meminimalkan risiko pengiriman barang ke pelanggan yang sedang bermasalah atau memiliki piutang macet melalui visibilitas status FRAUD atau `OWING` secara langsung.
- **Keputusan Transaksi Lebih Cepat:** Admin sales tidak perlu lagi membuka aplikasi Accurate secara terpisah untuk mengecek riwayat pembayaran pelanggan.

---

## 🛠️ Tech Choices
- **Core Language:** Python 3.8+
- **Data Processing & Manipulation:** `pandas`, `openpyxl`, `xlrd` (Engine pembaca & pemroses spreadsheet/legacy `.xls`)
- **Google Integration:** `gspread`, `google-auth` (Google Sheets API v4 via Service Account)
- **Utilities & Automation:** `requests`, `configparser`, `re`, `datetime`

---

## 🖼️ Screenshot / Output Visual
**Visualisasi Output pada Google Sheets:**
- **Nilai Sel Target:** Berisi nominal IDR total Sisa Piutang (misal: `1.234.567`).
- **Cell Note Hover:** Ringkasan terstruktur memuat *Header Pelanggan*, *Ringkasan Performa Piutang* (Plafon, Rata-rata Bayar, History Hari), serta *Daftar Rincian Faktur Aktif* beserta indikator `(OWING)` dan `(JT DD/MM/YY)`.

[Versi Machine Learning](https://github.com/ACC-TAX-REIGHTEEN/AR-Orderan-MachineLearning)

