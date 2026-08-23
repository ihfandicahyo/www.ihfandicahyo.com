# [📧 Case Study & Portofolio: Email Claim Sender — Shell Attachment](https://github.com/ACC-TAX-REIGHTEEN/Email-Claim-Sender-Attch-Shell/tree/main)

Documentasi studi kasus dan portofolio proyek pengiriman email klaim otomatis berbasis Python.

---

## 1. 🎯 Problem (Masalah)

Proses pengajuan klaim pembayaran/invoice ke mitra (seperti Shell) sering kali dilakukan secara massal dan manual dengan kendala-kendala berikut:

* **Inefisiensi Waktu:** Staf harus membuka email, menyusun subjek, mencari PDF invoice yang sesuai, melampirkan file, dan mengirimkannya satu per satu secara manual.
* **Tingginya Risiko *Human Error*:** Adanya potensi salah melampirkan PDF ke invoice yang salah, melewatkan baris klaim, atau terjadi kesalahan ketik pada subjek email.
* **Skalabilitas Rendah:** Proses manual tidak praktis ketika jumlah invoice mencapai puluhan atau ratusan dokumen setiap bulannya.

---

## 2. 👤 User (Pengguna Sasaran)

* **Tim Finance / Operations / Admin Klaim:** Staf operasional yang bertanggung jawab memproses dan mengoperkan berkas tagihan/invoice klaim berkala.
* **Manajemen & Supervisor:** Membutuhkan kepastian bahwa seluruh tagihan klaim terkirim secara konsisten, akurat, dan dapat dilacak.

---

## 3. 💡 Solution (Solusi)

Mengembangkan **Email Claim Sender**, sebuah aplikasi pengirim email otomatis dua tahap berbasis Python yang mengekstraksi data klaim dari file Excel (`.xlsm`), memetakan dokumen PDF lampiran secara otomatis, dan mengirimkan email secara massal melalui protokol SMTP SSL.

```
[Surat Saffiela.xlsm] + [File-File PDF]
           │
           ▼
 [1_EkstrakData.py]  ──► Ekstraksi sheet "Isian" ke isian_temp.xlsx
           │
           ▼
 [2_GmailSender.py]  ──► Pencocokan PDF, Pembuatan Body HTML, Pengiriman via SMTP SSL
           │
           ▼
    [Email Terkirim] + [Pembersihan Otomatis Folder Dapur]
```

---

## 4. ✨ Key Features (Fitur Utama)

* **Pengiriman 1-Invoice-1-Email:** Setiap baris data invoice diolah secara terisolasi menjadi satu email khusus.
* **Pencocokan PDF Otomatis:** Nama file PDF diturunkan otomatis dari nomor invoice dengan menghilangkan karakter khusus (`-` dan `/`).
* **Templating Body HTML Dynamic:** Mendukung penulisan template body HTML dengan *placeholder* dinamis `{nama_program}`.
* **Subjek Email Terstruktur:** Mengenerate subjek resmi secara otomatis (`INV PT. ABCXYZ {No Invoice Klaim}`).
* **Sesi SMTP Efisien:** Menggunakan satu koneksi SMTP SSL (port 465) terautentikasi untuk seluruh pengiriman, menghemat waktu proses.
* **Header Auto-Detection:** Mengidentifikasi posisi kolom header `No Invoice Klaim` dan `Nama Program Klaim` secara fleksibel tanpa bergantung pada posisi sel statis.
* **Pembersihan Otomatis (*Auto-Cleanup*):** Menghapus file sementara di folder kerja setelah pengiriman selesai untuk menjaga kerapian direktori.

---

## 5. ⚡ Challenge (Tantangan Eksekusi & Solusi)

| Tantangan Teknis | Solusi yang Diterapkan |
|---|---|
| **Struktur Excel Variatif:** Posisi baris header tidak selalu berada di baris pertama (`A1`). | Menerapkan pemindaian sel (*sheet scanning*) dinamis untuk menemukan koordinat header secara otomatis. |
| **File Excel Ber-macro (`.xlsm`):** Pengambilan data rumus/formula bisa corrupt jika dibaca langsung. | Menggunakan opsi `data_only=True` pada `openpyxl` untuk mengambil nilai akhir sel tanpa mengeksekusi macro. |
| **Penanganan File PDF Rusak/Hilang:** Lampiran yang tidak ditemukan dapat menghentikan seluruh antrean pengiriman. | Menerapkan *exception handling* & *logging warning*: melewatinya secara aman dan melanjutkan ke baris berikutnya. |
| **Batas Keamanan Gmail SMTP:** Login password biasa ditolak oleh sistem keamanan Google. | Mengintegrasikan autentikasi **App Password** 16-digit berbasis SSL/TLS. |

---

## 6. 📈 Impact (Dampak & Hasil)

* **Penghematan Waktu hingga 90%:** Mengubah proses manual yang memakan waktu jam menjadi hanya hitungan detik/menit per batch pengiriman.
* **Akurasi 100% pada Lampiran:** Memastikan setiap penerima menerima berkas PDF yang sesuai dengan nomor invoice klaimnya.
* **Konsistensi Format:** Seluruh email keluar memiliki struktur subjek dan format tampilan yang seragam dan profesional.

---

## 7. 🛠️ Tech Choices (Pilihan Teknologi)

* **Python 3.8+:** Bahasa utama karena keandalan scripting dan dukungan perpustakaan ekosistem yang luas.
* **`openpyxl`:** Library untuk membedah, membaca, dan mentransfer data spreadsheet Excel (`.xlsm` / `.xlsx`).
* **`smtplib` & `email.message`:** Module bawaan Python untuk mengelola koneksi SSL/TLS ke server SMTP Gmail dan menyusun email format MIME.
* **`configparser`:** Pengelola konfigurasi external (`config.conf`) untuk memisahkan logika kode dari kredensial sensitif.

---

## 8. 🖼️ Screenshot & Alur Sistem (Visual Overview)

### Tampilan Eksekusi Terminal (CLI Output)
```text
--> Memulai eksekusi 1_EkstrakData.py
--> Memulai proses ekstraksi file
--> Menyalin data ke sheet baru
--> Melakukan auto-fit pada kolom
--> File isian_temp.xlsx berhasil dibuat dan disimpan
--> Memulai eksekusi 2_GmailSender.py
--> Membaca file konfigurasi config.conf
--> Membaca data dari isian_temp.xlsx
--> Membuka koneksi SMTP Gmail
--> Email untuk INV-001/2026 berhasil dikirim beserta lampiran INV0012026.pdf
--> Email untuk INV-002/2026 berhasil dikirim beserta lampiran INV0022026.pdf
--> Seluruh proses pengiriman selesai
```
