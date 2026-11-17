# 🎓 SISTEM PRESENSI DIGITAL & KESISWAAN TERPADU
### SMK NU WARGABINANGUN CIREBON
> **Proyek Akhir Rekayasa Perangkat Lunak (RPL) — S1 Teknik Informatika**  
> *Sistem Absensi Real-Time Multi-Role Terintegrasi Telegram Bot & Kecerdasan Artifisial NVIDIA NIM (Nemotron LLM)*

---

## 📌 Identitas Pengembang & Proyek

| Informasi | Keterangan |
| :--- | :--- |
| **Nama Mahasiswa** | **Winarti** |
| **NIM** | **41257168** |
| **Kelas / Angkatan** | TI-1C / Angkatan 2025 |
| **Program Studi** | S1 Teknik Informatika |
| **Dosen Pengampu** | **Yudhistira Arie Wijaya, M.Kom., CDS.** |
| **Kelompok** | **Sunyaragi Nexus** |
| **Tema Topik Kelompok** | **E-Learning & Smart Education AI** |
| **Repositori GitHub** | [https://github.com/Wwinaa/proyek-rpl-absensiQR](https://github.com/Wwinaa/proyek-rpl-absensiQR) |
| **Mata Kuliah** | Rekayasa Perangkat Lunak (RPL) |
| **Prototipe UI/UX (Figma)** | [Tautan Desain Figma](https://www.figma.com/design/HJ3KQU3E6jYNRBvK1X99aU/Untitled?node-id=0-1) |
| **Tahun Akademik** | 2025 / 2026 |

---

## 📺 Link Video Presentasi & Live Demo (YouTube)

Tonton presentasi komprehensif, pembedahan arsitektur UML & SDD, serta demonstrasi langsung (live demo) sistem:

- 🔗 **Tautan Langsung YouTube**: [https://youtu.be/d8n-Pi-SUn4?si=7W71KGDXO2uolbbd](https://youtu.be/d8n-Pi-SUn4?si=7W71KGDXO2uolbbd)
- 🎬 **Video Final Demo**: Menyajikan Elevator Pitch, demonstrasi scan QR Code kamera secara langsung (*live*), verifikasi notifikasi Telegram bot `@HadirioBot`, transaksi atomik *dual-sync* presensi wali kelas, pengajuan dan persetujuan surat izin sakit mandiri, percakapan cerdas dengan Asisten AI Hadir.io (4 bahasa), serta validasi pengujian sistem.  

---

## 📖 Deskripsi Aplikasi

**Sistem Presensi Digital SMK NU Wargabinangun** adalah platform presensi dan tata kelola kesiswaan berbasis web yang dirancang khusus untuk mentransformasi sistem absensi manual (kertas/buku tanda tangan) menjadi sistem digital terotomasi, transparan, dan bebas kecurangan (*zero fraud*).

Aplikasi ini menggunakan pola arsitektur **MVC (Model-View-Controller)** murni dengan native PHP 8.2 dan database relasional MariaDB/MySQL ternormalisasi 3NF sebanyak **20 tabel**. Sistem mendukung otorisasi berlapis berbasis peran (**Role-Based Access Control / RBAC**) untuk 5 tingkatan entitas pengguna:
1. **Administrator (Superuser)**: Pengelolaan penuh data akademik (tahun ajaran, jurusan, kelas), master pendidik, siswa, pembuatan kartu QR/Barcode & RFID, manajemen akun pengguna, hingga konfigurasi sistem.
2. **Kepala Sekolah (Eksekutif)**: Monitoring supervisi KBM harian, status baca jurnal guru, serta rekapitulasi kehadiran eksekutif mingguan/bulanan berstandar hari efektif (Senin–Jumat).
3. **Wali Kelas**: Manajemen presensi rombel binaan dengan fitur **Dual-Sync Otomatis**, approval perizinan siswa, dan pencatatan bimbingan/konseling pembinaan karakter.
4. **Guru Pengajar**: Pengisian jurnal pembelajaran kelas harian, presensi mandiri berbasis QR/Barcode, dan monitoring riwayat presensi.
5. **Peserta Didik (Siswa)**: Presensi kilat via scan kartu barcode/QR, monitoring riwayat kehadiran pribadi, pengajuan dispensasi/izin sakit mandiri dengan bukti lampiran foto, serta interaksi dengan Asisten Cerdas Sekolah.

---

## 🤖 Fitur Kecerdasan Artifisial (AI) — Asisten Cerdas Hadir.io

Sistem ini diperkuat oleh asisten cerdas virtual bernama **Hadir.io** yang tertanam langsung pada antarmuka aplikasi. Asisten AI ini menggunakan arsitektur pemrosesan hybrid:

1. **Fast-Path SQL Extraction (Zero Halusinasi)**:  
   Untuk pertanyaan faktual seputar data sekolah (misalnya: jumlah kelas, total guru, rekap presensi hari ini, jadwal KBM), sistem mengambil data riil secara instan dari database lokal 20 tabel.
2. **NVIDIA NIM AI Processing (Nemotron-3.5-bolt-e4b)**:  
   Data teragregasi diformulasikan ke dalam prompt terstruktur dan diproses menggunakan model bahasa besar (**Large Language Model / LLM**) melalui NVIDIA NIM REST API untuk menghasilkan jawaban yang ramah, santun, ringkas (*to-the-point*), dan kontekstual.
3. **Dukungan 4 Bahasa (Multilingual)**:  
   Asisten Hadir.io mampu memahami dan merespons dalam:
   - 🇮🇩 **Bahasa Indonesia** (Resmi & Baku)
   - 🇬🇧 **English** (Bilingual)
   - 🇸🇺 **Basa Sunda** (Kearifan Lokal Jawa Barat)
   - 🇮🇩 **Basa Jawa** (Kearifan Lokal Cirebonan)
4. **RBAC Query Firewall**:  
   Lapisan keamanan controller yang memfilter dan menolak kueri yang berpotensi membocorkan data pribadi (seperti tanggal lahir, password, dan struktur kode internal).

---

## ⚡ Fitur-Fitur Unggulan Sistem

- 📷 **Scan Presensi Barcode & RFID Terintegrasi**: Menggunakan pustaka scanner kamera HTML5 interaktif berkecepatan tinggi dengan audio feedback real-time.
- 📱 **Otomasi Notifikasi Webhook Telegram**: Setiap kali siswa atau guru berhasil melakukan scan presensi, sistem secara otomatis mengirimkan notifikasi instan langsung ke kanal/grup Telegram sekolah melalui Bot `@HadirioBot`.
- 🔄 **Dual-Sync Presensi Wali Kelas**: Guru yang merangkap sebagai wali kelas otomatis mendapatkan sinkronisasi presensi ganda tanpa perlu melakukan scan berulang kali.
- 📊 **Rekapitulasi Presensi Presisi (Senin–Jumat)**: Algoritma rekapitulasi kehadiran mengecualikan hari libur akhir pekan (Sabtu & Minggu), serta mendukung ekspor satu klik ke format **Microsoft Excel (.xlsx)** dan **Cetak PDF**.
- 📋 **Pengajuan & Approval Perizinan Mandiri**: Siswa dapat mengunggah bukti surat dokter/izin yang langsung masuk ke dashboard wali kelas untuk diverifikasi.

---

## 📐 Pemodelan Sistem & Diagram UML

| Diagram | Berkas / Tautan | Deskripsi |
| :--- | :--- | :--- |
| **Use Case Diagram** | `Usecase_Diagram_Sistem_Absensi_SMK_NU_Wargabinangun.png` | Memetakan 14 use case fungsional untuk 5 aktor manusia di sisi kiri dan 2 aktor eksternal (Telegram & NVIDIA NIM). |
| **Activity Diagram** | `Activity_Diagram_Sistem_Absensi_SMK_NU_Wargabinangun.png` | Alur kerja swimlane bisnis presensi scan dan pemrosesan kueri Asisten AI Hadir.io. |
| **Class Diagram** | `Class_Diagram_Sistem_Absensi_SMK_NU_Wargabinangun.png` | Struktur kelas berorientasi objek MVC, Service, Model, dan Controller berprinsip SOLID. |
| **Sequence Diagram** | `Sequence_Diagram_Sistem_Absensi_SMK_NU_Wargabinangun.png` | Interaksi dinamis pertukaran pesan antara User, UI, Controller, MariaDB, Telegram API, dan NVIDIA NIM. |
| **Entity Relationship Diagram (ERD)** | [`ERD_Database_Absensi_SMK.png`](ERD_Database_Absensi_SMK.png) | Struktur relasional 20 tabel database `db_absensi_smk` ternormalisasi 3NF dengan integritas Foreign Key. |
| **Prototipe Desain UI/UX** | [`Desain_UI_UX_Figma_Sistem_Presensi_SMK.png`](Desain_UI_UX_Figma_Sistem_Presensi_SMK.png) | Visual mockup antarmuka mobile-first bertema Navy Blue & Sky Blue modern. |
| **Viewer Prototipe Web Interaktif** | [`desain_ui_ux_figma.html`](desain_ui_ux_figma.html) | File viewer interaktif simulasi alur antarmuka Figma di peramban web. |

---

## 🎨 Desain & Prototipe Interaktif UI/UX (Figma)

Antarmuka Sistem Presensi SMK NU Wargabinangun dirancang dengan pendekatan **Mobile-First & Clean Dashboard** menggunakan kombinasi warna profesional *Navy Blue* (`#0F172A`) dan *Sky Blue* (`#0284C7`). Desain ini memastikan tata letak tombol pemindai kamera responsif, kartu presensi mudah dibaca, dan navigasi ramah pengguna bagi seluruh tingkatan pengguna (Admin, Kepala Sekolah, Wali Kelas, Guru, dan Siswa).

- 🔗 **Tautan Berkas Desain Figma**: [https://www.figma.com/design/HJ3KQU3E6jYNRBvK1X99aU/Untitled?node-id=0-1](https://www.figma.com/design/HJ3KQU3E6jYNRBvK1X99aU/Untitled?node-id=0-1)  
  *(Akses terbuka publik untuk pengujian & evaluasi UI/UX interaktif)*
- 🌐 **Prototipe Web Interaktif**: Buka berkas [`desain_ui_ux_figma.html`](desain_ui_ux_figma.html) di browser untuk melihat interaksi mockup antarmuka.
- 🖼️ **Tinjauan Visual Desain UI/UX**:

<p align="center">
  <img src="Desain_UI_UX_Figma_Sistem_Presensi_SMK.png" alt="Desain UI/UX Figma Sistem Presensi SMK NU Wargabinangun" width="100%">
</p>

---

## 🛠️ Petunjuk Instalasi & Menjalankan Aplikasi

Ikuti langkah-langkah berikut untuk menjalankan sistem di lingkungan lokal komputer Anda:

### 1. Prasyarat Sistem (Prerequisites)
Pastikan komputer Anda sudah terinstal:
- **XAMPP** (dengan PHP versi 8.1 / 8.2 atau lebih baru dan MySQL/MariaDB)
- **Web Browser** modern (Google Chrome, Microsoft Edge, Mozilla Firefox)
- **Git** (opsional untuk cloning)

---

### 2. Unduh / Clone Repositori
Buka terminal (Git Bash, Command Prompt, atau PowerShell) di direktori `htdocs` XAMPP Anda:
```bash
cd c:/xampp/htdocs
git clone https://github.com/Wwinaa/proyek-rpl-absensiQR.git absen
cd absen
```
*(Atau ekstrak file zip proyek ke dalam folder `c:\xampp\htdocs\absen`)*

---

### 3. Konfigurasi Basis Data (Database Setup)
1. Nyalakan modul **Apache** dan **MySQL** pada XAMPP Control Panel.
2. Buka browser dan akses phpMyAdmin di `http://localhost/phpmyadmin`.
3. Buat database baru bernama: `db_absensi_smk`.
4. Klik tab **Import**, pilih berkas:
   - `database/schema.sql` (untuk membuat 20 tabel relasional).
   - `database/seed.sql` (untuk mengisi master data awal, akun, jurusan, kelas, dan jadwal).
5. Klik tombol **Go / Import** hingga seluruh tabel berhasil dimuat.

---

### 4. Konfigurasi Lingkungan (.env)
Salin berkas `.env.example` menjadi `.env`:
```bash
cp .env.example .env
```
Buka file `.env` dan pastikan kredensial database sesuai dengan konfigurasi XAMPP Anda:
```env
DB_HOST=127.0.0.1
DB_USER=root
DB_PASS=
DB_NAME=db_absensi_smk

APP_URL=http://localhost/absen
APP_NAME="Sistem Presensi SMK NU Wargabinangun"
APP_ENV=development

# Konfigurasi Opsional untuk Fitur AI & Telegram
NVIDIA_API_KEY=your_nvidia_api_key_here
NVIDIA_MODEL=nvidia/nemotron-3.5-bolt-e4b
TELEGRAM_CHAT_ID="your_chat_id"
TELEGRAM_BOT_TOKEN="your_bot_token"
```

---

### 5. Menjalankan Aplikasi di Browser
Buka browser favorit Anda dan akses alamat:
```
http://localhost/absen
```

---

## 🔑 Kredensial Login Default (Untuk Pengujian & Penilaian)

Sistem telah dilengkapi dengan akun demonstrasi untuk seluruh 5 tingkatan hak akses:

| Peran (Role) | Username | Password | Deskripsi Tugas |
| :--- | :--- | :--- | :--- |
| **Administrator** | `admin` | `admin123` | Hak akses penuh superuser |
| **Kepala Sekolah** | `kepsek` | `kepsek123` | Supervisi KBM & rekapitulasi eksekutif |
| **Wali Kelas** | `walikelas` | `walikelas123` | Monitoring kelas, approval izin, pembinaan |
| **Guru Pengajar** | `guru` | `guru123` | Jurnal mengajar & presensi mandiri |
| **Peserta Didik (Siswa)** | `siswa` | `siswa123` | Presensi scan, pantau riwayat, chat AI |

---

## 📂 Struktur Direktori Proyek

```text
absen/
├── app/
│   ├── config/          # Konfigurasi basis data & environment
│   ├── controllers/     # Controller logika bisnis (MVC)
│   ├── core/            # Router, base controller, database wrapper
│   ├── helpers/         # Helper autentikasi, URL, dan flash messages
│   ├── libraries/       # Pustaka eksternal (SimpleXLSXGen)
│   ├── models/          # Model data 20 tabel database
│   └── views/           # Template antarmuka responsif (Admin, Guru, Siswa, dll.)
├── database/
│   ├── schema.sql       # Skema DDL 20 tabel relasional
│   └── seed.sql         # Data awal (dummy/master data lengkap)
├── public/
│   ├── assets/          # CSS, JavaScript, icons, logo
│   └── uploads/         # Upload foto profil dan bukti surat izin
├── chat_api.php         # Endpoint Asisten AI Hadir.io (NVIDIA NIM)
├── index.php            # Front Controller & Application Entry Point
├── .env.example         # Template konfigurasi environment
├── .gitignore           # Aturan pengabaian file Git
└── README.md            # Dokumentasi lengkap proyek
```

---

## 📄 Dokumen Laporan Rekayasa Perangkat Lunak & Deliverables
Laporan proyek akhir lengkap (PRD/SRS, Pemodelan UML, Desain UI/UX Figma, dan Laporan Pengujian/CI-CD) tersedia dalam format:
- 📕 **Laporan Lengkap (PDF Tunggal)**: [Laporan_Proyek_Akhir_RPL_Sistem_Presensi_SMK_NU_Wargabinangun.pdf](https://github.com/Wwinaa/proyek-rpl-absensiQR/blob/main/Laporan_Proyek_Akhir_RPL_Sistem_Presensi_SMK_NU_Wargabinangun.pdf)
- 📥 **Unduh Langsung Berkas PDF**: [Download PDF](https://raw.githubusercontent.com/Wwinaa/proyek-rpl-absensiQR/main/Laporan_Proyek_Akhir_RPL_Sistem_Presensi_SMK_NU_Wargabinangun.pdf)
- 📝 **Dokumen Microsoft Word (DOCX)**: [Laporan_Proyek_Akhir_RPL_Sistem_Presensi_SMK_NU_Wargabinangun.docx](https://github.com/Wwinaa/proyek-rpl-absensiQR/blob/main/Laporan_Proyek_Akhir_RPL_Sistem_Presensi_SMK_NU_Wargabinangun.docx)
- 🎬 **Naskah & Panduan Demo**: [`Naskah_Video_Presentasi_RPL_Winarti.docx`](https://github.com/Wwinaa/proyek-rpl-absensiQR/blob/main/Naskah_Video_Presentasi_RPL_Winarti.docx)

---

*Dikembangkan dengan penuh dedikasi oleh **Winarti** (NIM: 41257168) untuk kemajuan digitalisasi SMK NU Wargabinangun Cirebon.*
