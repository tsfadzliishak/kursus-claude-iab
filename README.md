# Makmal Claude IAB

Bahan kursus **"Claude untuk Kerja Pengurusan Pendidikan"** — kursus 2 hari untuk pensyarah dan pegawai Institut Aminuddin Baki (IAB) Genting Highlands.

Repositori ini mengandungi laman web pembelajaran interaktif, fail data latihan, kod sistem Google Apps Script dan slaid kursus.

> Semua nama, angka dan data dalam repositori ini adalah **rekaan** untuk tujuan latihan sahaja.

---

## Laman web pembelajaran

Laman web satu halaman (`index.html`) yang mengandungi:

- **12 modul** dengan objektif, konsep, langkah aktiviti dan prompt yang boleh disalin satu klik
- **Penjejak kemajuan** — 39 aktiviti hands-on, disimpan dalam pelayar peserta
- **Pemasa aktiviti** untuk setiap sesi hands-on
- **Makmal Prompt** — pembina prompt kerangka PKTFCK
- **Kanvas Agen** — borang reka bentuk agen AI
- **Kuiz imbas kembali** 12 soalan dengan penjelasan serta-merta
- **Pusat fail** — semua fail latihan dijana dalam pelayar semasa muat turun

Laman ini tidak memerlukan pelayan, pangkalan data atau pemasangan. Buka `index.html` dalam mana-mana pelayar.

### Hoskan melalui GitHub Pages

1. Buka **Settings → Pages** dalam repositori ini.
2. **Source:** Deploy from a branch. **Branch:** `main`, folder `/ (root)`. Klik **Save**.
3. Tunggu 1–2 minit. Laman akan tersedia di `https://<nama-pengguna>.github.io/<nama-repo>/`.

Fail `.nojekyll` sudah disertakan supaya GitHub Pages menghidangkan fail apa adanya.

---

## Struktur repositori

| Folder | Kandungan |
| --- | --- |
| `index.html` | Laman web pembelajaran (satu fail, lengkap) |
| `fail/` | 9 fail data latihan dalam format asal + 2 pakej ZIP |
| `kod/` | Kod Google Apps Script untuk 2 sistem, agen AI dan skill |
| `slaid/` | Dek PowerPoint 52 slaid dengan nota penceramah |
| `sumber/` | Kod sumber laman web dan skrip penjana data |

---

## Modul kursus

### Hari 1

| Kod | Modul | Tempoh |
| --- | --- | --- |
| M1 | Claude Cowork | 1 jam |
| M1A | Aplikasi Claude Desktop (Windows & Mac) | 30 min |
| M2 | Advanced Prompt Engineering | 1 jam |
| M3 | Multi-file Intelligence | 1 jam |
| M4 | Document Pipeline | 2 jam |
| MCP | MCP & Penyambung Microsoft 365 / Google Workspace | 45 min |

### Hari 2

| Kod | Modul | Tempoh |
| --- | --- | --- |
| M5 | Task Orchestration | 2 jam |
| M6 | Data Intelligence | 2 jam |
| M7 | Performance Reporting | 1 jam |
| M8 | Integrated Management Workflow | 2 jam |

### Modul lanjutan (pilihan)

| Kod | Modul | Tempoh |
| --- | --- | --- |
| T1 | Bina Sistem dengan Claude: Google Sheets + Apps Script | 2 jam |
| T2 | Bina Agentic AI dengan Claude | 2 jam |

Semua modul berpaut pada satu senario: peserta bertindak sebagai urus setia **Program Kepimpinan Pengetua Baharu Siri 3/2026** (5 hari, 40 peserta, 7–11 Disember 2026). Output setiap modul menjadi input modul seterusnya.

---

## Fail data latihan

| Fail | Kandungan | Digunakan dalam |
| --- | --- | --- |
| `01_Profil_Program_PKPB3.docx` | Latar belakang, 5 objektif, sasaran KPI | M1 · M3 · M4 · MCP |
| `02_Senarai_Penceramah.xlsx` | 5 penceramah dengan tarikh tersedia | M3 · M4 · M8 |
| `03_Komen_Peserta.txt` | 20 komen terbuka peserta | M2 · M6 |
| `04_Jadual_Program_Draf.xlsx` | Jadual 5 hari, **mengandungi 4 kesilapan sengaja** | M3 |
| `05_Nota_Mesyuarat_JK.txt` | Nota kasar mesyuarat jawatankuasa | M4 · MCP · T2 |
| `06_Senarai_Peserta.xlsx` | 40 peserta: negeri, jantina, jenis sekolah | M5 |
| `07_Penilaian_PKPB3_mentah.xlsx` | Respons penilaian, **data belum dibersihkan** | M6 |
| `08_Ujian_Pra_Pasca_PKPB3.xlsx` | Skor pra dan pasca (/50) bagi 40 peserta | M6 · M7 |
| `09_Data_Operasi_PKPB3.xlsx` | Kehadiran, perbelanjaan, KPI Siri 2/2026 | M7 · M8 |

### Kesilapan sengaja (kunci jawapan jurulatih)

<details>
<summary>Klik untuk melihat — jangan edarkan kepada peserta sebelum aktiviti</summary>

**`04_Jadual_Program_Draf.xlsx` (Aktiviti 3.1)**

1. Sesi 11 *Pengurusan Aset dan Perolehan*: penceramah En. Rosli bin Hamid tiada dalam Senarai Penceramah.
2. Sesi 5 *Coaching dan Mentoring* pada 8 Dis: Dr. Siti Hajar hanya tersedia 9 dan 10 Dis.
3. Sesi 7 (9.00 pagi – 12.30 tgh) dan Sesi 8 (11.00 pagi – 12.30 tgh) bertindih di Bilik Kuliah 2 pada 9 Dis.
4. Objektif O4 (konflik dan krisis) tiada sesi berkaitan.

**`07_Penilaian_PKPB3_mentah.xlsx` (Aktiviti 6.1)**

- Pendua: P10 dan P22 (setiap satu muncul 2 kali)
- Nilai luar julat: P07 S07 = 0, P18 S11 = 6, P29 U5 = 55
- 5 sel kosong, nama negeri tidak seragam (Johor / JOHOR / Jhr), format tarikh bercampur

**Dapatan utama**

- Sesi min terendah: S08 Data dan Pelaporan 3.44, S07 Pengurusan Kewangan 3.46, S11 Aset dan Perolehan 3.50
- Pra-pasca: min pra 26.8, min pasca 35.4, peningkatan 8.6 markah; 60% skor pasca ≥ 35
- Status KPI: kehadiran 96.5% (capai), kepuasan 4.21 (sempadan), % pasca ≥ 35 = 60% (tidak capai), perbelanjaan 97.6% (capai)

</details>

---

## Kod sistem

### `kod/1_sistem_pengurusan_program/` — Modul Tambahan 1 (Aktiviti B.3)

Aplikasi web pendaftaran peserta, borang penilaian dan papan pemuka min skor.

1. Cipta Google Sheet baharu `DB_Sistem_PKPB3`, buka **Extensions → Apps Script**.
2. Tampal `Code.gs`. Tambah fail HTML bernama **`Index`** (tanpa `.html`), tampal `Index.html`.
3. Jalankan fungsi `setupSheets` dan beri kebenaran.
4. **Deploy → New deployment → Web app**. *Execute as:* Me. *Who has access:* dalam organisasi.

### `kod/2_sistem_pemantauan_tugasan/` — Aktiviti 5.1

Papan pemuka pemantauan tugasan live dengan auto-refresh 30 saat, tapisan dan jejak audit.

1. Cipta Google Sheet `DB_Pemantauan_PKPB3`, tampal `Code.gs` dan fail HTML `Index`.
2. Jalankan `setupSheet`.
3. Import `WBS_Contoh_PKPB3.csv` ke helaian **Tugasan** (File → Import → Replace current sheet), kemudian jalankan `setupSheet` sekali lagi.
4. **Deploy → New deployment → Web app**, kongsi pautan kepada pegawai.

### `kod/3_agen_ai/` — Modul Tambahan 2 (Tahap 4)

Agen autonomi yang menyusuli tugasan lewat setiap pagi menggunakan Claude API dengan tool use.

1. Tambah `Agen.gs` ke projek Apps Script **Sistem Pemantauan Tugasan**.
2. **Project Settings → Script Properties**, tambah `ANTHROPIC_API_KEY` (kunci dari Claude Console).
3. Jalankan `setupAgen`, isi e-mel dalam helaian **Pegawai**.
4. Muat semula Sheet, guna menu **Agen AI**.

Had keselamatan agen: hanya mencipta draf e-mel (tidak menghantar), tidak mengubah status tugasan, berhenti selepas 8 pusingan, dan merekod setiap tindakan dalam helaian Log Agen.

### `kod/4_skill/` — Modul Tambahan 2 (Tahap 2)

Skill contoh `minit-mesyuarat-iab`. Zip folder `minit-mesyuarat-iab` dan muat naik di **Customize → Skills**.

---

## Membina semula laman web

```bash
cd sumber
python3 build.py        # menjana ../index.html daripada index.src.html + data
```

`build.py` membaca fail dalam `fail/`, menukarnya kepada data JSON berstruktur dan membenamkannya ke dalam laman. Laman kemudian menjana semula fail Word, Excel dan ZIP dalam pelayar semasa muat turun.

`jana_fail.py` menjana semula kesemua 9 fail data latihan dari awal.

Keperluan: Python 3 dengan `openpyxl` dan `python-docx`.

---

## Nota kepada jurulatih

- Semak ciri Claude terkini di [support.claude.com](https://support.claude.com) seminggu sebelum kursus. Ciri dan nama menu mungkin berubah.
- Penyambung Microsoft 365 dan Google Workspace memerlukan kelulusan pentadbir ICT. Mulakan permohonan 2 minggu awal.
- Uji pemasangan Claude Desktop dan Apps Script pada komputer peserta sebelum kursus.
- Jangan muat naik dokumen Terhad, Sulit atau Rahsia semasa latihan.
- Sasaran KPI dan jadual dalam bahan ini adalah contoh. Gantikan dengan sasaran rasmi IAB.

## Lesen

Bahan latihan dalaman. Data rekaan sepenuhnya.
