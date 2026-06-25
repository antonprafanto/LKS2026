# SOP JURI — Robot Bergerak Otonom LKS Nasional 2026

**Versi:** 1.3  
**Acuan:** Kisi-kisi LKS 2026 + Deskripsi Teknis v1.0 (WorldSkills Skill 23)  
**Tema proyek uji:** Service Robot at Library

---

## 1. Ruang Lingkup

SOP ini mengatur peran juri dalam penilaian **5 modul** (A–E) selama 3 hari lomba + H-1. Berlaku untuk semua juri bidang Autonomous Mobile Robotics.

### Dokumen & file wajib juri

| Prioritas | Dokumen / file |
|-----------|----------------|
| Wajib | **Marking Scheme Excel CIS** (sumber kebenaran poin detail) |
| Wajib | `templates/LKS2026-Penilaian-Juri.xlsx` (1 file per tim, Save As) |
| Wajib | Work Instructions / Test Project hari H |
| Wajib | Petunjuk Teknis Umum LKS 2026 |
| Acuan | Kisi-kisi & Deskripsi Teknis — [Puspresnas](https://smk.pusatprestasinasional.kemdikdasmen.go.id) |
| Briefing | `docs/Briefing-Juri-LKS2026-Robot-Otonom.pptx` |
| Latihan | `templates/LKS2026-Penilaian-Juri-CONTOH.xlsx` + `docs/SKENARIO-SOAL-CONTOH.md` |
| Google Sheets | `docs/PANDUAN-GOOGLE-SHEETS.md` (stopwatch pakai Apps Script) |

> **Jangan pakai file CSV** di folder `templates/` — gunakan workbook Excel saja.

### Paket share ke juri lain

1. `LKS2026-Penilaian-Juri.xlsx`  
2. `SOP-JURI.md` (dokumen ini)  
3. `Briefing-Juri-LKS2026-Robot-Otonom.pptx`  

**Panduan online (GitHub):** [github.com/antonprafanto/LKS2026](https://github.com/antonprafanto/LKS2026) — setiap sheet Excel punya link bantuan; daftar lengkap di sheet **Panduan**.

---

## 2. Prinsip Penilaian

### 2.1 Dua metode

| Metode | Kode | Skala | Digunakan untuk |
|--------|------|-------|-----------------|
| **Judgement** | J | 0–3 | Kualitas proses, kerapian, kolaborasi, dokumentasi |
| **Measurement** | M | 0 atau 1 | Akurasi gerak, sukses misi, penempatan kubus |

### 2.2 Skala Judgement (0–3)

| Skor | Kisi-kisi | Deskripsi Teknis |
|------|-----------|------------------|
| 0 | Sangat kurang / tidak ada | Di bawah standar industri / tugas tidak dilakukan |
| 1 | Cukup | Memenuhi standar industri |
| 2 | Baik | Melampaui standar industri |
| 3 | Sangat baik | Luar biasa, setara ekspektasi industri terkini |

### 2.3 Prosedur 3 juri (Judgement)

1. Setiap aspek Judgement dinilai **3 juri** secara independen.
2. Skor final = **rata-rata** tiga juri.
3. Jika selisih skor antar juri **≥ 2 poin** pada aspek yang sama → **penilaian diulang** hingga selisih < 2.
4. Semua skor dimasukkan ke **CIS** (skala 0–100 → konversi skala 700).

### 2.4 Measurement

- **1** = kriteria terpenuhi  
- **0** = kriteria tidak terpenuhi  
- Tidak ada nilai tengah.

---

## 3. Peran Juri per Fase

### 3.1 H-30 (±30 hari sebelum lomba)

- Technical Meeting 1 — pahami Marking Scheme CIS
- **Klarifikasi** ke Chief Expert: ukuran arena, tegangan baterai, jadwal Modul D, agregasi skor Modul E

### 3.2 H-1

| Waktu | Tugas juri |
|-------|------------|
| Technical Meeting 2 | Sampaikan prosedur SIAP/START/RETRY; jelaskan 30% rule |
| Familiarisasi (maks. 2 jam) | Pastikan 4 set arena lengkap; verifikasi aksesoris |

**Checklist arena per set:**

- [ ] Arena (multiplex 18 mm, permukaan putih matte)
- [ ] 3 rak + marker warna di atas slot
- [ ] 9 kubus (3 warna × 3 bentuk: ○ □ △)
- [ ] Kubus stand (80×80×100 mm)
- [ ] 10 obstacle
- [ ] Reference line (isolasi 18×300 mm)
- [ ] 1 gate (profil aluminium 2020)
- [ ] Penerangan: tiang + lampu 100 W × 2

### 3.3 Hari 1 (H1)

| Waktu | Modul | Tugas juri |
|-------|-------|------------|
| 08.00–08.30 | — | Briefing & test project |
| 08.30–12.00 | C | Awasi perakitan L-Channel + lift (batas **2 jam**); verifikasi K3 |
| Sepanjang H1 | A, B | Observasi 5S/kolaborasi; terima & nilai jurnal |
| 13.00–16.00 | D | Pra-uji gerakan dasar (penilaian resmi di **H2**) |

**Verifikasi Modul C sebelum robot dioperasikan:**

- [ ] L-Channel terpasang sebagai penyangga utama sistem angkat
- [ ] Motor kompatibel JGA-25 / PG36 (jika dipakai)
- [ ] Tidak ada kabel terbuka / terjepit
- [ ] Tombol STOP merah berfungsi
- [ ] Rewiring rapi & aman

### 3.4 Hari 2 (H2)

| Waktu | Modul | Tugas juri |
|-------|-------|------------|
| 07.30–08.30 | — | Briefing & test project |
| Pagi–siang | D | Penilaian resmi **15 menit/tim** |
| 08.30–15.00 | E | Supervisi **trial** otonom |
| 15.00–17.30 | E | Marking trial |

### 3.5 Hari 3 (H3)

| Waktu | Modul | Tugas juri |
|-------|-------|------------|
| 07.30–08.30 | — | Briefing + variasi soal (**30% rule**) |
| 08.30–15.00 | E | Supervisi **final** otonom |
| 15.00–17.30 | E | Marking final |
| Sepanjang H3 | A | Observasi berkelanjutan |

---

## 4. Cara Pakai Excel Penilaian

1. **Save As** — duplikat `LKS2026-Penilaian-Juri.xlsx` per tim.
2. Isi header (Tim No, nama, juri) di setiap sheet.
3. **Undian Kubus** — isi warna marker (baris 24), lalu ID baris 28–36 + Rak/Stand + Slot; kolom lain otomatis.
4. **Log Run Otonom** — checklist SIAP/START + kronologi; skor tarik dari Modul E.
5. **Modul E** — kolom B–G otomatis; juri isi **Hasil Run (1/0)** di kolom H.
6. **Rekapitulasi** — total otomatis; Chief Expert input ke CIS.

> **Excel 2010+** (Microsoft Excel desktop). Tidak memakai fungsi `TEXTJOIN`.  
> **Latihan:** buka `LKS2026-Penilaian-Juri-CONTOH.xlsx` + [contoh Modul E di GitHub](https://github.com/antonprafanto/LKS2026/blob/main/docs/CONTOH-EXCEL-MODUL-E.md).

| Sheet Excel | Modul | Panduan GitHub |
|-------------|-------|----------------|
| [Modul A](https://github.com/antonprafanto/LKS2026/blob/main/docs/SOP-JURI.md#5-sop-penilaian-modul-a--organisasi--manajemen-kerja-8-poin-j) | Organisasi & manajemen kerja | §5 |
| [Modul B](https://github.com/antonprafanto/LKS2026/blob/main/docs/SOP-JURI.md#6-sop-penilaian-modul-b--jurnal-teknis-10-poin-j-8--m-2) | Jurnal teknis | §6 |
| [Modul C](https://github.com/antonprafanto/LKS2026/blob/main/docs/SOP-JURI.md#7-sop-penilaian-modul-c--perakitan-robot-10-poin-j) | Perakitan robot | §7 |
| [Modul D](https://github.com/antonprafanto/LKS2026/blob/main/docs/SOP-JURI.md#8-sop-penilaian-modul-d--gerakan-dasar-12-poin-m) | Gerakan dasar | §8 |
| [Modul E](PANDUAN-MODUL-E.md) | Performa otonom | Panduan lengkap |
| [Undian Kubus](SOP-JURI.md#10-sop-pengacakan-kubus) | Pengacakan posisi | §10 |
| [Log Run Otonom](PANDUAN-LOG-RUN.md) | Log per run | Panduan lengkap |
| [Rekapitulasi](https://github.com/antonprafanto/LKS2026/blob/main/docs/SOP-JURI.md#12-rekapitulasi--cis) | Total A–E | §12 |

| Kolom Modul E | Arti | Siapa isi |
|---------------|------|-----------|
| Posisi Awal | Lokasi kubus setelah undian | Otomatis |
| Slot Target | Slot rak yang marker-nya cocok warna kubus | Otomatis |
| OK Awal (1/0) | Sudah benar sebelum run dimulai | Otomatis (referensi) |
| **Hasil Run (1/0)** | Sukses setelah run otonom | **Juri** |

---

## 5. SOP Penilaian Modul A — Organisasi & Manajemen Kerja (8 poin, J)

**Hari:** H1, H2, H3  
**Sheet:** `Modul A`  
**Panduan:** [Modul A di GitHub](https://github.com/antonprafanto/LKS2026/blob/main/docs/SOP-JURI.md#5-sop-penilaian-modul-a--organisasi--manajemen-kerja-8-poin-j)

| Sub | Kriteria | Max |
|-----|----------|-----|
| A1 | Efisiensi penjadwalan kerja | 2.0 |
| A2 | Kolaborasi tim & manajemen tugas | 2.0 |
| A3 | Organisasi & kerapian area kerja (5S) | 2.0 |
| A4 | Restorasi area kerja | 2.0 |

### Rubrik cepat A2 — Kolaborasi Tim

| 0 | 1 | 2 | 3 |
|---|---|---|---|
| Tidak ada pembagian kerja; dominasi/passive; komunikasi buruk | Peran jelas (mekanik/programmer); komunikasi aktif | Koordinasi sinkron; solusi cepat; manajemen waktu efisien | Kepemimpinan profesional; feedback konstruktif; krisis ditangani tenang |

### Rubrik cepat A3 — Kerapian Area Kerja

| 0 | 1 | 2 | 3 |
|---|---|---|---|
| Sangat berantakan; alat berserakan | Cukup rapi; alat dikembalikan | Sangat terorganisir; kabel rapi; efisiensi ruang | Standar lab industri; 5S konsisten; restorasi instan |

---

## 6. SOP Penilaian Modul B — Jurnal Teknis (10 poin: J 8 + M 2)

**Hari:** H1  
**Sheet:** `Modul B`  
**Panduan:** [Modul B di GitHub](https://github.com/antonprafanto/LKS2026/blob/main/docs/SOP-JURI.md#6-sop-penilaian-modul-b--jurnal-teknis-10-poin-j-8--m-2)  
**Batas:** Maks. 4 lembar isi (8 halaman A4), tidak termasuk sampul

### Sub-kriteria Judgement (sheet Modul B)

| Kode | Aspek | Max |
|------|-------|-----|
| B-J1 | Kualitas fabrikasi rangka & stabilitas | 1.5 |
| B-J2 | Manajemen pengkabelan | 1.5 |
| B-J3 | Integrasi sistem penggerak | 1.5 |
| B-J4 | Kalibrasi sensor & aktuator | 1.5 |
| B-J5 | Optimasi software-hardware | 1.5 |
| B-J6 | Kelengkapan dokumentasi | 0.5 |

### Checklist kelengkapan jurnal

- [ ] Analisis tugas & strategi
- [ ] Desain mekanik (CAD/sketsa + dimensi)
- [ ] Diagram blok sistem & BOM
- [ ] Dokumentasi fabrikasi & wiring
- [ ] Penempatan sensor (justifikasi)
- [ ] Arsitektur software, algoritma navigasi & visi
- [ ] Flowchart, data kalibrasi, log troubleshooting

### B-M1 — Ketepatan waktu (Measurement, 2 poin)

| Kondisi | Isi di Excel (sel G4) |
|---------|------------------------|
| Tepat waktu | 1 |
| Terlambat / tidak terkumpul | 0 |

> Poin Measurement sisanya (jika ada) mengacu Marking Scheme CIS.

---

## 7. SOP Penilaian Modul C — Perakitan Robot (10 poin, J)

**Hari:** H1  
**Sheet:** `Modul C`  
**Panduan:** [Modul C di GitHub](https://github.com/antonprafanto/LKS2026/blob/main/docs/SOP-JURI.md#7-sop-penilaian-modul-c--perakitan-robot-10-poin-j)

### Alur inspeksi

1. **Sebelum 2 jam:** Catat progress perakitan L-Channel.
2. **Setelah 2 jam:** Perakitan mekanik + rewiring harus selesai.
3. **Inspeksi akhir:** L-Channel + lift + kabel + aktuator.

| Aspek | Indikator lulus |
|-------|-----------------|
| L-Channel | Penyangga utama sistem angkat |
| Mekanik | Stabil, tidak goyang berlebihan |
| Wiring | Rapi, tidak terbuka/terjepit |
| K3 | Tombol STOP merah accessible & berfungsi |

CAD L-Channel: [Google Drive resmi](https://drive.google.com/drive/folders/1xpbziLSLvmT-wKWaHIRK2A7eIz5fIM9P?usp=share_link)

---

## 8. SOP Penilaian Modul D — Gerakan Dasar (12 poin, M)

**Hari:** H2 (resmi)  
**Waktu:** 15 menit/tim  
**Sheet:** `Modul D`  
**Panduan:** [Modul D di GitHub](https://github.com/antonprafanto/LKS2026/blob/main/docs/SOP-JURI.md#8-sop-penilaian-modul-d--gerakan-dasar-12-poin-m)

### Prosedur sesi

1. Panggil tim ke zona uji.
2. Pastikan monitor laptop terlihat juri (item D4–D8).
3. Baca item satu per satu; catat **1** atau **0** di Excel.
4. Tidak ada retry item (kecuali Chief Expert menetapkan lain).

### Tabel item & bobot mentah

| ID | Tugas | Bukti | Bobot |
|----|-------|-------|-------|
| D1a–D1d | Gerak maju/mundur/kiri/kanan 200–300 mm | Observasi | 0,50 each |
| D2a | Maju, stop 50–100 mm dari dinding | Observasi | 0,30 |
| D2b | Kanan, stop 50–100 mm dari dinding | Observasi | 0,60 |
| D3 | Ikuti garis U | Line follower | 1,00 |
| D4 | Deteksi merah → teks "merah" | Monitor | 1,00 |
| D5 | Deteksi hijau → teks "hijau" | Monitor | 1,00 |
| D6 | Deteksi biru → teks "biru" | Monitor | 1,00 |
| D7 | Deteksi lingkaran merah → "O-merah" | Monitor | 1,00 |
| D8 | Deteksi lingkaran hijau → "O-hijau" | Monitor | 1,00 |
| D9 | Pick kubus dari rak otonom | Observasi arm | 2,00 |
| D10 | Place kubus ke standbox | Observasi arm | 2,00 |
| | **Jumlah bobot mentah** | | **12,9** |

> Excel **diskalakan ke 12 poin** (`SUM × 12/12,9`). Bentuk persegi/segitiga tidak diuji eksplisit di Modul D.

---

## 9. SOP Penilaian Modul E — Performansi Otonom (60 poin, M)

**Hari:** H2 (trial) + H3 (final)  
**Sheet:** `Modul E`, `Undian Kubus`, `Log Run Otonom`  
**Panduan:** [Modul E §9.0](SOP-JURI.md#90-panduan-lengkap-modul-e-untuk-juri) · **[Panduan lengkap Modul E](PANDUAN-MODUL-E.md)** · **[Panduan Log Run](PANDUAN-LOG-RUN.md)** · [Contoh Excel](CONTOH-EXCEL-MODUL-E.md)

> **Untuk juri baru:** baca [`PANDUAN-MODUL-E.md`](PANDUAN-MODUL-E.md) dan [`PANDUAN-LOG-RUN.md`](PANDUAN-LOG-RUN.md) — penjelasan runut lengkap. File latihan: `LKS2026-Penilaian-Juri-CONTOH.xlsx`.

### 9.0 Panduan lengkap Modul E untuk juri

#### Inti (1 kalimat)

**Juri mengisi sheet Undian Kubus dulu → Modul E terisi otomatis → juri cukup isi kolom Hasil Run (1/0).**

#### Tiga sheet yang dipakai

```
┌──────────────────────────┐
│  1. UNDIAN KUBUS         │  ← juri ISI di sini (setup per run)
│     marker, kubus, obs   │
└────────────┬─────────────┘
             │ rumus Excel
             ▼
┌──────────────────────────┐
│  2. MODUL E              │  ← juri CEK + isi Hasil Run (kolom H)
│     penilaian 60 poin    │
└────────────┬─────────────┘
             │
             ▼
┌──────────────────────────┐
│  3. LOG RUN OTONOM       │  ← checklist SIAP/START + kronologi
└──────────────────────────┘
```

#### Urutan waktu satu run (contoh)

| Waktu | Apa yang terjadi | Sheet / aksi juri |
|-------|------------------|-------------------|
| −15 mnt | Briefing layout arena ke tim | Jelaskan START, rak, gate, obstacle |
| −5 mnt | Tim programming selesai → **SIAP** | Pastikan laptop tidak disentuh lagi |
| 0 mnt | Juri **acak** posisi kubus | **Isi Undian Kubus** (baris 28–36 + obstacle) |
| +1 mnt | Sinyal **START** → tombol di robot | Centang checklist di Log Run |
| +1–15 mnt | Robot berjalan **otonom** | Pantau; jangan sentuh robot |
| Selesai | Catat hasil per kubus | **Modul E kolom H** + Log Run |

#### A. Isi Undian Kubus (SEBELUM / SETELAH acak, per run)

**① Baris 24 (kuning) — warna marker per slot**

| | Kiri | Tengah | Kanan |
|---|------|--------|-------|
| Rak 1 | Merah/Hijau/Biru | … | … |
| Rak 2 | … | … | … |
| Rak 3 | … | … | … |

Kode di Excel: **R1–R3** = Rak 1–3, **S1–S3** = slot kiri/tengah/kanan.  
**Marker** = warna label di atas slot — kubus sukses jika ditaruh di slot yang marker-nya **sama warnanya** dengan kubus.

Contoh baris 24: `Merah | Hijau | Biru | Biru | Merah | Hijau | Hijau | Biru | Merah`

**② Baris 28–36 — posisi kubus setelah diacak**

| Kolom | Isi | Contoh |
|-------|-----|--------|
| **A** | ID kubus | `K-M2` |
| **E** | Rak/Stand | `Rak 1`, `Stand A` (dropdown) |
| **F** | Slot | `kiri`, `tengah`, `kanan` (dropdown) |

Kolom B, C, D, H, I, J di Undian **otomatis** — jangan ketik manual.

**③ Obstacle (Obs 1–10)** — catat lokasi singkat obstacle yang dipakai

Contoh: `kiri GATE` | `tengah arena` | `dekat Rak 2` — kolom tidak dipakai dikosongkan.

#### B. Sheet Modul E — kolom mana untuk apa?

| Kolom | Nama | Siapa isi? | Fungsi |
|-------|------|------------|--------|
| B | ID Kubus | Otomatis | Dari Undian |
| C–D | Warna, Bentuk | Otomatis | Dari katalog |
| E | Posisi Awal | Otomatis | Rak/Stand + Slot dari Undian |
| F | Slot Target | Otomatis | Daftar slot valid untuk warna kubus |
| G | OK Awal (1/0) | Otomatis | Referensi: sudah benar **sebelum** run? |
| **H** | **Hasil Run (1/0)** | **JURI** | **Skor resmi** setelah run |
| I–J | Bobot, Skor | Otomatis | 60÷9 per kubus, total max 60 |

**Yang juri ketik manual di Modul E: hanya kolom H.**

#### C. Cara isi Hasil Run (1/0)

Tanya per kubus setelah run selesai:

> Apakah kubus berhasil di slot rak yang **marker warnanya cocok** dengan warna kubus, dan **stabil**?

| Isi | Artinya |
|-----|---------|
| **1** | Sukses |
| **0** | Gagal (salah slot, jatuh, tidak terpenuhi, dll.) |

**OK Awal (G)** hanya membantu melihat posisi awal — **bukan skor**. Yang menghitung poin = **Hasil Run (H)**.

#### D. Contoh penilaian (Tim 05, trial)

| ID | Posisi Awal | OK Awal | Hasil Run | Keterangan |
|----|-------------|---------|-----------|------------|
| K-M1 | Rak 1 — kiri | 1 | **1** | Sudah benar, tetap stabil |
| K-M2 | Stand A — depan | 0 | **1** | Robot berhasil pindahkan |
| K-H2 | Stand B — depan | 0 | **0** | Jatuh saat place |
| … | … | … | … | 9 baris total |

Skor run ≈ jumlah Hasil Run × (60÷9). Contoh: 8 sukses ≈ **53,33** / 60.

#### E. Jika Modul E menampilkan petunjuk (bukan angka/nama)

| Tampilan | Penyebab | Solusi |
|----------|----------|--------|
| `→ isi Rak/Stand & Slot` | Kolom E/F Undian kosong | Lengkapi Undian |
| `→ isi warna marker baris 24` | Baris 24 kosong | Isi warna marker |
| Data normal (posisi, slot target) | Undian lengkap | Lanjut isi **Hasil Run** |

#### F. Checklist juri per run Modul E

- [ ] Baris 24 marker terisi
- [ ] Posisi 9 kubus terisi (A, E, F) setelah acak
- [ ] Obstacle dicatat (jika dipakai)
- [ ] Briefing layout → SIAP → acak → START sesuai SOP
- [ ] Modul E kolom H (Hasil Run) terisi semua baris
- [ ] Log Run Otonom checklist + kronologi terisi
- [ ] Tanda tangan juri & peserta (Undian / Log Run)

---

### 9.1 Alur run operasional (WAJIB)

```
1. Briefing layout arena
2. Peserta programming (BOLEH sentuh laptop)
3. Peserta nyatakan SIAP
4. STOP sentuh laptop
5. Juri ACAK ULANG posisi kubus → isi sheet Undian Kubus
6. Sinyal START
7. Peserta tekan tombol START di robot
8. Robot berjalan OTOMONOM
9. Isi **Hasil Run (1/0)** di Modul E kolom H + catat di Log Run
```

### 9.2 Prosedur RETRY

| Aturan | Detail |
|--------|--------|
| Permintaan | Peserta minta ke juri (kebijakan CE) |
| Ubah program | Diizinkan sebelum run ulang |
| Posisi kubus | TIDAK dikembalikan ke kondisi awal |
| Kubus di robot | Harus dikeluarkan ke luar arena |
| Setelah SIAP lagi | Posisi kubus diacak ulang |

### 9.3 Kriteria sukses penempatan kubus

Kubus **berhasil** (isi **Hasil Run = 1**) jika:

1. Di slot rak dengan **marker warna yang cocok** dengan warna kubus, dan  
2. Kubus stabil (tidak jatuh / setengah keluar), dan  
3. Penempatan selama run otonom (tanpa sentuh robot setelah START)

**OK Awal vs Hasil Run:**

| Kolom Excel | Kapan | Untuk apa |
|-------------|-------|-----------|
| OK Awal (1/0) | Sebelum START | Referensi — apakah kubus sudah benar di rak |
| **Hasil Run (1/0)** | Setelah run | **Skor resmi** — keputusan juri |

**Klarifikasi CE wajib:** Apakah kubus **OK Awal = 1** otomatis mendapat **Hasil Run = 1** tanpa dipindah robot?

### 9.4 Skema poin (template Excel)

| Item | Bobot |
|------|-------|
| Per kubus berhasil | **60 ÷ 9** poin (rumus `=60/9` di Excel) |
| **TOTAL** | **60** |

> Jika Marking Scheme CIS berbeda, **CIS yang mengikat**.

### 9.5 Agregasi skor H2 vs H3

- H2: trial Modul E (marking 15.00–17.30)
- H3: final Modul E + variasi 30%
- **Keputusan CE:** skor trial/final dijumlahkan, diratakan, atau hanya final yang dihitung

### 9.6 Pelanggaran

| Pelanggaran | Tindakan juri |
|-------------|---------------|
| Sentuh laptop setelah SIAP | Run batal / skor 0 run (konfirmasi CE) |
| Sentuh robot setelah START | Run batal / skor 0 run |
| Tidak start via tombol robot | Run batal |
| Place manual oleh peserta | Item/run batal |

---

## 10. SOP Pengacakan Kubus

**Sheet:** `Undian Kubus`  
**Panduan:** [Undian Kubus di GitHub](https://github.com/antonprafanto/LKS2026/blob/main/docs/SOP-JURI.md#10-sop-pengacakan-kubus)

1. Siapkan 9 kubus (3 merah, 3 hijau, 3 biru × ○ □ △).
2. Isi **baris marker** (baris 24, baris kuning): warna Merah/Hijau/Biru per slot rak.
3. Acak posisi di rak + stand — **beberapa sengaja salah**.
4. Isi **baris 28–36**: kolom **ID Kubus**, **Rak/Stand**, dan **Slot** (kiri/tengah/kanan).
5. Kolom Lokasi Awal, Sudah Benar, Target Rak/Slot otomatis di Undian.
6. Sheet **Modul E** menampilkan Posisi Awal, Slot Target, OK Awal — otomatis dari Undian.
7. Catat posisi **obstacle** (Obs 1–10) jika dipakai di run tersebut.
8. Setelah peserta **SIAP** → acak ulang untuk run tersebut.

### Variasi 30% rule (H3)

- Urutan pengiriman  
- Posisi obstacle  
- Konfigurasi marker warna di rak  
- Posisi start  

---

## 11. SOP K3 untuk Juri

- [ ] Larangan air mineral terbuka dekat peralatan listrik
- [ ] Verifikasi tombol STOP setiap tim sebelum run
- [ ] Reset obstacle yang jatuh sebelum tim berikutnya
- [ ] Laporkan kabel terbuka / baterai aus ke panitia
- [ ] Antisipasi **perubahan intensitas cahaya** venue (pengaruh kamera/Lidar)

---

## 12. Rekapitulasi & CIS

**Panduan:** [Rekapitulasi di GitHub](https://github.com/antonprafanto/LKS2026/blob/main/docs/SOP-JURI.md#12-rekapitulasi--cis)

1. Pastikan semua sheet Excel tim terisi.
2. Periksa sheet **Rekapitulasi** (total otomatis dari Modul A–E).
3. Chief Expert review outlier (selisih juri ≥ 2) & pelanggaran.
4. Input skor resmi ke **CIS** → konversi skala 700.
5. Arsipkan file Excel per tim.

| Peringkat | Kriteria |
|-----------|----------|
| Juara 1–3 | Skor CIS tertinggi |
| Medallion of Excellence | Skor > 700 |

---

## 13. Klarifikasi Wajib ke Chief Expert

| Isu | Catatan |
|-----|---------|
| Ukuran arena | 4040×2040 mm (Kisi-kisi) vs 2400×4800 mm (Daftar Bahan) |
| Tegangan maks. baterai | Referensi @12V; hal. 19 Teknis terpotong |
| Bobot per kubus Modul E | Excel 60/9 vs CIS |
| Kubus sudah benar di awal | Dapat poin atau tidak? |
| Agregasi skor trial + final | H2 + H3 cara hitung |
| Maksimum RETRY per tim | Tidak disebut di dokumen resmi |

---

*Dokumen ini disusun berdasarkan Kisi-kisi & Deskripsi Teknis LKS 2026. Marking Scheme CIS tetap acuan utama untuk poin resmi.*
