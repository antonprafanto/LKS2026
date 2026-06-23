# SOP JURI — Robot Bergerak Otonom LKS Nasional 2026

**Versi:** 1.0  
**Acuan:** Kisi-kisi LKS 2026 + Deskripsi Teknis v1.0 (WorldSkills Skill 23)  
**Tema proyek uji:** Service Robot at Library

---

## 1. Ruang Lingkup

SOP ini mengatur peran juri dalam penilaian **5 modul** (A–E) selama 3 hari lomba + H-1. Berlaku untuk semua juri bidang Autonomous Mobile Robotics.

### Dokumen wajib dipegang juri

- [ ] Marking Scheme Excel CIS (sumber kebenaran poin detail)
- [ ] Kisi-kisi & Deskripsi Teknis (tersedia di repo ini)
- [ ] Work Instructions / Test Project hari H
- [ ] Petunjuk Teknis Umum LKS 2026

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

### 3.1 H-30: Technical Meeting 1

- Pahami Marking Scheme CIS
- Klafikasi ke Chief Expert: ukuran arena, tegangan baterai, jadwal Modul D

### 3.2 H-1

| Waktu | Tugas juri |
|-------|------------|
| Technical Meeting 2 | Sampaikan prosedur SIAP/START/RETRY; jelaskan 30% rule |
| Familiarisasi (maks. 2 jam) | Pastikan 4 set arena lengkap; verifikasi aksesoris |

**Checklist arena per set:**

- [ ] Arena (multiplex 18 mm, permukaan putih matte)
- [ ] 3 rak + marker warna di atas
- [ ] 9 kubus (3 warna × 3 bentuk)
- [ ] Kubus stand (80×80×100 mm)
- [ ] 10 obstacle
- [ ] Reference line (isolasi 18×300 mm)
- [ ] 1 gate (profil 2020)
- [ ] Penerangan: tiang + lampu 100 W × 2

### 3.3 Hari 1 (H1)

| Sesi | Modul | Tugas juri |
|------|-------|------------|
| 08.30–12.00 | C | Awasi perakitan L-Channel + lift (batas **2 jam**); verifikasi kelayakan mekanik & K3 |
| Sepanjang H1 | A | Observasi 5S, kolaborasi, penjadwalan |
| Sepanjang H1 | B | Terima jurnal; nilai kelengkapan & kualitas |
| 13.00–16.00 | D (pra-uji) | Fasilitasi latihan gerakan dasar (penilaian resmi H2) |

**Verifikasi Modul C sebelum robot dioperasikan:**

- [ ] L-Channel terpasang sebagai penyangga utama sistem angkat
- [ ] Tidak ada kabel terbuka / terjepit
- [ ] Tombol STOP merah berfungsi
- [ ] Rewiring rapi & aman

### 3.4 Hari 2 (H2)

| Sesi | Modul | Tugas juri |
|------|-------|------------|
| Pagi–siang | D | Penilaian resmi 15 menit/tim (lihat Template Modul D) |
| 08.30–15.00 | E | Supervisi trial otonom |
| 15.00–17.30 | E | Marking trial |

### 3.5 Hari 3 (H3)

| Sesi | Modul | Tugas juri |
|------|-------|------------|
| Briefing | E | Terapkan variasi soal (30% rule) |
| 08.30–15.00 | E | Supervisi run final |
| 15.00–17.30 | E | Marking final |
| Sepanjang H3 | A | Observasi berkelanjutan |

---

## 4. SOP Penilaian Modul A — Organisasi & Manajemen Kerja (8 poin, J)

**Hari penilaian:** H1, H2, H3  
**Template:** sheet `Modul A` di `templates/LKS2026-Penilaian-Juri.xlsx`

| Sub | Kriteria | Max (contoh CIS) |
|-----|----------|------------------|
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

## 5. SOP Penilaian Modul B — Jurnal Teknis (10 poin: J 8 + M 2)

**Hari penilaian:** H1  
**Template:** sheet `Modul B` di `templates/LKS2026-Penilaian-Juri.xlsx`  
**Batas:** Maks. 4 lembar isi (8 halaman A4), tidak termasuk sampul

### Checklist kelengkapan (Judgement)

- [ ] Analisis tugas & strategi
- [ ] Desain mekanik (CAD/sketsa + dimensi)
- [ ] Diagram blok sistem
- [ ] Bill of Materials
- [ ] Dokumentasi fabrikasi & wiring
- [ ] Penempatan sensor (justifikasi)
- [ ] Arsitektur software & library
- [ ] Algoritma navigasi & visi komputer
- [ ] Flowchart program
- [ ] Data kalibrasi & tuning
- [ ] Log troubleshooting

### B6 — Ketepatan waktu pengumpulan (Measurement)

| Kondisi | Skor M |
|---------|--------|
| Jurnal terkumpul sebelum deadline | 1 |
| Terlambat / tidak terkumpul | 0 |

> Detail poin Measurement 2 poin lainnya mengacu Marking Scheme CIS.

---

## 6. SOP Penilaian Modul C — Perakitan Robot (10 poin, J)

**Hari penilaian:** H1  
**Template:** sheet `Modul C` di `templates/LKS2026-Penilaian-Juri.xlsx`

### Alur inspeksi

1. **Sebelum run 2 jam:** Catat progress perakitan L-Channel.
2. **Setelah 2 jam:** Waktu perakitan mekanik + rewiring harus selesai.
3. **Inspeksi akhir:** L-Channel + sistem angkat + kabel + integrasi aktuator.

| Aspek | Indikator lulus inspeksi |
|-------|--------------------------|
| L-Channel | Terpasang sebagai penyangga utama lift |
| Mekanik | Stabil, tidak goyang berlebihan |
| Wiring | Rapi, tidak terbuka, tidak terjepit |
| Motor | Kompatibel JGA-25 / PG36 (jika dipakai) |
| K3 | Tombol STOP merah accessible & berfungsi |

---

## 7. SOP Penilaian Modul D — Gerakan Dasar & Manajemen Objek (12 poin, M)

**Hari penilaian:** H2 (resmi)  
**Waktu per tim:** 15 menit  
**Template:** `templates/modul-d-gerakan-dasar.csv`

### Prosedur sesi

1. Panggil tim ke zona uji.
2. Pastikan monitor laptop terlihat juri (untuk D4–D8).
3. Baca item satu per satu; catat **1** atau **0**.
4. Tidak ada retry untuk item yang sudah dinilai (kecuali Chief Expert menetapkan lain).

### Tabel item & bobot

| ID | Tugas | Bukti | Bobot |
|----|-------|-------|-------|
| D1a | Gerak maju 200–300 mm | Observasi | 0.50 |
| D1b | Gerak mundur 200–300 mm | Observasi | 0.50 |
| D1c | Gerak kanan 200–300 mm | Observasi | 0.50 |
| D1d | Gerak kiri 200–300 mm | Observasi | 0.50 |
| D2a | Maju, stop 50–100 mm dari dinding | Observasi | 0.30 |
| D2b | Kanan, stop 50–100 mm dari dinding | Observasi | 0.60 |
| D3 | Ikuti garis hitam bentuk U | Line follower | 1.00 |
| D4 | Deteksi kubus merah → teks "merah" di monitor | Screenshot/lihat | 1.00 |
| D5 | Deteksi kubus hijau → teks "hijau" | Screenshot/lihat | 1.00 |
| D6 | Deteksi kubus biru → teks "biru" | Screenshot/lihat | 1.00 |
| D7 | Deteksi lingkaran merah → "O-merah" | Screenshot/lihat | 1.00 |
| D8 | Deteksi lingkaran hijau → "O-hijau" | Screenshot/lihat | 1.00 |
| D9 | Ambil kubus dari rak secara otonom | Observasi arm | 2.00 |
| D10 | Letakkan kubus ke standbox | Observasi arm | 2.00 |
| | **TOTAL** | | **12.00** |

> **Catatan:** Modul D tidak menguji bentuk persegi/segitiga secara eksplisit. Semua 9 kombinasi kubus relevan untuk Modul E.

---

## 8. SOP Penilaian Modul E — Performansi Otonom (60 poin, M)

**Hari penilaian:** H2 (trial) + H3 (final)  
**Template:** `templates/modul-e-otonom.csv`, `templates/lembar-run-otonom.csv`, `templates/lembar-undian-kubus.csv`

### 8.1 Alur run (WAJIB diikuti)

```
┌──────────────────────────────────────────────────────────────┐
│ 1. Briefing layout arena (start, rak, gate, obstacle)        │
│ 2. Peserta selesai programming (boleh sentuh laptop)         │
│ 3. Peserta nyatakan SIAP                                     │
│ 4. ► Peserta TIDAK BOLEH sentuh laptop setelah SIAP          │
│ 5. Juri ACAK ULANG posisi kubus (undian)                     │
│ 6. Juri beri sinyal START                                    │
│ 7. Peserta tekan tombol START di robot                       │
│ 8. Robot berjalan OTOMONOM                                   │
│ 9. Juri catat hasil penempatan kubus                         │
└──────────────────────────────────────────────────────────────┘
```

### 8.2 Prosedur RETRY

| Aturan | Detail |
|--------|--------|
| Permintaan | Peserta minta ke juri sebelum/during run (sesuai kebijakan CE) |
| Ubah program | **Diizinkan** sebelum run ulang |
| Posisi kubus | **TIDAK** dikembalikan ke kondisi awal |
| Kubus di robot | **Harus dikeluarkan** ke luar arena |
| Setelah SIAP lagi | Posisi kubus **diacak ulang** |

### 8.3 Kriteria sukses penempatan kubus

Kubus dianggap **berhasil** jika:

1. Kubus berada di **rak yang benar** sesuai **marker warna** di atas slot rak, **dan**
2. Kubus **tidak jatuh** / tidak setengah keluar slot, **dan**
3. Penempatan dilakukan **selama run otonom** (tanpa sentuh robot setelah START)

### 8.4 Skema poin (template default)

| Item | Bobot |
|------|-------|
| Penempatan kubus 1 benar | 6.67 |
| Penempatan kubus 2 benar | 6.67 |
| … (hingga 9 kubus) | 6.67 |
| **TOTAL** | **60.00** |

> Jika Marking Scheme CIS menggunakan bobot berbeda per kubus/warna, **CIS yang mengikat**.

### 8.5 Pelanggaran → skor 0 untuk run/item

- Peserta menyentuh laptop setelah SIAP
- Peserta menyentuh robot setelah START (tanpa izin)
- Robot tidak dimulai dengan tombol START di robot
- Kubus diangkat/ditempatkan manual oleh peserta

---

## 9. SOP Pengacakan Kubus

**Template:** `templates/lembar-undian-kubus.csv`

### Langkah undian

1. Siapkan 9 kubus: 3 merah, 3 hijau, 3 biru (masing-masing ○ □ △).
2. Tentukan posisi awal: rak (3 rak × 3 slot) + stand di lapangan.
3. Acak dengan undian/box digital — **beberapa kubus sengaja di posisi salah**.
4. Catat di lembar undian: ID kubus, warna, bentuk, posisi awal, posisi target (marker).
5. Setelah peserta SIAP → **ulangi undian** untuk run tersebut.

### Variasi 30% rule (H3)

Chief Expert dapat mengubah hingga 30%:

- Urutan pengiriman
- Posisi obstacle
- Konfigurasi marker warna di rak
- Posisi start

---

## 10. SOP K3 untuk Juri

- [ ] Pastikan tidak ada air mineral terbuka dekat listrik
- [ ] Verifikasi tombol STOP setiap tim sebelum run
- [ ] Obstacle yang ditabrak jatuh — reset sebelum tim berikutnya
- [ ] Laporkan kabel terbuka / baterai aus ke panitia

---

## 11. Rekapitulasi & CIS

1. Kumpulkan semua lembar penilaian per modul.
2. Isi `templates/rekapitulasi-skor.csv`.
3. Chief Expert review outlier (selisih juri ≥ 2).
4. Input ke CIS → konversi skala 700.
5. Arsipkan lembar undian + run otonom per tim.

| Peringkat | Kriteria |
|-----------|----------|
| Juara 1–3 | Skor CIS tertinggi |
| Medallion of Excellence | Skor > 700 |

---

## 12. Klarifikasi Wajib ke Chief Expert

Sebelum lomba, pastikan keputusan tertulis untuk:

| Isu | Opsi di dokumen |
|-----|-----------------|
| Ukuran arena | 4040×2040 mm (Kisi-kisi) vs 2400×4800 mm (Daftar Bahan) |
| Tegangan maks. | Hal. 19 terpotong; referensi baterai @12V |
| Bobot per kubus Modul E | Template repo vs Excel CIS |
| Jumlah retry per tim | Tidak disebutkan di dokumen |

---

*Dokumen ini disusun berdasarkan Kisi-kisi & Deskripsi Teknis LKS 2026. Untuk poin resmi, Marking Scheme CIS tetap menjadi acuan utama.*
