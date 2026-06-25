# Panduan Lengkap Modul E — Performansi Otonom (untuk Juri)

**Bobot:** 60 poin (60% total skor)  
**Hari:** H2 trial + H3 final  
**Sheet Excel:** `Undian Kubus` → `Modul E` → `Log Run Otonom`  
**File latihan:** [`LKS2026-Penilaian-Juri-CONTOH.xlsx`](../templates/LKS2026-Penilaian-Juri-CONTOH.xlsx)

**Dokumen terkait:**
- [Panduan Log Run Otonom](PANDUAN-LOG-RUN.md) — checklist SIAP/START + kronologi per run
- [SOP-JURI.md §9](SOP-JURI.md#9-sop-penilaian-modul-e--performansi-otonom-60-poin-m) — aturan resmi lomba
- [Skenario latihan](SKENARIO-SOAL-CONTOH.md) — praktik H-1

---

## Inti (baca ini dulu)

> **Modul E hampir tidak diisi manual.** Juri mengisi **Undian Kubus** dulu → kolom B–G di Modul E terisi otomatis → juri hanya isi **Hasil Run (1/0)** kolom H setelah run selesai.

```
Undian Kubus (juri ISI)  ──rumus──▶  Modul E (juri CEK + isi kolom H)
                                           │
                                           ▼
                                    Log Run (juri ISI checklist + kronologi)
```

---

## Bagian 1 — Urutan waktu satu run (runut)

Ikuti urutan ini **setiap run** Modul E:

| # | Waktu (contoh) | Yang terjadi | Apa yang juri lakukan |
|---|----------------|--------------|------------------------|
| 1 | −15 mnt | Briefing layout arena | Jelaskan posisi START, 3 rak, gate, obstacle ke tim |
| 2 | −10 s/d 0 | Tim programming | Tim **boleh** sentuh laptop |
| 3 | 0 | Tim nyatakan **SIAP** | Pastikan laptop **tidak disentuh lagi** |
| 4 | +0 mnt | Juri **acak** posisi 9 kubus | Isi sheet **Undian Kubus** (baris 24, 28–36, obstacle) |
| 5 | +1 mnt | Sinyal **START** | Tim tekan tombol START di robot (bukan laptop) |
| 6 | +1–15 mnt | Robot berjalan **otonom** | Pantau; **jangan sentuh robot** |
| 7 | Selesai | Run berakhir | Isi **Modul E kolom H** + **Log Run Otonom** |
| 8 | Opsional | Tim minta **RETRY** | Ikuti prosedur retry (Bagian 8) |

---

## Bagian 2 — Isi Undian Kubus (WAJIB sebelum Modul E)

Buka sheet **Undian Kubus**. Isi header run (Run ID, Tim No, Hari, dll.) jika belum.

### Langkah 2.1 — Warna marker (baris 24, baris kuning)

Baris ini mendefinisikan **warna label di atas setiap slot rak**. Kubus sukses jika ditaruh di slot yang marker-nya **sama warna** dengan kubus.

**Kode di Excel (baris 23):**

| Kode | Artinya |
|------|---------|
| **R1, R2, R3** | Rak 1, Rak 2, Rak 3 |
| **S1, S2, S3** | Slot kiri, tengah, kanan (dilihat dari depan rak) |

**Layout baris 24:**

| | Kiri (S1) | Tengah (S2) | Kanan (S3) |
|---|-----------|-------------|------------|
| **Rak 1** | ketik warna | ketik warna | ketik warna |
| **Rak 2** | ketik warna | ketik warna | ketik warna |
| **Rak 3** | ketik warna | ketik warna | ketik warna |

Ketik persis: `Merah`, `Hijau`, atau `Biru` (sesuai dropdown/soal).

**Contoh isian kolom A–I baris 24:**

```
Merah | Hijau | Biru | Biru | Merah | Hijau | Hijau | Biru | Merah
```

Tanpa baris ini → kolom **Slot Target** di Modul E menampilkan petunjuk `→ isi warna marker baris 24`.

---

### Langkah 2.2 — Posisi kubus setelah diacak (baris 28–36)

Setelah tim **SIAP**, acak posisi 9 kubus (beberapa **sengaja salah** — di stand/lantai atau slot marker tidak cocok).

Isi **hanya 3 kolom** per baris:

| Kolom | Isi | Cara |
|-------|-----|------|
| **A** | ID kubus | `K-M1`, `K-H2`, `K-B3`, … (9 baris) |
| **E** | Rak/Stand | Dropdown: `Rak 1`–`Rak 3`, `Stand A`–`C`, `Lantai` |
| **F** | Slot | Dropdown: `kiri`, `tengah`, `kanan` (stand: `depan`) |

**Jangan ketik manual** kolom B, C, D, G, H, I, J di Undian — terisi rumus.

**Contoh 9 kubus (Tim 05, trial):**

| Baris | ID | Rak/Stand | Slot | Sudah benar? |
|-------|-----|-----------|------|--------------|
| 28 | K-M1 | Rak 1 | kiri | 1 ✅ |
| 29 | K-M2 | Stand A | depan | 0 ❌ |
| 30 | K-M3 | Rak 2 | tengah | 1 ✅ |
| 31 | K-H1 | Rak 3 | kiri | 1 ✅ |
| 32 | K-H2 | Stand B | depan | 0 ❌ |
| 33 | K-H3 | Rak 1 | kanan | 0 ❌ |
| 34 | K-B1 | Rak 2 | kiri | 1 ✅ |
| 35 | K-B2 | Rak 3 | tengah | 0 ❌ |
| 36 | K-B3 | Stand C | depan | 0 ❌ |

---

### Langkah 2.3 — Obstacle (Obs 1–10)

Catat lokasi singkat obstacle yang dipakai di run ini. Biasanya 6–10 buah sesuai soal hari itu.

**Contoh isian:**

| Obs 1 | Obs 2 | Obs 3 | Obs 4 | Obs 5 | Obs 6 |
|-------|-------|-------|-------|-------|-------|
| kiri GATE | tengah arena | dekat Rak 2 | kanan bawah | antara START–Rak1 | (kosong) |

Kolom tidak dipakai → dikosongkan.

---

## Bagian 3 — Sheet Modul E (penilaian 60 poin)

### Langkah 3.1 — Buka Modul E, cek kolom otomatis

Setelah Undian lengkap, buka sheet **Modul E**. Baris 7–15 (9 kubus) harus terisi:

| Kolom | Nama | Siapa isi? | Fungsi |
|-------|------|------------|--------|
| A | No | Otomatis | 1–9 |
| B | ID Kubus | Otomatis | Dari Undian A28–A36 |
| C | Warna | Otomatis | Dari katalog |
| D | Bentuk | Otomatis | Dari katalog |
| E | Posisi Awal | Otomatis | Rak/Stand + Slot |
| F | Slot Target | Otomatis | Daftar slot valid untuk warna kubus |
| G | OK Awal (1/0) | Otomatis | Referensi: sudah benar **sebelum** run? |
| **H** | **Hasil Run (1/0)** | **JURI** | **Skor resmi** |
| I | Bobot | Otomatis | 60÷9 ≈ 6,67 |
| J | Skor | Otomatis | Hasil × Bobot |

**Yang juri ketik manual: hanya kolom H.**

### Langkah 3.2 — Contoh baris otomatis

**K-M2 (merah di Stand A — belum benar):**

| Kolom | Nilai |
|-------|-------|
| Posisi Awal | `Stand A — depan` |
| Slot Target | `R1-S1 (Rak1 kiri), R1-S2, …` (semua slot marker merah) |
| OK Awal | **0** |

**K-M1 (merah di Rak 1 kiri — sudah benar):**

| Kolom | Nilai |
|-------|-------|
| Posisi Awal | `Rak 1 — kiri` |
| OK Awal | **1** |

---

## Bagian 4 — Isi Hasil Run (1/0) setelah run selesai

### Kriteria sukses (isi **1**)

Kubus **berhasil** jika **semua** terpenuhi:

1. Ditaruh di slot rak yang **marker warnanya cocok** dengan warna kubus  
2. Kubus **stabil** (tidak jatuh / setengah keluar slot)  
3. Penempatan terjadi **selama run otonom** (tanpa sentuh robot setelah START)

### Kriteria gagal (isi **0**)

Salah slot, jatuh, tidak dipindah robot, place manual peserta, atau run dibatalkan.

### Cara menilai per kubus

| Situasi | Hasil Run |
|---------|-----------|
| Robot berhasil pindahkan ke slot benar + stabil | **1** |
| Robot gagal / jatuh / salah slot | **0** |
| OK Awal = 1, kubus tidak dipindah, tetap stabil | **1** *(konfirmasi CE)* |
| OK Awal = 1 tapi robot memindahkan lalu gagal | **0** |

> **Klarifikasi wajib ke Chief Expert:** Apakah kubus OK Awal = 1 otomatis dapat Hasil Run = 1 tanpa dipindah robot?

### Contoh penilaian Run 1 (Tim 05)

| ID | OK Awal | Hasil Run | Keterangan |
|----|---------|-----------|------------|
| K-M1 | 1 | **1** | Sudah benar, tetap stabil |
| K-M2 | 0 | **1** | Robot berhasil pindahkan |
| K-H2 | 0 | **0** | Jatuh saat place |
| K-M3 | 1 | **1** | … |
| … | … | … | 9 baris total |

**Skor run:** jumlah Hasil Run × (60÷9). Contoh 8 sukses ≈ **53,33** / 60 (lihat sel J16).

---

## Bagian 5 — Troubleshooting Modul E

| Tampilan di Modul E | Penyebab | Solusi |
|---------------------|----------|--------|
| `→ isi Rak/Stand & Slot` | Kolom E/F Undian kosong | Lengkapi baris 28–36 |
| `→ isi warna marker baris 24` | Baris 24 kosong | Isi warna marker |
| ID/warna/posisi normal | Undian lengkap | Lanjut isi **Hasil Run** kolom H |
| `#NAME?` atau error | Excel terlalu lama | Pakai Excel 2010+ desktop |

---

## Bagian 6 — Checklist juri per run

- [ ] Baris 24 marker terisi (9 sel)
- [ ] 9 kubus terisi (kolom A, E, F) setelah acak
- [ ] Obstacle dicatat (jika dipakai)
- [ ] Briefing → SIAP → acak → START sesuai SOP
- [ ] Modul E kolom H terisi semua 9 baris
- [ ] [Log Run Otonom](PANDUAN-LOG-RUN.md) checklist + kronologi terisi
- [ ] Tanda tangan juri & peserta

---

## Bagian 7 — Skema poin

| Item | Nilai |
|------|-------|
| Per kubus berhasil | 60 ÷ 9 ≈ **6,67** poin |
| Maksimum per run | **60** poin |
| Rumus Excel | `=60/9` di kolom I, total di J16 |

> Jika Marking Scheme CIS berbeda, **CIS yang mengikat**.

**Agregasi H2 vs H3:** keputusan Chief Expert — trial/final dijumlahkan, diratakan, atau hanya final.

---

## Bagian 8 — Prosedur RETRY

| Aturan | Detail |
|--------|--------|
| Permintaan | Peserta minta ke juri (kebijakan CE) |
| Ubah program | Diizinkan sebelum run ulang |
| Posisi kubus | **TIDAK** dikembalikan ke kondisi Run sebelumnya |
| Kubus di robot | Harus dikeluarkan ke luar arena |
| Setelah SIAP lagi | Posisi kubus **diacak ulang** → update Undian → run baru |

Catat run retry di **Log Run** (Retry = 1, Retry No = 2, dll.).

---

## Bagian 9 — Pelanggaran

| Pelanggaran | Tindakan juri |
|-------------|---------------|
| Sentuh laptop setelah SIAP | Run batal / skor 0 run (konfirmasi CE) |
| Sentuh robot setelah START | Run batal / skor 0 run |
| Tidak start via tombol robot | Run batal |
| Place manual oleh peserta | Item/run batal |

Centang pelanggaran di bagian bawah sheet Modul E dan catat di Log Run.

---

## Diagram ringkas

```
┌─────────────────────────────────────────┐
│ UNDIAN KUBUS                            │
│ ① Baris 24 → warna marker (9 sel)       │
│ ② Baris 28–36 → ID | Rak/Stand | Slot   │
│ ③ Obs 1–10 → lokasi obstacle            │
└──────────────────┬──────────────────────┘
                   │ rumus Excel
                   ▼
┌─────────────────────────────────────────┐
│ MODUL E                                 │
│ Otomatis: B–G, I–J                      │
│ Juri isi: H (Hasil Run 1/0)             │
└──────────────────┬──────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────┐
│ LOG RUN OTONOM                          │
│ Checklist 8 langkah + kronologi + ttd   │
└─────────────────────────────────────────┘
```

---

*Panduan ini untuk juri LKS Nasional 2026 — Robot Bergerak Otonom. Acuan resmi: SOP-JURI.md + Marking Scheme CIS.*
