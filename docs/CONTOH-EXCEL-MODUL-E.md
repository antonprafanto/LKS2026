# Contoh Pengisian Excel — Modul E (untuk Juri)

**Panduan lengkap (runut):** [PANDUAN-MODUL-E.md](PANDUAN-MODUL-E.md)  
**Panduan Log Run:** [PANDUAN-LOG-RUN.md](PANDUAN-LOG-RUN.md)  
**File latihan:** [`templates/LKS2026-Penilaian-Juri-CONTOH.xlsx`](../templates/LKS2026-Penilaian-Juri-CONTOH.xlsx)

Dokumen ini adalah **ringkasan contoh angka** — untuk penjelasan runut lengkap, baca **PANDUAN-MODUL-E.md**.

---

## Inti konsep (1 kalimat)

> **Modul E tidak diisi manual untuk posisi/target** — juri mengisi **sheet Undian Kubus** dulu, lalu Modul E menampilkan hasilnya otomatis.

---

## Langkah 1 — Isi warna marker (baris 24)

Buka sheet **Undian Kubus**. Di **baris 24** (baris kuning), isi warna tiap slot rak.

**Kode R dan S (baris 23):**
- **R1, R2, R3** = **Rak 1, Rak 2, Rak 3** (tiga rak di arena)
- **S1, S2, S3** = **Slot kiri, tengah, kanan** (dilihat dari depan rak)
- **Marker** = warna label di **atas** slot — kubus sukses jika ditaruh di slot yang marker-nya sama dengan warna kubus

| | Slot kiri | Slot tengah | Slot kanan |
|---|----------|------------|-----------|
| **Rak 1** | isi warna | isi warna | isi warna |
| **Rak 2** | isi warna | isi warna | isi warna |
| **Rak 3** | isi warna | isi warna | isi warna |

Contoh isian baris 24 kolom A–I:

```
Merah | Hijau | Biru | Biru | Merah | Hijau | Hijau | Biru | Merah
```

Tanpa baris ini → Slot Target di Modul E **kosong** atau menampilkan petunjuk.

---

## Langkah 2 — Isi posisi kubus setelah diacak

Di **baris 28–36**, isi kolom A, E, F:

| Kolom | Isi apa | Contoh |
|-------|---------|--------|
| **A** | ID kubus | `K-M2` |
| **E** | Rak atau Stand | `Stand A` atau `Rak 1` |
| **F** | Slot | `kiri`, `tengah`, `kanan` (atau `depan` untuk stand) |

### Contoh lengkap 9 kubus (Tim 05, trial H2)

| Baris | A (ID) | E (Rak/Stand) | F (Slot) | Sudah benar? |
|-------|--------|---------------|----------|--------------|
| 28 | K-M1 | Rak 1 | kiri | 1 ✅ |
| 29 | K-M2 | Stand A | depan | 0 ❌ |
| 30 | K-M3 | Rak 2 | tengah | 1 ✅ |
| 31 | K-H1 | Rak 3 | kiri | 1 ✅ |
| 32 | K-H2 | Stand B | depan | 0 ❌ |
| 33 | K-H3 | Rak 1 | kanan | 0 ❌ |
| 34 | K-B1 | Rak 2 | kiri | 1 ✅ |
| 35 | K-B2 | Rak 3 | tengah | 0 ❌ |
| 36 | K-B3 | Stand C | depan | 0 ❌ |

Kolom **B, C, D, H, I, J** terisi otomatis — jangan ketik manual.

---

## Langkah 3 — Lihat hasil di Modul E

Buka sheet **Modul E**. Kolom B–G terisi sendiri.

### Contoh baris K-M2 (kubus merah di Stand A)

| Kolom | Nilai otomatis | Artinya |
|-------|----------------|---------|
| ID Kubus | K-M2 | dari Undian A26 |
| Warna | Merah | dari katalog |
| Bentuk | Persegi | dari katalog |
| **Posisi Awal** | `Stand A — depan` | dari Undian E+F |
| **Slot Target** | `R1-S1 (kiri), R2-S2, R3-S3` | slot valid untuk kubus merah |
| **OK Awal (1/0)** | `0` | belum di rak yang benar |

### Contoh baris K-M1 (sudah benar di rak)

| Kolom | Nilai |
|-------|-------|
| Posisi Awal | `Rak 1 — kiri` |
| Slot Target | `R1-S1 (kiri), R2-S2, R3-S3` |
| OK Awal | **1** |
| **Hasil Run (1/0)** | juri isi **1** (sudah benar + tetap stabil) |

---

## Langkah 4 — Juri isi skor saja

Di Modul E, kolom **Hasil Run (1/0)** (kolom H):

| ID | Dipindah robot? | Berhasil? | Isi |
|----|-----------------|-----------|-----|
| K-M2 | Ya | Ya | **1** |
| K-H2 | Ya | Jatuh | **0** |
| K-M1 | Tidak (sudah benar) | — | **1** |

Skor per kubus = `Hasil × 6,67` → total max 60.

---

## Diagram alur

```
┌─────────────────────────────────────┐
│  SHEET: UNDIAN KUBUS                │
│  ① Baris 24 → warna marker          │
│  ② Kolom A  → ID kubus              │
│  ③ Kolom E  → Rak/Stand               │
│  ④ Kolom F  → Slot                    │
└──────────────┬──────────────────────┘
               │ rumus Excel
               ▼
┌─────────────────────────────────────┐
│  SHEET: MODUL E                     │
│  Otomatis: ID, Warna, Posisi, Target│
│  Juri isi: Hasil Run (1/0) saja     │
└─────────────────────────────────────┘
```

---

## Kenapa screenshot kamu tampil petunjuk?

Itu artinya **Langkah 1 atau 2 belum selesai** di Undian Kubus:

| Tampilan di Modul E | Penyebab | Solusi |
|---------------------|----------|--------|
| `→ isi Rak/Stand & Slot` | Kolom E/F kosong | Isi di Undian |
| `→ isi warna marker baris 24` | Baris 24 kosong | Isi warna marker |
| Slot target / posisi asli | Undian sudah lengkap | Lanjut isi Hasil Run |

---

## Regenerasi file contoh

```bash
python scripts/generate_excel_templates.py
python scripts/generate_example_workbook.py
```
