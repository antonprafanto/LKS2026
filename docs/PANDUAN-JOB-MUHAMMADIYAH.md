# Panduan Job Tetap — TIM SMK Muhammadiyah

**File:** `templates/LKS2026-Penilaian-Juri TIM SMK Muhammadiyah.xlsm`  
**Versi:** Job internal / latihan — **bukan** template undian resmi LKS nasional.

---

## Perbedaan dengan template utama

| | Template nasional | Muhammadiyah |
|---|-------------------|--------------|
| Undian Kubus | Acak tiap run, isi manual | **Tidak ada** — pakai **Job Arena** |
| Modul E | Otomatis dari Undian | **Manual** — B–G terisi, juri isi **Hasil Run** |
| Layout | Marker acak per slot | **Tetap** KS1–KS9 + rak per bentuk |

---

## Denah job (ringkas)

### Stand awal (KS)

| Stand | Kubus | Warna | Bentuk |
|-------|-------|-------|--------|
| KS1 | K-M3 | Merah | Segitiga |
| KS2 | K-H3 | Hijau | Segitiga |
| KS3 | K-B3 | Biru | Segitiga |
| KS4 | K-M2 | Merah | Persegi |
| KS5 | K-H2 | Hijau | Persegi |
| KS6 | K-B2 | Biru | Persegi |
| KS7 | K-M1 | Merah | Lingkaran |
| KS8 | K-H1 | Hijau | Lingkaran |
| KS9 | K-B1 | Biru | Lingkaran |

### Rak target (per bentuk)

| Rak | Bentuk kubus | Marker slot |
|-----|--------------|-------------|
| Rak 1 | Lingkaran | Merah · Hijau · Biru |
| Rak 2 | Persegi | Merah · Hijau · Biru |
| Rak 3 | Segitiga | Merah · Hijau · Biru |

### Aturan job

1. Pindahkan semua kubus ke rak yang **bentuknya cocok** + **slot warna cocok**.
2. **Bawa kubus** → boleh lewat **Gate**.
3. **Tanpa kubus** → wajib lewat **Obstacle**.

---

## Cara penilaian Modul E

1. Buka sheet **Job Arena** — cek denah.
2. Jalankan run (Log Run: SIAP → START → SELESAI).
3. Sheet **Modul E** — isi kolom **H (Hasil Run 1/0)** per kubus:

| Isi | Artinya |
|-----|---------|
| **1** | Kubus di rak bentuk benar + warna slot cocok + stabil |
| **0** | Gagal |

Skor otomatis: `Hasil × (60÷9)` — total max **60**.

---

## Regenerasi file

```bash
python scripts/generate_muhammadiyah_workbook.py
```

Gambar denah: `assets/job-muhammadiyah-arena.png`

---

*Untuk lomba resmi LKS nasional tetap gunakan `LKS2026-Penilaian-Juri.xlsm` + Undian Kubus.*
