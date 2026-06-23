# Template Penilaian — Panduan

## Pakai file ini saja

**[`LKS2026-Penilaian-Juri.xlsx`](LKS2026-Penilaian-Juri.xlsx)** — satu workbook Excel berisi semua lembar penilaian dengan kolom terpisah dan rumus otomatis.

> File `.csv` di folder ini adalah format lama. Buka langsung di Excel akan tampil mentah (satu kolom). **Gunakan file `.xlsx` di atas.**

## Sheet di dalam Excel

| Sheet | Fungsi |
|-------|--------|
| **Panduan** | Cara pakai singkat |
| **Modul A–E** | Penilaian semua modul + rumus skor |
| **Undian Kubus** | Pengacakan posisi — ID kubus → warna/bentuk otomatis |
| **Log Run Otonom** | Checklist SIAP/START + kronologi + tarik skor dari Modul E |
| **Rekapitulasi** | Total A–E otomatis |

## Alur pakai (per tim)

1. **Save As** — duplikat 1 file Excel per tim.
2. Isi **Undian Kubus** setelah pengacakan (kolom ID kubus).
3. Saat run otonom: isi **Log Run Otonom** (checklist + kronologi).
4. Isi **Modul E** — ketik ID kubus di kolom B → warna/bentuk/posisi awal terisi otomatis.
5. Centang **Hasil (1/0)** per kubus → skor terhitung.
6. **Rekapitulasi** mengambil total semua modul.
7. Input skor resmi ke **CIS**.

## Rumus penting

- **Undian:** `VLOOKUP` ID kubus → katalog warna & bentuk
- **Modul E:** terhubung ke sheet Undian Kubus
- **Log Run:** `COUNTIF` kubus sukses + skor dari Modul E
- **Modul A/B/C:** rata-rata 3 juri → skor akhir

## Regenerasi file Excel

```bash
python scripts/generate_excel_templates.py
```

Acuan: [`docs/SOP-JURI.md`](../docs/SOP-JURI.md)
