# Template Penilaian — Panduan

## Pakai file ini saja

**[`LKS2026-Penilaian-Juri.xlsx`](LKS2026-Penilaian-Juri.xlsx)** — satu workbook Excel berisi semua lembar penilaian dengan kolom terpisah dan rumus otomatis.

**[`LKS2026-Penilaian-Juri-CONTOH.xlsx`](LKS2026-Penilaian-Juri-CONTOH.xlsx)** — versi **sudah terisi contoh** (buka ini dulu kalau bingung). Panduan: [`docs/CONTOH-EXCEL-MODUL-E.md`](../docs/CONTOH-EXCEL-MODUL-E.md)

> File `.csv` di folder ini adalah format lama. Buka langsung di Excel akan tampil mentah (satu kolom). **Gunakan file `.xlsx` di atas.**

## Sheet di dalam Excel

| Sheet | Fungsi |
|-------|--------|
| **Panduan** | Cara pakai singkat |
| **Modul A–E** | Penilaian semua modul + rumus skor |
| **Undian Kubus** | Pengacakan posisi — isi marker, ID, Rak/Stand, Slot → sisanya otomatis |
| **Log Run Otonom** | Checklist SIAP/START + kronologi + tarik skor dari Modul E |
| **Rekapitulasi** | Total A–E otomatis |

## Alur pakai (per tim)

1. **Save As** — duplikat 1 file Excel per tim.
2. Isi **Undian Kubus**:
   - Baris **marker** (22): warna Merah/Hijau/Biru per slot rak.
   - Kolom **ID Kubus** + **Rak/Stand** + **Slot** per kubus.
   - Otomatis: Warna, Bentuk, Lokasi Awal, Sudah Benar, Target Rak, Target Marker.
3. Saat run otonom: isi **Log Run Otonom** (checklist + kronologi).
4. **Modul E** — kolom B–G terisi otomatis dari Undian; juri cukup isi **Hasil (1/0)**.
5. Skor per kubus terhitung otomatis.
6. **Rekapitulasi** mengambil total semua modul.
7. Input skor resmi ke **CIS**.

## Rumus penting

- **Undian:** `VLOOKUP` ID → warna/bentuk; Rak+Slot → Lokasi Awal; warna kubus + marker → Target
- **Modul E:** tarik Posisi Awal, Slot Target, OK Awal dari Undian Kubus
- **Log Run:** `COUNTIF` kubus sukses + skor dari Modul E
- **Modul A/B/C:** rata-rata 3 juri → skor akhir

## Regenerasi file Excel

```bash
python scripts/generate_excel_templates.py
```

Acuan: [`docs/SOP-JURI.md`](../docs/SOP-JURI.md)
