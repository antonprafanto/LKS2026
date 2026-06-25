# Panduan Google Sheets — Penilaian Juri LKS 2026

**Penting:** Template resmi penilaian juri adalah **Microsoft Excel** (`.xlsm` / `.xlsx`).  
Jika Anda pakai **Google Sheets**, beberapa fitur **tidak otomatis jalan** setelah upload/import.

---

## Kenapa stopwatch tidak jalan di Google Sheets?

| Fitur di Excel | Di Google Sheets |
|----------------|------------------|
| Tombol **▶ SIAP / START / SELESAI** (macro VBA) | **Tidak ada** — VBA tidak didukung |
| File **.xlsm** | Macro **dihapus** saat di-upload |
| **F9** refresh durasi | **Tidak ada** — tidak ada tombol F9 |
| Rumus `NOW()` durasi live | Hanya update saat sel **diedit**, bukan live otomatis |

Itu sebabnya di spreadsheet Anda tombol stopwatch **kosong / tidak ada**, dan kolom **K7 (Durasi live)** tidak bergerak.

---

## Solusi: Apps Script stopwatch (gratis, ~5 menit setup)

### Langkah 2 — Pasang script (sekali)

1. Buka spreadsheet Anda di Google Sheets  
2. **Extensions → Apps Script**  
3. Buat file **`Code.gs`** — salin isi dari repo:  
   [`scripts/google-apps-script/LogRunStopwatch.gs`](../scripts/google-apps-script/LogRunStopwatch.gs)  
4. Klik **+** → **HTML** → nama file **`StopwatchSidebar`**  
   Salin isi: [`scripts/google-apps-script/StopwatchSidebar.html`](../scripts/google-apps-script/StopwatchSidebar.html)  
5. **Save** → di Script Editor pilih fungsi **`setupLogRunSheet`** → klik **Run** (izinkan akses saat diminta)  
6. **Refresh** spreadsheet (F5) — menu **⏱ Stopwatch Run** muncul

> **Setup otomatis:** langkah 5 mengisi rumus K7/K8/K9, teks petunjuk, dan skor Modul E.  
> Bisa juga lewat menu **⏱ Stopwatch Run → ⚙ Setup sheet Log Run (sekali)** setelah refresh.

### Langkah 3 — Pakai saat run

1. Menu baru: **⏱ Stopwatch Run** (pojok kanan atas)  
2. **▶ SIAP** → waktu masuk **B8**  
3. **▶ START** → waktu masuk **E8**  
4. **Buka stopwatch live (sidebar)** → angka jalan tiap detik (opsional)  
5. **■ SELESAI** → waktu masuk **H8**

Atau buka sidebar dulu — di sana ada tombol SIAP/START/SELESAI + jam live.

### Langkah 4 — Rumus (jika setup belum dijalankan)

Di sheet **Log Run Otonom**, paste rumus ini:

| Sel | Rumus |
|-----|--------|
| **K7** | `=IF(E8="","",IF(H8<>"",H8-E8,NOW()-E8))` |
| **K8** | `=IF(AND(E8<>"",H8<>""),H8-E8,"")` |
| **K9** | `=IF(K8="","",ROUND(K8*24*60,1))` |

Format sel K7/K8: **Format → Number → Duration**

### Langkah 5 — Update teks petunjuk (opsional)

Ganti baris 6 dan 10 yang masih bertuliskan «file .xlsm, Enable Macro» menjadi:

> **STOPWATCH:** menu **⏱ Stopwatch Run** (Google Sheets) — bukan tombol Excel.

---

## Pintasan waktu manual (tanpa menu)

| Aksi | Pintasan Google Sheets |
|------|------------------------|
| Catat waktu | Klik sel B8/E8/H8 → **`Ctrl+Shift+;`** |

---

## Hal lain yang sering rusak di Google Sheets

| Masalah | Penyebab | Solusi |
|---------|----------|--------|
| Skor Modul E = 0 | Referensi antar-sheet / import | Pastikan nama tab sama persis: `Modul E`, `Undian Kubus` |
| Checklist «BELUM» | Rumus COUNTIF | Isi **1** di kolom C baris 13–20 |
| Link GitHub baris 2 tidak jalan | Hyperlink Excel | Buka manual: [PANDUAN-LOG-RUN.md](PANDUAN-LOG-RUN.md) |

---

## Rekomendasi panitia

| Situasi | Saran |
|---------|--------|
| Hari lomba, juri di lapangan | **Microsoft Excel .xlsm** (stopwatch + rumus paling stabil) |
| Backup cloud / banyak juri online | Google Sheets **+ Apps Script** (panduan ini) |
| Tanpa internet | Excel desktop wajib |

---

## File sumber script

- [`LogRunStopwatch.gs`](../scripts/google-apps-script/LogRunStopwatch.gs)  
- [`StopwatchSidebar.html`](../scripts/google-apps-script/StopwatchSidebar.html)

*Acuan penilaian tetap: SOP-JURI.md + Marking Scheme CIS.*
