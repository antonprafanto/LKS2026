# Panduan Google Sheets — Penilaian Juri LKS 2026

**Penting:** Template resmi penilaian juri adalah **Microsoft Excel** (`.xlsm` / `.xlsx`).  
Jika Anda pakai **Google Sheets**, beberapa fitur **tidak otomatis jalan** setelah upload/import.

---

## Kenapa stopwatch tidak jalan di Google Sheets?

| Fitur di Excel | Di Google Sheets |
|----------------|------------------|
| Tombol **▶ SIAP / START / SELESAI** (macro VBA) | **Tidak ada** — VBA tidak didukung |
| File **.xlsm** | Macro **dihapus** saat di-upload |
| **F9** refresh durasi | **Tidak ada** |
| Rumus `NOW()` durasi live | Hanya update saat sel **diedit** |

---

## Instalasi stopwatch (SATU FILE saja)

### Langkah A — Paste script

1. Buka spreadsheet di Google Sheets  
2. **Extensions → Apps Script** *(wajib dari spreadsheet ini, bukan script.google.com)*  
3. Hapus **semua** isi file `Code.gs`  
4. Paste **seluruh** isi:  
   https://raw.githubusercontent.com/antonprafanto/LKS2026/main/scripts/google-apps-script/LogRunStopwatch.gs  
5. Klik **Save** (ikon disket)

> **Tidak perlu** file HTML terpisah — versi terbaru sudah satu file.

### Langkah B — Buat menu (WAJIB, sekali)

Menu **tidak muncul otomatis** sampai Anda Run + izinkan akses:

1. Di Apps Script, dropdown fungsi → pilih **`buatMenuStopwatch`**
2. Klik **Run** ▶
3. **Authorize / Izinkan** (Review Permissions → pilih akun → Advanced → Allow)
4. Harus muncul popup: *"Menu LKS Stopwatch sudah dibuat"*
5. **Kembali ke tab spreadsheet** — menu **`LKS Stopwatch`** ada di **bar atas** (sebelah Help)

### Langkah C — Setup rumus (sekali)

1. Apps Script → pilih **`setupLogRunSheet`** → **Run**
2. Kembali ke tab **Log Run Otonom** — kolom K7/K8/K9 terisi rumus

### Langkah D — Tes

1. Menu **LKS Stopwatch → SIAP** → sel **B8** terisi waktu  
2. **START** → **E8** terisi  
3. **SELESAI** → **H8** terisi, **K8** durasi final

---

## Kenapa Jam SIAP bukan 00:00:00?

**B8, E8, H8** mencatat **jam di dinding** (kapan Anda klik), **bukan** stopwatch dari nol.

Contoh dari run Anda:

| Sel | Nilai | Artinya |
|-----|-------|---------|
| B8 Jam SIAP | `01:15:53` | Pukul 01:15:53 tim SIAP |
| E8 Jam START | `01:16:03` | 10 detik kemudian, START |
| H8 Jam selesai | `01:16:07` | Run selesai |
| **K8 Durasi run** | **`0:00:04`** | **Ini stopwatch dari 0** — run hanya 4 detik |

Yang **dari 0** hanya **kolom K** (durasi START → SELESAI). Jam SIAP sengaja jam nyata supaya log run bisa diaudit CE (*"tim SIAP jam berapa?"*).

---

## Jam SIAP tampil 01:16 padahal sekarang 09:18? (timezone)

Google Sheets default sering **UTC**. Samarinda pakai **WITA (UTC+8)**.

| Yang Anda lihat | Penyebab |
|-----------------|----------|
| `01:16` di B8 | Waktu **UTC** (bukan jam Samarinda) |
| Sekarang `09:18` Samarinda | = `01:18` UTC → selisih **8 jam** |

**Perbaikan:**

1. Spreadsheet → **File → Settings** → **Time zone** → **`(GMT+08:00) Asia/Makassar`** (WITA)  
2. Apps Script → Run **`setupLogRunSheet`** (script terbaru set timezone otomatis)  
3. Menu **LKS Stopwatch → Reset waktu**  
4. Klik **SIAP / START / SELESAI** lagi → harus tampil **~09:18**, bukan 01:18  

> Durasi di **K7/K8** tetap benar meski timezone salah — hanya **jam tampilan** B8/E8/H8 yang melenceng.

---

## Menu tidak muncul? Checklist

| # | Cek | Solusi |
|---|-----|--------|
| 1 | Script dibuka dari **script.google.com** terpisah | Hapus project itu. Buka spreadsheet → **Extensions → Apps Script** dari file yang benar |
| 2 | Belum **Run + Authorize** | Run **`buatMenuStopwatch`** → izinkan semua |
| 3 | Masih di tab Apps Script | Klik **tab spreadsheet** (bukan tab script) |
| 4 | Tab belum di-refresh | **Tutup** spreadsheet → buka lagi dari Drive |
| 5 | Error saat Run | Apps Script → **Executions** (ikon jam) → baca error merah |
| 6 | Tes koneksi script | Run **`testScriptOK`** — harus popup "Script OK" |
| 7 | Nama tab salah | Harus ada tab persis **`Log Run Otonom`** (spasi, huruf besar kecil sama) |

---

## Pintasan manual (tanpa menu)

Klik sel **B8** / **E8** / **H8** → **`Ctrl+Shift+;`**

---

## Rumus manual (jika setup gagal)

| Sel | Rumus |
|-----|--------|
| **K7** | `=IF(E8="","",IF(H8<>"",H8-E8,JAM_WITA()-E8))` |
| **K8** | `=IF(AND(E8<>"",H8<>""),H8-E8,"")` |
| **K9** | `=IF(K8="","",ROUND(K8*24*60,1))` |

---

## Rekomendasi

| Situasi | Saran |
|---------|--------|
| Hari lomba | **Excel .xlsm** (paling stabil) |
| Google Sheets | Script di atas + menu **LKS Stopwatch** |

Script: [`LogRunStopwatch.gs`](../scripts/google-apps-script/LogRunStopwatch.gs)
