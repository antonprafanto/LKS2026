# Panduan Lengkap Log Run Otonom (untuk Juri)

**Sheet Excel:** `Log Run Otonom`  
**Fungsi:** Bukti audit trail setiap run Modul E — checklist prosedur SIAP/START, kronologi, dan ringkasan skor.  
**File latihan:** [`LKS2026-Penilaian-Juri-CONTOH.xlsx`](../templates/LKS2026-Penilaian-Juri-CONTOH.xlsx)

**Dokumen terkait:**
- [Panduan Modul E](PANDUAN-MODUL-E.md) — pengisian Undian + penilaian Hasil Run
- [SOP-JURI.md §9.1](SOP-JURI.md#91-alur-run-operasional-wajib) — alur run wajib
- [Skenario latihan](SKENARIO-SOAL-CONTOH.md) — contoh Run 1 & RETRY

---

## Inti (baca ini dulu)

> **Log Run = catatan resmi satu run otonom.** Isi **bersamaan** dengan Modul E setelah run selesai. Checklist memastikan juri tidak melewatkan langkah SIAP/START; kronologi mendokumentasikan apa yang terjadi; bagian bawah **menarik skor otomatis** dari Modul E.

```
Tim SIAP → Juri acak (Undian) → START → Run otonom → Selesai
                                                      │
                    ┌─────────────────────────────────┴────────────────────┐
                    ▼                                                      ▼
            Modul E (Hasil Run)                              Log Run (checklist + kronologi)
```

**Satu sheet Log Run = satu run.** Jika tim punya beberapa run (trial + final, atau RETRY), isi **Run No** berbeda atau duplikat workbook (kebijakan CE).

---

## Bagian 1 — Kapan mengisi Log Run

| Momen | Isi apa |
|-------|---------|
| **Sebelum run** | Header tim, Run ID, Trial/Final |
| **Saat SIAP** | Waktu SIAP, centang checklist item 1–3 |
| **Saat acak kubus** | Centang item 4, isi Undian ulang = 1 |
| **Saat START** | Waktu START, centang item 5–6 |
| **Selama run** | Centang item 7, catat kronologi |
| **Setelah run** | Waktu Selesai, centang item 8, isi kronologi lengkap, pelanggaran, tanda tangan |

---

## Bagian 2 — Header identitas (baris 3–4)

| Sel | Isi | Contoh |
|-----|-----|--------|
| Tim No | Nomor tim | `05` |
| Nama Tim | Nama resmi tim | `SMK Robotik A` |
| Kontingen | Provinsi/kota | `DKI Jakarta` |
| Arena No | Nomor arena | `2` |
| Hari | H2 / H3 | `H2` |
| Tanggal | Tanggal run | `24/06/2026` |
| Juri Penilai | Nama juri utama | `Budi` |
| Juri Pendamping | Nama juri pendamping | `Ani` |

---

## Bagian 3 — DATA RUN (baris 7–9)

| Field | Isi | Kapan / catatan |
|-------|-----|-----------------|
| **Run No** | Urutan run tim hari itu | `1`, `2`, … |
| **Run ID** | ID unik run (samakan Undian Kubus) | `H2-T05-R1` |
| **Trial/Final** | Fase lomba | `Trial` atau `Final` |
| **Waktu SIAP** | Jam tim nyatakan SIAP | `08:40` |
| **Waktu START** | Jam tombol robot ditekan | `08:42` |
| **Waktu Selesai** | Jam run berakhir | `08:55` |
| **Undian ulang (1/0)** | Apakah kubus diacak ulang run ini? | `1` = ya (normal setelah SIAP) |
| **Retry (1/0)** | Apakah ini run ulang setelah gagal? | `0` = run pertama, `1` = retry |
| **Retry No** | Urutan retry | Kosong jika bukan retry; `2` = retry ke-2 |

**Tips:** Run ID yang sama di **Undian Kubus**, **Modul E**, dan **Log Run** memudahkan audit CE.

---

## Bagian 4 — Checklist prosedur juri (8 item)

Centang **1** di kolom OK saat langkah selesai. Excel menampilkan **LENGKAP** jika ke-8 item = 1.

| No | Item checklist | Kapan centang 1 |
|----|----------------|-----------------|
| 1 | Layout arena dijelaskan ke peserta | Setelah briefing layout |
| 2 | Peserta nyatakan SIAP | Tim bilang SIAP |
| 3 | Laptop tidak disentuh setelah SIAP | Juri konfirmasi visual |
| 4 | Posisi kubus diacak / diacak ulang | Setelah acak + Undian terisi |
| 5 | Sinyal START diberikan | Juri memberi sinyal verbal/gesture |
| 6 | Peserta tekan tombol START di robot | Tombol fisik di robot ditekan |
| 7 | Run berjalan otonom | Robot bergerak tanpa sentuhan |
| 8 | Hasil dicatat di sheet Modul E | Kolom Hasil Run (H) terisi |

> Checklist ini **bukan skor** — ini memastikan prosedur diikuti. Skor ada di Modul E.

---

## Bagian 5 — Kronologi kejadian

Isi tabel **Waktu | Kejadian | Catatan** selama atau setelah run.

**Contoh Run 1 — Tim 05, trial H2:**

| Waktu | Kejadian | Catatan |
|-------|----------|---------|
| 08:42 | Robot start dari posisi start | Navigasi ke Stand A |
| 08:44 | Pick K-M2 dari Stand A | Berhasil grip |
| 08:47 | Place K-M2 ke Rak 1 kiri | Marker merah — sukses |
| 08:49 | Pick K-H2 dari Stand B | Berhasil grip |
| 08:51 | Place K-H2 ke Rak 2 tengah | Kubus jatuh — gagal |
| 08:55 | Robot selesai / kembali home | 8/9 kubus sukses |

Baris kosong di bawah contoh bisa diisi jika ada kejadian tambahan (tabrakan obstacle, timeout, STOP darurat).

**Yang perlu dicatat jika terjadi:**
- Pick/place gagal atau jatuh
- Tabrakan obstacle
- Peserta minta RETRY
- Pelanggaran (sentuh laptop/robot)
- Intervensi CE

---

## Bagian 6 — Hasil singkat (otomatis)

Bagian ini **terisi rumus** dari Modul E — jangan ketik manual:

| Field | Sumber |
|-------|--------|
| Kubus berhasil ditempatkan benar | `=COUNTIF('Modul E'!H7:H15,1)` → tampil `8 / 9` |
| Skor Modul E run ini | `='Modul E'!J16` → tampil `53,33 / 60` |

**Isi manual juri:**

| Field | Isi |
|-------|-----|
| Pelanggaran (ya/tidak) | `ya` atau `tidak` |
| Jenis pelanggaran | Contoh: «Sentuh laptop setelah SIAP» — kosong jika tidak ada |

---

## Bagian 7 — Tanda tangan

| Pihak | Kapan |
|-------|-------|
| Juri | Setelah Log Run + Modul E lengkap |
| Peserta | Saksi run (biasanya kapten tim) |
| Chief Expert | Opsional — audit sampling |

---

## Bagian 8 — Urutan lengkap satu run (contoh runut)

Contoh **Run 1, Tim 05, H2 trial** — ikuti langkah berurutan:

| Langkah | Waktu | Sheet | Aksi |
|---------|-------|-------|------|
| 1 | 08:35 | Log Run | Isi header tim + Run No 1, Run ID `H2-T05-R1`, Trial |
| 2 | 08:35 | Log Run | Centang checklist **1** (briefing layout) |
| 3 | 08:40 | Log Run | Tim SIAP → isi Waktu SIAP `08:40`, centang **2** dan **3** |
| 4 | 08:41 | Undian | Acak kubus → isi baris 24, 28–36, obstacle |
| 5 | 08:41 | Log Run | Undian ulang = `1`, centang **4** |
| 6 | 08:42 | Log Run | START → Waktu START `08:42`, centang **5** dan **6** |
| 7 | 08:42–55 | — | Pantau run otonom → centang **7** |
| 8 | 08:55 | Modul E | Isi Hasil Run (1/0) kolom H untuk 9 kubus |
| 9 | 08:55 | Log Run | Waktu Selesai `08:55`, centang **8**, isi kronologi |
| 10 | 08:56 | Log Run | Cek Hasil singkat otomatis, isi pelanggaran, tanda tangan |

---

## Bagian 9 — Run RETRY (contoh)

**Situasi:** Tim 05 minta retry setelah K-H2 jatuh di Run 1.

| Aturan | Pelaksanaan |
|--------|-------------|
| Kubus di robot | K-H2 dikeluarkan ke arena/luar |
| Posisi kubus | **Tidak** dikembalikan ke posisi Run 1 |
| Programming | Diizinkan (waktu sesuai CE) |
| Setelah SIAP lagi | Acak ulang → **update Undian Kubus** |

**Log Run Run 2 (retry):**

| Field | Nilai |
|-------|-------|
| Run No | `2` |
| Run ID | `H2-T05-R2` |
| Retry | `1` |
| Retry No | `2` |
| Undian ulang | `1` |

Ulangi checklist 1–8 dari awal. Modul E diisi ulang Hasil Run untuk run retry ini (kebijakan CE: overwrite atau sheet terpisah).

---

## Bagian 10 — Hubungan dengan sheet lain

| Sheet | Hubungan dengan Log Run |
|-------|-------------------------|
| **Undian Kubus** | Run ID sama; posisi kubus harus sudah diisi sebelum START |
| **Modul E** | Hasil Run (H) harus terisi sebelum checklist item 8 = 1 |
| **Rekapitulasi** | Total Modul E masuk skor akhir tim (agregasi H2/H3 per CE) |

---

## Bagian 11 — Checklist juri sebelum lanjut tim berikutnya

- [ ] Semua field DATA RUN terisi
- [ ] Checklist 8 item = LENGKAP
- [ ] Kronologi minimal 3 baris (start, aktivitas, selesai)
- [ ] Modul E Hasil Run terisi → skor otomatis muncul di Hasil singkat
- [ ] Pelanggaran dicatat (jika ada)
- [ ] Tanda tangan juri + peserta

---

## Troubleshooting

| Masalah | Penyebab | Solusi |
|---------|----------|--------|
| Skor Log Run = 0 / kosong | Modul E kolom H belum diisi | Isi Hasil Run dulu |
| Checklist «BELUM» | Ada item OK ≠ 1 | Lengkapi centang |
| Kubus berhasil 0/9 padahal ada sukses | Rumus Modul E error | Cek Undian + Modul E |
| Run ID tidak cocok Undian | Salin typo | Samakan Run ID di semua sheet |

---

*Panduan ini untuk juri LKS Nasional 2026 — Robot Bergerak Otonom. Acuan resmi: SOP-JURI.md + Marking Scheme CIS.*
