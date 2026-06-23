# Skenario Soal Contoh — Latihan Juri
## Robot Bergerak Otonom LKS 2026 | Tema: Service Robot at Library

**Tujuan:** Melatih juri mengisi template penilaian, menjalankan SOP SIAP/START/RETRY, dan menghitung skor Modul E.

**Level:** Trial H2 (bukan final). Variasi 30% belum diterapkan penuh.

---

## 1. Konfigurasi Arena

| Elemen | Setting skenario ini |
|--------|----------------------|
| Arena | Set 2 (4040 × 2040 mm) |
| Start robot | Sudut kiri-bawah (dekat gate) |
| Gate | Tengah partisi, robot harus melewati |
| Rak 1 | Kiri arena — marker: **Merah, Hijau, Biru** (slot kiri→kanan) |
| Rak 2 | Atas-tengah — marker: **Biru, Merah, Hijau** |
| Rak 3 | Kanan arena — marker: **Hijau, Biru, Merah** |
| Obstacle | 6 buah diletakkan (lihat peta di bawah) |

### Peta obstacle (kasar)

```
    [Rak 2]
       |
  O    |    O
-------GATE-------
  O         O
       |
[Rak 1]     [Rak 3]
    S = START
```

- **O** = obstacle (10 tersedia, pakai 6 untuk latihan)
- **GATE** = 500 mm, tengah arena

---

## 2. Undian Posisi Awal Kubus

Beberapa kubus **sengaja salah posisi** — robot harus memindahkan ke rak/slot yang sesuai marker warna **di atas slot**, bukan warna kubus sembarang.

| ID | Warna | Bentuk | Posisi Awal | Benar? | Target |
|----|-------|--------|-------------|--------|--------|
| K-M1 | Merah | Lingkaran | Rak 1, slot kiri | ✅ Sudah benar | — (tidak perlu dipindah) |
| K-M2 | Merah | Persegi | Stand A (depan rak 1) | ❌ | Rak 1 slot kiri **atau** slot dengan marker merah di rak mana pun* |
| K-M3 | Merah | Segitiga | Rak 2, slot tengah | ❌ | Rak dengan marker merah |
| K-H1 | Hijau | Lingkaran | Rak 3, slot kiri | ✅ | — |
| K-H2 | Hijau | Persegi | Stand B | ❌ | Slot marker hijau |
| K-H3 | Hijau | Segitiga | Rak 1, slot kanan | ❌ | Slot marker hijau |
| K-B1 | Biru | Lingkaran | Rak 2, slot kiri | ✅ | — |
| K-B2 | Biru | Persegi | Rak 3, slot tengah | ❌ | Slot marker biru |
| K-B3 | Biru | Segitiga | Stand C | ❌ | Slot marker biru |

\* **Keputusan Chief Expert:** Apakah kubus harus ke rak **mana pun** yang marker-nya cocok dengan warna kubus, atau ke **rak tertentu** di soal. Untuk latihan ini: **warna kubus harus cocok dengan marker warna di slot rak**.

### Kubus yang harus dipindahkan (6 buah)

K-M2, K-M3, K-H2, K-H3, K-B2, K-B3

### Jawaban target (kunci juri)

| ID | Target Rak | Target Slot (marker) |
|----|------------|----------------------|
| K-M2 | Rak 1 | Kiri (merah) — **sudah ada K-M1** → gunakan Rak 2 tengah (merah) atau Rak 3 kanan (merah) |
| K-M3 | Rak 2 | Tengah (merah) |
| K-H2 | Rak 1 | Tengah (hijau) |
| K-H3 | Rak 3 | Kiri (hijau) — **sudah ada K-H1** → Rak 1 kanan (hijau) atau Rak 2 kanan (hijau) |
| K-B2 | Rak 1 | Kanan (biru) |
| K-B3 | Rak 2 | Kiri (biru) — **sudah ada K-B1** → konflik slot: juri catat **penempatan di slot marker biru kosong** |

> Skenario ini sengaja membuat **konflik slot** agar juri berlatih menilai: sukses = kubus di slot dengan marker warna yang cocok **dan** slot tidak double-booked.

---

## 3. Alur Run Latihan (Modul E)

### Run 1 — Trial normal

| Langkah | Waktu | Aksi |
|---------|-------|------|
| 1 | 08:35 | Briefing layout ke Tim **05** |
| 2 | 08:40 | Tim programming selesai, nyatakan **SIAP** |
| 3 | 08:41 | Juri **acak ulang**: tukar posisi K-H2 ↔ K-B3 |
| 4 | 08:42 | Sinyal **START** — tim tekan tombol di robot |
| 5 | 08:42–08:55 | Run otonom |
| 6 | 08:55 | Catat hasil per kubus di `modul-e-otonom.csv` |

### Hasil contoh Run 1 (untuk latihan penilaian)

| ID | Dipindah? | Berhasil? | Skor (6,67) |
|----|-----------|-----------|-------------|
| K-M2 | Ya | Ya | 6,67 |
| K-M3 | Ya | Ya | 6,67 |
| K-H2 | Ya | Tidak (jatuh) | 0 |
| K-H3 | Ya | Ya | 6,67 |
| K-B2 | Ya | Ya | 6,67 |
| K-B3 | Ya | Ya | 6,67 |
| K-M1, K-H1, K-B1 | Tidak | Sudah benar | 6,67 each |

**Skor Run 1:** 6 × 6,67 + 3 × 6,67 (yang sudah benar) = ... 

**Catatan penilaian:** Kubus yang **sudah benar sejak awal** — putuskan dengan CE apakah dapat poin penuh atau hanya kubus yang **dipindahkan** yang dinilai. Template repo memberi **9 baris**; isi **1** jika kondisi sukses terpenuhi per baris.

**Skor contoh (semua 9 baris dinilai):** 8 sukses × 6,67 ≈ **53,36** / 60

---

## 4. Run 2 — Latihan RETRY

**Situasi:** Tim 05 minta RETRY setelah K-H2 jatuh.

| Aturan | Pelaksanaan |
|--------|-------------|
| Kubus di robot | K-H2 dikeluarkan ke luar arena |
| Posisi kubus | **Tidak** dikembalikan ke Run 1 — K-H2 tetap di lantai di mana jatuh |
| Ubah program | Diizinkan 10 menit |
| Setelah SIAP | Posisi **diacak ulang** lagi |

**Pelanggaran latihan:** Peserta menyentuh laptop setelah SIAP → juri beri peringatan; jika tetap, **run batal** (skor 0 run tersebut).

---

## 5. Latihan Modul D (cuplikan)

Gunakan **1 tim** yang sama. Centang hasil:

| ID | Hasil latihan (contoh) |
|----|------------------------|
| D1a–D1d | Semua 1 |
| D2a | 1 |
| D2b | 0 (stop terlalu jauh) |
| D3 | 1 |
| D4–D6 | 1 |
| D7 | 1 |
| D8 | 0 (teks tidak muncul) |
| D9–D10 | 1 |

**Skor Modul D:** 12 − 0,6 − 1,0 = **10,4** / 12

---

## 6. Latihan Modul A (cuplikan)

| Sub | J1 | J2 | J3 | Rata | Skor |
|-----|----|----|-----|------|------|
| A2 | 2 | 3 | 2 | 2,33 | 1,56 |
| A3 | 2 | 2 | 1 | 1,67 | 1,11 |

Selisih A2: max−min = 1 → **tidak perlu** ulang.  
Selisih A3: jika J3=0 → selisih 2 → **ulangi penilaian A3**.

---

## 7. Checklist Juri Setelah Latihan

- [ ] Semua lembar terisi (undian, run, modul D/E)
- [ ] Tanda tangan juri + peserta (run otonom)
- [ ] Rekap masuk `rekapitulasi-skor.csv` atau sheet Excel
- [ ] Diskusi dengan CE: konflik slot, kubus sudah benar, bobot CIS

---

## 8. Variasi untuk Sesi Latihan 2 (H3 simulation)

Terapkan **30% rule:**

1. Tukar marker Rak 2 menjadi: **Hijau, Biru, Merah**
2. Tambah 4 obstacle
3. Ubah start ke sudut kanan-bawah
4. Urutan pengiriman: utamakan kubus **bentuk segitiga** dulu (jika CE menetapkan)

---

*Skenario ini untuk **latihan internal juri**, bukan soal resmi lomba.*
