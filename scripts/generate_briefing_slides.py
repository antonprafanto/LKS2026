"""Generate briefing PowerPoint for LKS 2026 Autonomous Mobile Robotics judges."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

OUT = Path(__file__).resolve().parent.parent / "docs" / "Briefing-Juri-LKS2026-Robot-Otonom.pptx"

NAVY = RGBColor(0x1F, 0x4E, 0x79)
ORANGE = RGBColor(0xED, 0x7D, 0x31)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x33, 0x33, 0x33)


def add_title_slide(prs: Presentation, title: str, subtitle: str) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(1.4))
    shape.fill.solid()
    shape.fill.fore_color.rgb = NAVY
    shape.line.fill.background()

    box = slide.shapes.add_textbox(Inches(0.6), Inches(1.8), Inches(8.8), Inches(1.5))
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = NAVY

    box2 = slide.shapes.add_textbox(Inches(0.6), Inches(3.3), Inches(8.8), Inches(1.2))
    tf2 = box2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = subtitle
    p2.font.size = Pt(18)
    p2.font.color.rgb = ORANGE


def add_section(prs: Presentation, title: str) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(1, 0, Inches(2.5), prs.slide_width, Inches(1.2))
    shape.fill.solid()
    shape.fill.fore_color.rgb = NAVY
    shape.line.fill.background()
    box = slide.shapes.add_textbox(Inches(0.5), Inches(2.55), Inches(9), Inches(1))
    p = box.text_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER


def add_bullets(prs: Presentation, title: str, bullets: list[str], sub: str = "") -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    bar = slide.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(0.9))
    bar.fill.solid()
    bar.fill.fore_color.rgb = NAVY
    bar.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(9), Inches(0.7))
    p = tb.text_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = WHITE

    body = slide.shapes.add_textbox(Inches(0.6), Inches(1.1), Inches(8.8), Inches(5.5))
    tf = body.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(bullets):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.text = bullet
        para.level = 0
        para.font.size = Pt(16)
        para.font.color.rgb = DARK
        para.space_after = Pt(8)

    if sub:
        para = tf.add_paragraph()
        para.text = sub
        para.font.size = Pt(12)
        para.font.italic = True
        para.font.color.rgb = ORANGE


def add_table_slide(prs: Presentation, title: str, headers: list[str], rows: list[list[str]]) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bar = slide.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(0.9))
    bar.fill.solid()
    bar.fill.fore_color.rgb = NAVY
    bar.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(9), Inches(0.7))
    tb.text_frame.paragraphs[0].text = title
    tb.text_frame.paragraphs[0].font.size = Pt(22)
    tb.text_frame.paragraphs[0].font.bold = True
    tb.text_frame.paragraphs[0].font.color.rgb = WHITE

    cols, row_count = len(headers), len(rows) + 1
    table = slide.shapes.add_table(row_count, cols, Inches(0.4), Inches(1.1), Inches(9.2), Inches(0.4 * row_count + 0.5)).table

    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        for p in cell.text_frame.paragraphs:
            p.font.bold = True
            p.font.size = Pt(11)
            p.font.color.rgb = WHITE

    for r, row in enumerate(rows, 1):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = val
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(11)


def main() -> None:
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    add_title_slide(
        prs,
        "Briefing Juri",
        "Robot Bergerak Otonom (Autonomous Mobile Robotics)\nLKS Nasional XXXIV 2026",
    )

    add_bullets(prs, "Agenda", [
        "Identitas & tema proyek uji",
        "Struktur modul & bobot penilaian",
        "Sistem penilaian (Judgement & Measurement)",
        "Alur 3 hari lomba",
        "SOP Modul A–E (fokus run otonom)",
        "Arena, robot, K3",
        "CIS & hal yang perlu diklarifikasi CE",
    ])

    add_bullets(prs, "Identitas Lomba", [
        "Bidang: Autonomous Mobile Robotics (Robot Bergerak Otonom)",
        "Tim: 2 peserta | Format: offline (luring)",
        "Total waktu kerja: ±15 jam (3 hari)",
        "Standar acuan: WorldSkills WSOS Skill 23 — Mobile Robotics",
        "Tema 2026: Service Robot at Library",
        "Robot = asisten perpustakaan: ambil kubus (buku) → tempatkan ke rak sesuai marker warna",
    ])

    add_table_slide(
        prs,
        "Bobot Modul Penilaian",
        ["Modul", "Kriteria", "Hari", "Bobot", "Metode"],
        [
            ["A", "Organisasi & Manajemen Kerja", "H1–H3", "8%", "Judgement"],
            ["B", "Jurnal / Laporan Teknis", "H1", "10%", "J + M"],
            ["C", "Perakitan Robot (L-Channel)", "H1", "10%", "Judgement"],
            ["D", "Gerakan Dasar & Manajemen Objek", "H2", "12%", "Measurement"],
            ["E", "Performansi Otonom", "H2–H3", "60%", "Measurement"],
        ],
    )

    add_bullets(prs, "Dua Metode Penilaian", [
        "Judgement (J): skala 0–3 — proses, kerapian, kolaborasi, kualitas dokumentasi",
        "Measurement (M): biner 1/0 — akurasi gerak, sukses misi, penempatan kubus",
        "Komposisi total: Judgement 26 poin + Measurement 74 poin = 100",
        "3 juri untuk Judgement → rata-rata; selisih ≥2 → ulangi penilaian",
        "Skor dimasukkan ke CIS → konversi skala 700; Medallion of Excellence jika >700",
    ])

    add_table_slide(
        prs,
        "Jadwal Operasional",
        ["Hari", "Waktu utama", "Kegiatan penilaian"],
        [
            ["H-1", "TM + 2 jam", "Familiarisasi arena & alat"],
            ["H1", "08.30–12.00", "Modul C: perakitan L-Channel (maks. 2 jam)"],
            ["H1", "Sepanjang hari", "Modul A, B"],
            ["H2", "15 menit/tim", "Modul D — gerakan dasar"],
            ["H2–H3", "Trial + marking", "Modul E — performa otonom (60%)"],
        ],
    )

    add_section(prs, "Modul A — Organisasi & Manajemen Kerja")

    add_bullets(prs, "Modul A (8 poin)", [
        "A1: Efisiensi penjadwalan (mekanik / elektronik / debugging)",
        "A2: Kolaborasi tim & pembagian peran",
        "A3: Kerapian area kerja — penerapan 5S",
        "A4: Restorasi area kerja setelah selesai",
        "Dinilai H1, H2, H3 secara observasi berkelanjutan",
    ], "Template: modul-a-organisasi.csv | Sheet Modul A (Excel)")

    add_section(prs, "Modul B & C — Hari Pertama")

    add_bullets(prs, "Modul B — Jurnal Teknis (10 poin)", [
        "Judgement 8 poin: fabrikasi, wiring, integrasi, kalibrasi, optimasi, kelengkapan dokumen",
        "Measurement 2 poin: ketepatan pengumpulan sesuai deadline",
        "Maks. 4 lembar isi (8 halaman A4) + sampul",
        "Wajib: desain mekanik, diagram blok, BOM, algoritma navigasi & visi, flowchart, data kalibrasi",
    ])

    add_bullets(prs, "Modul C — Perakitan Robot (10 poin)", [
        "WAJIB: L-Channel panitia sebagai penyangga utama sistem angkat",
        "Integrasi lift/gripper + rewiring di lokasi — batas 2 jam",
        "Motor kompatibel: JGA-25 / PG36 gear head DC",
        "Inspeksi: L-Channel benar, kabel aman, tombol STOP merah berfungsi",
        "CAD: drive.google.com (folder L-Channel resmi)",
    ])

    add_section(prs, "Modul D — Gerakan Dasar (H2)")

    add_bullets(prs, "Modul D (12 poin, 15 menit/tim)", [
        "D1: gerak maju/mundur/kiri/kanan (200–300 mm)",
        "D2: deteksi dinding, stop 50–100 mm",
        "D3: ikuti garis hitam bentuk U",
        "D4–D6: deteksi warna merah/hijau/biru di monitor",
        "D7–D8: deteksi lingkaran merah/hijau (O-merah, O-hijau)",
        "D9–D10: pick dari rak & place ke standbox — otonom",
        "Catatan: bentuk persegi/segitiga tidak diuji eksplisit di Modul D",
    ])

    add_section(prs, "Modul E — Performa Otonom (60%)")

    add_bullets(prs, "SOP Run Otonom — WAJIB", [
        "1. Briefing layout arena ke peserta",
        "2. Peserta selesai programming → nyatakan SIAP",
        "3. Setelah SIAP: peserta TIDAK BOLEH sentuh laptop",
        "4. Juri ACAK ULANG posisi kubus",
        "5. Sinyal START → peserta tekan tombol START di robot",
        "6. Robot berjalan sepenuhnya otonom",
        "7. Catat kubus berhasil di slot marker yang benar",
    ])

    add_bullets(prs, "Aturan RETRY", [
        "Peserta boleh minta RETRY dan ubah program",
        "Posisi kubus TIDAK dikembalikan ke kondisi awal",
        "Kubus di badan robot harus dikeluarkan ke luar arena",
        "Setelah SIAP lagi → posisi kubus diacak ulang",
        "Pelanggaran: sentuh laptop setelah SIAP / sentuh robot setelah START → run batal",
    ])

    add_bullets(prs, "Arena & Aksesoris", [
        "Arena: multiplex 18 mm, putih matte (~4040×2040 mm — klarifikasi CE)",
        "3 rak (marker warna di atas slot), 9 kubus PLA 65³ mm",
        "10 obstacle, 1 gate (aluminium 2020), reference line isolasi hitam",
        "Posisi kubus & marker dapat berubah (30% rule di H3)",
        "4 set arena | penerangan 100 W × 2 per lapangan",
    ])

    add_bullets(prs, "Spesifikasi Robot Peserta", [
        "Controller: Raspberry Pi atau myRIO",
        "Bahasa: Python, C++, atau LabVIEW",
        "Sensor: max 1 kamera, 4 IR, 2 US, 4 line, 1 lidar, 1 IMU",
        "Aktuator: max 4 motor DC, 3 servo, 2 continuous servo",
        "Baterai ≤12V nominal | tombol STOP merah wajib",
        "Dimensi & berat: tidak dibatasi",
    ])

    add_bullets(prs, "K3 — Prioritas Juri", [
        "Larang air mineral terbuka dekat komponen listrik",
        "Verifikasi tombol STOP sebelum setiap run",
        "Reset obstacle yang jatuh sebelum tim berikutnya",
        "Area kerja: jalur evakuasi jelas, kabel rapi",
        "Robot tidak boleh membahayakan peserta, juri, atau penonton",
    ])

    add_bullets(prs, "Klarifikasi ke Chief Expert", [
        "Ukuran arena: 4040×2040 (Kisi-kisi) vs 2400×4800 (Daftar Bahan)",
        "Bobot per kubus Modul E di CIS vs template (6,67×9)",
        "Apakah kubus sudah benar di awal mendapat poin penuh?",
        "Jumlah maksimum RETRY per tim",
        "Agregasi skor multi-run H2 vs H3 (trial/final)",
    ])

    add_bullets(prs, "Materi Pendukung Juri", [
        "SOP lengkap: docs/SOP-JURI.md",
        "Template Excel: templates/LKS2026-Penilaian-Juri.xlsx",
        "Skenario latihan: docs/SKENARIO-SOAL-CONTOH.md",
        "Repo GitHub: github.com/antonprafanto/LKS2026",
        "Dokumen resmi: smk.pusatprestasinasional.kemdikdasmen.go.id",
        "Marking Scheme Excel CIS — sumber kebenaran poin detail",
    ], "Terima kasih — semoga lancar sebagai juri!")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    print(f"Created: {OUT}")


if __name__ == "__main__":
    main()
