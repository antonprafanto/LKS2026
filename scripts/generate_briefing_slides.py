"""Generate branded briefing PowerPoint for LKS 2026 judges."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"
OUT = ROOT / "docs" / "Briefing-Juri-LKS2026-Robot-Otonom.pptx"

NAVY = RGBColor(0x1F, 0x4E, 0x79)
ORANGE = RGBColor(0xED, 0x7D, 0x31)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x33, 0x33, 0x33)
LIGHT = RGBColor(0xF5, 0xF5, 0xF5)

LOGO_KEMEN = ASSETS / "logo-kemendikdasmen.png"
LOGO_LKS = ASSETS / "logo-lks-dikmen.png"
LOGO_GEO = ASSETS / "logo-lks-geometric.png"


def _blank_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def add_branding(slide, slide_width, title: str = "", footer: str = "LKS Nasional 2026 — Robot Bergerak Otonom | Kemendikdasmen") -> None:
    """Header bar, logos, optional title, footer stripe."""
    bar = slide.shapes.add_shape(1, 0, 0, slide_width, Inches(0.85))
    bar.fill.solid()
    bar.fill.fore_color.rgb = NAVY
    bar.line.fill.background()

    # orange accent line
    accent = slide.shapes.add_shape(1, 0, Inches(0.85), slide_width, Inches(0.06))
    accent.fill.solid()
    accent.fill.fore_color.rgb = ORANGE
    accent.line.fill.background()

    if LOGO_KEMEN.exists():
        slide.shapes.add_picture(str(LOGO_KEMEN), Inches(0.15), Inches(0.08), height=Inches(0.55))

    if LOGO_LKS.exists():
        slide.shapes.add_picture(str(LOGO_LKS), Inches(8.55), Inches(0.05), height=Inches(0.72))

    if title:
        tb = slide.shapes.add_textbox(Inches(2.5), Inches(0.12), Inches(5.8), Inches(0.65))
        p = tb.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

    # footer
    foot_bar = slide.shapes.add_shape(
        1, 0, Inches(7.05), slide_width, Inches(0.45)
    )
    foot_bar.fill.solid()
    foot_bar.fill.fore_color.rgb = GREEN
    foot_bar.line.fill.background()

    ft = slide.shapes.add_textbox(Inches(0.3), Inches(7.1), Inches(9.4), Inches(0.35))
    fp = ft.text_frame.paragraphs[0]
    fp.text = footer
    fp.font.size = Pt(10)
    fp.font.color.rgb = WHITE
    fp.alignment = PP_ALIGN.CENTER


def add_title_slide(prs: Presentation) -> None:
    slide = _blank_slide(prs)

    # background panel
    panel = slide.shapes.add_shape(1, 0, Inches(1.2), prs.slide_width, Inches(5.2))
    panel.fill.solid()
    panel.fill.fore_color.rgb = LIGHT
    panel.line.fill.background()

    if LOGO_KEMEN.exists():
        slide.shapes.add_picture(str(LOGO_KEMEN), Inches(0.5), Inches(0.15), height=Inches(0.7))
    if LOGO_LKS.exists():
        slide.shapes.add_picture(str(LOGO_LKS), Inches(4.1), Inches(1.5), height=Inches(1.6))
    if LOGO_GEO.exists():
        slide.shapes.add_picture(str(LOGO_GEO), Inches(7.2), Inches(1.55), height=Inches(1.5))

    tb = slide.shapes.add_textbox(Inches(0.6), Inches(3.4), Inches(8.8), Inches(1.2))
    p = tb.text_frame.paragraphs[0]
    p.text = "Briefing Juri"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.alignment = PP_ALIGN.CENTER

    tb2 = slide.shapes.add_textbox(Inches(0.6), Inches(4.5), Inches(8.8), Inches(1.5))
    tf = tb2.text_frame
    tf.word_wrap = True
    for i, line in enumerate([
        "Robot Bergerak Otonom (Autonomous Mobile Robotics)",
        "LKS Pendidikan Menengah Tingkat Nasional XXXIV 2026",
        "Kemendikdasmen — Pusat Prestasi Nasional",
    ]):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.text = line
        para.font.size = Pt(16 if i else 18)
        para.font.color.rgb = ORANGE if i else DARK
        para.alignment = PP_ALIGN.CENTER

    add_branding(slide, prs.slide_width, footer="Briefing Juri | Robot Bergerak Otonom | LKS 2026")


def add_section(prs: Presentation, title: str) -> None:
    slide = _blank_slide(prs)
    add_branding(slide, prs.slide_width)
    shape = slide.shapes.add_shape(1, Inches(0.8), Inches(2.8), Inches(8.4), Inches(1.3))
    shape.fill.solid()
    shape.fill.fore_color.rgb = NAVY
    shape.line.fill.background()
    box = slide.shapes.add_textbox(Inches(1), Inches(2.95), Inches(8), Inches(1))
    p = box.text_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER


def add_bullets(prs: Presentation, title: str, bullets: list[str], sub: str = "") -> None:
    slide = _blank_slide(prs)
    add_branding(slide, prs.slide_width, title=title)

    body = slide.shapes.add_textbox(Inches(0.55), Inches(1.15), Inches(8.9), Inches(5.7))
    tf = body.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(bullets):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.text = bullet
        para.font.size = Pt(15)
        para.font.color.rgb = DARK
        para.space_after = Pt(7)

    if sub:
        para = tf.add_paragraph()
        para.text = sub
        para.font.size = Pt(11)
        para.font.italic = True
        para.font.color.rgb = ORANGE


def add_numbered_steps(prs: Presentation, title: str, steps: list[tuple[str, str]]) -> None:
    """Step slide: (step label, description)."""
    slide = _blank_slide(prs)
    add_branding(slide, prs.slide_width, title=title)

    y = 1.15
    for num, (label, desc) in enumerate(steps, 1):
        badge = slide.shapes.add_shape(1, Inches(0.5), Inches(y), Inches(0.45), Inches(0.45))
        badge.fill.solid()
        badge.fill.fore_color.rgb = ORANGE
        badge.line.fill.background()
        bt = slide.shapes.add_textbox(Inches(0.5), Inches(y + 0.02), Inches(0.45), Inches(0.4))
        bp = bt.text_frame.paragraphs[0]
        bp.text = str(num)
        bp.font.bold = True
        bp.font.size = Pt(14)
        bp.font.color.rgb = WHITE
        bp.alignment = PP_ALIGN.CENTER

        box = slide.shapes.add_textbox(Inches(1.1), Inches(y - 0.05), Inches(8.3), Inches(0.9))
        tf = box.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = label
        p1.font.bold = True
        p1.font.size = Pt(14)
        p1.font.color.rgb = NAVY
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(12)
        p2.font.color.rgb = DARK
        y += 0.95


def add_demo_step(prs: Presentation, step_no: int, title: str, bullets: list[str], note: str = "") -> None:
    """Single Modul E demo step — one slide per phase."""
    slide = _blank_slide(prs)
    add_branding(slide, prs.slide_width, title=f"Demo Modul E — Langkah {step_no}")

    badge = slide.shapes.add_shape(1, Inches(0.5), Inches(1.1), Inches(0.7), Inches(0.7))
    badge.fill.solid()
    badge.fill.fore_color.rgb = ORANGE
    badge.line.fill.background()
    bt = slide.shapes.add_textbox(Inches(0.5), Inches(1.18), Inches(0.7), Inches(0.55))
    bp = bt.text_frame.paragraphs[0]
    bp.text = str(step_no)
    bp.font.size = Pt(22)
    bp.font.bold = True
    bp.font.color.rgb = WHITE
    bp.alignment = PP_ALIGN.CENTER

    tb = slide.shapes.add_textbox(Inches(1.4), Inches(1.05), Inches(8), Inches(0.8))
    tp = tb.text_frame.paragraphs[0]
    tp.text = title
    tp.font.size = Pt(22)
    tp.font.bold = True
    tp.font.color.rgb = NAVY

    body = slide.shapes.add_textbox(Inches(0.6), Inches(2.0), Inches(8.8), Inches(4.5))
    tf = body.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.text = f"• {b}"
        para.font.size = Pt(15)
        para.space_after = Pt(10)

    if note:
        nb = slide.shapes.add_shape(1, Inches(0.5), Inches(6.2), Inches(9), Inches(0.7))
        nb.fill.solid()
        nb.fill.fore_color.rgb = RGBColor(0xFF, 0xF3, 0xE0)
        nb.line.color.rgb = ORANGE
        nt = slide.shapes.add_textbox(Inches(0.65), Inches(6.28), Inches(8.7), Inches(0.55))
        np = nt.text_frame.paragraphs[0]
        np.text = f"⚠ Juri: {note}"
        np.font.size = Pt(12)
        np.font.bold = True
        np.font.color.rgb = ORANGE


def add_table_slide(prs: Presentation, title: str, headers: list[str], rows: list[list[str]]) -> None:
    slide = _blank_slide(prs)
    add_branding(slide, prs.slide_width, title=title)

    cols, row_count = len(headers), len(rows) + 1
    height = min(Inches(5.5), Inches(0.38 * row_count + 0.3))
    table = slide.shapes.add_table(
        row_count, cols, Inches(0.35), Inches(1.1), Inches(9.3), height
    ).table

    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        for p in cell.text_frame.paragraphs:
            p.font.bold = True
            p.font.size = Pt(10)
            p.font.color.rgb = WHITE

    for r, row in enumerate(rows, 1):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = val
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(10)


def main() -> None:
    if not LOGO_KEMEN.exists():
        import subprocess
        subprocess.run(["python", str(ROOT / "scripts" / "extract_branding_assets.py")], check=True)

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    add_title_slide(prs)

    add_bullets(prs, "Agenda Briefing", [
        "Identitas & tema proyek uji",
        "Struktur modul & bobot penilaian",
        "Sistem penilaian (Judgement & Measurement)",
        "Alur 3 hari lomba",
        "★ Demo penilaian Modul E (step-by-step)",
        "Arena, robot, K3, CIS",
    ])

    add_bullets(prs, "Identitas Lomba", [
        "Bidang: Autonomous Mobile Robotics (Robot Bergerak Otonom)",
        "Penyelenggara: Kemendikdasmen — Pusat Prestasi Nasional",
        "Tim: 2 peserta | Format: offline (luring) | ±15 jam / 3 hari",
        "Standar: WorldSkills WSOS Skill 23 — Mobile Robotics",
        "Tema 2026: Service Robot at Library",
        "Misi: robot mengambil kubus (buku) dan menempatkannya ke rak sesuai marker warna — sepenuhnya otonom",
    ])

    add_table_slide(
        prs, "Bobot Modul",
        ["Modul", "Kriteria", "Hari", "Bobot", "Metode"],
        [
            ["A", "Organisasi & Manajemen Kerja", "H1–H3", "8%", "Judgement"],
            ["B", "Jurnal / Laporan Teknis", "H1", "10%", "J + M"],
            ["C", "Perakitan Robot (L-Channel)", "H1", "10%", "Judgement"],
            ["D", "Gerakan Dasar & Manajemen Objek", "H2", "12%", "Measurement"],
            ["E", "Performansi Otonom", "H2–H3", "60%", "Measurement"],
        ],
    )

    add_bullets(prs, "Sistem Penilaian", [
        "Judgement (J): skala 0–3 — proses, kerapian, kolaborasi",
        "Measurement (M): biner 1/0 — akurasi gerak & sukses misi",
        "3 juri untuk J → rata-rata; selisih ≥2 → ulangi penilaian",
        "Input ke CIS → skala 700; Medallion of Excellence jika >700",
        "60% skor dari Modul E — penilaian Anda sangat menentukan juara",
    ])

    # ========== DEMO MODUL E ==========
    add_section(prs, "Demo Penilaian Modul E\n(Step-by-Step untuk Juri)")

    add_demo_step(prs, 1, "Briefing Layout Arena", [
        "Jelaskan ke peserta: posisi START, orientasi 3 rak, gate tengah, obstacle",
        "Tunjukkan marker warna di atas setiap slot rak (dapat berubah per soal)",
        "Pastikan penerangan arena menyala (lampu 100 W × 2)",
        "Catat konfigurasi marker hari ini di lembar undian",
    ], "Jangan mulai run sebelum peserta paham layout.")

    add_demo_step(prs, 2, "Programming & Persiapan", [
        "Peserta memprogram navigasi, visi, dan pick-place sesuai layout",
        "Peserta BOLEH menyentuh laptop pada tahap ini",
        "Pastikan tombol START di robot dan tombol STOP merah siap",
        "Juri observasi Modul A (kolaborasi, 5S) selama sesi ini",
    ])

    add_demo_step(prs, 3, "Sinyal SIAP", [
        "Peserta menyatakan SIAP ketika program siap dijalankan",
        "Setelah SIAP → peserta TIDAK BOLEH menyentuh laptop lagi",
        "Juri verifikasi: tangan peserta menjauh dari laptop",
        "Siapkan lembar undian untuk pengacakan kubus",
    ], "Sentuh laptop setelah SIAP = pelanggaran (run batal).")

    add_demo_step(prs, 4, "Pengacakan Posisi Kubus", [
        "Juri/acara mengacak posisi 9 kubus (rak + stand)",
        "Beberapa kubus sengaja di posisi SALAH — perlu dipindahkan robot",
        "Isi templates/lembar-undian-kubus.csv atau Excel Modul E",
        "Peserta TIDAK ikut mengatur posisi setelah SIAP",
    ], "Setelah SIAP, acak ulang — bukan hanya sebelum run pertama.")

    add_demo_step(prs, 5, "Sinyal START & Run Otonom", [
        "Juri beri sinyal START",
        "Peserta menekan tombol START pada robot (bukan keyboard)",
        "Robot bergerak, navigasi, pick & place secara otonom",
        "Juri catat kronologi di lembar-run-otonom.csv",
    ], "Peserta tidak boleh menyentuh robot setelah START.")

    add_demo_step(prs, 6, "Penilaian Per Kubus (Measurement)", [
        "Kubus SUKSES = di slot rak dengan marker warna yang cocok + stabil",
        "Isi kolom Hasil (1/0) di Excel Modul E untuk tiap kubus",
        "Default: 6,67 poin per kubus × 9 = 60 poin (sesuaikan CIS jika beda)",
        "Kubus sudah benar sejak awal — konfirmasi ke CE apakah dapat poin",
    ])

    add_table_slide(
        prs, "Contoh Penilaian Run (Latihan)",
        ["Kubus", "Awal", "Target", "Hasil", "Skor"],
        [
            ["K-M2 Merah□", "Stand A", "Slot marker merah", "Berhasil", "6,67"],
            ["K-H2 Hijau□", "Stand B", "Slot marker hijau", "Gagal (jatuh)", "0"],
            ["K-B3 Biru△", "Stand C", "Slot marker biru", "Berhasil", "6,67"],
            ["K-M1 Merah○", "Rak 1 (benar)", "—", "Sudah benar*", "6,67*"],
            ["...", "...", "...", "...", "..."],
        ],
    )

    add_demo_step(prs, 7, "Prosedur RETRY", [
        "Peserta minta RETRY → boleh ubah program",
        "Kubus di badan robot HARUS dikeluarkan ke luar arena",
        "Posisi kubus TIDAK dikembalikan ke kondisi awal run",
        "Setelah SIAP lagi → acak ulang posisi kubus → START baru",
    ])

    add_demo_step(prs, 8, "Rekap & Input CIS", [
        "Kumpulkan lembar undian + run + Modul E per tim",
        "Transfer skor ke rekapitulasi-skor.csv / sheet Rekapitulasi",
        "Chief Expert review outlier & pelanggaran",
        "Input skor final ke CIS — penentuan juara & Medallion",
    ])

    add_bullets(prs, "Latihan Juri — Kerjakan Bersama", [
        "1. Buka docs/SKENARIO-SOAL-CONTOH.md — skenario Tim 05",
        "2. Isi lembar undian untuk Run 1",
        "3. Mainkan peran: SIAP → acak → START → catat 9 kubus",
        "4. Hitung skor di Excel Modul E",
        "5. Simulasikan RETRY + pelanggaran sentuh laptop",
        "6. Diskusi dengan CE: konflik slot, bobot CIS",
    ], "Materi: github.com/antonprafanto/LKS2026")

    # ========== remaining slides ==========
    add_section(prs, "Modul A–D — Ringkasan")

    add_bullets(prs, "Modul D (12 poin, 15 menit/tim)", [
        "D1: gerak 4 arah | D2: stop 50–100 mm dari dinding",
        "D3: line follower U | D4–D6: deteksi warna di monitor",
        "D7–D8: O-merah, O-hijau | D9–D10: pick rak & place stand",
        "Tidak menguji bentuk persegi/segitiga secara eksplisit",
    ])

    add_bullets(prs, "Arena & Robot", [
        "Arena ~4040×2040 mm, 3 rak, 9 kubus, 10 obstacle, 1 gate",
        "Controller: Raspberry Pi / myRIO | Python, C++, LabVIEW",
        "Baterai ≤12V | tombol STOP merah wajib",
        "30% rule H3: posisi objek, urutan, marker dapat berubah",
    ])

    add_bullets(prs, "K3 & Klarifikasi CE", [
        "K3: no air terbuka dekat listrik; reset obstacle jatuh",
        "Klarifikasi: ukuran arena, bobot per kubus CIS, max RETRY",
        "Dokumen resmi: smk.pusatprestasinasional.kemdikdasmen.go.id",
        "SOP: docs/SOP-JURI.md | Excel: templates/LKS2026-Penilaian-Juri.xlsx",
    ], "Terima kasih — selamat bertugas sebagai juri LKS 2026!")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    print(f"Created: {OUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
