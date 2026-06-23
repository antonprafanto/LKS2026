"""Generate Excel scoring templates with auto formulas for LKS 2026 judges."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = Path(__file__).resolve().parent.parent / "templates" / "LKS2026-Penilaian-Juri.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True)
THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header_row(ws, row: int, cols: int) -> None:
    for col in range(1, cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def set_widths(ws, widths: dict[int, float]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def sheet_modul_a(wb: Workbook) -> None:
    ws = wb.active
    ws.title = "Modul A"
    ws["A1"] = "MODUL A — Organisasi & Manajemen Kerja (8 poin, Judgement)"
    ws.merge_cells("A1:I1")
    ws["A1"].font = Font(bold=True, size=14)

    ws["A3"], ws["C3"], ws["E3"], ws["G3"] = "Tim No:", "Nama Tim:", "Kontingen:", "Tanggal:"
    ws["A4"], ws["C4"], ws["E4"] = "Juri 1:", "Juri 2:", "Juri 3:"

    headers = [
        "Sub", "Kriteria", "Max", "Juri 1", "Juri 2", "Juri 3",
        "Rata-rata", "Skor Akhir", "Selisih", "Catatan",
    ]
    start = 6
    for i, h in enumerate(headers, 1):
        ws.cell(row=start, column=i, value=h)
    style_header_row(ws, start, len(headers))

    rows = [
        ("A1", "Efisiensi Penjadwalan Kerja", 2.0),
        ("A2", "Kolaborasi Tim & Manajemen Tugas", 2.0),
        ("A3", "Organisasi & Kerapian Area Kerja (5S)", 2.0),
        ("A4", "Restorasi Area Kerja", 2.0),
    ]
    for idx, (sub, name, mx) in enumerate(rows, start + 1):
        r = idx
        ws.cell(r, 1, sub)
        ws.cell(r, 2, name)
        ws.cell(r, 3, mx)
        ws.cell(r, 7, f"=IF(COUNT(D{r}:F{r})=0,\"\",AVERAGE(D{r}:F{r}))")
        ws.cell(r, 8, f"=IF(G{r}=\"\",0,G{r}*C{r}/3)")
        ws.cell(r, 9, f"=IF(COUNT(D{r}:F{r})<2,\"\",MAX(D{r}:F{r})-MIN(D{r}:F{r}))")
        for c in range(4, 7):
            ws.cell(r, c).number_format = "0.0"

    total_row = start + 1 + len(rows)
    ws.cell(total_row, 2, "TOTAL MODUL A")
    ws.cell(total_row, 3, 8.0)
    ws.cell(total_row, 8, f"=SUM(H{start+1}:H{total_row-1})")
    ws.cell(total_row, 2).font = Font(bold=True)

    set_widths(ws, {1: 6, 2: 38, 3: 8, 4: 10, 5: 10, 6: 10, 7: 12, 8: 12, 9: 10, 10: 24})


def sheet_modul_d(wb: Workbook) -> None:
    ws = wb.create_sheet("Modul D")
    ws["A1"] = "MODUL D — Gerakan Dasar & Manajemen Objek (12 poin, Measurement)"
    ws.merge_cells("A1:G1")
    ws["A1"].font = Font(bold=True, size=14)

    ws["A3"], ws["C3"], ws["E3"] = "Tim No:", "Nama Tim:", "Tanggal:"
    ws["A4"] = "Juri Penilai:"

    headers = ["ID", "Evaluasi", "Indikator", "Bobot", "Hasil (1/0)", "Skor", "Catatan"]
    start = 6
    for i, h in enumerate(headers, 1):
        ws.cell(start, i, h)
    style_header_row(ws, start, len(headers))

    items = [
        ("D1a", "Gerak maju 200-300 mm", "Observasi", 0.5),
        ("D1b", "Gerak mundur 200-300 mm", "Observasi", 0.5),
        ("D1c", "Gerak kanan 200-300 mm", "Observasi", 0.5),
        ("D1d", "Gerak kiri 200-300 mm", "Observasi", 0.5),
        ("D2a", "Maju stop 50-100 mm dari dinding", "Observasi", 0.3),
        ("D2b", "Kanan stop 50-100 mm dari dinding", "Observasi", 0.6),
        ("D3", "Ikuti garis U", "Line follower", 1.0),
        ("D4", "Deteksi merah → teks merah", "Monitor", 1.0),
        ("D5", "Deteksi hijau → teks hijau", "Monitor", 1.0),
        ("D6", "Deteksi biru → teks biru", "Monitor", 1.0),
        ("D7", "Deteksi O-merah", "Monitor", 1.0),
        ("D8", "Deteksi O-hijau", "Monitor", 1.0),
        ("D9", "Pick kubus dari rak otonom", "Observasi arm", 2.0),
        ("D10", "Place kubus ke standbox", "Observasi arm", 2.0),
    ]
    for i, (id_, ev, ind, bobot) in enumerate(items, start + 1):
        ws.cell(i, 1, id_)
        ws.cell(i, 2, ev)
        ws.cell(i, 3, ind)
        ws.cell(i, 4, bobot)
        ws.cell(i, 6, f"=IF(E{i}=\"\",0,E{i}*D{i})")

    tr = start + 1 + len(items)
    ws.cell(tr, 2, "TOTAL")
    ws.cell(tr, 4, 12.0)
    ws.cell(tr, 6, f"=SUM(F{start+1}:F{tr-1})")
    ws.cell(tr, 2).font = Font(bold=True)

    set_widths(ws, {1: 8, 2: 34, 3: 16, 4: 8, 5: 12, 6: 10, 7: 24})


def sheet_modul_e(wb: Workbook) -> None:
    ws = wb.create_sheet("Modul E")
    ws["A1"] = "MODUL E — Performansi Otonom (60 poin, Measurement)"
    ws.merge_cells("A1:J1")
    ws["A1"].font = Font(bold=True, size=14)

    ws["A3"], ws["C3"], ws["E3"], ws["G3"] = "Tim No:", "Run ID:", "Hari:", "Trial/Final:"
    ws["A4"], ws["C4"] = "Arena No:", "Waktu START:"

    headers = [
        "No", "ID Kubus", "Warna", "Bentuk", "Posisi Awal", "Target Rak",
        "Target Marker", "Hasil (1/0)", "Bobot", "Skor", "Catatan",
    ]
    start = 6
    for i, h in enumerate(headers, 1):
        ws.cell(start, i, h)
    style_header_row(ws, start, len(headers))

    for n in range(1, 10):
        r = start + n
        ws.cell(r, 1, n)
        ws.cell(r, 9, 6.67)
        ws.cell(r, 10, f"=IF(H{r}=\"\",0,H{r}*I{r})")

    tr = start + 10
    ws.cell(tr, 5, "TOTAL MODUL E")
    ws.cell(tr, 9, 60.0)
    ws.cell(tr, 10, f"=SUM(J{start+1}:J{tr-1})")
    ws.cell(tr, 5).font = Font(bold=True)

    ws.cell(tr + 2, 1, "Pelanggaran (1=ya):")
    pelanggaran = [
        "Sentuh laptop setelah SIAP",
        "Sentuh robot setelah START",
        "Tidak start via tombol robot",
        "Place manual oleh peserta",
    ]
    for i, p in enumerate(pelanggaran, tr + 3):
        ws.cell(i, 1, p)
        ws.cell(i, 2, 0)

    set_widths(ws, {1: 5, 2: 10, 3: 10, 4: 10, 5: 14, 6: 12, 7: 14, 8: 12, 9: 8, 10: 8, 11: 20})


def sheet_rekap(wb: Workbook) -> None:
    ws = wb.create_sheet("Rekapitulasi")
    ws["A1"] = "REKAPITULASI SKOR — Robot Bergerak Otonom LKS 2026"
    ws.merge_cells("A1:F1")
    ws["A1"].font = Font(bold=True, size=14)

    ws["A3"], ws["C3"], ws["E3"] = "Tim No:", "Nama Tim:", "Kontingen:"

    headers = ["Modul", "Kriteria", "Bobot Max", "Skor", "Persen", "Metode"]
    start = 5
    for i, h in enumerate(headers, 1):
        ws.cell(start, i, h)
    style_header_row(ws, start, len(headers))

    data = [
        ("A", "Organisasi & Manajemen Kerja", 8, "=\'Modul A\'!H11", "J"),
        ("B", "Jurnal Teknis", 10, 0, "J+M"),
        ("C", "Perakitan Robot", 10, 0, "J"),
        ("D", "Gerakan Dasar", 12, "=\'Modul D\'!F21", "M"),
        ("E", "Performa Otonom", 60, "=\'Modul E\'!J16", "M"),
    ]
    for i, (mod, crit, mx, formula, met) in enumerate(data, start + 1):
        ws.cell(i, 1, mod)
        ws.cell(i, 2, crit)
        ws.cell(i, 3, mx)
        if isinstance(formula, str) and formula.startswith("="):
            ws.cell(i, 4, formula)
        else:
            ws.cell(i, 4, formula)
        ws.cell(i, 5, f"=IF(C{i}=0,0,D{i}/C{i}*100)")
        ws.cell(i, 6, met)

    tr = start + 1 + len(data)
    ws.cell(tr, 2, "TOTAL")
    ws.cell(tr, 3, 100)
    ws.cell(tr, 4, f"=SUM(D{start+1}:D{tr-1})")
    ws.cell(tr, 5, f"=D{tr}")
    ws.cell(tr, 2).font = Font(bold=True)

    ws.cell(tr + 2, 1, "Skor CIS (0-100):")
    ws.cell(tr + 3, 1, "Skor Konversi (700):")
    ws.cell(tr + 4, 1, "Medallion (>700):")

    set_widths(ws, {1: 8, 2: 32, 3: 12, 4: 12, 5: 10, 6: 10})


def sheet_panduan(wb: Workbook) -> None:
    ws = wb.create_sheet("Panduan", 0)
    ws["A1"] = "PANDUAN PENGGUNAAN — LKS 2026 Penilaian Juri"
    ws["A1"].font = Font(bold=True, size=14)
    lines = [
        "",
        "1. Isi header (Tim No, Nama, Juri) di setiap sheet modul.",
        "2. Modul A: isi skor Juri 1-3 (0-3). Rata-rata & Skor Akhir otomatis.",
        "3. Modul D/E: isi Hasil dengan 1 (berhasil) atau 0 (gagal). Skor otomatis.",
        "4. Modul E: isi posisi kubus dari lembar Undian / skenario soal.",
        "5. Sheet Rekapitulasi menarik total dari Modul A, D, E (B & C isi manual).",
        "6. Skor resmi tetap diinput ke CIS sesuai Marking Scheme Chief Expert.",
        "",
        "Acuan: docs/SOP-JURI.md | Skenario latihan: docs/SKENARIO-SOAL-CONTOH.md",
    ]
    for i, line in enumerate(lines, 2):
        ws.cell(i, 1, line)
    ws.column_dimensions["A"].width = 80


def main() -> None:
    wb = Workbook()
    sheet_panduan(wb)
    sheet_modul_a(wb)
    sheet_modul_d(wb)
    sheet_modul_e(wb)
    sheet_rekap(wb)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"Created: {OUT}")


if __name__ == "__main__":
    main()
