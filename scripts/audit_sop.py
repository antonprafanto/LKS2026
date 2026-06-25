"""Audit docs/SOP-JURI.md for consistency with Excel, PPT, and official rules."""

from __future__ import annotations

import re
from pathlib import Path

SOP = Path(__file__).resolve().parent.parent / "docs" / "SOP-JURI.md"

BUGS: list[str] = []
WARNS: list[str] = []


def bug(msg: str) -> None:
    BUGS.append(msg)


def warn(msg: str) -> None:
    WARNS.append(msg)


def main() -> int:
    text = SOP.read_text(encoding="utf-8")

    if "6,67" in text or "6.67" in text:
        bug("Masih menyebut bobot 6,67 per kubus (Excel memakai 60/9)")
    if ".csv" in text.lower():
        bug("Masih ada referensi file CSV")
    if "tersedia di repo ini" in text and "Kisi-kisi" in text:
        warn("Menyebut Kisi-kisi di repo — file PNG sudah dihapus; arahkan ke Puspresnas")
    if "Briefing" not in text and "pptx" not in text.lower():
        warn("Tidak menyebut file briefing PPTX")
    if "SKENARIO" not in text:
        warn("Tidak menyebut skenario latihan juri")
    if "12,9" not in text and "12.9" not in text:
        bug("Modul D: bobot mentah 12,9 tidak dijelaskan (Excel diskalakan ke 12)")
    if "Save As" not in text and "duplikat" not in text.lower():
        warn("Tidak menjelaskan duplikat Excel per tim")
    if "Log Run" not in text:
        bug("Sheet Log Run Otonom tidak disebut di alur Modul E")
    if "Klafikasi" in text:
        warn("Typo: Klafikasi → Klarifikasi")
    if "H-30" in text and "30 hari" not in text:
        warn("H-30 tidak dijelaskan (= 30 hari sebelum lomba)")
    if "intensitas cahaya" not in text.lower() and "pencahayaan" not in text.lower():
        warn("Tidak ada catatan faktor pencahayaan venue (visi/Lidar)")
    if "agregasi" not in text.lower() and "multi-run" not in text.lower():
        warn("Tidak menjelaskan agregasi skor Modul E H2 vs H3 (trial/final)")
    if "Medallion" not in text:
        bug("Medallion of Excellence tidak disebut")

    # Modul D item sum
    weights = re.findall(r"\| 0\.\d+ \|", text)
    d_section = text.split("## 8. SOP Penilaian Modul D")[1].split("## 9.")[0]
    nums = [float(x.strip("| ")) for x in re.findall(r"\| (\d+\.\d+) \|", d_section) if "TOTAL" not in d_section]
    item_weights = [0.5, 0.5, 0.5, 0.5, 0.3, 0.6, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 2.0, 2.0]
    if abs(sum(item_weights) - 12.9) < 0.01:
        if "12.00" in d_section and "12,9" not in d_section:
            bug("Tabel Modul D TOTAL 12,00 padahal jumlah bobot item = 12,9")

    import sys

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    print(f"=== AUDIT {SOP.name} ===")
    for b in BUGS:
        print("[BUG]", b)
    for w in WARNS:
        print("[WARN]", w)
    if not BUGS and not WARNS:
        print("Tidak ada masalah terdeteksi.")
    return 1 if BUGS else 0


if __name__ == "__main__":
    raise SystemExit(main())
