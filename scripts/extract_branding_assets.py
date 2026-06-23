"""Extract branding images from official LKS 2026 cover pages."""

from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"
ASSETS.mkdir(exist_ok=True)


def crop_save(src: Path, box: tuple[int, int, int, int], out: Path) -> None:
    img = Image.open(src).convert("RGBA")
    cropped = img.crop(box)
    cropped.save(out, "PNG")
    print(f"Saved {out.name} {cropped.size}")


def main() -> None:
    kisi1 = ROOT / "Kisi-kisi_page_1.png"
    kisi2 = ROOT / "Kisi-kisi_page_2.png"

    if not kisi1.exists() or not kisi2.exists():
        print(
            "Lewati ekstraksi: letakkan Kisi-kisi_page_1.png dan _2.png di root repo "
            "(unduh dari Puspresnas) lalu jalankan ulang."
        )
        return

    # Kemendikdasmen / Pusat Prestasi Nasional header (top-left)
    crop_save(kisi1, (60, 55, 720, 260), ASSETS / "logo-kemendikdasmen.png")

    # LKS Dikmen circular logo (left on cover)
    crop_save(kisi1, (55, 1180, 430, 1555), ASSETS / "logo-lks-dikmen.png")

    # LKS geometric logo from kisi-kisi page 2
    crop_save(kisi2, (620, 520, 1165, 1065), ASSETS / "logo-lks-geometric.png")

    # Orange accent bar sample for footer stripe
    crop_save(kisi1, (0, 2460, 1785, 2526), ASSETS / "footer-stripe.png")


if __name__ == "__main__":
    main()
