"""Sample generation script creating a printable label DOCX with a sample product."""

import sys
from pathlib import Path
from PIL import Image, ImageDraw

from src.session import ProductItem
from src.docx_gen import create_label_sheet
from src.pdf_converter import convert_docx_to_pdf


def generate_sample_photo(target_path: Path) -> Path:
    """Generates a clean sample product image with a white background."""
    target_path.parent.mkdir(parents=True, exist_ok=True)
    # 400x400 white canvas
    im = Image.new("RGB", (400, 400), (255, 255, 255))
    draw = ImageDraw.Draw(im)

    # Draw transition pipe reducer (Perexod 76-50mm) silhouette / illustration
    # Outer metal body in industrial blue-gray
    body_color = (60, 90, 120)
    highlight_color = (100, 140, 180)
    rim_color = (40, 60, 80)

    # Upper wide flange (76mm)
    draw.ellipse([100, 60, 300, 100], fill=body_color, outline=rim_color, width=3)
    draw.ellipse([120, 70, 280, 90], fill=(220, 230, 240), outline=rim_color, width=2)

    # Conical body
    draw.polygon(
        [(100, 80), (300, 80), (250, 280), (150, 280)],
        fill=body_color,
        outline=rim_color,
    )
    # Shading / highlight stripe
    draw.polygon([(170, 80), (230, 80), (200, 280), (175, 280)], fill=highlight_color)

    # Lower narrow flange (50mm)
    draw.ellipse([150, 270, 250, 300], fill=body_color, outline=rim_color, width=3)
    draw.ellipse([165, 275, 235, 295], fill=(200, 210, 220), outline=rim_color, width=2)

    # Text label on body
    draw.text((155, 175), "76 -> 50", fill=(255, 255, 255))

    im.save(target_path, format="PNG")
    return target_path


def generate_sample_docx(
    output_docx: str | Path = "sample_perexod_labels.docx",
    convert_pdf: bool = True
) -> Path:
    """Generates a sample DOCX with the required sample product."""
    output_docx = Path(output_docx).resolve()

    # Create sample image
    sample_img_path = output_docx.parent / "sample_perexod.png"
    generate_sample_photo(sample_img_path)

    # Sample product details as specified in requirements
    sample_product = ProductItem(
        name="Perexod 76-50mm",
        barcode="1000049",
        image_path=str(sample_img_path),
    )

    print(f"Generating sample label sheet with product: '{sample_product.name}' (Barcode: {sample_product.barcode})...")
    created_docx = create_label_sheet([sample_product], output_docx)
    print(f"✅ Sample DOCX created successfully: {created_docx} ({created_docx.stat().st_size} bytes)")

    if convert_pdf:
        pdf_path, err = convert_docx_to_pdf(created_docx)
        if pdf_path and pdf_path.is_file():
            print(f"✅ Sample PDF created successfully: {pdf_path} ({pdf_path.stat().st_size} bytes)")
        else:
            print(f"ℹ️ PDF conversion skipped: {err}")

    return created_docx


if __name__ == "__main__":
    out_file = sys.argv[1] if len(sys.argv) > 1 else "sample_perexod_labels.docx"
    generate_sample_docx(out_file)
