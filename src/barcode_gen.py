"""Programmatic barcode generator supporting Code128 and EAN-13."""

from pathlib import Path
from typing import Tuple
import barcode
from barcode.writer import ImageWriter
from PIL import Image


def is_valid_ean13(code: str) -> bool:
    """Checks if the code is a valid 13-digit EAN-13 barcode with correct checksum."""
    if len(code) != 13 or not code.isdigit():
        return False
    digits = [int(d) for d in code]
    # Sum odd positions (0-indexed: 0, 2, 4...) * 1, even positions (1, 3, 5...) * 3
    weighted_sum = sum(digits[i] for i in range(0, 12, 2)) + sum(digits[i] * 3 for i in range(1, 12, 2))
    expected_check = (10 - (weighted_sum % 10)) % 10
    return digits[12] == expected_check


def _prepare_barcode(code_clean: str, writer: ImageWriter):
    """Creates a barcode object and verifies that the encoded value is unchanged."""
    barcode_type = "Code128"
    if is_valid_ean13(code_clean):
        try:
            bc = barcode.get_barcode_class("ean13")(code_clean, writer=writer)
            barcode_type = "EAN-13"
        except Exception:
            bc = barcode.get_barcode_class("code128")(code_clean, writer=writer)
            barcode_type = "Code128"
    else:
        bc = barcode.get_barcode_class("code128")(code_clean, writer=writer)

    full_code = bc.get_fullcode()
    if full_code != code_clean:
        raise ValueError(
            f"Barcode encoder changed the value: expected {code_clean!r}, got {full_code!r}."
        )
    return bc, barcode_type, full_code


def get_barcode_encoding(code: str) -> Tuple[str, str]:
    """Returns the barcode type and exact value that will be encoded."""
    code_clean = str(code).strip()
    if not code_clean:
        raise ValueError("Barcode code cannot be empty.")
    _bc, barcode_type, full_code = _prepare_barcode(code_clean, ImageWriter())
    return barcode_type, full_code


def generate_barcode_image(
    code: str,
    output_path: str | Path,
    dpi: int = 300
) -> Tuple[Path, str, Tuple[int, int]]:
    """
    Generates a scannable barcode image (PNG).
    - Uses EAN-13 if code has 13 numeric digits with a valid checksum.
    - Otherwise uses Code128 (ideal for short/internal codes like '1000049').
    Returns (saved_path, barcode_type, (width_px, height_px)).
    """
    code_clean = str(code).strip()
    if not code_clean:
        raise ValueError("Barcode code cannot be empty.")

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    writer = ImageWriter()
    writer_options = {
        "write_text": False,  # Text is rendered cleanly in Word table
        "quiet_zone": 4.0,    # Scannable quiet zones on left and right
        "module_height": 14.0,
        "module_width": 0.28,
        "dpi": dpi,
    }

    bc, barcode_type, _full_code = _prepare_barcode(code_clean, writer)

    # Render image with PIL
    img: Image.Image = bc.render(writer_options=writer_options)

    # Ensure RGB black-and-white
    if img.mode != "RGB":
        img = img.convert("RGB")

    # Save to file
    final_path = output_path.with_suffix(".png")
    img.save(final_path, format="PNG", dpi=(dpi, dpi))

    return final_path, barcode_type, img.size
