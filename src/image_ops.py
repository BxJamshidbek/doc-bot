"""Image processing operations for product photos."""

from pathlib import Path
from typing import Tuple, Optional
from PIL import Image, ImageChops, ImageOps

from src.config import MAX_UPLOAD_IMAGE_SIDE_PX


def trim_white_background(
    im: Image.Image,
    threshold: int = 240,
    padding: int = 8
) -> Image.Image:
    """
    Crops white or near-white borders around the product object.
    Supports RGB, RGBA, and transparent backgrounds.
    Preserves original image if no white background is detected.
    """
    # Normalize EXIF orientation (common in smartphone photos)
    im = ImageOps.exif_transpose(im)

    if im.mode not in ("RGB", "RGBA"):
        im = im.convert("RGB")

    w, h = im.size
    if w <= 10 or h <= 10:
        return im

    # Check corners to see if image has a light/transparent background
    corners = [
        im.getpixel((0, 0)),
        im.getpixel((w - 1, 0)),
        im.getpixel((0, h - 1)),
        im.getpixel((w - 1, h - 1)),
    ]

    def is_white_or_transparent(px) -> bool:
        if im.mode == "RGBA" and len(px) >= 4 and px[3] < 30:
            return True
        return px[0] >= threshold and px[1] >= threshold and px[2] >= threshold

    white_corner_count = sum(1 for c in corners if is_white_or_transparent(c))

    # If at least 2 corners are light or transparent, treat as object on background
    if white_corner_count >= 2:
        tolerance = 255 - threshold
        bbox = None

        if im.mode == "RGBA":
            r, g, b, a = im.split()
            corners_a = [
                a.getpixel((0, 0)),
                a.getpixel((w - 1, 0)),
                a.getpixel((0, h - 1)),
                a.getpixel((w - 1, h - 1)),
            ]
            if sum(1 for ca in corners_a if ca < 30) >= 2:
                # Transparent background
                bbox = a.point(lambda p: 255 if p > 30 else 0).getbbox()
            else:
                # Opaque white background
                bg_white = Image.new("RGB", im.size, (255, 255, 255))
                diff = ImageChops.difference(im.convert("RGB"), bg_white)
                r_diff, g_diff, b_diff = diff.split()
                color_diff = ImageChops.lighter(ImageChops.lighter(r_diff, g_diff), b_diff)
                color_mask = color_diff.point(lambda p: 255 if p > tolerance else 0)
                bbox = color_mask.getbbox()
        else:
            bg_white = Image.new("RGB", im.size, (255, 255, 255))
            diff = ImageChops.difference(im, bg_white)
            r_diff, g_diff, b_diff = diff.split()
            max_diff = ImageChops.lighter(ImageChops.lighter(r_diff, g_diff), b_diff)
            mask = max_diff.point(lambda p: 255 if p > tolerance else 0)
            bbox = mask.getbbox()

        if bbox:
            min_x, min_y, max_x, max_y = bbox
            min_x = max(0, min_x - padding)
            min_y = max(0, min_y - padding)
            max_x = min(w, max_x + padding)
            max_y = min(h, max_y + padding)
            if max_x > min_x and max_y > min_y:
                return im.crop((min_x, min_y, max_x, max_y))

    return im


def process_product_image(
    input_path: str | Path,
    output_path: str | Path,
    *,
    max_side_px: Optional[int] = None,
    target_max_size_px: Optional[Tuple[int, int]] = None,
    output_format: str = "PNG",
    jpeg_quality: int = 82,
) -> Tuple[int, int]:
    """
    Loads, processes (EXIF transpose + white background trim),
    and saves normalized image. Returns (width, height) of result.
    """
    input_path = Path(input_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with Image.open(input_path) as im:
        im = ImageOps.exif_transpose(im)
        initial_max_side_px = max_side_px or MAX_UPLOAD_IMAGE_SIDE_PX
        if max(im.size) > initial_max_side_px:
            im.thumbnail(
                (initial_max_side_px, initial_max_side_px),
                Image.Resampling.LANCZOS,
            )
        processed = trim_white_background(im)
        if processed.mode == "RGBA":
            bg = Image.new("RGB", processed.size, (255, 255, 255))
            bg.paste(processed, mask=processed.split()[-1])
            processed = bg
        elif processed.mode != "RGB":
            processed = processed.convert("RGB")

        if target_max_size_px:
            target_w, target_h = target_max_size_px
            target_w = max(1, int(target_w))
            target_h = max(1, int(target_h))
            if processed.width > target_w or processed.height > target_h:
                processed.thumbnail((target_w, target_h), Image.Resampling.LANCZOS)

        normalized_format = (output_format or "PNG").strip().upper()
        if normalized_format in {"JPG", "JPEG"}:
            if processed.mode != "RGB":
                processed = processed.convert("RGB")
            processed.save(
                output_path,
                format="JPEG",
                quality=max(60, min(95, int(jpeg_quality))),
                optimize=True,
                progressive=True,
            )
        else:
            processed.save(output_path, format="PNG", optimize=True)
        return processed.size


def calculate_fit_dimensions(
    img_width_px: int,
    img_height_px: int,
    max_width_mm: float,
    max_height_mm: float
) -> Tuple[float, float]:
    """
    Calculates display dimensions in millimeters that scale the image
    proportionally without distortion, fitting strictly inside the bounds.
    """
    if img_width_px <= 0 or img_height_px <= 0:
        return max_width_mm, max_height_mm

    scale = min(max_width_mm / img_width_px, max_height_mm / img_height_px)
    fit_w = round(img_width_px * scale, 2)
    fit_h = round(img_height_px * scale, 2)
    return fit_w, fit_h
