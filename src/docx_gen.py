"""DOCX A4 12-label (3x4 grid) sheet generator."""

import hashlib
import math
import shutil
import tempfile
import uuid
from pathlib import Path
from typing import List, Sequence

import docx
from docx.shared import Mm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.section import WD_SECTION_START
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from PIL import Image, ImageDraw

from src.config import (
    PAGE_WIDTH_MM,
    PAGE_HEIGHT_MM,
    PAGE_MARGIN_TOP_MM,
    PAGE_MARGIN_BOTTOM_MM,
    PAGE_MARGIN_LEFT_MM,
    PAGE_MARGIN_RIGHT_MM,
    GRID_COLS,
    GRID_ROWS,
    LABELS_PER_PAGE,
    COL_WIDTHS_MM,
    ROW_HEIGHT_MM,
    ROW_HEIGHT_DXA,
    MAX_PHOTO_WIDTH_MM,
    MAX_PHOTO_HEIGHT_MM,
    MAX_BARCODE_WIDTH_MM,
    MAX_BARCODE_HEIGHT_MM,
    FONT_NAME,
    FONT_SIZE_NAME_PT,
    FONT_SIZE_CODE_PT,
    BORDER_COLOR_OUTER,
    BORDER_COLOR_INNER,
)
from src.image_ops import process_product_image, calculate_fit_dimensions
from src.barcode_gen import generate_barcode_image
from src.session import ProductItem


def _apply_section_page_setup(section) -> None:
    """Applies A4 page dimensions and margins to a Word section."""
    section.page_width = Mm(PAGE_WIDTH_MM)
    section.page_height = Mm(PAGE_HEIGHT_MM)
    section.top_margin = Mm(PAGE_MARGIN_TOP_MM)
    section.bottom_margin = Mm(PAGE_MARGIN_BOTTOM_MM)
    section.left_margin = Mm(PAGE_MARGIN_LEFT_MM)
    section.right_margin = Mm(PAGE_MARGIN_RIGHT_MM)


def _setup_table_properties(table) -> None:
    """Applies fixed layout and light gray cutting guide borders to the table."""
    tblPr = table._tbl.tblPr
    # Fixed table layout in OOXML
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

    # Light gray borders for cutting guides
    borders_xml = f"""
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="{BORDER_COLOR_OUTER}"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="{BORDER_COLOR_OUTER}"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{BORDER_COLOR_OUTER}"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="{BORDER_COLOR_OUTER}"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{BORDER_COLOR_INNER}"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="{BORDER_COLOR_INNER}"/>
        </w:tblBorders>
    """
    tblPr.append(parse_xml(borders_xml))

    # Explicit column widths in tblGrid
    tblGrid = table._tbl.find(qn("w:tblGrid"))
    if tblGrid is None:
        grid_cols_xml = "".join(
            f'<w:gridCol {nsdecls("w")} w:w="{int(round(w * 56.6929))}"/>'
            for w in COL_WIDTHS_MM
        )
        grid_element = parse_xml(f'<w:tblGrid {nsdecls("w")}>{grid_cols_xml}</w:tblGrid>')
        table._tbl.insert(1, grid_element)


def _set_cell_properties(cell, width_mm: float) -> None:
    """Configures cell width, margins, and vertical alignment with explicit OOXML."""
    cell.width = Mm(width_mm)
    tcPr = cell._tc.get_or_add_tcPr()

    # Explicit cell width in OOXML w:tcW
    width_dxa = int(round(width_mm * 56.6929))
    tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{width_dxa}" w:type="dxa"/>'))

    # Center vertical alignment
    tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))

    # Cell internal padding (dxa)
    tcPr.append(parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'  <w:top w:w="40" w:type="dxa"/>'
        f'  <w:bottom w:w="40" w:type="dxa"/>'
        f'  <w:left w:w="80" w:type="dxa"/>'
        f'  <w:right w:w="80" w:type="dxa"/>'
        f'</w:tcMar>'
    ))


def _populate_label_cell(
    cell,
    product: ProductItem,
    temp_dir: Path
) -> None:
    """
    Fills a single cell with product photo, name, barcode, and barcode number.
    Uses normal proportional line spacing (line_spacing = 1.0) so images and
    text render with full visibility and zero clipping.
    """
    # 1. Product Photo (Inline, centered, normal line spacing)
    p_img = cell.paragraphs[0]
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(1)
    p_img.paragraph_format.space_after = Pt(2)
    p_img.paragraph_format.line_spacing = 1.0

    photo_path = Path(product.image_path)
    # Safe temp filename using uuid4 to support any original filename
    processed_photo = temp_dir / f"proc_{uuid.uuid4().hex}.png"

    try:
        if photo_path.is_file():
            img_w, img_h = process_product_image(photo_path, processed_photo)
        else:
            placeholder = Image.new("RGB", (200, 200), (245, 245, 245))
            d = ImageDraw.Draw(placeholder)
            d.text((50, 90), "No Image", fill=(120, 120, 120))
            placeholder.save(processed_photo)
            img_w, img_h = 200, 200
    except Exception:
        placeholder = Image.new("RGB", (200, 200), (245, 245, 245))
        d = ImageDraw.Draw(placeholder)
        d.text((50, 90), "No Image", fill=(120, 120, 120))
        placeholder.save(processed_photo)
        img_w, img_h = 200, 200

    fit_w, fit_h = calculate_fit_dimensions(
        img_w, img_h, MAX_PHOTO_WIDTH_MM, MAX_PHOTO_HEIGHT_MM
    )
    r_img = p_img.add_run()
    r_img.add_picture(str(processed_photo), width=Mm(fit_w), height=Mm(fit_h))

    # 2. Product Name (Bold, Centered, normal line spacing)
    p_name = cell.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(1)
    p_name.paragraph_format.space_after = Pt(2)
    p_name.paragraph_format.line_spacing = 1.0

    r_name = p_name.add_run(product.name)
    r_name.font.name = FONT_NAME
    name_size = FONT_SIZE_NAME_PT if len(product.name) <= 30 else FONT_SIZE_NAME_PT - 1.0
    r_name.font.size = Pt(name_size)
    r_name.bold = True
    r_name.font.color.rgb = RGBColor(20, 20, 20)

    # 3. Barcode Image (Inline, centered, normal line spacing)
    p_bc = cell.add_paragraph()
    p_bc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_bc.paragraph_format.space_before = Pt(1)
    p_bc.paragraph_format.space_after = Pt(1)
    p_bc.paragraph_format.line_spacing = 1.0

    # Safe temp barcode filename using sha256 to support slashes, colons, spaces, etc.
    bc_hash = hashlib.sha256(product.barcode.encode("utf-8", errors="replace")).hexdigest()[:16]
    bc_output_path = temp_dir / f"bc_{bc_hash}_{uuid.uuid4().hex[:6]}.png"
    bc_path, _, bc_size = generate_barcode_image(product.barcode, bc_output_path)

    bc_w, bc_h = calculate_fit_dimensions(
        bc_size[0], bc_size[1], MAX_BARCODE_WIDTH_MM, MAX_BARCODE_HEIGHT_MM
    )
    r_bc = p_bc.add_run()
    r_bc.add_picture(str(bc_path), width=Mm(bc_w), height=Mm(bc_h))

    # 4. Barcode Code Text (Centered, Clear, normal line spacing)
    p_code = cell.add_paragraph()
    p_code.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_code.paragraph_format.space_before = Pt(0)
    p_code.paragraph_format.space_after = Pt(1)
    p_code.paragraph_format.line_spacing = 1.0

    r_code = p_code.add_run(product.barcode)
    r_code.font.name = FONT_NAME
    r_code.font.size = Pt(FONT_SIZE_CODE_PT)
    r_code.bold = True
    r_code.font.color.rgb = RGBColor(30, 30, 30)


def create_label_sheet(
    products: Sequence[ProductItem],
    output_path: str | Path
) -> Path:
    """
    Creates an A4 document containing 12 labels per page (3 cols x 4 rows).
    Preserves the exact order of products.
    Returns the resolved output Path.
    """
    if not products:
        raise ValueError("Cannot create label sheet with empty product list.")

    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    temp_render_dir = Path(tempfile.mkdtemp(prefix="label_render_"))

    try:
        doc = docx.Document()
        total_products = len(products)
        num_pages = math.ceil(total_products / LABELS_PER_PAGE)

        for page_idx in range(num_pages):
            if page_idx == 0:
                sec = doc.sections[0]
            else:
                sec = doc.add_section(WD_SECTION_START.NEW_PAGE)

            _apply_section_page_setup(sec)

            table = doc.add_table(rows=GRID_ROWS, cols=GRID_COLS)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            _setup_table_properties(table)

            # Configure rows with exact OOXML trHeight and cantSplit
            for row in table.rows:
                trPr = row._tr.get_or_add_trPr()
                trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
                trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{ROW_HEIGHT_DXA}" w:hRule="exact"/>'))
                row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
                row.height = Mm(ROW_HEIGHT_MM)

            # Fill 12 cells for current page
            page_start = page_idx * LABELS_PER_PAGE
            for i in range(LABELS_PER_PAGE):
                row_idx = i // GRID_COLS
                col_idx = i % GRID_COLS
                cell = table.cell(row_idx, col_idx)
                _set_cell_properties(cell, COL_WIDTHS_MM[col_idx])

                prod_idx = page_start + i
                if prod_idx < total_products:
                    _populate_label_cell(cell, products[prod_idx], temp_render_dir)
                else:
                    # Empty cell retains cutting borders
                    cell.paragraphs[0].text = ""

        doc.save(str(output_path))
        return output_path
    finally:
        shutil.rmtree(temp_render_dir, ignore_errors=True)
