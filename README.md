# Printable Product Label Sheet Telegram Bot

A Python Telegram bot that generates professional, cashier-ready printable A4 product label sheets without using any AI APIs. Users submit products one by one (Photo ➔ Name ➔ Barcode), and the bot compiles them into a printable document where each page contains **12 labels arranged in a 3 columns × 4 rows grid**.

---

## Features

- **12 Labels per A4 Page (3 × 4 Grid)**: Formatted to fit standard A4 paper (`210mm × 297mm`) with margins, exact row heights, and `cantSplit` protection so rows never break across pages.
- **Order Preservation**: Products appear on the sheet in the exact order submitted by the user.
- **Multi-Page Pagination**: If more than 12 products are added, subsequent pages are automatically created with cutting grid guides.
- **Pure Programmatic Barcodes (Zero AI)**:
  - Generates scannable, high-resolution **Code128** barcodes for short/internal codes like `1000049`.
  - Supports **EAN-13** for 13-digit numeric codes with valid checksums (with automatic fallback to Code128 if checksum is non-standard, ensuring no user digits are altered).
  - Includes certified quiet-zone margins and 300 DPI rendering for laser/inkjet cashier scanners.
- **Automatic White Background Trimming**:
  - Automatically detects and crops white or near-white borders around products using pure Pillow image operations.
  - Transparent PNGs are cleanly flattened and trimmed.
  - Proportional scaling ensures product images are never stretched or distorted.
- **Dual Export (DOCX & PDF)**:
  - Generates DOCX files with light gray cutting guide borders (`#C8C8C8` outer, `#E0E0E0` inner).
  - Converts to PDF automatically if LibreOffice is installed; if not, sends the DOCX and provides a clear installation tip.
- **Per-User Session Management**:
  - Each Telegram user has an isolated session and directory.
  - Temporary files and drafts are safely cleaned up upon `/new` or `/cancel`.

---

## Label Design Specifications

Each label cell contains:
1. **Product Photo**: Centered horizontally, white background trimmed, aspect ratio preserved (max 48mm × 25mm).
2. **Product Name**: Centered, bold Arial (8.5pt).
3. **Barcode Image**: Centered, high-contrast black-on-white with scannable quiet zones (max 46mm × 13mm).
4. **Barcode Number**: Centered, bold Arial (8.0pt).
5. **Cutting Guides**: Light gray borders surrounding each label for easy trimming with scissors or a paper cutter.

---

## Project Structure

```
├── .env.example          # Sample environment variables
├── .env                  # Environment configuration (TELEGRAM_BOT_TOKEN)
├── requirements.txt      # Python dependencies
├── README.md             # Documentation and instructions
├── test_sample.py        # Standalone sample generator (creates sample DOCX)
├── tests/
│   └── test_pipeline.py  # Unit tests for image trimming, barcodes, sessions, docx
└── src/
    ├── __init__.py
    ├── config.py         # Dimensions, typography, and path constants
    ├── image_ops.py      # White-background trimming and aspect scaling
    ├── barcode_gen.py    # Programmatic Code128 / EAN-13 generator (300 DPI)
    ├── docx_gen.py       # A4 12-label (3x4) DOCX sheet generator
    ├── pdf_converter.py  # Headless LibreOffice PDF conversion with fallback
    ├── session.py        # Per-user session tracking & temporary storage
    └── bot.py            # Telegram bot application & conversation handlers
```

---

## Requirements & Prerequisites

- **Python**: 3.10 or higher (tested on Python 3.13)
- **Telegram Bot Token**: Obtain a bot token from [@BotFather](https://t.me/botfather) on Telegram.
- **Optional for PDF conversion**: LibreOffice
  - **macOS**: `brew install --cask libreoffice`
  - **Ubuntu/Debian**: `sudo apt-get update && sudo apt-get install -y libreoffice`
  - **Windows**: Download installer from [libreoffice.org](https://www.libreoffice.org)

*(Note: If LibreOffice is not installed, the bot still works completely and sends the formatted DOCX file).*

---

## Installation & Setup

1. **Clone or navigate to the project directory**:
   ```bash
   cd "/Users/user/Desktop/doc bot"
   ```

2. **Create and activate a virtual environment**:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   # On Windows: .venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure environment variables**:
   Copy `.env.example` to `.env` and enter your Telegram Bot Token:
   ```bash
   cp .env.example .env
   ```
   Open `.env` in your editor and set:
   ```env
   TELEGRAM_BOT_TOKEN=your_bot_token_here
   ```

---

## Running the Bot

Run the bot with the standard module execution command:
```bash
source .venv/bin/activate
python -m src.bot
```
*(Direct script execution `python src/bot.py` is also supported via a package-path bootstrap).*

The bot will initialize, register UI commands with Telegram, and begin polling for updates.

---

## Bot Commands & Usage Guide

| Command | Description |
| :--- | :--- |
| `/start` | Displays welcome message, features, and command guide. |
| `/new` | Clears all products and starts a fresh label sheet. |
| `/add` | Initiates the 3-step prompt flow to add a new product. |
| `/list` | Displays the current ordered list of products and page count. |
| `/remove <N>` | Removes product number `N` from the sheet (e.g. `/remove 2`). |
| `/preview` | Generates and sends a preview DOCX/PDF without clearing the list. |
| `/finish` | Compiles final printable DOCX & PDF files and sends them to the user. |
| `/cancel` | Cancels the current `/add` product input flow. |

> **Note on `/done`**: `/done` is not needed in the active workflow. Products are automatically committed and confirmed immediately after you submit the barcode value.

### Step-by-Step Flow:
1. Send `/add`.
2. **Step 1/3 (Photo)**: Upload a photo of the product (as a photo or uncompressed image document).
3. **Step 2/3 (Name)**: Type the product name (e.g., `Perexod 76-50mm`).
4. **Step 3/3 (Barcode)**: Type the barcode or code (e.g., `1000049`, `5901234123457`, or internal code).
5. The product is **automatically saved and confirmed** with its position on the sheet.
6. Repeat `/add` for more products, or send `/finish` to generate your printable label sheet.

---

## Standalone Test & Verification Scripts

### 1. Sample Product Generator (`test_sample.py`)
Generates a test DOCX document with the required sample product (`Perexod 76-50mm`, Barcode: `1000049`):

```bash
source .venv/bin/activate
python test_sample.py
```

### 2. Render Verification Script (`verify_render.py`)
Generates and validates 1-product, 12-product, and 13-product documents:
- **Verifies Real PDF Page Count using `pypdf`**:
  - `render_1_product.pdf`: exactly 1 page
  - `render_12_products.pdf`: exactly 1 page
  - `render_13_products.pdf`: exactly 2 pages
- **Confirms OOXML Layout**:
  - `<w:tblLayout w:type="fixed"/>` fixed table layout
  - Exact row heights (`w:trHeight`) and `w:cantSplit`
  - Explicit cell widths (`w:tcW`)
  - Absence of fixed 1pt line height
- **Validates Safe Barcode Filenames**: tests barcodes with special Code128 characters (`/`, `:`, spaces, dashes).
- **Visual Thumbnail Verification**: confirms non-clipped, high-contrast images and scannable barcode bars.

```bash
source .venv/bin/activate
python verify_render.py
```

---

## Running Automated Tests

Run the comprehensive unit test suite:
```bash
source .venv/bin/activate
python -m unittest tests/test_pipeline.py
```

Tests cover:
- White-background trimming and transparency handling.
- Fit dimension scaling without distortion.
- Code128 generation for short codes.
- EAN-13 checksum validation and fallback to Code128.
- Ordered session management and product removal.
- Single-page and multi-page DOCX pagination (12 labels/page).
- Fixed table layout and absence of 1pt line clipping.
- Markdown special character escaping.

---

## Printing Tips

- **Paper Size**: A4 (`210mm × 297mm`).
- **Scale**: Print at **100% (Actual size)**; do **not** select "Fit to page" or "Shrink oversized pages" to preserve exact label dimensions.
- **Cutting**: Use the light gray grid lines as alignment guides for scissors or a guillotine cutter.
