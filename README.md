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
  - Normalizes smartphone EXIF orientation and downscales very large phone photos before processing.
  - Transparent PNGs are cleanly flattened and trimmed.
  - Proportional scaling ensures product images are never stretched or distorted.
- **Dual Export (DOCX & PDF)**:
  - Generates DOCX files with light gray cutting guide borders (`#C8C8C8` outer, `#E0E0E0` inner).
  - Embeds product photos as print-quality compact JPEGs so 100+ product sheets stay small enough for Telegram.
  - If an export is still too large, automatically splits the DOCX into ordered parts instead of failing with `Request Entity Too Large`.
  - Converts to PDF automatically if LibreOffice is installed; if not, sends the DOCX and provides a clear installation tip.
  - Preview and final exports use unique filenames so files are not overwritten.
- **Named Documents & Previous Files**:
  - `/new` asks for a document name before the first product is collected.
  - Generated filenames include the document name when available.
  - `/files` shows recent preview/final documents and lets users download DOCX/PDF again.
- **Product Correction Flow**:
  - `/list` shows each added product with inline edit buttons.
  - `/edit <N>` lets users correct product photo, name, or barcode without rebuilding the whole document.
  - Corrected products keep their original order on the sheet.
- **Optional Bottom Menu**:
  - The large Telegram reply keyboard is not persistent by default.
  - `/menu` opens the bottom menu manually when needed; normal bot replies keep it hidden.
- **Per-User Session Management**:
  - Each Telegram user has an isolated session and directory.
  - Generated DOCX/PDF files are stored separately under `data/generated/<telegram_user_id>/`.
  - `/new` clears the active product list but keeps generated DOCX/PDF files until the retention cleanup removes them.
  - Generated files and stale drafts older than 30 days are cleaned automatically.

---

## Label Design Specifications

Each label cell contains:
1. **Product Photo**: Centered horizontally, white background trimmed, aspect ratio preserved (max 56mm × 36mm).
2. **Product Name**: Centered, bold Arial (up to 10.2pt, reduced automatically for long names).
3. **Barcode Image**: Centered, high-contrast black-on-white with scannable quiet zones (max 56mm × 20mm).
4. **Barcode Number**: Centered, bold Arial (9.4pt).
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
   SESSION_DIR=data/sessions
   GENERATED_DIR=data/generated
   GENERATED_RETENTION_DAYS=30
   CLEANUP_INTERVAL_HOURS=24
   MAX_UPLOAD_IMAGE_SIDE_PX=2400
   MAX_EMBEDDED_PHOTO_DPI=300
   PHOTO_JPEG_QUALITY=82
   TELEGRAM_SAFE_DOCUMENT_SIZE_MB=45
   MAX_PRODUCTS_PER_DOCX_PART=72
   ```

   On a server, keep `SESSION_DIR` and `GENERATED_DIR` on persistent disk, not in a temporary directory.

---

## Running the Bot

Run the bot with the standard module execution command:
```bash
source .venv/bin/activate
python -m src.bot
```
*(Direct script execution `python src/bot.py` is also supported via a package-path bootstrap).*

By default, the bot starts in polling mode. For webhook deployment, set `BOT_RUN_MODE=webhook` and configure the webhook variables shown below.

---

## Bot Commands & Usage Guide

| Command | Description |
| :--- | :--- |
| `/start` | Displays welcome message, features, and command guide. |
| `/new` | Clears active products, starts a fresh label sheet, and asks for a document name. |
| `/add` | Initiates the 3-step prompt flow to add a new product. |
| `/list` | Displays the current ordered list of products and page count. |
| `/remove <N>` | Removes product number `N` from the sheet (e.g. `/remove 2`). |
| `/preview` | Generates and sends a preview DOCX/PDF without clearing the list. |
| `/finish` | Compiles final printable DOCX & PDF files and sends them to the user without clearing the list. |
| `/files` | Shows previous preview/final documents and lets the user download them again. |
| `/edit <N>` | Opens correction controls for product number `N` (photo, name, or barcode). |
| `/cancel` | Cancels the current `/add` product input flow. |
| `/menu` | Opens the optional bottom Telegram keyboard manually. |

> **Note on `/done`**: `/done` is not needed in the active workflow. Products are automatically committed and confirmed immediately after you submit the barcode value.

### Step-by-Step Flow:
1. Send `/new` and type the document name.
2. Upload a photo of the first product (as a photo or uncompressed image document).
3. Type the product name (e.g., `Perexod 76-50mm`).
4. Type the barcode or code (e.g., `1000049`, `5901234123457`, or internal code).
5. The product is **automatically saved and confirmed** with its position on the sheet.
6. Use `/add` to add the next product. The current product list is shown before the next photo prompt.
7. Use `/preview` at any time to download a test document without stopping the current list.
8. Use `/finish` at any time to download the final document. The list is kept; use `/new` only when you want to start a fresh sheet.
9. Use `/files` to download previous preview/final documents again while they are still within the retention period.

### Correcting a Product

If a product photo, name, or barcode is wrong:
1. Send `/list`.
2. Press the product's `✏️` button, or send `/edit <N>` (for example, `/edit 2`).
3. Choose `Rasmni tuzatish`, `Nomni tuzatish`, or `Shtrixni tuzatish`.
4. Send the replacement value. The bot updates the same product position and shows the refreshed list.

### Photo Quality Guidance

The bot accepts normal phone photos. For best results:
- Put the product on a white or light background.
- Keep the product centered and well lit.
- Avoid strong shadows and busy backgrounds.
- Send the image as a file/document when you want Telegram to avoid compressing it.

The bot preserves the image aspect ratio, corrects phone rotation metadata, trims white borders when safe, and fits the product into the label without stretching.

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
- Smartphone orientation normalization and large-photo downscaling.
- Fit dimension scaling without distortion.
- Code128 generation for short codes.
- Exact barcode encoded-value preservation for codes like `1000049`.
- EAN-13 checksum validation and fallback to Code128.
- Ordered session management and product removal.
- Document naming and previous export history.
- Unique preview/final export filenames under `data/generated/<telegram_user_id>/`.
- 30-day cleanup for old generated files and stale drafts.
- Single-page and multi-page DOCX pagination (12 labels/page).
- Fixed table layout and absence of 1pt line clipping.
- Markdown special character escaping.

---

## Server Deployment Notes

Recommended Ubuntu package setup:

```bash
sudo apt-get update
sudo apt-get install -y python3 python3-venv python3-pip libreoffice poppler-utils git nginx certbot python3-certbot-nginx
cd /root
git clone https://github.com/BxJamshidbek/doc-bot.git "doc bot"
cd "doc bot"
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Edit `.env` and set the real Telegram token. Do not commit `.env`.

### Polling mode

Use this when you do not have a domain and HTTPS yet:

```env
BOT_RUN_MODE=polling
```

### Webhook mode

Webhook requires a domain or subdomain with DNS `A` record pointing to the server IP.

Example DNS:

```text
bot.example.com  A  87.192.253.111
```

Example `.env` values:

```env
BOT_RUN_MODE=webhook
WEBHOOK_BASE_URL=https://bot.example.com
WEBHOOK_PATH=/tg/replace_with_random_path
WEBHOOK_LISTEN_HOST=127.0.0.1
WEBHOOK_LISTEN_PORT=8091
WEBHOOK_SECRET_TOKEN=replace_with_random_secret
SESSION_DIR=/root/doc bot/data/sessions
GENERATED_DIR=/root/doc bot/data/generated
GENERATED_RETENTION_DAYS=30
CLEANUP_INTERVAL_HOURS=24
```

Example Nginx reverse proxy:

```nginx
server {
    server_name bot.example.com;

    client_max_body_size 50M;

    location /tg/replace_with_random_path {
        proxy_pass http://127.0.0.1:8091/tg/replace_with_random_path;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

After DNS points to the server, issue HTTPS with:

```bash
sudo certbot --nginx -d bot.example.com
```

Example systemd service:

```ini
[Unit]
Description=Product Label Telegram Bot
After=network.target

[Service]
WorkingDirectory=/root/doc bot
EnvironmentFile=/root/doc bot/.env
ExecStart=/root/doc bot/.venv/bin/python -m src.bot
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

After saving the service file:

```bash
sudo systemctl daemon-reload
sudo systemctl enable doc-bot
sudo systemctl start doc-bot
sudo systemctl status doc-bot
```

---

## Printing Tips

- **Paper Size**: A4 (`210mm × 297mm`).
- **Scale**: Print at **100% (Actual size)**; do **not** select "Fit to page" or "Shrink oversized pages" to preserve exact label dimensions.
- **Cutting**: Use the light gray grid lines as alignment guides for scissors or a guillotine cutter.
