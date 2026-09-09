"""Configuration settings for the Label Bot."""

import os
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Telegram Bot Token
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()

# Base project path
BASE_DIR = Path(__file__).resolve().parent.parent


def _env_int(name: str, default: int) -> int:
    """Reads an integer environment variable with a safe fallback."""
    try:
        return int(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        return default


# Session storage directory
SESSION_BASE_DIR = Path(os.getenv("SESSION_DIR", str(BASE_DIR / "data" / "sessions"))).resolve()

# Persistent generated DOCX/PDF storage. Files are cleaned by retention policy.
GENERATED_BASE_DIR = Path(os.getenv("GENERATED_DIR", str(BASE_DIR / "data" / "generated"))).resolve()
GENERATED_RETENTION_DAYS = _env_int("GENERATED_RETENTION_DAYS", 30)
CLEANUP_INTERVAL_HOURS = _env_int("CLEANUP_INTERVAL_HOURS", 24)

# Incoming phone photos can be very large; downscale before processing to avoid memory spikes.
MAX_UPLOAD_IMAGE_SIDE_PX = _env_int("MAX_UPLOAD_IMAGE_SIDE_PX", 2400)

# Runtime mode. Use "polling" for simple deployment or "webhook" behind HTTPS/Nginx.
BOT_RUN_MODE = os.getenv("BOT_RUN_MODE", "polling").strip().lower()
WEBHOOK_BASE_URL = os.getenv("WEBHOOK_BASE_URL", "").strip().rstrip("/")
WEBHOOK_PATH = os.getenv("WEBHOOK_PATH", "/telegram-webhook").strip() or "/telegram-webhook"
if not WEBHOOK_PATH.startswith("/"):
    WEBHOOK_PATH = f"/{WEBHOOK_PATH}"
WEBHOOK_LISTEN_HOST = os.getenv("WEBHOOK_LISTEN_HOST", "127.0.0.1").strip() or "127.0.0.1"
WEBHOOK_LISTEN_PORT = _env_int("WEBHOOK_LISTEN_PORT", 8091)
WEBHOOK_SECRET_TOKEN = os.getenv("WEBHOOK_SECRET_TOKEN", "").strip()

# Page and Label layout dimensions (A4 standard)
PAGE_WIDTH_MM = 210.0
PAGE_HEIGHT_MM = 297.0

# Margins
PAGE_MARGIN_TOP_MM = 6.5
PAGE_MARGIN_BOTTOM_MM = 6.5
PAGE_MARGIN_LEFT_MM = 10.0
PAGE_MARGIN_RIGHT_MM = 10.0

# Label grid structure
GRID_COLS = 3
GRID_ROWS = 4
LABELS_PER_PAGE = GRID_COLS * GRID_ROWS  # 12 labels per page

# Exact column widths and row height to fill page cleanly without overflow
COL_WIDTHS_MM = [63.3, 63.4, 63.3]
ROW_HEIGHT_MM = 67.5

# OOXML dxa units (1 mm = 56.6929 dxa / twips)
COL_WIDTHS_DXA = [int(round(w * 56.6929)) for w in COL_WIDTHS_MM]
ROW_HEIGHT_DXA = int(round(ROW_HEIGHT_MM * 56.6929))

# Image & Barcode dimension constraints inside label
MAX_PHOTO_WIDTH_MM = 48.0
MAX_PHOTO_HEIGHT_MM = 25.0

MAX_BARCODE_WIDTH_MM = 46.0
MAX_BARCODE_HEIGHT_MM = 15.0

# Typography
FONT_NAME = "Arial"
FONT_SIZE_NAME_PT = 8.5
FONT_SIZE_CODE_PT = 8.0

# Border styling for cutting guides
BORDER_COLOR_OUTER = "C8C8C8"  # Light gray cutting border
BORDER_COLOR_INNER = "E0E0E0"  # Subtle inner cell borders
