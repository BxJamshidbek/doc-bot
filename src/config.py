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

# Session storage directory
SESSION_BASE_DIR = Path(os.getenv("SESSION_DIR", str(BASE_DIR / "data" / "sessions"))).resolve()

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
