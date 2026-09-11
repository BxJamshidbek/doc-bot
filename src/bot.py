"""Telegram Bot for generating printable A4 product label sheets."""

import asyncio
import logging
import sys
from pathlib import Path
from typing import Optional

# Enable direct script execution (e.g. `python src/bot.py`)
if __package__ is None or __package__ == "":
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from telegram import (
    Update,
    BotCommand,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
)
from telegram.constants import ParseMode
from telegram.helpers import escape_markdown
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)
from telegram.warnings import PTBUserWarning

import warnings
warnings.filterwarnings("ignore", category=PTBUserWarning)

from src.config import (
    TELEGRAM_BOT_TOKEN,
    SESSION_BASE_DIR,
    GENERATED_BASE_DIR,
    GENERATED_RETENTION_DAYS,
    CLEANUP_INTERVAL_HOURS,
    BOT_RUN_MODE,
    WEBHOOK_BASE_URL,
    WEBHOOK_PATH,
    WEBHOOK_LISTEN_HOST,
    WEBHOOK_LISTEN_PORT,
    WEBHOOK_SECRET_TOKEN,
    LABELS_PER_PAGE,
    TELEGRAM_SAFE_DOCUMENT_SIZE_BYTES,
    TELEGRAM_SAFE_DOCUMENT_SIZE_MB,
    MAX_PRODUCTS_PER_DOCX_PART,
)
from src.session import ExportRecord, SessionManager, UserSession, normalize_document_name
from src.docx_gen import create_label_sheet
from src.pdf_converter import convert_docx_to_pdf
from src.barcode_gen import generate_barcode_image

# Configure logging
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)


class TelegramDocumentTooLargeError(Exception):
    """Raised before/when Telegram rejects a document upload because it is too large."""

    def __init__(self, path: Path, size_bytes: int, original_error: Exception | None = None):
        self.path = Path(path)
        self.size_bytes = int(size_bytes)
        self.original_error = original_error
        super().__init__(f"{self.path.name} is too large for Telegram ({format_file_size(self.size_bytes)})")

# Conversation states
(
    STATE_WAITING_DOCUMENT_NAME,
    STATE_WAITING_PHOTO,
    STATE_WAITING_NAME,
    STATE_WAITING_BARCODE,
    STATE_WAITING_EDIT_PHOTO,
    STATE_WAITING_EDIT_NAME,
    STATE_WAITING_EDIT_BARCODE,
    STATE_WAITING_SWAP_TARGET,
    STATE_WAITING_MOVE_TARGET,
) = range(9)

# Global session manager
session_mgr = SessionManager(SESSION_BASE_DIR, GENERATED_BASE_DIR)

SUPPORTED_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp"}
IMAGE_MIME_EXTENSIONS = {
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
    "image/png": ".png",
    "image/webp": ".webp",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
}

# Keyboards
MAIN_REPLY_KEYBOARD = ReplyKeyboardMarkup(
    [
        ["➕ Keyingi mahsulot", "👀 Preview olish"],
        ["✅ Final yuklab olish", "📋 Ro‘yxat"],
        ["🆕 Yangi hujjat", "🗂 Oldingi fayllar"],
    ],
    resize_keyboard=True,
    one_time_keyboard=True,
    is_persistent=False,
    input_field_placeholder="Kerakli amalni tanlang",
)
HIDE_REPLY_KEYBOARD = ReplyKeyboardRemove()

MAIN_INLINE_KEYBOARD = InlineKeyboardMarkup(
    [
        [
            InlineKeyboardButton("➕ Keyingi mahsulot", callback_data="btn_add"),
            InlineKeyboardButton("👀 Preview olish", callback_data="btn_preview"),
        ],
        [
            InlineKeyboardButton("✅ Final yuklab olish", callback_data="btn_finish"),
            InlineKeyboardButton("📋 Ro‘yxat", callback_data="btn_list"),
        ],
        [
            InlineKeyboardButton("🆕 Yangi hujjat", callback_data="btn_new"),
            InlineKeyboardButton("🗂 Oldingi fayllar", callback_data="btn_files"),
        ],
        [
            InlineKeyboardButton("☰ Pastki menyuni ochish", callback_data="btn_menu"),
        ],
    ]
)


def esc(text: object) -> str:
    """Safely escapes user text for Telegram Markdown V1."""
    return escape_markdown(str(text), version=1)


def format_file_size(size_bytes: int) -> str:
    """Formats a byte count for Telegram-facing status messages."""
    size = float(size_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.1f} {unit}" if unit != "B" else f"{int(size)} B"
        size /= 1024
    return f"{size_bytes} B"


def is_request_entity_too_large(exc: Exception) -> bool:
    """Detects Telegram/httpx 413 errors without depending on one exception class."""
    text = str(exc).lower()
    return "request entity too large" in text or "413" in text


def safe_docx_part_size(total_products: int) -> int:
    """Returns the first chunk size to try when a generated DOCX is too large."""
    if total_products <= 0:
        return 1
    configured = min(max(1, MAX_PRODUCTS_PER_DOCX_PART), total_products)
    if configured >= LABELS_PER_PAGE:
        configured = max(LABELS_PER_PAGE, (configured // LABELS_PER_PAGE) * LABELS_PER_PAGE)
    return max(1, configured)


def chunk_products_for_document_parts(products: list, max_products_per_part: int) -> list[list]:
    """Splits products into ordered chunks for multi-part DOCX delivery."""
    chunk_size = max(1, int(max_products_per_part))
    return [list(products[index:index + chunk_size]) for index in range(0, len(products), chunk_size)]


def next_smaller_docx_part_size(current_size: int) -> int:
    """Reduces a DOCX split size while preserving full pages when practical."""
    current_size = max(1, int(current_size))
    if current_size > LABELS_PER_PAGE:
        half = current_size // 2
        return max(LABELS_PER_PAGE, (half // LABELS_PER_PAGE) * LABELS_PER_PAGE)
    if current_size > 1:
        return max(1, current_size // 2)
    return 1


def safe_image_extension(file_name: Optional[str], mime_type: Optional[str]) -> str:
    """Returns a safe image extension for storing uploaded Telegram files."""
    suffix = Path(file_name or "").suffix.lower()
    if suffix in SUPPORTED_IMAGE_EXTENSIONS:
        return suffix
    return IMAGE_MIME_EXTENSIONS.get((mime_type or "").lower(), ".png")


def normalize_webhook_path(path: str) -> str:
    """Normalizes webhook path to '/path' form for Nginx and Telegram."""
    clean = (path or "/telegram-webhook").strip()
    if not clean.startswith("/"):
        clean = f"/{clean}"
    return clean


def build_webhook_public_url(base_url: str, path: str) -> str:
    """Builds the final public HTTPS webhook URL."""
    return f"{base_url.strip().rstrip('/')}{normalize_webhook_path(path)}"


def parse_index_from_callback(data: str) -> Optional[int]:
    """Parses callback data in '<action>:<1-based-index>' format."""
    try:
        return int(str(data).rsplit(":", 1)[1])
    except (IndexError, TypeError, ValueError):
        return None


def format_products_list(session: UserSession) -> str:
    """Builds the current active product list text."""
    if not session.products:
        return (
            f"📋 *Hujjat:* {esc(session.get_document_title())}\n"
            "📭 Hozircha mahsulotlar ro‘yxati bo‘sh."
        )

    lines = [
        f"📋 *Hujjat:* {esc(session.get_document_title())}",
        f"*Kiritilgan mahsulotlar ({len(session.products)} ta):*\n",
    ]
    for i, product in enumerate(session.products, start=1):
        lines.append(f"{i}. *{esc(product.name)}* — `{esc(product.barcode)}`")

    pages = (len(session.products) + 11) // 12
    lines.append(f"\n📄 *Jami betlar:* {pages} ta (A4, 12 ta/bet)")
    lines.append("✏️ Tuzatish uchun pastdagi mahsulot raqamini bosing yoki `/edit 2` yozing.")
    lines.append("↔️ Joyini almashtirish uchun `/swap 2 7` yozing.")
    lines.append("📍 Aniq o‘ringa qo‘yish uchun `/move 7 1` yozing — eski #1 avtomatik #2 ga suriladi.")
    return "\n".join(lines)


def build_products_keyboard(session: UserSession) -> InlineKeyboardMarkup:
    """Builds inline controls for editing products and common actions."""
    buttons = []
    edit_row = []
    for index, product in enumerate(session.products, start=1):
        label = product.name
        if len(label) > 12:
            label = f"{label[:11]}…"
        edit_row.append(InlineKeyboardButton(f"✏️ {index}. {label}", callback_data=f"edit:{index}"))
        if len(edit_row) == 2:
            buttons.append(edit_row)
            edit_row = []
    if edit_row:
        buttons.append(edit_row)

    buttons.extend(MAIN_INLINE_KEYBOARD.inline_keyboard)
    return InlineKeyboardMarkup(buttons)


def build_product_edit_keyboard(index: int) -> InlineKeyboardMarkup:
    """Builds inline field-specific edit controls for one product."""
    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton("🖼 Rasmni tuzatish", callback_data=f"edit_photo:{index}"),
            ],
            [
                InlineKeyboardButton("📝 Nomni tuzatish", callback_data=f"edit_name:{index}"),
                InlineKeyboardButton("🔢 Shtrixni tuzatish", callback_data=f"edit_barcode:{index}"),
            ],
            [
                InlineKeyboardButton("↔️ Joyini almashtirish", callback_data=f"swap:{index}"),
            ],
            [
                InlineKeyboardButton("📍 Aniq o‘ringa qo‘yish", callback_data=f"move:{index}"),
            ],
            [
                InlineKeyboardButton("🗑 O‘chirish", callback_data=f"remove:{index}"),
                InlineKeyboardButton("📋 Ro‘yxatga qaytish", callback_data="btn_list"),
            ],
        ]
    )


async def send_products_list_message(
    *,
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: int,
    session: UserSession,
    intro: Optional[str] = None,
) -> None:
    """Sends the active product list with inline edit controls."""
    text = format_products_list(session)
    if intro:
        text = f"{intro}\n\n{text}"
    await context.bot.send_message(
        chat_id=chat_id,
        text=text,
        reply_markup=build_products_keyboard(session) if session.products else MAIN_INLINE_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )


async def save_incoming_image(update: Update, user_id: int) -> Optional[str]:
    """Downloads and stores an incoming Telegram image, returning its local path."""
    file_bytes = None
    ext = ".jpg"

    if update.message.photo:
        photo = update.message.photo[-1]
        file_obj = await photo.get_file()
        file_bytes = await file_obj.download_as_bytearray()
        ext = ".jpg"
    elif update.message.document and (
        update.message.document.mime_type
        and update.message.document.mime_type.startswith("image/")
    ):
        doc = update.message.document
        file_obj = await doc.get_file()
        file_bytes = await file_obj.download_as_bytearray()
        ext = safe_image_extension(doc.file_name, doc.mime_type)

    if not file_bytes:
        return None
    return session_mgr.save_draft_photo(user_id, bytes(file_bytes), ext=ext)


async def send_prompt(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str, **kwargs) -> None:
    """Sends a prompt from either a normal message or callback query."""
    kwargs.setdefault("reply_markup", HIDE_REPLY_KEYBOARD)
    if update.callback_query:
        await update.callback_query.answer()
        await context.bot.send_message(chat_id=update.effective_chat.id, text=text, **kwargs)
        return
    await update.message.reply_text(text, **kwargs)


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles /start command: explains usage and features."""
    help_text = (
        "🏷️ *Mahsulot yorlig‘i botiga xush kelibsiz!*\n\n"
        "Men A4 formatda *3 × 4 grid* bo‘yicha, ya’ni har betga 12 ta yorliq qilib DOCX/PDF hujjat tayyorlayman.\n"
        "Siz hatto 1 ta mahsulot qo‘shib ham darhol preview yoki final hujjatni yuklab olishingiz mumkin.\n\n"
        "*Har bir yorliqda:*\n"
        "• mahsulot rasmi markazda joylashadi\n"
        "• mahsulot nomi qalin yoziladi\n"
        "• Code128 / EAN-13 barcode yaratiladi\n"
        "• barcode raqami alohida ko‘rsatiladi\n"
        "• kesish uchun yengil chiziqlar bo‘ladi\n\n"
        "📋 *Ish tartibi:*\n"
        "1. /new bosing va hujjat nomini kiriting\n"
        "2. *Rasm* ➔ *Nomi* ➔ *Barcode* yuboring\n"
        "3. Mahsulot avtomatik saqlanadi\n"
        "4. Istalgan vaqtda /preview yoki /finish orqali hujjatni oling\n"
        "5. /files orqali oldingi hujjatlarni qayta yuklab oling.\n\n"
        "⚙️ *Buyruqlar:*\n"
        "/add — Yangi mahsulot qo‘shish\n"
        "/preview — Hozirgi ro‘yxatni ko‘rish uchun DOCX/PDF olish\n"
        "/finish — Final DOCX & PDF fayllarni yuklab olish\n"
        "/files — Oldingi hujjatlar ro‘yxati va qayta yuklash\n"
        "/list — Ro‘yxatni ko‘rish\n"
        "/edit `<N>` — Mahsulot rasmi, nomi yoki shtrixini tuzatish\n"
        "/swap `<A>` `<B>` — Ikki mahsulot joyini almashtirish (masalan: `/swap 2 7`)\n"
        "/move `<A>` `<B>` — A-mahsulotni B-o‘ringa qo‘yish, qolganlarini surish (masalan: `/move 7 1`)\n"
        "/remove `<N>` — Mahsulotni o‘chirish (masalan: `/remove 1`)\n"
        "/menu — Pastki katta menyuni qo‘lda ochish\n"
        "/new — Yangi hujjat boshlash va nom berish\n"
        "/cancel — Joriy kiritishni bekor qilish"
    )
    await update.message.reply_text(
        help_text, reply_markup=HIDE_REPLY_KEYBOARD, parse_mode=ParseMode.MARKDOWN
    )
    await update.message.reply_text(
        "Tezkor amallar:",
        reply_markup=MAIN_INLINE_KEYBOARD,
    )


async def cmd_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Shows the optional bottom reply keyboard only when explicitly requested."""
    await update.message.reply_text(
        "☰ Pastki menyu ochildi. Tugmani bosganingizdan keyin u yana yopiladi.",
        reply_markup=MAIN_REPLY_KEYBOARD,
    )


async def new_document_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Starts a new document: clears active products and asks for document name."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    session.clear()
    session.reset_draft()
    await send_prompt(
        update,
        context,
        "🆕 *Yangi hujjat boshlaymiz.*\n\n"
        "Avval hujjat nomini yuboring. Masalan: `Santexnika perexodlar`.\n\n"
        "Bekor qilish uchun /cancel bosing.",
        parse_mode=ParseMode.MARKDOWN,
    )
    return STATE_WAITING_DOCUMENT_NAME


async def cmd_list(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles /list command: shows collected products in order with safe markdown."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    if update.callback_query:
        await update.callback_query.answer()
    await send_products_list_message(
        context=context,
        chat_id=update.effective_chat.id,
        session=session,
    )


async def cmd_remove(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles /remove N command: removes product number N with safe markdown."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    if not context.args or not context.args[0].isdigit():
        await update.message.reply_text(
            "⚠️ Iltimos, o‘chiriladigan mahsulot raqamini ko‘rsating.\n"
            "Masalan: `/remove 1`.\n"
            "Raqamlarni ko‘rish uchun /list buyrug‘idan foydalaning.",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return

    idx = int(context.args[0])
    total = len(session.products)
    removed = session.remove_product(idx)

    if not removed:
        await update.message.reply_text(
            f"⚠️ #{idx} raqamli mahsulot topilmadi. Sizda jami {total} ta mahsulot mavjud.\n"
            "Ro‘yxatni ko‘rish uchun /list bosing.",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return

    await update.message.reply_text(
        f"🗑️ #{idx} o‘chirildi: *{esc(removed.name)}* (`{esc(removed.barcode)}`).\n"
        f"Qolgan mahsulotlar: {len(session.products)} ta.",
        reply_markup=HIDE_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )
    await send_products_list_message(
        context=context,
        chat_id=update.effective_chat.id,
        session=session,
    )


async def send_swap_result(
    *,
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    session: UserSession,
    first_index: int,
    second_index: int,
) -> int:
    """Swaps two product positions, confirms the result, and shows the refreshed list."""
    total = len(session.products)
    swapped = session.swap_products(first_index, second_index)
    clear_edit_context(context)

    if not swapped:
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=(
                f"⚠️ Mahsulot raqami noto‘g‘ri. Sizda jami {total} ta mahsulot bor.\n"
                "Masalan: `/swap 2 7`"
            ),
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return ConversationHandler.END

    first_product, second_product = swapped
    if first_index == second_index:
        intro = f"ℹ️ #{first_index} va #{second_index} bir xil raqam. Tartib o‘zgarmadi."
    else:
        intro = (
            f"✅ #{first_index} ↔ #{second_index} joyi almashtirildi.\n"
            f"#{first_index} endi: *{esc(second_product.name)}* (`{esc(second_product.barcode)}`)\n"
            f"#{second_index} endi: *{esc(first_product.name)}* (`{esc(first_product.barcode)}`)"
        )

    await send_products_list_message(
        context=context,
        chat_id=update.effective_chat.id,
        session=session,
        intro=intro,
    )
    return ConversationHandler.END


async def cmd_swap(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handles /swap A B command or starts an interactive swap when only A is provided."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    if not session.products:
        await update.message.reply_text(
            "⚠️ Hozircha mahsulotlar ro‘yxati bo‘sh. Avval /add orqali mahsulot qo‘shing.",
            reply_markup=HIDE_REPLY_KEYBOARD,
        )
        return ConversationHandler.END

    if not context.args or len(context.args) > 2 or not all(arg.isdigit() for arg in context.args):
        await update.message.reply_text(
            "⚠️ Ikki mahsulot raqamini yuboring.\n"
            "Masalan: `/swap 2 7`\n\n"
            "Raqamlarni ko‘rish uchun /list bosing.",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return ConversationHandler.END

    first_index = int(context.args[0])
    if len(context.args) == 1:
        product = session.get_product(first_index)
        if not product:
            await update.message.reply_text(
                f"⚠️ #{first_index} raqamli mahsulot topilmadi. Sizda jami {len(session.products)} ta mahsulot bor.",
                reply_markup=HIDE_REPLY_KEYBOARD,
                parse_mode=ParseMode.MARKDOWN,
            )
            return ConversationHandler.END
        context.user_data["swap_source_index"] = first_index
        await update.message.reply_text(
            (
                f"↔️ *#{first_index} mahsulot joyi almashtiriladi:*\n"
                f"*{esc(product.name)}* (`{esc(product.barcode)}`)\n\n"
                f"Qaysi raqam bilan almashtiramiz? 1 dan {len(session.products)} gacha raqam yuboring.\n"
                "Bekor qilish uchun /cancel bosing."
            ),
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_SWAP_TARGET

    return await send_swap_result(
        update=update,
        context=context,
        session=session,
        first_index=first_index,
        second_index=int(context.args[1]),
    )


async def swap_product_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Starts interactive position swapping from a product edit inline button."""
    query = update.callback_query
    index = parse_index_from_callback(query.data if query else "")
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    if not index:
        if query:
            await query.answer("Noto‘g‘ri mahsulot raqami", show_alert=True)
        return ConversationHandler.END

    product = session.get_product(index)
    if not product:
        if query:
            await query.answer("Mahsulot topilmadi", show_alert=True)
        return ConversationHandler.END

    context.user_data["swap_source_index"] = index
    await send_prompt(
        update,
        context,
        (
            f"↔️ *#{index} mahsulot joyi almashtiriladi:*\n"
            f"*{esc(product.name)}* (`{esc(product.barcode)}`)\n\n"
            f"Qaysi raqam bilan almashtiramiz? 1 dan {len(session.products)} gacha raqam yuboring.\n"
            "Masalan: `7`\n\n"
            "Bekor qilish uchun /cancel bosing."
        ),
        parse_mode=ParseMode.MARKDOWN,
    )
    return STATE_WAITING_SWAP_TARGET


async def swap_target_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Receives the second index for an interactive product position swap."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    source_index = int(context.user_data.get("swap_source_index") or 0)
    raw_text = (update.message.text or "").strip()
    tokens = raw_text.split()
    first_token = tokens[0] if tokens else ""

    if not source_index or not session.get_product(source_index):
        clear_edit_context(context)
        await update.message.reply_text(
            "⚠️ Tanlangan mahsulot topilmadi. Iltimos, /list orqali qaytadan tanlang.",
            reply_markup=HIDE_REPLY_KEYBOARD,
        )
        return ConversationHandler.END

    if not first_token.isdigit():
        await update.message.reply_text(
            f"⚠️ Faqat raqam yuboring. Masalan: `7`.\n"
            f"Sizda jami {len(session.products)} ta mahsulot bor.",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_SWAP_TARGET

    return await send_swap_result(
        update=update,
        context=context,
        session=session,
        first_index=source_index,
        second_index=int(first_token),
    )


async def send_move_result(
    *,
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    session: UserSession,
    source_index: int,
    target_index: int,
) -> int:
    """Moves one product to an exact position, confirms the result, and shows the refreshed list."""
    total = len(session.products)
    moved = session.move_product_to_position(source_index, target_index)
    clear_edit_context(context)

    if not moved:
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=(
                f"⚠️ Mahsulot raqami noto‘g‘ri. Sizda jami {total} ta mahsulot bor.\n"
                "Masalan: `/move 7 1`"
            ),
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return ConversationHandler.END

    product, old_position, new_position = moved
    if old_position == new_position:
        intro = f"ℹ️ #{old_position} mahsulot shu o‘rinda turibdi. Tartib o‘zgarmadi."
    else:
        intro = (
            f"✅ #{old_position} mahsulot #{new_position}-o‘ringa qo‘yildi.\n"
            f"📍 #{new_position} endi: *{esc(product.name)}* (`{esc(product.barcode)}`)\n"
            "Qolgan mahsulotlar avtomatik tartib bilan surildi."
        )

    await send_products_list_message(
        context=context,
        chat_id=update.effective_chat.id,
        session=session,
        intro=intro,
    )
    return ConversationHandler.END


async def cmd_move(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handles /move A B command or starts an interactive exact-position move."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    if not session.products:
        await update.message.reply_text(
            "⚠️ Hozircha mahsulotlar ro‘yxati bo‘sh. Avval /add orqali mahsulot qo‘shing.",
            reply_markup=HIDE_REPLY_KEYBOARD,
        )
        return ConversationHandler.END

    if not context.args or len(context.args) > 2 or not all(arg.isdigit() for arg in context.args):
        await update.message.reply_text(
            "⚠️ Qaysi mahsulotni qaysi o‘ringa qo‘yishni yuboring.\n"
            "Masalan: `/move 7 1` — #7 mahsulot #1-o‘ringa tushadi, eski #1 #2 ga suriladi.\n\n"
            "Raqamlarni ko‘rish uchun /list bosing.",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return ConversationHandler.END

    source_index = int(context.args[0])
    if len(context.args) == 1:
        product = session.get_product(source_index)
        if not product:
            await update.message.reply_text(
                f"⚠️ #{source_index} raqamli mahsulot topilmadi. Sizda jami {len(session.products)} ta mahsulot bor.",
                reply_markup=HIDE_REPLY_KEYBOARD,
                parse_mode=ParseMode.MARKDOWN,
            )
            return ConversationHandler.END
        context.user_data["move_source_index"] = source_index
        await update.message.reply_text(
            (
                f"📍 *#{source_index} mahsulot aniq o‘ringa qo‘yiladi:*\n"
                f"*{esc(product.name)}* (`{esc(product.barcode)}`)\n\n"
                f"Qaysi o‘ringa qo‘yamiz? 1 dan {len(session.products)} gacha raqam yuboring.\n"
                "Masalan: `1` — shu mahsulot birinchi bo‘ladi, qolganlari pastga suriladi.\n\n"
                "Bekor qilish uchun /cancel bosing."
            ),
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_MOVE_TARGET

    return await send_move_result(
        update=update,
        context=context,
        session=session,
        source_index=source_index,
        target_index=int(context.args[1]),
    )


async def move_product_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Starts interactive exact-position movement from a product edit inline button."""
    query = update.callback_query
    index = parse_index_from_callback(query.data if query else "")
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    if not index:
        if query:
            await query.answer("Noto‘g‘ri mahsulot raqami", show_alert=True)
        return ConversationHandler.END

    product = session.get_product(index)
    if not product:
        if query:
            await query.answer("Mahsulot topilmadi", show_alert=True)
        return ConversationHandler.END

    context.user_data["move_source_index"] = index
    await send_prompt(
        update,
        context,
        (
            f"📍 *#{index} mahsulot aniq o‘ringa qo‘yiladi:*\n"
            f"*{esc(product.name)}* (`{esc(product.barcode)}`)\n\n"
            f"Qaysi o‘ringa qo‘yamiz? 1 dan {len(session.products)} gacha raqam yuboring.\n"
            "Masalan: `1` — shu mahsulot birinchi bo‘ladi, qolganlari pastga suriladi.\n\n"
            "Bekor qilish uchun /cancel bosing."
        ),
        parse_mode=ParseMode.MARKDOWN,
    )
    return STATE_WAITING_MOVE_TARGET


async def move_target_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Receives the target index for an interactive exact-position move."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    source_index = int(context.user_data.get("move_source_index") or 0)
    raw_text = (update.message.text or "").strip()
    tokens = raw_text.split()
    first_token = tokens[0] if tokens else ""

    if not source_index or not session.get_product(source_index):
        clear_edit_context(context)
        await update.message.reply_text(
            "⚠️ Tanlangan mahsulot topilmadi. Iltimos, /list orqali qaytadan tanlang.",
            reply_markup=HIDE_REPLY_KEYBOARD,
        )
        return ConversationHandler.END

    if not first_token.isdigit():
        await update.message.reply_text(
            f"⚠️ Faqat raqam yuboring. Masalan: `1`.\n"
            f"Sizda jami {len(session.products)} ta mahsulot bor.",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_MOVE_TARGET

    return await send_move_result(
        update=update,
        context=context,
        session=session,
        source_index=source_index,
        target_index=int(first_token),
    )


async def show_product_edit_options(
    *,
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    index: int,
) -> None:
    """Shows field-specific edit options for one product."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    product = session.get_product(index)

    if not product:
        text = f"⚠️ #{index} raqamli mahsulot topilmadi. Ro‘yxatni yangilash uchun /list bosing."
        if update.callback_query:
            await update.callback_query.answer("Mahsulot topilmadi", show_alert=True)
            await context.bot.send_message(
                chat_id=update.effective_chat.id,
                text=text,
                reply_markup=MAIN_INLINE_KEYBOARD,
            )
        else:
            await update.message.reply_text(
                text,
                reply_markup=HIDE_REPLY_KEYBOARD,
            )
        return

    text = (
        f"✏️ *#{index} mahsulotni tuzatish*\n\n"
        f"📦 Nomi: *{esc(product.name)}*\n"
        f"🔢 Shtrix: `{esc(product.barcode)}`\n\n"
        "Qaysi qismini tuzatmoqchisiz?"
    )
    if update.callback_query:
        await update.callback_query.answer()
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=text,
            reply_markup=build_product_edit_keyboard(index),
            parse_mode=ParseMode.MARKDOWN,
        )
    else:
        await update.message.reply_text(
            text,
            reply_markup=build_product_edit_keyboard(index),
            parse_mode=ParseMode.MARKDOWN,
        )


async def cmd_edit(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles /edit N command by showing field-specific edit options."""
    if not context.args or not context.args[0].isdigit():
        await update.message.reply_text(
            "⚠️ Qaysi mahsulotni tuzatishni raqam bilan yuboring.\n"
            "Masalan: `/edit 2`.\n\n"
            "Raqamlarni ko‘rish uchun /list bosing.",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return

    await show_product_edit_options(
        update=update,
        context=context,
        index=int(context.args[0]),
    )


async def start_edit_field(update: Update, context: ContextTypes.DEFAULT_TYPE, field: str) -> int:
    """Starts an edit conversation for one product field."""
    query = update.callback_query
    index = parse_index_from_callback(query.data if query else "")
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    if not index or not session.get_product(index):
        if query:
            await query.answer("Mahsulot topilmadi", show_alert=True)
        return ConversationHandler.END

    context.user_data["edit_index"] = index
    context.user_data["edit_field"] = field

    if field == "photo":
        text = (
            f"🖼 *#{index} mahsulot rasmi tuzatilyapti.*\n\n"
            "Yangi rasmni yuboring. Bekor qilish uchun /cancel bosing."
        )
        next_state = STATE_WAITING_EDIT_PHOTO
    elif field == "name":
        text = (
            f"📝 *#{index} mahsulot nomi tuzatilyapti.*\n\n"
            "Yangi nomni yuboring. Bekor qilish uchun /cancel bosing."
        )
        next_state = STATE_WAITING_EDIT_NAME
    else:
        text = (
            f"🔢 *#{index} mahsulot shtrix-kodi tuzatilyapti.*\n\n"
            "Yangi barcode/kodni yuboring. Bekor qilish uchun /cancel bosing."
        )
        next_state = STATE_WAITING_EDIT_BARCODE

    await send_prompt(update, context, text, parse_mode=ParseMode.MARKDOWN)
    return next_state


async def edit_photo_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Starts product image edit."""
    return await start_edit_field(update, context, "photo")


async def edit_name_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Starts product name edit."""
    return await start_edit_field(update, context, "name")


async def edit_barcode_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Starts product barcode edit."""
    return await start_edit_field(update, context, "barcode")


def clear_edit_context(context: ContextTypes.DEFAULT_TYPE) -> None:
    """Clears edit state stored in Telegram user_data."""
    context.user_data.pop("edit_index", None)
    context.user_data.pop("edit_field", None)
    context.user_data.pop("swap_source_index", None)
    context.user_data.pop("move_source_index", None)


async def finish_product_edit(
    *,
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    session: UserSession,
    message: str,
) -> int:
    """Sends edit confirmation and refreshed list."""
    clear_edit_context(context)
    await update.message.reply_text(
        message,
        reply_markup=HIDE_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )
    await send_products_list_message(
        context=context,
        chat_id=update.effective_chat.id,
        session=session,
    )
    return ConversationHandler.END


async def edit_photo_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Receives and saves a replacement product photo."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    index = int(context.user_data.get("edit_index") or 0)

    if not session.get_product(index):
        clear_edit_context(context)
        await update.message.reply_text("⚠️ Mahsulot topilmadi. Iltimos, /list orqali qaytadan tanlang.")
        return ConversationHandler.END

    photo_path = await save_incoming_image(update, user_id)
    if not photo_path:
        await update.message.reply_text(
            "⚠️ Iltimos, yangi mahsulot rasmini yuboring, yoki /cancel bosing.",
            reply_markup=HIDE_REPLY_KEYBOARD,
        )
        return STATE_WAITING_EDIT_PHOTO

    updated = session.update_product_image(index, photo_path)
    if not updated:
        clear_edit_context(context)
        await update.message.reply_text("⚠️ Mahsulot topilmadi. Iltimos, /list orqali qaytadan tanlang.")
        return ConversationHandler.END

    return await finish_product_edit(
        update=update,
        context=context,
        session=session,
        message=f"✅ #{index} mahsulot rasmi yangilandi: *{esc(updated.name)}*.",
    )


async def edit_name_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Receives and saves a replacement product name."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    index = int(context.user_data.get("edit_index") or 0)
    new_name = update.message.text.strip()

    if not new_name:
        await update.message.reply_text(
            "⚠️ Mahsulot nomi bo‘sh bo‘lishi mumkin emas. Yangi nomni yuboring:",
            reply_markup=HIDE_REPLY_KEYBOARD,
        )
        return STATE_WAITING_EDIT_NAME

    updated = session.update_product_name(index, new_name)
    if not updated:
        clear_edit_context(context)
        await update.message.reply_text("⚠️ Mahsulot topilmadi. Iltimos, /list orqali qaytadan tanlang.")
        return ConversationHandler.END

    return await finish_product_edit(
        update=update,
        context=context,
        session=session,
        message=f"✅ #{index} mahsulot nomi yangilandi: *{esc(updated.name)}*.",
    )


async def edit_barcode_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Receives, validates, and saves a replacement barcode/code value."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    index = int(context.user_data.get("edit_index") or 0)
    new_barcode = update.message.text.strip()

    if not new_barcode:
        await update.message.reply_text(
            "⚠️ Barcode bo‘sh bo‘lishi mumkin emas. Yangi kodni yuboring:",
            reply_markup=HIDE_REPLY_KEYBOARD,
        )
        return STATE_WAITING_EDIT_BARCODE

    temp_check = session.session_dir / "temp_check.png"
    try:
        generate_barcode_image(new_barcode, temp_check)
        if temp_check.is_file():
            temp_check.unlink(missing_ok=True)
    except Exception as e:
        await update.message.reply_text(
            f"⚠️ Barcode hosil qilib bo‘lmadi `{esc(new_barcode)}`: {esc(e)}\n"
            "Iltimos, yaroqli kod yuboring:",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_EDIT_BARCODE

    updated = session.update_product_barcode(index, new_barcode)
    if not updated:
        clear_edit_context(context)
        await update.message.reply_text("⚠️ Mahsulot topilmadi. Iltimos, /list orqali qaytadan tanlang.")
        return ConversationHandler.END

    return await finish_product_edit(
        update=update,
        context=context,
        session=session,
        message=f"✅ #{index} mahsulot shtrixi yangilandi: `{esc(updated.barcode)}`.",
    )


async def remove_product_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Removes a product from an inline edit button."""
    query = update.callback_query
    index = parse_index_from_callback(query.data if query else "")
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    if not index:
        await query.answer("Noto‘g‘ri mahsulot raqami", show_alert=True)
        return

    removed = session.remove_product(index)
    if not removed:
        await query.answer("Mahsulot topilmadi", show_alert=True)
        return

    await query.answer("O‘chirildi")
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=f"🗑️ #{index} o‘chirildi: *{esc(removed.name)}* (`{esc(removed.barcode)}`).",
        reply_markup=HIDE_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )
    await send_products_list_message(
        context=context,
        chat_id=update.effective_chat.id,
        session=session,
    )


# --- /add Conversation Handlers ---

async def ask_product_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Prompts the user for the product photo."""
    text = (
        "📸 *1/3-qadam: Mahsulot fotosi*\n\n"
        "Iltimos, mahsulot rasmini yuboring.\n"
        "Mahsulotni oq yoki yengil fonda, yaxshi yorug‘likda, kadr markaziga olib rasmga oling.\n"
        "_Oq fon avtomatik qirqiladi. Sifat yaxshi bo‘lishi uchun rasmni file/document sifatida yuborishingiz ham mumkin._\n\n"
        "Bekor qilish uchun /cancel bosing."
    )
    await send_prompt(update, context, text, parse_mode=ParseMode.MARKDOWN)
    return STATE_WAITING_PHOTO


async def add_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Starts the /add flow: requests product photo."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    session.reset_draft()

    if not session.document_name:
        await send_prompt(
            update,
            context,
            "📝 *Hujjat nomi kerak.*\n\n"
            "Bu hujjatga nom bering. Masalan: `Santexnika perexodlar`.\n"
            "Keyin mahsulot rasmini qabul qilaman.\n\n"
            "Bekor qilish uchun /cancel bosing.",
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_DOCUMENT_NAME

    await send_products_list_message(
        context=context,
        chat_id=update.effective_chat.id,
        session=session,
        intro="➕ *Yangi mahsulot qo‘shishdan oldin hozirgi ro‘yxat:*",
    )
    return await ask_product_photo(update, context)


async def document_name_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Stores the active document name before collecting the first product."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    document_name = normalize_document_name(update.message.text)
    if not document_name:
        await update.message.reply_text(
            "⚠️ Hujjat nomi bo‘sh bo‘lishi mumkin emas. Iltimos, nom yuboring:",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_DOCUMENT_NAME

    session.set_document_name(document_name)
    session.reset_draft()
    await update.message.reply_text(
        f"✅ Hujjat nomi saqlandi: *{esc(document_name)}*\n\n"
        "Endi birinchi mahsulot rasmini yuboring.",
        reply_markup=HIDE_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )
    return STATE_WAITING_PHOTO


async def add_photo_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Processes uploaded photo."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    photo_path = await save_incoming_image(update, user_id)
    if not photo_path:
        await update.message.reply_text(
            "⚠️ Iltimos, mahsulot rasmini yuboring, yoki /cancel bosing.",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_PHOTO

    session.current_draft["photo_path"] = photo_path

    await update.message.reply_text(
        "✅ Rasm qabul qilindi!\n\n"
        "📝 *2/3-qadam: Mahsulot nomi*\n"
        "Iltimos, mahsulot nomini yuboring (masalan: `Perexod 76-50mm`):",
        reply_markup=HIDE_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )
    return STATE_WAITING_NAME


async def add_name_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Processes product name."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    name = update.message.text.strip()
    if not name:
        await update.message.reply_text(
            "⚠️ Mahsulot nomi bo‘sh bo‘lishi mumkin emas. Iltimos, nomni yuboring:",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_NAME

    session.current_draft["name"] = name

    await update.message.reply_text(
        f"✅ Nom saqlandi: *{esc(name)}*\n\n"
        "🔢 *3/3-qadam: Barcode / Kod qiymati*\n"
        "Iltimos, shtrix-kod yoki mahsulot kodini yuboring (masalan: `1000049` yoki 13 xonali EAN):",
        reply_markup=HIDE_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )
    return STATE_WAITING_BARCODE


async def add_barcode_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Processes barcode value and completes product addition automatically."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    barcode_val = update.message.text.strip()
    if not barcode_val:
        await update.message.reply_text(
            "⚠️ Barcode bo‘sh bo‘lishi mumkin emas. Iltimos, kodni yuboring:",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_BARCODE

    # Verify barcode can be rendered cleanly
    temp_check = session.session_dir / "temp_check.png"
    try:
        generate_barcode_image(barcode_val, temp_check)
        if temp_check.is_file():
            temp_check.unlink(missing_ok=True)
    except Exception as e:
        await update.message.reply_text(
            f"⚠️ Barcode hosil qilib bo‘lmadi `{esc(barcode_val)}`: {esc(e)}\n"
            "Iltimos, yaroqli kod yuboring:",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_BARCODE

    # Validate that photo and name exist in draft
    photo_path = session.current_draft.get("photo_path")
    prod_name = session.current_draft.get("name")
    if not photo_path or not prod_name:
        await update.message.reply_text(
            "⚠️ Ma’lumotlar to‘liq emas. Iltimos, /add orqali qaytadan boshlang.",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        session.reset_draft()
        return ConversationHandler.END

    session.current_draft["barcode"] = barcode_val

    # Automatically commit product
    count = session.add_product(name=prod_name, barcode=barcode_val, image_path=photo_path)
    session.reset_draft()

    # Confirmation message as requested by user
    confirmation_text = (
        f"✅ Mahsulot #{count} saqlandi.\n"
        f"📄 Hujjat: *{esc(session.get_document_title())}*\n"
        f"📦 Nomi: *{esc(prod_name)}*\n"
        f"🔢 Barcode: `{esc(barcode_val)}`\n\n"
        "Keyingi amalni tanlang:\n"
        "➕ /add — Keyingi mahsulotni qo‘shish\n"
        "👀 /preview — Hozirgi holatni tekshirib ko‘rish uchun DOCX/PDF olish\n"
        "✅ /finish — Hozirgi ro‘yxatni final hujjat qilib yuklab olish\n"
        "📋 /list — Ro‘yxatni ko‘rish"
    )

    await update.message.reply_text(
        confirmation_text,
        reply_markup=HIDE_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )
    # Also attach inline keyboard for quick 1-tap action
    await update.message.reply_text(
        "Tezkor amallar:",
        reply_markup=MAIN_INLINE_KEYBOARD,
    )
    return ConversationHandler.END


async def cmd_done_in_conv(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handles /done inside conversation: informs user that save is automatic."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    draft = session.current_draft

    if not draft.get("photo_path"):
        next_state = STATE_WAITING_PHOTO
        next_step = "hozir mahsulot rasmini yuboring."
    elif not draft.get("name"):
        next_state = STATE_WAITING_NAME
        next_step = "hozir mahsulot nomini yuboring."
    else:
        next_state = STATE_WAITING_BARCODE
        next_step = "hozir barcode yoki mahsulot kodini yuboring."

    await update.message.reply_text(
        "ℹ️ `/done` buyrug‘i kerak emas! Barcode kiritilishi bilanoq mahsulot avtomatik saqlanadi.\n"
        f"Iltimos, {next_step} Bekor qilish uchun /cancel bosing.",
        reply_markup=HIDE_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )
    return next_state


async def cmd_done_global(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles /done outside of conversation: explains that save is automatic."""
    await update.message.reply_text(
        "ℹ️ `/done` buyrug‘i kerak emas!\n"
        "Mahsulot barcode kiritilishi bilanoq avtomatik saqlanadi.\n\n"
        "• /add — Yangi mahsulot qo‘shish\n"
        "• /preview — Tekshirib ko‘rish\n"
        "• /finish — Hujjatni yuklab olish",
        reply_markup=HIDE_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )


async def cmd_cancel_in_conv(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Cancels current /add flow inside conversation."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    session.reset_draft()
    clear_edit_context(context)
    await update.message.reply_text(
        "❌ Mahsulot kiritish bekor qilindi.\n"
        "Yangi mahsulot kiritish uchun /add bosing yoki mavjudlarini olish uchun /finish bosing.",
        reply_markup=HIDE_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )
    return ConversationHandler.END


async def cmd_cancel_global(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles /cancel outside of conversation."""
    clear_edit_context(context)
    await update.message.reply_text(
        "ℹ️ Hozirda bekor qilinadigan jarayon yo‘q.\n"
        "Mahsulot qo‘shish uchun /add yoki ro‘yxatni ko‘rish uchun /list bosing.",
        reply_markup=HIDE_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )


# --- Document Generation & Delivery ---

def export_kind_label(kind: str) -> str:
    """Human-readable export kind label."""
    return "Final" if kind == "final" else "Preview"


def format_export_created_at(record: ExportRecord) -> str:
    """Formats stored ISO datetime safely for Telegram messages."""
    if not record.created_at:
        return "vaqt noma’lum"
    return record.created_at.replace("T", " ")[:16]


def format_export_details(record: ExportRecord) -> str:
    """Formats product/page counts, including legacy files without metadata."""
    details = []
    if record.product_count > 0:
        details.append(f"{record.product_count} ta mahsulot")
    if record.page_count > 0:
        details.append(f"{record.page_count} bet")
    return " | ".join(details) if details else "ma’lumot noma’lum"


def build_exports_keyboard(records: list[ExportRecord]) -> InlineKeyboardMarkup:
    """Builds inline keyboard for downloading previous exports."""
    buttons = []
    for index, record in enumerate(records, start=1):
        title = record.title
        if len(title) > 28:
            title = f"{title[:25]}..."
        buttons.append([
            InlineKeyboardButton(
                f"📥 {index}. {title}",
                callback_data=f"file:{record.export_id}",
            )
        ])
    buttons.append([InlineKeyboardButton("🆕 Yangi hujjat", callback_data="btn_new")])
    return InlineKeyboardMarkup(buttons)


async def send_document_file_with_limit(
    *,
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: int,
    path: Path,
    filename: str,
    caption: str,
) -> None:
    """Sends one document only when it is safely below Telegram's request limit."""
    file_size = path.stat().st_size
    if file_size > TELEGRAM_SAFE_DOCUMENT_SIZE_BYTES:
        raise TelegramDocumentTooLargeError(path, file_size)

    try:
        with open(path, "rb") as f_doc:
            await context.bot.send_document(
                chat_id=chat_id,
                document=f_doc,
                filename=filename,
                caption=caption,
                parse_mode=ParseMode.MARKDOWN,
            )
    except Exception as exc:
        if is_request_entity_too_large(exc):
            raise TelegramDocumentTooLargeError(path, file_size, original_error=exc) from exc
        raise


def split_docx_path(base_docx_path: Path, part_index: int, part_count: int) -> Path:
    """Builds a stable split DOCX filename next to the original export."""
    return base_docx_path.with_name(
        f"{base_docx_path.stem}_part{part_index:02d}of{part_count:02d}{base_docx_path.suffix}"
    )


def remove_stale_split_files(base_docx_path: Path) -> None:
    """Removes split files from previous retries for the same generated export name."""
    for stale in base_docx_path.parent.glob(f"{base_docx_path.stem}_part*"):
        if stale.is_file() and UserSession._is_within(stale, base_docx_path.parent):
            stale.unlink(missing_ok=True)


async def create_split_export_records(
    *,
    session: UserSession,
    base_docx_path: Path,
    products: list,
    kind: str,
) -> list[ExportRecord]:
    """
    Creates ordered split DOCX/PDF export records. It retries with smaller
    chunks if any generated part still exceeds the Telegram-safe size.
    """
    part_size = safe_docx_part_size(len(products))
    remove_stale_split_files(base_docx_path)

    while True:
        chunks = chunk_products_for_document_parts(products, part_size)
        generated_parts = []
        too_large = False

        for idx, chunk in enumerate(chunks, start=1):
            part_docx = split_docx_path(base_docx_path, idx, len(chunks))
            await asyncio.to_thread(create_label_sheet, chunk, part_docx)
            if part_docx.stat().st_size > TELEGRAM_SAFE_DOCUMENT_SIZE_BYTES:
                too_large = True
                part_docx.unlink(missing_ok=True)
                break

            part_pdf, _part_pdf_note = await asyncio.to_thread(convert_docx_to_pdf, part_docx)
            if part_pdf and part_pdf.is_file() and part_pdf.stat().st_size > TELEGRAM_SAFE_DOCUMENT_SIZE_BYTES:
                part_pdf.unlink(missing_ok=True)
                part_pdf = None

            generated_parts.append(
                {
                    "index": idx,
                    "count": len(chunks),
                    "chunk": chunk,
                    "docx": part_docx,
                    "pdf": part_pdf if part_pdf and part_pdf.is_file() else None,
                }
            )

        if not too_large:
            records_by_index = {}
            # Insert records from the last part to the first so /files lists
            # part 1 before part 2.
            for part in reversed(generated_parts):
                title = f"{session.get_document_title()} qism {part['index']}/{part['count']}"
                records_by_index[part["index"]] = session.record_export(
                    kind=kind,
                    docx_path=part["docx"],
                    pdf_path=part["pdf"],
                    product_count=len(part["chunk"]),
                    page_count=(len(part["chunk"]) + LABELS_PER_PAGE - 1) // LABELS_PER_PAGE,
                    title=title,
                )
            return [records_by_index[index] for index in range(1, len(generated_parts) + 1)]

        for part in generated_parts:
            Path(part["docx"]).unlink(missing_ok=True)
            if part["pdf"]:
                Path(part["pdf"]).unlink(missing_ok=True)

        next_part_size = next_smaller_docx_part_size(part_size)
        if next_part_size == part_size:
            raise TelegramDocumentTooLargeError(part_docx, part_docx.stat().st_size)
        part_size = next_part_size


async def send_export_record_docx(
    *,
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: int,
    record: ExportRecord,
    caption: str,
) -> None:
    """Sends the DOCX file for one export record with a size guard."""
    await send_document_file_with_limit(
        context=context,
        chat_id=chat_id,
        path=Path(record.docx_path),
        filename=Path(record.docx_path).name,
        caption=caption,
    )


async def send_export_record_documents(
    *,
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: int,
    session: UserSession,
    record: ExportRecord,
) -> None:
    """Sends stored DOCX/PDF files for an export record."""
    sent_any = False
    kind_label = export_kind_label(record.kind)
    title = esc(record.title)
    allowed_export_dirs = [session.generated_dir, session.session_dir]

    docx_path = Path(record.docx_path)
    if docx_path.is_file() and any(UserSession._is_within(docx_path, base_dir) for base_dir in allowed_export_dirs):
        try:
            await send_export_record_docx(
                context=context,
                chat_id=chat_id,
                record=record,
                caption=(
                    f"📄 *{kind_label} DOCX:* {title}\n"
                    f"• {esc(format_export_details(record))}\n"
                    f"• Hajmi: {esc(format_file_size(docx_path.stat().st_size))}"
                ),
            )
            sent_any = True
        except TelegramDocumentTooLargeError as exc:
            await context.bot.send_message(
                chat_id=chat_id,
                text=(
                    f"⚠️ Bu DOCX Telegram uchun juda katta: {esc(format_file_size(exc.size_bytes))}.\n"
                    f"Limitdan xavfsiz past yuborish chegarasi: {TELEGRAM_SAFE_DOCUMENT_SIZE_MB} MB.\n\n"
                    "Yangi /preview yoki /finish bosing — yangi versiyada rasmlar ixchamlanadi "
                    "va kerak bo‘lsa hujjat avtomatik qismlarga bo‘linadi."
                ),
                reply_markup=HIDE_REPLY_KEYBOARD,
                parse_mode=ParseMode.MARKDOWN,
            )

    if record.pdf_path:
        pdf_path = Path(record.pdf_path)
        if pdf_path.is_file() and any(UserSession._is_within(pdf_path, base_dir) for base_dir in allowed_export_dirs):
            try:
                await send_document_file_with_limit(
                    context=context,
                    chat_id=chat_id,
                    path=pdf_path,
                    filename=pdf_path.name,
                    caption=(
                        f"📑 *{kind_label} PDF:* {title}\n"
                        "• Chop etishga tayyor A4 format."
                    ),
                )
                sent_any = True
            except TelegramDocumentTooLargeError as exc:
                await context.bot.send_message(
                    chat_id=chat_id,
                    text=f"⚠️ Bu PDF Telegram uchun juda katta: {esc(format_file_size(exc.size_bytes))}.",
                    reply_markup=HIDE_REPLY_KEYBOARD,
                    parse_mode=ParseMode.MARKDOWN,
                )

    if not sent_any:
        await context.bot.send_message(
            chat_id=chat_id,
            text="⚠️ Bu hujjat fayllari topilmadi. U o‘chirilgan yoki 30 kunlik muddatdan o‘tgan bo‘lishi mumkin.",
            reply_markup=HIDE_REPLY_KEYBOARD,
        )


async def _generate_and_send_documents(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    session: UserSession,
    is_finish: bool,
) -> None:
    """Generates unique DOCX and PDF files non-blockingly, then sends them to user."""
    chat_id = update.effective_chat.id

    if not session.products:
        await context.bot.send_message(
            chat_id=chat_id,
            text="⚠️ Hozircha hech qanday mahsulot qo‘shilmagan.\nKamida bitta mahsulot qo‘shish uchun /add bosing.",
            reply_markup=HIDE_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return

    status_msg = await context.bot.send_message(
        chat_id=chat_id,
        text=f"⏳ *{esc(session.get_document_title())}* hujjati tayyorlanmoqda...",
        parse_mode=ParseMode.MARKDOWN,
    )

    prefix = "product_labels" if is_finish else "labels_preview"
    docx_path = session.new_export_docx_path(prefix)

    try:
        # Offload file generation to worker thread to avoid blocking bot event loop
        await asyncio.to_thread(create_label_sheet, session.products, docx_path)
    except Exception as exc:
        logger.exception("Error creating label sheet DOCX: %s", exc)
        await status_msg.edit_text(f"❌ DOCX yaratishda xatolik yuz berdi: {esc(exc)}")
        return

    page_count = (len(session.products) + LABELS_PER_PAGE - 1) // LABELS_PER_PAGE
    export_kind = "final" if is_finish else "preview"
    export_records: list[ExportRecord] = []
    split_mode = False
    pdf_note = None

    if docx_path.stat().st_size > TELEGRAM_SAFE_DOCUMENT_SIZE_BYTES:
        split_mode = True
        await status_msg.edit_text(
            (
                f"⚠️ DOCX hajmi katta: {esc(format_file_size(docx_path.stat().st_size))}.\n"
                "Telegram limitidan oshmasligi uchun hujjat qismlarga bo‘linmoqda..."
            ),
            parse_mode=ParseMode.MARKDOWN,
        )
        try:
            export_records = await create_split_export_records(
                session=session,
                base_docx_path=docx_path,
                products=list(session.products),
                kind=export_kind,
            )
            docx_path.unlink(missing_ok=True)
        except Exception as exc:
            logger.exception("Error creating split DOCX files: %s", exc)
            await status_msg.edit_text(
                (
                    f"❌ DOCX juda katta va qismlarga bo‘lishda xatolik yuz berdi: {esc(exc)}\n\n"
                    "Iltimos, rasmlarni biroz yengilroq qilib yuboring yoki mahsulotlarni kamroq qism bilan final qiling."
                ),
                parse_mode=ParseMode.MARKDOWN,
            )
            return
    else:
        # Offload PDF conversion to worker thread for the normal one-file export.
        pdf_path, pdf_note = await asyncio.to_thread(convert_docx_to_pdf, docx_path)
        safe_pdf_path = None
        if pdf_path and pdf_path.is_file() and pdf_path.stat().st_size <= TELEGRAM_SAFE_DOCUMENT_SIZE_BYTES:
            safe_pdf_path = pdf_path
        export_records = [
            session.record_export(
                kind=export_kind,
                docx_path=docx_path,
                pdf_path=safe_pdf_path,
                product_count=len(session.products),
                page_count=page_count,
            )
        ]

    # Send generated DOCX/PDF file(s)
    for record_index, export_record in enumerate(export_records, start=1):
        docx_file = Path(export_record.docx_path)
        part_suffix = f" | qism {record_index}/{len(export_records)}" if split_mode else ""
        try:
            await send_export_record_docx(
                context=context,
                chat_id=chat_id,
                record=export_record,
                caption=(
                    f"📄 *{export_kind_label(export_record.kind)} DOCX:* {esc(export_record.title)}\n"
                    f"• {export_record.product_count} ta mahsulot | {export_record.page_count} bet{esc(part_suffix)}\n"
                    f"• Hajmi: {esc(format_file_size(docx_file.stat().st_size))}\n"
                    f"• A4 format, 12 ta yorliq (3 ustun × 4 qator)"
                ),
            )
        except TelegramDocumentTooLargeError as exc:
            logger.exception("DOCX still too large after size guard: %s", exc)
            await status_msg.edit_text(
                (
                    f"❌ DOCX hali ham Telegram uchun katta: {esc(format_file_size(exc.size_bytes))}.\n"
                    "Bot faylni serverda saqladi, lekin Telegram orqali yuborib bo‘lmadi."
                ),
                parse_mode=ParseMode.MARKDOWN,
            )
            return
        except Exception as exc:
            logger.exception("Error sending DOCX file: %s", exc)
            await status_msg.edit_text(f"❌ DOCX yuborishda xatolik: {esc(exc)}")
            return

        if export_record.pdf_path:
            pdf_file = Path(export_record.pdf_path)
            try:
                await send_document_file_with_limit(
                    context=context,
                    chat_id=chat_id,
                    path=pdf_file,
                    filename=pdf_file.name,
                    caption=(
                        f"📑 *{export_kind_label(export_record.kind)} PDF:* {esc(export_record.title)}\n"
                        "• A4 formatda to‘g‘ridan-to‘g‘ri chop etish uchun."
                    ),
                )
            except TelegramDocumentTooLargeError as exc:
                logger.warning("PDF too large to send: %s", exc)
                await context.bot.send_message(
                    chat_id=chat_id,
                    text=f"⚠️ PDF juda katta bo‘lgani uchun yuborilmadi: {esc(format_file_size(exc.size_bytes))}.",
                    parse_mode=ParseMode.MARKDOWN,
                )
            except Exception as exc:
                logger.warning("Error sending PDF file: %s", exc)

    # If PDF wasn't generated and a note is available
    if not split_mode and pdf_note:
        await context.bot.send_message(
            chat_id=chat_id,
            text=f"ℹ️ *PDF ma’lumoti:*\n{esc(pdf_note)}\n\n*DOCX faylingiz* tayyor va uni bemalol chop etishingiz mumkin.",
            parse_mode=ParseMode.MARKDOWN,
        )

    result_title = export_records[0].title if export_records else session.get_document_title()
    split_notice = f"\nDOCX {len(export_records)} qismga bo‘lib yuborildi." if split_mode else ""

    # Post-generation messages adhering to requirements
    if is_finish:
        finish_text = (
            f"🎉 *Hujjat tayyor:* {esc(result_title)}\n"
            f"Jami: *{len(session.products)}* ta mahsulot, *{page_count}* bet.{esc(split_notice)}\n\n"
            "Agar yangi varaq boshlamoqchi bo‘lsangiz /new bosing.\n"
            "Agar shu ro‘yxatga yana mahsulot qo‘shmoqchi bo‘lsangiz /add bosing.\n"
            "Oldingi fayllar uchun /files bosing."
        )
        await status_msg.edit_text(
            finish_text,
            reply_markup=MAIN_INLINE_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
    else:
        preview_text = (
            f"👀 *Preview tayyor:* {esc(result_title)}\n"
            f"Hozirgi holat: *{len(session.products)}* ta mahsulot.{esc(split_notice)}\n\n"
            "Ro‘yxat saqlanib qoldi. Yana mahsulot qo‘shish uchun /add bosing yoki yakuniy fayllarni olish uchun /finish bosing.\n"
            "Oldingi fayllar uchun /files bosing."
        )
        await status_msg.edit_text(
            preview_text,
            reply_markup=MAIN_INLINE_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )


async def cmd_preview(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Generates preview documents without clearing products."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    await _generate_and_send_documents(update, context, session, is_finish=False)


async def cmd_finish(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Generates final documents without clearing products."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    await _generate_and_send_documents(update, context, session, is_finish=True)


async def cmd_files(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Shows previous generated documents and lets the user download them again."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    records = session.list_exports(limit=10)

    if not records:
        text = (
            "🗂 Oldingi hujjatlar hali yo‘q.\n\n"
            "Preview yoki final fayl yaratganingizdan keyin ular shu yerda chiqadi. "
            "Fayllar serverda 30 kun saqlanadi."
        )
        if update.callback_query:
            await update.callback_query.answer()
            await context.bot.send_message(
                chat_id=update.effective_chat.id,
                text=text,
                reply_markup=HIDE_REPLY_KEYBOARD,
            )
        else:
            await update.message.reply_text(text, reply_markup=HIDE_REPLY_KEYBOARD)
        return

    lines = ["🗂 *Oldingi hujjatlar:*\n"]
    for index, record in enumerate(records, start=1):
        lines.append(
            f"{index}. *{esc(record.title)}* — {export_kind_label(record.kind)}\n"
            f"   {esc(format_export_details(record))} | {esc(format_export_created_at(record))}"
        )
    lines.append("\nKerakli hujjatni qayta yuklab olish uchun pastdagi tugmani bosing.")

    if update.callback_query:
        await update.callback_query.answer()
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text="\n".join(lines),
            reply_markup=build_exports_keyboard(records),
            parse_mode=ParseMode.MARKDOWN,
        )
    else:
        await update.message.reply_text(
            "\n".join(lines),
            reply_markup=build_exports_keyboard(records),
            parse_mode=ParseMode.MARKDOWN,
        )


async def handle_inline_callbacks(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles inline buttons for preview, finish, list, files, and downloads."""
    query = update.callback_query
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    if query.data.startswith("file:"):
        export_id = query.data.split(":", 1)[1]
        record = session.get_export(export_id)
        if not record:
            await query.answer("Fayl topilmadi", show_alert=True)
            return
        await query.answer("Yuborilyapti...")
        await send_export_record_documents(
            context=context,
            chat_id=update.effective_chat.id,
            session=session,
            record=record,
        )
    elif query.data == "btn_preview":
        await query.answer()
        await _generate_and_send_documents(update, context, session, is_finish=False)
    elif query.data == "btn_finish":
        await query.answer()
        await _generate_and_send_documents(update, context, session, is_finish=True)
    elif query.data == "btn_list":
        await cmd_list(update, context)
    elif query.data == "btn_files":
        await cmd_files(update, context)
    elif query.data == "btn_menu":
        await query.answer()
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text="☰ Pastki menyu ochildi. Tugmani bosganingizdan keyin u yana yopiladi.",
            reply_markup=MAIN_REPLY_KEYBOARD,
        )
    elif query.data.startswith("edit:"):
        index = parse_index_from_callback(query.data)
        if not index:
            await query.answer("Noto‘g‘ri mahsulot raqami", show_alert=True)
            return
        await show_product_edit_options(update=update, context=context, index=index)
    elif query.data.startswith("remove:"):
        await remove_product_callback(update, context)


async def cleanup_old_files_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    """Background retention cleanup for generated files and stale drafts."""
    await asyncio.to_thread(session_mgr.cleanup_all_expired_files, GENERATED_RETENTION_DAYS)


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Global error handler."""
    logger.error("Exception while handling update: %s", context.error, exc_info=context.error)
    if isinstance(update, Update) and update.effective_message:
        try:
            await update.effective_message.reply_text(
                "⚠️ Kutilmagan xatolik yuz berdi. Iltimos, qaytadan urinib ko‘ring yoki /cancel bosing."
            )
        except Exception:
            pass


async def post_init(application) -> None:
    """Sets bot commands in Telegram UI menu."""
    commands = [
        BotCommand("start", "Foydalanish qo‘llanmasi"),
        BotCommand("add", "Yangi mahsulot qo‘shish"),
        BotCommand("preview", "Hozirgi holatni ko‘rish (DOCX/PDF)"),
        BotCommand("finish", "Final hujjatlarni yuklab olish"),
        BotCommand("files", "Oldingi hujjatlarni qayta yuklash"),
        BotCommand("list", "Kiritilgan mahsulotlar ro‘yxati"),
        BotCommand("edit", "Mahsulotni tuzatish"),
        BotCommand("swap", "Ikki mahsulot joyini almashtirish"),
        BotCommand("move", "Mahsulotni aniq o‘ringa qo‘yish"),
        BotCommand("remove", "N-raqamli mahsulotni o‘chirish"),
        BotCommand("new", "Yangi hujjat boshlash"),
        BotCommand("menu", "Pastki menyuni qo‘lda ochish"),
        BotCommand("cancel", "Kiritishni bekor qilish"),
    ]
    try:
        await application.bot.set_my_commands(commands)
    except Exception as exc:
        logger.warning("Could not set bot commands: %s", exc)

    try:
        await asyncio.to_thread(session_mgr.cleanup_all_expired_files, GENERATED_RETENTION_DAYS)
    except Exception as exc:
        logger.warning("Startup cleanup failed: %s", exc)

    if application.job_queue:
        application.job_queue.run_repeating(
            cleanup_old_files_job,
            interval=CLEANUP_INTERVAL_HOURS * 3600,
            first=CLEANUP_INTERVAL_HOURS * 3600,
            name="cleanup_old_generated_files",
        )
    else:
        logger.warning("JobQueue is not available; scheduled cleanup is disabled.")


def create_bot_application() -> object:
    """Builds and configures the Telegram Application."""
    if not TELEGRAM_BOT_TOKEN:
        print("ERROR: TELEGRAM_BOT_TOKEN is not set in environment or .env file.", file=sys.stderr)
        print("Please copy .env.example to .env and set your bot token.", file=sys.stderr)
        sys.exit(1)

    app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).post_init(post_init).build()

    # Add product conversation handler
    conv_handler = ConversationHandler(
        entry_points=[
            CommandHandler("new", new_document_start),
            CommandHandler("add", add_start),
            CommandHandler("swap", cmd_swap),
            CommandHandler("move", cmd_move),
            MessageHandler(filters.Regex(r"^(🆕 Yangi hujjat|🆕 /new)$"), new_document_start),
            MessageHandler(filters.Regex(r"^(➕ Keyingi mahsulot|➕ /add)$"), add_start),
            CallbackQueryHandler(new_document_start, pattern="^btn_new$"),
            CallbackQueryHandler(add_start, pattern="^btn_add$"),
            CallbackQueryHandler(edit_photo_start, pattern=r"^edit_photo:\d+$"),
            CallbackQueryHandler(edit_name_start, pattern=r"^edit_name:\d+$"),
            CallbackQueryHandler(edit_barcode_start, pattern=r"^edit_barcode:\d+$"),
            CallbackQueryHandler(swap_product_start, pattern=r"^swap:\d+$"),
            CallbackQueryHandler(move_product_start, pattern=r"^move:\d+$"),
        ],
        states={
            STATE_WAITING_DOCUMENT_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, document_name_received),
            ],
            STATE_WAITING_PHOTO: [
                MessageHandler(filters.PHOTO | (filters.Document.IMAGE), add_photo_received),
                CommandHandler("done", cmd_done_in_conv),
            ],
            STATE_WAITING_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_name_received),
                MessageHandler(filters.PHOTO | (filters.Document.IMAGE), add_photo_received),
                CommandHandler("done", cmd_done_in_conv),
            ],
            STATE_WAITING_BARCODE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, add_barcode_received),
                CommandHandler("done", cmd_done_in_conv),
            ],
            STATE_WAITING_EDIT_PHOTO: [
                MessageHandler(filters.PHOTO | (filters.Document.IMAGE), edit_photo_received),
            ],
            STATE_WAITING_EDIT_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, edit_name_received),
            ],
            STATE_WAITING_EDIT_BARCODE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, edit_barcode_received),
            ],
            STATE_WAITING_SWAP_TARGET: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, swap_target_received),
            ],
            STATE_WAITING_MOVE_TARGET: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, move_target_received),
            ],
        },
        fallbacks=[CommandHandler("cancel", cmd_cancel_in_conv)],
        per_chat=True,
        per_user=True,
        per_message=False,
        allow_reentry=True,
    )

    # Register conversation handler FIRST so it intercepts in-conversation commands
    app.add_handler(conv_handler)

    # Register inline button callbacks
    app.add_handler(CallbackQueryHandler(handle_inline_callbacks, pattern=r"^(btn_(preview|finish|list|files|menu)|file:[a-f0-9]{12}|edit:\d+|remove:\d+)$"))

    # Register button text handlers
    app.add_handler(MessageHandler(filters.Regex(r"^(👀 Preview olish|👀 /preview)$"), cmd_preview))
    app.add_handler(MessageHandler(filters.Regex(r"^(✅ Final yuklab olish|✅ /finish)$"), cmd_finish))
    app.add_handler(MessageHandler(filters.Regex(r"^(📋 Ro‘yxat|📋 /list)$"), cmd_list))
    app.add_handler(MessageHandler(filters.Regex(r"^(🗂 Oldingi fayllar|🗂 /files)$"), cmd_files))

    # Register global command handlers
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("list", cmd_list))
    app.add_handler(CommandHandler("edit", cmd_edit))
    app.add_handler(CommandHandler("swap", cmd_swap))
    app.add_handler(CommandHandler("move", cmd_move))
    app.add_handler(CommandHandler("remove", cmd_remove))
    app.add_handler(CommandHandler("preview", cmd_preview))
    app.add_handler(CommandHandler("finish", cmd_finish))
    app.add_handler(CommandHandler("files", cmd_files))
    app.add_handler(CommandHandler("menu", cmd_menu))
    app.add_handler(CommandHandler("done", cmd_done_global))
    app.add_handler(CommandHandler("cancel", cmd_cancel_global))

    app.add_error_handler(error_handler)
    return app


def main() -> None:
    """Bot runner."""
    application = create_bot_application()
    run_mode = BOT_RUN_MODE.lower()

    if run_mode == "webhook":
        webhook_path = normalize_webhook_path(WEBHOOK_PATH)
        if not WEBHOOK_BASE_URL.startswith("https://"):
            print("ERROR: WEBHOOK_BASE_URL must start with https:// in webhook mode.", file=sys.stderr)
            sys.exit(1)

        public_url = build_webhook_public_url(WEBHOOK_BASE_URL, webhook_path)
        logger.info(
            "Starting Telegram Product Label Bot webhook on %s:%s with configured webhook path",
            WEBHOOK_LISTEN_HOST,
            WEBHOOK_LISTEN_PORT,
        )
        application.run_webhook(
            listen=WEBHOOK_LISTEN_HOST,
            port=WEBHOOK_LISTEN_PORT,
            url_path=webhook_path.lstrip("/"),
            webhook_url=public_url,
            secret_token=WEBHOOK_SECRET_TOKEN or None,
        )
        return

    if run_mode != "polling":
        print("ERROR: BOT_RUN_MODE must be either 'polling' or 'webhook'.", file=sys.stderr)
        sys.exit(1)

    logger.info("Starting Telegram Product Label Bot polling...")
    application.run_polling()


if __name__ == "__main__":
    main()
