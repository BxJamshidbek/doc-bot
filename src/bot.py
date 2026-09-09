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
)
from src.session import SessionManager, UserSession
from src.docx_gen import create_label_sheet
from src.pdf_converter import convert_docx_to_pdf
from src.barcode_gen import generate_barcode_image

# Configure logging
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# Conversation states
STATE_WAITING_PHOTO, STATE_WAITING_NAME, STATE_WAITING_BARCODE = range(3)

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
    ],
    resize_keyboard=True,
)

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
    ]
)


def esc(text: object) -> str:
    """Safely escapes user text for Telegram Markdown V1."""
    return escape_markdown(str(text), version=1)


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
        "1. /add bosing yoki *[➕ Keyingi mahsulot]* tugmasini bosing\n"
        "2. *Rasm* ➔ *Nomi* ➔ *Barcode* yuboring\n"
        "3. Mahsulot avtomatik saqlanadi!\n"
        "4. Istalgan vaqtda /preview yoki /finish orqali hujjatni oling.\n\n"
        "⚙️ *Buyruqlar:*\n"
        "/add — Yangi mahsulot qo‘shish\n"
        "/preview — Hozirgi ro‘yxatni ko‘rish uchun DOCX/PDF olish\n"
        "/finish — Final DOCX & PDF fayllarni yuklab olish\n"
        "/list — Ro‘yxatni ko‘rish\n"
        "/remove `<N>` — Mahsulotni o‘chirish (masalan: `/remove 1`)\n"
        "/new — Yangi varaq boshlash (ro‘yxatni tozalash)\n"
        "/cancel — Joriy kiritishni bekor qilish"
    )
    await update.message.reply_text(
        help_text, reply_markup=MAIN_REPLY_KEYBOARD, parse_mode=ParseMode.MARKDOWN
    )


async def cmd_new(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles /new command: clears active products and resets session while preserving past exports."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    session.clear()
    await update.message.reply_text(
        "🆕 *Yangi varaq boshlandi!*\nOldingi mahsulotlar ro‘yxati tozalandi.\n\n"
        "Birinchi mahsulotni qo‘shish uchun /add bosing yoki quyidagi tugmadan foydalaning:",
        reply_markup=MAIN_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )


async def cmd_list(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles /list command: shows collected products in order with safe markdown."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    target_msg = update.message or (update.callback_query.message if update.callback_query else None)

    if not session.products:
        text = "📭 Hozircha mahsulotlar ro‘yxati bo‘sh.\nMahsulot qo‘shish uchun /add bosing."
        if update.callback_query:
            await update.callback_query.answer()
            await context.bot.send_message(
                chat_id=update.effective_chat.id,
                text=text,
                reply_markup=MAIN_REPLY_KEYBOARD,
            )
        else:
            await update.message.reply_text(
                text, reply_markup=MAIN_REPLY_KEYBOARD, parse_mode=ParseMode.MARKDOWN
            )
        return

    lines = [f"📋 *Kiritilgan mahsulotlar ({len(session.products)} ta):*\n"]
    for i, p in enumerate(session.products, start=1):
        lines.append(f"{i}. *{esc(p.name)}* — `{esc(p.barcode)}`")

    pages = (len(session.products) + 11) // 12
    lines.append(f"\n📄 *Jami betlar:* {pages} ta (A4, 12 ta/bet)")
    lines.append("\n👉 /add — Qo‘shish | /preview — Tekshirish | /finish — Yakunlash")

    msg_text = "\n".join(lines)
    if update.callback_query:
        await update.callback_query.answer()
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=msg_text,
            reply_markup=MAIN_INLINE_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
    else:
        await update.message.reply_text(
            msg_text,
            reply_markup=MAIN_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
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
            parse_mode=ParseMode.MARKDOWN,
        )
        return

    await update.message.reply_text(
        f"🗑️ #{idx} o‘chirildi: *{esc(removed.name)}* (`{esc(removed.barcode)}`).\n"
        f"Qolgan mahsulotlar: {len(session.products)} ta.",
        reply_markup=MAIN_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )


# --- /add Conversation Handlers ---

async def add_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Starts the /add flow: requests product photo."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    session.reset_draft()

    text = (
        "📸 *1/3-qadam: Mahsulot fotosi*\n\n"
        "Iltimos, mahsulot rasmini yuboring.\n"
        "Mahsulotni oq yoki yengil fonda, yaxshi yorug‘likda, kadr markaziga olib rasmga oling.\n"
        "_Oq fon avtomatik qirqiladi. Sifat yaxshi bo‘lishi uchun rasmni file/document sifatida yuborishingiz ham mumkin._\n\n"
        "Bekor qilish uchun /cancel bosing."
    )

    if update.callback_query:
        await update.callback_query.answer()
        await context.bot.send_message(
            chat_id=update.effective_chat.id, text=text, parse_mode=ParseMode.MARKDOWN
        )
    else:
        await update.message.reply_text(text, parse_mode=ParseMode.MARKDOWN)

    return STATE_WAITING_PHOTO


async def add_photo_received(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Processes uploaded photo."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

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
        await update.message.reply_text(
            "⚠️ Iltimos, mahsulot rasmini yuboring, yoki /cancel bosing.",
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_PHOTO

    photo_path = session_mgr.save_draft_photo(user_id, bytes(file_bytes), ext=ext)
    session.current_draft["photo_path"] = photo_path

    await update.message.reply_text(
        "✅ Rasm qabul qilindi!\n\n"
        "📝 *2/3-qadam: Mahsulot nomi*\n"
        "Iltimos, mahsulot nomini yuboring (masalan: `Perexod 76-50mm`):",
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
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_NAME

    session.current_draft["name"] = name

    await update.message.reply_text(
        f"✅ Nom saqlandi: *{esc(name)}*\n\n"
        "🔢 *3/3-qadam: Barcode / Kod qiymati*\n"
        "Iltimos, shtrix-kod yoki mahsulot kodini yuboring (masalan: `1000049` yoki 13 xonali EAN):",
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
            parse_mode=ParseMode.MARKDOWN,
        )
        return STATE_WAITING_BARCODE

    # Validate that photo and name exist in draft
    photo_path = session.current_draft.get("photo_path")
    prod_name = session.current_draft.get("name")
    if not photo_path or not prod_name:
        await update.message.reply_text(
            "⚠️ Ma’lumotlar to‘liq emas. Iltimos, /add orqali qaytadan boshlang.",
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
        reply_markup=MAIN_REPLY_KEYBOARD,
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
        reply_markup=MAIN_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )


async def cmd_cancel_in_conv(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Cancels current /add flow inside conversation."""
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)
    session.reset_draft()
    await update.message.reply_text(
        "❌ Mahsulot kiritish bekor qilindi.\n"
        "Yangi mahsulot kiritish uchun /add bosing yoki mavjudlarini olish uchun /finish bosing.",
        reply_markup=MAIN_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )
    return ConversationHandler.END


async def cmd_cancel_global(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles /cancel outside of conversation."""
    await update.message.reply_text(
        "ℹ️ Hozirda bekor qilinadigan jarayon yo‘q.\n"
        "Mahsulot qo‘shish uchun /add yoki ro‘yxatni ko‘rish uchun /list bosing.",
        reply_markup=MAIN_REPLY_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN,
    )


# --- Document Generation & Delivery ---

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
            reply_markup=MAIN_REPLY_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
        return

    status_msg = await context.bot.send_message(
        chat_id=chat_id,
        text="⏳ A4 formatdagi 3×4 yorliqlar varag‘i tayyorlanmoqda...",
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

    # Offload PDF conversion to worker thread
    pdf_path, pdf_note = await asyncio.to_thread(convert_docx_to_pdf, docx_path)

    # Send DOCX file
    try:
        with open(docx_path, "rb") as f_docx:
            await context.bot.send_document(
                chat_id=chat_id,
                document=f_docx,
                filename=docx_path.name,
                caption=(
                    f"📄 *{'Final' if is_finish else 'Preview'} mahsulot yorliqlari (DOCX)*\n"
                    f"• {len(session.products)} ta mahsulot | {(len(session.products) + 11) // 12} bet\n"
                    f"• A4 format, 12 ta yorliq (3 ustun × 4 qator)"
                ),
                parse_mode=ParseMode.MARKDOWN,
            )
    except Exception as exc:
        logger.exception("Error sending DOCX file: %s", exc)
        await status_msg.edit_text(f"❌ DOCX yuborishda xatolik: {esc(exc)}")
        return

    # Send PDF file if generated
    if pdf_path and pdf_path.is_file():
        try:
            with open(pdf_path, "rb") as f_pdf:
                await context.bot.send_document(
                    chat_id=chat_id,
                    document=f_pdf,
                    filename=pdf_path.name,
                    caption="📑 *Chop etish uchun tayyor (PDF)*\n• A4 formatda to‘g‘ridan-to‘g‘ri chop etish uchun.",
                    parse_mode=ParseMode.MARKDOWN,
                )
        except Exception as exc:
            logger.warning("Error sending PDF file: %s", exc)

    # If PDF wasn't generated and a note is available
    if not pdf_path and pdf_note:
        await context.bot.send_message(
            chat_id=chat_id,
            text=f"ℹ️ *PDF ma’lumoti:*\n{esc(pdf_note)}\n\n*DOCX faylingiz* tayyor va uni bemalol chop etishingiz mumkin.",
            parse_mode=ParseMode.MARKDOWN,
        )

    # Post-generation messages adhering to requirements
    if is_finish:
        finish_text = (
            f"🎉 *Hujjat tayyor!* (Jami: *{len(session.products)}* ta mahsulot, *{(len(session.products) + 11) // 12}* bet).\n\n"
            "Agar yangi varaq boshlamoqchi bo‘lsangiz /new bosing.\n"
            "Agar shu ro‘yxatga yana mahsulot qo‘shmoqchi bo‘lsangiz /add bosing."
        )
        await status_msg.edit_text(
            finish_text,
            reply_markup=MAIN_INLINE_KEYBOARD,
            parse_mode=ParseMode.MARKDOWN,
        )
    else:
        preview_text = (
            f"👀 *Preview tayyor!* Hozirgi holat: *{len(session.products)}* ta mahsulot.\n\n"
            "Ro‘yxat saqlanib qoldi. Yana mahsulot qo‘shish uchun /add bosing yoki yakuniy fayllarni olish uchun /finish bosing."
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


async def handle_inline_callbacks(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles inline buttons for preview, finish, list."""
    query = update.callback_query
    user_id = update.effective_user.id
    session = session_mgr.get_session(user_id)

    if query.data == "btn_preview":
        await query.answer()
        await _generate_and_send_documents(update, context, session, is_finish=False)
    elif query.data == "btn_finish":
        await query.answer()
        await _generate_and_send_documents(update, context, session, is_finish=True)
    elif query.data == "btn_list":
        await cmd_list(update, context)


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
        BotCommand("list", "Kiritilgan mahsulotlar ro‘yxati"),
        BotCommand("remove", "N-raqamli mahsulotni o‘chirish"),
        BotCommand("new", "Yangi varaq boshlash"),
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
            CommandHandler("add", add_start),
            MessageHandler(filters.Regex(r"^(➕ Keyingi mahsulot|➕ /add)$"), add_start),
            CallbackQueryHandler(add_start, pattern="^btn_add$"),
        ],
        states={
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
        },
        fallbacks=[CommandHandler("cancel", cmd_cancel_in_conv)],
        per_chat=True,
        per_user=True,
        per_message=False,
    )

    # Register conversation handler FIRST so it intercepts in-conversation commands
    app.add_handler(conv_handler)

    # Register inline button callbacks
    app.add_handler(CallbackQueryHandler(handle_inline_callbacks, pattern="^btn_(preview|finish|list)$"))

    # Register button text handlers
    app.add_handler(MessageHandler(filters.Regex(r"^(👀 Preview olish|👀 /preview)$"), cmd_preview))
    app.add_handler(MessageHandler(filters.Regex(r"^(✅ Final yuklab olish|✅ /finish)$"), cmd_finish))
    app.add_handler(MessageHandler(filters.Regex(r"^(📋 Ro‘yxat|📋 /list)$"), cmd_list))

    # Register global command handlers
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("new", cmd_new))
    app.add_handler(CommandHandler("list", cmd_list))
    app.add_handler(CommandHandler("remove", cmd_remove))
    app.add_handler(CommandHandler("preview", cmd_preview))
    app.add_handler(CommandHandler("finish", cmd_finish))
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
            "Starting Telegram Product Label Bot webhook on %s:%s%s",
            WEBHOOK_LISTEN_HOST,
            WEBHOOK_LISTEN_PORT,
            webhook_path,
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
