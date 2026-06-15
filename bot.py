from __future__ import annotations

import logging
import os
from pathlib import Path

from dotenv import load_dotenv
from telegram import Update
from telegram.constants import ChatAction
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from office_tools import is_supported_file, run_office_task, supported_file_types


load_dotenv()

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

DATA_DIR = Path(os.getenv("BOT_DATA_DIR", "data"))


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Salam! Mən ofis işləri üçün köməkçi botam.\n\n"
        "Excel, Word, PowerPoint, PDF, CSV və ya mətn faylı göndərin, sonra tapşırıq yazın. Məsələn:\n"
        "- xülasə ver\n"
        "- sütunları göstər\n"
        "- status = ödənilib filtr et\n"
        "- məbləğ sütununa görə sırala\n"
        "- mətni çıxart\n"
        "- müqavilə sözünü axtar\n"
        "- bu CV-ni analiz et\n"
        "- bu müqavilədə riskləri tap\n"
        "- csv et\n\n"
        "Sərbəst AI tapşırıqları üçün kompüterdə Ollama işləməlidir."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.message
    if message is None or message.document is None:
        return

    chat_dir = _chat_dir(update.effective_chat.id)
    chat_dir.mkdir(parents=True, exist_ok=True)

    file_name = message.document.file_name or "uploaded_file"
    file_path = chat_dir / file_name

    if not is_supported_file(file_path):
        await message.reply_text(f"Bu fayl tipi hələ dəstəklənmir. Bunlardan birini göndərin: `{supported_file_types()}`.")
        return

    await message.chat.send_action(ChatAction.UPLOAD_DOCUMENT)
    telegram_file = await message.document.get_file()
    await telegram_file.download_to_drive(file_path)
    context.user_data["latest_file"] = str(file_path)

    if message.caption:
        await _run_task(update, context, message.caption)
    else:
        await message.reply_text(
            "Faylı aldım. İndi nə etməyimi yazın. Məsələn: `xülasə ver`, `mətni çıxart` və ya `sözünü axtar`."
        )


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.message
    if message is None or message.text is None:
        return

    await _run_task(update, context, message.text)


async def _run_task(update: Update, context: ContextTypes.DEFAULT_TYPE, instruction: str) -> None:
    message = update.message
    if message is None:
        return

    latest_file = context.user_data.get("latest_file")
    if not latest_file:
        await message.reply_text("Əvvəlcə ofis faylı göndərin, sonra tapşırığı yazın.")
        return

    file_path = Path(latest_file)
    if not file_path.exists():
        await message.reply_text("Son göndərilən faylı tapa bilmədim. Zəhmət olmasa faylı yenidən göndərin.")
        return

    await message.chat.send_action(ChatAction.TYPING)
    output_dir = _chat_dir(update.effective_chat.id) / "outputs"

    try:
        result = run_office_task(file_path, instruction, output_dir)
    except Exception:
        logger.exception("Task failed")
        await message.reply_text("Tapşırığı yerinə yetirərkən xəta oldu. Faylın strukturunu yoxlayıb yenidən cəhd edin.")
        return

    await message.reply_text(result.message, parse_mode="Markdown")
    if result.output_path:
        await message.chat.send_action(ChatAction.UPLOAD_DOCUMENT)
        await message.reply_document(document=result.output_path.open("rb"), filename=result.output_path.name)


def _chat_dir(chat_id: int) -> Path:
    return DATA_DIR / str(chat_id)


def main() -> None:
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not token:
        raise RuntimeError("TELEGRAM_BOT_TOKEN `.env` faylında yazılmalıdır.")

    app = Application.builder().token(token).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.run_polling()


if __name__ == "__main__":
    main()
