"""Telegram-бот с контекстным AI-ассистентом на Gemini API."""

import json
import logging
import os
from collections.abc import Iterable
from datetime import timedelta
from io import BytesIO

from dotenv import load_dotenv
from google import genai
from google.genai import errors, types
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from telegram import Update
from telegram.constants import ChatAction
from telegram.error import TelegramError
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
MODEL = os.getenv("GEMINI_MODEL", "gemini-3.6-flash").strip()
IMAGE_MODEL = os.getenv("GEMINI_IMAGE_MODEL", "gemini-3.1-flash-image").strip()
MAX_HISTORY_MESSAGES = 12
TELEGRAM_MESSAGE_LIMIT = 4096

SYSTEM_PROMPT = """
Ты — современный, умный и дружелюбный AI-ассистент в Telegram.

ПРАВИЛА ОБЩЕНИЯ И СТИЛЯ:
1. Говори просто, живо и естественно, как опытный старший товарищ или крутой ментор. Избегай сухого академического языка, канцеляризмов и старых учебниковых формулировок ("Представляем вам...", "Давай разберем...").
2. Пиши кратко, емко и по делу. Не используй лишнее вступление и воду.
3. Категорически НЕ используй LaTeX и знаки доллара ($ или $$). Telegram их не поддерживает.
4. Для математических формул используй обычные символы Unicode (x², x₁, x₂, -b/a, √D).
5. Разбивай текст на короткие абзацы и списки, чтобы удобно было читать с экрана телефона.
"""



logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None


def _safe_reply(message, text: str) -> None:
    """Reply safely even if Telegram temporarily rejects the request."""
    try:
        if hasattr(message, "reply_text"):
            message.reply_text(text)
    except Exception:
        logger.exception("Failed to send Telegram reply")


def split_message(text: str) -> Iterable[str]:
    """Split a long response into Telegram-safe pieces."""
    while text:
        if len(text) <= TELEGRAM_MESSAGE_LIMIT:
            yield text
            return

        split_at = text.rfind("\n", 0, TELEGRAM_MESSAGE_LIMIT)
        if split_at == -1:
            split_at = text.rfind(" ", 0, TELEGRAM_MESSAGE_LIMIT)
        if split_at == -1:
            split_at = TELEGRAM_MESSAGE_LIMIT

        yield text[:split_at]
        text = text[split_at:].lstrip()


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_message.reply_text(
        "👋 **Welcome / Привет!**\n\n"
        "🇬🇧 I am an AI coding mentor and study assistant. Send me a message, a snippet of code, or a photo to get started!\n\n"
        "🇷🇺 Я AI-ментор и помощник в учебе. Напиши мне вопрос, отправь код или фото задачи!\n\n"
        "📌 **Commands / Команды:**\n"
        "/reset — Clear context / Очистить историю\n"
        "/slides <topic> — Create PPTX / Создать презентацию\n"
        "/image <prompt> — Generate image / AI-картинка\n"
        "/remind <min> <text> — Set reminder / Напоминание\n"
        "/help — Full instructions / Помощь"
    )
    


async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data.pop("history", None)
    await update.effective_message.reply_text("История диалога очищена. Начнём заново!")


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_message.reply_text(
        "Отправьте текст, фотографию или изображение файлом — я отвечу с учётом "
        "нашего диалога. К изображению можно добавить подпись с вопросом.\n\n"
        "Напоминание: /remind <минуты> <текст>\n"
        "Пример: /remind 30 Выпить воды\n\n"
        "AI-изображение: /image <описание>\n"
        "Презентация: /slides <тема>. Можно отправить фото с подписью «/slides "
        "— оно станет основой презентации и попадёт на титульный слайд.\n\n"
        "Используйте /reset, если хотите начать новую тему."
    )


async def send_reminder(context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a scheduled reminder to the chat that created it."""
    job = context.job
    await context.bot.send_message(
        chat_id=job.chat_id,
        text=f"⏰ Напоминание: {job.data['text']}",
    )


async def remind(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Schedule a one-time reminder in minutes."""
    if not context.args or len(context.args) < 2:
        await update.effective_message.reply_text(
            "Формат: /remind <минуты> <текст>\n"
            "Например: /remind 30 Выпить воды"
        )
        return

    try:
        minutes = int(context.args[0])
        if minutes < 1 or minutes > 43_200:
            raise ValueError
    except ValueError:
        await update.effective_message.reply_text(
            "Укажите целое число минут от 1 до 43200 (30 дней)."
        )
        return

    if context.job_queue is None:
        await update.effective_message.reply_text(
            "Напоминания не настроены. Установите зависимости из requirements.txt "
            "и перезапустите бота."
        )
        return

    reminder_text = " ".join(context.args[1:])
    context.job_queue.run_once(
        send_reminder,
        when=timedelta(minutes=minutes),
        chat_id=update.effective_chat.id,
        data={"text": reminder_text},
        name=f"reminder-{update.effective_chat.id}",
    )
    await update.effective_message.reply_text(
        f"Хорошо, напомню через {minutes} мин.: {reminder_text}"
    )


def image_from_response(response: object) -> bytes | None:
    """Extract the first generated image from a Gemini response."""
    for part in getattr(response, "parts", None) or []:
        inline_data = getattr(part, "inline_data", None)
        if inline_data and getattr(inline_data, "data", None):
            return bytes(inline_data.data)
    return None


async def image_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Create an image from a text prompt and send it back to Telegram."""
    if client is None:
        await update.effective_message.reply_text("Gemini ещё не настроен.")
        return
    if not context.args:
        await update.effective_message.reply_text(
            "Напишите, что создать. Например: /image уютный домик в горах зимой"
        )
        return

    prompt = " ".join(context.args)
    try:
        await update.effective_message.reply_text("🎨 Создаю изображение, это может занять до минуты…")
        await update.effective_chat.send_action(ChatAction.UPLOAD_PHOTO)
        response = await client.aio.models.generate_content(
            model=IMAGE_MODEL,
            contents=prompt,
            config=types.GenerateContentConfig(response_modalities=["IMAGE"]),
        )
        image_bytes = image_from_response(response)
        if image_bytes is None:
            await update.effective_message.reply_text(
                "Не удалось получить изображение. Попробуйте изменить описание."
            )
            return
        await update.effective_message.reply_photo(
            photo=image_bytes,
            caption=f"AI-изображение: {prompt[:900]}",
        )
    except errors.APIError as error:
        logger.exception("Gemini image generation failed: %s", error.code)
        if str(error.code) == "429":
            await update.effective_message.reply_text(
                "🎨 Gemini отклонил запрос: для генерации изображений у этого ключа "
                "исчерпана или не выдана квота (ошибка 429). Включите биллинг либо "
                "квоту для gemini-3.1-flash-image и повторите запрос."
            )
            return
        await update.effective_message.reply_text(
            "Не удалось создать изображение. Проверьте доступ Gemini API и попробуйте ещё раз."
        )


def add_textbox(slide, text: str, left, top, width, height, size: int, color) -> None:
    box = slide.shapes.add_textbox(left, top, width, height)
    paragraph = box.text_frame.paragraphs[0]
    paragraph.text = text
    paragraph.font.size = Pt(size)
    paragraph.font.color.rgb = RGBColor(*color)
    return box


def create_presentation(slides: list[dict[str, object]], cover_image: bytes | None) -> BytesIO:
    """Build a visual, card-based PowerPoint file from the AI slide outline."""
    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    blank_layout = presentation.slide_layouts[6]
    palettes = [
        ((250, 247, 255), (109, 40, 217), (237, 233, 254)),
        ((240, 249, 255), (3, 105, 161), (224, 242, 254)),
        ((240, 253, 244), (21, 128, 61), (220, 252, 231)),
        ((255, 247, 237), (194, 65, 12), (255, 237, 213)),
    ]

    for index, slide_data in enumerate(slides):
        slide = presentation.slides.add_slide(blank_layout)
        background_color, accent_color, pale_accent = palettes[index % len(palettes)]
        background = slide.background.fill
        background.solid()
        background.fore_color.rgb = RGBColor(*background_color)

        blob = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(10.85), Inches(-0.8), Inches(3.0), Inches(3.0)
        )
        blob.fill.solid()
        blob.fill.fore_color.rgb = RGBColor(*pale_accent)
        blob.line.fill.background()
        accent_bar = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(0.55), Inches(0.78), Inches(0.15)
        )
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = RGBColor(*accent_color)
        accent_bar.line.fill.background()

        title = str(slide_data.get("title", "Без названия"))[:120]
        add_textbox(
            slide, title, Inches(0.7), Inches(0.85), Inches(10.8), Inches(0.8), 29, (30, 41, 59)
        )

        if index == 0 and cover_image:
            try:
                cover_card = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.35), Inches(1.65), Inches(4.2), Inches(4.75)
                )
                cover_card.fill.solid()
                cover_card.fill.fore_color.rgb = RGBColor(*pale_accent)
                cover_card.line.fill.background()
                slide.shapes.add_picture(
                    BytesIO(cover_image), Inches(8.55), Inches(1.85), width=Inches(3.8)
                )
            except (OSError, ValueError):
                logger.warning("Could not insert cover image into presentation")

        bullets = slide_data.get("bullets", [])
        if not isinstance(bullets, list):
            bullets = []
        visible_bullets = bullets[:4] or ["Основная идея", "Ключевой вывод"]
        content_width = Inches(7.1) if index == 0 and cover_image else Inches(11.7)
        columns = 1 if len(visible_bullets) < 3 or (index == 0 and cover_image) else 2
        card_width = (content_width - Inches(0.25 * (columns - 1))) / columns
        for bullet_index, bullet in enumerate(visible_bullets):
            column = bullet_index % columns
            row = bullet_index // columns
            card_height = Inches(1.38)
            card_left = Inches(0.7) + column * (card_width + Inches(0.25))
            card_top = Inches(1.85) + row * Inches(1.62)
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, card_left, card_top, card_width, card_height
            )
            card.fill.solid()
            card.fill.fore_color.rgb = RGBColor(255, 255, 255)
            card.line.color.rgb = RGBColor(*pale_accent)
            number = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, card_left + Inches(0.22), card_top + Inches(0.25), Inches(0.48), Inches(0.48)
            )
            number.fill.solid()
            number.fill.fore_color.rgb = RGBColor(*accent_color)
            number.line.fill.background()
            number_text = add_textbox(
                slide, str(bullet_index + 1), card_left + Inches(0.22), card_top + Inches(0.33), Inches(0.48), Inches(0.2), 10, (255, 255, 255)
            )
            number_text.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            card_text = add_textbox(
                slide, str(bullet)[:150], card_left + Inches(0.88), card_top + Inches(0.2), card_width - Inches(1.08), Inches(0.95), 16, (51, 65, 85)
            )
            card_text.text_frame.word_wrap = True

        footer = add_textbox(
            slide, f"{index + 1} / {len(slides)}", Inches(11.7), Inches(6.85), Inches(0.8), Inches(0.3), 10, (100, 116, 139)
        )
        footer.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

    file = BytesIO()
    presentation.save(file)
    file.seek(0)
    return file


async def slides_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Generate a PowerPoint presentation, optionally based on a submitted image."""
    if client is None:
        await update.effective_message.reply_text("Gemini ещё не настроен.")
        return
    command_args = context.args
    caption = update.effective_message.caption or ""
    if not command_args and caption.startswith("/slides"):
        command_args = caption.split()[1:]
    if not command_args:
        await update.effective_message.reply_text(
            "Укажите тему: /slides Экология. Можно отправить фото с этой подписью."
        )
        return

    message = update.effective_message
    image = message.photo[-1] if message.photo else message.document
    image_bytes = None
    image_part = None
    if image:
        try:
            file = await image.get_file()
            image_bytes = bytes(await file.download_as_bytearray())
            image_part = types.Part.from_bytes(
                data=image_bytes,
                mime_type=getattr(image, "mime_type", None) or "image/jpeg",
            )
        except (OSError, ValueError, TelegramError):
            await message.reply_text("Не удалось скачать изображение для презентации.")
            return

    topic = " ".join(command_args)
    prompt = (
        "Составь план презентации на русском языке по теме: " + topic + ". "
        "Верни только корректный JSON без Markdown: объект с ключом slides. "
        "slides — массив из 5–7 объектов; у каждого title (короткий заголовок) "
        "и bullets (массив из 3–5 коротких тезисов). "
        "Первый слайд — титульный, последний — выводы."
    )
    contents = [types.Part.from_text(text=prompt)]
    if image_part:
        contents.append(image_part)
        contents.append(
            types.Part.from_text(
                text="Проанализируй приложенное изображение и используй его как контекст презентации."
            )
        )

    try:
        await update.effective_chat.send_action(ChatAction.TYPING)
        response = await client.aio.models.generate_content(
            model=MODEL,
            contents=types.Content(role="user", parts=contents),
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.4,
                max_output_tokens=2048,
            ),
        )
        outline = json.loads((response.text or "").strip())
        slides = outline.get("slides", [])
        if not isinstance(slides, list) or not 2 <= len(slides) <= 10:
            raise ValueError("invalid slide outline")
        document = create_presentation(slides, image_bytes)
    except (errors.APIError, OSError, ValueError, json.JSONDecodeError):
        logger.exception("Presentation generation failed")
        await message.reply_text(
            "Не удалось создать презентацию. Попробуйте сделать тему короче и отправить ещё раз."
        )
        return

    await update.effective_chat.send_action(ChatAction.UPLOAD_DOCUMENT)
    await message.reply_document(
        document=document,
        filename="ai_presentation.pptx",
        caption=f"Готово: презентация по теме «{topic[:120]}»."
    )


async def chat(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if client is None:
        await update.effective_message.reply_text(
            "Бот ещё не настроен: добавьте GEMINI_API_KEY в файл .env и перезапустите его."
        )
        return

    message = update.effective_message
    user_text = (message.text or message.caption or "").strip()
    if message.caption and message.caption.startswith("/slides"):
        await slides_command(update, context)
        return
    photo = message.photo[-1] if message.photo else None
    image = photo or message.document
    if not user_text and image is None:
        return

    history: list[dict[str, str]] = context.user_data.get("history", [])
    contents = [
        types.Content(
            role=item["role"],
            parts=[types.Part.from_text(text=item["content"])],
        )
        for item in history
    ]
    user_parts = []
    if image:
        try:
            file = await image.get_file()
            image_bytes = await file.download_as_bytearray()
        except (OSError, ValueError, TelegramError):
            logger.exception("Telegram photo download failed")
            await message.reply_text(
                "Не удалось скачать фотографию. Пожалуйста, отправьте её ещё раз."
            )
            return

        user_parts.append(
            types.Part.from_bytes(
                data=bytes(image_bytes),
                mime_type=getattr(image, "mime_type", None) or "image/jpeg",
            )
        )
        user_parts.append(
            types.Part.from_text(
                text=user_text or "Опиши и проанализируй это изображение."
            )
        )
    else:
        user_parts.append(types.Part.from_text(text=user_text))

    contents.append(types.Content(role="user", parts=user_parts))

    try:
        await update.effective_chat.send_action(
            ChatAction.UPLOAD_PHOTO if image else ChatAction.TYPING
        )
        response = await client.aio.models.generate_content(
            model=MODEL,
            contents=contents,
            config=types.GenerateContentConfig(
                system_instruction=SYSTEM_PROMPT,
                temperature=0.7,
                max_output_tokens=2048,
            ),
        )
        reply = (response.text or "").strip() or (
            "Не удалось сформировать ответ. Попробуйте переформулировать вопрос."
        )
    except errors.APIError as error:
        logger.exception("Gemini API request failed: %s", error.code)
        if error.code == 429:
            message = "Слишком много запросов к Gemini. Подождите немного и попробуйте ещё раз."
        elif error.code in {401, 403}:
            message = "Не удалось авторизоваться в Gemini. Проверьте GEMINI_API_KEY в файле .env."
        else:
            message = "Gemini временно недоступен. Пожалуйста, попробуйте ещё раз чуть позже."
        await update.effective_message.reply_text(message)
        return
    except (OSError, ValueError, TypeError):
        logger.exception("Gemini request failed")
        await update.effective_message.reply_text(
            "Не удалось получить ответ от Gemini. Проверьте интернет и попробуйте ещё раз."
        )
        return

    updated_history = [
        *history,
        {
            "role": "user",
            "content": user_text or "[Пользователь отправил фотографию для анализа]",
        },
        {"role": "model", "content": reply},
    ]
    context.user_data["history"] = updated_history[-MAX_HISTORY_MESSAGES:]

    for part in split_message(reply):
        try:
            await update.effective_message.reply_text(part)
        except Exception:
            logger.exception("Failed to send message chunk to Telegram")
            break


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    logger.exception("Unhandled exception while processing an update", exc_info=context.error)


def main() -> None:
    if not TELEGRAM_BOT_TOKEN or TELEGRAM_BOT_TOKEN == "your_telegram_bot_token":
        raise RuntimeError("TELEGRAM_BOT_TOKEN is missing from .env")
    if not GEMINI_API_KEY or GEMINI_API_KEY == "your_gemini_api_key":
        raise RuntimeError("GEMINI_API_KEY is missing from .env")

    application = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("reset", reset))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("remind", remind))
    application.add_handler(CommandHandler("image", image_command))
    application.add_handler(CommandHandler("slides", slides_command))
    application.add_handler(
        MessageHandler(
            (filters.TEXT & ~filters.COMMAND) | filters.PHOTO | filters.Document.IMAGE,
            chat,
        )
    )
    application.add_error_handler(error_handler)

    logger.info("Bot is starting")
    application.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()