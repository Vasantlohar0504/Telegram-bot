import os
import asyncio
from functools import partial
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
)

# ------------------ IMPORT CONVERTERS ------------------
from config import BOT_TOKEN
from converters.pdf_to_image import pdf_to_images
from converters.pdf_to_word import pdf_to_word
from converters.pdf_to_excel import pdf_to_excel
from converters.image_to_pdf import images_to_pdf, compress_image
from converters.word_to_pdf import word_to_pdf
from converters.compress_pdf import compress_pdf
from converters.office_to_pdf import office_to_pdf
from converters.image_to_word import image_to_word

# ------------------ DIRECTORIES ------------------
DOWNLOAD_DIR = "downloads"
IMAGE_DIR = "outputs/images"
WORD_DIR = "outputs/words"
EXCEL_DIR = "outputs/excel"
IMAGES_TO_PDF_DIR = "outputs/images_to_pdf"
WORDS_TO_PDF_DIR = "outputs/words_to_pdf"
OFFICE_PDF_DIR = "outputs/office_to_pdf"
PDF_COMPRESS_DIR = "outputs/compressed_pdf"
COMPRESSED_IMAGE_DIR = "outputs/compressed_images"
PPT_DIR = "outputs/ppt"

for folder in [
    DOWNLOAD_DIR, IMAGE_DIR, WORD_DIR, EXCEL_DIR,
    IMAGES_TO_PDF_DIR, WORDS_TO_PDF_DIR, OFFICE_PDF_DIR,
    PDF_COMPRESS_DIR, COMPRESSED_IMAGE_DIR, PPT_DIR
]:
    os.makedirs(folder, exist_ok=True)

# ------------------ START COMMAND ------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📄 Send a file (PDF, DOCX, PPTX, Image)\n\n"
        "I can convert:\n"
        "📄 PDF → Images / Word / Excel\n"
        "🖼 Image → PDF / Compress / Word\n"
        "📄 Word → PDF\n"
        "📄 PPTX → PDF\n"
        "📄 Compress PDF"
    )

# ------------------ FILE HANDLER ------------------
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if not document:
        await update.message.reply_text("❌ Please send a valid file.")
        return

    file = await document.get_file()
    file_name = document.file_name
    ext = file_name.split(".")[-1].lower()

    if ext == "pdf":
        path = os.path.join(DOWNLOAD_DIR, file_name)
        await file.download_to_drive(path)
        context.user_data["pdf_path"] = path

        keyboard = [
            [InlineKeyboardButton("🖼 PDF → Images", callback_data="to_image")],
            [InlineKeyboardButton("📄 PDF → Word", callback_data="to_word")],
            [InlineKeyboardButton("📊 PDF → Excel", callback_data="to_excel")],
            [InlineKeyboardButton("📄 Compress PDF", callback_data="compress_pdf")]
        ]
        await update.message.reply_text(
            "✅ PDF received. Choose conversion:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif ext == "docx":
        path = os.path.join(WORD_DIR, file_name)
        await file.download_to_drive(path)
        context.user_data["word_path"] = path

        keyboard = [[InlineKeyboardButton("📄 Word → PDF", callback_data="word_to_pdf")]]
        await update.message.reply_text(
            "✅ Word file received. Choose conversion:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif ext == "pptx":
        path = os.path.join(PPT_DIR, file_name)
        await file.download_to_drive(path)
        context.user_data["ppt_path"] = path

        keyboard = [[InlineKeyboardButton("📄 PPTX → PDF", callback_data="office_to_pdf")]]
        await update.message.reply_text(
            "✅ PPTX file received. Choose conversion:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    elif ext in ["jpg", "jpeg", "png"]:
        path = os.path.join(IMAGE_DIR, file_name)
        await file.download_to_drive(path)
        context.user_data.setdefault("images_list", []).append(path)

        keyboard = [
            [InlineKeyboardButton("🖼 Images → PDF", callback_data="images_to_pdf")],
            [InlineKeyboardButton("📄 Image → Word", callback_data="image_to_word")],
            [InlineKeyboardButton("🖼 Compress Image", callback_data="compress_image")]
        ]
        await update.message.reply_text(
            "✅ Image received. Choose conversion:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

    else:
        await update.message.reply_text(
            "❌ Unsupported file type. Allowed: PDF, DOCX, PPTX, JPG, JPEG, PNG."
        )

# ------------------ BUTTON HANDLER ------------------
async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    loop = asyncio.get_event_loop()

    pdf_path = context.user_data.get("pdf_path")
    if query.data in ["to_image", "to_word", "to_excel", "compress_pdf"] and (not pdf_path or not os.path.exists(pdf_path)):
        await query.message.reply_text("❌ Please upload a PDF first.")
        return

    if query.data == "to_image":
        await query.message.reply_text("⏳ Converting PDF to Images...")
        images = pdf_to_images(pdf_path, IMAGE_DIR)
        for img in images:
            await query.message.reply_photo(photo=open(img, "rb"))

    elif query.data == "to_word":
        await query.message.reply_text("⏳ Converting PDF to Word...")
        out = os.path.join(WORD_DIR, f"{os.path.splitext(os.path.basename(pdf_path))[0]}_converted.docx")
        pdf_to_word(pdf_path, out)
        await query.message.reply_document(open(out, "rb"))

    elif query.data == "to_excel":
        await query.message.reply_text("⏳ Converting PDF to Excel...")
        out = os.path.join(EXCEL_DIR, "converted.xlsx")
        try:
            pdf_to_excel(pdf_path, out)
            await query.message.reply_document(open(out, "rb"))
        except Exception:
            await query.message.reply_text("❌ PDF not suitable for Excel conversion.")

    elif query.data == "compress_pdf":
        await query.message.reply_text("⏳ Compressing PDF...")
        out = os.path.join(PDF_COMPRESS_DIR, "compressed.pdf")
        await loop.run_in_executor(None, partial(compress_pdf, pdf_path, out))
        await query.message.reply_document(open(out, "rb"))

    elif query.data == "images_to_pdf":
        images_list = context.user_data.get("images_list", [])
        if not images_list:
            await query.message.reply_text("❌ No images uploaded to convert.")
            return
        await query.message.reply_text("⏳ Converting Images to PDF...")
        out = os.path.join(IMAGES_TO_PDF_DIR, "images_converted.pdf")
        images_to_pdf(images_list, out)
        await query.message.reply_document(open(out, "rb"))

    elif query.data == "compress_image":
        images_list = context.user_data.get("images_list", [])
        if not images_list:
            await query.message.reply_text("❌ No images uploaded to compress.")
            return
        await query.message.reply_text("⏳ Compressing Images...")
        for img in images_list:
            out_path = os.path.join(COMPRESSED_IMAGE_DIR, os.path.basename(img))
            compress_image(img, out_path, quality=50)
            await query.message.reply_photo(photo=open(out_path, "rb"))

    elif query.data == "image_to_word":
        images_list = context.user_data.get("images_list", [])
        if not images_list:
            await query.message.reply_text("❌ No image uploaded.")
            return
        await query.message.reply_text("⏳ Converting Image to Word...")
        image_path = images_list[-1]
        out = os.path.join(
            WORD_DIR,
            f"{os.path.splitext(os.path.basename(image_path))[0]}_image.docx"
        )
        image_to_word(image_path, out)
        await query.message.reply_document(open(out, "rb"))

    elif query.data == "word_to_pdf":
        word_path = context.user_data.get("word_path")
        if not word_path or not os.path.exists(word_path):
            await query.message.reply_text("❌ Upload a Word file first.")
            return
        await query.message.reply_text("⏳ Converting Word to PDF...")
        out = os.path.join(WORDS_TO_PDF_DIR, "word_converted.pdf")
        word_to_pdf(word_path, out)
        await query.message.reply_document(open(out, "rb"))

    elif query.data == "office_to_pdf":
        file_path = context.user_data.get("ppt_path") or context.user_data.get("word_path")
        if not file_path or not os.path.exists(file_path):
            await query.message.reply_text("❌ Upload DOCX or PPTX first.")
            return

        await query.message.reply_text("⏳ Converting Office file to PDF...")

        pdf_path = office_to_pdf(file_path, OFFICE_PDF_DIR)

        if not os.path.exists(pdf_path):
            await query.message.reply_text("❌ Failed to convert file to PDF.")
            return

        await query.message.reply_document(open(pdf_path, "rb"))

# ------------------ MAIN ------------------
def main():
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN not set in config.py")

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
    app.add_handler(CallbackQueryHandler(button_handler))

    print("🤖 Bot is running...")
    app.run_polling()


if __name__ == "__main__":
    main()
