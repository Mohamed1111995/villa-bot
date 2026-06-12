import os
import shutil
import logging
from telegram import Update
from telegram.ext import Application, MessageHandler, CommandHandler, filters, ContextTypes, ConversationHandler
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
from datetime import datetime

logging.basicConfig(level=logging.INFO)

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
TEMPLATE_PATH = "TEMPLATE_VILLA.xlsx"

VILLA_NUMBER, PHOTO, NOTES = range(3)
user_data_temp = {}

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "مرحباً في بوت فحص الفيلات!\n\n"
        "1 - أدخل رقم الفيلا (مثال: C110)\n"
        "2 - ارسل الصور\n"
        "3 - اكتب /done\n"
        "4 - اكتب الملاحظات"
    )
    return VILLA_NUMBER

async def receive_villa_number(update: Update, context: ContextTypes.DEFAULT_TYPE):
    villa_number = update.message.text.strip().upper()
    user_id = update.effective_user.id
    user_data_temp[user_id] = {
        "villa": villa_number,
        "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "photo_count": 0
    }
    await update.message.reply_text(f"رقم الفيلا: {villa_number}\n\nارسل الصور، ثم اكتب /done")
    return PHOTO

async def receive_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    data = user_data_temp[user_id]
    if data["photo_count"] >= 44:
        await update.message.reply_text("الحد الأقصى 44 صورة. اكتب /done")
        return PHOTO
    os.makedirs("photos", exist_ok=True)
    photo_file = await update.message.photo[-1].get_file()
    raw_path = f"photos/{user_id}_{data['photo_count']+1}.jpg"
    await photo_file.download_to_drive(raw_path)
    img = PILImage.open(raw_path)
    img_resized = img.resize((547, 403), PILImage.Resampling.LANCZOS)
    processed_path = raw_path.replace(".jpg", "_processed.jpg")
    img_resized.save(processed_path, quality=95)
    data["photo_count"] += 1
    data[f"photo_{data['photo_count']}"] = processed_path
    await update.message.reply_text(f"صورة {data['photo_count']} - ارسل أخرى أو /done")
    return PHOTO

async def done_photos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_data_temp[user_id]["photo_count"] == 0:
        await update.message.reply_text("ارسل صورة أولاً!")
        return PHOTO
    await update.message.reply_text(f"تم استقبال {user_data_temp[user_id]['photo_count']} صورة\n\nاكتب الملاحظات:")
    return NOTES

async def receive_notes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    data = user_data_temp[user_id]
    data["notes"] = update.message.text
    output_path = await save_to_template(data)
    await update.message.reply_text("جاري إعداد الملف...")
    with open(output_path, 'rb') as f:
        await update.message.reply_document(
            document=f,
            filename=f"{data['villa']}.xlsx",
            caption=f"تم!\nالفيلا: {data['villa']}\nصور: {data['photo_count']}\n\n/start لفيلا جديدة"
        )
    os.remove(output_path)
    del user_data_temp[user_id]
    return ConversationHandler.END

async def save_to_template(data):
    output_path = f"temp_{data['villa']}.xlsx"
    shutil.copy(TEMPLATE_PATH, output_path)
    wb = openpyxl.load_workbook(output_path)
    ws_table = wb["الجدول"]
    ws_table['A1'] = f"فيلا رقم {data['villa']} - {data['date']}"
    notes_lines = data["notes"].strip().split("\n")
    for i, note in enumerate(notes_lines):
        row = i + 3
        if row > 52:
            break
        ws_table[f'A{row}'] = i + 1
        ws_table[f'B{row}'] = note.strip()
    for i in range(1, data["photo_count"] + 1):
        photo_path = data.get(f"photo_{i}")
        if not photo_path or not os.path.exists(photo_path):
            continue
        pair_start = i if i % 2 != 0 else i - 1
        pair_end = pair_start + 1
        sheet_name = f"ملاحظة {pair_start},{pair_end}"
        if sheet_name not in wb.sheetnames:
            continue
        ws_img = wb[sheet_name]
        img_xl = XLImage(photo_path)
        img_xl.width = 547
        img_xl.height = 403
        anchor_row = 'A4' if i % 2 != 0 else 'A29'
        ws_img.add_image(img_xl, anchor_row)
    wb.save(output_path)
    return output_path

if __name__ == '__main__':
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            VILLA_NUMBER: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_villa_number)],
            PHOTO: [
                MessageHandler(filters.PHOTO, receive_photo),
                CommandHandler("done", done_photos),
            ],
            NOTES: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_notes)],
        },
        fallbacks=[],
    )
    app.add_handler(conv_handler)
    print("البوت يعمل!")
    app.run_polling(allowed_updates=Update.ALL_TYPES)
