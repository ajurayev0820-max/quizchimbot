import csv
import logging
import os
import random
from threading import Thread
import docx
from flask import Flask
import openpyxl
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)
import xlrd

BOT_TOKEN = "8885503132:AAFyCJmyo0oLDLNiK0agg0URqc1rDuG7DHQ"

logging.basicConfig(level=logging.INFO)

# --- RENDER UCHUN FLASK SERVER ---
web_app = Flask("")


@web_app.route("/")
def home():
    return "Bot 24/7 ishlamoqda!"


def run_web():
    port = int(os.environ.get("PORT", 8080))
    web_app.run(host="0.0.0.0", port=port)


# ---------------------------------


def clean_row(row):
    """Jadvaldagi bo'sh va keraksiz kataklarni tozalash"""
    cleaned = [str(cell).strip() for cell in row if cell is not None]
    if not cleaned:
        return []

    # Agar 1-ustun tartib raqami bo'lsa (masalan: 1, 2, №, T/r), uni olib tashlaymiz
    first_cell = cleaned[0].lower().replace(".", "")
    if (
        first_cell.isdigit()
        or first_cell in ["№", "nr", "t/r", "no"]
        or "test" in first_cell
    ):
        cleaned = cleaned[1:]

    return cleaned


def read_file_data(file_path):
    rows = []

    if file_path.endswith(".csv"):
        with open(file_path, mode="r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            for row in reader:
                r = clean_row(row)
                if len(r) >= 2:
                    rows.append(r)

    elif file_path.endswith(".xlsx"):
        wb = openpyxl.load_workbook(file_path, data_only=False)
        sheet = wb.active
        for row in sheet.iter_rows():
            row_vals = [cell.value for cell in row]
            r = clean_row(row_vals)
            if len(r) >= 2:
                rows.append(r)

    elif file_path.endswith(".xls"):
        wb = xlrd.open_workbook(file_path, formatting_info=True)
        sheet = wb.sheet_by_index(0)
        for row_idx in range(sheet.nrows):
            row_vals = [
                sheet.cell(row_idx, col_idx).value
                for col_idx in range(sheet.ncols)
            ]
            r = clean_row(row_vals)
            if len(r) >= 2:
                rows.append(r)

    elif file_path.endswith(".docx") or file_path.endswith(".doc"):
        # docx kutubxonasi orqali jadvallarni o'qish
        try:
            doc = docx.Document(file_path)
            for table in doc.tables:
                for row in table.rows:
                    row_vals = [cell.text for cell in row.cells]
                    r = clean_row(row_vals)
                    if len(r) >= 2:
                        rows.append(r)
        except Exception:
            pass

    return rows


def convert_data(rows):
    quizmaker_text = ""
    assyst_text = ""
    valid_q_num = 1

    for row in rows:
        # Sarlavha qatorlarini (masalan: Savol, To'g'ri javob) o'tkazib yuborish
        if "savol" in row[0].lower() or "topshiriq" in row[0].lower():
            continue

        question = row[0]
        correct_ans = row[1]
        wrong_answers = [
            cell for cell in row[2:] if cell and cell.lower() != "none"
        ]

        if not question or not correct_ans:
            continue

        # 1. QuizMaker
        all_options = [correct_ans] + wrong_answers
        random.shuffle(all_options)
        correct_option_number = all_options.index(correct_ans) + 1

        labels = ["A", "B", "C", "D", "E", "F"]
        quizmaker_text += f"{valid_q_num}. {question}\n"
        for idx, opt in enumerate(all_options):
            label = labels[idx] if idx < len(labels) else f"Option {idx+1}"
            quizmaker_text += f"{label}. {opt}\n"
        quizmaker_text += f"{correct_option_number}\n\n"

        # 2. Assyst
        assyst_text += f"? {question}\n"
        assyst_text += f"+ {correct_ans}\n"
        for w_ans in wrong_answers:
            assyst_text += f"- {w_ans}\n"
        assyst_text += "\n"

        valid_q_num += 1

    return quizmaker_text.strip(), assyst_text.strip()


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Xush kelibsiz! Menga testlar joylashgan Excel (.xlsx, .xls), CSV yoki Word (.docx, .doc) faylini yuboring."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    file_name = doc.file_name
    ext = os.path.splitext(file_name)[1].lower()

    if ext not in [".xlsx", ".xls", ".csv", ".docx", ".doc"]:
        await update.message.reply_text(
            "Faqat .xlsx, .xls, .csv, .docx yoki .doc formatidagi fayllarni yuboring!"
        )
        return

    await update.message.reply_text("Fayl qabul qilindi, ishlanmoqda...")

    tg_file = await context.bot.get_file(doc.file_id)
    download_path = f"temp_{file_name}"
    await tg_file.download_to_drive(download_path)

    try:
        rows = read_file_data(download_path)

        if not rows:
            await update.message.reply_text(
                "Fayl bo'sh yoki jadval formati mos kelmadi."
            )
            os.remove(download_path)
            return

        quizmaker_txt, assyst_txt = convert_data(rows)

        base_name = os.path.splitext(file_name)[0]
        qm_path = f"{base_name}_QuizMaker.txt"
        assyst_path = f"{base_name}_Assyst.txt"

        with open(qm_path, "w", encoding="utf-8") as f:
            f.write(quizmaker_txt)

        with open(assyst_path, "w", encoding="utf-8") as f:
            f.write(assyst_txt)

        await update.message.reply_document(
            document=open(qm_path, "rb"), caption="QuizMaker formati"
        )
        await update.message.reply_document(
            document=open(assyst_path, "rb"), caption="Assyst formati"
        )

        if os.path.exists(download_path):
            os.remove(download_path)
        if os.path.exists(qm_path):
            os.remove(qm_path)
        if os.path.exists(assyst_path):
            os.remove(assyst_path)

    except Exception as e:
        await update.message.reply_text(f"Xatolik yuz berdi: {e}")
        if os.path.exists(download_path):
            os.remove(download_path)


if __name__ == "__main__":
    Thread(target=run_web).start()

    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(
        MessageHandler(filters.Document.ALL, handle_document)
    )

    print("Bot ishga tushdi...")
    app.run_polling()
    
