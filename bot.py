import os
import re
from datetime import datetime, timedelta, time
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID"))
EXCEL_FILE = os.getenv("EXCEL_FILE", "pressure_journal.xlsx")

def is_admin(update: Update) -> bool:
    return update.effective_user.id == ADMIN_ID

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Давление"
        headers = ["Дата", "Время", "Систолическое", "Диастолическое", "Пульс"]
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
            ws.cell(row=1, column=col).font = Font(bold=True)
        wb.save(EXCEL_FILE)

def save_to_excel(systolic: int, diastolic: int, pulse: int = None):
    now = datetime.now()
    
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Давление"
        headers = ["Дата", "Время", "Систолическое", "Диастолическое", "Пульс"]
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
    
    next_row = ws.max_row + 1
    ws.cell(row=next_row, column=1, value=now.strftime("%Y-%m-%d"))
    ws.cell(row=next_row, column=2, value=now.strftime("%H:%M"))
    ws.cell(row=next_row, column=3, value=systolic)
    ws.cell(row=next_row, column=4, value=diastolic)
    ws.cell(row=next_row, column=5, value=pulse if pulse else "")
    
    wb.save(EXCEL_FILE)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("⛔ Доступ запрещён.")
        return
    
    await update.message.reply_text(
        "🩸 Бот контроля давления\n\n"
        "Введите показания:\n"
        "130 85 72 — давление и пульс\n"
        "130/85 — только давление\n\n"
        "/table — получить Excel файл с журналом"
    )

async def table_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    if not os.path.exists(EXCEL_FILE):
        await update.message.reply_text("Журнал пуст.")
        return
    
    with open(EXCEL_FILE, 'rb') as f:
        await update.message.reply_document(
            document=f,
            filename="pressure_journal.xlsx",
            caption="📊 Журнал давления"
        )

async def handle_pressure(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    text = update.message.text.strip()
    numbers = re.findall(r'\d+', text)
    
    systolic = None
    diastolic = None
    pulse = None
    
    slash_match = re.search(r'(\d{2,3})/(\d{2,3})', text)
    if slash_match:
        systolic = int(slash_match.group(1))
        diastolic = int(slash_match.group(2))
    elif len(numbers) >= 2:
        systolic = int(numbers[0])
        diastolic = int(numbers[1])
    
    if systolic and diastolic:
        for num in numbers:
            num_int = int(num)
            if 40 <= num_int <= 150 and num_int != systolic and num_int != diastolic:
                pulse = num_int
                break
    
    if not systolic or not diastolic:
        await update.message.reply_text("Не понял. Пример: 130 85 72 или 130/85")
        return
    
    save_to_excel(systolic, diastolic, pulse)
    await update.message.reply_text(f"✅ Записано: {systolic}/{diastolic}" + (f", пульс {pulse}" if pulse else ""))

async def send_reminder(context: ContextTypes.DEFAULT_TYPE):
    await context.bot.send_message(
        chat_id=ADMIN_ID,
        text="🔔 Напоминание: пора измерить давление"
    )

def main():
    init_excel()
    
    app = Application.builder().token(TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("table", table_command))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pressure))
    
    job_queue = app.job_queue
    if job_queue:
        # 8:00 по Москве (UTC+3)
        job_queue.run_daily(send_reminder, time=time(5, 0))  # 5:00 UTC = 8:00 MSK
        print("⏰ Напоминание установлено на 8:00 МСК")
    
    print("🤖 Бот запущен")
    app.run_polling()

if __name__ == "__main__":
    main()