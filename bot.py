import os
import re
import asyncio
from datetime import datetime, timedelta
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import pandas as pd

load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID"))
EXCEL_FILE = os.getenv("EXCEL_FILE", "pressure_journal.xlsx")

# Проверка прав доступа
def is_admin(update: Update) -> bool:
    return update.effective_user.id == ADMIN_ID

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=[
            "Дата", "Время", "Систолическое", "Диастолическое",
            "Пульс", "Самочувствие", "Примечания"
        ])
        df.to_excel(EXCEL_FILE, index=False)

def save_to_excel(systolic, diastolic, pulse, feeling="", notes=""):
    now = datetime.now()
    new_row = pd.DataFrame([{
        "Дата": now.strftime("%Y-%m-%d"),
        "Время": now.strftime("%H:%M"),
        "Систолическое": systolic,
        "Диастолическое": diastolic,
        "Пульс": pulse if pulse else "",
        "Самочувствие": feeling,
        "Примечания": notes
    }])
    
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        df = pd.concat([df, new_row], ignore_index=True)
    else:
        df = new_row
    df.to_excel(EXCEL_FILE, index=False)

def get_recent_table(days=7):
    if not os.path.exists(EXCEL_FILE):
        return None
    df = pd.read_excel(EXCEL_FILE)
    cutoff = datetime.now() - timedelta(days=days)
    df["Дата"] = pd.to_datetime(df["Дата"])
    recent = df[df["Дата"] >= cutoff].copy()
    recent["Дата"] = recent["Дата"].dt.strftime("%d.%m")
    return recent[["Дата", "Время", "Систолическое", "Диастолическое", "Пульс"]]

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("⛔ Доступ запрещён. Вы не админ.")
        return
    
    await update.message.reply_text(
        "🩸 **Бот контроля давления**\n\n"
        "Введите показания в любом формате:\n"
        "`130 85 72` — давление и пульс\n"
        "`130/85` — только давление\n"
        "`145/90 голова болит` — с самочувствием\n\n"
        "**Команды:**\n"
        "/table — таблица за 7 дней\n"
        "/table 30 — таблица за 30 дней\n"
        "/help — помощь",
        parse_mode="Markdown"
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    await update.message.reply_text(
        "📖 **Как пользоваться:**\n\n"
        "1️⃣ Измерьте давление\n"
        "2️⃣ Напишите боту цифры\n"
        "3️⃣ По желанию добавьте самочувствие\n\n"
        "Примеры:\n"
        "• `125 80 68`\n"
        "• `130/85`\n"
        "• `145/90 кружится голова`"
    )

async def table_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    args = context.args
    days = int(args[0]) if args and args[0].isdigit() else 7
    
    df = get_recent_table(days)
    if df is None or df.empty:
        await update.message.reply_text("📭 Нет записей за указанный период.")
        return
    
    # Формируем текстовую таблицу
    table_text = f"📊 *Журнал давления (последние {days} дней)*\n\n"
    for _, row in df.iterrows():
        table_text += f"`{row['Дата']} {row['Время']}  {int(row['Систолическое'])}/{int(row['Диастолическое'])}  пульс {int(row['Пульс'])}`\n"
    
    if len(table_text) > 4000:
        # Отправляем файлом, если текст слишком длинный
        with open(EXCEL_FILE, "rb") as f:
            await update.message.reply_document(f, filename=f"давление_{days}дней.xlsx")
    else:
        await update.message.reply_text(table_text, parse_mode="Markdown")

async def handle_pressure(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    text = update.message.text.strip()
    numbers = re.findall(r'\d+', text)
    
    # Парсинг давления и пульса
    systolic = None
    diastolic = None
    pulse = None
    
    # Ищем форматы: 130/85 или 130 85
    slash_match = re.search(r'(\d{2,3})/(\d{2,3})', text)
    if slash_match:
        systolic = int(slash_match.group(1))
        diastolic = int(slash_match.group(2))
    elif len(numbers) >= 2:
        systolic = int(numbers[0])
        diastolic = int(numbers[1])
    
    # Ищем пульс (обычно 50-150)
    for num in numbers:
        if 40 <= int(num) <= 150 and int(num) != systolic and int(num) != diastolic:
            pulse = int(num)
            break
    
    if not systolic or not diastolic:
        await update.message.reply_text(
            "❌ Не распознал показания.\n"
            "Пример: `130 85 72` или `130/85`",
            parse_mode="Markdown"
        )
        return
    
    # Определяем самочувствие
    feeling = ""
    if any(word in text.lower() for word in ["голов", "болит", "давит"]):
        feeling = "головная боль"
    elif "круж" in text.lower():
        feeling = "головокружение"
    elif "слабость" in text.lower():
        feeling = "слабость"
    elif "норм" in text.lower() or "хорош" in text.lower():
        feeling = "хорошо"
    
    save_to_excel(systolic, diastolic, pulse, feeling)
    
    # Оценка давления
    if systolic < 120 and diastolic < 80:
        status = "✅ Норма"
    elif systolic < 130 and diastolic < 85:
        status = "🟢 Повышенная норма"
    elif systolic < 140 or diastolic < 90:
        status = "⚠️ Высокое нормальное"
    else:
        status = "🔴 Повышенное! Обратите внимание"
    
    response = f"✅ Записано: *{systolic}/{diastolic}*"
    if pulse:
        response += f", пульс *{pulse}*"
    response += f"\n{status}"
    
    await update.message.reply_text(response, parse_mode="Markdown")

async def send_reminder(context: ContextTypes.DEFAULT_TYPE):
    """Функция для напоминаний (запускается по расписанию)"""
    await context.bot.send_message(
        chat_id=ADMIN_ID,
        text="🔔 **Напоминание**\nПора измерить давление!\n\n"
             "После измерения просто напишите мне цифры, например:\n"
             "`125 80 68`",
        parse_mode="Markdown"
    )

def main():
    init_excel()
    
    app = Application.builder().token(TOKEN).build()
    
    # Команды
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("table", table_command))
    
    # Обработка сообщений с показаниями
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pressure))
    
    # Настройка напоминаний (утром и вечером)
    job_queue = app.job_queue
    
    # Утреннее напоминание в 9:00
    job_queue.run_daily(
        send_reminder,
        time=datetime.time(hour=9, minute=0),
        days=tuple(range(7))
    )
    
    # Вечернее напоминание в 20:00
    job_queue.run_daily(
        send_reminder,
        time=datetime.time(hour=20, minute=0),
        days=tuple(range(7))
    )
    
    print("🤖 Бот запущен...")
    app.run_polling()

if __name__ == "__main__":
    main()