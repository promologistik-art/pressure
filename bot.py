import os
import re
from datetime import datetime, timedelta, time  # ← правильный импорт time
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID"))
EXCEL_FILE = os.getenv("EXCEL_FILE", "pressure_journal.xlsx")

# ==================== ПРОВЕРКА ПРАВ ====================
def is_admin(update: Update) -> bool:
    return update.effective_user.id == ADMIN_ID

# ==================== РАБОТА С EXCEL ====================
def init_excel():
    """Создаёт Excel-файл с заголовками, если его нет"""
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Давление"
        
        headers = ["Дата", "Время", "Систолическое", "Диастолическое", "Пульс", "Самочувствие", "Примечания"]
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
            ws.cell(row=1, column=col).font = Font(bold=True)
        
        wb.save(EXCEL_FILE)
        print(f"✅ Создан файл {EXCEL_FILE}")

def save_to_excel(systolic: int, diastolic: int, pulse: int = None, feeling: str = "", notes: str = ""):
    """Сохраняет показания в Excel"""
    now = datetime.now()
    
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Давление"
        headers = ["Дата", "Время", "Систолическое", "Диастолическое", "Пульс", "Самочувствие", "Примечания"]
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
    
    next_row = ws.max_row + 1
    
    ws.cell(row=next_row, column=1, value=now.strftime("%Y-%m-%d"))
    ws.cell(row=next_row, column=2, value=now.strftime("%H:%M"))
    ws.cell(row=next_row, column=3, value=systolic)
    ws.cell(row=next_row, column=4, value=diastolic)
    ws.cell(row=next_row, column=5, value=pulse if pulse else "")
    ws.cell(row=next_row, column=6, value=feeling)
    ws.cell(row=next_row, column=7, value=notes)
    
    wb.save(EXCEL_FILE)

def get_recent_table(days: int = 7):
    """Возвращает список записей за последние N дней"""
    if not os.path.exists(EXCEL_FILE):
        return []
    
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    
    records = []
    cutoff = datetime.now() - timedelta(days=days)
    
    for row in range(2, ws.max_row + 1):
        date_str = ws.cell(row=row, column=1).value
        if not date_str:
            continue
        
        try:
            record_date = datetime.strptime(str(date_str), "%Y-%m-%d")
            if record_date >= cutoff:
                records.append({
                    "Дата": record_date.strftime("%d.%m"),
                    "Время": ws.cell(row=row, column=2).value,
                    "Систолическое": ws.cell(row=row, column=3).value,
                    "Диастолическое": ws.cell(row=row, column=4).value,
                    "Пульс": ws.cell(row=row, column=5).value
                })
        except:
            continue
    
    return records

def get_statistics(days: int = None):
    """Возвращает статистику за период"""
    if not os.path.exists(EXCEL_FILE):
        return None
    
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    
    records = []
    cutoff = datetime.now() - timedelta(days=days) if days else None
    
    for row in range(2, ws.max_row + 1):
        date_str = ws.cell(row=row, column=1).value
        if not date_str:
            continue
        
        try:
            record_date = datetime.strptime(str(date_str), "%Y-%m-%d")
            if cutoff is None or record_date >= cutoff:
                systolic = ws.cell(row=row, column=3).value
                diastolic = ws.cell(row=row, column=4).value
                pulse = ws.cell(row=row, column=5).value
                
                if systolic and diastolic:
                    records.append({
                        "systolic": float(systolic),
                        "diastolic": float(diastolic),
                        "pulse": float(pulse) if pulse else None
                    })
        except:
            continue
    
    if not records:
        return None
    
    systolic_list = [r["systolic"] for r in records]
    diastolic_list = [r["diastolic"] for r in records]
    pulse_list = [r["pulse"] for r in records if r["pulse"]]
    
    high_count = sum(1 for s, d in zip(systolic_list, diastolic_list) if s >= 135 or d >= 85)
    
    period_text = f"последние {days} дней" if days else "всё время"
    
    return {
        "period": period_text,
        "count": len(records),
        "avg_systolic": sum(systolic_list) / len(systolic_list),
        "avg_diastolic": sum(diastolic_list) / len(diastolic_list),
        "avg_pulse": sum(pulse_list) / len(pulse_list) if pulse_list else 0,
        "max_systolic": max(systolic_list),
        "min_systolic": min(systolic_list),
        "max_diastolic": max(diastolic_list),
        "high_percent": (high_count / len(records)) * 100
    }

# ==================== ОБРАБОТЧИКИ КОМАНД ====================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("⛔ Доступ запрещён.")
        return
    
    await update.message.reply_text(
        "🩸 **Бот контроля давления**\n\n"
        "📝 **Ввод показаний:**\n"
        "`130 85 72` — давление и пульс\n"
        "`130/85` — только давление\n"
        "`145/90 голова болит` — с самочувствием\n\n"
        "📋 **Команды:**\n"
        "/table — таблица за 7 дней\n"
        "/table 14 — таблица за 14 дней\n"
        "/stats — статистика\n"
        "/remind — напоминание\n"
        "/help — помощь",
        parse_mode="Markdown"
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    await update.message.reply_text(
        "📖 **Инструкция**\n\n"
        "**Форматы ввода:**\n"
        "• `125 80` — давление\n"
        "• `125 80 68` — давление и пульс\n"
        "• `130/85` — через слеш\n"
        "• `145/90 кружится голова` — с самочувствием\n\n"
        "**Команды:**\n"
        "• `/table 5` — таблица за 5 дней\n"
        "• `/table неделя` — за неделю\n"
        "• `/table месяц` — за месяц\n"
        "• `/stats 14` — статистика за 14 дней\n"
        "• `/remind` — получить напоминание\n\n"
        "**Оценка давления:**\n"
        "• ✅ менее 120/80 — норма\n"
        "• 🟢 120-129/80-84 — повыш. норма\n"
        "• ⚠️ 130-139/85-89 — высокое норм.\n"
        "• 🔴 140+ / 90+ — повышено",
        parse_mode="Markdown"
    )

async def table_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    args = context.args
    days = 7
    
    if args:
        arg = args[0].lower()
        if arg in ["неделя", "week", "7"]:
            days = 7
        elif arg in ["две", "2недели", "2weeks", "14"]:
            days = 14
        elif arg in ["месяц", "month", "30"]:
            days = 30
        elif arg.isdigit():
            days = int(arg)
        else:
            await update.message.reply_text("❌ Используйте: `/table 14` или `/table неделя`", parse_mode="Markdown")
            return
    
    records = get_recent_table(days)
    
    if not records:
        await update.message.reply_text(f"📭 Нет записей за последние {days} дней.")
        return
    
    table_text = f"📊 *Журнал давления (последние {days} дней)*\n\n"
    table_text += "```\n"
    table_text += f"{'Дата':<8} {'Время':<6} {'Давление':<9} {'Пульс':<6}\n"
    table_text += "-" * 35 + "\n"
    
    for r in records:
        pressure = f"{int(r['Систолическое'])}/{int(r['Диастолическое'])}"
        pulse = f"{int(r['Пульс'])}" if r['Пульс'] else "-"
        table_text += f"{r['Дата']:<8} {r['Время']:<6} {pressure:<9} {pulse:<6}\n"
    
    table_text += "```"
    
    if len(table_text) > 4000:
        with open(EXCEL_FILE, "rb") as f:
            await update.message.reply_document(f, filename=f"давление_{days}дней.xlsx")
    else:
        await update.message.reply_text(table_text, parse_mode="Markdown")

async def stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    args = context.args
    days = None
    
    if args:
        arg = args[0].lower()
        if arg in ["неделя", "week", "7"]:
            days = 7
        elif arg in ["две", "2недели", "2weeks", "14"]:
            days = 14
        elif arg in ["месяц", "month", "30"]:
            days = 30
        elif arg.isdigit():
            days = int(arg)
    
    stats = get_statistics(days)
    
    if not stats:
        period = f"за последние {days} дней" if days else ""
        await update.message.reply_text(f"📭 Нет данных {period}.")
        return
    
    response = f"📊 **Статистика давления ({stats['period']})**\n\n"
    response += f"📏 **Всего измерений:** {stats['count']}\n\n"
    response += f"🩸 **Среднее давление:** {stats['avg_systolic']:.0f}/{stats['avg_diastolic']:.0f}\n"
    response += f"❤️ **Средний пульс:** {stats['avg_pulse']:.0f}\n\n"
    response += f"📈 **Максимальное давление:** {stats['max_systolic']:.0f}/{stats['max_diastolic']:.0f}\n"
    response += f"📉 **Минимальное давление:** {stats['min_systolic']:.0f}\n\n"
    response += f"⚠️ **Выше 135/85:** {stats['high_percent']:.1f}%\n"
    
    if stats['avg_systolic'] < 120 and stats['avg_diastolic'] < 80:
        response += "\n✅ Отличный контроль!"
    elif stats['avg_systolic'] < 130 and stats['avg_diastolic'] < 85:
        response += "\n🟢 Хороший контроль"
    elif stats['avg_systolic'] < 140 or stats['avg_diastolic'] < 90:
        response += "\n⚠️ Требуется внимание"
    else:
        response += "\n🔴 Проконсультируйтесь с врачом!"
    
    await update.message.reply_text(response, parse_mode="Markdown")

async def remind_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    await update.message.reply_text(
        "🔔 **Напоминание**\n\nПора измерить давление!\n\nНапишите цифры, например:\n`125 80 68`",
        parse_mode="Markdown"
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
        await update.message.reply_text(
            "❌ Не распознал.\nФорматы: `130 85` или `130/85`",
            parse_mode="Markdown"
        )
        return
    
    feeling = ""
    if any(w in text.lower() for w in ["голов", "болит"]):
        feeling = "головная боль"
    elif "круж" in text.lower():
        feeling = "головокружение"
    elif "слабость" in text.lower():
        feeling = "слабость"
    elif "норм" in text.lower() or "хорош" in text.lower():
        feeling = "хорошо"
    
    save_to_excel(systolic, diastolic, pulse, feeling)
    
    if systolic < 120 and diastolic < 80:
        status = "✅ Норма"
    elif systolic < 130 and diastolic < 85:
        status = "🟢 Повышенная норма"
    elif systolic < 140 or diastolic < 90:
        status = "⚠️ Высокое нормальное"
    else:
        status = "🔴 Повышенное!"
    
    response = f"✅ Записано: *{systolic}/{diastolic}*"
    if pulse:
        response += f", пульс *{pulse}*"
    response += f"\n{status}"
    
    if systolic >= 160 or diastolic >= 100:
        response += "\n\n🚨 **ВНИМАНИЕ!** Очень высокое давление!"
    
    await update.message.reply_text(response, parse_mode="Markdown")

async def send_scheduled_reminder(context: ContextTypes.DEFAULT_TYPE):
    await context.bot.send_message(
        chat_id=ADMIN_ID,
        text="🔔 **Плановое напоминание**\n\nПора измерить давление!\nНапишите цифры, например: `125 80 68`",
        parse_mode="Markdown"
    )

# ==================== ЗАПУСК БОТА ====================
def main():
    init_excel()
    
    app = Application.builder().token(TOKEN).build()
    
    # Команды
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("table", table_command))
    app.add_handler(CommandHandler("stats", stats_command))
    app.add_handler(CommandHandler("remind", remind_command))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pressure))
    
    # Настройка напоминаний
    job_queue = app.job_queue
    if job_queue:
        # Утро в 9:00 — используем time из импорта
        job_queue.run_daily(
            send_scheduled_reminder,
            time=time(9, 0),  # ← исправлено: time(9, 0) вместо datetime.time()
            days=tuple(range(7))
        )
        # Вечер в 20:00
        job_queue.run_daily(
            send_scheduled_reminder,
            time=time(20, 0),  # ← исправлено
            days=tuple(range(7))
        )
        print("⏰ Напоминания установлены на 9:00 и 20:00")
    else:
        print("⚠️ JobQueue не доступен, напоминания работать не будут")
    
    print(f"🤖 Бот запущен! Админ ID: {ADMIN_ID}")
    app.run_polling()

if __name__ == "__main__":
    main()