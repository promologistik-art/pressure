import os
import re
from datetime import datetime, timedelta, time
import pytz
import asyncio
from dotenv import load_dotenv
from telegram import Update, BotCommand
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment

load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID"))
EXCEL_FILE = os.getenv("EXCEL_FILE", "pressure_journal.xlsx")

MSK = pytz.timezone('Europe/Moscow')

# Хранилище выбранного периода для пользователя
user_period = {}

# ==================== ПРОВЕРКА ПРАВ ====================
def is_admin(update: Update) -> bool:
    return update.effective_user.id == ADMIN_ID

# ==================== СОЗДАНИЕ EXCEL ПО ШАБЛОНУ ====================
def init_excel():
    """Создаёт Excel файл с шапкой: утро, обед, вечер"""
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        
        # Удаляем дефолтный лист
        default_sheet = wb.active
        wb.remove(default_sheet)
        
        # Создаём основной лист
        ws = wb.create_sheet("Давление")
        
        # === ШАПКА ===
        # Строка 1: объединённые ячейки для "утро", "обед", "вечер"
        ws.merge_cells('A1:E1')
        ws['A1'] = 'утро'
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws['A1'].font = Font(bold=True)
        
        ws.merge_cells('F1:J1')
        ws['F1'] = 'обед'
        ws['F1'].alignment = Alignment(horizontal='center', vertical='center')
        ws['F1'].font = Font(bold=True)
        
        ws.merge_cells('K1:O1')
        ws['K1'] = 'вечер'
        ws['K1'].alignment = Alignment(horizontal='center', vertical='center')
        ws['K1'].font = Font(bold=True)
        
        # Строка 2: подзаголовки
        headers_row2 = [
            'Дата', 'Время', 'Систолическое', 'Диастолическое', 'Пульс',
            'Время', 'Систолическое', 'Диастолическое', 'Пульс',
            'Время', 'Систолическое', 'Диастолическое', 'Пульс'
        ]
        
        for col, header in enumerate(headers_row2, 1):
            cell = ws.cell(row=2, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Настройка ширины колонок
        column_widths = {
            'A': 12,  # Дата
            'B': 10,  # Время утро
            'C': 14,  # Систолическое утро
            'D': 14,  # Диастолическое утро
            'E': 8,   # Пульс утро
            'F': 10,  # Время обед
            'G': 14,  # Систолическое обед
            'H': 14,  # Диастолическое обед
            'I': 8,   # Пульс обед
            'J': 10,  # Время вечер
            'K': 14,  # Систолическое вечер
            'L': 14,  # Диастолическое вечер
            'M': 8    # Пульс вечер
        }
        
        for col_letter, width in column_widths.items():
            ws.column_dimensions[col_letter].width = width
        
        # Высота строк
        ws.row_dimensions[1].height = 25
        ws.row_dimensions[2].height = 20
        
        wb.save(EXCEL_FILE)
        print(f"✅ Создан файл {EXCEL_FILE} по шаблону")

def save_to_excel(period: str, systolic: int, diastolic: int, pulse: int = None):
    """
    Сохраняет показания в Excel
    period: 'утро', 'обед', 'вечер'
    """
    now = datetime.now(MSK)
    date_str = now.strftime("%d-%m-%Y")
    time_str = now.strftime("%H:%M:%S")
    
    # Колонки: A=1 Дата, B=2 Время утро, C=3 САД утро, D=4 ДАД утро, E=5 Пульс утро
    #          F=6 Время обед, G=7 САД обед, H=8 ДАД обед, I=9 Пульс обед
    #          J=10 Время вечер, K=11 САД вечер, L=12 ДАД вечер, M=13 Пульс вечер
    period_cols = {
        'утро':   {'time_col': 2, 'systolic_col': 3, 'diastolic_col': 4, 'pulse_col': 5},
        'обед':   {'time_col': 6, 'systolic_col': 7, 'diastolic_col': 8, 'pulse_col': 9},
        'вечер':  {'time_col': 10, 'systolic_col': 11, 'diastolic_col': 12, 'pulse_col': 13}
    }
    
    cols = period_cols[period]
    
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb["Давление"]
    else:
        init_excel()
        wb = load_workbook(EXCEL_FILE)
        ws = wb["Давление"]
    
    # Ищем строку с такой же датой
    target_row = None
    for row in range(3, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == date_str:
            target_row = row
            break
    
    # Если даты нет, создаём новую строку
    if target_row is None:
        target_row = ws.max_row + 1
        ws.cell(row=target_row, column=1, value=date_str)
    
    # Заполняем данные
    ws.cell(row=target_row, column=cols['time_col'], value=time_str)
    ws.cell(row=target_row, column=cols['systolic_col'], value=systolic)
    ws.cell(row=target_row, column=cols['diastolic_col'], value=diastolic)
    if pulse:
        ws.cell(row=target_row, column=cols['pulse_col'], value=pulse)
    
    wb.save(EXCEL_FILE)

# ==================== КОМАНДЫ ====================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        await update.message.reply_text("⛔ Доступ запрещён.")
        return
    
    await update.message.reply_text(
        "🩸 **Бот контроля давления**\n\n"
        "**Доступные команды:**\n"
        "/morning — записать утреннее давление\n"
        "/afternoon — записать обеденное давление\n"
        "/evening — записать вечернее давление\n"
        "/table — получить Excel файл с журналом\n"
        "/remind — получить напоминание сейчас\n"
        "/help — помощь\n\n"
        "**Пример ввода:**\n"
        "`130 85` — только давление\n"
        "`130 85 72` — давление и пульс\n"
        "`130/85` — через слеш",
        parse_mode="Markdown"
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    await update.message.reply_text(
        "📖 **Инструкция**\n\n"
        "**1. Запись давления:**\n"
        "/morning — утренний замер\n"
        "/afternoon — обеденный замер\n"
        "/evening — вечерний замер\n\n"
        "**2. Форматы ввода:**\n"
        "• `130 85` — давление\n"
        "• `130 85 72` — давление и пульс\n"
        "• `130/85` — через слеш\n\n"
        "**3. Получение данных:**\n"
        "/table — скачать Excel файл\n\n"
        "**4. Напоминания:**\n"
        "Автоматические в 8:00, 14:00, 20:00 МСК\n"
        "/remind — ручное напоминание\n\n"
        "**Формат даты:** ДД-ММ-ГГГГ (российский)",
        parse_mode="Markdown"
    )

async def table_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    if not os.path.exists(EXCEL_FILE):
        await update.message.reply_text("📭 Журнал пока пуст. Добавьте хотя бы одно измерение.")
        return
    
    with open(EXCEL_FILE, 'rb') as f:
        await update.message.reply_document(
            document=f,
            filename="pressure_journal.xlsx",
            caption="📊 Журнал давления\n\n📅 Формат даты: ДД-ММ-ГГГГ\n🌅 Утро | ☀️ Обед | 🌙 Вечер"
        )

async def remind_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    await update.message.reply_text(
        "🔔 **Напоминание**\n\n"
        "Пора измерить давление!\n\n"
        "Используйте команды:\n"
        "/morning — утро\n"
        "/afternoon — обед\n"
        "/evening — вечер",
        parse_mode="Markdown"
    )

async def morning_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    user_id = update.effective_user.id
    user_period[user_id] = 'утро'
    await update.message.reply_text(
        "🌅 **Утренний замер**\n\n"
        "Введите показания в формате:\n"
        "• `130 85` — давление\n"
        "• `130 85 72` — давление и пульс\n"
        "• `130/85` — через слеш\n\n"
        "Например: `120 80 68`",
        parse_mode="Markdown"
    )

async def afternoon_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    user_id = update.effective_user.id
    user_period[user_id] = 'обед'
    await update.message.reply_text(
        "☀️ **Обеденный замер**\n\n"
        "Введите показания в формате:\n"
        "• `130 85` — давление\n"
        "• `130 85 72` — давление и пульс\n"
        "• `130/85` — через слеш\n\n"
        "Например: `120 80 68`",
        parse_mode="Markdown"
    )

async def evening_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    user_id = update.effective_user.id
    user_period[user_id] = 'вечер'
    await update.message.reply_text(
        "🌙 **Вечерний замер**\n\n"
        "Введите показания в формате:\n"
        "• `130 85` — давление\n"
        "• `130 85 72` — давление и пульс\n"
        "• `130/85` — через слеш\n\n"
        "Например: `120 80 68`",
        parse_mode="Markdown"
    )

async def handle_pressure(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update):
        return
    
    user_id = update.effective_user.id
    
    # Проверяем, выбран ли период
    if user_id not in user_period:
        await update.message.reply_text(
            "❌ Сначала выберите время замера.\n\n"
            "Используйте команды:\n"
            "/morning — утро\n"
            "/afternoon — обед\n"
            "/evening — вечер"
        )
        return
    
    period = user_period[user_id]
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
            "❌ Не понял показания.\n\n"
            "Форматы:\n"
            "• `130 85`\n"
            "• `130 85 72`\n"
            "• `130/85`",
            parse_mode="Markdown"
        )
        return
    
    # Сохраняем в Excel
    save_to_excel(period, systolic, diastolic, pulse)
    
    # Очищаем выбранный период
    del user_period[user_id]
    
    period_emoji = {'утро': '🌅', 'обед': '☀️', 'вечер': '🌙'}
    now = datetime.now(MSK)
    
    await update.message.reply_text(
        f"✅ **Записано!**\n\n"
        f"{period_emoji[period]} {period}: {systolic}/{diastolic}" + (f", пульс {pulse}" if pulse else "") + f"\n\n"
        f"📅 {now.strftime('%d-%m-%Y')} {now.strftime('%H:%M:%S')}",
        parse_mode="Markdown"
    )

async def send_scheduled_reminder(context: ContextTypes.DEFAULT_TYPE):
    await context.bot.send_message(
        chat_id=ADMIN_ID,
        text="🔔 **Плановое напоминание**\n\nПора измерить давление!\n\n"
             "Используйте команды:\n"
             "/morning — утро\n"
             "/afternoon — обед\n"
             "/evening — вечер",
        parse_mode="Markdown"
    )

async def set_commands(app):
    """Устанавливает команды для кнопки меню (≡) в левом нижнем углу"""
    commands = [
        BotCommand("start", "Главное меню"),
        BotCommand("morning", "Записать утреннее давление"),
        BotCommand("afternoon", "Записать обеденное давление"),
        BotCommand("evening", "Записать вечернее давление"),
        BotCommand("table", "Получить журнал Excel"),
        BotCommand("remind", "Напоминание измерить давление"),
        BotCommand("help", "Помощь"),
    ]
    await app.bot.set_my_commands(commands)

# ==================== ЗАПУСК ====================
def main():
    init_excel()
    
    # Создаём новый событийный цикл
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    
    app = Application.builder().token(TOKEN).build()
    
    # Команды
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("table", table_command))
    app.add_handler(CommandHandler("remind", remind_command))
    app.add_handler(CommandHandler("morning", morning_command))
    app.add_handler(CommandHandler("afternoon", afternoon_command))
    app.add_handler(CommandHandler("evening", evening_command))
    
    # Обработчик текстовых сообщений (показаний)
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pressure))
    
    # Устанавливаем команды для кнопки меню (≡)
    loop.run_until_complete(set_commands(app))
    
    # Напоминания по МСК (8:00, 14:00, 20:00)
    job_queue = app.job_queue
    if job_queue:
        job_queue.run_daily(send_scheduled_reminder, time=time(5, 0))   # 8:00 МСК
        job_queue.run_daily(send_scheduled_reminder, time=time(11, 0))  # 14:00 МСК
        job_queue.run_daily(send_scheduled_reminder, time=time(17, 0))  # 20:00 МСК
        print("⏰ Напоминания: 8:00, 14:00, 20:00 МСК")
    
    print("🤖 Бот запущен")
    print("📋 Команды появятся в кнопке меню (≡) в левом нижнем углу")
    
    # Запускаем бота
    app.run_polling()

if __name__ == "__main__":
    main()