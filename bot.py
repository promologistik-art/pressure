import os
import re
import json
import asyncio
from datetime import datetime, timedelta, time
import pytz
from dotenv import load_dotenv
from telegram import Update, BotCommand, BotCommandScopeChat
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment

load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID"))
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "silverzen")
EXCEL_FILE = os.getenv("EXCEL_FILE", "pressure_journal.xlsx")
USERS_DB = os.getenv("USERS_DB", "users.json")

# Время МСК+1 (UTC+4)
MSK_PLUS_1 = pytz.timezone('Europe/Samara')

# ==================== РАБОТА С БАЗОЙ ПОЛЬЗОВАТЕЛЕЙ ====================
def load_users():
    if os.path.exists(USERS_DB):
        with open(USERS_DB, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_users(users):
    with open(USERS_DB, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def add_user(user_id, username):
    users = load_users()
    if str(user_id) not in users:
        now = datetime.now(MSK_PLUS_1)
        users[str(user_id)] = {
            "username": username,
            "joined": now.strftime("%d-%m-%Y %H:%M:%S"),
            "status": "active",
            "days_count": 1,
            "last_reminder_sent": "",
            "access_until": None
        }
        save_users(users)
        return True
    else:
        if users[str(user_id)]["username"] != username:
            users[str(user_id)]["username"] = username
            save_users(users)
    return False

def update_user_days(user_id):
    users = load_users()
    user = users.get(str(user_id))
    if user and user["status"] == "active":
        joined_str = user["joined"]
        joined = datetime.strptime(joined_str, "%d-%m-%Y %H:%M:%S")
        joined = MSK_PLUS_1.localize(joined)
        days = (datetime.now(MSK_PLUS_1) - joined).days + 1
        user["days_count"] = days
        save_users(users)
        return days
    return 0

async def check_and_send_3day_reminder(user_id, app):
    users = load_users()
    user = users.get(str(user_id))
    if user and user["status"] == "active":
        joined_str = user["joined"]
        joined = datetime.strptime(joined_str, "%d-%m-%Y %H:%M:%S")
        joined = MSK_PLUS_1.localize(joined)
        days = (datetime.now(MSK_PLUS_1) - joined).days + 1
        last_reminder = user.get("last_reminder_sent", "")
        
        if days >= 3 and last_reminder != datetime.now(MSK_PLUS_1).strftime("%d-%m-%Y"):
            user["last_reminder_sent"] = datetime.now(MSK_PLUS_1).strftime("%d-%m-%Y")
            save_users(users)
            try:
                await app.bot.send_message(
                    chat_id=int(user_id),
                    text="Вы уже 3 дня пользуетесь ботом Журнал давления.\n\nЕсли хотите продолжить, есть предложения или замечания, свяжитесь с админом"
                )
            except:
                pass
            return True
    return False

def check_access(user_id):
    users = load_users()
    user = users.get(str(user_id))
    if not user:
        return False
    
    if str(user_id) == str(ADMIN_ID):
        return True
    
    if user["status"] == "active":
        return True
    
    if user["status"] == "access":
        if user.get("access_until"):
            access_until = datetime.strptime(user["access_until"], "%d-%m-%Y")
            access_until = MSK_PLUS_1.localize(access_until)
            if datetime.now(MSK_PLUS_1) <= access_until:
                return True
            else:
                user["status"] = "blocked"
                save_users(users)
                return False
        return True
    
    return False

def grant_access(user_id, days):
    users = load_users()
    user = users.get(str(user_id))
    if user:
        user["status"] = "access"
        user["access_until"] = (datetime.now(MSK_PLUS_1) + timedelta(days=days)).strftime("%d-%m-%Y")
        save_users(users)
        return True
    return False

def is_admin(user_id):
    return user_id == ADMIN_ID

def get_all_users():
    users = load_users()
    return users

# ==================== ОПРЕДЕЛЕНИЕ ПЕРИОДА ПО ВРЕМЕНИ ====================
def get_period_by_time():
    now = datetime.now(MSK_PLUS_1)
    hour = now.hour
    
    if 6 <= hour < 12:
        return "Утро"
    elif 12 <= hour < 18:
        return "День"
    else:
        return "Вечер"

# ==================== РАБОТА С EXCEL (НОВАЯ ВЕРСИЯ) ====================
def init_excel():
    """Создаёт Excel файл с новой структурой: каждая запись в отдельной строке"""
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        default_sheet = wb.active
        wb.remove(default_sheet)
        
        ws = wb.create_sheet("Давление")
        
        # Заголовки
        headers = ['Дата', 'Время', 'Период', 'Верхнее', 'Нижнее', 'Пульс', 'Комментарий']
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Ширина колонок
        column_widths = {
            'A': 12,   # Дата
            'B': 10,   # Время
            'C': 8,    # Период
            'D': 10,   # Верхнее
            'E': 10,   # Нижнее
            'F': 8,    # Пульс
            'G': 40    # Комментарий
        }
        
        for col_letter, width in column_widths.items():
            ws.column_dimensions[col_letter].width = width
        
        ws.row_dimensions[1].height = 20
        
        wb.save(EXCEL_FILE)
        print(f"Создан файл {EXCEL_FILE} с новой структурой")

def save_to_excel(user_id, period, systolic, diastolic, pulse, comment):
    """Сохраняет показания в Excel (каждая запись в новой строке)"""
    now = datetime.now(MSK_PLUS_1)
    
    # Если время с 00:00 до 5:59 - относим к предыдущему дню
    if now.hour < 6:
        date_str = (now - timedelta(days=1)).strftime("%d-%m-%Y")
    else:
        date_str = now.strftime("%d-%m-%Y")
    
    time_str = now.strftime("%H:%M:%S")
    
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb["Давление"]
    else:
        init_excel()
        wb = load_workbook(EXCEL_FILE)
        ws = wb["Давление"]
    
    # Находим следующую пустую строку
    next_row = ws.max_row + 1
    
    ws.cell(row=next_row, column=1, value=date_str)
    ws.cell(row=next_row, column=2, value=time_str)
    ws.cell(row=next_row, column=3, value=period)
    ws.cell(row=next_row, column=4, value=systolic)
    ws.cell(row=next_row, column=5, value=diastolic)
    ws.cell(row=next_row, column=6, value=pulse if pulse else "")
    ws.cell(row=next_row, column=7, value=comment if comment else "")
    
    wb.save(EXCEL_FILE)

def get_today_report():
    if not os.path.exists(EXCEL_FILE):
        return None
    
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Давление"]
    today = datetime.now(MSK_PLUS_1).strftime("%d-%m-%Y")
    
    report = f"📊 Отчет за {today}\n\n"
    has_data = False
    
    for row in range(2, ws.max_row + 1):
        date = ws.cell(row=row, column=1).value
        if date == today:
            has_data = True
            time_val = ws.cell(row=row, column=2).value or "-"
            period = ws.cell(row=row, column=3).value or "-"
            systolic = ws.cell(row=row, column=4).value or "-"
            diastolic = ws.cell(row=row, column=5).value or "-"
            pulse = ws.cell(row=row, column=6).value or "-"
            comment = ws.cell(row=row, column=7).value or ""
            
            period_emoji = {"Утро": "🌅", "День": "☀️", "Вечер": "🌙"}
            emoji = period_emoji.get(period, "")
            
            report += f"{emoji} {period} {time_val}: {systolic}/{diastolic}"
            if pulse != "-":
                report += f", пульс {pulse}"
            if comment:
                report += f"\n   📝 {comment}"
            report += "\n\n"
    
    if not has_data:
        return f"📊 Отчет за {today}\n\nНет данных. Добавьте измерения."
    
    report += f"Для полного журнала нажмите /table"
    return report

# ==================== АДМИН КОМАНДЫ ====================
async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    await update.message.reply_text(
        "👑 Админ панель\n\n"
        "Доступные команды:\n"
        "/users - список пользователей\n"
        "/users_excel - выгрузить пользователей в Excel\n"
        "/grant username дни - выдать доступ (5,7,30)"
    )

async def admin_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    users = get_all_users()
    if not users:
        await update.message.reply_text("Нет пользователей.")
        return
    
    text = "👥 Список пользователей:\n\n"
    for uid, data in users.items():
        text += f"🆔 ID: {uid}\n"
        text += f"📝 Username: {data.get('username', '-')}\n"
        text += f"📅 Подключен: {data.get('joined', '-')}\n"
        text += f"🔒 Статус: {data.get('status', '-')}\n"
        text += f"📊 Дней: {data.get('days_count', '-')}\n"
        if data.get('access_until'):
            text += f"⏰ Доступ до: {data['access_until']}\n"
        text += "-" * 30 + "\n"
    
    await update.message.reply_text(text)

async def admin_users_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    users = get_all_users()
    if not users:
        await update.message.reply_text("Нет пользователей.")
        return
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Пользователи"
    
    headers = ["ID", "Username", "Дата подключения", "Статус", "Дней", "Доступ до"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
        ws.cell(row=1, column=col).font = Font(bold=True)
    
    row = 2
    for uid, data in users.items():
        ws.cell(row=row, column=1, value=uid)
        ws.cell(row=row, column=2, value=data.get('username', '-'))
        ws.cell(row=row, column=3, value=data.get('joined', '-'))
        ws.cell(row=row, column=4, value=data.get('status', '-'))
        ws.cell(row=row, column=5, value=data.get('days_count', '-'))
        ws.cell(row=row, column=6, value=data.get('access_until', '-'))
        row += 1
    
    filename = f"users_{datetime.now(MSK_PLUS_1).strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(filename)
    
    with open(filename, 'rb') as f:
        await update.message.reply_document(f, filename=filename)
    
    os.remove(filename)

async def admin_grant(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    args = context.args
    if len(args) < 2:
        await update.message.reply_text(
            "📝 Используйте: /grant username дни\n"
            "Дни: 5, 7, 30\n"
            "Пример: /grant john 7"
        )
        return
    
    username = args[0]
    days = int(args[1])
    
    if days not in [5, 7, 30]:
        await update.message.reply_text("❌ Доступны дни: 5, 7, 30")
        return
    
    users = get_all_users()
    found = None
    for uid, data in users.items():
        if data.get('username', '').lower() == username.lower():
            found = uid
            break
    
    if found:
        grant_access(found, days)
        await update.message.reply_text(f"✅ Пользователю {username} выдан доступ на {days} дней.")
    else:
        await update.message.reply_text(f"❌ Пользователь {username} не найден.")

# ==================== ОСНОВНЫЕ КОМАНДЫ ====================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    username = update.effective_user.username or update.effective_user.first_name
    
    is_new = add_user(user_id, username)
    
    # Если новый пользователь, уведомляем админа
    if is_new and not is_admin(user_id):
        now = datetime.now(MSK_PLUS_1)
        await context.bot.send_message(
            chat_id=ADMIN_ID,
            text=f"🆕 Новый пользователь!\n\n"
                 f"👤 Username: @{username}\n"
                 f"🆔 ID: {user_id}\n"
                 f"📅 Дата: {now.strftime('%d-%m-%Y %H:%M:%S')}"
        )
    
    # Проверяем 3-й день
    await check_and_send_3day_reminder(user_id, context.application)
    
    await update.message.reply_text(
        "📊 Я помогу вести журнал вашего артериального давления.\n\n"
        "Напишите мне свои показатели в формате:\n"
        "120 80 68 - давление и пульс\n"
        "120 80 - только давление\n"
        "120/80 - через слеш\n\n"
        "Можно добавить комментарий:\n"
        "120 80 68 выпил таблетку\n\n"
        "Бот сам определит время суток (Утро, День, Вечер)\n"
        "Каждый замер сохраняется отдельной строкой"
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📖 Помощь\n\n"
        "Как пользоваться:\n"
        "1. Отправьте показания давления в любом формате\n"
        "2. Бот сам определит время суток (Утро, День, Вечер)\n"
        "3. Каждый замер сохраняется отдельной строкой\n\n"
        "Форматы ввода:\n"
        "• 130 85 - давление\n"
        "• 130 85 72 - давление и пульс\n"
        "• 130/85 - через слеш\n"
        "• 130 85 72 выпил таблетку - с комментарием\n\n"
        "Команды:\n"
        "/table - получить Excel файл\n"
        "/report - отчет за сегодня\n\n"
        f"По вопросам и предложениям пишите администратору @{ADMIN_USERNAME}"
    )

async def report_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    report = get_today_report()
    await update.message.reply_text(report)

async def table_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
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
    user_id = update.effective_user.id
    username = update.effective_user.username or update.effective_user.first_name
    
    # Добавляем пользователя если его нет
    add_user(user_id, username)
    
    # Проверяем доступ
    if not check_access(user_id):
        await update.message.reply_text(
            "⛔ Доступ временно приостановлен.\n"
            f"Свяжитесь с администратором @{ADMIN_USERNAME}"
        )
        return
    
    # Проверяем 3-й день
    await check_and_send_3day_reminder(user_id, context.application)
    
    # Обновляем количество дней
    update_user_days(user_id)
    
    text = update.message.text.strip()
    
    # Парсим числа
    numbers = re.findall(r'\d+', text)
    
    systolic = None
    diastolic = None
    pulse = None
    
    # Ищем давление через слеш
    slash_match = re.search(r'(\d{2,3})/(\d{2,3})', text)
    if slash_match:
        systolic = int(slash_match.group(1))
        diastolic = int(slash_match.group(2))
    elif len(numbers) >= 2:
        systolic = int(numbers[0])
        diastolic = int(numbers[1])
    
    if not systolic or not diastolic:
        await update.message.reply_text(
            "❌ Не понял. Примеры:\n"
            "120 80\n"
            "120 80 68\n"
            "120/80\n"
            "120 80 68 выпил таблетку"
        )
        return
    
    # Ищем пульс (третье число)
    if len(numbers) >= 3:
        pulse = int(numbers[2])
    
    # Извлекаем комментарий (всё после первых трёх чисел или после двух)
    comment = ""
    # Удаляем все числа из текста
    comment_text = re.sub(r'\d+', '', text)
    # Удаляем пробелы и слеши в начале
    comment_text = re.sub(r'^[\s/]+', '', comment_text)
    # Если остался текст - это комментарий
    if comment_text.strip():
        comment = comment_text.strip()
    
    # Определяем период по времени
    period = get_period_by_time()
    
    # Сохраняем в Excel
    save_to_excel(user_id, period, systolic, diastolic, pulse, comment)
    
    period_emoji = {"Утро": "🌅", "День": "☀️", "Вечер": "🌙"}
    now = datetime.now(MSK_PLUS_1)
    
    response = f"✅ Записано! {period_emoji.get(period, '')} {period}: {systolic}/{diastolic}"
    if pulse:
        response += f", пульс {pulse}"
    if comment:
        response += f"\n📝 Комментарий: {comment}"
    response += f"\n📅 {now.strftime('%d-%m-%Y %H:%M:%S')}"
    
    await update.message.reply_text(response)

async def send_scheduled_reminder(context: ContextTypes.DEFAULT_TYPE):
    users = get_all_users()
    now_time = datetime.now(MSK_PLUS_1)
    
    # Отправляем напоминания только в 8:00, 14:00, 20:00
    if now_time.hour not in [8, 14, 20]:
        return
    
    for uid, data in users.items():
        if check_access(int(uid)):
            try:
                await context.bot.send_message(
                    chat_id=int(uid),
                    text="🔔 Напоминание: пора измерить давление!\n\nПросто отправьте мне показания"
                )
            except:
                pass

async def set_commands(app):
    # Все команды для админа
    admin_commands = [
        BotCommand("start", "Главное меню"),
        BotCommand("table", "Получить Excel журнал"),
        BotCommand("report", "Отчет за сегодня"),
        BotCommand("help", "Помощь"),
        BotCommand("admin", "Админ панель"),
        BotCommand("users", "Список пользователей"),
        BotCommand("users_excel", "Выгрузить пользователей в Excel"),
        BotCommand("grant", "Выдать доступ (username дни)"),
    ]
    
    # Обычные команды для всех пользователей
    default_commands = [
        BotCommand("start", "Главное меню"),
        BotCommand("table", "Получить Excel журнал"),
        BotCommand("report", "Отчет за сегодня"),
        BotCommand("help", "Помощь"),
    ]
    
    # Сначала устанавливаем команды по умолчанию (для всех)
    await app.bot.set_my_commands(default_commands)
    
    # Затем переопределяем команды для админа (полностью заменяем)
    await app.bot.set_my_commands(admin_commands, scope=BotCommandScopeChat(chat_id=ADMIN_ID))

def main():
    init_excel()
    
    app = Application.builder().token(TOKEN).build()
    
    # Основные команды
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("report", report_command))
    app.add_handler(CommandHandler("table", table_command))
    
    # Админ команды
    app.add_handler(CommandHandler("admin", admin_panel))
    app.add_handler(CommandHandler("users", admin_users))
    app.add_handler(CommandHandler("users_excel", admin_users_excel))
    app.add_handler(CommandHandler("grant", admin_grant))
    
    # Обработчик текстовых сообщений (показаний)
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pressure))
    
    # Устанавливаем команды для меню
    asyncio.get_event_loop().run_until_complete(set_commands(app))
    
    # Напоминания
    job_queue = app.job_queue
    if job_queue:
        job_queue.run_daily(send_scheduled_reminder, time(5, 0))   # 8:00 МСК+1
        job_queue.run_daily(send_scheduled_reminder, time(11, 0))  # 14:00 МСК+1
        job_queue.run_daily(send_scheduled_reminder, time(17, 0))  # 20:00 МСК+1
        print("Напоминания: 8:00, 14:00, 20:00 (МСК+1)")
    
    print("Бот запущен")
    app.run_polling()

if __name__ == "__main__":
    main()