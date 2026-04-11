import os
import re
import json
import asyncio
from datetime import datetime, timedelta, time
import pytz
from dotenv import load_dotenv
from telegram import Update, BotCommand
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from telegram import Update, BotCommand, BotCommandScopeChat
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

# ==================== РАБОТА С EXCEL ====================
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        default_sheet = wb.active
        wb.remove(default_sheet)
        
        ws = wb.create_sheet("Давление")
        
        # Строка 1: B1=Утро, F1=День, J1=Вечер
        ws.cell(row=1, column=2, value="Утро")
        ws.cell(row=1, column=6, value="День")
        ws.cell(row=1, column=10, value="Вечер")
        
        for col in [2, 6, 10]:
            ws.cell(row=1, column=col).font = Font(bold=True)
            ws.cell(row=1, column=col).alignment = Alignment(horizontal='center', vertical='center')
        
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
        
        # Строка 1, ячейка O1 - пояснение
        explanation = ("Систолическое («верхнее») давление показывает силу давления крови на стенки артерий при сокращении сердца, "
                       "а диастолическое («нижнее») — давление в сосудах во время его расслабления. "
                       "Норма составляет около 120/80 мм рт. ст. Верхнее число отражает работу сердца, а нижнее — тонус сосудов")
        
        ws.cell(row=1, column=15, value=explanation)
        ws.cell(row=1, column=15).alignment = Alignment(wrap_text=True, vertical='top')
        ws.row_dimensions[1].height = 60
        
        # Ширина колонок
        column_widths = {
            'A': 12,   # Дата
            'B': 10,   # Время Утро
            'C': 15,   # Систолическое Утро
            'D': 16,   # Диастолическое Утро
            'E': 7,    # Пульс Утро
            'F': 10,   # Время День
            'G': 15,   # Систолическое День
            'H': 16,   # Диастолическое День
            'I': 7,    # Пульс День
            'J': 10,   # Время Вечер
            'K': 15,   # Систолическое Вечер
            'L': 16,   # Диастолическое Вечер
            'M': 7,    # Пульс Вечер
            'O': 50    # Пояснение
        }
        
        for col_letter, width in column_widths.items():
            ws.column_dimensions[col_letter].width = width
        
        ws.row_dimensions[2].height = 20
        
        wb.save(EXCEL_FILE)
        print(f"Создан файл {EXCEL_FILE}")

def save_to_excel(user_id, period, systolic, diastolic, pulse=None):
    now = datetime.now(MSK_PLUS_1)
    
    # Если время с 00:00 до 5:59 - относим к предыдущему дню
    if now.hour < 6:
        date_str = (now - timedelta(days=1)).strftime("%d-%m-%Y")
    else:
        date_str = now.strftime("%d-%m-%Y")
    
    time_str = now.strftime("%H:%M:%S")
    
    period_cols = {
        "Утро":   {'time_col': 2, 'systolic_col': 3, 'diastolic_col': 4, 'pulse_col': 5},
        "День":   {'time_col': 6, 'systolic_col': 7, 'diastolic_col': 8, 'pulse_col': 9},
        "Вечер":  {'time_col': 10, 'systolic_col': 11, 'diastolic_col': 12, 'pulse_col': 13}
    }
    
    cols = period_cols[period]
    
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb["Давление"]
    else:
        init_excel()
        wb = load_workbook(EXCEL_FILE)
        ws = wb["Давление"]
    
    target_row = None
    for row in range(3, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == date_str:
            target_row = row
            break
    
    if target_row is None:
        target_row = ws.max_row + 1
        ws.cell(row=target_row, column=1, value=date_str)
    
    ws.cell(row=target_row, column=cols['time_col'], value=time_str)
    ws.cell(row=target_row, column=cols['systolic_col'], value=systolic)
    ws.cell(row=target_row, column=cols['diastolic_col'], value=diastolic)
    if pulse:
        ws.cell(row=target_row, column=cols['pulse_col'], value=pulse)
    
    wb.save(EXCEL_FILE)

def get_today_report():
    if not os.path.exists(EXCEL_FILE):
        return None
    
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Давление"]
    today = datetime.now(MSK_PLUS_1).strftime("%d-%m-%Y")
    
    for row in range(3, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == today:
            morning_time = ws.cell(row=row, column=2).value or "-"
            morning_sys = ws.cell(row=row, column=3).value or "-"
            morning_dia = ws.cell(row=row, column=4).value or "-"
            morning_pulse = ws.cell(row=row, column=5).value or "-"
            
            afternoon_time = ws.cell(row=row, column=6).value or "-"
            afternoon_sys = ws.cell(row=row, column=7).value or "-"
            afternoon_dia = ws.cell(row=row, column=8).value or "-"
            afternoon_pulse = ws.cell(row=row, column=9).value or "-"
            
            evening_time = ws.cell(row=row, column=10).value or "-"
            evening_sys = ws.cell(row=row, column=11).value or "-"
            evening_dia = ws.cell(row=row, column=12).value or "-"
            evening_pulse = ws.cell(row=row, column=13).value or "-"
            
            report = f"📊 Отчет за {today}\n\n"
            report += f"🌅 Утро:\n"
            report += f"   Время: {morning_time}\n"
            report += f"   Давление: {morning_sys}/{morning_dia}\n"
            report += f"   Пульс: {morning_pulse}\n\n"
            report += f"☀️ День:\n"
            report += f"   Время: {afternoon_time}\n"
            report += f"   Давление: {afternoon_sys}/{afternoon_dia}\n"
            report += f"   Пульс: {afternoon_pulse}\n\n"
            report += f"🌙 Вечер:\n"
            report += f"   Время: {evening_time}\n"
            report += f"   Давление: {evening_sys}/{evening_dia}\n"
            report += f"   Пульс: {evening_pulse}\n\n"
            report += f"Для полного журнала нажмите /table"
            
            return report
    
    return f"📊 Отчет за {today}\n\nНет данных. Добавьте измерения."

# ==================== АДМИН КОМАНДЫ ====================
async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    await update.message.reply_text(
        "👑 Админ панель\n\n"
        "Команды:\n"
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
    
    text = "Список пользователей:\n\n"
    for uid, data in users.items():
        text += f"ID: {uid}\n"
        text += f"Username: {data.get('username', '-')}\n"
        text += f"Подключен: {data.get('joined', '-')}\n"
        text += f"Статус: {data.get('status', '-')}\n"
        text += f"Дней: {data.get('days_count', '-')}\n"
        if data.get('access_until'):
            text += f"Доступ до: {data['access_until']}\n"
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
            "Используйте: /grant username дни\n"
            "Дни: 5, 7, 30\n"
            "Пример: /grant john 7"
        )
        return
    
    username = args[0]
    days = int(args[1])
    
    if days not in [5, 7, 30]:
        await update.message.reply_text("Доступны дни: 5, 7, 30")
        return
    
    users = get_all_users()
    found = None
    for uid, data in users.items():
        if data.get('username', '').lower() == username.lower():
            found = uid
            break
    
    if found:
        grant_access(found, days)
        await update.message.reply_text(f"Пользователю {username} выдан доступ на {days} дней.")
    else:
        await update.message.reply_text(f"Пользователь {username} не найден.")

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
                 f"Username: @{username}\n"
                 f"ID: {user_id}\n"
                 f"Дата: {now.strftime('%d-%m-%Y %H:%M:%S')}"
        )
    
    # Проверяем 3-й день
    await check_and_send_3day_reminder(user_id, context.application)
    
    await update.message.reply_text(
        "Я помогу вести журнал вашего артериального давления. Напишите мне свои показатели давления и пульса в формате 120 80 68, я сохраню и буду вести ваш журнал.\n\n"
        "Как пользоваться:\n"
        "1. Отправьте показания давления в любом формате\n"
        "2. Бот сам определит время суток (Утро 6-12, День 12-18, Вечер 18-6)\n"
        "3. Данные сохранятся в базе, доступной только Вам и могут быть выгружены в Excel\n\n"
        
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📖 Помощь\n\n"
        "Как пользоваться:\n"
        "1. Отправьте показания давления в любом формате\n"
        "2. Бот сам определит время суток (Утро 6-12, День 12-18, Вечер 18-6)\n"
        "3. Данные сохранятся в Excel\n\n"
        "Форматы ввода:\n"
        "• 130 85 - давление\n"
        "• 130 85 72 - давление и пульс\n"
        "• 130/85 - через слеш\n\n"
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
            caption="Журнал давления"
        )

async def handle_pressure(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    username = update.effective_user.username or update.effective_user.first_name
    
    # Добавляем пользователя если его нет
    add_user(user_id, username)
    
    # Проверяем доступ
    if not check_access(user_id):
        await update.message.reply_text(
            "Доступ временно приостановлен.\n"
            f"Свяжитесь с администратором @{ADMIN_USERNAME}"
        )
        return
    
    # Проверяем 3-й день
    await check_and_send_3day_reminder(user_id, context.application)
    
    # Обновляем количество дней
    update_user_days(user_id)
    
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
            "Не понял. Примеры:\n"
            "120 80\n"
            "120 80 68\n"
            "120/80"
        )
        return
    
    # Определяем период по времени
    period = get_period_by_time()
    
    # Сохраняем в Excel
    save_to_excel(user_id, period, systolic, diastolic, pulse)
    
    period_emoji = {"Утро": "🌅", "День": "☀️", "Вечер": "🌙"}
    now = datetime.now(MSK_PLUS_1)
    
    await update.message.reply_text(
        f"✅ Записано! {period_emoji.get(period, '')} {period}: {systolic}/{diastolic}" +
        (f", пульс {pulse}" if pulse else "") +
        f"\n📅 {now.strftime('%d-%m-%Y %H:%M:%S')}"
    )

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
    # Команды для всех пользователей
    commands = [
        BotCommand("start", "Главное меню"),
        BotCommand("table", "Получить Excel журнал"),
        BotCommand("report", "Отчет за сегодня"),
        BotCommand("help", "Помощь"),
    ]
    await app.bot.set_my_commands(commands)
    
    # Дополнительные команды только для админа
    admin_commands = [
        BotCommand("admin", "Админ панель"),
    ]
    await app.bot.set_my_commands(admin_commands, scope=telegram.BotCommandScopeChat(chat_id=ADMIN_ID))

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