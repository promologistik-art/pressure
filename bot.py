import os
import re
import json
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
USERS_DB = os.getenv("USERS_DB", "users.json")

MSK = pytz.timezone('Europe/Moscow')

user_period = {}

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
        users[str(user_id)] = {
            "username": username,
            "joined": datetime.now(MSK).strftime("%d-%m-%Y %H:%M:%S"),
            "status": "trial",
            "trial_until": (datetime.now(MSK) + timedelta(days=3)).strftime("%d-%m-%Y"),
            "access_until": None
        }
        save_users(users)
        return True
    return False

def check_access(user_id):
    users = load_users()
    user = users.get(str(user_id))
    if not user:
        return False
    
    if user["status"] == "admin":
        return True
    
    if user["status"] == "trial":
        trial_until = datetime.strptime(user["trial_until"], "%d-%m-%Y")
        if datetime.now(MSK) <= trial_until:
            return True
        else:
            user["status"] = "blocked"
            save_users(users)
            return False
    
    if user["status"] == "access":
        access_until = datetime.strptime(user["access_until"], "%d-%m-%Y")
        if datetime.now(MSK) <= access_until:
            return True
        else:
            user["status"] = "blocked"
            save_users(users)
            return False
    
    return False

def grant_access(user_id, days):
    users = load_users()
    user = users.get(str(user_id))
    if user:
        user["status"] = "access"
        user["access_until"] = (datetime.now(MSK) + timedelta(days=days)).strftime("%d-%m-%Y")
        save_users(users)
        return True
    return False

def is_admin(user_id):
    return user_id == ADMIN_ID

def get_all_users():
    users = load_users()
    return users

# ==================== РАБОТА С EXCEL ====================
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        default_sheet = wb.active
        wb.remove(default_sheet)
        
        ws = wb.create_sheet("Давление")
        
        # Строка 1: только B1=утро, F1=обед, J1=вечер
        ws.cell(row=1, column=2, value="утро")  # B1
        ws.cell(row=1, column=6, value="обед")  # F1
        ws.cell(row=1, column=10, value="вечер") # J1
        
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
        
        # Ширина колонок
        column_widths = {
            'A': 12, 'B': 10, 'C': 14, 'D': 14, 'E': 8,
            'F': 10, 'G': 14, 'H': 14, 'I': 8,
            'J': 10, 'K': 14, 'L': 14, 'M': 8
        }
        
        for col_letter, width in column_widths.items():
            ws.column_dimensions[col_letter].width = width
        
        ws.row_dimensions[1].height = 25
        ws.row_dimensions[2].height = 20
        
        wb.save(EXCEL_FILE)
        print(f"Создан файл {EXCEL_FILE}")

def save_to_excel(user_id, period, systolic, diastolic, pulse=None):
    now = datetime.now(MSK)
    date_str = now.strftime("%d-%m-%Y")
    time_str = now.strftime("%H:%M:%S")
    
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
    today = datetime.now(MSK).strftime("%d-%m-%Y")
    
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
            report += f"☀️ Обед:\n"
            report += f"   Время: {afternoon_time}\n"
            report += f"   Давление: {afternoon_sys}/{afternoon_dia}\n"
            report += f"   Пульс: {afternoon_pulse}\n\n"
            report += f"🌙 Вечер:\n"
            report += f"   Время: {evening_time}\n"
            report += f"   Давление: {evening_sys}/{evening_dia}\n"
            report += f"   Пульс: {evening_pulse}\n\n"
            report += f"Для полного журнала нажмите /table"
            
            return report
    
    return f"📊 Отчет за {today}\n\nНет данных. Добавьте измерения через /morning, /afternoon, /evening"

# ==================== АДМИН КОМАНДЫ ====================
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
        if data.get('access_until'):
            text += f"Доступ до: {data['access_until']}\n"
        if data.get('trial_until'):
            text += f"Триал до: {data['trial_until']}\n"
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
    
    headers = ["ID", "Username", "Дата подключения", "Статус", "Триал до", "Доступ до"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
        ws.cell(row=1, column=col).font = Font(bold=True)
    
    row = 2
    for uid, data in users.items():
        ws.cell(row=row, column=1, value=uid)
        ws.cell(row=row, column=2, value=data.get('username', '-'))
        ws.cell(row=row, column=3, value=data.get('joined', '-'))
        ws.cell(row=row, column=4, value=data.get('status', '-'))
        ws.cell(row=row, column=5, value=data.get('trial_until', '-'))
        ws.cell(row=row, column=6, value=data.get('access_until', '-'))
        row += 1
    
    filename = f"users_{datetime.now(MSK).strftime('%Y%m%d_%H%M%S')}.xlsx"
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
    
    if is_admin(user_id):
        await update.message.reply_text(
            "👑 Админ панель\n\n"
            "Доступные команды:\n"
            "/users - список пользователей\n"
            "/users_excel - выгрузить пользователей в Excel\n"
            "/grant username дни - выдать доступ (5,7,30)\n\n"
            "Основные команды:\n"
            "/morning - утро\n/afternoon - обед\n/evening - вечер\n"
            "/table - Excel журнал\n/report - отчет за сегодня\n/remind - напоминание\n/help - помощь"
        )
        return
    
    new_user = add_user(user_id, username)
    if new_user:
        await update.message.reply_text(
            "Добро пожаловать!\n\n"
            "Вам доступен 3-дневный пробный период.\n\n"
            "Команды:\n"
            "/morning - утреннее давление\n"
            "/afternoon - обеденное давление\n"
            "/evening - вечернее давление\n"
            "/table - получить Excel журнал\n"
            "/report - отчет за сегодня\n"
            "/remind - напоминание\n"
            "/help - помощь\n\n"
            f"Пробный период: 3 дня (до {(datetime.now(MSK) + timedelta(days=3)).strftime('%d-%m-%Y')})"
        )
        
        await context.bot.send_message(
            chat_id=ADMIN_ID,
            text=f"Новый пользователь!\n\n"
                 f"Username: @{username}\n"
                 f"ID: {user_id}\n"
                 f"Дата: {datetime.now(MSK).strftime('%d-%m-%Y %H:%M:%S')}\n\n"
                 f"Для выдачи доступа: /grant {username} 7"
        )
    else:
        if check_access(user_id):
            await update.message.reply_text(
                "Бот контроля давления\n\n"
                "Команды:\n"
                "/morning - утро\n/afternoon - обед\n/evening - вечер\n"
                "/table - Excel журнал\n/report - отчет за сегодня\n/remind - напоминание\n/help - помощь"
            )
        else:
            await update.message.reply_text(
                "Доступ заблокирован\n\n"
                "Ваш пробный период истек. Свяжитесь с администратором для получения доступа."
            )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if not check_access(user_id) and not is_admin(user_id):
        await update.message.reply_text("Доступ запрещён. Свяжитесь с администратором.")
        return
    
    await update.message.reply_text(
        "Инструкция\n\n"
        "1. Запись давления:\n"
        "/morning - утро\n/afternoon - обед\n/evening - вечер\n\n"
        "2. Форматы ввода:\n"
        "• 130 85 - давление\n"
        "• 130 85 72 - давление и пульс\n"
        "• 130/85 - через слеш\n\n"
        "3. Получение данных:\n"
        "/table - Excel файл\n/report - отчет за сегодня\n\n"
        "4. Напоминания:\n"
        "Автоматические в 8:00, 14:00, 20:00 МСК\n"
        "/remind - ручное напоминание"
    )

async def report_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if not check_access(user_id) and not is_admin(user_id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    report = get_today_report()
    await update.message.reply_text(report)

async def table_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if not check_access(user_id) and not is_admin(user_id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    if not os.path.exists(EXCEL_FILE):
        await update.message.reply_text("Журнал пуст.")
        return
    
    with open(EXCEL_FILE, 'rb') as f:
        await update.message.reply_document(
            document=f,
            filename="pressure_journal.xlsx",
            caption="Журнал давления"
        )

async def remind_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if not check_access(user_id) and not is_admin(user_id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    await update.message.reply_text(
        "Напоминание\n\nПора измерить давление!\n\n"
        "/morning - утро\n/afternoon - обед\n/evening - вечер"
    )

async def morning_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if not check_access(user_id) and not is_admin(user_id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    user_period[user_id] = 'утро'
    await update.message.reply_text("Утренний замер\n\nВведите показания в формате:\n120 80 68 или 120/80")

async def afternoon_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if not check_access(user_id) and not is_admin(user_id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    user_period[user_id] = 'обед'
    await update.message.reply_text("Обеденный замер\n\nВведите показания в формате:\n120 80 68 или 120/80")

async def evening_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if not check_access(user_id) and not is_admin(user_id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    user_period[user_id] = 'вечер'
    await update.message.reply_text("Вечерний замер\n\nВведите показания в формате:\n120 80 68 или 120/80")

async def handle_pressure(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if not check_access(user_id) and not is_admin(user_id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    if user_id not in user_period:
        await update.message.reply_text("Сначала выберите /morning, /afternoon или /evening")
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
        await update.message.reply_text("Не понял. Пример: 130 85 или 130/85")
        return
    
    save_to_excel(user_id, period, systolic, diastolic, pulse)
    del user_period[user_id]
    
    period_emoji = {'утро': '🌅', 'обед': '☀️', 'вечер': '🌙'}
    now = datetime.now(MSK)
    
    await update.message.reply_text(
        f"Записано! {period_emoji[period]} {period}: {systolic}/{diastolic}" +
        (f", пульс {pulse}" if pulse else "") +
        f"\nДата: {now.strftime('%d-%m-%Y %H:%M:%S')}"
    )

async def send_scheduled_reminder(context: ContextTypes.DEFAULT_TYPE):
    users = get_all_users()
    for uid, data in users.items():
        if check_access(int(uid)):
            try:
                await context.bot.send_message(
                    chat_id=int(uid),
                    text="Напоминание: пора измерить давление!\n/morning - утро\n/afternoon - обед\n/evening - вечер"
                )
            except:
                pass

async def set_commands(app):
    commands = [
        BotCommand("start", "Главное меню"),
        BotCommand("morning", "Утреннее давление"),
        BotCommand("afternoon", "Обеденное давление"),
        BotCommand("evening", "Вечернее давление"),
        BotCommand("table", "Получить Excel"),
        BotCommand("report", "Отчет за сегодня"),
        BotCommand("remind", "Напоминание"),
        BotCommand("help", "Помощь"),
    ]
    await app.bot.set_my_commands(commands)

def main():
    init_excel()
    
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    
    app = Application.builder().token(TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("report", report_command))
    app.add_handler(CommandHandler("table", table_command))
    app.add_handler(CommandHandler("remind", remind_command))
    app.add_handler(CommandHandler("morning", morning_command))
    app.add_handler(CommandHandler("afternoon", afternoon_command))
    app.add_handler(CommandHandler("evening", evening_command))
    
    app.add_handler(CommandHandler("users", admin_users))
    app.add_handler(CommandHandler("users_excel", admin_users_excel))
    app.add_handler(CommandHandler("grant", admin_grant))
    
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pressure))
    
    loop.run_until_complete(set_commands(app))
    
    job_queue = app.job_queue
    if job_queue:
        job_queue.run_daily(send_scheduled_reminder, time(5, 0))
        job_queue.run_daily(send_scheduled_reminder, time(11, 0))
        job_queue.run_daily(send_scheduled_reminder, time(17, 0))
        print("Напоминания: 8:00, 14:00, 20:00 МСК")
    
    print("Бот запущен")
    app.run_polling()

if __name__ == "__main__":
    main()