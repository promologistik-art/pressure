import os
import re
import json
import asyncio
import shutil
import zipfile
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
EXCEL_FILE = os.getenv("EXCEL_FILE", "medical_journal.xlsx")
USERS_DB = os.getenv("USERS_DB", "users.json")

# Время МСК+1 (UTC+4)
MSK_PLUS_1 = pytz.timezone('Europe/Samara')

# Типы замеров глюкозы
GLUCOSE_TYPES = ["натощак", "через 2 часа после еды", "перед едой", "перед сном", "ночью"]

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
            "access_until": None,
            "payment_info": None,
            "payment_date": None
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
                    text="Вы уже 3 дня пользуетесь ботом. Если хотите продолжить, есть предложения или замечания, свяжитесь с админом."
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

# ==================== РАБОТА С EXCEL (ОДИН ФАЙЛ, ДВА ЛИСТА) ====================
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        
        default_sheet = wb.active
        wb.remove(default_sheet)
        
        # Лист 1: Давление
        ws_pressure = wb.create_sheet("Давление")
        headers_pressure = ['Дата', 'Время', 'Период', 'Верхнее', 'Нижнее', 'Пульс', 'Комментарий']
        for col, header in enumerate(headers_pressure, 1):
            cell = ws_pressure.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        col_widths_pressure = {'A': 12, 'B': 10, 'C': 8, 'D': 10, 'E': 10, 'F': 8, 'G': 40}
        for col_letter, width in col_widths_pressure.items():
            ws_pressure.column_dimensions[col_letter].width = width
        ws_pressure.row_dimensions[1].height = 20
        
        # Лист 2: Глюкоза
        ws_glucose = wb.create_sheet("Глюкоза")
        headers_glucose = ['Дата', 'Время', 'Период', 'Глюкоза', 'Тип замера', 'Комментарий']
        for col, header in enumerate(headers_glucose, 1):
            cell = ws_glucose.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        col_widths_glucose = {'A': 12, 'B': 10, 'C': 8, 'D': 10, 'E': 25, 'F': 40}
        for col_letter, width in col_widths_glucose.items():
            ws_glucose.column_dimensions[col_letter].width = width
        ws_glucose.row_dimensions[1].height = 20
        
        wb.save(EXCEL_FILE)
        print(f"Создан файл {EXCEL_FILE} с двумя листами")

def save_pressure_to_excel(user_id, period, systolic, diastolic, pulse, comment):
    now = datetime.now(MSK_PLUS_1)
    
    if now.hour < 6:
        date_str = (now - timedelta(days=1)).strftime("%d-%m-%Y")
    else:
        date_str = now.strftime("%d-%m-%Y")
    
    time_str = now.strftime("%H:%M:%S")
    
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        if "Давление" not in wb.sheetnames:
            wb.create_sheet("Давление")
        ws = wb["Давление"]
    else:
        init_excel()
        wb = load_workbook(EXCEL_FILE)
        ws = wb["Давление"]
    
    next_row = ws.max_row + 1
    
    ws.cell(row=next_row, column=1, value=date_str)
    ws.cell(row=next_row, column=2, value=time_str)
    ws.cell(row=next_row, column=3, value=period)
    ws.cell(row=next_row, column=4, value=systolic)
    ws.cell(row=next_row, column=5, value=diastolic)
    ws.cell(row=next_row, column=6, value=pulse if pulse else "")
    ws.cell(row=next_row, column=7, value=comment if comment else "")
    
    wb.save(EXCEL_FILE)

def save_glucose_to_excel(user_id, period, glucose, glucose_type, comment):
    now = datetime.now(MSK_PLUS_1)
    
    if now.hour < 6:
        date_str = (now - timedelta(days=1)).strftime("%d-%m-%Y")
    else:
        date_str = now.strftime("%d-%m-%Y")
    
    time_str = now.strftime("%H:%M:%S")
    
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        if "Глюкоза" not in wb.sheetnames:
            wb.create_sheet("Глюкоза")
        ws = wb["Глюкоза"]
    else:
        init_excel()
        wb = load_workbook(EXCEL_FILE)
        ws = wb["Глюкоза"]
    
    next_row = ws.max_row + 1
    
    ws.cell(row=next_row, column=1, value=date_str)
    ws.cell(row=next_row, column=2, value=time_str)
    ws.cell(row=next_row, column=3, value=period)
    ws.cell(row=next_row, column=4, value=glucose)
    ws.cell(row=next_row, column=5, value=glucose_type)
    ws.cell(row=next_row, column=6, value=comment if comment else "")
    
    wb.save(EXCEL_FILE)

def get_today_pressure_report():
    if not os.path.exists(EXCEL_FILE):
        return None
    
    wb = load_workbook(EXCEL_FILE)
    if "Давление" not in wb.sheetnames:
        return "📊 Отчет по давлению за сегодня\n\nНет данных."
    
    ws = wb["Давление"]
    today = datetime.now(MSK_PLUS_1).strftime("%d-%m-%Y")
    
    report = f"📊 Отчет по давлению за {today}\n\n"
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
        return f"📊 Отчет по давлению за {today}\n\nНет данных."
    
    return report

def get_today_glucose_report():
    if not os.path.exists(EXCEL_FILE):
        return None
    
    wb = load_workbook(EXCEL_FILE)
    if "Глюкоза" not in wb.sheetnames:
        return "📊 Отчет по глюкозе за сегодня\n\nНет данных."
    
    ws = wb["Глюкоза"]
    today = datetime.now(MSK_PLUS_1).strftime("%d-%m-%Y")
    
    report = f"📊 Отчет по глюкозе за {today}\n\n"
    has_data = False
    
    for row in range(2, ws.max_row + 1):
        date = ws.cell(row=row, column=1).value
        if date == today:
            has_data = True
            time_val = ws.cell(row=row, column=2).value or "-"
            period = ws.cell(row=row, column=3).value or "-"
            glucose = ws.cell(row=row, column=4).value or "-"
            glucose_type = ws.cell(row=row, column=5).value or "-"
            comment = ws.cell(row=row, column=6).value or ""
            
            period_emoji = {"Утро": "🌅", "День": "☀️", "Вечер": "🌙"}
            emoji = period_emoji.get(period, "")
            
            report += f"{emoji} {period} {time_val}: глюкоза {glucose}"
            if glucose_type != "-":
                report += f" ({glucose_type})"
            if comment:
                report += f"\n   📝 {comment}"
            report += "\n\n"
    
    if not has_data:
        return f"📊 Отчет по глюкозе за {today}\n\nНет данных."
    
    return report

# ==================== ГЛЮКОЗА - ОПРЕДЕЛЕНИЕ ТИПА ЗАМЕРА ====================
def detect_glucose_type(text):
    text_lower = text.lower()
    if "натощак" in text_lower or "на тощак" in text_lower:
        return "натощак"
    elif "через 2 часа" in text_lower or "после еды" in text_lower:
        return "через 2 часа после еды"
    elif "перед едой" in text_lower:
        return "перед едой"
    elif "перед сном" in text_lower:
        return "перед сном"
    elif "ночью" in text_lower or "ночь" in text_lower:
        return "ночью"
    else:
        now = datetime.now(MSK_PLUS_1)
        if 6 <= now.hour < 12:
            return "натощак"
        return "без указания"

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
        "/grant username дни - выдать доступ (5,7,30)\n"
        "/backup - создать резервную копию (Excel + users.json)\n"
        "/restore - восстановить данные из zip-архива\n"
        "/test_remind - тестовая отправка напоминаний"
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
        if data.get('payment_info'):
            text += f"💳 Оплата: {data['payment_info']}\n"
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
    
    headers = ["ID", "Username", "Дата подключения", "Статус", "Дней", "Доступ до", "Оплата", "Дата оплаты"]
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
        ws.cell(row=row, column=7, value=data.get('payment_info', '-'))
        ws.cell(row=row, column=8, value=data.get('payment_date', '-'))
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

async def backup_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Создаёт резервную копию (Excel + users.json) в zip-архиве"""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    files_to_backup = []
    if os.path.exists(EXCEL_FILE):
        files_to_backup.append(EXCEL_FILE)
    else:
        await update.message.reply_text(f"❌ Файл {EXCEL_FILE} не найден.")
        return
    
    if os.path.exists(USERS_DB):
        files_to_backup.append(USERS_DB)
    else:
        await update.message.reply_text(f"⚠️ Файл {USERS_DB} не найден, будет создан новый при запуске.")
    
    timestamp = datetime.now(MSK_PLUS_1).strftime("%Y%m%d_%H%M%S")
    zip_filename = f"backup_{timestamp}.zip"
    
    try:
        with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file in files_to_backup:
                zipf.write(file, os.path.basename(file))
        
        with open(zip_filename, 'rb') as f:
            await update.message.reply_document(
                document=f,
                filename=zip_filename,
                caption=f"📦 Резервная копия данных от {datetime.now(MSK_PLUS_1).strftime('%d-%m-%Y %H:%M:%S')}\n\n"
                        f"Содержит:\n"
                        f"• {os.path.basename(EXCEL_FILE)} - медицинский журнал\n"
                        f"• {os.path.basename(USERS_DB)} - база пользователей"
            )
        
        os.remove(zip_filename)
        print(f"Создана резервная копия: {zip_filename}")
        
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка при создании резервной копии: {e}")

async def restore_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Восстанавливает данные из zip-архива (только админ)"""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    await update.message.reply_text(
        "📤 Отправьте zip-архив с резервной копией (созданный командой /backup).\n\n"
        "⚠️ ВНИМАНИЕ: текущие данные будут ПЕРЕЗАПИСАНЫ!"
    )
    context.user_data['awaiting_restore'] = True

async def handle_restore_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик загруженного zip-файла для восстановления"""
    if not is_admin(update.effective_user.id):
        return
    
    if not context.user_data.get('awaiting_restore'):
        return
    
    document = update.message.document
    if not document or not document.file_name.endswith('.zip'):
        await update.message.reply_text("❌ Пожалуйста, отправьте zip-архив (созданный командой /backup)")
        return
    
    try:
        file = await context.bot.get_file(document.file_id)
        
        temp_zip = f"temp_restore_{datetime.now(MSK_PLUS_1).strftime('%Y%m%d_%H%M%S')}.zip"
        await file.download_to_drive(temp_zip)
        
        extract_dir = f"extract_{datetime.now(MSK_PLUS_1).strftime('%Y%m%d_%H%M%S')}"
        os.makedirs(extract_dir, exist_ok=True)
        
        with zipfile.ZipFile(temp_zip, 'r') as zipf:
            zipf.extractall(extract_dir)
        
        restored_files = []
        
        # Восстанавливаем Excel файл
        extracted_excel = os.path.join(extract_dir, os.path.basename(EXCEL_FILE))
        if os.path.exists(extracted_excel):
            wb = load_workbook(extracted_excel)
            sheet_names = wb.sheetnames
            
            if "Глюкоза" not in sheet_names:
                ws_glucose = wb.create_sheet("Глюкоза")
                headers_glucose = ['Дата', 'Время', 'Период', 'Глюкоза', 'Тип замера', 'Комментарий']
                for col, header in enumerate(headers_glucose, 1):
                    cell = ws_glucose.cell(row=1, column=col, value=header)
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                wb.save(extracted_excel)
                print("Добавлен отсутствующий лист 'Глюкоза'")
            
            shutil.copy2(extracted_excel, EXCEL_FILE)
            restored_files.append(os.path.basename(EXCEL_FILE))
        
        # Восстанавливаем users.json
        extracted_users = os.path.join(extract_dir, os.path.basename(USERS_DB))
        if os.path.exists(extracted_users):
            shutil.copy2(extracted_users, USERS_DB)
            restored_files.append(os.path.basename(USERS_DB))
        
        os.remove(temp_zip)
        shutil.rmtree(extract_dir)
        
        if restored_files:
            await update.message.reply_text(
                f"✅ Данные восстановлены из файла {document.file_name}\n\n"
                f"Восстановлено: {', '.join(restored_files)}"
            )
        else:
            await update.message.reply_text("❌ Архив не содержит нужных файлов")
        
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка при восстановлении: {e}")
    finally:
        context.user_data['awaiting_restore'] = False

async def test_remind_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Тестовая отправка напоминания всем (только админ)"""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    await update.message.reply_text("🔄 Отправляю тестовые напоминания всем пользователям...")
    
    users = get_all_users()
    sent = 0
    for uid, data in users.items():
        if check_access(int(uid)):
            try:
                await context.bot.send_message(
                    chat_id=int(uid),
                    text="🧪 ТЕСТ: Напоминание работает! Если вы это видите — бот исправен."
                )
                sent += 1
            except:
                pass
    
    await update.message.reply_text(f"✅ Тестовое напоминание отправлено {sent} пользователям")

# ==================== ОСНОВНЫЕ КОМАНДЫ ====================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    username = update.effective_user.username or update.effective_user.first_name
    
    is_new = add_user(user_id, username)
    
    if is_new and not is_admin(user_id):
        now = datetime.now(MSK_PLUS_1)
        await context.bot.send_message(
            chat_id=ADMIN_ID,
            text=f"🆕 Новый пользователь!\n\n"
                 f"👤 Username: @{username}\n"
                 f"🆔 ID: {user_id}\n"
                 f"📅 Дата: {now.strftime('%d-%m-%Y %H:%M:%S')}"
        )
    
    await check_and_send_3day_reminder(user_id, context.application)
    
    await update.message.reply_text(
        "📊 Я помогу вести журнал вашего артериального давления и уровня глюкозы.\n\n"
        "📝 Форматы ввода давления:\n"
        "• 120 80 - давление\n"
        "• 120 80 68 - давление и пульс\n"
        "• 120 80 выпил таблетку - с комментарием\n\n"
        "📝 Форматы ввода глюкозы:\n"
        "• 5.5 - глюкоза (период определится автоматически)\n"
        "• 5.5 натощак - глюкоза с типом замера\n"
        "• 5.5 через 2 часа после еды - глюкоза с типом замера\n\n"
        "🌅 Бот сам определит время суток (Утро, День, Вечер)\n"
        "💾 Давление и глюкоза сохраняются в одном файле на разных листах\n\n"
        "Команды:\n"
        "/table - получить Excel файл (2 листа: давление и глюкоза)\n"
        "/report - отчет по давлению за сегодня\n"
        "/glucose_report - отчет по глюкозе за сегодня\n"
        "/help - помощь"
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "📖 Помощь\n\n"
        "Как пользоваться:\n"
        "1. Отправьте показания давления или глюкозы\n"
        "2. Бот сам определит время суток (Утро, День, Вечер)\n"
        "3. Каждый замер сохраняется отдельной строкой\n\n"
        "Форматы ввода давления:\n"
        "• 130 85 - давление\n"
        "• 130 85 72 - давление и пульс\n"
        "• 130 85 выпил таблетку - с комментарием\n\n"
        "Форматы ввода глюкозы:\n"
        "• 5.5 - глюкоза\n"
        "• 5.5 натощак - глюкоза с типом замера\n"
        "• 5.5 через 2 часа после еды\n\n"
        "Рекомендации по измерению глюкозы:\n"
        "• Утром натощак\n"
        "• Перед каждым приёмом пищи\n"
        "• Через 2 часа после еды\n"
        "• Перед сном\n\n"
        "Целевые показатели:\n"
        "• Натощак: 4.0–7.0 ммоль/л\n"
        "• Через 2 часа после еды: менее 10.0 ммоль/л\n\n"
        "Команды:\n"
        "/table - Excel файл (давление и глюкоза)\n"
        "/report - отчет по давлению за сегодня\n"
        "/glucose_report - отчет по глюкозе за сегодня\n\n"
        f"📢 <a href='https://t.me/+MAuGbcnBQmgxZTIy'>Больше наших ботов в канале</a>"
    )
    
    await update.message.reply_text(text, parse_mode="HTML", disable_web_page_preview=True)

async def report_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    report = get_today_pressure_report()
    await update.message.reply_text(report)

async def glucose_report_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    report = get_today_glucose_report()
    await update.message.reply_text(report)

async def table_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not os.path.exists(EXCEL_FILE):
        await update.message.reply_text("Журнал пуст.")
        return
    
    with open(EXCEL_FILE, 'rb') as f:
        await update.message.reply_document(
            document=f,
            filename="medical_journal.xlsx",
            caption="📊 Медицинский журнал (давление и глюкоза)"
        )

async def handle_pressure_glucose(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    username = update.effective_user.username or update.effective_user.first_name
    
    add_user(user_id, username)
    
    if not check_access(user_id):
        await update.message.reply_text(
            f"⛔ Доступ временно приостановлен.\nСвяжитесь с администратором @{ADMIN_USERNAME}"
        )
        return
    
    await check_and_send_3day_reminder(user_id, context.application)
    update_user_days(user_id)
    
    text = update.message.text.strip()
    
    numbers = re.findall(r'\d+[.,]?\d*', text)
    numbers = [float(n.replace(',', '.')) for n in numbers]
    
    # Проверяем на глюкозу (одно число 1-30)
    if len(numbers) == 1 and 1 <= numbers[0] <= 30:
        glucose = numbers[0]
        glucose_type = detect_glucose_type(text)
        period = get_period_by_time()
        
        comment = re.sub(r'\d+[.,]?\d*', '', text)
        comment = re.sub(r'натощак|через 2 часа после еды|перед едой|перед сном|ночью', '', comment, flags=re.IGNORECASE)
        comment = re.sub(r'[\s/]+', ' ', comment).strip()
        
        save_glucose_to_excel(user_id, period, glucose, glucose_type, comment)
        
        period_emoji = {"Утро": "🌅", "День": "☀️", "Вечер": "🌙"}
        now = datetime.now(MSK_PLUS_1)
        
        response = f"✅ Записано! {period_emoji.get(period, '')} {period}: глюкоза {glucose}"
        if glucose_type != "без указания":
            response += f" ({glucose_type})"
        if comment:
            response += f"\n📝 {comment}"
        response += f"\n📅 {now.strftime('%d-%m-%Y %H:%M:%S')}"
        
        await update.message.reply_text(response)
        return
    
    # Давление
    systolic = None
    diastolic = None
    pulse = None
    
    slash_match = re.search(r'(\d{2,3})/(\d{2,3})', text)
    if slash_match:
        systolic = int(slash_match.group(1))
        diastolic = int(slash_match.group(2))
        numbers = [n for n in numbers if n not in [systolic, diastolic]]
    elif len(numbers) >= 2:
        systolic = int(numbers[0])
        diastolic = int(numbers[1])
        numbers = numbers[2:]
    
    if not systolic or not diastolic:
        await update.message.reply_text(
            "❌ Не понял. Примеры:\n"
            "120 80 - давление\n"
            "120 80 68 - давление и пульс\n"
            "5.5 - глюкоза\n"
            "120 80 выпил таблетку - с комментарием"
        )
        return
    
    for n in numbers:
        if 40 <= n <= 150:
            pulse = int(n)
            break
    
    comment = re.sub(r'\d+[.,]?\d*', '', text)
    comment = re.sub(r'[\s/]+', ' ', comment).strip()
    
    period = get_period_by_time()
    save_pressure_to_excel(user_id, period, systolic, diastolic, pulse, comment)
    
    period_emoji = {"Утро": "🌅", "День": "☀️", "Вечер": "🌙"}
    now = datetime.now(MSK_PLUS_1)
    
    response = f"✅ Записано! {period_emoji.get(period, '')} {period}: {systolic}/{diastolic}"
    if pulse:
        response += f", пульс {pulse}"
    if comment:
        response += f"\n📝 {comment}"
    response += f"\n📅 {now.strftime('%d-%m-%Y %H:%M:%S')}"
    
    await update.message.reply_text(response)

async def send_scheduled_reminder(context: ContextTypes.DEFAULT_TYPE):
    """Отправляет напоминания всем активным пользователям"""
    now_time = datetime.now(MSK_PLUS_1)
    current_hour = now_time.hour
    
    if current_hour not in [8, 14, 20]:
        return
    
    print(f"[{now_time.strftime('%Y-%m-%d %H:%M:%S')}] Запуск напоминаний, час: {current_hour}")
    
    users = get_all_users()
    sent_count = 0
    active_count = 0
    
    for uid, data in users.items():
        if check_access(int(uid)):
            active_count += 1
            try:
                await context.bot.send_message(
                    chat_id=int(uid),
                    text="🔔 Напоминание: пора измерить давление и глюкозу!\n\n"
                         "Просто отправьте мне показания:\n"
                         "• 120 80 - давление\n"
                         "• 120 80 68 - давление и пульс\n"
                         "• 5.5 - глюкоза\n"
                         "• 5.5 натощак - глюкоза с типом замера"
                )
                sent_count += 1
                print(f"  → Напоминание отправлено пользователю {uid}")
            except Exception as e:
                print(f"  ✗ Ошибка отправки пользователю {uid}: {e}")
    
    print(f"Активных пользователей: {active_count}, отправлено напоминаний: {sent_count}")

async def set_commands(app):
    admin_commands = [
        BotCommand("start", "Главное меню"),
        BotCommand("table", "Excel журнал (давление+глюкоза)"),
        BotCommand("report", "Отчет по давлению за сегодня"),
        BotCommand("glucose_report", "Отчет по глюкозе за сегодня"),
        BotCommand("help", "Помощь"),
        BotCommand("admin", "Админ панель"),
        BotCommand("users", "Список пользователей"),
        BotCommand("users_excel", "Выгрузить пользователей в Excel"),
        BotCommand("grant", "Выдать доступ (username дни)"),
        BotCommand("backup", "Резервная копия данных"),
        BotCommand("restore", "Восстановить данные"),
        BotCommand("test_remind", "Тест напоминаний"),
    ]
    
    default_commands = [
        BotCommand("start", "Главное меню"),
        BotCommand("table", "Excel журнал (давление+глюкоза)"),
        BotCommand("report", "Отчет по давлению за сегодня"),
        BotCommand("glucose_report", "Отчет по глюкозе за сегодня"),
        BotCommand("help", "Помощь"),
    ]
    
    await app.bot.set_my_commands(default_commands)
    await app.bot.set_my_commands(admin_commands, scope=BotCommandScopeChat(chat_id=ADMIN_ID))

def main():
    init_excel()
    
    app = Application.builder().token(TOKEN).build()
    
    if app.job_queue is None:
        print("❌ ОШИБКА: JobQueue не создан! Напоминания работать не будут")
    else:
        print("✅ JobQueue создан успешно")
        print(f"   Текущее время сервера: {datetime.now(MSK_PLUS_1).strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"   Часовой пояс: Europe/Samara (МСК+1)")
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("report", report_command))
    app.add_handler(CommandHandler("glucose_report", glucose_report_command))
    app.add_handler(CommandHandler("table", table_command))
    
    app.add_handler(CommandHandler("admin", admin_panel))
    app.add_handler(CommandHandler("users", admin_users))
    app.add_handler(CommandHandler("users_excel", admin_users_excel))
    app.add_handler(CommandHandler("grant", admin_grant))
    app.add_handler(CommandHandler("backup", backup_command))
    app.add_handler(CommandHandler("restore", restore_command))
    app.add_handler(CommandHandler("test_remind", test_remind_all))
    
    app.add_handler(MessageHandler(filters.Document.ALL, handle_restore_file))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pressure_glucose))
    
    asyncio.get_event_loop().run_until_complete(set_commands(app))
    
    job_queue = app.job_queue
    if job_queue:
        job_queue.run_daily(send_scheduled_reminder, time(5, 0))
        job_queue.run_daily(send_scheduled_reminder, time(11, 0))
        job_queue.run_daily(send_scheduled_reminder, time(17, 0))
        print("Напоминания: 8:00, 14:00, 20:00 (МСК+1)")
    else:
        print("ОШИБКА: job_queue не создан! Напоминания работать не будут")
    
    print("Бот запущен")
    app.run_polling()

if __name__ == "__main__":
    main()