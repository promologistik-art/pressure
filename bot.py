import os
import re
import asyncio
import tempfile
import subprocess
from datetime import datetime, timedelta, time
from typing import Dict, List, Optional, Tuple, Any
import pytz
from dotenv import load_dotenv
from telegram import Update, BotCommand, BotCommandScopeChat
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import asyncpg

load_dotenv()

# ==================== КОНФИГУРАЦИЯ ====================
TOKEN = os.getenv("BOT_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID"))
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "silverzen")
DATABASE_URL = os.getenv("DATABASE_URL")

# Время МСК+1 (UTC+4)
MSK_PLUS_1 = pytz.timezone('Europe/Samara')

# Глобальная переменная для пула подключений к БД
db_pool: Optional[asyncpg.Pool] = None

# ==================== ИНИЦИАЛИЗАЦИЯ БД ====================
async def init_db() -> None:
    """Инициализация подключения к БД и создание таблиц"""
    global db_pool
    
    if db_pool is None:
        db_pool = await asyncpg.create_pool(DATABASE_URL, min_size=1, max_size=5)
        print("✅ Подключение к PostgreSQL установлено")
    
    async with db_pool.acquire() as conn:
        # Таблица пользователей
        await conn.execute('''
            CREATE TABLE IF NOT EXISTS users (
                user_id BIGINT PRIMARY KEY,
                username TEXT,
                joined TIMESTAMP,
                status TEXT DEFAULT 'active',
                days_count INT DEFAULT 1,
                last_reminder_sent DATE,
                access_until DATE,
                payment_info TEXT,
                payment_date DATE
            )
        ''')
        
        # Таблица давления
        await conn.execute('''
            CREATE TABLE IF NOT EXISTS pressure (
                id SERIAL PRIMARY KEY,
                user_id BIGINT REFERENCES users(user_id) ON DELETE CASCADE,
                date DATE NOT NULL,
                time TIME NOT NULL,
                period TEXT CHECK (period IN ('Утро', 'День', 'Вечер')),
                systolic INT,
                diastolic INT,
                pulse INT,
                comment TEXT,
                created_at TIMESTAMP DEFAULT NOW()
            )
        ''')
        
        # Таблица глюкозы
        await conn.execute('''
            CREATE TABLE IF NOT EXISTS glucose (
                id SERIAL PRIMARY KEY,
                user_id BIGINT REFERENCES users(user_id) ON DELETE CASCADE,
                date DATE NOT NULL,
                time TIME NOT NULL,
                period TEXT CHECK (period IN ('Утро', 'День', 'Вечер')),
                glucose_value DECIMAL(4,1),
                glucose_type TEXT,
                comment TEXT,
                created_at TIMESTAMP DEFAULT NOW()
            )
        ''')
        
        # Добавляем новые столбцы для инсулина если их нет
        try:
            await conn.execute('ALTER TABLE glucose ADD COLUMN insulin_dose DECIMAL(4,1)')
            print("✅ Добавлен столбец insulin_dose")
        except Exception as e:
            if 'duplicate column' in str(e).lower() or 'already exists' in str(e).lower():
                print("ℹ️ Столбец insulin_dose уже существует")
            else:
                print(f"⚠️ Ошибка при добавлении insulin_dose: {e}")
        
        try:
            await conn.execute('ALTER TABLE glucose ADD COLUMN insulin_recommendation TEXT')
            print("✅ Добавлен столбец insulin_recommendation")
        except Exception as e:
            if 'duplicate column' in str(e).lower() or 'already exists' in str(e).lower():
                print("ℹ️ Столбец insulin_recommendation уже существует")
            else:
                print(f"⚠️ Ошибка при добавлении insulin_recommendation: {e}")
        
        # Таблица: связь наставник-подопечный
        await conn.execute('''
            CREATE TABLE IF NOT EXISTS mentor_students (
                mentor_id BIGINT REFERENCES users(user_id) ON DELETE CASCADE,
                student_id BIGINT REFERENCES users(user_id) ON DELETE CASCADE,
                created_at TIMESTAMP DEFAULT NOW(),
                PRIMARY KEY (mentor_id, student_id)
            )
        ''')
        
        # Индексы
        await conn.execute('CREATE INDEX IF NOT EXISTS idx_pressure_user_id ON pressure(user_id)')
        await conn.execute('CREATE INDEX IF NOT EXISTS idx_glucose_user_id ON glucose(user_id)')
        await conn.execute('CREATE INDEX IF NOT EXISTS idx_mentor_students_mentor ON mentor_students(mentor_id)')
        await conn.execute('CREATE INDEX IF NOT EXISTS idx_mentor_students_student ON mentor_students(student_id)')
        
        print("✅ Таблицы созданы/проверены")

async def get_db_pool() -> Optional[asyncpg.Pool]:
    """Возвращает пул подключений к БД"""
    return db_pool

# ==================== РАБОТА С ПОЛЬЗОВАТЕЛЯМИ ====================
async def add_user(user_id: int, username: str) -> bool:
    """Добавляет пользователя в БД, возвращает True если новый"""
    users = await get_all_users()
    if str(user_id) not in users:
        now = datetime.now(MSK_PLUS_1).replace(tzinfo=None)
        async with db_pool.acquire() as conn:
            await conn.execute('''
                INSERT INTO users (user_id, username, joined, status, days_count)
                VALUES ($1, $2, $3, 'active', 1)
            ''', user_id, username, now)
        return True
    else:
        if users.get(str(user_id), {}).get("username") != username:
            async with db_pool.acquire() as conn:
                await conn.execute('UPDATE users SET username = $1 WHERE user_id = $2', username, user_id)
    return False

async def get_all_users() -> Dict[str, Dict[str, Any]]:
    """Возвращает всех пользователей в виде словаря"""
    async with db_pool.acquire() as conn:
        rows = await conn.fetch('SELECT * FROM users')
        users = {}
        for row in rows:
            uid = str(row['user_id'])
            users[uid] = {
                "username": row['username'],
                "joined": row['joined'].strftime("%d-%m-%Y %H:%M:%S") if row['joined'] else None,
                "status": row['status'],
                "days_count": row['days_count'],
                "last_reminder_sent": row['last_reminder_sent'].strftime("%d-%m-%Y") if row['last_reminder_sent'] else "",
                "access_until": row['access_until'].strftime("%d-%m-%Y") if row['access_until'] else None,
                "payment_info": row['payment_info'],
                "payment_date": row['payment_date'].strftime("%d-%m-%Y") if row['payment_date'] else None
            }
        return users

async def get_user_by_username(username: str) -> Optional[int]:
    """Возвращает user_id по username"""
    async with db_pool.acquire() as conn:
        row = await conn.fetchrow('SELECT user_id FROM users WHERE username = $1', username)
        if row:
            return row['user_id']
    return None

async def update_user_days(user_id: int) -> int:
    """Обновляет количество дней пользователя"""
    async with db_pool.acquire() as conn:
        row = await conn.fetchrow('SELECT joined FROM users WHERE user_id = $1', user_id)
        if row:
            joined = row['joined']
            days = (datetime.now(MSK_PLUS_1).replace(tzinfo=None) - joined).days + 1
            await conn.execute('UPDATE users SET days_count = $1 WHERE user_id = $2', days, user_id)
            return days
    return 0

async def check_and_send_3day_reminder(user_id: int, app: Application) -> bool:
    """Проверяет и отправляет напоминание на 3-й день"""
    if str(user_id) == str(ADMIN_ID):
        return False
    
    async with db_pool.acquire() as conn:
        row = await conn.fetchrow('SELECT joined, status, last_reminder_sent FROM users WHERE user_id = $1', user_id)
        if row and row['status'] == 'active':
            joined = row['joined']
            days = (datetime.now(MSK_PLUS_1).replace(tzinfo=None) - joined).days + 1
            last_reminder = row['last_reminder_sent']
            
            if days >= 3 and last_reminder != datetime.now(MSK_PLUS_1).date():
                await conn.execute('UPDATE users SET last_reminder_sent = $1 WHERE user_id = $2', 
                                  datetime.now(MSK_PLUS_1).date(), user_id)
                try:
                    await app.bot.send_message(
                        chat_id=int(user_id),
                        text="Вы уже 3 дня пользуетесь ботом. Если хотите продолжить, есть предложения или замечания, свяжитесь с админом."
                    )
                except Exception:
                    pass
                return True
    return False

async def check_access(user_id: int) -> bool:
    """Проверяет доступ пользователя"""
    if str(user_id) == str(ADMIN_ID):
        return True
    
    async with db_pool.acquire() as conn:
        row = await conn.fetchrow('SELECT status, access_until FROM users WHERE user_id = $1', user_id)
        if not row:
            return False
        
        if row['status'] == 'active':
            return True
        
        if row['status'] == 'access':
            if row['access_until']:
                if datetime.now(MSK_PLUS_1).date() <= row['access_until']:
                    return True
                else:
                    await conn.execute('UPDATE users SET status = $1 WHERE user_id = $2', 'blocked', user_id)
                    return False
            return True
    
    return False

async def grant_access(user_id: int, days: int) -> bool:
    """Выдаёт доступ пользователю на указанное количество дней"""
    async with db_pool.acquire() as conn:
        row = await conn.fetchrow('SELECT * FROM users WHERE user_id = $1', user_id)
        if row:
            access_until = (datetime.now(MSK_PLUS_1).replace(tzinfo=None) + timedelta(days=days)).date()
            await conn.execute('UPDATE users SET status = $1, access_until = $2 WHERE user_id = $3', 
                              'access', access_until, user_id)
            return True
    return False

def is_admin(user_id: int) -> bool:
    """Проверяет, является ли пользователь администратором"""
    return user_id == ADMIN_ID

# ==================== РАБОТА С НАСТАВНИКАМИ И ПОДОПЕЧНЫМИ ====================
async def add_student(mentor_id: int, student_username: str) -> Tuple[bool, str]:
    """
    Добавляет подопечного наставнику.
    Возвращает (успех, сообщение)
    """
    student_id = await get_user_by_username(student_username)
    if not student_id:
        return False, f"❌ Пользователь @{student_username} не найден в системе."
    
    if student_id == mentor_id:
        return False, "❌ Нельзя добавить самого себя в подопечные."
    
    async with db_pool.acquire() as conn:
        exists = await conn.fetchrow(
            'SELECT 1 FROM mentor_students WHERE mentor_id = $1 AND student_id = $2',
            mentor_id, student_id
        )
        if exists:
            return False, f"❌ Пользователь @{student_username} уже является вашим подопечным."
        
        await conn.execute(
            'INSERT INTO mentor_students (mentor_id, student_id) VALUES ($1, $2)',
            mentor_id, student_id
        )
    
    return True, f"✅ Пользователь @{student_username} добавлен в ваши подопечные."

async def remove_student(mentor_id: int, student_username: str) -> Tuple[bool, str]:
    """Удаляет подопечного у наставника"""
    student_id = await get_user_by_username(student_username)
    if not student_id:
        return False, f"❌ Пользователь @{student_username} не найден."
    
    async with db_pool.acquire() as conn:
        result = await conn.execute(
            'DELETE FROM mentor_students WHERE mentor_id = $1 AND student_id = $2',
            mentor_id, student_id
        )
        if result == "DELETE 0":
            return False, f"❌ Пользователь @{student_username} не является вашим подопечным."
    
    return True, f"✅ Пользователь @{student_username} удалён из ваших подопечных."

async def get_mentor_students(mentor_id: int) -> List[Dict[str, Any]]:
    """Возвращает список подопечных наставника"""
    async with db_pool.acquire() as conn:
        rows = await conn.fetch('''
            SELECT u.user_id, u.username, u.joined, ms.created_at
            FROM mentor_students ms
            JOIN users u ON ms.student_id = u.user_id
            WHERE ms.mentor_id = $1
            ORDER BY ms.created_at DESC
        ''', mentor_id)
        
        students = []
        for row in rows:
            students.append({
                "user_id": row['user_id'],
                "username": row['username'],
                "joined": row['joined'].strftime("%d-%m-%Y %H:%M:%S") if row['joined'] else None,
                "added_at": row['created_at'].strftime("%d-%m-%Y %H:%M:%S") if row['created_at'] else None
            })
        return students

async def is_mentor_for_student(mentor_id: int, student_id: int) -> bool:
    """Проверяет, является ли пользователь наставником для подопечного"""
    async with db_pool.acquire() as conn:
        row = await conn.fetchrow(
            'SELECT 1 FROM mentor_students WHERE mentor_id = $1 AND student_id = $2',
            mentor_id, student_id
        )
        return row is not None

async def get_student_mentors(student_id: int) -> List[int]:
    """Возвращает список наставников подопечного"""
    async with db_pool.acquire() as conn:
        rows = await conn.fetch(
            'SELECT mentor_id FROM mentor_students WHERE student_id = $1',
            student_id
        )
        return [row['mentor_id'] for row in rows]

async def notify_mentors_about_measurement(student_id: int, measurement_type: str, value: str, app: Application) -> None:
    """Уведомляет всех наставников подопечного о новом замере"""
    mentors = await get_student_mentors(student_id)
    if not mentors:
        return
    
    users = await get_all_users()
    student_info = users.get(str(student_id), {})
    student_username = student_info.get('username', 'Неизвестный пользователь')
    
    now = datetime.now(MSK_PLUS_1).strftime('%d-%m-%Y %H:%M:%S')
    
    for mentor_id in mentors:
        try:
            await app.bot.send_message(
                chat_id=mentor_id,
                text=f"📊 Новый замер от подопечного @{student_username}\n\n"
                     f"📝 Тип: {measurement_type}\n"
                     f"📊 Показания: {value}\n"
                     f"🕐 Время: {now}"
            )
        except Exception as e:
            print(f"❌ Не удалось уведомить наставника {mentor_id}: {e}")

# ==================== РАСЧЁТ ИНСУЛИНА ====================
def calculate_insulin_recommendation(period: str, glucose_value: float, user_dose: Optional[float] = None) -> Dict[str, Any]:
    """
    Рассчитывает рекомендацию по инсулину.
    Возвращает словарь с рекомендацией.
    """
    # Базовые параметры
    if period == "Утро":
        target = 7.0
        base_dose = 10.0
    elif period == "Вечер":
        target = 13.0
        base_dose = 8.0
    else:
        # Для "День" используем вечерние параметры
        target = 13.0
        base_dose = 8.0
    
    # Используем дозу пользователя или базовую
    current_dose = user_dose if user_dose is not None else base_dose
    
    # Вычисляем отклонение
    deviation = glucose_value - target
    adjustment = 0
    recommended_dose = current_dose
    recommendation_text = ""
    
    if deviation >= 1.0:
        adjustment = 2
        recommended_dose = current_dose + 2
        recommendation_text = f"Целевой сахар: {target}\nТекущий сахар: {glucose_value} (выше на {deviation:.1f})\nВаша доза: {current_dose:.0f} ед.\n→ +2 ед.\n✅ Рекомендуемая доза: {recommended_dose:.0f} ед."
    elif deviation <= -1.0:
        adjustment = -2
        recommended_dose = current_dose - 2
        recommendation_text = f"Целевой сахар: {target}\nТекущий сахар: {glucose_value} (ниже на {abs(deviation):.1f})\nВаша доза: {current_dose:.0f} ед.\n→ -2 ед.\n✅ Рекомендуемая доза: {recommended_dose:.0f} ед."
    else:
        recommendation_text = f"Целевой сахар: {target}\nТекущий сахар: {glucose_value} (в пределах нормы)\nВаша доза: {current_dose:.0f} ед.\n→ без изменений\n✅ Рекомендуемая доза: {current_dose:.0f} ед."
    
    return {
        "period": period,
        "target": target,
        "base_dose": base_dose,
        "current_dose": current_dose,
        "glucose_value": glucose_value,
        "deviation": deviation,
        "adjustment": adjustment,
        "recommended_dose": recommended_dose,
        "recommendation_text": recommendation_text,
        "user_provided_dose": user_dose is not None
    }

# ==================== ПОЛУЧЕНИЕ ДАННЫХ ДЛЯ ПРОСМОТРА ====================
async def get_student_full_history(student_id: int) -> str:
    """Возвращает полную историю замеров подопечного в текстовом формате"""
    async with db_pool.acquire() as conn:
        pressure_rows = await conn.fetch('''
            SELECT date, time, period, systolic, diastolic, pulse, comment 
            FROM pressure 
            WHERE user_id = $1 
            ORDER BY date DESC, time DESC
            LIMIT 50
        ''', student_id)
        
        glucose_rows = await conn.fetch('''
            SELECT date, time, period, glucose_value, glucose_type, comment, insulin_dose, insulin_recommendation
            FROM glucose 
            WHERE user_id = $1 
            ORDER BY date DESC, time DESC
            LIMIT 50
        ''', student_id)
    
    if not pressure_rows and not glucose_rows:
        return "📊 У подопечного пока нет записей."
    
    result = "📊 ПОЛНЫЙ ЖУРНАЛ ПОДОПЕЧНОГО\n"
    result += "=" * 40 + "\n\n"
    
    if pressure_rows:
        result += "🩸 ДАВЛЕНИЕ (последние 50 записей):\n"
        result += "-" * 30 + "\n"
        for row in pressure_rows[:20]:
            result += f"📅 {row['date'].strftime('%d-%m-%Y')} {row['time'].strftime('%H:%M:%S')} "
            result += f"({row['period']}): {row['systolic']}/{row['diastolic']}"
            if row['pulse']:
                result += f", пульс {row['pulse']}"
            if row['comment']:
                result += f"\n   📝 {row['comment']}"
            result += "\n"
        if len(pressure_rows) > 20:
            result += f"\n... и ещё {len(pressure_rows) - 20} записей\n"
        result += "\n"
    
    if glucose_rows:
        result += "🩸 ГЛЮКОЗА (последние 50 записей):\n"
        result += "-" * 30 + "\n"
        for row in glucose_rows[:20]:
            result += f"📅 {row['date'].strftime('%d-%m-%Y')} {row['time'].strftime('%H:%M:%S')} "
            result += f"({row['period']}): {float(row['glucose_value'])} ммоль/л"
            if row['glucose_type']:
                result += f" ({row['glucose_type']})"
            if row['comment']:
                result += f"\n   📝 {row['comment']}"
            if row['insulin_dose']:
                result += f"\n   💉 Доза: {float(row['insulin_dose'])} ед."
            if row['insulin_recommendation']:
                result += f"\n   📋 {row['insulin_recommendation']}"
            result += "\n"
        if len(glucose_rows) > 20:
            result += f"\n... и ещё {len(glucose_rows) - 20} записей\n"
    
    return result

async def generate_student_excel(student_id: int) -> str:
    """Генерирует Excel файл с данными подопечного (давление и глюкоза)"""
    wb = Workbook()
    
    ws_pressure = wb.create_sheet("Давление", 0)
    ws_glucose = wb.create_sheet("Глюкоза", 1)
    
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    
    # Заголовки для давления
    headers_pressure = ['Дата', 'Время', 'Период', 'Верхнее', 'Нижнее', 'Пульс', 'Комментарий']
    for col, header in enumerate(headers_pressure, 1):
        cell = ws_pressure.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    col_widths_pressure = {'A': 12, 'B': 10, 'C': 8, 'D': 10, 'E': 10, 'F': 8, 'G': 40}
    for col_letter, width in col_widths_pressure.items():
        ws_pressure.column_dimensions[col_letter].width = width
    ws_pressure.row_dimensions[1].height = 20
    
    # Заголовки для глюкозы (с новыми столбцами для инсулина)
    headers_glucose = ['Дата', 'Время', 'Период', 'Глюкоза', 'Тип замера', 'Комментарий', 'Доза инсулина', 'Рекомендация']
    for col, header in enumerate(headers_glucose, 1):
        cell = ws_glucose.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    col_widths_glucose = {'A': 12, 'B': 10, 'C': 8, 'D': 10, 'E': 25, 'F': 40, 'G': 15, 'H': 50}
    for col_letter, width in col_widths_glucose.items():
        ws_glucose.column_dimensions[col_letter].width = width
    ws_glucose.row_dimensions[1].height = 20
    
    async with db_pool.acquire() as conn:
        # Давление
        rows = await conn.fetch('''
            SELECT date, time, period, systolic, diastolic, pulse, comment 
            FROM pressure 
            WHERE user_id = $1 
            ORDER BY date DESC, time DESC
        ''', student_id)
        
        row_num = 2
        for row in rows:
            ws_pressure.cell(row=row_num, column=1, value=row['date'].strftime("%d-%m-%Y"))
            ws_pressure.cell(row=row_num, column=2, value=row['time'].strftime("%H:%M:%S"))
            ws_pressure.cell(row=row_num, column=3, value=row['period'])
            ws_pressure.cell(row=row_num, column=4, value=row['systolic'])
            ws_pressure.cell(row=row_num, column=5, value=row['diastolic'])
            ws_pressure.cell(row=row_num, column=6, value=row['pulse'] if row['pulse'] else "")
            ws_pressure.cell(row=row_num, column=7, value=row['comment'] if row['comment'] else "")
            row_num += 1
        
        # Глюкоза
        rows = await conn.fetch('''
            SELECT date, time, period, glucose_value, glucose_type, comment, insulin_dose, insulin_recommendation
            FROM glucose 
            WHERE user_id = $1 
            ORDER BY date DESC, time DESC
        ''', student_id)
        
        row_num = 2
        for row in rows:
            ws_glucose.cell(row=row_num, column=1, value=row['date'].strftime("%d-%m-%Y"))
            ws_glucose.cell(row=row_num, column=2, value=row['time'].strftime("%H:%M:%S"))
            ws_glucose.cell(row=row_num, column=3, value=row['period'])
            ws_glucose.cell(row=row_num, column=4, value=float(row['glucose_value']))
            ws_glucose.cell(row=row_num, column=5, value=row['glucose_type'] if row['glucose_type'] else "")
            ws_glucose.cell(row=row_num, column=6, value=row['comment'] if row['comment'] else "")
            ws_glucose.cell(row=row_num, column=7, value=float(row['insulin_dose']) if row['insulin_dose'] else "")
            ws_glucose.cell(row=row_num, column=8, value=row['insulin_recommendation'] if row['insulin_recommendation'] else "")
            row_num += 1
    
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        wb.save(tmp.name)
        return tmp.name

# ==================== ОПРЕДЕЛЕНИЕ ПЕРИОДА ====================
def get_period_by_time() -> str:
    """Определяет период дня по текущему времени"""
    now = datetime.now(MSK_PLUS_1)
    hour = now.hour
    
    if 6 <= hour < 12:
        return "Утро"
    elif 12 <= hour < 18:
        return "День"
    else:
        return "Вечер"

# ==================== РАБОТА С ДАННЫМИ ====================
def detect_glucose_type(text: str) -> str:
    """Определяет тип замера глюкозы по тексту"""
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

def extract_insulin_dose(text: str) -> Optional[float]:
    """Извлекает дозу инсулина из текста (формат: инсулин Xед или X ед)"""
    patterns = [
        r'инсулин\s*(\d+[.,]?\d*)\s*ед',
        r'инсулин\s*(\d+[.,]?\d*)',
        r'(\d+[.,]?\d*)\s*ед\s*инсулин',
        r'(\d+[.,]?\d*)\s*ед'
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return float(match.group(1).replace(',', '.'))
    return None

async def save_pressure_to_db(user_id: int, period: str, systolic: int, diastolic: int, pulse: int, comment: str) -> None:
    """Сохраняет показания давления в БД"""
    now = datetime.now(MSK_PLUS_1)
    
    if now.hour < 6:
        date_val = (now - timedelta(days=1)).date()
    else:
        date_val = now.date()
    
    time_val = now.time()
    
    async with db_pool.acquire() as conn:
        await conn.execute('''
            INSERT INTO pressure (user_id, date, time, period, systolic, diastolic, pulse, comment)
            VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
        ''', user_id, date_val, time_val, period, systolic, diastolic, pulse, comment)

async def save_glucose_to_db(user_id: int, period: str, glucose: float, glucose_type: str, comment: str, insulin_dose: Optional[float] = None, insulin_recommendation: Optional[str] = None) -> None:
    """Сохраняет показания глюкозы в БД"""
    now = datetime.now(MSK_PLUS_1)
    
    if now.hour < 6:
        date_val = (now - timedelta(days=1)).date()
    else:
        date_val = now.date()
    
    time_val = now.time()
    
    async with db_pool.acquire() as conn:
        try:
            await conn.execute('''
                INSERT INTO glucose (user_id, date, time, period, glucose_value, glucose_type, comment, insulin_dose, insulin_recommendation)
                VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
            ''', user_id, date_val, time_val, period, glucose, glucose_type, comment, insulin_dose, insulin_recommendation)
        except Exception as e:
            # Если столбцов нет - сохраняем без них
            if 'column "insulin_dose" does not exist' in str(e) or 'column "insulin_recommendation" does not exist' in str(e):
                await conn.execute('''
                    INSERT INTO glucose (user_id, date, time, period, glucose_value, glucose_type, comment)
                    VALUES ($1, $2, $3, $4, $5, $6, $7)
                ''', user_id, date_val, time_val, period, glucose, glucose_type, comment)
            else:
                raise e

async def get_today_pressure_report(user_id: int) -> str:
    """Возвращает отчёт по давлению за сегодня"""
    today = datetime.now(MSK_PLUS_1).date()
    
    async with db_pool.acquire() as conn:
        rows = await conn.fetch('''
            SELECT date, time, period, systolic, diastolic, pulse, comment 
            FROM pressure 
            WHERE user_id = $1 AND date = $2
            ORDER BY time ASC
        ''', user_id, today)
    
    if not rows:
        return f"📊 Отчет по давлению за {today.strftime('%d-%m-%Y')}\n\nНет данных."
    
    report = f"📊 Отчет по давлению за {today.strftime('%d-%m-%Y')}\n\n"
    for row in rows:
        period_emoji = {"Утро": "🌅", "День": "☀️", "Вечер": "🌙"}
        emoji = period_emoji.get(row['period'], "")
        
        report += f"{emoji} {row['period']} {row['time'].strftime('%H:%M:%S')}: {row['systolic']}/{row['diastolic']}"
        if row['pulse']:
            report += f", пульс {row['pulse']}"
        if row['comment']:
            report += f"\n   📝 {row['comment']}"
        report += "\n\n"
    
    return report

async def get_today_glucose_report(user_id: int) -> str:
    """Возвращает отчёт по глюкозе за сегодня"""
    today = datetime.now(MSK_PLUS_1).date()
    
    async with db_pool.acquire() as conn:
        rows = await conn.fetch('''
            SELECT date, time, period, glucose_value, glucose_type, comment, insulin_dose, insulin_recommendation
            FROM glucose 
            WHERE user_id = $1 AND date = $2
            ORDER BY time ASC
        ''', user_id, today)
    
    if not rows:
        return f"📊 Отчет по глюкозе за {today.strftime('%d-%m-%Y')}\n\nНет данных."
    
    report = f"📊 Отчет по глюкозе за {today.strftime('%d-%m-%Y')}\n\n"
    for row in rows:
        period_emoji = {"Утро": "🌅", "День": "☀️", "Вечер": "🌙"}
        emoji = period_emoji.get(row['period'], "")
        
        report += f"{emoji} {row['period']} {row['time'].strftime('%H:%M:%S')}: глюкоза {float(row['glucose_value'])}"
        if row['glucose_type']:
            report += f" ({row['glucose_type']})"
        if row['comment']:
            report += f"\n   📝 {row['comment']}"
        if row['insulin_dose']:
            report += f"\n   💉 Доза: {float(row['insulin_dose'])} ед."
        if row['insulin_recommendation']:
            report += f"\n   📋 {row['insulin_recommendation']}"
        report += "\n\n"
    
    return report

# ==================== АДМИН КОМАНДЫ ====================
async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Админ панель"""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    await update.message.reply_text(
        "👑 Админ панель\n\n"
        "Доступные команды:\n"
        "/users - список пользователей\n"
        "/users_excel - выгрузить пользователей в Excel\n"
        "/grant username дни - выдать доступ (5,7,30)\n"
        "/backup - создать резервную копию (SQL дамп)\n"
        "/restore - восстановить данные из SQL дампа\n"
        "/status - статус бота\n"
        "/test_remind - тестовая отправка напоминаний"
    )

async def admin_users(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Список пользователей (админ)"""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    users = await get_all_users()
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

async def admin_users_excel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Выгрузка пользователей в Excel (админ)"""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    users = await get_all_users()
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

async def admin_grant(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Выдать доступ пользователю (админ)"""
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
    
    users = await get_all_users()
    found = None
    for uid, data in users.items():
        if data.get('username', '').lower() == username.lower():
            found = int(uid)
            break
    
    if found:
        await grant_access(found, days)
        await update.message.reply_text(f"✅ Пользователю {username} выдан доступ на {days} дней.")
    else:
        await update.message.reply_text(f"❌ Пользователь {username} не найден.")

async def backup_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Создаёт резервную копию БД (админ)"""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    try:
        timestamp = datetime.now(MSK_PLUS_1).strftime("%Y%m%d_%H%M%S")
        filename = f"backup_{timestamp}.sql"
        
        result = subprocess.run(
            ['pg_dump', DATABASE_URL, '--clean', '--if-exists', '-f', filename],
            capture_output=True,
            text=True
        )
        
        if result.returncode != 0:
            await update.message.reply_text(f"❌ Ошибка при создании дампа: {result.stderr}")
            return
        
        with open(filename, 'rb') as f:
            await update.message.reply_document(
                document=f,
                filename=filename,
                caption=f"📦 Резервная копия БД от {datetime.now(MSK_PLUS_1).strftime('%d-%m-%Y %H:%M:%S')}"
            )
        
        os.remove(filename)
        print(f"Создана резервная копия БД: {filename}")
        
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка при создании резервной копии: {e}")

async def restore_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Восстанавливает данные из SQL дампа (админ)"""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    await update.message.reply_text(
        "📤 Отправьте SQL дамп (созданный командой /backup).\n\n"
        "⚠️ ВНИМАНИЕ: текущие данные будут ПЕРЕЗАПИСАНЫ!"
    )
    context.user_data['awaiting_restore'] = 'sql'

async def status_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Проверка статуса бота (админ)"""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    status = "работает" if context.application.job_queue else "НЕ РАБОТАЕТ"
    await update.message.reply_text(f"🤖 Статус бота:\n\nJobQueue: {status}\nБаза данных: ✅ подключена")

async def test_remind_all(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Тестовая отправка напоминаний (админ)"""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Доступ запрещён.")
        return
    
    await update.message.reply_text("🔄 Отправляю тестовые напоминания всем пользователям...")
    
    users = await get_all_users()
    sent = 0
    for uid, data in users.items():
        if await check_access(int(uid)):
            try:
                await context.bot.send_message(
                    chat_id=int(uid),
                    text="🧪 ТЕСТ: Напоминание работает! Если вы это видите — бот исправен."
                )
                sent += 1
            except Exception:
                pass
    
    await update.message.reply_text(f"✅ Тестовое напоминание отправлено {sent} пользователям")

# ==================== КОМАНДЫ НАСТАВНИКА ====================
async def add_student_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Добавить подопечного (команда для наставника)"""
    user_id = update.effective_user.id
    
    args = context.args
    if len(args) < 1:
        await update.message.reply_text(
            "📝 Используйте: /add_student @username\n"
            "Пример: /add_student @john_doe"
        )
        return
    
    username = args[0].lstrip('@')
    
    success, message = await add_student(user_id, username)
    await update.message.reply_text(message)

async def remove_student_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Удалить подопечного (команда для наставника)"""
    user_id = update.effective_user.id
    
    args = context.args
    if len(args) < 1:
        await update.message.reply_text(
            "📝 Используйте: /remove_student @username\n"
            "Пример: /remove_student @john_doe"
        )
        return
    
    username = args[0].lstrip('@')
    
    success, message = await remove_student(user_id, username)
    await update.message.reply_text(message)

async def list_students_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Показать список подопечных (команда для наставника)"""
    user_id = update.effective_user.id
    
    students = await get_mentor_students(user_id)
    
    if not students:
        await update.message.reply_text("📋 У вас пока нет подопечных.\n\nДобавьте командой /add_student @username")
        return
    
    text = "📋 ВАШИ ПОДОПЕЧНЫЕ:\n\n"
    for i, student in enumerate(students, 1):
        text += f"{i}. @{student['username']}\n"
        text += f"   🆔 ID: {student['user_id']}\n"
        text += f"   📅 Добавлен: {student['added_at']}\n\n"
    
    text += "\nКоманды для просмотра:\n"
    text += "/view @username - просмотреть журнал\n"
    text += "/view_excel @username - скачать Excel-файл"
    
    await update.message.reply_text(text)

async def view_student_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Просмотреть данные подопечного (команда для наставника)"""
    mentor_id = update.effective_user.id
    
    args = context.args
    if len(args) < 1:
        await update.message.reply_text(
            "📝 Используйте: /view @username\n"
            "Пример: /view @john_doe"
        )
        return
    
    username = args[0].lstrip('@')
    
    student_id = await get_user_by_username(username)
    if not student_id:
        await update.message.reply_text(f"❌ Пользователь @{username} не найден.")
        return
    
    if not await is_mentor_for_student(mentor_id, student_id):
        await update.message.reply_text(f"❌ Пользователь @{username} не является вашим подопечным.")
        return
    
    await update.message.reply_text("📊 Загружаю данные подопечного...")
    
    history = await get_student_full_history(student_id)
    
    if len(history) > 4000:
        parts = [history[i:i+4000] for i in range(0, len(history), 4000)]
        for part in parts:
            await update.message.reply_text(part)
    else:
        await update.message.reply_text(history)

async def view_student_excel_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Выгрузить Excel-файл подопечного (команда для наставника)"""
    mentor_id = update.effective_user.id
    
    args = context.args
    if len(args) < 1:
        await update.message.reply_text(
            "📝 Используйте: /view_excel @username\n"
            "Пример: /view_excel @john_doe"
        )
        return
    
    username = args[0].lstrip('@')
    
    student_id = await get_user_by_username(username)
    if not student_id:
        await update.message.reply_text(f"❌ Пользователь @{username} не найден.")
        return
    
    if not await is_mentor_for_student(mentor_id, student_id):
        await update.message.reply_text(f"❌ Пользователь @{username} не является вашим подопечным.")
        return
    
    await update.message.reply_text("📊 Генерирую Excel-файл подопечного...")
    
    try:
        filename = await generate_student_excel(student_id)
        
        with open(filename, 'rb') as f:
            await update.message.reply_document(
                document=f,
                filename=f"student_{username}_{datetime.now(MSK_PLUS_1).strftime('%Y%m%d')}.xlsx",
                caption=f"📊 Данные подопечного @{username}\nДата выгрузки: {datetime.now(MSK_PLUS_1).strftime('%d-%m-%Y %H:%M:%S')}"
            )
        
        os.remove(filename)
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка при создании файла: {e}")

# ==================== ОСНОВНЫЕ КОМАНДЫ ====================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Стартовая команда"""
    user_id = update.effective_user.id
    username = update.effective_user.username or update.effective_user.first_name
    
    is_new = await add_user(user_id, username)
    
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
        "• 5.5 через 2 часа после еды - глюкоза с типом замера\n"
        "• 8.5 инсулин 10ед - глюкоза с дозой инсулина\n\n"
        "🌅 Бот сам определит время суток (Утро, День, Вечер)\n"
        "💉 Для глюкозы будет рассчитана рекомендация по инсулину\n"
        "💾 Все данные хранятся в защищённой базе данных\n\n"
        "📋 Если вы наставник, используйте команды:\n"
        "/add_student @username - добавить подопечного\n"
        "/students - список подопечных\n"
        "/view @username - просмотреть журнал подопечного\n"
        "/view_excel @username - скачать Excel-файл подопечного\n\n"
        "Основные команды:\n"
        "/table - получить Excel файл (давление и глюкоза)\n"
        "/report - отчет по давлению за сегодня\n"
        "/glucose_report - отчет по глюкозе за сегодня\n"
        "/help - помощь"
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Помощь"""
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
        "• 5.5 через 2 часа после еды\n"
        "• 8.5 инсулин 10ед - глюкоза с дозой инсулина\n\n"
        "💉 Рекомендация по инсулину:\n"
        "• Утро: цель 7.0, базовая доза 10 ед.\n"
        "• Вечер: цель 13.0, базовая доза 8 ед.\n"
        "• Если сахар выше цели на ≥1 → +2 ед.\n"
        "• Если сахар ниже цели на ≥1 → -2 ед.\n\n"
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
        "📋 Команды наставника:\n"
        "/add_student @username - добавить подопечного\n"
        "/remove_student @username - удалить подопечного\n"
        "/students - список подопечных\n"
        "/view @username - журнал подопечного\n"
        "/view_excel @username - Excel-файл подопечного\n\n"
        f"📢 <a href='https://t.me/+MAuGbcnBQmgxZTIy'>Больше наших ботов в канале</a>"
    )
    
    await update.message.reply_text(text, parse_mode="HTML", disable_web_page_preview=True)

async def report_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Отчет по давлению за сегодня"""
    user_id = update.effective_user.id
    report = await get_today_pressure_report(user_id)
    await update.message.reply_text(report)

async def glucose_report_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Отчет по глюкозе за сегодня"""
    user_id = update.effective_user.id
    report = await get_today_glucose_report(user_id)
    await update.message.reply_text(report)

async def table_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Excel файл для пользователя"""
    user_id = update.effective_user.id
    
    await update.message.reply_text("🔄 Генерирую ваш Excel-файл...")
    
    try:
        filename = await generate_student_excel(user_id)
        
        with open(filename, 'rb') as f:
            await update.message.reply_document(
                document=f,
                filename="medical_journal.xlsx",
                caption="📊 Ваш медицинский журнал (давление и глюкоза)\n\n"
                        "Все ваши данные — только ваши. Другие пользователи не видят их."
            )
        
        os.remove(filename)
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка при создании файла: {e}")

# ==================== ОБРАБОТКА СООБЩЕНИЙ ====================
async def handle_pressure_glucose(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработка текстовых сообщений (давление и глюкоза)"""
    user_id = update.effective_user.id
    username = update.effective_user.username or update.effective_user.first_name
    
    await add_user(user_id, username)
    
    if not await check_access(user_id):
        await update.message.reply_text(
            f"⛔ Доступ временно приостановлен.\nСвяжитесь с администратором @{ADMIN_USERNAME}"
        )
        return
    
    await check_and_send_3day_reminder(user_id, context.application)
    await update_user_days(user_id)
    
    text = update.message.text.strip()
    text_lower = text.lower()
    
    # Извлекаем дозу инсулина из текста
    insulin_dose = extract_insulin_dose(text)
    
    # Ищем все числа (включая десятичные)
    numbers = re.findall(r'\d+[.,]?\d*', text)
    numbers = [float(n.replace(',', '.')) for n in numbers]
    
    # Формируем ответ для уведомления наставников
    response_value = ""
    
    # Если есть упоминание инсулина или одно число 1-30 — это глюкоза
    is_insulin = any(word in text_lower for word in ["инсулин", "инсулина", "ед", "единиц", "iu", "ме", "мед", "единица"])
    
    if is_insulin or (len(numbers) >= 1 and 1 <= numbers[0] <= 30):
        # Это глюкоза
        glucose = None
        
        for n in numbers:
            if glucose is None and 1 <= n <= 30:
                glucose = n
                break
        
        if glucose is None and numbers:
            glucose = numbers[0]
        
        glucose_type = detect_glucose_type(text)
        period = get_period_by_time()
        
        # Удаляем числа и ключевые слова из комментария
        comment = re.sub(r'\d+[.,]?\d*', '', text)
        comment = re.sub(r'инсулин|инсулина|ед|единиц|iu|ме|мед', '', comment, flags=re.IGNORECASE)
        comment = re.sub(r'натощак|через 2 часа после еды|перед едой|перед сном|ночью', '', comment, flags=re.IGNORECASE)
        comment = re.sub(r'[\s/]+', ' ', comment).strip()
        
        # Рассчитываем рекомендацию по инсулину
        insulin_rec = calculate_insulin_recommendation(period, glucose, insulin_dose)
        
        # Сохраняем в БД
        await save_glucose_to_db(
            user_id, 
            period, 
            glucose, 
            glucose_type, 
            comment, 
            insulin_rec['current_dose'], 
            insulin_rec['recommendation_text']
        )
        
        period_emoji = {"Утро": "🌅", "День": "☀️", "Вечер": "🌙"}
        now = datetime.now(MSK_PLUS_1)
        
        response_value = f"глюкоза {glucose}"
        if glucose_type != "без указания":
            response_value += f" ({glucose_type})"
        if insulin_dose:
            response_value += f", доза {insulin_dose} ед."
        
        # Формируем ответ
        response = f"✅ Записано! {period_emoji.get(period, '')} {period}: {response_value}\n"
        response += f"📅 {now.strftime('%d-%m-%Y %H:%M:%S')}\n"
        if comment:
            response += f"📝 {comment}\n"
        response += f"\n💉 Рекомендация по инсулину:\n{insulin_rec['recommendation_text']}"
        
        await update.message.reply_text(response)
        
        # Уведомляем наставников
        await notify_mentors_about_measurement(user_id, "Глюкоза", response_value, context.application)
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
            "8.5 инсулин 10ед - глюкоза с дозой инсулина\n"
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
    await save_pressure_to_db(user_id, period, systolic, diastolic, pulse, comment)
    
    period_emoji = {"Утро": "🌅", "День": "☀️", "Вечер": "🌙"}
    now = datetime.now(MSK_PLUS_1)
    
    response_value = f"{systolic}/{diastolic}"
    if pulse:
        response_value += f", пульс {pulse}"
    
    response = f"✅ Записано! {period_emoji.get(period, '')} {period}: {response_value}\n"
    if comment:
        response += f"📝 {comment}\n"
    response += f"📅 {now.strftime('%d-%m-%Y %H:%M:%S')}"
    
    await update.message.reply_text(response)
    
    # Уведомляем наставников
    await notify_mentors_about_measurement(user_id, "Давление", response_value, context.application)

# ==================== ВОССТАНОВЛЕНИЕ ИЗ SQL ДАМПА ====================
async def handle_restore_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработчик загруженного SQL файла для восстановления"""
    if not is_admin(update.effective_user.id):
        return
    
    if context.user_data.get('awaiting_restore') != 'sql':
        return
    
    document = update.message.document
    if not document or not document.file_name.endswith('.sql'):
        await update.message.reply_text("❌ Пожалуйста, отправьте SQL дамп (созданный командой /backup)")
        return
    
    try:
        file = await context.bot.get_file(document.file_id)
        
        temp_file = f"temp_restore_{datetime.now(MSK_PLUS_1).strftime('%Y%m%d_%H%M%S')}.sql"
        await file.download_to_drive(temp_file)
        
        result = subprocess.run(
            ['psql', DATABASE_URL, '-f', temp_file],
            capture_output=True,
            text=True
        )
        
        os.remove(temp_file)
        
        if result.returncode != 0:
            await update.message.reply_text(f"❌ Ошибка при восстановлении: {result.stderr}")
        else:
            await update.message.reply_text(f"✅ Данные восстановлены из файла {document.file_name}")
        
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка при восстановлении: {e}")
    finally:
        context.user_data['awaiting_restore'] = None

# ==================== НАПОМИНАНИЯ ====================
async def send_scheduled_reminder(context: ContextTypes.DEFAULT_TYPE) -> None:
    """Отправляет напоминания всем активным пользователям"""
    now_time = datetime.now(MSK_PLUS_1)
    current_hour = now_time.hour
    
    if current_hour not in [8, 14, 20]:
        return
    
    print(f"[{now_time.strftime('%Y-%m-%d %H:%M:%S')}] Запуск напоминаний, час: {current_hour}")
    
    users = await get_all_users()
    sent_count = 0
    active_count = 0
    
    for uid, data in users.items():
        if await check_access(int(uid)):
            active_count += 1
            try:
                await context.bot.send_message(
                    chat_id=int(uid),
                    text="🔔 Напоминание: пора измерить давление и глюкозу!\n\n"
                         "Просто отправьте мне показания:\n"
                         "• 120 80 - давление\n"
                         "• 120 80 68 - давление и пульс\n"
                         "• 5.5 - глюкоза\n"
                         "• 5.5 натощак - глюкоза с типом замера\n"
                         "• 8.5 инсулин 10ед - глюкоза с дозой инсулина"
                )
                sent_count += 1
                print(f"  → Напоминание отправлено пользователю {uid}")
            except Exception as e:
                print(f"  ✗ Ошибка отправки пользователю {uid}: {e}")
    
    print(f"Активных пользователей: {active_count}, отправлено напоминаний: {sent_count}")

# ==================== КОМАНДЫ МЕНЮ ====================
async def set_commands(app: Application) -> None:
    """Устанавливает команды для бота"""
    
    # Команды для всех пользователей
    default_commands = [
        BotCommand("start", "Главное меню"),
        BotCommand("table", "Excel журнал (давление+глюкоза)"),
        BotCommand("report", "Отчет по давлению за сегодня"),
        BotCommand("glucose_report", "Отчет по глюкозе за сегодня"),
        BotCommand("help", "Помощь"),
        BotCommand("add_student", "Добавить подопечного"),
        BotCommand("remove_student", "Удалить подопечного"),
        BotCommand("students", "Список подопечных"),
        BotCommand("view", "Просмотреть журнал подопечного"),
        BotCommand("view_excel", "Скачать Excel подопечного"),
    ]
    
    # Команды для админа
    admin_commands = [
        BotCommand("start", "Главное меню"),
        BotCommand("table", "Excel журнал (давление+глюкоза)"),
        BotCommand("report", "Отчет по давлению за сегодня"),
        BotCommand("glucose_report", "Отчет по глюкозе за сегодня"),
        BotCommand("help", "Помощь"),
        BotCommand("add_student", "Добавить подопечного"),
        BotCommand("remove_student", "Удалить подопечного"),
        BotCommand("students", "Список подопечных"),
        BotCommand("view", "Просмотреть журнал подопечного"),
        BotCommand("view_excel", "Скачать Excel подопечного"),
        BotCommand("admin", "Админ панель"),
        BotCommand("users", "Список пользователей"),
        BotCommand("users_excel", "Выгрузить пользователей в Excel"),
        BotCommand("grant", "Выдать доступ (username дни)"),
        BotCommand("backup", "Резервная копия БД"),
        BotCommand("restore", "Восстановить БД"),
        BotCommand("status", "Статус бота"),
        BotCommand("test_remind", "Тест напоминаний"),
    ]
    
    # Устанавливаем команды для всех
    await app.bot.set_my_commands(default_commands)
    
    # Дополнительные команды только для админа
    await app.bot.set_my_commands(admin_commands, scope=BotCommandScopeChat(chat_id=ADMIN_ID))

# ==================== ЗАПУСК ====================
def main() -> None:
    """Запуск бота"""
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.run_until_complete(init_db())
    
    app = Application.builder().token(TOKEN).build()
    
    if app.job_queue is None:
        print("❌ ОШИБКА: JobQueue не создан! Напоминания работать не будут")
    else:
        print("✅ JobQueue создан успешно")
        print(f"   Текущее время сервера: {datetime.now(MSK_PLUS_1).strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"   Часовой пояс: Europe/Samara (МСК+1)")
    
    # Регистрируем обработчики основных команд
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("report", report_command))
    app.add_handler(CommandHandler("glucose_report", glucose_report_command))
    app.add_handler(CommandHandler("table", table_command))
    
    # Команды наставника
    app.add_handler(CommandHandler("add_student", add_student_command))
    app.add_handler(CommandHandler("remove_student", remove_student_command))
    app.add_handler(CommandHandler("students", list_students_command))
    app.add_handler(CommandHandler("view", view_student_command))
    app.add_handler(CommandHandler("view_excel", view_student_excel_command))
    
    # Админ команды
    app.add_handler(CommandHandler("admin", admin_panel))
    app.add_handler(CommandHandler("users", admin_users))
    app.add_handler(CommandHandler("users_excel", admin_users_excel))
    app.add_handler(CommandHandler("grant", admin_grant))
    app.add_handler(CommandHandler("backup", backup_command))
    app.add_handler(CommandHandler("restore", restore_command))
    app.add_handler(CommandHandler("status", status_command))
    app.add_handler(CommandHandler("test_remind", test_remind_all))
    
    # Обработчики сообщений
    app.add_handler(MessageHandler(filters.Document.ALL, handle_restore_file))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_pressure_glucose))
    
    # Настройка напоминаний
    job_queue = app.job_queue
    if job_queue:
        job_queue.run_daily(send_scheduled_reminder, time(5, 0))
        job_queue.run_daily(send_scheduled_reminder, time(11, 0))
        job_queue.run_daily(send_scheduled_reminder, time(17, 0))
        print("Напоминания: 8:00, 14:00, 20:00 (МСК+1)")
    else:
        print("ОШИБКА: job_queue не создан! Напоминания работать не будут")
    
    print("🤖 Бот запущен")
    
    app.run_polling()

if __name__ == "__main__":
    main()