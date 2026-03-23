import asyncio
import os
import sys
from datetime import datetime
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.utils.keyboard import ReplyKeyboardBuilder
from dotenv import load_dotenv
from gsheets_storage import GoogleSheetsStorage

from aiogram.types import FSInputFile

load_dotenv()

# Constants
API_TOKEN = os.getenv("BOT_TOKEN")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")
ADMIN_ID = os.getenv("ADMIN_ID")

# Bot and Dispatcher
bot = Bot(token=API_TOKEN)
dp = Dispatcher()

# Storage client
storage = GoogleSheetsStorage(GOOGLE_CREDENTIALS_FILE, GOOGLE_SHEET_ID)

# States
class Registration(StatesGroup):
    waiting_for_name = State()

class AttendanceStates(StatesGroup):
    waiting_for_hours = State()

# Main menu keyboard
def get_main_menu_keyboard():
    builder = ReplyKeyboardBuilder()
    builder.button(text="Пришел")
    builder.button(text="Ушел раньше")
    builder.adjust(2)
    return builder.as_markup(resize_keyboard=True)

# Handlers
@dp.message(Command("start"))
async def start_command(message: types.Message, state: FSMContext):
    """
    Greets the user and asks for registration if necessary.
    """
    user_id = message.from_user.id
    
    # Check if user is already registered in Excel
    registered_name = storage.get_employee_name(user_id)
    
    if registered_name:
        await message.answer(
            f"С возвращением, {registered_name}! Пожалуйста, используй кнопки ниже.", 
            reply_markup=get_main_menu_keyboard()
        )
    else:
        await message.answer("Привет! Ты еще не зарегистрирован. Пожалуйста, введи свою Фамилию и Инициалы (например: Иванов И.И.)")
        await state.set_state(Registration.waiting_for_name)

@dp.message(Registration.waiting_for_name)
async def process_name(message: types.Message, state: FSMContext):
    """
    Saves the user's name and completes registration.
    """
    full_name = message.text.strip()
    user_id = message.from_user.id
    
    if len(full_name) < 3:
        await message.answer("Пожалуйста, введите корректную Фамилию и Инициалы.")
        return
        
    storage.register_employee(user_id, full_name)
    await state.clear()
    
    await message.answer(
        f"Регистрация прошла успешно, {full_name}! Теперь ты можешь отмечать посещаемость.", 
        reply_markup=get_main_menu_keyboard()
    )

@dp.message(F.text == "Пришел")
async def check_in(message: types.Message):
    """
    Logs check-in time.
    """
    user_id = message.from_user.id
    full_name = storage.get_employee_name(user_id)
    
    if not full_name:
        await message.answer("Сначала зарегистрируйся! Напиши /start")
        return
        
    try:
        storage.add_attendance(user_id, full_name, "Пришел")
        await message.answer(f"Записано! Удачной смены, {full_name}!")
    except Exception as e:
        await message.answer(f"Ошибка при записи: {e}")

@dp.message(F.text == "Ушел раньше")
async def process_early_leave(message: types.Message, state: FSMContext):
    """
    Asks how many hours the employee worked.
    """
    user_id = message.from_user.id
    full_name = storage.get_employee_name(user_id)
    
    if not full_name:
        await message.answer("Сначала зарегистрируйся! Напиши /start")
        return
        
    await message.answer("Сколько часов вы отработали сегодня? (Напишите число, например: 4)")
    await state.set_state(AttendanceStates.waiting_for_hours)

@dp.message(AttendanceStates.waiting_for_hours)
async def save_hours(message: types.Message, state: FSMContext):
    """
    Saves the hours instead of a plus in the summary sheet.
    """
    hours = message.text.strip()
    user_id = message.from_user.id
    full_name = storage.get_employee_name(user_id)
    
    try:
        # Save to Excel with the provided hours
        storage.add_attendance(user_id, full_name, "Ушел раньше", value=hours)
        await state.clear()
        await message.answer(f"Записано! Отработано часов: {hours}. Хорошего отдыха!")
    except Exception as e:
        await message.answer(f"Ошибка при записи: {e}")

@dp.message(Command("report"))
async def send_report(message: types.Message):
    """
    Sends the Google Sheet link to the admin.
    """
    if str(message.from_user.id) != str(ADMIN_ID):
        await message.answer("У вас нет прав для получения отчета.")
        return

    sheet_url = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/edit"
    await message.answer(f"Ссылка на таблицу посещаемости:\n{sheet_url}")

@dp.message(Command("archives"))
async def list_archives(message: types.Message):
    """
    Archives are managed by Google Sheets history.
    """
    await message.answer("В Google Sheets архивация происходит автоматически через историю изменений таблицы.")

@dp.message(F.text.startswith("/download_"))
async def download_archive(message: types.Message):
    """
    Handles downloading a specific archive file.
    """
    if str(message.from_user.id) != str(ADMIN_ID):
        await message.answer("У вас нет прав.")
        return

    filename = message.text.replace("/download_", "") + ".xlsx"
    filepath = os.path.join("archive", filename)

    if os.path.exists(filepath):
        file = FSInputFile(filepath)
        await message.answer_document(file, caption=f"Архивный отчет: {filename}")
    else:
        await message.answer("Файл не найден.")

async def main():
    # Удаляем вебхук и все пропущенные обновления перед запуском
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    print("Бот запускается...")
    sys.stdout.flush()
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("Бот остановлен.")
    except Exception as e:
        print(f"Ошибка при запуске: {e}")
        sys.stderr.flush()
