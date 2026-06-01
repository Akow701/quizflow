import os
import asyncio
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message
from aiogram.filters import CommandStart
from config import BOT_TOKEN
from processor import extract_text, generate_quiz

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
UPLOAD_DIR = "downloads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

@dp.message(CommandStart())
async def start_handler(message: Message):
    await message.answer("PDF, DOCX, TXT yoki XLSX yuboring.")

@dp.message(F.document)
async def document_handler(message: Message):
    try:
        document = message.document
        file_path = os.path.join(UPLOAD_DIR, document.file_name)
        await bot.download(document, destination=file_path)

        text = extract_text(file_path)
        if not text:
            await message.answer("Fayldan matn chiqarilmadi")
            return

        quiz_data = generate_quiz(text)
        for q in quiz_data:
            await bot.send_poll(
                chat_id=message.chat.id,
                question=q["question"],
                options=q["options"],
                type="quiz",
                correct_option_id=q["correct"],
                is_anonymous=False
            )
            await asyncio.sleep(1)

    except Exception as e:
        await message.answer(f"Xatolik: {e}")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
