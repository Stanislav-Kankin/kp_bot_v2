import asyncio
import logging
from bot import dp, bot

# Настройка логирования
logging.basicConfig(level=logging.INFO)

async def main():
    print("🤖 Бот для создания КП запускается...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())