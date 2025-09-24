from aiogram import Router, types, F
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.types import (
    InlineKeyboardButton, InlineKeyboardMarkup,
    FSInputFile
    )
from bot.states import FormKP
from bot.ppt_service import ppt_service
import asyncio
import os

router = Router()


@router.message(Command("start"))
async def start(message: types.Message):
    await message.answer(
        "🤖 <b>Бот для создания коммерческих предложений HRlink</b>\n\n"
        "▶️ <b>Нажмите /make_kp чтобы начать</b>",
        parse_mode='HTML'
    )


@router.message(Command("make_kp"))
async def make_kp(message: types.Message, state: FSMContext):
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(
            text="📊 КП длинный вариант (общий)",
            callback_data="template_long"
        )],
        [InlineKeyboardButton(
            text="📈 КП короткий вариант (общий)",
            callback_data="template_short"
        )]
    ])

    await message.answer(
        "🎯 <b>Выберите вариант коммерческого предложения:</b>",
        reply_markup=keyboard,
        parse_mode='HTML'
    )
    await state.set_state(FormKP.template_type)


async def delete_file_with_delay(file_path: str, delay: int = 600):
    """Удаляет файл с задержкой (по умолчанию 10 минут)"""
    await asyncio.sleep(delay)
    if os.path.exists(file_path):
        os.remove(file_path)


@router.callback_query(FormKP.template_type)
async def process_template_choice(callback: types.CallbackQuery,
                                  state: FSMContext):
    template_type = callback.data.split("_")[1]
    await state.update_data(template_type=template_type)
    await state.set_state(FormKP.company_name)

    await callback.message.answer(
        "🏢 <b>Введите название компании:</b>",
        parse_mode='HTML'
    )
    await callback.answer()


@router.message(FormKP.company_name)
async def process_company_name(message: types.Message, state: FSMContext):
    await state.update_data(company_name=message.text)
    await state.set_state(FormKP.hr_licenses)

    await message.answer(
        "👥 <b>Введите количество лицензий кадровика:</b>",
        parse_mode='HTML'
    )


@router.message(FormKP.hr_licenses)
async def process_hr_licenses(message: types.Message, state: FSMContext):
    try:
        hr_licenses = int(message.text.strip())
        if hr_licenses <= 0:
            raise ValueError
    except ValueError:
        await message.answer(
            "❌ Пожалуйста, введите корректное число лицензий:"
            )
        return

    await state.update_data(hr_licenses=hr_licenses)
    await state.set_state(FormKP.employee_licenses)

    await message.answer(
        "👥 <b>Введите количество лицензий сотрудников:</b>",
        parse_mode='HTML'
    )


@router.message(FormKP.employee_licenses)
async def process_employee_licenses(message: types.Message, state: FSMContext):
    try:
        employee_licenses = int(message.text.strip())
        if employee_licenses <= 0:
            raise ValueError
    except ValueError:
        await message.answer(
            "❌ Пожалуйста, введите корректное число лицензий:"
            )
        return

    await state.update_data(employee_licenses=employee_licenses)
    await state.set_state(FormKP.on_premises)

    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(
            text="✅ Да", callback_data="on_premises_yes")],
        [InlineKeyboardButton(
            text="❌ Нет", callback_data="on_premises_no")]
    ])

    await message.answer(
        "🏢 <b>Нужен on-premises?</b>",
        reply_markup=keyboard,
        parse_mode='HTML'
    )


@router.callback_query(FormKP.on_premises)
async def process_on_premises(callback: types.CallbackQuery,
                              state: FSMContext):
    on_premises = "Да" if callback.data == "on_premises_yes" else "Нет"
    await state.update_data(on_premises=on_premises)

    # Получаем все данные
    data = await state.get_data()

    # Создаем презентацию
    await callback.message.answer(
        "🔄 <b>Создаю коммерческое предложение...</b>", parse_mode='HTML')

    presentation_path = ppt_service.create_kp_presentation(
        data['template_type'],
        data
    )

    if presentation_path and os.path.exists(presentation_path):
        file = FSInputFile(presentation_path)
        keyboard = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(
                text="📄 Сделать PDF",
                callback_data=f"make_pdf_{os.path.basename(presentation_path)}"
            )]
        ])
        await callback.message.answer_document(
            document=file,
            caption=f"✅ <b>Коммерческое предложение готово!</b>\n\n"
                    f"🏢 <b>Компания:</b> {data['company_name']}\n"
                    f"👥 <b>Лицензии кадровика:</b> {data['hr_licenses']}\n"
                    f"👥 <b>Лицензии сотрудников:</b> {
                        data['employee_licenses']}\n"
                    f"🏢 <b>On-premises:</b> {data['on_premises']}\n\n"
                    f"<i>Для создания нового КП нажмите /make_kp</i>",
            reply_markup=keyboard,
            parse_mode='HTML'
        )

        # Планируем удаление презентации через 10 минут
        asyncio.create_task(delete_file_with_delay(presentation_path))
    else:
        await callback.message.answer(
            "❌ <b>Ошибка при создании презентации</b>\n"
            "Попробуйте позже или обратитесь к администратору.",
            parse_mode='HTML'
        )

    await state.clear()


@router.callback_query(F.data.startswith("make_pdf_"))
async def make_pdf_handler(callback: types.CallbackQuery):
    filename = callback.data.replace("make_pdf_", "")
    pptx_path = os.path.join("templates", "output", filename)

    await callback.message.answer(
        "🔄 <b>Конвертирую в PDF...</b>", parse_mode='HTML')

    # Конвертируем в PDF
    pdf_path = ppt_service.convert_to_pdf(pptx_path)

    if pdf_path and os.path.exists(pdf_path):
        file = FSInputFile(pdf_path)
        await callback.message.answer_document(
            document=file,
            caption="📄 <b>PDF версия коммерческого предложения готова!</b>",
            parse_mode='HTML'
        )

        # Удаляем PDF и презентацию после отправки
        os.remove(pdf_path)
        os.remove(pptx_path)
    else:
        await callback.message.answer(
            "❌ <b>Ошибка при конвертации в PDF</b>\n"
            "Попробуйте позже или обратитесь к администратору.",
            parse_mode='HTML'
        )

    await callback.answer()
