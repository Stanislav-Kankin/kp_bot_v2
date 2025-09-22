from aiogram import Router, types, Bot
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup
from bot.states import FormKP
from bot.google_slides_service import slides_service
from datetime import datetime
import re

router = Router()

@router.message(Command("start"))
async def start(message: types.Message):
    await message.answer(
        "🤖 <b>Бот для создания коммерческих предложений</b>\n\n"
        "Я помогу создать профессиональное КП в формате Google Slides.\n\n"
        "▶️ <b>Нажмите /kp чтобы начать</b>",
        parse_mode='HTML'
    )

@router.message(Command("kp"))
async def start_kp(message: types.Message, state: FSMContext):
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(
            text="📊 Длинный вариант КП", 
            callback_data="template_long"
        )],
        [InlineKeyboardButton(
            text="📈 Короткий вариант КП", 
            callback_data="template_short"
        )]
    ])
    
    await message.answer(
        "🎯 <b>Выберите тип коммерческого предложения:</b>\n\n"
        "• <b>Длинный вариант</b> - детальное предложение со всеми услугами\n"
        "• <b>Короткий вариант</b> - краткое предложение с основными пунктами",
        reply_markup=keyboard,
        parse_mode='HTML'
    )
    await state.set_state(FormKP.template_type)

@router.callback_query(FormKP.template_type)
async def process_template_choice(callback: types.CallbackQuery, state: FSMContext):
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
    await state.set_state(FormKP.contact_person)
    
    await message.answer(
        "👤 <b>Введите ФИО контактного лица:</b>",
        parse_mode='HTML'
    )

@router.message(FormKP.contact_person)
async def process_contact_person(message: types.Message, state: FSMContext):
    await state.update_data(contact_person=message.text)
    await state.set_state(FormKP.phone)
    
    await message.answer(
        "📞 <b>Введите номер телефона:</b>",
        parse_mode='HTML'
    )

@router.message(FormKP.phone)
async def process_phone(message: types.Message, state: FSMContext):
    phone = message.text.strip()
    if not re.match(r'^[\d\s\-\+\(\)]+$', phone):
        await message.answer("❌ Пожалуйста, введите корректный номер телефона:")
        return
    
    await state.update_data(phone=phone)
    await state.set_state(FormKP.email)
    
    await message.answer(
        "📧 <b>Введите email:</b>",
        parse_mode='HTML'
    )

@router.message(FormKP.email)
async def process_email(message: types.Message, state: FSMContext):
    email = message.text.strip()
    if not re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email):
        await message.answer("❌ Пожалуйста, введите корректный email:")
        return
    
    await state.update_data(email=email)
    await state.set_state(FormKP.employees_count)
    
    await message.answer(
        "👥 <b>Введите количество сотрудников в компании:</b>",
        parse_mode='HTML'
    )

@router.message(FormKP.employees_count)
async def process_employees_count(message: types.Message, state: FSMContext):
    try:
        count = int(message.text.strip())
        if count <= 0:
            raise ValueError
    except ValueError:
        await message.answer("❌ Пожалуйста, введите корректное число сотрудников:")
        return
    
    await state.update_data(employees_count=count)
    await state.set_state(FormKP.service_type)
    
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="HR-автоматизация", callback_data="service_hr")],
        [InlineKeyboardButton(text="Документооборот", callback_data="service_docs")],
        [InlineKeyboardButton(text="Оба решения", callback_data="service_both")]
    ])
    
    await message.answer(
        "🛠️ <b>Выберите тип услуги:</b>",
        reply_markup=keyboard,
        parse_mode='HTML'
    )

@router.callback_query(FormKP.service_type)
async def process_service_type(callback: types.CallbackQuery, state: FSMContext):
    service_map = {
        'service_hr': 'HR-автоматизация',
        'service_docs': 'Документооборот', 
        'service_both': 'HR-автоматизация + Документооборот'
    }
    
    service_type = service_map.get(callback.data, 'Не указано')
    await state.update_data(service_type=service_type)
    await state.set_state(FormKP.implementation_period)
    
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="1 месяц", callback_data="period_1")],
        [InlineKeyboardButton(text="2-3 месяца", callback_data="period_2_3")],
        [InlineKeyboardButton(text="4-6 месяцев", callback_data="period_4_6")]
    ])
    
    await callback.message.answer(
        "⏱️ <b>Выберите срок внедрения:</b>",
        reply_markup=keyboard,
        parse_mode='HTML'
    )
    await callback.answer()

@router.callback_query(FormKP.implementation_period)
async def process_implementation_period(callback: types.CallbackQuery, state: FSMContext):
    period_map = {
        'period_1': '1 месяц',
        'period_2_3': '2-3 месяца',
        'period_4_6': '4-6 месяцев'
    }
    
    period = period_map.get(callback.data, 'Не указано')
    await state.update_data(implementation_period=period)
    await state.set_state(FormKP.additional_services)
    
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Обучение сотрудников", callback_data="add_training")],
        [InlineKeyboardButton(text="Техническая поддержка", callback_data="add_support")],
        [InlineKeyboardButton(text="Интеграция с 1С", callback_data="add_1c")],
        [InlineKeyboardButton(text="Пропустить", callback_data="add_skip")]
    ])
    
    await callback.message.answer(
        "🎁 <b>Выберите дополнительные услуги (можно несколько):</b>\n"
        "Нажмите кнопки по очереди, затем 'Готово'",
        reply_markup=keyboard,
        parse_mode='HTML'
    )
    await callback.answer()

@router.callback_query(FormKP.additional_services)
async def process_additional_services(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    current_services = data.get('additional_services', [])
    
    service_map = {
        'add_training': 'Обучение сотрудников',
        'add_support': 'Техническая поддержка', 
        'add_1c': 'Интеграция с 1С'
    }
    
    if callback.data == 'add_skip':
        services_text = 'Не требуются' if not current_services else ', '.join(current_services)
        await state.update_data(additional_services=services_text)
        await state.set_state(FormKP.total_budget)
        
        await callback.message.answer(
            "💰 <b>Введите общий бюджет проекта (в рублях):</b>",
            parse_mode='HTML'
        )
    else:
        service = service_map.get(callback.data)
        if service and service not in current_services:
            current_services.append(service)
            await state.update_data(additional_services=current_services)
            
            keyboard = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="Обучение сотрудников", callback_data="add_training")],
                [InlineKeyboardButton(text="Техническая поддержка", callback_data="add_support")],
                [InlineKeyboardButton(text="Интеграция с 1С", callback_data="add_1c")],
                [InlineKeyboardButton(text="✅ Готово", callback_data="add_skip")]
            ])
            
            await callback.message.edit_text(
                f"🎁 <b>Выбранные услуги:</b>\n{', '.join(current_services)}\n\n"
                "Продолжайте выбирать или нажмите 'Готово':",
                reply_markup=keyboard,
                parse_mode='HTML'
            )
    
    await callback.answer()

@router.message(FormKP.total_budget)
async def process_total_budget(message: types.Message, state: FSMContext):
    try:
        budget = int(message.text.strip().replace(' ', '').replace(',', ''))
        if budget <= 0:
            raise ValueError
    except ValueError:
        await message.answer("❌ Пожалуйста, введите корректную сумму бюджета:")
        return
    
    await state.update_data(total_budget=budget)
    
    # Получаем все данные
    data = await state.get_data()
    data['current_date'] = datetime.now().strftime("%d.%m.%Y")
    
    # Создаем презентацию
    await message.answer("🔄 <b>Создаю коммерческое предложение...</b>", parse_mode='HTML')
    
    presentation_url = slides_service.create_kp_presentation(
        data['template_type'], 
        data
    )
    
    if presentation_url:
        await message.answer(
            f"✅ <b>Коммерческое предложение готово!</b>\n\n"
            f"🔗 <b>Ссылка для редактирования:</b>\n{presentation_url}\n\n"
            f"📊 <b>Данные КП:</b>\n"
            f"• Компания: {data['company_name']}\n"
            f"• Контакт: {data['contact_person']}\n"
            f"• Сотрудников: {data['employees_count']}\n"
            f"• Услуга: {data['service_type']}\n"
            f"• Бюджет: {data['total_budget']:,} ₽\n\n"
            f"<i>Для создания нового КП нажмите /kp</i>",
            parse_mode='HTML'
        )
    else:
        await message.answer(
            "❌ <b>Ошибка при создании презентации</b>\n"
            "Попробуйте позже или обратитесь к администратору.",
            parse_mode='HTML'
        )
    
    await state.clear()