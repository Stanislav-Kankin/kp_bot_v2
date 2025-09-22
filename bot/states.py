from aiogram.fsm.state import State, StatesGroup


class FormKP(StatesGroup):
    template_type = State()
    company_name = State()
    contact_person = State()
    phone = State()
    email = State()
    employees_count = State()
    service_type = State()
    implementation_period = State()
    additional_services = State()
    total_budget = State()
