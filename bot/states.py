from aiogram.fsm.state import State, StatesGroup


class FormKP(StatesGroup):
    template_type = State()
    company_name = State()
    hr_licenses = State()
    employee_licenses = State()
    on_premises = State()
