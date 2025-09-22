from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from config.config import SERVICE_ACCOUNT_FILE, SCOPES, LONG_TEMPLATE_ID, SHORT_TEMPLATE_ID


class GoogleSlidesService:
    def __init__(self):
        self.credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES
        )
        self.slides_service = build('slides', 'v1', credentials=self.credentials)
        self.drive_service = build('drive', 'v3', credentials=self.credentials)

    def create_presentation_copy(self, template_id, new_title):
        """Создает копию презентации"""
        try:
            # Создаем копию файла
            body = {'name': new_title}
            copied_file = self.drive_service.files().copy(
                fileId=template_id, body=body
            ).execute()

            return copied_file['id']
        except HttpError as error:
            print(f"Ошибка при создании копии: {error}")
            return None

    def replace_text_in_slide(self, presentation_id, replacements):
        """Заменяет текст в презентации"""
        try:
            requests = []

            for placeholder, new_text in replacements.items():
                requests.append({
                    'replaceAllText': {
                        'containsText': {'text': '{{' + placeholder + '}}'},
                        'replaceText': str(new_text)
                    }
                })

            # Добавляем запросы для замены одиночных фигурных скобок
            for placeholder, new_text in replacements.items():
                requests.append({
                    'replaceAllText': {
                        'containsText': {'text': '{' + placeholder + '}'},
                        'replaceText': str(new_text)
                    }
                })

            body = {'requests': requests}
            result = self.slides_service.presentations().batchUpdate(
                presentationId=presentation_id, body=body
            ).execute()

            return result
        except HttpError as error:
            print(f"Ошибка при замене текста: {error}")
            return None

    def create_kp_presentation(self, template_type, kp_data):
        """Создает КП на основе выбранного шаблона"""
        template_id = LONG_TEMPLATE_ID if template_type == 'long' else SHORT_TEMPLATE_ID

        # Создаем название файла
        new_title = f"КП для {kp_data['company_name']} - {kp_data['contact_person']}"

        # Создаем копию презентации
        presentation_id = self.create_presentation_copy(template_id, new_title)
        if not presentation_id:
            return None

        # Подготавливаем данные для замены
        replacements = {
            'COMPANY_NAME': kp_data.get('company_name', ''),
            'CONTACT_PERSON': kp_data.get('contact_person', ''),
            'PHONE': kp_data.get('phone', ''),
            'EMAIL': kp_data.get('email', ''),
            'EMPLOYEES_COUNT': kp_data.get('employees_count', ''),
            'SERVICE_TYPE': kp_data.get('service_type', ''),
            'IMPLEMENTATION_PERIOD': kp_data.get('implementation_period', ''),
            'ADDITIONAL_SERVICES': kp_data.get('additional_services', ''),
            'TOTAL_BUDGET': kp_data.get('total_budget', ''),
            'CURRENT_DATE': kp_data.get('current_date', '')
        }

        # Заменяем текст в презентации
        self.replace_text_in_slide(presentation_id, replacements)

        # Создаем ссылку для просмотра
        presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"

        return presentation_url


# Создаем глобальный экземпляр сервиса
slides_service = GoogleSlidesService()
