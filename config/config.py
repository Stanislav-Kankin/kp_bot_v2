import os
from dotenv import load_dotenv

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
LONG_TEMPLATE_ID = os.getenv("LONG_TEMPLATE_ID")
SHORT_TEMPLATE_ID = os.getenv("SHORT_TEMPLATE_ID")

# Настройки Google API
SERVICE_ACCOUNT_FILE = 'credentials/service_account.json'
SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive'
]
