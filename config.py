import os

# Trendyol Go Entegrasyon Bilgileri
SUPPLIER_ID      = os.environ["SUPPLIER_ID"]
API_KEY          = os.environ["API_KEY"]
API_SECRET       = os.environ["API_SECRET"]
INTEGRATOR_NAME  = os.environ.get("INTEGRATOR_NAME", "SelfIntegration")

# Telegram Bilgileri
TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHAT_ID   = os.environ["TELEGRAM_CHAT_ID"]

# Uygulama Ayarları
POLL_INTERVAL_SECONDS = int(os.environ.get("POLL_INTERVAL_SECONDS", "30"))
API_BASE_URL = "https://api.tgoapis.com"
