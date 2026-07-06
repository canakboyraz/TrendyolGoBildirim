"""
Trendyol Go Sipariş Bildirim Servisi
=====================================
Her 30 saniyede bir yeni siparişleri kontrol eder,
yeni sipariş gelince Telegram üzerinden bildirim gönderir.
"""

import time
import signal
import sys
from datetime import datetime
import pytz
from dotenv import load_dotenv
load_dotenv()  # .env dosyasını yükle (local geliştirme için)

from config import POLL_INTERVAL_SECONDS, SUPPLIER_ID
from trendyolgo_client import get_new_orders, get_stores
from telegram_notifier import send_message, format_order_message
from daily_report import send_daily_report


# Daha önce bildirimi gönderilmiş sipariş ID'lerini tutar
seen_order_ids: set = set()

# Servis çalışma durumu
running = True

# Günlük raporun bugün gönderilip gönderilmediğini takip eder
daily_report_sent_date: str = ""
DAILY_REPORT_HOUR = 23
DAILY_REPORT_MINUTE = 45


def handle_shutdown(signum, frame):
    """CTRL+C veya kill sinyalinde temiz kapanış."""
    global running
    print("\n[BİLGİ] Servis durduruluyor...")
    running = False


def check_daily_report():
    """Her gün 23:45'te bir kez günlük rapor gönderir (Türkiye saati)."""
    global daily_report_sent_date
    now_dt = datetime.now(TURKEY_TZ)
    today_str = now_dt.strftime("%Y-%m-%d")

    if (
        now_dt.hour == DAILY_REPORT_HOUR
        and now_dt.minute == DAILY_REPORT_MINUTE
        and daily_report_sent_date != today_str
    ):
        daily_report_sent_date = today_str
        print(f"[{now()}] 📊 Günlük rapor saati geldi, gönderiliyor...")
        send_daily_report()


def check_and_notify():
    """Yeni siparişleri kontrol eder ve bildirim gönderir."""
    orders = get_new_orders()

    new_count = 0
    for order in orders:
        order_id = order.get("id")
        if not order_id:
            continue

        # Her sipariş ID'si için yalnızca bir kez bildirim gönder
        if order_id not in seen_order_ids:
            seen_order_ids.add(order_id)
            message = format_order_message(order)
            success = send_message(message)
            if success:
                order_number = order.get("orderNumber", "N/A")
                status = order.get("packageStatus", "?")
                print(f"[{now()}] ✅ Bildirim gönderildi → Sipariş #{order_number} (Statü: {status})")
            else:
                print(f"[{now()}] ❌ Bildirim gönderilemedi → Sipariş ID: {order_id}")
            new_count += 1

    if new_count == 0:
        print(f"[{now()}] 🔍 Yeni sipariş yok. (Toplam takip edilen: {len(seen_order_ids)})")


TURKEY_TZ = pytz.timezone("Europe/Istanbul")

def now() -> str:
    return datetime.now(TURKEY_TZ).strftime("%H:%M:%S")


def startup_check() -> bool:
    """Başlangıçta API ve Telegram bağlantısını test eder."""
    print("=" * 50)
    print("  Trendyol Go Sipariş Bildirim Servisi")
    print("=" * 50)
    print(f"  Supplier ID : {SUPPLIER_ID}")
    print(f"  Kontrol aralığı: her {POLL_INTERVAL_SECONDS} saniye")
    print("=" * 50)

    # API bağlantı testi
    print("\n[BAŞLANGIÇ] Trendyol Go API bağlantısı test ediliyor...")
    stores = get_stores()
    if stores is not None:
        print(f"[BAŞLANGIÇ] ✅ API bağlantısı başarılı. {len(stores)} restoran bulundu.")
        for store in stores:
            print(f"           → {store.get('name', '?')} (ID: {store.get('id', '?')}) — {store.get('workingStatus', '?')}")
    else:
        print("[BAŞLANGIÇ] ⚠️  API bağlantısı test edilemedi.")

    # Telegram testi
    print("\n[BAŞLANGIÇ] Telegram bağlantısı test ediliyor...")
    test_msg = (
        "✅ <b>Trendyol Go Bildirim Servisi Başladı!</b>\n\n"
        f"🏪 Supplier ID: <code>{SUPPLIER_ID}</code>\n"
        f"⏱️ Kontrol aralığı: her <b>{POLL_INTERVAL_SECONDS} saniye</b>\n\n"
        "Yeni sipariş geldiğinde buradan bildirim alacaksınız. 🚀"
    )
    if send_message(test_msg):
        print("[BAŞLANGIÇ] ✅ Telegram bağlantısı başarılı. Test mesajı gönderildi.")
        return True
    else:
        print("[BAŞLANGIÇ] ❌ Telegram mesajı gönderilemedi. Token ve Chat ID'yi kontrol edin.")
        return False


def main():
    global running

    # Temiz kapanış için sinyal dinleyicileri
    signal.signal(signal.SIGINT, handle_shutdown)
    signal.signal(signal.SIGTERM, handle_shutdown)

    # Başlangıç kontrolü
    if not startup_check():
        sys.exit(1)

    print(f"\n[{now()}] Servis çalışıyor. Siparişler izleniyor...\n")

    # Ana döngü
    while running:
        try:
            check_and_notify()
            check_daily_report()
        except Exception as e:
            print(f"[{now()}] ❌ Beklenmeyen hata: {e}")

        # Bekleme — interrupt'a duyarlı
        for _ in range(POLL_INTERVAL_SECONDS):
            if not running:
                break
            time.sleep(1)

    print(f"[{now()}] Servis durduruldu.")


if __name__ == "__main__":
    main()
