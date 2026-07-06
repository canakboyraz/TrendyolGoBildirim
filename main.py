"""
Trendyol Go Sipariş Bildirim Servisi
=====================================
- Her 30 saniyede yeni sipariş kontrolü + Telegram bildirimi
- Telegram bot komutları (/durum, /rapor, /excel, /ac, /kapat ...)
- Her gece 23:45 → günlük özet mesajı + Excel (dün ile karşılaştırmalı)
- Her Pazartesi 08:00 → haftalık Excel
- Ayın 1'i 08:00 → aylık Excel
"""
import time, signal, sys
from datetime import datetime
import pytz
from dotenv import load_dotenv
load_dotenv()

from config import POLL_INTERVAL_SECONDS, SUPPLIER_ID
from trendyolgo_client import get_new_orders
from trendyolgo_client import get_stores
from telegram_notifier import send_message, format_order_message
from daily_report import send_daily_report
from excel_report import generate_daily, generate_weekly, generate_monthly
from database import upsert_order, is_notified, mark_notified, get_all_order_ids, init_db
from bot_commands import process_updates

TURKEY_TZ = pytz.timezone("Europe/Istanbul")

running = True
DAILY_H, DAILY_M     = 23, 45   # Günlük rapor saati
WEEKLY_H, WEEKLY_M   =  8,  0   # Haftalık rapor saati (Pazartesi)
MONTHLY_H, MONTHLY_M =  8,  0   # Aylık rapor saati (Ayın 1'i)

sent_today:   str = ""
sent_weekly:  str = ""
sent_monthly: str = ""


def handle_shutdown(signum, frame):
    global running
    print("\n[BİLGİ] Servis durduruluyor...")
    running = False


def now_str() -> str:
    return datetime.now(TURKEY_TZ).strftime("%H:%M:%S")


def check_and_notify():
    """Yeni siparişleri kontrol eder, DB'ye kaydeder, bildirim gönderir."""
    orders = get_new_orders()
    new_count = 0
    for order in orders:
        order_id = order.get("id")
        if not order_id:
            continue
        upsert_order(order)               # her durumda DB'ye kaydet / güncelle
        if not is_notified(order_id):
            msg = format_order_message(order)
            if send_message(msg):
                mark_notified(order_id)
                print(f"[{now_str()}] ✅ Bildirim → #{order.get('orderNumber')} ({order.get('packageStatus')})")
            else:
                print(f"[{now_str()}] ❌ Bildirim gönderilemedi → {order_id}")
            new_count += 1

    if new_count == 0:
        notified_count = len(get_all_order_ids())
        print(f"[{now_str()}] 🔍 Yeni sipariş yok. (DB'de toplam: {notified_count})")


def check_scheduled_reports():
    """Zamanlı raporları kontrol eder."""
    global sent_today, sent_weekly, sent_monthly
    now = datetime.now(TURKEY_TZ)
    today = now.strftime("%Y-%m-%d")

    # Günlük — her gece 23:45
    if now.hour == DAILY_H and now.minute == DAILY_M and sent_today != today:
        sent_today = today
        print(f"[{now_str()}] 📊 Günlük rapor gönderiliyor...")
        send_daily_report()
        generate_daily()

    # Haftalık — her Pazartesi 08:00
    week_key = f"{now.isocalendar()[1]}-{now.year}"
    if now.weekday() == 0 and now.hour == WEEKLY_H and now.minute == WEEKLY_M and sent_weekly != week_key:
        sent_weekly = week_key
        print(f"[{now_str()}] 📅 Haftalık rapor gönderiliyor...")
        generate_weekly()

    # Aylık — ayın 1'i 08:00
    month_key = now.strftime("%Y-%m")
    if now.day == 1 and now.hour == MONTHLY_H and now.minute == MONTHLY_M and sent_monthly != month_key:
        sent_monthly = month_key
        print(f"[{now_str()}] 🗓️ Aylık rapor gönderiliyor...")
        generate_monthly()


def startup_check() -> bool:
    print("=" * 50)
    print("  Trendyol Go Sipariş Bildirim Servisi")
    print("=" * 50)
    print(f"  Supplier ID      : {SUPPLIER_ID}")
    print(f"  Kontrol aralığı  : her {POLL_INTERVAL_SECONDS} saniye")
    print(f"  Günlük rapor     : 23:45 (TR)")
    print(f"  Haftalık rapor   : Pazartesi 08:00 (TR)")
    print(f"  Aylık rapor      : Ayın 1'i 08:00 (TR)")
    print("=" * 50)

    print("\n[BAŞLANGIÇ] DB başlatılıyor...")
    init_db()
    print("[BAŞLANGIÇ] ✅ Veritabanı hazır.")

    print("[BAŞLANGIÇ] Trendyol Go API test ediliyor...")
    stores = get_stores()
    if stores is not None:
        print(f"[BAŞLANGIÇ] ✅ API bağlantısı başarılı. {len(stores)} restoran.")
        for s in stores:
            print(f"           → {s.get('name')} (ID:{s.get('id')}) — {s.get('workingStatus')}")
    else:
        print("[BAŞLANGIÇ] ⚠️  API test edilemedi.")

    print("[BAŞLANGIÇ] Telegram test ediliyor...")
    msg = (
        "✅ <b>Trendyol Go Bildirim Servisi Başladı!</b>\n\n"
        f"🏪 Supplier ID: <code>{SUPPLIER_ID}</code>\n"
        f"⏱️ Kontrol: her <b>{POLL_INTERVAL_SECONDS} sn</b>\n\n"
        "Komutlar için /yardim yaz. 🚀"
    )
    if send_message(msg):
        print("[BAŞLANGIÇ] ✅ Telegram hazır.")
        return True
    else:
        print("[BAŞLANGIÇ] ❌ Telegram mesajı gönderilemedi.")
        return False


def main():
    global running
    signal.signal(signal.SIGINT,  handle_shutdown)
    signal.signal(signal.SIGTERM, handle_shutdown)

    if not startup_check():
        sys.exit(1)

    print(f"\n[{now_str()}] Servis çalışıyor...\n")

    while running:
        try:
            check_and_notify()
            check_scheduled_reports()
            process_updates()          # Telegram komutlarını işle
        except Exception as e:
            print(f"[{now_str()}] ❌ Beklenmeyen hata: {e}")

        for _ in range(POLL_INTERVAL_SECONDS):
            if not running: break
            time.sleep(1)

    print(f"[{now_str()}] Servis durduruldu.")


if __name__ == "__main__":
    main()
