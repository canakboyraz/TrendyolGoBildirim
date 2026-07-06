"""
Günlük sipariş özet raporu.
Bugün gelen tüm siparişleri çeker, ürün bazında özetler ve Telegram'a gönderir.
"""

import base64
import requests
import pytz
from dotenv import load_dotenv
load_dotenv()
from datetime import datetime, time
from collections import defaultdict
from config import SUPPLIER_ID, API_KEY, API_SECRET, INTEGRATOR_NAME, API_BASE_URL
from telegram_notifier import send_message


def _get_headers() -> dict:
    credentials = f"{API_KEY}:{API_SECRET}"
    encoded = base64.b64encode(credentials.encode()).decode()
    return {
        "Authorization": f"Basic {encoded}",
        "User-Agent": f"{SUPPLIER_ID} - {INTEGRATOR_NAME}",
        "x-agentname": INTEGRATOR_NAME,
        "x-executor-user": "integration@selfservice.com",
        "Content-Type": "application/json",
    }


TURKEY_TZ = pytz.timezone("Europe/Istanbul")


def get_today_range_ms() -> tuple:
    """Bugünün 00:00 - 23:59 aralığını Türkiye saatiyle epoch milliseconds olarak döner."""
    today = datetime.now(TURKEY_TZ).date()
    start = TURKEY_TZ.localize(datetime.combine(today, time.min))
    end   = TURKEY_TZ.localize(datetime.combine(today, time.max))
    start_ms = int(start.timestamp() * 1000)
    end_ms   = int(end.timestamp()   * 1000)
    return start_ms, end_ms


def fetch_all_orders_today() -> list:
    """Bugünkü tüm siparişleri (tüm statüler) sayfalayarak çeker."""
    start_ms, end_ms = get_today_range_ms()
    all_orders = []
    page = 0
    all_statuses = "Created,Picking,Invoiced,Shipped,Delivered,Cancelled,UnSupplied"

    while True:
        url = f"{API_BASE_URL}/integrator/order/meal/suppliers/{SUPPLIER_ID}/packages"
        params = {
            "packageStatuses": all_statuses,
            "packageModificationStartDate": start_ms,
            "packageModificationEndDate": end_ms,
            "page": page,
            "size": 50,
        }
        try:
            r = requests.get(url, headers=_get_headers(), params=params, timeout=15)
            r.raise_for_status()
            data = r.json()
        except requests.exceptions.RequestException as e:
            print(f"[RAPOR HATA] Sipariş çekilemedi (sayfa {page}): {e}")
            break

        content = data.get("content", [])
        all_orders.extend(content)

        total_pages = data.get("totalPages", 1)
        if page >= total_pages - 1:
            break
        page += 1

    return all_orders


def build_report(orders: list) -> str:
    """Sipariş listesinden Telegram mesajı oluşturur."""
    today_str = datetime.now(TURKEY_TZ).strftime("%d.%m.%Y")

    if not orders:
        return (
            f"📊 <b>Günlük Sipariş Raporu — {today_str}</b>\n\n"
            "Bugün hiç sipariş alınmadı."
        )

    # İptal / geçerli ayrımı
    cancelled_statuses = {"Cancelled", "UnSupplied"}
    active_orders   = [o for o in orders if o.get("packageStatus") not in cancelled_statuses]
    cancelled_orders = [o for o in orders if o.get("packageStatus") in cancelled_statuses]

    # Toplam tutar (iptal olmayanlar)
    total_revenue = sum(o.get("totalPrice", 0) for o in active_orders)

    # Ürün bazında adet ve tutar sayacı
    product_counts  = defaultdict(int)
    product_revenue = defaultdict(float)

    for order in active_orders:
        for line in order.get("lines", []):
            name     = line.get("name", "Bilinmeyen Ürün")
            price    = line.get("price", 0.0)
            quantity = len(line.get("items", [])) or 1
            product_counts[name]  += quantity
            product_revenue[name] += price * quantity

    # Ürünleri adete göre sıralı listele
    sorted_products = sorted(product_counts.items(), key=lambda x: x[1], reverse=True)

    products_text = ""
    for name, qty in sorted_products:
        rev = product_revenue[name]
        products_text += f"  • {name}: <b>{qty} adet</b> — {rev:.2f} ₺\n"

    # Ödeme tipi dağılımı
    payment_counts = defaultdict(int)
    payment_map = {
        "PAY_WITH_CARD":        "💳 Online Kart",
        "PAY_WITH_ON_DELIVERY": "🚪 Kapıda Ödeme",
        "PAY_WITH_MEAL_CARD":   "🍽️ Yemek Kartı",
    }
    for order in active_orders:
        raw = order.get("payment", {}).get("paymentType", "Bilinmiyor")
        label = payment_map.get(raw, raw)
        payment_counts[label] += 1

    payment_text = ""
    for label, cnt in sorted(payment_counts.items(), key=lambda x: x[1], reverse=True):
        payment_text += f"  • {label}: {cnt} sipariş\n"

    # Uygulama kaynağı dağılımı
    app_map = {
        "Trendyol":   "Trendyol Uygulaması",
        "TrendyolGo": "Trendyol Go",
        "Galaxy":     "Getir Yemek by Uber Eats",
    }
    app_counts = defaultdict(int)
    for order in active_orders:
        raw = order.get("userInformation", {}).get("appName", "Bilinmiyor")
        label = app_map.get(raw, raw)
        app_counts[label] += 1

    app_text = ""
    for label, cnt in sorted(app_counts.items(), key=lambda x: x[1], reverse=True):
        app_text += f"  • {label}: {cnt} sipariş\n"

    message = (
        f"📊 <b>Günlük Sipariş Raporu — {today_str}</b>\n"
        f"{'━' * 30}\n"
        f"✅ <b>Geçerli Sipariş:</b> {len(active_orders)} adet\n"
        f"❌ <b>İptal Sipariş:</b> {len(cancelled_orders)} adet\n"
        f"💰 <b>Toplam Ciro:</b> {total_revenue:.2f} ₺\n"
        f"{'━' * 30}\n"
        f"🛍️ <b>Ürün Bazında Özet:</b>\n{products_text}"
        f"{'━' * 30}\n"
        f"💳 <b>Ödeme Yöntemleri:</b>\n{payment_text}"
        f"{'━' * 30}\n"
        f"📱 <b>Sipariş Kaynakları:</b>\n{app_text}"
        f"{'━' * 30}\n"
        f"🕙 <i>Rapor saati: {datetime.now(TURKEY_TZ).strftime('%H:%M')} (TR)</i>"
    )

    return message


def send_daily_report():
    """Günlük raporu oluşturup Telegram'a gönderir."""
    print(f"[RAPOR] Günlük rapor hazırlanıyor...")
    orders = fetch_all_orders_today()
    print(f"[RAPOR] Toplam {len(orders)} sipariş çekildi.")
    message = build_report(orders)
    success = send_message(message)
    if success:
        print(f"[RAPOR] ✅ Günlük rapor Telegram'a gönderildi.")
    else:
        print(f"[RAPOR] ❌ Rapor gönderilemedi.")


if __name__ == "__main__":
    # Doğrudan çalıştırılırsa hemen rapor gönder (test için)
    send_daily_report()
