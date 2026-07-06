import requests
from config import TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID


TELEGRAM_API_URL = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"


def send_message(text: str) -> bool:
    """Telegram botuna mesaj gönderir. Başarılıysa True döner."""
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
    }
    try:
        response = requests.post(TELEGRAM_API_URL, json=payload, timeout=10)
        response.raise_for_status()
        return True
    except requests.exceptions.RequestException as e:
        print(f"[TELEGRAM HATA] Mesaj gönderilemedi: {e}")
        return False


def format_order_message(order: dict) -> str:
    """Sipariş verisini okunabilir bir Telegram mesajına dönüştürür."""
    order_number = order.get("orderNumber", "N/A")
    order_code = order.get("orderCode", "N/A")
    total_price = order.get("totalPrice", 0)
    delivery_type = order.get("deliveryType", "N/A")
    eta = order.get("eta", "N/A")
    customer_note = order.get("customerNote", "")
    store_id = order.get("storeId", "N/A")
    app_name = order.get("userInformation", {}).get("appName", "N/A")

    # Ödeme bilgisi
    payment = order.get("payment", {})
    payment_type_raw = payment.get("paymentType", "")
    payment_map = {
        "PAY_WITH_CARD": "💳 Online Kart",
        "PAY_WITH_ON_DELIVERY": "🚪 Kapıda Ödeme",
        "PAY_WITH_MEAL_CARD": "🍽️ Yemek Kartı",
    }
    payment_text = payment_map.get(payment_type_raw, payment_type_raw)

    if payment_type_raw == "PAY_WITH_ON_DELIVERY":
        on_delivery = payment.get("onDelivery", {})
        sub_type = on_delivery.get("paymentType", "")
        sub_map = {
            "CASH": "Nakit",
            "CARD": "Kredi Kartı",
        }
        sub_text = sub_map.get(sub_type, sub_type)
        if sub_text:
            payment_text += f" ({sub_text})"

    # Teslimat tipi
    delivery_map = {
        "GO": "🛵 Trendyol Go Kuryesi",
        "STORE": "🏪 Restoran Kuryesi",
    }
    delivery_text = delivery_map.get(delivery_type, delivery_type)

    # Ürün listesi
    lines = order.get("lines", [])
    items_text = ""
    for line in lines:
        name = line.get("name", "Ürün")
        price = line.get("price", 0)
        quantity = len(line.get("items", [])) or 1
        items_text += f"  • {name} x{quantity} — {price:.2f} ₺\n"

        # Modifier ürünler (ekstra seçenekler)
        for modifier in line.get("modifierProducts", []):
            mod_name = modifier.get("name", "")
            items_text += f"    ↳ {mod_name}\n"

    if not items_text:
        items_text = "  (Ürün bilgisi yok)\n"

    # Uygulama kaynağı
    app_map = {
        "Trendyol": "Trendyol Uygulaması",
        "TrendyolGo": "Trendyol Go",
        "Galaxy": "Getir Yemek by Uber Eats",
    }
    app_text = app_map.get(app_name, app_name)

    # Gel-al sipariş kontrolü
    store_pickup = order.get("storePickupSelected", False)
    pickup_text = "🏃 Gel-Al Sipariş" if store_pickup else ""

    message = (
        f"🆕 <b>YENİ SİPARİŞ GELDİ!</b>\n"
        f"{'━' * 28}\n"
        f"📋 <b>Sipariş No:</b> #{order_number}\n"
        f"🔑 <b>Kod:</b> {order_code}\n"
        f"🏪 <b>Şube ID:</b> {store_id}\n"
        f"📱 <b>Kaynak:</b> {app_text}\n"
        f"{'━' * 28}\n"
        f"🛍️ <b>Ürünler:</b>\n{items_text}"
        f"{'━' * 28}\n"
        f"💰 <b>Toplam:</b> {total_price:.2f} ₺\n"
        f"💳 <b>Ödeme:</b> {payment_text}\n"
        f"🚀 <b>Teslimat:</b> {delivery_text}\n"
        f"⏱️ <b>Tahmini Süre:</b> {eta}\n"
    )

    if pickup_text:
        message += f"{pickup_text}\n"

    if customer_note:
        message += f"📝 <b>Müşteri Notu:</b> {customer_note}\n"

    return message
