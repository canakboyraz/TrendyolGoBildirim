"""
Telegram Bot Komut Handler'ı
/durum   — Restoran durumu + bekleyen siparişler
/rapor   — Anlık Telegram özet raporu
/excel   — Anlık Excel raporu
/siparisler — Aktif siparişleri listele
/ac      — Restoranı aç
/kapat   — Restoranı kapat
/yardim  — Komut listesi
"""
import requests
import threading
from datetime import datetime
from dotenv import load_dotenv
load_dotenv()

import pytz
from config import TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
from trendyolgo_client import get_stores, get_active_orders, update_store_status
from telegram_notifier import send_message, format_order_message
from daily_report import send_daily_report
from excel_report import generate_daily

TURKEY_TZ  = pytz.timezone("Europe/Istanbul")
BASE_URL   = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}"
_last_update_id = 0   # işlenen son update ID


def _get_updates(offset: int) -> list:
    try:
        r = requests.get(f"{BASE_URL}/getUpdates",
                         params={"offset": offset, "timeout": 5, "limit": 10},
                         timeout=10)
        r.raise_for_status()
        return r.json().get("result", [])
    except Exception:
        return []


def _reply(text: str, chat_id: str = None):
    send_message(text) if not chat_id else _send(chat_id, text)


def _send(chat_id, text):
    try:
        requests.post(f"{BASE_URL}/sendMessage",
                      json={"chat_id": chat_id, "text": text, "parse_mode": "HTML"},
                      timeout=10)
    except Exception:
        pass


def _authorize(chat_id) -> bool:
    """Yalnızca yetkili chat_id'den gelen komutları kabul et."""
    return str(chat_id) == str(TELEGRAM_CHAT_ID)


# ── Komut işleyiciler ─────────────────────────────────────────────────────────
def cmd_durum(chat_id):
    stores  = get_stores()
    active  = get_active_orders()
    now_str = datetime.now(TURKEY_TZ).strftime("%H:%M")

    if not stores:
        _send(chat_id, "⚠️ Restoran bilgisi alınamadı.")
        return

    lines = [f"🕐 <b>Saat:</b> {now_str}\n"]
    for s in stores:
        status_emoji = "🟢" if s.get("workingStatus") == "OPEN" else "🔴"
        lines.append(
            f"{status_emoji} <b>{s.get('name','?')}</b>\n"
            f"   Durum: {s.get('workingStatus','?')}\n"
            f"   Ortalama hazırlık: {s.get('averageOrderPreparationTimeInMin','?')} dk"
        )

    lines.append(f"\n📦 <b>Bekleyen sipariş:</b> {len(active)} adet")
    _send(chat_id, "\n".join(lines))


def cmd_siparisler(chat_id):
    active = get_active_orders()
    if not active:
        _send(chat_id, "📭 Şu an bekleyen sipariş yok.")
        return
    _send(chat_id, f"📦 <b>{len(active)} aktif sipariş:</b>")
    for order in active:
        _send(chat_id, format_order_message(order))


def cmd_rapor(chat_id):
    _send(chat_id, "📊 Rapor hazırlanıyor...")
    # Ayrı thread — bloke etmemesi için
    threading.Thread(target=send_daily_report, daemon=True).start()


def cmd_excel(chat_id):
    _send(chat_id, "📎 Excel hazırlanıyor, birkaç saniye...")
    threading.Thread(target=generate_daily, daemon=True).start()


def cmd_ac(chat_id):
    stores = get_stores()
    if not stores:
        _send(chat_id, "⚠️ Restoran bilgisi alınamadı."); return
    msgs = []
    for s in stores:
        ok, msg = update_store_status(s["id"], "OPEN")
        msgs.append(f"{'✅' if ok else '❌'} {s.get('name','?')}: {msg}")
    _send(chat_id, "\n".join(msgs))


def cmd_kapat(chat_id):
    stores = get_stores()
    if not stores:
        _send(chat_id, "⚠️ Restoran bilgisi alınamadı."); return
    msgs = []
    for s in stores:
        ok, msg = update_store_status(s["id"], "CLOSED")
        msgs.append(f"{'✅' if ok else '❌'} {s.get('name','?')}: {msg}")
    _send(chat_id, "\n".join(msgs))


def cmd_yardim(chat_id):
    _send(chat_id,
        "🤖 <b>Kullanılabilir Komutlar</b>\n\n"
        "/durum — Restoran durumu ve bekleyen siparişler\n"
        "/siparisler — Aktif siparişleri listele\n"
        "/rapor — Anlık özet raporu gönder\n"
        "/excel — Anlık Excel dosyası gönder\n"
        "/ac — Restoranı aç\n"
        "/kapat — Restoranı kapat\n"
        "/yardim — Bu listeyi göster"
    )


# ── Komut yönlendirici ────────────────────────────────────────────────────────
COMMANDS = {
    "/durum":      cmd_durum,
    "/siparisler": cmd_siparisler,
    "/rapor":      cmd_rapor,
    "/excel":      cmd_excel,
    "/ac":         cmd_ac,
    "/kapat":      cmd_kapat,
    "/yardim":     cmd_yardim,
    "/start":      cmd_yardim,
}


def process_updates():
    """
    Gelen Telegram mesajlarını kontrol eder ve komutları çalıştırır.
    main.py'deki ana döngüden her turda çağrılır.
    """
    global _last_update_id
    updates = _get_updates(_last_update_id + 1)

    for update in updates:
        _last_update_id = update["update_id"]
        msg = update.get("message", {})
        if not msg:
            continue

        chat_id = str(msg.get("chat", {}).get("id", ""))
        text    = msg.get("text", "").strip()

        if not text.startswith("/"):
            continue

        # /komut@botismi formatını temizle
        command = text.split()[0].split("@")[0].lower()

        if not _authorize(chat_id):
            _send(chat_id, "⛔ Yetkisiz erişim.")
            continue

        handler = COMMANDS.get(command)
        if handler:
            print(f"[BOT] Komut alındı: {command}")
            threading.Thread(target=handler, args=(chat_id,), daemon=True).start()
        else:
            _send(chat_id, f"❓ Bilinmeyen komut: {command}\n/yardim ile listeye bakabilirsin.")
