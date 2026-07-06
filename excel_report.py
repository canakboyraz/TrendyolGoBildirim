"""
Günlük sipariş Excel raporu.
Tüm siparişleri çekip formatlı bir .xlsx dosyası oluşturur ve Telegram'a gönderir.
"""

import os
import requests
from datetime import datetime, date, time as dtime
from collections import defaultdict
import pytz
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv
load_dotenv()

from config import (
    SUPPLIER_ID, API_KEY, API_SECRET, INTEGRATOR_NAME,
    API_BASE_URL, TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
)

import base64

TURKEY_TZ = pytz.timezone("Europe/Istanbul")

# ── Renkler ──────────────────────────────────────────────────────────────────
CLR_HEADER      = "1F3864"   # Koyu lacivert
CLR_HEADER_FONT = "FFFFFF"
CLR_SUBHEADER   = "2E75B6"   # Mavi
CLR_SUBHDR_FONT = "FFFFFF"
CLR_ROW_ODD     = "EBF3FB"
CLR_ROW_EVEN    = "FFFFFF"
CLR_CANCELLED   = "FFD7D7"
CLR_TOTAL_ROW   = "FFF2CC"
CLR_SUMMARY_HDR = "375623"
CLR_SUMMARY_FNT = "FFFFFF"


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


def get_today_range_ms() -> tuple:
    today = datetime.now(TURKEY_TZ).date()
    start = TURKEY_TZ.localize(datetime.combine(today, dtime.min))
    end   = TURKEY_TZ.localize(datetime.combine(today, dtime.max))
    return int(start.timestamp() * 1000), int(end.timestamp() * 1000)


def fetch_all_orders_today() -> list:
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
            print(f"[EXCEL] API hatası: {e}")
            break
        all_orders.extend(data.get("content", []))
        if page >= data.get("totalPages", 1) - 1:
            break
        page += 1
    return all_orders


def _border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)


def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def _font(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, size=size, name="Calibri")


def _align(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def _set_col_width(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width


def _payment_label(order):
    payment = order.get("payment", {})
    raw = payment.get("paymentType", "")
    mapping = {
        "PAY_WITH_CARD":        "Online Kart",
        "PAY_WITH_ON_DELIVERY": "Kapıda Ödeme",
        "PAY_WITH_MEAL_CARD":   "Yemek Kartı",
    }
    label = mapping.get(raw, raw)
    if raw == "PAY_WITH_ON_DELIVERY":
        sub = payment.get("onDelivery", {}) or {}
        sub_map = {"CASH": "Nakit", "CARD": "Kredi Kartı"}
        sub_label = sub_map.get(sub.get("paymentType", ""), "")
        if sub_label:
            label += f" - {sub_label}"
    elif raw == "PAY_WITH_MEAL_CARD":
        card = payment.get("mealCard", {}) or {}
        card_type = card.get("cardSourceType", "")
        if card_type:
            label += f" ({card_type})"
    return label


def _status_label(status):
    return {
        "Created":    "Yeni",
        "Picking":    "Hazırlanıyor",
        "Invoiced":   "Hazır",
        "Shipped":    "Yolda",
        "Delivered":  "Teslim Edildi",
        "Cancelled":  "İptal",
        "UnSupplied": "Restoran İptali",
    }.get(status, status)


def _app_label(app):
    return {
        "Trendyol":   "Trendyol",
        "TrendyolGo": "Trendyol Go",
        "Galaxy":     "Getir Yemek",
    }.get(app, app or "-")


def _delivery_label(d):
    return {"GO": "TGo Kuryesi", "STORE": "Restoran Kuryesi"}.get(d, d)


def _ts_to_str(ts_ms):
    if not ts_ms:
        return "-"
    try:
        dt = datetime.fromtimestamp(ts_ms / 1000, tz=TURKEY_TZ)
        return dt.strftime("%H:%M:%S")
    except Exception:
        return "-"


def build_excel(orders: list, filepath: str):
    wb = openpyxl.Workbook()
    today_str = datetime.now(TURKEY_TZ).strftime("%d.%m.%Y")

    # ═══════════════════════════════════════════════════════════
    # SAYFA 1 — SİPARİŞ LİSTESİ
    # ═══════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "Sipariş Listesi"

    # Başlık satırı
    ws1.merge_cells("A1:N1")
    title_cell = ws1["A1"]
    title_cell.value = f"🛍️  Trendyol Go — Günlük Sipariş Raporu   {today_str}"
    title_cell.font      = _font(bold=True, color=CLR_HEADER_FONT, size=13)
    title_cell.fill      = _fill(CLR_HEADER)
    title_cell.alignment = _align("center")
    ws1.row_dimensions[1].height = 28

    # Kolon başlıkları
    columns = [
        ("Sipariş No",       16),
        ("Kod",               8),
        ("Sipariş Saati",    14),
        ("Statü",            16),
        ("Ürünler",          40),
        ("Adet",              6),
        ("Ürün Tutarı (₺)",  16),
        ("Toplam (₺)",       14),
        ("Ödeme",            22),
        ("Teslimat",         18),
        ("Kaynak",           16),
        ("Müşteri",          16),
        ("Müşteri Notu",     28),
        ("Test mi?",          9),
    ]

    for col_idx, (col_name, col_width) in enumerate(columns, start=1):
        cell = ws1.cell(row=2, column=col_idx, value=col_name)
        cell.font      = _font(bold=True, color=CLR_SUBHDR_FONT, size=10)
        cell.fill      = _fill(CLR_SUBHEADER)
        cell.alignment = _align("center")
        cell.border    = _border()
        _set_col_width(ws1, col_idx, col_width)
    ws1.row_dimensions[2].height = 18

    ws1.freeze_panes = "A3"

    cancelled_statuses = {"Cancelled", "UnSupplied"}
    row = 3

    for order in orders:
        status       = order.get("packageStatus", "")
        is_cancelled = status in cancelled_statuses
        row_fill     = _fill(CLR_CANCELLED) if is_cancelled else \
                       _fill(CLR_ROW_ODD if row % 2 == 1 else CLR_ROW_EVEN)

        order_number  = order.get("orderNumber", "-")
        order_code    = order.get("orderCode", "-")
        created_ts    = order.get("packageCreationDate", 0)
        total_price   = order.get("totalPrice", 0)
        customer      = order.get("customer", {})
        customer_name = f"{customer.get('firstName','')} {customer.get('lastName','')}".strip()
        note          = order.get("customerNote", "") or ""
        app_name      = _app_label(order.get("userInformation", {}).get("appName", ""))
        delivery      = _delivery_label(order.get("deliveryType", ""))
        is_test       = "Evet" if order.get("testPackage") else "Hayır"

        lines = order.get("lines", [])
        product_names  = []
        total_qty      = 0
        total_line_rev = 0.0

        for line in lines:
            name  = line.get("name", "?")
            qty   = len(line.get("items", [])) or 1
            price = line.get("price", 0)
            mods  = [m.get("name", "") for m in line.get("modifierProducts", [])]
            mod_str = f" ({', '.join(mods)})" if mods else ""
            product_names.append(f"{name}{mod_str} x{qty}")
            total_qty      += qty
            total_line_rev += price * qty

        products_str = "\n".join(product_names)

        data_row = [
            order_number,
            order_code,
            _ts_to_str(created_ts),
            _status_label(status),
            products_str,
            total_qty,
            round(total_line_rev, 2),
            round(total_price, 2),
            _payment_label(order),
            delivery,
            app_name,
            customer_name,
            note,
            is_test,
        ]

        for col_idx, value in enumerate(data_row, start=1):
            cell = ws1.cell(row=row, column=col_idx, value=value)
            cell.fill      = row_fill
            cell.border    = _border()
            cell.alignment = _align(
                h="left" if col_idx in (5, 9, 10, 11, 13) else "center",
                wrap=col_idx in (5, 13)
            )
            cell.font = _font(size=9)
            if col_idx in (7, 8):
                cell.number_format = '#,##0.00 ₺'

        # Satır yüksekliği — ürün sayısına göre
        line_count = max(len(product_names), 1)
        ws1.row_dimensions[row].height = max(16, line_count * 14)
        row += 1

    # Toplam satırı
    if orders:
        active_orders = [o for o in orders if o.get("packageStatus") not in cancelled_statuses]
        total_revenue = sum(o.get("totalPrice", 0) for o in active_orders)

        ws1.merge_cells(f"A{row}:G{row}")
        total_label = ws1.cell(row=row, column=1, value="TOPLAM (İptal hariç)")
        total_label.font      = _font(bold=True, size=10)
        total_label.fill      = _fill(CLR_TOTAL_ROW)
        total_label.alignment = _align("right")
        total_label.border    = _border()

        total_val = ws1.cell(row=row, column=8, value=round(total_revenue, 2))
        total_val.font          = _font(bold=True, size=10)
        total_val.fill          = _fill(CLR_TOTAL_ROW)
        total_val.alignment     = _align("center")
        total_val.border        = _border()
        total_val.number_format = '#,##0.00 ₺'

        for col_idx in range(9, 15):
            c = ws1.cell(row=row, column=col_idx)
            c.fill   = _fill(CLR_TOTAL_ROW)
            c.border = _border()

    # ═══════════════════════════════════════════════════════════
    # SAYFA 2 — ÜRÜN BAZINDA ÖZET
    # ═══════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("Ürün Özeti")

    ws2.merge_cells("A1:E1")
    t2 = ws2["A1"]
    t2.value     = f"Ürün Bazında Satış Özeti — {today_str}"
    t2.font      = _font(bold=True, color=CLR_HEADER_FONT, size=12)
    t2.fill      = _fill(CLR_HEADER)
    t2.alignment = _align("center")
    ws2.row_dimensions[1].height = 24

    prod_cols = [("Ürün Adı", 40), ("Adet", 10), ("Toplam Tutar (₺)", 20), ("Ort. Birim Fiyat (₺)", 22), ("% Pay", 12)]
    for ci, (cn, cw) in enumerate(prod_cols, 1):
        cell = ws2.cell(row=2, column=ci, value=cn)
        cell.font      = _font(bold=True, color=CLR_SUBHDR_FONT)
        cell.fill      = _fill(CLR_SUBHEADER)
        cell.alignment = _align("center")
        cell.border    = _border()
        _set_col_width(ws2, ci, cw)

    active_orders = [o for o in orders if o.get("packageStatus") not in cancelled_statuses]
    product_qty = defaultdict(int)
    product_rev = defaultdict(float)

    for order in active_orders:
        for line in order.get("lines", []):
            name  = line.get("name", "?")
            qty   = len(line.get("items", [])) or 1
            price = line.get("price", 0)
            product_qty[name] += qty
            product_rev[name] += price * qty

    grand_total = sum(product_rev.values()) or 1
    sorted_prods = sorted(product_qty.items(), key=lambda x: x[1], reverse=True)

    for ri, (name, qty) in enumerate(sorted_prods, start=3):
        rev      = product_rev[name]
        avg      = rev / qty if qty else 0
        pct      = rev / grand_total * 100
        row_fill = _fill(CLR_ROW_ODD if ri % 2 == 0 else CLR_ROW_EVEN)

        row_data = [name, qty, round(rev, 2), round(avg, 2), round(pct, 1)]
        for ci, val in enumerate(row_data, 1):
            cell = ws2.cell(row=ri, column=ci, value=val)
            cell.fill      = row_fill
            cell.border    = _border()
            cell.font      = _font(size=9)
            cell.alignment = _align(h="left" if ci == 1 else "center")
            if ci in (3, 4):
                cell.number_format = '#,##0.00 ₺'
            if ci == 5:
                cell.number_format = '0.0"%"'

    # Özet toplam
    total_row = len(sorted_prods) + 3
    ws2.cell(row=total_row, column=1, value="TOPLAM").font = _font(bold=True)
    ws2.cell(row=total_row, column=1).fill = _fill(CLR_TOTAL_ROW)
    ws2.cell(row=total_row, column=1).border = _border()
    ws2.cell(row=total_row, column=1).alignment = _align()

    total_qty_val = sum(product_qty.values())
    ws2.cell(row=total_row, column=2, value=total_qty_val).fill  = _fill(CLR_TOTAL_ROW)
    ws2.cell(row=total_row, column=2).border    = _border()
    ws2.cell(row=total_row, column=2).font      = _font(bold=True)
    ws2.cell(row=total_row, column=2).alignment = _align()

    total_rev_cell = ws2.cell(row=total_row, column=3, value=round(grand_total, 2))
    total_rev_cell.fill          = _fill(CLR_TOTAL_ROW)
    total_rev_cell.border        = _border()
    total_rev_cell.font          = _font(bold=True)
    total_rev_cell.number_format = '#,##0.00 ₺'
    total_rev_cell.alignment     = _align()

    for ci in (4, 5):
        ws2.cell(row=total_row, column=ci).fill   = _fill(CLR_TOTAL_ROW)
        ws2.cell(row=total_row, column=ci).border = _border()

    # ═══════════════════════════════════════════════════════════
    # SAYFA 3 — GENEL ÖZET
    # ═══════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("Genel Özet")
    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 20

    ws3.merge_cells("A1:B1")
    t3 = ws3["A1"]
    t3.value     = f"Genel Özet — {today_str}"
    t3.font      = _font(bold=True, color=CLR_HEADER_FONT, size=12)
    t3.fill      = _fill(CLR_HEADER)
    t3.alignment = _align("center")
    ws3.row_dimensions[1].height = 24

    cancelled_statuses_set = {"Cancelled", "UnSupplied"}
    active   = [o for o in orders if o.get("packageStatus") not in cancelled_statuses_set]
    cancelled = [o for o in orders if o.get("packageStatus") in cancelled_statuses_set]
    revenue   = sum(o.get("totalPrice", 0) for o in active)

    payment_counts = defaultdict(int)
    for o in active:
        payment_counts[_payment_label(o)] += 1

    app_counts = defaultdict(int)
    for o in active:
        app_counts[_app_label(o.get("userInformation", {}).get("appName", ""))] += 1

    summary_rows = [
        ("📦 Toplam Sipariş", len(orders)),
        ("✅ Geçerli Sipariş", len(active)),
        ("❌ İptal Sipariş", len(cancelled)),
        ("💰 Toplam Ciro (₺)", round(revenue, 2)),
        ("", ""),
        ("— Ödeme Yöntemleri —", ""),
    ]
    for label, cnt in sorted(payment_counts.items(), key=lambda x: x[1], reverse=True):
        summary_rows.append((f"  {label}", cnt))

    summary_rows.append(("", ""))
    summary_rows.append(("— Sipariş Kaynakları —", ""))
    for label, cnt in sorted(app_counts.items(), key=lambda x: x[1], reverse=True):
        summary_rows.append((f"  {label}", cnt))

    for ri, (label, value) in enumerate(summary_rows, start=2):
        la = ws3.cell(row=ri, column=1, value=label)
        va = ws3.cell(row=ri, column=2, value=value)

        is_section = str(label).startswith("—")
        is_main    = ri in (2, 3, 4, 5)

        if is_section:
            la.font = _font(bold=True, color=CLR_SUMMARY_FNT, size=10)
            la.fill = _fill(CLR_SUMMARY_HDR)
            va.fill = _fill(CLR_SUMMARY_HDR)
        elif is_main:
            la.font = _font(bold=True, size=10)
            la.fill = _fill(CLR_TOTAL_ROW)
            va.fill = _fill(CLR_TOTAL_ROW)
            va.font = _font(bold=True, size=10)
        else:
            row_fill2 = _fill(CLR_ROW_ODD if ri % 2 == 0 else CLR_ROW_EVEN)
            la.fill  = row_fill2
            va.fill  = row_fill2
            la.font  = _font(size=10)
            va.font  = _font(size=10)

        la.border    = _border()
        va.border    = _border()
        la.alignment = _align(h="left")
        va.alignment = _align(h="center")

        if label == "💰 Toplam Ciro (₺)":
            va.number_format = '#,##0.00 ₺'

    wb.save(filepath)
    print(f"[EXCEL] Dosya oluşturuldu: {filepath}")


def send_excel_to_telegram(filepath: str, today_str: str):
    """Excel dosyasını Telegram'a gönderir."""
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    caption = (
        f"📊 <b>Günlük Sipariş Raporu — {today_str}</b>\n"
        f"Trendyol Go sipariş detayları ektedir."
    )
    try:
        with open(filepath, "rb") as f:
            response = requests.post(
                url,
                data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption, "parse_mode": "HTML"},
                files={"document": (os.path.basename(filepath), f,
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                timeout=30,
            )
        response.raise_for_status()
        print(f"[EXCEL] ✅ Excel Telegram'a gönderildi.")
        return True
    except requests.exceptions.RequestException as e:
        print(f"[EXCEL] ❌ Gönderilemedi: {e}")
        return False


def generate_and_send():
    """Ana fonksiyon: sipariş çek → Excel oluştur → Telegram'a gönder → dosyayı sil."""
    today_str = datetime.now(TURKEY_TZ).strftime("%d.%m.%Y")
    filename  = f"TrendyolGo_Rapor_{today_str.replace('.', '-')}.xlsx"

    print(f"[EXCEL] Siparişler çekiliyor...")
    orders = fetch_all_orders_today()
    print(f"[EXCEL] {len(orders)} sipariş bulundu.")

    build_excel(orders, filename)
    send_excel_to_telegram(filename, today_str)

    # Geçici dosyayı temizle
    if os.path.exists(filename):
        os.remove(filename)
        print(f"[EXCEL] Geçici dosya silindi.")


if __name__ == "__main__":
    generate_and_send()
