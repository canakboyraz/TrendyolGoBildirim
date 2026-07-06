"""
SQLite veritabanı — siparişleri kalıcı olarak saklar.
Servis yeniden başlasa bile geçmiş siparişler korunur.
Haftalık / aylık raporlar için de kullanılır.
"""

import sqlite3
import json
import os
from datetime import datetime
import pytz
from dotenv import load_dotenv
load_dotenv()

TURKEY_TZ = pytz.timezone("Europe/Istanbul")
DB_PATH = os.environ.get("DB_PATH", "orders.db")


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Tabloları oluşturur (yoksa)."""
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                id                TEXT PRIMARY KEY,
                order_number      TEXT,
                order_code        TEXT,
                store_id          INTEGER,
                supplier_id       INTEGER,
                status            TEXT,
                total_price       REAL,
                delivery_type     TEXT,
                payment_type      TEXT,
                app_name          TEXT,
                customer_name     TEXT,
                customer_note     TEXT,
                is_test           INTEGER,
                is_cancelled      INTEGER,
                created_ts        INTEGER,
                lines_json        TEXT,
                raw_json          TEXT,
                notified          INTEGER DEFAULT 0,
                notified_statuses TEXT DEFAULT '',
                recorded_at       TEXT
            )
        """)
        # Eski DB'lere kolon ekle (migration)
        try:
            conn.execute("ALTER TABLE orders ADD COLUMN notified_statuses TEXT DEFAULT ''")
        except Exception:
            pass  # Zaten varsa hata vermez
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_orders_created
            ON orders(created_ts)
        """)
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_orders_status
            ON orders(status)
        """)
        conn.commit()


def upsert_order(order: dict):
    """Siparişi ekler ya da günceller."""
    cancelled_statuses = {"Cancelled", "UnSupplied"}
    status    = order.get("packageStatus", "")
    customer  = order.get("customer", {})
    user_info = order.get("userInformation", {}) or {}

    payment = order.get("payment", {}) or {}
    raw_pt  = payment.get("paymentType", "")
    pt_map  = {
        "PAY_WITH_CARD":        "Online Kart",
        "PAY_WITH_ON_DELIVERY": "Kapıda Ödeme",
        "PAY_WITH_MEAL_CARD":   "Yemek Kartı",
    }
    payment_label = pt_map.get(raw_pt, raw_pt)

    now_str = datetime.now(TURKEY_TZ).strftime("%Y-%m-%d %H:%M:%S")

    with get_conn() as conn:
        conn.execute("""
            INSERT INTO orders (
                id, order_number, order_code, store_id, supplier_id,
                status, total_price, delivery_type, payment_type,
                app_name, customer_name, customer_note, is_test,
                is_cancelled, created_ts, lines_json, raw_json,
                notified, recorded_at
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(id) DO UPDATE SET
                status       = excluded.status,
                is_cancelled = excluded.is_cancelled,
                raw_json     = excluded.raw_json
        """, (
            order.get("id"),
            order.get("orderNumber"),
            order.get("orderCode"),
            order.get("storeId"),
            order.get("supplierId"),
            status,
            order.get("totalPrice", 0),
            order.get("deliveryType"),
            payment_label,
            user_info.get("appName"),
            f"{customer.get('firstName','')} {customer.get('lastName','')}".strip(),
            order.get("customerNote", ""),
            1 if order.get("testPackage") else 0,
            1 if status in cancelled_statuses else 0,
            order.get("packageCreationDate", 0),
            json.dumps(order.get("lines", []), ensure_ascii=False),
            json.dumps(order, ensure_ascii=False),
            0,
            now_str,
        ))
        conn.commit()


def is_notified(order_id: str) -> bool:
    """Sipariş için herhangi bir bildirim gönderildi mi?"""
    with get_conn() as conn:
        row = conn.execute(
            "SELECT notified FROM orders WHERE id = ?", (order_id,)
        ).fetchone()
        return bool(row and row["notified"])


def is_status_notified(order_id: str, status: str) -> bool:
    """Bu sipariş için bu statü bildirimi daha önce gönderildi mi?"""
    with get_conn() as conn:
        row = conn.execute(
            "SELECT notified_statuses FROM orders WHERE id = ?", (order_id,)
        ).fetchone()
        if not row:
            return False
        sent = row["notified_statuses"] or ""
        return status in sent.split(",")


def mark_notified(order_id: str):
    """Siparişi genel olarak bildirildi say."""
    with get_conn() as conn:
        conn.execute(
            "UPDATE orders SET notified = 1 WHERE id = ?", (order_id,)
        )
        conn.commit()


def mark_status_notified(order_id: str, status: str):
    """Bu sipariş için bu statüyü bildirildi olarak işaretle."""
    with get_conn() as conn:
        row = conn.execute(
            "SELECT notified_statuses FROM orders WHERE id = ?", (order_id,)
        ).fetchone()
        existing = (row["notified_statuses"] or "") if row else ""
        statuses = set(existing.split(",")) if existing else set()
        statuses.add(status)
        conn.execute(
            "UPDATE orders SET notified = 1, notified_statuses = ? WHERE id = ?",
            (",".join(statuses), order_id)
        )
        conn.commit()


def get_orders_between(start_ms: int, end_ms: int) -> list:
    """Verilen zaman aralığındaki siparişleri döner."""
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT raw_json FROM orders
            WHERE created_ts BETWEEN ? AND ?
            ORDER BY created_ts ASC
        """, (start_ms, end_ms)).fetchall()
    result = []
    for row in rows:
        try:
            result.append(json.loads(row["raw_json"]))
        except Exception:
            pass
    return result


def get_orders_by_date_str(date_str: str) -> list:
    """'YYYY-MM-DD' formatında tarih alır, o günün siparişlerini döner."""
    from datetime import date, time as dtime
    from datetime import datetime as dt
    d = dt.strptime(date_str, "%Y-%m-%d").date()
    start = TURKEY_TZ.localize(dt.combine(d, dtime.min))
    end   = TURKEY_TZ.localize(dt.combine(d, dtime.max))
    return get_orders_between(
        int(start.timestamp() * 1000),
        int(end.timestamp()   * 1000),
    )


def get_all_order_ids() -> set:
    """Veritabanındaki tüm sipariş ID'lerini döner."""
    with get_conn() as conn:
        rows = conn.execute("SELECT id FROM orders").fetchall()
    return {row["id"] for row in rows}


# Başlangıçta DB'yi hazırla
init_db()
