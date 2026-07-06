"""Tüm statülerdeki siparişleri ham JSON olarak gösterir."""
import json
import requests
import base64
from config import SUPPLIER_ID, API_KEY, API_SECRET, INTEGRATOR_NAME, API_BASE_URL


def get_headers():
    credentials = f"{API_KEY}:{API_SECRET}"
    encoded = base64.b64encode(credentials.encode()).decode()
    return {
        "Authorization": f"Basic {encoded}",
        "User-Agent": f"{SUPPLIER_ID} - {INTEGRATOR_NAME}",
        "x-agentname": INTEGRATOR_NAME,
        "x-executor-user": "integration@selfservice.com",
        "Content-Type": "application/json",
    }


# Tüm statüleri dene
statuses = ["Created", "Picking", "Invoiced", "Shipped", "Delivered", "Cancelled", "UnSupplied"]

for status in statuses:
    url = f"{API_BASE_URL}/integrator/order/meal/suppliers/{SUPPLIER_ID}/packages"
    params = {"packageStatuses": status, "page": 0, "size": 10}
    r = requests.get(url, headers=get_headers(), params=params, timeout=15)
    data = r.json()
    count = data.get("totalCount", 0)
    if count and count > 0:
        print(f"\n✅ '{status}' statüsünde {count} sipariş bulundu!")
        print(json.dumps(data.get("content", [{}])[0], indent=2, ensure_ascii=False))
    else:
        print(f"   '{status}' → sipariş yok")
