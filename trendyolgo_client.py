import base64
import requests
from config import SUPPLIER_ID, API_KEY, API_SECRET, INTEGRATOR_NAME, API_BASE_URL


def _get_auth_header() -> str:
    credentials = f"{API_KEY}:{API_SECRET}"
    encoded = base64.b64encode(credentials.encode("utf-8")).decode("utf-8")
    return f"Basic {encoded}"


def _get_headers() -> dict:
    return {
        "Authorization": _get_auth_header(),
        "User-Agent": f"{SUPPLIER_ID} - {INTEGRATOR_NAME}",
        "x-agentname": INTEGRATOR_NAME,
        "x-executor-user": "integration@selfservice.com",
        "Content-Type": "application/json",
    }


def get_new_orders() -> list:
    """Aktif statülerdeki siparişleri çeker."""
    url = f"{API_BASE_URL}/integrator/order/meal/suppliers/{SUPPLIER_ID}/packages"
    params = {
        "packageStatuses": "Created,Picking,Invoiced,Shipped",
        "page": 0,
        "size": 50,
    }
    try:
        response = requests.get(url, headers=_get_headers(), params=params, timeout=15)
        response.raise_for_status()
        return response.json().get("content", [])
    except requests.exceptions.HTTPError as e:
        print(f"[API HATA] HTTP {e.response.status_code}: {e.response.text}")
        return []
    except requests.exceptions.RequestException as e:
        print(f"[API HATA] Bağlantı hatası: {e}")
        return []


def get_active_orders() -> list:
    """Bekleyen (Created + Picking) siparişleri döner."""
    url = f"{API_BASE_URL}/integrator/order/meal/suppliers/{SUPPLIER_ID}/packages"
    params = {"packageStatuses": "Created,Picking", "page": 0, "size": 50}
    try:
        response = requests.get(url, headers=_get_headers(), params=params, timeout=15)
        response.raise_for_status()
        return response.json().get("content", [])
    except requests.exceptions.RequestException as e:
        print(f"[API HATA] Aktif sipariş alınamadı: {e}")
        return []


def get_stores() -> list:
    """Restoranları listeler."""
    url = f"{API_BASE_URL}/integrator/store/meal/suppliers/{SUPPLIER_ID}/stores"
    params = {"page": 0, "size": 10}
    try:
        response = requests.get(url, headers=_get_headers(), params=params, timeout=15)
        response.raise_for_status()
        return response.json().get("restaurants", [])
    except requests.exceptions.RequestException as e:
        print(f"[API HATA] Restoran listesi alınamadı: {e}")
        return []


def update_store_status(store_id: int, status: str) -> tuple:
    """
    Restoranı açar veya kapatır.
    status: "OPEN" veya "CLOSED"
    Döner: (başarılı_mı: bool, mesaj: str)
    """
    url = (
        f"{API_BASE_URL}/integrator/store/meal/suppliers/{SUPPLIER_ID}"
        f"/stores/{store_id}/status"
    )
    try:
        response = requests.put(
            url,
            headers=_get_headers(),
            json={"status": status},
            timeout=15,
        )
        if response.status_code == 200:
            label = "AÇILDI ✅" if status == "OPEN" else "KAPATILDI 🔴"
            return True, f"Restoran başarıyla {label}"
        else:
            return False, f"Hata {response.status_code}: {response.text}"
    except requests.exceptions.RequestException as e:
        return False, f"Bağlantı hatası: {e}"
