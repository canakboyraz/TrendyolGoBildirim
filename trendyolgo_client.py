import base64
import requests
from config import SUPPLIER_ID, API_KEY, API_SECRET, INTEGRATOR_NAME, API_BASE_URL


def _get_auth_header() -> str:
    """Basic Auth header'ı oluşturur."""
    credentials = f"{API_KEY}:{API_SECRET}"
    encoded = base64.b64encode(credentials.encode("utf-8")).decode("utf-8")
    return f"Basic {encoded}"


def _get_headers() -> dict:
    """Tüm isteklerde kullanılacak header'ları döner."""
    return {
        "Authorization": _get_auth_header(),
        "User-Agent": f"{SUPPLIER_ID} - {INTEGRATOR_NAME}",
        "x-agentname": INTEGRATOR_NAME,
        "x-executor-user": "integration@selfservice.com",
        "Content-Type": "application/json",
    }


def get_new_orders() -> list:
    """
    Aktif tüm statülerdeki sipariş paketlerini çeker.
    (Created, Picking, Invoiced, Shipped)
    Hata durumunda boş liste döner.
    """
    url = f"{API_BASE_URL}/integrator/order/meal/suppliers/{SUPPLIER_ID}/packages"
    params = {
        "packageStatuses": "Created,Picking,Invoiced,Shipped",
        "page": 0,
        "size": 50,
    }

    try:
        response = requests.get(
            url,
            headers=_get_headers(),
            params=params,
            timeout=15,
        )
        response.raise_for_status()
        data = response.json()
        return data.get("content", [])
    except requests.exceptions.HTTPError as e:
        print(f"[API HATA] HTTP {e.response.status_code}: {e.response.text}")
        return []
    except requests.exceptions.RequestException as e:
        print(f"[API HATA] Bağlantı hatası: {e}")
        return []


def get_stores() -> list:
    """Restoranları listeler (başlangıçta bağlantı testi için kullanılır)."""
    url = f"{API_BASE_URL}/integrator/store/meal/suppliers/{SUPPLIER_ID}/stores"
    params = {"page": 0, "size": 10}

    try:
        response = requests.get(
            url,
            headers=_get_headers(),
            params=params,
            timeout=15,
        )
        response.raise_for_status()
        data = response.json()
        return data.get("restaurants", [])
    except requests.exceptions.RequestException as e:
        print(f"[API HATA] Restoran listesi alınamadı: {e}")
        return []
