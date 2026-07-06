# Trendyol Go Sipariş Bildirim Servisi

Yeni Trendyol Go siparişi geldiğinde Telegram üzerinden anlık bildirim gönderir.

## Kurulum

```bash
cd trendyolgo_notifier
pip install -r requirements.txt
```

## Çalıştırma

```bash
python main.py
```

## Dosya Yapısı

```
trendyolgo_notifier/
├── main.py              # Ana uygulama, döngü ve başlangıç kontrolü
├── trendyolgo_client.py # Trendyol Go API istekleri
├── telegram_notifier.py # Telegram mesaj gönderimi ve formatlama
├── config.py            # API anahtarları ve ayarlar
└── requirements.txt     # Bağımlılıklar
```

## Ayarlar (config.py)

| Parametre | Açıklama |
|---|---|
| `SUPPLIER_ID` | Trendyol Go Satıcı ID |
| `API_KEY` | API anahtarı |
| `API_SECRET` | API gizli anahtarı |
| `TELEGRAM_BOT_TOKEN` | Telegram bot token |
| `TELEGRAM_CHAT_ID` | Bildirim gönderilecek chat ID |
| `POLL_INTERVAL_SECONDS` | Kaç saniyede bir kontrol edilsin (varsayılan: 30) |

## Nasıl Çalışır?

1. Servis başlarken API ve Telegram bağlantısını test eder
2. Her 30 saniyede bir `Created` statüsündeki siparişleri çeker
3. Daha önce görülmemiş yeni sipariş varsa Telegram'a bildirim gönderir
4. `CTRL+C` ile temiz şekilde durdurulabilir

## Önemli Notlar

- Servis yeniden başlatılırsa, başlatılmadan önce gelen siparişler tekrar bildirim göndermez
- Trendyol Go API limiti: aynı endpoint'e 10 saniyede max 50 istek
- Bildirim mesajı HTML formatında gönderilir (Telegram destekler)
