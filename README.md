# Telegram Office Assistant Bot

Bu layihə Telegram üzərindən ofis faylları ilə işləyən sadə bot skeletidir. Bot Excel/CSV faylı qəbul edir, sonra yazdığınız tapşırığa görə faylı analiz edir və nəticəni geri göndərir.

## Hazır imkanlar

- `.xlsx`, `.xls`, `.csv` fayllarını qəbul edir
- Fayl haqqında qısa xülasə verir
- Sütunları göstərir
- Sütuna görə sıralayır
- Sadə filtr tətbiq edir
- CSV faylını Excel-ə çevirir
- Excel faylını CSV-yə çevirir

## Quraşdırma

1. Telegram-da `@BotFather` ilə yeni bot yaradın və token alın.
2. Layihə qovluğunda virtual mühit yaradın:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

3. `.env.example` faylını `.env` kimi kopyalayın və tokeni yazın:

```bash
cp .env.example .env
```

4. Botu başladın:

```bash
python bot.py
```

## İstifadə nümunələri

Faylı bota göndərin və caption kimi və ya sonra ayrıca mesajda belə tapşırıq yazın:

- `xülasə ver`
- `sütunları göstər`
- `satış sütununa görə sırala`
- `status = ödənilib olanları filtr et`
- `csv et`
- `excel et`

## Genişləndirmə ideyası

Növbəti mərhələdə bu bota OpenAI API əlavə edib daha sərbəst əmrləri başa düşdürmək olar, məsələn:

- `Bu Excel-də aylıq satışları hesabla və ayrıca hesabat yarat`
- `Borcu qalan müştəriləri tap, toplam məbləği yaz`
- `Bu sənəddən rəhbər üçün qısa xülasə hazırla`

