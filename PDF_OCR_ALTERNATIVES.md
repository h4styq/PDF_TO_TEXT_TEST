# Бесплатные и условно-бесплатные API для PDF (УПД / счета-фактуры)

Обзор для интеграции в Google Apps Script (`UrlFetchApp`, ключи в свойствах скрипта).  
**Сейчас в скрипте:** **OCR.space** распознаёт PDF → **Gemini** разбирает только OCR-текст в колонки листа. PDF в Gemini не отправляется.

Критерии для наших PDF:

- кириллица, таблицы УПД (много колонок, «796», «шт», суммы);
- сканы и текстовые PDF;
- вызов из Apps Script (POST, API key, лимит размера blob ~50 МБ, таймаут выполнения ~6 мин).

---

## Уже используем (рекомендуется)

| Сервис | Бесплатно | Плюсы | Минусы |
|--------|-----------|-------|--------|
| **Google Gemini API** | Квота в AI Studio (меняется; часто достаточно для пакетной обработки) | PDF нативно, можно задать структуру `===TABLE===` | HTTP 429 при частых запросах → пауза в скрипте |
| **OCR.space** | Да, с регистрацией ключа | Простой REST, русский (`language=rus`) | ~1 МБ на файл на free; Engine 2/3 лучше для таблиц |

Ключи: `GEMINI_API_KEY`, `OCR_SPACE_API_KEY`.

---

## Кандидаты на добавление в скрипт (приоритет)

### 1. Google Cloud Vision API (Document Text Detection)

- **Бесплатно:** первые **1000 единиц/мес** (1 страница PDF = 1 единица).
- **Плюсы:** сильный OCR, PDF до 2000 стр., хорошо для сканов.
- **Минусы:** нужен GCP-проект, OAuth2 или API key + отдельная настройка async для PDF (GCS bucket); в Apps Script сложнее, чем один POST с ключом.
- Документация: https://cloud.google.com/vision/docs/pdf

**Вердикт:** имеет смысл, если готовы завести GCP; для «одного ключа в свойствах» — тяжеловато.

### 2. Azure Document Intelligence (бывш. Form Recognizer)

- **Бесплатно (F0):** **500 страниц/мес**, но в одном запросе часто только **первые 2 страницы** PDF.
- **Плюсы:** layout/tables, prebuilt invoice.
- **Минусы:** многостраничные УПД на F0 режутся; нужен Azure resource + ключ/endpoint.
- Цены: https://azure.microsoft.com/pricing/details/ai-document-intelligence/

**Вердикт:** для типовых 1–2 стр. УПД — ок; для длинных накладных — слабо на free.

### 3. Mistral OCR (`mistral-ocr-latest`)

- **Бесплатно:** обычно только стартовые кредиты в console.mistral.ai, далее ~$1 / 1000 стр.
- **Плюсы:** markdown + таблицы, до 2000 стр./мин, хорошее качество на документах.
- **Минусы:** не «вечно бесплатный»; отдельный биллинг.
- API: https://docs.mistral.ai/api/endpoint/ocr

**Вердикт:** хороший платный запасной вариант, не замена OCR.space на нулевой бюджет.

### 4. OpenRouter (PDF → markdown / vision)

- **Бесплатно:** роутер `openrouter/free` и отдельные `:free` модели; PDF через engines (`cloudflare-ai` — конвертация в markdown заявлена как free tier в доке).
- **Плюсы:** один API на разные модели; PDF URL или base64.
- **Минусы:** лимиты free непредсказуемы; для production нужен платный баланс.
- Документация: https://openrouter.ai/docs/guides/overview/multimodal/pdfs

**Вердикт:** можно прототипировать второй «умный» путь рядом с Gemini, если появится стабильная free-модель с PDF.

### 5. OCRAPI.cloud / Optiic / EasyOCR (сторонние OCR API)

| Сервис | Free | Заметки |
|--------|------|---------|
| OCRAPI.cloud | tier с лимитами | PDF async, таблицы — проверить русский |
| Optiic | free tier | PDF, много языков |
| EasyOCR.org | без ключа (ограниченно) | в основном **изображения**, не полноценный PDF pipeline |

**Вердикт:** дублируют OCR.space; имеет смысл только если OCR.space не устраивает по качеству/лимиту.

---

## Слабее под наш сценарий

| Сервис | Почему |
|--------|--------|
| **AnyParser / CambioML** | API только с **платной** подпиской в app.cambioml.com |
| **Google Document AI** | мощно, но настройка и квоты не «один ключ за 5 минут» |
| **AWS Textract** | free tier очень ограничен, сложная авторизация из GAS |
| **Tesseract локально** | в Apps Script нет своего сервера |
| **Docling, Unstructured** | self-hosted, не UrlFetchApp |

---

## Практическая рекомендация

1. **Оставить основной стек:** Gemini (структура + сложные PDF) + OCR.space (сканы, fallback).
2. **Улучшить без нового API:** сжатие PDF >1 МБ перед OCR.space; `OCR_TRY_DRIVE_URL_FOR_LARGE`; пауза при 429 Gemini.
3. **Если нужен ещё один бесплатный OCR:** сначала **Cloud Vision** (1000 стр./мес) или второй ключ **OCR.space** (PRO) — не платить за Mistral/AnyParser.
4. **Если появится бюджет:** Mistral OCR или Azure S0 — для таблиц и многостраничных УПД.

---

## Что нужно для интеграции нового API в `InvoicePdfToSheet.gs`

1. Свойство скрипта, например `VISION_API_KEY` или `MISTRAL_API_KEY`.
2. Функция `tryXxxPdfExtract_(pdfFileId, apiKey)` → `{ text, source: 'xxx' }`.
3. Вставка в цепочку `tryExternalTextExtractionGeminiFirst_` или отдельный пункт меню.
4. Лимит размера и язык `rus`/кириллица в запросе.
5. Запись в `RECOGNITION.md` и справку `showRecognitionSetupHelp`.

При выборе кандидата — напишите, какой сервис пробуем первым (Vision / Azure / OpenRouter / другой).
