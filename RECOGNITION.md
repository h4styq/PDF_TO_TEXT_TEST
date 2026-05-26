# Распознавание УПД / счёт-фактур

## Поток

```
PDF (Drive)
  → OCR.space (upload, при неудаче URL; все страницы ParsedResults)
  → Gemini (HEADER/TABLE TAB или JSON; повтор TAB при битом JSON)
  → лист «Счета_фактуры» (15 граф CANONICAL_UPD_HEADERS)
```

## Ключи (свойства скрипта)

- `OCR_SPACE_API_KEY` — https://ocr.space/ocrapi
- `GEMINI_API_KEY` — https://aistudio.google.com/apikey

## Меню

**Загрузить из папки Drive** — оба ключа обязательны.

Версия: см. `SCRIPT_VERSION` в `InvoicePdfToSheet.gs`.

## Примечания

- Крупные PDF: при ошибке upload OCR.space запрашивает файл по временной ссылке Drive.
- Очень длинный OCR (>120 000 симв.) обрезается только для запроса к Gemini (лимит Apps Script), не на этапе OCR.
- Эвристики по вендорам, golden-check и парсинг PDF через Gemini **удалены** — только OCR + структура от Gemini.
