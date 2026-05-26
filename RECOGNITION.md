# Распознавание УПД / счёт-фактур

## Поток (гибрид, как рекомендует Gemini)

```
PDF → OCR.space (сырой текст, без структуры)
    → Gemini (только текст, не PDF — меньше токенов и 429)
    → ===HEADER=== / ===TABLE=== (TAB, предпочтительно) или JSON
    → при битом JSON — повтор запроса только в TAB-формате; repair частичного JSON
    → Google Таблица (батч setValues, не построчно)
```

OCR **не** раскладывает колонки — только «картинка → текст». Таблицу восстанавливает Gemini.

## Ключи

- `OCR_SPACE_API_KEY` — https://ocr.space/ocrapi
- `GEMINI_API_KEY` — https://aistudio.google.com/apikey

## Запись в Sheets

- Данные собираются в памяти, на лист пишутся **пакетами** (`SHEETS_FLUSH_EVERY_N_PDF`, по умолчанию 25 PDF) одним вызовом `setValues` на блок строк.
- Не используется `appendRow` / построчная запись (избегаем 429 от Google Sheets API).

## Меню

**Загрузить из папки Drive** — оба ключа обязательны.

Версия: `SCRIPT_VERSION` в `InvoicePdfToSheet.gs`.
