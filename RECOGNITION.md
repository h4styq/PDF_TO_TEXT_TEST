# Схема распознавания УПД / счёт-фактур

Цель — **общий конвейер** для любых PDF, без отдельных кнопок «сверки с эталоном» в меню.

## Запуск из таблицы

| Пункт меню | Функция | Поведение |
|------------|---------|-----------|
| **Загрузить из папки (Gemini)** | `runProcessFolderGemini` | Сначала Gemini (PDF), при сбое — OCR.space. Между PDF пауза ~20 с (лимит 429). |
| **Загрузить из папки (OCR.space, без паузы 20 с)** | `runProcessFolderOcr` | Только OCR.space, Gemini не вызывается, паузы нет. |

`runProcessFolder` оставлен как alias на Gemini (совместимость).

Ключи в **Свойствах скрипта**: `GEMINI_API_KEY`, `OCR_SPACE_API_KEY`.

## Поток данных

```
PDF → pdfToExtracted_(fileId, mode)
        mode = 'gemini' | 'ocr'
      → parseInvoiceData_
      → normalizeGoodsTableRows_ (общие правила + редкий vendor-слой)
      → лист «Счета_фактуры»
```

## Таблица товаров (общее ядро)

1. **semanticMapGoodsRow_** — токены в 15 граф УПД.
2. **repairScrambledOcrRow_** — локальные исправления OCR.
3. **normalizeGoodsTableRowOrderGeneric_** — тип строки (товар / доставка), порядок, дубли наименований в блоке УПД.
4. **repairGenericMappedRow_** — наименование из плоского текста, суммы, НДС.
5. **applyVendorDocumentRepairs_** — только при явных маркерах в тексте (сужается со временем).

## Выбор таблицы OCR

`pickBestOcrTable_`: блок УПД vs построчная эвристика; штраф блоку за **дублирующиеся наименования** в соседних строках.

## Версия

`SCRIPT_VERSION` в `InvoicePdfToSheet.gs` (Ctrl+F в редакторе Apps Script).
