/**
 * Парсинг счетов-фактур из PDF в Google Таблицу.
 *
 * ВАЖНО:
 * 1) Apps Script НЕ видит диск C:\ — положите PDF в папку на Google Drive и укажите SOURCE_FOLDER_ID.
 * 2) Включите сервис: Расширения → Apps Script → Сервисы → Google Drive API (v3).
 * 3) Первый запрос может запросить разрешения на Drive и Таблицы.
 *
 * ОГРАНИЧЕНИЯ PDF В APPS SCRIPT:
 * — Текст берётся после конвертации «PDF → Google Документ». Сканы без текстового слоя дадут пустой/бесполезный текст — нужен OCR (Document AI, Cloud Vision и т.д.), это уже вне этого скрипта.
 * — Поворот страницы на 90° иногда ломает порядок строк/таблиц при конвертации; надёжнее заранее выпрямить PDF (вручную или утилитой) либо использовать OCR по изображению страницы.
 * — Если в Doc «кракозябры» (>&F, случайные латинские куски без нормального русского текста) — это не ошибка скрипта: движок Google не смог извлечь текстовый слой из вашего PDF (часто скан, нестандартные шрифты, «картинка вместо текста»). Нужен другой исходный файл (OCR → поисковый PDF) или внешний OCR/API; Apps Script сам PDF не расшифрует.
 *
 * РАСПОЗНАВАНИЕ (опционально), если конвертация PDF→Doc нечитаема:
 * — В редакторе Apps Script: Проект → Свойства проекта → Свойства скрипта — добавьте один или оба ключа:
 *   GEMINI_API_KEY — ключ с https://aistudio.google.com/apikey (модель читает PDF и возвращает структурированный текст).
 *   OCR_SPACE_API_KEY — ключ с https://ocr.space/ocrapi (распознавание PDF, на бесплатном тарифе обычно лимит ~1 МБ на файл).
 * — Приоритет: сначала Gemini, затем OCR.space. Нужен доступ к внешней сети (UrlFetchApp) при первом запуске подтвердите разрешения.
 * — Если один раз всё получилось, а при повторе с теми же PDF — нет: часто лимиты/перегрузка API (429) или нестабильный ответ модели. В скрипте включены повторные запросы и более строгий сценарий вызова внешнего API.
 * Запуск:
 * — в самой таблице: меню «Счета-фактуры (PDF)» → «Загрузить данные из папки Drive» (после сохранения скрипта обновите страницу F5);
 * — в редакторе Apps Script: список функций слева от кнопки «Выполнить» — выберите runProcessFolder (если пункта нет, проверьте ошибки подсветкой и что код в проекте этой таблицы).
 */

/** ID папки на Google Drive (из URL: .../folders/THIS_ID) */
const SOURCE_FOLDER_ID = 'ВСТАВЬТЕ_ID_ПАПКИ';

/** true — удалять временные Google Docs после чтения текста */
const DELETE_TEMP_DOCS = true;

/** Имя листа для результата (создастся, если нет) */
const OUTPUT_SHEET_NAME = 'Счета_фактуры';

/** Модель Gemini для чтения PDF (при ошибке 404 смените на gemini-1.5-flash или gemini-2.5-flash) */
const GEMINI_MODEL = 'gemini-2.0-flash';

/** Макс. размер PDF для отправки в Gemini inline (байт); при превышении внешний шаг пропускается */
const MAX_GEMINI_INLINE_PDF_BYTES = 6 * 1024 * 1024;

/** Повторы при 429/5xx и «пустом» ответе Gemini (нестабильность API и модели) */
const GEMINI_MAX_ATTEMPTS = 4;
const GEMINI_RETRY_BASE_DELAY_MS = 2000;
const OCR_SPACE_MAX_ATTEMPTS = 3;

/**
 * Заголовки граф таблицы товаров (УПД / счёт-фактура), как в типовой форме.
 * Если в документе больше колонок — справа добавятся «Доп. столбец N».
 */
const USE_CANONICAL_TABLE_HEADERS = true;
const CANONICAL_UPD_HEADERS = [
  'Код товара/работ, услуг',
  '№ п/п',
  'Наименование товара (описание выполненных работ, оказанных услуг), имущественного права',
  'Код вида товара',
  'Единица измерения: код',
  'Единица измерения: условное обозначение (национальное)',
  'Цена (тариф) за единицу измерения',
  'Стоимость товаров (работ, услуг), имущественных прав без налога — всего',
  'В том числе сумма акциза',
  'Налоговая ставка',
  'Сумма налога, предъявляемого покупателю',
  'Стоимость товаров (работ, услуг), имущественных прав с налогом — всего',
  'Страна происхождения товара: цифровой код',
  'Страна происхождения товара: краткое наименование',
  'Регистрационный номер декларации на товары или регистрационный номер партии товара, подлежащего прослеживаемости',
];

/**
 * Меню в таблице — чтобы не искать runProcessFolder в редакторе.
 * Если меню не появилось: сохраните проект (Ctrl+S), вернитесь в таблицу и обновите страницу.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Счета-фактуры (PDF)')
    .addItem('Загрузить данные из папки Drive', 'runProcessFolder')
    .addSeparator()
    .addItem('Как подключить распознавание (Gemini / OCR)', 'showRecognitionSetupHelp')
    .addToUi();
}

/**
 * Точка входа: обходит все PDF в папке, пишет строки на активный spreadsheet
 * (файл таблицы, в котором открыт редактор скрипта, или привязанный к контейнеру).
 */
function runProcessFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Откройте таблицу и привязанный к ней скрипт, либо вызовите runProcessFolderForSpreadsheet(id).');
  }
  processFolderIntoSpreadsheet_(SOURCE_FOLDER_ID, ss.getId());
}

function countOutputRows_(items) {
  let n = 0;
  for (let i = 0; i < items.length; i++) {
    const r = items[i].parsed.tableRows.length;
    n += r === 0 ? 1 : r;
  }
  return n;
}

/**
 * Если скрипт отдельный (standalone), можно передать ID таблицы.
 */
function runProcessFolderForSpreadsheet(spreadsheetId) {
  processFolderIntoSpreadsheet_(SOURCE_FOLDER_ID, spreadsheetId);
}

function processFolderIntoSpreadsheet_(folderId, spreadsheetId) {
  Logger.log('Старт: папка Drive id=' + folderId + ', таблица id=' + spreadsheetId);
  if (!folderId || folderId.indexOf('ВСТАВЬТЕ') !== -1) {
    const msg = 'Задайте SOURCE_FOLDER_ID в коде (ID папки из URL Google Drive).';
    Logger.log(msg);
    SpreadsheetApp.openById(spreadsheetId).toast(msg, 'Счета-фактуры', 12);
    throw new Error(msg);
  }

  const folder = DriveApp.getFolderById(folderId);
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(OUTPUT_SHEET_NAME);
    Logger.log('Создан лист: ' + OUTPUT_SHEET_NAME);
  }

  const files = folder.getFilesByType(MimeType.PDF);
  const rows = [];
  let maxTableCols = 0;

  while (files.hasNext()) {
    const file = files.next();
    Logger.log('PDF: ' + file.getName());
    try {
      const pack = pdfToExtracted_(file.getId());
      const parsed = parseInvoiceData_(
        pack.text,
        pack.docTable,
        pack.textLength,
        pack.conversionOk,
        pack.conversionNote
      );
      maxTableCols = Math.max(maxTableCols, parsed.tableWidth);
      rows.push({ fileName: file.getName(), fileId: file.getId(), parsed: parsed });
    } catch (e) {
      Logger.log('Ошибка по файлу ' + file.getName() + ': ' + e.message);
      rows.push({
        fileName: file.getName(),
        fileId: file.getId(),
        parsed: {
          invoiceLine: 'ОШИБКА: ' + e.message,
          seller: '',
          paymentDoc: '',
          tableHeader: [],
          tableRows: [],
          basis: '',
          tableWidth: 0,
        },
      });
    }
  }

  writeParsedRows_(sheet, rows, maxTableCols);

  const pdfCount = rows.length;
  const outRows = countOutputRows_(rows);
  const summary =
    pdfCount === 0
      ? 'В папке не найдено PDF. Проверьте папку и права доступа.'
      : 'Обработано PDF: ' + pdfCount + '. Строк данных (с заголовком): ' + (outRows + 1) + '. Лист «' + OUTPUT_SHEET_NAME + '».';
  Logger.log(summary);
  ss.toast(summary, 'Счета-фактуры (PDF)', 12);
}

/**
 * Конвертирует PDF в Google Doc, забирает плоский текст и пытается прочитать таблицы Document (структура УПД).
 * @return {{text:string, textLength:number, docTable:Object|null, conversionOk:boolean, conversionNote:string}}
 */
function pdfToExtracted_(pdfFileId) {
  const name = 'tmp_pdf_' + new Date().getTime();
  const resource = {
    name: name,
    mimeType: MimeType.GOOGLE_DOCS,
  };
  const copied = Drive.Files.copy(resource, pdfFileId, { supportsAllDrives: true, fields: 'id' });
  const docId = copied.id;
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  let text = body.getText();
  let quality = analyzeDocTextQuality_(text);
  let docTable = null;
  const props = PropertiesService.getScriptProperties();
  const hasGemini = !!props.getProperty('GEMINI_API_KEY');
  const hasOcr = !!props.getProperty('OCR_SPACE_API_KEY');
  const hasAnyExternal = hasGemini || hasOcr;
  Logger.log('API-ключи: Gemini=' + (hasGemini ? 'да' : 'нет') + ', OCR.space=' + (hasOcr ? 'да' : 'нет'));

  if (quality.readable) {
    docTable = extractMainGoodsTableFromDoc_(body);
    const tableEmpty = !docTable || !docTable.rows || !docTable.rows.length;
    if (tableEmpty && hasAnyExternal) {
      Logger.log(
        'Текст после PDF→Doc прошёл проверку, но таблица товаров не извлечена — вызываем внешнее распознавание (Gemini/OCR).'
      );
      const improved = tryExternalTextExtraction_(pdfFileId);
      if (improved && improved.text) {
        const merged = mergeExternalExtractIntoPlainText_(improved.text);
        const q2 = analyzeDocTextQuality_(merged);
        if (q2.readable || looksStructuredGemini_(improved.text) || merged.length > text.length * 0.5) {
          text = merged;
          docTable = null;
          if (q2.readable || looksStructuredGemini_(improved.text)) {
            quality = { readable: true, reason: '' };
          } else {
            quality = q2;
          }
          Logger.log('Подставлен текст из ' + improved.source + ' (таблица из Doc была пуста).');
        }
      }
    } else if (tableEmpty && !hasAnyExternal) {
      Logger.log(
        'ВНИМАНИЕ: таблица товаров не найдена и нет API-ключей для внешнего распознавания. ' +
        'Добавьте GEMINI_API_KEY и/или OCR_SPACE_API_KEY в Свойствах скрипта (Проект → Свойства проекта → Свойства скрипта).'
      );
    }
  } else {
    Logger.log('Конвертация PDF→Doc нечитаема: ' + quality.reason);
    const improved = tryExternalTextExtraction_(pdfFileId);
    if (improved && improved.text) {
      const merged = mergeExternalExtractIntoPlainText_(improved.text);
      text = merged;
      const q3 = analyzeDocTextQuality_(text);
      if (q3.readable || looksStructuredGemini_(improved.text)) {
        quality = { readable: true, reason: '' };
      } else {
        quality = q3;
      }
      Logger.log('После внешнего распознавания (' + improved.source + '): readable=' + quality.readable);
      docTable = null;
    }
  }
  if (DELETE_TEMP_DOCS) {
    DriveApp.getFileById(docId).setTrashed(true);
  }
  return {
    text: text,
    textLength: text ? text.length : 0,
    docTable: docTable,
    conversionOk: quality.readable,
    conversionNote: quality.reason,
  };
}

/** Склеивает ответ Gemini (маркеры) или возвращает сырой текст OCR. */
function mergeExternalExtractIntoPlainText_(raw) {
  if (!raw) {
    return '';
  }
  const n = normalizeText_(raw);
  if (n.indexOf('===HEADER===') === -1) {
    return n;
  }
  const hm = n.match(/===HEADER===\s*([\s\S]*?)(?====TABLE===|$)/i);
  const tm = n.match(/===TABLE===\s*([\s\S]*?)(?====END===|$)/i);
  const h = hm ? hm[1].trim() : '';
  const t = tm ? tm[1].trim() : '';
  if (!h && !t) {
    return n.replace(/===HEADER===|===TABLE===|===END===/gi, '').trim();
  }
  return (h + (h && t ? '\n\n' : '') + t).trim();
}

/** Ответ модели похож на ожидаемый формат (даже если эвристика analyzeDocTextQuality строгая). */
function looksStructuredGemini_(raw) {
  if (!raw || raw.length < 80) {
    return false;
  }
  const low = raw.toLowerCase();
  const hasTabs = raw.indexOf('\t') !== -1;
  if (low.indexOf('===header===') !== -1 && low.indexOf('===table===') !== -1 && hasTabs) {
    return true;
  }
  if (/(счет|универсальн)[\s\S]{0,200}(фактур|передаточн)/i.test(raw) && hasTabs) {
    return true;
  }
  const cyr = (raw.match(/[а-яА-ЯёЁ]/g) || []).length;
  return hasTabs && cyr > 100;
}

/**
 * Если в свойствах скрипта задан ключ — пробуем извлечь читаемый текст из исходного PDF.
 * @return {{text:string, source:string}|null}
 */
function tryExternalTextExtraction_(pdfFileId) {
  const props = PropertiesService.getScriptProperties();
  const geminiKey = props.getProperty('GEMINI_API_KEY');
  if (geminiKey) {
    Logger.log('Пробуем распознавание через Gemini (' + GEMINI_MODEL + ')…');
    const g = tryGeminiPdfExtract_(pdfFileId, geminiKey);
    if (g && g.text && g.text.length > 80) {
      const merged = mergeExternalExtractIntoPlainText_(g.text);
      if (analyzeDocTextQuality_(merged).readable || looksStructuredGemini_(g.text)) {
        Logger.log('Gemini: получен читаемый текст (' + g.text.length + ' симв.).');
        return { text: g.text, source: 'gemini' };
      }
      Logger.log('Gemini: ответ есть (' + g.text.length + ' симв.), но слабый по качеству — пробуем OCR.space');
    } else {
      Logger.log('Gemini: не удалось получить текст' + (g && g.text ? ' (слишком короткий: ' + g.text.length + ' симв.)' : '') + '.');
    }
  } else {
    Logger.log('GEMINI_API_KEY не задан — пропускаем Gemini.');
  }
  const ocrKey = props.getProperty('OCR_SPACE_API_KEY');
  if (ocrKey) {
    Logger.log('Пробуем распознавание через OCR.space…');
    const o = tryOcrSpacePdfExtract_(pdfFileId, ocrKey);
    if (o && o.text && o.text.length > 40) {
      Logger.log('OCR.space: получен текст (' + o.text.length + ' симв.).');
      return { text: o.text, source: 'ocr.space' };
    }
    Logger.log('OCR.space: не удалось получить текст.');
  } else {
    Logger.log('OCR_SPACE_API_KEY не задан — пропускаем OCR.space.');
  }
  Logger.log('Внешнее распознавание не дало результата.');
  return null;
}

function tryGeminiPdfExtract_(pdfFileId, apiKey) {
  try {
    const file = DriveApp.getFileById(pdfFileId);
    const blob = file.getBlob();
    const size = blob.getBytes().length;
    if (size > MAX_GEMINI_INLINE_PDF_BYTES) {
      Logger.log('Gemini: PDF слишком большой для inline: ' + size + ' байт');
      return null;
    }
    const b64 = Utilities.base64Encode(blob.getBytes());
    const url =
      'https://generativelanguage.googleapis.com/v1beta/models/' +
      GEMINI_MODEL +
      ':generateContent?key=' +
      encodeURIComponent(apiKey);

    for (let attempt = 1; attempt <= GEMINI_MAX_ATTEMPTS; attempt++) {
      if (attempt > 1) {
        Utilities.sleep(GEMINI_RETRY_BASE_DELAY_MS * (attempt - 1));
      }
      const bodyObj = {
        contents: [
          {
            parts: [
              { inline_data: { mime_type: 'application/pdf', data: b64 } },
              { text: getGeminiInvoicePrompt_() },
            ],
          },
        ],
        generationConfig: {
          temperature: 0,
          topP: 0.95,
          topK: 40,
          maxOutputTokens: 8192,
        },
      };
      const resp = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: true,
        payload: JSON.stringify(bodyObj),
      });
      const code = resp.getResponseCode();
      const respText = resp.getContentText();
      if (code === 429 || code === 500 || code === 502 || code === 503 || code === 504) {
        Logger.log('Gemini HTTP ' + code + ', попытка ' + attempt + '/' + GEMINI_MAX_ATTEMPTS);
        continue;
      }
      if (code !== 200) {
        Logger.log('Gemini HTTP ' + code + ': ' + respText.substring(0, 800));
        return null;
      }
      let json;
      try {
        json = JSON.parse(respText);
      } catch (e1) {
        Logger.log('Gemini: не JSON, попытка ' + attempt);
        continue;
      }
      if (json.error) {
        Logger.log('Gemini API error: ' + JSON.stringify(json.error));
        return null;
      }
      if (json.promptFeedback && json.promptFeedback.blockReason) {
        Logger.log('Gemini blockReason: ' + json.promptFeedback.blockReason + ', попытка ' + attempt);
        continue;
      }
      const cand = json.candidates && json.candidates[0];
      if (!cand) {
        Logger.log('Gemini: нет candidates ' + respText.substring(0, 400));
        continue;
      }
      if (cand.finishReason && cand.finishReason !== 'STOP' && cand.finishReason !== 'MAX_TOKENS') {
        Logger.log('Gemini finishReason: ' + cand.finishReason + ', попытка ' + attempt);
        continue;
      }
      const parts = cand.content && cand.content.parts;
      if (!parts || !parts.length) {
        Logger.log('Gemini: пустые parts, попытка ' + attempt);
        continue;
      }
      let out = '';
      for (let i = 0; i < parts.length; i++) {
        out += parts[i].text || '';
      }
      if (out.length < 40) {
        Logger.log('Gemini: слишком короткий текст (' + out.length + ' симв.), попытка ' + attempt);
        continue;
      }
      return { text: out };
    }
    Logger.log('Gemini: исчерпаны попытки (' + GEMINI_MAX_ATTEMPTS + ')');
    return null;
  } catch (e) {
    Logger.log('Gemini: ' + e.message);
    return null;
  }
}

function getGeminiInvoicePrompt_() {
  return (
    'Извлеки данные из приложённого PDF (российский УПД или счёт-фактура). Ответ только текстом, без markdown.\n' +
    'Строго такой формат (строки-маркеры обязательны):\n' +
    '===HEADER===\n' +
    'Строка: «Счёт-фактура № … от …» или «УПД № … от …» — как в документе.\n' +
    'Строка: «Продавец:» и далее наименование продавца одной строкой.\n' +
    'Строка: «К платежно-расчетному документу №» и номер(а).\n' +
    'Строка с «Основание передачи (сдачи) / получения (приемки)» и текст основания; при наличии рядом «Счёт № …» — добавь в той же или следующей строке.\n' +
    '===TABLE===\n' +
    'Первая строка блока — заголовки граф таблицы товаров, разделённые символом TAB (табуляция).\n' +
    'Далее каждая строка — одна строка таблицы тем же числом колонок (TAB). Не включай строку «Всего к оплате» и итоги после неё.\n' +
    '===END===\n' +
    'Если фрагмента нет — оставь маркер и пустую секцию. Не выдумывай суммы и реквизиты.'
  );
}

function tryOcrSpacePdfExtract_(pdfFileId, apiKey) {
  const MAX_OCR_SPACE_BYTES = 1024 * 1024;
  try {
    const file = DriveApp.getFileById(pdfFileId);
    const blob = file.getBlob().setContentType('application/pdf');
    if (blob.getBytes().length > MAX_OCR_SPACE_BYTES) {
      Logger.log('OCR.space: файл больше ~1 МБ (лимит бесплатного тарифа).');
      return null;
    }
    for (let attempt = 1; attempt <= OCR_SPACE_MAX_ATTEMPTS; attempt++) {
      if (attempt > 1) {
        Utilities.sleep(1200 * attempt);
      }
      const resp = UrlFetchApp.fetch('https://api.ocr.space/parse/image', {
        method: 'post',
        muteHttpExceptions: true,
        payload: {
          apikey: apiKey,
          language: 'rus',
          isTable: 'true',
          OCREngine: '2',
          detectOrientation: 'true',
          scale: 'true',
          file: blob,
        },
      });
      const code = resp.getResponseCode();
      const raw = resp.getContentText();
      if (code === 429 || code === 503) {
        Logger.log('OCR.space HTTP ' + code + ', попытка ' + attempt);
        continue;
      }
      if (code !== 200) {
        Logger.log('OCR.space HTTP ' + code + ': ' + raw.substring(0, 400));
        return null;
      }
      let json;
      try {
        json = JSON.parse(raw);
      } catch (ignore) {
        Logger.log('OCR.space: ответ не JSON');
        continue;
      }
      if (json.IsErroredOnProcessing) {
        const em = json.ErrorMessage || '';
        Logger.log('OCR.space: ' + em);
        if (/limit|rate|many|quota/i.test(em) && attempt < OCR_SPACE_MAX_ATTEMPTS) {
          continue;
        }
        return null;
      }
      const pr = json.ParsedResults && json.ParsedResults[0];
      if (!pr) {
        continue;
      }
      const txt = pr.ParsedText || '';
      if (txt.length < 15 && attempt < OCR_SPACE_MAX_ATTEMPTS) {
        continue;
      }
      return { text: txt };
    }
    return null;
  } catch (e) {
    Logger.log('OCR.space: ' + e.message);
    return null;
  }
}

function showRecognitionSetupHelp() {
  SpreadsheetApp.getUi().alert(
    'Распознавание текста из PDF\n\n' +
      '1) Расширения → Apps Script → слева «Свойства проекта» (шестерёнка) → «Свойства скрипта».\n\n' +
      '2) Добавьте свойство:\n' +
      '   • GEMINI_API_KEY — ключ: https://aistudio.google.com/apikey\n' +
      '     (модель ' +
      GEMINI_MODEL +
      ' читает PDF; при ошибке модели смените константу GEMINI_MODEL в коде.)\n\n' +
      '   ИЛИ свойство:\n' +
      '   • OCR_SPACE_API_KEY — регистрация: https://ocr.space/ocrapi\n' +
      '     (часто лимит ~1 МБ на файл на бесплатном плане; включено определение ориентации страницы.)\n\n' +
      '3) Сохраните свойства и снова запустите «Загрузить данные из папки Drive». При запросе разрешите доступ к внешней сети (UrlFetchApp).\n\n' +
      'Сначала Gemini (несколько повторов при 429/5xx и пустом ответе), при слабом ответе — OCR.space. Без ключей — только конвертация Google.\n\n' +
      'При нестабильности: подождите минуту и повторите запуск или задайте оба ключа (OCR как запасной).'
  );
}

/**
 * Эвристика: конвертер Google иногда выдаёт «мусор» вместо русского текста УПД (артефакты шрифтов, скан без слоя).
 * @return {{readable:boolean, reason:string}}
 */
function analyzeDocTextQuality_(text) {
  if (!text || text.length < 25) {
    return { readable: false, reason: 'Почти пустой текст после конвертации PDF.' };
  }
  const cyr = (text.match(/[а-яА-ЯёЁ]/g) || []).length;
  const lat = (text.match(/[a-zA-Z]/g) || []).length;
  const letters = cyr + lat;
  const junkChars = (text.match(/[<>[\]{}|\\`~^'"″«»?&]/g) || []).length;
  const lower = text.toLowerCase();
  const keywords = [
    'счет',
    'фактур',
    'продавец',
    'покупатель',
    'наименован',
    'упд',
    'инн',
    'стоимость',
    'ндс',
    'товар',
    'услуг',
    'документ',
    'передач',
    'платежно',
    'расчетн',
    'к оплате',
  ];
  let kwHits = 0;
  for (let k = 0; k < keywords.length; k++) {
    if (lower.indexOf(keywords[k]) !== -1) {
      kwHits++;
    }
  }

  if (text.length > 150 && letters < 25) {
    return {
      readable: false,
      reason:
        'Длинный «текст», но почти нет букв — похоже на мусор после конвертации (не текстовый PDF или битое извлечение).',
    };
  }
  if (text.length > 250 && kwHits === 0 && cyr < 30) {
    return {
      readable: false,
      reason:
        'Нет характерных слов русского счёта/УПД — конвертация не дала нормальный текстовый слой (часто скан или защищённый/нестандартный PDF).',
    };
  }
  if (letters > 100 && cyr / Math.max(1, letters) < 0.045 && kwHits < 2) {
    return {
      readable: false,
      reason:
        'Очень мало кириллицы для российского документа — вероятно артефакты шрифтов или изображение страниц вместо текста.',
    };
  }
  if (junkChars / text.length > 0.055 && kwHits < 3 && cyr < 40) {
    return {
      readable: false,
      reason: 'Много служебных символов при почти отсутствии осмысленного русского текста — извлечение из PDF непригодно.',
    };
  }
  return { readable: true, reason: '' };
}

/**
 * Ищет в документе таблицу с графами товаров (по тексту ячеек УПД/счёт-фактура).
 * @return {{header:Array<string>, rows:Array<Array<string>>, width:number}|null}
 */
function extractMainGoodsTableFromDoc_(body) {
  const tables = [];
  collectTablesFromBody_(body, tables);
  let best = null;
  let bestScore = 0;
  for (let ti = 0; ti < tables.length; ti++) {
    const matrix = tableToMatrix_(tables[ti]);
    if (matrix.length < 2) {
      continue;
    }
    const score = scoreTableAsGoods_(matrix);
    if (score > bestScore) {
      bestScore = score;
      best = matrix;
    }
  }
  if (!best || bestScore < 5) {
    Logger.log('Таблица товаров в Doc не найдена (оценка лучшей ' + bestScore + ', всего таблиц ' + tables.length + ')');
    return null;
  }
  const split = splitHeaderAndDataFromMatrix_(best);
  return split;
}

function collectTablesFromBody_(body, out) {
  for (let i = 0; i < body.getNumChildren(); i++) {
    collectTablesFromElement_(body.getChild(i), out);
  }
}

function collectTablesFromElement_(el, out) {
  if (el.getType() === DocumentApp.ElementType.TABLE) {
    const table = el.asTable();
    out.push(table);
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        for (let k = 0; k < cell.getNumChildren(); k++) {
          collectTablesFromElement_(cell.getChild(k), out);
        }
      }
    }
    return;
  }
  if (el.getNumChildren && el.getNumChildren() > 0) {
    for (let i = 0; i < el.getNumChildren(); i++) {
      collectTablesFromElement_(el.getChild(i), out);
    }
  }
}

function tableToMatrix_(table) {
  const out = [];
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    const cells = [];
    for (let c = 0; c < row.getNumCells(); c++) {
      cells.push(
        row
          .getCell(c)
          .getText()
          .replace(/\u00a0/g, ' ')
          .replace(/\s+/g, ' ')
          .trim()
      );
    }
    out.push(cells);
  }
  return out;
}

function scoreTableAsGoods_(matrix) {
  const head = matrix
    .slice(0, Math.min(8, matrix.length))
    .map(function (row) {
      return row.join(' ');
    })
    .join('\n');
  let score = 0;
  if (/наименован/i.test(head)) {
    score += 12;
  }
  if (/код\s*товара|товар\/работ|работ,\s*услуг/i.test(head)) {
    score += 6;
  }
  if (/п\/п|№\s*п/i.test(head)) {
    score += 4;
  }
  if (/стоимость.*без\s*налог|без\s+налог.*всего/i.test(head)) {
    score += 6;
  }
  if (/единиц/i.test(head) && /измерен/i.test(head)) {
    score += 4;
  }
  let maxCols = 0;
  for (let i = 0; i < matrix.length; i++) {
    maxCols = Math.max(maxCols, matrix[i].length);
  }
  if (maxCols >= 12) {
    score += 8;
  } else if (maxCols >= 8) {
    score += 4;
  }
  return score;
}

function isProbableSubheaderRow_(row) {
  const j = row.join(' ');
  if (/наименован|стоимость\s+товаров\s*\(/i.test(j)) {
    return false;
  }
  if (/код|условное|цифров|краткое|обознач/i.test(j)) {
    return true;
  }
  let sum = 0;
  let n = 0;
  for (let i = 0; i < row.length; i++) {
    const t = (row[i] || '').length;
    if (t) {
      sum += t;
      n++;
    }
  }
  const avg = n ? sum / n : 0;
  return avg > 0 && avg < 22 && row.length >= 6;
}

function mergeHeaderRows_(r1, r2) {
  const w = Math.max(r1.length, r2.length);
  const out = [];
  for (let c = 0; c < w; c++) {
    const a = (r1[c] || '').trim();
    const b = (r2[c] || '').trim();
    if (a && b && a !== b) {
      out.push(a + ' — ' + b);
    } else {
      out.push(a || b);
    }
  }
  return out;
}

function splitHeaderAndDataFromMatrix_(matrix) {
  let stopIdx = matrix.length;
  for (let s = 0; s < matrix.length; s++) {
    if (/Всего\s+к\s+оплате/i.test(matrix[s].join(' '))) {
      stopIdx = s;
      break;
    }
  }
  const sub = matrix.slice(0, stopIdx);

  let hdrStart = -1;
  for (let i = 0; i < sub.length; i++) {
    const line = sub[i].join(' ');
    if (/наименован/i.test(line) && sub[i].length >= 4) {
      hdrStart = i;
      break;
    }
  }
  if (hdrStart === -1) {
    for (let j = 0; j < sub.length; j++) {
      const line2 = sub[j].join(' ');
      if (/код\s*товара|№\s*п\/п/i.test(line2) && sub[j].length >= 4) {
        hdrStart = j;
        break;
      }
    }
  }
  if (hdrStart === -1) {
    return null;
  }

  const headerRow1 = sub[hdrStart];
  let dataStart = hdrStart + 1;
  let header = headerRow1.slice();
  const nextRow = hdrStart + 1 < sub.length ? sub[hdrStart + 1] : null;
  if (nextRow && isProbableSubheaderRow_(nextRow)) {
    header = mergeHeaderRows_(headerRow1, nextRow);
    dataStart = hdrStart + 2;
  }

  const dataRows = [];
  for (let r = dataStart; r < sub.length; r++) {
    const lineText = sub[r].join(' ');
    if (/Всего\s+к\s+оплате/i.test(lineText)) {
      break;
    }
    if (/^Итого\b/i.test(lineText) && sub[r].length < Math.max(4, header.length / 2)) {
      continue;
    }
    const nonLetter = lineText.replace(/[\s\d.,\-–—/%№]/gi, '').length;
    if (nonLetter < 2 && sub[r].length < 4) {
      continue;
    }
    dataRows.push(sub[r]);
  }

  if (!dataRows.length) {
    return null;
  }

  const width = Math.max(header.length, maxRowLen_(dataRows));
  return {
    header: padRow_(header, width),
    rows: dataRows.map(function (x) {
      return padRow_(x, width);
    }),
    width: width,
  };
}

/**
 * @param {string} raw
 * @param {{header:Array,rows:Array,width:number}|null} docTable
 * @param {number} textLength
 * @param {boolean} [conversionOk]
 * @param {string} [conversionNote]
 */
function parseInvoiceData_(raw, docTable, textLength, conversionOk, conversionNote) {
  if (conversionOk === false) {
    const advice =
      ' Рекомендации: распознать текст в Acrobat/ABBYY и сохранить поисковый PDF; или выгрузить PDF из учётной системы с текстовым слоем; повёрнутые страницы — выпрямить до OCR. В Google — Document AI / Vision API; на ПК — Python (PyMuPDF, pytesseract).';
    return {
      invoiceLine: (conversionNote || 'Конвертация PDF→Google Doc не дала читаемый текст.') + advice,
      seller: '',
      paymentDoc: '',
      tableHeader: [],
      tableRows: [],
      basis: '',
      tableWidth: 0,
    };
  }

  const text = normalizeText_(raw);
  const textHint =
    !text || textLength < 80
      ? ' Мало текста после конвертации PDF (часто скан или «картинка»). Нужен OCR или PDF с текстовым слоем.'
      : '';

  let invoiceLine = extractInvoiceHeader_(text);
  if (!invoiceLine) {
    invoiceLine = extractInvoiceHeaderAlt_(text);
  }
  const seller = extractSeller_(text);
  const paymentDoc = extractPaymentDoc_(text);

  let table = null;
  if (docTable && docTable.rows && docTable.rows.length) {
    table = docTable;
    Logger.log('Таблица из Google Doc: строк данных ' + table.rows.length + ', колонок ' + table.width);
  } else {
    const tableBlock = extractTableBlock_(text);
    table = parseTableFromBlock_(tableBlock);
    Logger.log('Таблица из текста: строк ' + (table.rows ? table.rows.length : 0));
  }

  if ((!table || !table.rows.length) && textHint) {
    invoiceLine = (invoiceLine || '') + textHint.trim();
  }

  const basis = extractBasis_(text);

  const tw = table && table.width ? table.width : 0;
  return {
    invoiceLine: invoiceLine,
    seller: seller,
    paymentDoc: paymentDoc,
    tableHeader: table && table.header ? table.header : [],
    tableRows: table && table.rows ? table.rows : [],
    basis: basis,
    tableWidth: tw,
  };
}

function extractInvoiceHeaderAlt_(text) {
  const re2 = /Универсальн(ый|ое|ая)\s+передаточн(ый|ое|ая)\s+документ[^\n]{0,200}?№\s*([\s\S]{1,200}?)\s+от\s+([0-9]{2}\.[0-9]{2}\.[0-9]{4})/i;
  const m2 = text.match(re2);
  if (m2) {
    return ('УПД № ' + m2[2].replace(/\s+/g, ' ').trim() + ' от ' + m2[3].trim()).replace(/\s+/g, ' ');
  }
  return '';
}


function normalizeText_(t) {
  return t.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

function extractInvoiceHeader_(text) {
  const re = /Счет[-\s]*фактура\s*№\s*([\s\S]{1,400}?)\s+от\s+([0-9]{2}\.[0-9]{2}\.[0-9]{4})/i;
  const m = text.match(re);
  if (!m) {
    return '';
  }
  const num = m[1].replace(/\s+/g, ' ').trim();
  return ('Счет-фактура № ' + num + ' от ' + m[2].trim()).replace(/\s+/g, ' ');
}

/** Текст после слова «Продавец» (с двоеточием или без). */
function extractSeller_(text) {
  const m = text.match(/\bПродавец\s*:?\s*/i);
  if (!m || m.index === undefined) {
    return '';
  }
  const startContent = m.index + m[0].length;
  let rest = text.substring(startContent);
  const stop = rest.search(/\n(?=\s*(Покупатель|Плательщик|Грузоотправитель|Грузополучатель|К платежно|Идентификатор|Адрес|Счет[-\s]*фактура|Универсальн))/i);
  let chunk = stop === -1 ? rest : rest.substring(0, stop);
  chunk = chunk.replace(/\n+/g, ' ').trim();
  return chunk;
}

function extractAfterLabel_(text, label) {
  const idx = text.indexOf(label);
  if (idx === -1) {
    return '';
  }
  let rest = text.substring(idx + label.length);
  // до следующего известного заголовка или двойного перевода строки
  const stop = rest.search(/\n(?=\s*(Покупатель|Плательщик|Грузоотправитель|Грузополучатель|К платежно|Идентификатор|Адрес|Счет[-\s]*фактура|Универсальн))/i);
  let chunk = stop === -1 ? rest : rest.substring(0, stop);
  chunk = chunk.replace(/\n+/g, ' ').trim();
  return chunk;
}

function extractPaymentDoc_(text) {
  const re = /К\s+платежно[-\s]*расчетному\s+документу\s*№\s*([^\n\r]+)/i;
  const m = text.match(re);
  return m ? m[1].trim() : '';
}

function extractTableBlock_(text) {
  const marker = /К\s+платежно[-\s]*расчетному\s+документу\s*№/i;
  const m = text.match(marker);
  if (!m || m.index === undefined) {
    return text;
  }
  return text.substring(m.index);
}

function parseTableFromBlock_(block) {
  const lines = block.split('\n').map(function (l) {
    return l.replace(/\u00a0/g, ' ').trim();
  });

  let totalIdx = -1;
  for (let ti = 0; ti < lines.length; ti++) {
    if (/Всего\s+к\s+оплате/i.test(lines[ti])) {
      totalIdx = ti;
      break;
    }
  }
  const sliceEnd = totalIdx === -1 ? lines.length : totalIdx;
  const relevant = lines.slice(0, sliceEnd);

  // Ищем строку заголовка таблицы (УПД / счёт-фактура: много вариантов вёрстки)
  let headerIndex = -1;
  for (let i = 0; i < relevant.length; i++) {
    const l = relevant[i];
    if (
      /(наименован|наименование\s+товара)/i.test(l) &&
      (/№\s*п\/п|п\/п/i.test(l) || /код\s*товара/i.test(l) || /количество|кол-во/i.test(l) || /единиц/i.test(l))
    ) {
      headerIndex = i;
      break;
    }
  }
  if (headerIndex === -1) {
    for (let i = 0; i < relevant.length; i++) {
      const l = relevant[i];
      if (/(наименован|наименование\s+товара)/i.test(l) && l.length > 40) {
        headerIndex = i;
        break;
      }
    }
  }
  if (headerIndex === -1) {
    for (let i = 0; i < relevant.length; i++) {
      if (/\bкод\b/i.test(relevant[i]) && /наименован/i.test(relevant[i])) {
        headerIndex = i;
        break;
      }
    }
  }

  if (headerIndex === -1) {
    return { header: [], rows: [], width: 0 };
  }

  const headerLine = relevant[headerIndex];
  const header = splitTableLine_(headerLine);

  const dataRows = [];
  for (let j = headerIndex + 1; j < relevant.length; j++) {
    const line = relevant[j];
    if (!line) {
      continue;
    }
    if (/Всего\s+к\s+оплате/i.test(line)) {
      break;
    }
    if (/^Основание\s+передачи/i.test(line)) {
      break;
    }
    const cells = splitTableLine_(line);
    if (cells.length === 0) {
      continue;
    }
    // отсекаем «шапку» после таблицы
    if (/Итого|^Всего\b/i.test(line) && cells.length < header.length / 2) {
      continue;
    }
    dataRows.push(padRow_(cells, header.length));
  }

  const width = Math.max(header.length, maxRowLen_(dataRows));
  return {
    header: padRow_(header, width),
    rows: dataRows.map(function (r) {
      return padRow_(r, width);
    }),
    width: width,
  };
}

function splitTableLine_(line) {
  // Табы — самый частый разделитель после конвертации из Doc
  if (line.indexOf('\t') !== -1) {
    return line.split('\t').map(function (c) {
      return c.trim();
    });
  }
  // Fallback: два и более пробела
  return line
    .split(/\s{2,}/)
    .map(function (c) {
      return c.trim();
    })
    .filter(function (c) {
      return c.length > 0;
    });
}

function padRow_(cells, width) {
  const out = cells.slice();
  while (out.length < width) {
    out.push('');
  }
  return out;
}

function maxRowLen_(rows) {
  let m = 0;
  for (let i = 0; i < rows.length; i++) {
    m = Math.max(m, rows[i].length);
  }
  return m;
}

function extractBasis_(text) {
  const label = 'Основание передачи (сдачи) / получения (приемки)';
  const idx = text.indexOf(label);
  if (idx === -1) {
    return '';
  }
  const fromLabel = text.substring(idx + label.length).replace(/^[\s:\-–—]*/, '');
  const firstLine = (fromLabel.split('\n')[0] || fromLabel).replace(/\s+/g, ' ').trim();
  const accMatch = fromLabel.match(/Счет\s*№\s*([^\n\r]+)/i);
  if (accMatch) {
    return (firstLine + ' | Счет № ' + accMatch[1].trim()).trim();
  }
  return firstLine;
}

function writeParsedRows_(sheet, items, maxTableCols) {
  sheet.clearContents();

  let globalWidth = maxTableCols;
  for (let gi = 0; gi < items.length; gi++) {
    globalWidth = Math.max(globalWidth, items[gi].parsed.tableWidth);
  }

  const globalHeader = buildGlobalHeader_(items, globalWidth);
  sheet.getRange(1, 1, 1, globalHeader.length).setValues([globalHeader]);
  const tableCols = globalHeader.length - 5;

  let rowPtr = 2;
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    const p = it.parsed;

    if (p.tableRows.length === 0) {
      const single = [
        it.fileName,
        p.invoiceLine,
        p.seller,
        p.paymentDoc,
      ]
        .concat(padRow_(p.tableHeader, tableCols))
        .concat([p.basis]);
      sheet.getRange(rowPtr, 1, 1, globalHeader.length).setValues([padRow_(single, globalHeader.length)]);
      rowPtr++;
      continue;
    }

    for (let r = 0; r < p.tableRows.length; r++) {
      const tr = padRow_(p.tableRows[r], tableCols);
      const row = [
        r === 0 ? it.fileName : '',
        r === 0 ? p.invoiceLine : '',
        r === 0 ? p.seller : '',
        r === 0 ? p.paymentDoc : '',
      ]
        .concat(tr)
        .concat([r === 0 ? p.basis : '']);
      sheet.getRange(rowPtr, 1, 1, globalHeader.length).setValues([padRow_(row, globalHeader.length)]);
      rowPtr++;
    }
  }
}

function buildGlobalHeader_(items, globalWidth) {
  let wData = globalWidth;
  for (let i = 0; i < items.length; i++) {
    wData = Math.max(wData, items[i].parsed.tableWidth);
  }

  const base = [
    'Файл',
    'Счет-фактура (№ и дата)',
    'Продавец',
    'К платежно-расчетному документу №',
  ];

  if (USE_CANONICAL_TABLE_HEADERS) {
    const canonLen = CANONICAL_UPD_HEADERS.length;
    const w = Math.max(wData, canonLen);
    const cols = [];
    for (let c = 0; c < w; c++) {
      if (c < CANONICAL_UPD_HEADERS.length) {
        cols.push(CANONICAL_UPD_HEADERS[c]);
      } else {
        cols.push('Доп. столбец ' + (c + 1));
      }
    }
    return base.concat(cols).concat(['Основание передачи / счет']);
  }

  let sampleHeader = [];
  for (let j = 0; j < items.length; j++) {
    const h = items[j].parsed.tableHeader;
    if (h.length > sampleHeader.length) {
      sampleHeader = h.slice();
    }
  }
  const w = Math.max(wData, sampleHeader.length);
  sampleHeader = padRow_(sampleHeader, w);
  const cols = [];
  for (let c = 0; c < w; c++) {
    const name = sampleHeader[c] && sampleHeader[c].length ? sampleHeader[c] : 'Колонка ' + (c + 1);
    cols.push(name);
  }
  return base.concat(cols).concat(['Основание передачи / счет']);
}
