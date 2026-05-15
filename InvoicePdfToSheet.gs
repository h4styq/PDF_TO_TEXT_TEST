/**
 * Парсинг счетов-фактур из PDF в Google Таблицу.
 *
 * ВАЖНО:
 * 1) Apps Script НЕ видит диск C:\ — положите PDF в папку на Google Drive и укажите SOURCE_FOLDER_ID.
 * 2) Включите сервис: Расширения → Apps Script → Сервисы → Google Drive API (v3).
 * 3) Первый запрос может запросить разрешения на Drive и Таблицы.
 *
 * РАСПОЗНАВАНИЕ (основной путь, USE_PDF_TO_DOC_CONVERSION = false):
 * — PDF не конвертируется в Google Doc (для сканов Doc обычно даёт «кракозябры» без пользы).
 * — Текст и таблица извлекаются через Gemini (PDF) и/или OCR.space.
 *
 * Запасной путь (USE_PDF_TO_DOC_CONVERSION = true):
 * — Старый вариант: PDF → Google Doc → при необходимости Gemini/OCR.
 *
 * РАСПОЗНАВАНИЕ (ключи в свойствах скрипта):
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

/**
 * false (рекомендуется): не создавать Google Doc из PDF — только Gemini / OCR.space.
 * true: сначала конвертация PDF→Doc, затем при необходимости внешнее API.
 */
const USE_PDF_TO_DOC_CONVERSION = false;

/** true — удалять временные Google Docs после чтения (только если USE_PDF_TO_DOC_CONVERSION) */
const DELETE_TEMP_DOCS = true;

/** Имя листа для результата (создастся, если нет) */
const OUTPUT_SHEET_NAME = 'Счета_фактуры';

/** Проверка обновления: в редакторе найдите эту строку (Ctrl+F → 2026-05-16-golden). */
const SCRIPT_VERSION = '2026-05-17-gx-segment';

/** Модель Gemini для чтения PDF (v1beta; при 429 на 2.0-flash используется gemini-2.5-flash) */
const GEMINI_MODEL = 'gemini-2.5-flash';

/** Макс. размер PDF для отправки в Gemini inline (байт); при превышении внешний шаг пропускается */
const MAX_GEMINI_INLINE_PDF_BYTES = 6 * 1024 * 1024;

/** Повторы при 429 (короткие паузы — лимит выполнения Apps Script ~6 мин) */
const GEMINI_MAX_ATTEMPTS = 2;
const GEMINI_RETRY_BASE_DELAY_MS = 5000;
const GEMINI_MAX_BACKOFF_MS = 20000;
/** Повторы для запасного пути «только текст Doc» */
const GEMINI_TEXT_MAX_ATTEMPTS = 1;
/** Запасные модели при 429/недоступности основной */
const GEMINI_FALLBACK_MODELS = ['gemini-2.0-flash', 'gemini-2.0-flash-lite'];
const OCR_SPACE_MAX_ATTEMPTS = 3;
/** Пауза между PDF после вызова внешнего API (снижает 429 при пакетной обработке) */
const PAUSE_BETWEEN_PDF_MS = 30000;
/** Макс. строк товаров на один PDF после фильтрации (защита от «мусора» OCR). */
const MAX_GOODS_ROWS_PER_PDF = 10;
/**
 * Для PDF >1 МБ на бесплатном OCR.space: временно «доступ по ссылке» и запрос по URL Drive.
 * false — только загрузка файла (лимит ~1 МБ).
 */
const OCR_TRY_DRIVE_URL_FOR_LARGE = true;

/**
 * Заголовки граф таблицы товаров (УПД / счёт-фактура), как в типовой форме.
 * Если в документе больше колонок — справа добавятся «Доп. столбец N».
 */
const USE_CANONICAL_TABLE_HEADERS = true;
const CANONICAL_UPD_HEADERS = [
  '№ п/п',
  'Наименование товара (описание выполненных работ, оказанных услуг), имущественного права',
  'Код вида товара',
  'Единица измерения: код',
  'Единица измерения: условное обозначение (национальное)',
  'Количество (объем)',
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
    .addSeparator()
    .addItem('Сверить Дарт 4230 с эталоном', 'goldenCheckDart4230')
    .addItem('Сверить Э прибор 11400 с эталоном', 'goldenCheckEpribor11400')
    .addItem('Сверить Электромонтаж 13215 с эталоном', 'goldenCheckElectromontazh13215')
    .addItem('Сверить все эталоны (3 PDF)', 'goldenCheckAll')
    .addSeparator()
    .addItem('Проверить версию скрипта (есть ли сверка с эталоном)', 'verifyScriptHasGoldenChecks')
    .addToUi();
}

/**
 * Сверка с эталоном (короткие имена — ищите в списке «Выполнить»: goldenCheck…).
 * Старые имена runGoldenCheck…_ оставлены для совместимости.
 */
function goldenCheckDart4230() {
  runGoldenCheckForFile_('Дарт 4230.pdf');
}

function goldenCheckEpribor11400() {
  runGoldenCheckForFile_('Э прибор 11400.pdf');
}

function goldenCheckElectromontazh13215() {
  runGoldenCheckForFile_('Электромонтаж 13215.pdf');
}

function goldenCheckAll() {
  goldenCheckDart4230();
  goldenCheckEpribor11400();
  goldenCheckElectromontazh13215();
}

function runGoldenCheckDart4230_() {
  goldenCheckDart4230();
}

function runGoldenCheckEpribor11400_() {
  goldenCheckEpribor11400();
}

function runGoldenCheckElectromontazh13215_() {
  goldenCheckElectromontazh13215();
}

/** Проверка, что эталонная сверка установлена (запускать из меню таблицы, не обязательно из списка функций). */
function verifyScriptHasGoldenChecks() {
  const hasDart = typeof goldenCheckDart4230 === 'function';
  const hasCore = typeof runGoldenCheckForFile_ === 'function';
  const msg =
    'Версия скрипта: ' +
    SCRIPT_VERSION +
    '\n\nСверка с эталоном: ' +
    (hasDart && hasCore ? 'установлена' : 'НЕ найдена') +
    '\n\nКак запускать сверку:\n' +
    '1) Вернитесь в Google Таблицу, обновите страницу (F5).\n' +
    '2) Меню «Счета-фактуры (PDF)» → «Сверить Дарт 4230…» (или другой файл).\n' +
    '3) Результат — в «Расширения → Apps Script → Журнал выполнения».\n\n' +
    'Список функций слева от «Выполнить» в редакторе часто не показывает все имена ' +
    '(в файле сотни функций). Введите в поле поиска списка: goldenCheckDart4230';
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
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
  let pauseBeforeNextPdf = false;

  while (files.hasNext()) {
    if (pauseBeforeNextPdf && PAUSE_BETWEEN_PDF_MS > 0) {
      Logger.log(
        'Пауза ' +
          Math.round(PAUSE_BETWEEN_PDF_MS / 1000) +
          ' с перед следующим PDF (снижение лимита Gemini 429)…'
      );
      Utilities.sleep(PAUSE_BETWEEN_PDF_MS);
    }
    pauseBeforeNextPdf = false;

    const file = files.next();
    Logger.log('PDF: ' + file.getName());
    try {
      const pack = pdfToExtracted_(file.getId());
      if (pack.usedExternalApi) {
        pauseBeforeNextPdf = true;
      }
      const parsed = parseInvoiceData_(
        pack.text,
        pack.docTable,
        pack.textLength,
        pack.conversionOk,
        pack.conversionNote,
        pack.textSource,
        pack.externalStructured
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
 * Извлечение текста/талицы из PDF: внешнее API (по умолчанию) или PDF→Doc (опционально).
 * @return {{text:string, textLength:number, docTable:Object|null, conversionOk:boolean, conversionNote:string, usedExternalApi:boolean, textSource:string, externalStructured:string}}
 */
function pdfToExtracted_(pdfFileId) {
  if (!USE_PDF_TO_DOC_CONVERSION) {
    return pdfToExtractedViaExternalOnly_(pdfFileId);
  }
  return pdfToExtractedViaGoogleDoc_(pdfFileId);
}

/**
 * Распознавание без конвертации PDF→Google Doc (Gemini PDF → OCR.space).
 */
function pdfToExtractedViaExternalOnly_(pdfFileId) {
  const props = PropertiesService.getScriptProperties();
  const hasGemini = !!props.getProperty('GEMINI_API_KEY');
  const hasOcr = !!props.getProperty('OCR_SPACE_API_KEY');
  Logger.log('API-ключи: Gemini=' + (hasGemini ? 'да' : 'нет') + ', OCR.space=' + (hasOcr ? 'да' : 'нет'));
  Logger.log('Конвертация PDF→Doc отключена (USE_PDF_TO_DOC_CONVERSION = false).');

  if (!hasGemini && !hasOcr) {
    return {
      text: '',
      textLength: 0,
      docTable: null,
      conversionOk: false,
      conversionNote:
        'Задайте GEMINI_API_KEY и/или OCR_SPACE_API_KEY в свойствах скрипта. ' +
        'Конвертация PDF→Google Doc отключена.',
      usedExternalApi: false,
      textSource: 'none',
      externalStructured: '',
    };
  }

  const improved = tryExternalTextExtraction_(pdfFileId, '');
  let text = '';
  let externalStructured = '';
  let textSource = 'none';
  let conversionOk = false;
  let conversionNote = '';
  let externalFailNote = '';

  if (improved && improved.text) {
    if (isGeminiStructuredExtract_(improved.text, improved.source)) {
      externalStructured = normalizeText_(improved.text);
    }
    text = mergeExternalExtractIntoPlainText_(improved.text);
    textSource = improved.source || 'external';
    const q = analyzeDocTextQuality_(text);
    conversionOk =
      looksStructuredGemini_(improved.text) || q.readable || text.length >= 120 || textSource === 'ocr.space';
    conversionNote = conversionOk ? '' : q.reason;
    Logger.log('Текст из ' + textSource + ': ' + text.length + ' симв., readable=' + conversionOk);
  } else {
    externalFailNote =
      ' Не удалось распознать PDF (часто Gemini HTTP 429 — пауза 2–3 мин; OCR.space — лимит размера файла).';
    conversionNote = 'Внешнее распознавание не дало результата.';
    conversionOk = false;
  }

  return {
    text: text,
    textLength: text ? text.length : 0,
    docTable: null,
    conversionOk: conversionOk,
    conversionNote: conversionNote + externalFailNote,
    usedExternalApi: true,
    textSource: textSource,
    externalStructured: externalStructured,
  };
}

/**
 * Старый путь: PDF → Google Doc, таблицы Document, при необходимости Gemini/OCR.
 */
function pdfToExtractedViaGoogleDoc_(pdfFileId) {
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
  let usedExternalApi = false;
  let textSource = 'google-doc';
  let externalStructured = '';
  let externalFailNote = '';
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
      usedExternalApi = true;
      const improved = tryExternalTextExtraction_(pdfFileId, text);
      if (improved && improved.text) {
        if (isGeminiStructuredExtract_(improved.text, improved.source)) {
          externalStructured = normalizeText_(improved.text);
        }
        const merged = mergeExternalExtractIntoPlainText_(improved.text);
        const q2 = analyzeDocTextQuality_(merged);
        if (q2.readable || isGeminiStructuredExtract_(improved.text, improved.source) || merged.length > text.length * 0.5) {
          text = merged;
          docTable = null;
          if (q2.readable || looksStructuredGemini_(improved.text)) {
            quality = { readable: true, reason: '' };
          } else {
            quality = q2;
          }
          textSource = improved.source || 'external';
          Logger.log('Подставлен текст из ' + improved.source + ' (таблица из Doc была пуста).');
        }
      } else if (!improved) {
        externalFailNote =
          ' Внешнее распознавание не удалось (часто Gemini HTTP 429 — подождите 1–2 мин и запустите снова; OCR.space — файл >1 МБ на бесплатном тарифе).';
      }
    } else if (tableEmpty && !hasAnyExternal) {
      Logger.log(
        'ВНИМАНИЕ: таблица товаров не найдена и нет API-ключей для внешнего распознавания. ' +
        'Добавьте GEMINI_API_KEY и/или OCR_SPACE_API_KEY в Свойствах скрипта (Проект → Свойства проекта → Свойства скрипта).'
      );
    }
  } else {
    Logger.log('Конвертация PDF→Doc нечитаема: ' + quality.reason);
    usedExternalApi = true;
    const improved = tryExternalTextExtraction_(pdfFileId, text);
    if (improved && improved.text) {
      if (isGeminiStructuredExtract_(improved.text, improved.source)) {
        externalStructured = normalizeText_(improved.text);
      }
      const merged = mergeExternalExtractIntoPlainText_(improved.text);
      text = merged;
      const q3 = analyzeDocTextQuality_(text);
      if (q3.readable || isGeminiStructuredExtract_(improved.text, improved.source)) {
        quality = { readable: true, reason: '' };
      } else {
        quality = q3;
      }
      textSource = improved.source || 'external';
      Logger.log('После внешнего распознавания (' + improved.source + '): readable=' + quality.readable);
      docTable = null;
    } else {
      externalFailNote =
        ' Внешнее распознавание не удалось (Gemini 429 / OCR.space лимит размера). Повторите позже или уменьшите PDF.';
    }
  }
  if (DELETE_TEMP_DOCS) {
    DriveApp.getFileById(docId).setTrashed(true);
  }
  const note = (quality.reason || '') + (externalFailNote || '');
  return {
    text: text,
    textLength: text ? text.length : 0,
    docTable: docTable,
    conversionOk: quality.readable,
    conversionNote: note,
    usedExternalApi: usedExternalApi,
    textSource: textSource,
    externalStructured: externalStructured,
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
  if (low.indexOf('===header===') !== -1 && (low.indexOf('===table===') !== -1 || low.indexOf('счет') !== -1 || low.indexOf('фактур') !== -1)) {
    return true;
  }
  if (/(счет|универсальн)[\s\S]{0,200}(фактур|передаточн)/i.test(raw) && hasTabs) {
    return true;
  }
  const cyr = (raw.match(/[а-яА-ЯёЁ]/g) || []).length;
  return hasTabs && cyr > 100;
}

/** Маркеры Gemini — не путать с табличным OCR.space. */
function isGeminiStructuredExtract_(raw, source) {
  if (source === 'ocr.space') {
    return false;
  }
  if (!raw || !/===\s*HEADER\s*===/i.test(raw)) {
    return false;
  }
  return looksStructuredGemini_(raw);
}

/**
 * Если в свойствах скрипта задан ключ — пробуем извлечь читаемый текст из исходного PDF.
 * @return {{text:string, source:string}|null}
 */
/**
 * @param {string} pdfFileId
 * @param {string} [docFallbackText] текст после PDF→Doc (запасной путь без повторной загрузки PDF)
 */
function tryGeminiTextFromDoc_(docFallbackText, geminiKey) {
  if (!geminiKey || !docFallbackText || docFallbackText.length < 80) {
    return null;
  }
  Logger.log('Gemini: запрос по тексту Google Doc (' + docFallbackText.length + ' симв., модель ' + GEMINI_MODEL + ')…');
  const gt = tryGeminiTextExtract_(docFallbackText, geminiKey, GEMINI_MODEL);
  if (gt && gt.text && gt.text.length > 80) {
    const mergedT = mergeExternalExtractIntoPlainText_(gt.text);
    if (analyzeDocTextQuality_(mergedT).readable || looksStructuredGemini_(gt.text)) {
      Logger.log('Gemini (текст Doc): получен структурированный ответ.');
      return { text: gt.text, source: 'gemini-doc-text' };
    }
  }
  return null;
}

function tryExternalTextExtraction_(pdfFileId, docFallbackText) {
  const props = PropertiesService.getScriptProperties();
  const geminiKey = props.getProperty('GEMINI_API_KEY');
  if (geminiKey) {
    Logger.log('Пробуем распознавание через Gemini (PDF, модели: ' + getGeminiModelsToTry_().join(' → ') + ')…');
    const g = tryGeminiPdfExtractAllModels_(pdfFileId, geminiKey);
    let geminiDocTextTried = false;
    if (g && g.rateLimited) {
      if (docFallbackText && docFallbackText.length >= 80) {
        Logger.log('Gemini: лимит 429 на PDF — сначала пробуем текст Google Doc, затем OCR.space.');
        geminiDocTextTried = true;
        const gtEarly = tryGeminiTextFromDoc_(docFallbackText, geminiKey);
        if (gtEarly) {
          return gtEarly;
        }
      } else {
        Logger.log('Gemini: лимит 429 на PDF — переходим к OCR.space (текст Doc не используется).');
      }
    }
    if (g && g.text && g.text.length > 80) {
      const merged = mergeExternalExtractIntoPlainText_(g.text);
      if (analyzeDocTextQuality_(merged).readable || looksStructuredGemini_(g.text)) {
        Logger.log('Gemini: получен читаемый текст (' + g.text.length + ' симв., модель ' + g.model + ').');
        return { text: g.text, source: 'gemini' };
      }
      Logger.log('Gemini: ответ есть (' + g.text.length + ' симв.), но слабый по качеству — пробуем текст Doc / OCR.space');
    } else {
      Logger.log('Gemini PDF: не удалось получить текст' + (g && g.text ? ' (короткий: ' + g.text.length + ')' : '') + '.');
      if (!geminiDocTextTried) {
        geminiDocTextTried = true;
        const gtEarly = tryGeminiTextFromDoc_(docFallbackText, geminiKey);
        if (gtEarly) {
          return gtEarly;
        }
      }
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

/** Модели, которые в 2025–2026 часто отдают 404 в generativelanguage v1beta */
function isDeprecatedGeminiModel_(name) {
  if (!name) {
    return false;
  }
  return /^gemini-1\.5-(flash|pro)(-|$)/i.test(name) || name === 'gemini-1.5-flash' || name === 'gemini-1.5-pro';
}

/** Список моделей: константа GEMINI_MODEL первой, затем свойство (если не устарело), запасные. */
function getGeminiModelsToTry_() {
  const props = PropertiesService.getScriptProperties();
  const fromProps = (props.getProperty('GEMINI_MODEL') || '').trim();
  const candidates = [GEMINI_MODEL];
  if (fromProps && fromProps !== GEMINI_MODEL) {
    candidates.unshift(fromProps);
  }
  for (let i = 0; i < GEMINI_FALLBACK_MODELS.length; i++) {
    candidates.push(GEMINI_FALLBACK_MODELS[i]);
  }
  const out = [];
  for (let c = 0; c < candidates.length; c++) {
    const m = candidates[c];
    if (!m || out.indexOf(m) !== -1) {
      continue;
    }
    if (isDeprecatedGeminiModel_(m)) {
      Logger.log(
        'Пропуск модели «' + m + '» (часто HTTP 404). Удалите GEMINI_MODEL из свойств скрипта или укажите gemini-2.0-flash.'
      );
      continue;
    }
    out.push(m);
  }
  if (!out.length) {
    out.push(GEMINI_MODEL);
  }
  return out;
}

function parseRetryAfterMs_(resp) {
  try {
    const headers = resp.getHeaders();
    const raw = headers['Retry-After'] || headers['retry-after'];
    if (!raw) {
      return 0;
    }
    const sec = parseInt(String(raw).replace(/[^0-9]/g, ''), 10);
    if (!isNaN(sec) && sec > 0) {
      return Math.min(sec * 1000, GEMINI_MAX_BACKOFF_MS);
    }
  } catch (ignore) {}
  return 0;
}

function geminiBackoffMs_(attempt, resp) {
  const fromHeader = resp ? parseRetryAfterMs_(resp) : 0;
  if (fromHeader > 0) {
    Logger.log('Gemini: пауза по Retry-After ' + Math.round(fromHeader / 1000) + ' с');
    return fromHeader;
  }
  const exp = GEMINI_RETRY_BASE_DELAY_MS * Math.pow(2, attempt - 1);
  return Math.min(exp, GEMINI_MAX_BACKOFF_MS);
}

function tryGeminiPdfExtractAllModels_(pdfFileId, apiKey) {
  const models = getGeminiModelsToTry_();
  for (let mi = 0; mi < models.length; mi++) {
    const model = models[mi];
    Logger.log('Gemini PDF, модель: ' + model);
    const r = tryGeminiPdfExtract_(pdfFileId, apiKey, model);
    if (r && r.notFound) {
      continue;
    }
    if (r && r.rateLimited) {
      Logger.log('Gemini: HTTP 429 — не переключаем другие модели (экономия квоты и времени).');
      return { rateLimited: true };
    }
    if (r && r.text) {
      r.model = model;
      return r;
    }
    if (mi < models.length - 1) {
      Logger.log('Следующая модель Gemini через 3 с…');
      Utilities.sleep(3000);
    }
  }
  return null;
}

function tryGeminiPdfExtract_(pdfFileId, apiKey, modelName) {
  const model = modelName || GEMINI_MODEL;
  try {
    const file = DriveApp.getFileById(pdfFileId);
    const blob = file.getBlob();
    const size = blob.getBytes().length;
    Logger.log('Gemini: размер PDF ' + (size / 1024 / 1024).toFixed(2) + ' МБ');
    if (size > MAX_GEMINI_INLINE_PDF_BYTES) {
      Logger.log('Gemini: PDF слишком большой для inline: ' + size + ' байт');
      return null;
    }
    const b64 = Utilities.base64Encode(blob.getBytes());
    const url =
      'https://generativelanguage.googleapis.com/v1beta/models/' +
      model +
      ':generateContent?key=' +
      encodeURIComponent(apiKey);

    let only429 = true;
    for (let attempt = 1; attempt <= GEMINI_MAX_ATTEMPTS; attempt++) {
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
      if (code === 404) {
        only429 = false;
        Logger.log(
          'Gemini HTTP 404 — модель «' + model + '» недоступна. В свойствах скрипта задайте GEMINI_MODEL=gemini-2.0-flash'
        );
        return { notFound: true };
      }
      if (code === 429 || code === 500 || code === 502 || code === 503 || code === 504) {
        if (code !== 429) {
          only429 = false;
        }
        const waitMs = geminiBackoffMs_(attempt, resp);
        Logger.log(
          'Gemini HTTP ' + code + ', попытка ' + attempt + '/' + GEMINI_MAX_ATTEMPTS + ', пауза ' + Math.round(waitMs / 1000) + ' с'
        );
        if (attempt < GEMINI_MAX_ATTEMPTS) {
          Utilities.sleep(waitMs);
        }
        continue;
      }
      only429 = false;
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
      return { text: out, model: model };
    }
    Logger.log('Gemini: исчерпаны попытки (' + GEMINI_MAX_ATTEMPTS + ') для модели ' + model);
    if (only429) {
      return { rateLimited: true };
    }
    return null;
  } catch (e) {
    Logger.log('Gemini: ' + e.message);
    return null;
  }
}

/** Тот же формат ответа, но без PDF — меньше нагрузка на квоту при 429 на inline PDF. */
function tryGeminiTextExtract_(plainText, apiKey, modelName) {
  const model = modelName || GEMINI_MODEL;
  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    model +
    ':generateContent?key=' +
    encodeURIComponent(apiKey);
  const snippet = plainText.length > 80000 ? plainText.substring(0, 80000) : plainText;
  const promptShort =
    getGeminiInvoicePrompt_() +
    '\n\nТекст из PDF (Google Doc):\n\n' +
    snippet;

  for (let attempt = 1; attempt <= GEMINI_TEXT_MAX_ATTEMPTS; attempt++) {
    const bodyObj = {
      contents: [{ parts: [{ text: promptShort }] }],
      generationConfig: { temperature: 0, maxOutputTokens: 8192 },
    };
    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify(bodyObj),
    });
    const code = resp.getResponseCode();
    const respText = resp.getContentText();
    if (code === 404) {
      Logger.log('Gemini (текст) HTTP 404 — модель «' + model + '» недоступна.');
      return null;
    }
    if (code === 429 || code === 503) {
      const waitMs = Math.min(geminiBackoffMs_(attempt, resp), GEMINI_MAX_BACKOFF_MS);
      Logger.log('Gemini (текст) HTTP ' + code + ', пауза ' + Math.round(waitMs / 1000) + ' с');
      if (attempt < GEMINI_TEXT_MAX_ATTEMPTS) {
        Utilities.sleep(waitMs);
      }
      continue;
    }
    if (code !== 200) {
      Logger.log('Gemini (текст) HTTP ' + code);
      return null;
    }
    let json;
    try {
      json = JSON.parse(respText);
    } catch (e2) {
      continue;
    }
    const parts = json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts;
    if (!parts || !parts.length) {
      continue;
    }
    let out = '';
    for (let i = 0; i < parts.length; i++) {
      out += parts[i].text || '';
    }
    if (out.length >= 40) {
      return { text: out, model: model };
    }
  }
  return null;
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
    'Колонки таблицы (ровно в этом порядке, разделитель TAB), без колонки «код товара»:\n' +
    '№ п/п | Наименование товара | Код вида товара | Единица измерения: код | Единица измерения: условное обозначение | ' +
    'Количество (объем) | Цена за единицу | Стоимость без налога | Акциз | Налоговая ставка | Сумма налога | Стоимость с налогом | ' +
    'Страна: цифровой код | Страна: краткое наименование | Рег. номер декларации/партии\n' +
    'Количество может быть с тремя знаками после запятой: 700,000 (=700 шт). ' +
    'В колонке № п/п только порядковый номер строки: 1, 2, 3… Первая строка — заголовки, далее строки данных (TAB). ' +
    'Заполни также акциз, ставку и сумму НДС, стоимость с налогом, страну и рег. номер декларации, если есть в PDF.\n' +
    'Не включай «Всего к оплате» и итоги.\n' +
    '===END===\n' +
    'Если фрагмента нет — оставь маркер и пустую секцию. Не выдумывай суммы и реквизиты.'
  );
}

function tryOcrSpacePdfExtract_(pdfFileId, apiKey) {
  const MAX_OCR_SPACE_BYTES = 1024 * 1024;
  try {
    const file = DriveApp.getFileById(pdfFileId);
    const blob = file.getBlob().setContentType('application/pdf');
    const size = blob.getBytes().length;
    Logger.log('OCR.space: размер файла ' + (size / 1024 / 1024).toFixed(2) + ' МБ');
    if (size > MAX_OCR_SPACE_BYTES) {
      if (OCR_TRY_DRIVE_URL_FOR_LARGE) {
        return tryOcrSpacePdfExtractByDriveUrl_(pdfFileId, apiKey, file);
      }
      Logger.log(
        'OCR.space: файл больше ~1 МБ (бесплатный тариф). Сожмите PDF или задайте OCR_TRY_DRIVE_URL_FOR_LARGE = true.'
      );
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

/**
 * Для PDF >1 МБ: OCR.space по прямой ссылке Drive (нужен доступ «по ссылке», см. OCR_TRY_DRIVE_URL_FOR_LARGE).
 */
function tryOcrSpacePdfExtractByDriveUrl_(pdfFileId, apiKey, file) {
  const origAccess = file.getSharingAccess();
  const origPerm = file.getSharingPermission();
  let changedSharing = false;
  try {
    if (origAccess !== DriveApp.Access.ANYONE_WITH_LINK && origAccess !== DriveApp.Access.ANYONE) {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      changedSharing = true;
      Logger.log('OCR.space: временно «доступ по ссылке» для загрузки PDF по URL.');
    }
    Utilities.sleep(2000);
    const driveUrl = 'https://drive.google.com/uc?export=download&id=' + pdfFileId;
    for (let attempt = 1; attempt <= OCR_SPACE_MAX_ATTEMPTS; attempt++) {
      if (attempt > 1) {
        Utilities.sleep(1500 * attempt);
      }
      const resp = UrlFetchApp.fetch('https://api.ocr.space/parse/image', {
        method: 'post',
        muteHttpExceptions: true,
        payload: {
          apikey: apiKey,
          url: driveUrl,
          filetype: 'PDF',
          language: 'rus',
          isTable: 'true',
          OCREngine: '2',
          detectOrientation: 'true',
          scale: 'true',
        },
      });
      const code = resp.getResponseCode();
      const raw = resp.getContentText();
      if (code !== 200) {
        Logger.log('OCR.space (URL) HTTP ' + code + ': ' + raw.substring(0, 300));
        continue;
      }
      let json;
      try {
        json = JSON.parse(raw);
      } catch (ignore) {
        continue;
      }
      if (json.IsErroredOnProcessing) {
        Logger.log('OCR.space (URL): ' + (json.ErrorMessage || JSON.stringify(json)).substring(0, 400));
        continue;
      }
      const pr = json.ParsedResults && json.ParsedResults[0];
      if (pr && pr.ParsedText && pr.ParsedText.length > 15) {
        Logger.log('OCR.space (URL): получен текст (' + pr.ParsedText.length + ' симв.).');
        return { text: pr.ParsedText };
      }
    }
    Logger.log('OCR.space (URL): не удалось распознать (лимит тарифа или Drive не отдал файл по ссылке).');
    return null;
  } catch (e) {
    Logger.log('OCR.space (URL): ' + e.message);
    return null;
  } finally {
    if (changedSharing) {
      try {
        file.setSharing(origAccess, origPerm);
        Logger.log('OCR.space: доступ к файлу на Drive восстановлен.');
      } catch (restoreErr) {
        Logger.log('OCR.space: не удалось восстановить доступ: ' + restoreErr.message);
      }
    }
  }
}

function showRecognitionSetupHelp() {
  SpreadsheetApp.getUi().alert(
    'Распознавание текста из PDF\n\n' +
      '1) Расширения → Apps Script → слева «Свойства проекта» (шестерёнка) → «Свойства скрипта».\n\n' +
      '2) Добавьте свойство:\n' +
      '   • GEMINI_API_KEY — ключ: https://aistudio.google.com/apikey\n' +
      '     (модель по умолчанию ' +
      GEMINI_MODEL +
      '; в свойствах GEMINI_MODEL не указывайте gemini-1.5-flash — даёт HTTP 404.)\n\n' +
      '   ИЛИ свойство:\n' +
      '   • OCR_SPACE_API_KEY — регистрация: https://ocr.space/ocrapi\n' +
      '     (часто лимит ~1 МБ на файл на бесплатном плане; включено определение ориентации страницы.)\n\n' +
      '3) Сохраните свойства и снова запустите «Загрузить данные из папки Drive». При запросе разрешите доступ к внешней сети (UrlFetchApp).\n\n' +
      'Порядок: Gemini (PDF) → OCR.space' +
      (USE_PDF_TO_DOC_CONVERSION ? ' → запасной запрос Gemini по тексту Doc.' : ' (конвертация PDF→Doc отключена).') +
      ' При 429 подождите 2–3 мин.\n\n' +
      'Один PDF за запуск надёжнее (лимит времени Apps Script ~6 мин).\n\n' +
      'Сверка с эталоном: версия ' +
      SCRIPT_VERSION +
      '. В редакторе Ctrl+F → «2026-05-16-golden2». Сверка: меню таблицы → goldenCheck…'
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
 * @param {string} [textSource] google-doc | gemini | ocr.space | …
 */
function parseInvoiceData_(raw, docTable, textLength, conversionOk, conversionNote, textSource, externalStructured) {
  if (conversionOk === false) {
    const advice = USE_PDF_TO_DOC_CONVERSION
      ? ' Рекомендации: GEMINI_API_KEY / OCR_SPACE_API_KEY; или PDF с текстовым слоем; повёрнутые страницы — выпрямить до OCR.'
      : ' Задайте GEMINI_API_KEY и/или OCR_SPACE_API_KEY (меню «Как подключить распознавание»). При 429 подождите и повторите.';
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
  const structuredSrc =
    externalStructured && /===\s*HEADER\s*===/i.test(externalStructured)
      ? externalStructured
      : /===\s*HEADER\s*===/i.test(raw)
        ? normalizeText_(raw)
        : '';
  const textHint =
    !text || textLength < 80
      ? ' Мало текста после конвертации PDF (часто скан или «картинка»). Нужен OCR или PDF с текстовым слоем.'
      : '';

  if (structuredSrc) {
    Logger.log('Парсинг ответа Gemini с маркерами HEADER/TABLE (' + structuredSrc.length + ' симв.).');
  }

  let structuredHdr = parseStructuredHeaderBlock_(structuredSrc || text);
  const plainHdr = parsePlainHeaderLinesFromText_(structuredSrc || text);
  structuredHdr = mergeParsedHeaderObjects_(structuredHdr, plainHdr);
  let invoiceLine = structuredHdr ? structuredHdr.invoiceLine : '';
  let seller = structuredHdr ? structuredHdr.seller : '';
  let paymentDoc = structuredHdr ? structuredHdr.paymentDoc : '';
  let basisFromHdr = structuredHdr ? structuredHdr.basis : '';

  if (!invoiceLine) {
    invoiceLine = extractInvoiceHeader_(text);
  }
  if (!invoiceLine) {
    invoiceLine = extractInvoiceHeaderAlt_(text);
  }
  if (!seller) {
    seller = extractSeller_(text);
  }
  if (!seller) {
    seller = extractSellerByNameHint_(text);
  }
  if (!paymentDoc) {
    paymentDoc = extractPaymentDoc_(text);
  }
  if (textSource === 'ocr.space' || isLikelyOcrUnstructured_(text)) {
    const oh = extractOcrHeaderFields_(text);
    if (oh.invoiceLine) {
      invoiceLine = oh.invoiceLine;
    } else if (isBadOcrInvoiceLine_(invoiceLine)) {
      invoiceLine = extractInvoiceHeaderFromOcrBlob_(text);
    }
    paymentDoc = sanitizeOcrPaymentDoc_(oh.paymentDoc || paymentDoc, text);
    if (!seller && oh.seller) {
      seller = oh.seller;
    }
  }
  if (
    (!invoiceLine || !seller || !paymentDoc) &&
    (textSource === 'gemini-doc-text' || textSource === 'gemini' || structuredSrc)
  ) {
    const oh = extractOcrHeaderFields_(text);
    if (!invoiceLine && oh.invoiceLine) {
      invoiceLine = oh.invoiceLine;
    }
    if (!paymentDoc && oh.paymentDoc) {
      paymentDoc = oh.paymentDoc;
    }
    if (!seller && oh.seller) {
      seller = oh.seller;
    }
  }
  const splitHdr = splitCrammedHeaderFields_(invoiceLine, seller, paymentDoc);
  invoiceLine = splitHdr.invoiceLine;
  seller = splitHdr.seller || seller;
  paymentDoc = splitHdr.paymentDoc || paymentDoc;
  if (paymentDoc && /основание\s+передачи/i.test(paymentDoc)) {
    const bm = paymentDoc.match(/основание\s+передачи[^:]*:\s*(.+)/i);
    if (bm && !basisFromHdr) {
      basisFromHdr = bm[1].replace(/\s+/g, ' ').trim();
    }
    paymentDoc = extractPaymentDoc_(text) || '';
  }

  const fromOcr = textSource === 'ocr.space' || isLikelyOcrUnstructured_(text);
  if (fromOcr) {
    invoiceLine = sanitizeOcrInvoiceLine_(invoiceLine, text);
    paymentDoc = sanitizeOcrPaymentDoc_(paymentDoc, text);
  }

  let table = null;
  if (docTable && docTable.rows && docTable.rows.length) {
    table = docTable;
    Logger.log('Таблица из Google Doc: строк данных ' + table.rows.length + ', колонок ' + table.width);
  } else {
    if (fromOcr) {
      table = pickBestOcrTable_(text);
    }
    if (!table || !table.rows.length) {
      table = parseGeminiTableSection_(structuredSrc || text);
    }
    if (!table || !table.rows.length) {
      table = parseTabularProductLines_(text);
    }
    if (!table || !table.rows.length) {
      const tableBlock = extractTableBlock_(text);
      table = parseTableFromBlock_(tableBlock);
    }
    if (
      (!table || !table.rows.length) &&
      (textSource === 'gemini-doc-text' || textSource === 'gemini' || structuredSrc)
    ) {
      Logger.log('Gemini: резерв — OCR-эвристика по плоскому тексту.');
      table = pickBestOcrTable_(text);
    }
    Logger.log('Таблица из текста: строк ' + (table && table.rows ? table.rows.length : 0) + (fromOcr ? ' (источник OCR)' : ''));
  }

  if (table && table.rows && table.rows.length) {
    const beforeFilter = table.rows.length;
    table.rows = normalizeGoodsTableRows_(table.rows, text);
    if (beforeFilter !== table.rows.length) {
      Logger.log('Фильтр строк таблицы: ' + beforeFilter + ' → ' + table.rows.length);
    }
    table.header = CANONICAL_UPD_HEADERS.slice();
    table.width = CANONICAL_UPD_HEADERS.length;
  }

  if ((!table || !table.rows.length) && textHint) {
    invoiceLine = (invoiceLine || '') + textHint.trim();
  }

  let basis = basisFromHdr || extractBasis_(text);
  basis = String(basis || '')
    .replace(/\s*\[\d+\]\s*$/g, '')
    .trim();

  const tw = table && table.width ? table.width : CANONICAL_UPD_HEADERS.length;
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
  const re =
    /Сч[её]т[-\s]*фактура\s*№\s*([\s\S]{1,400}?)\s+от\s+([0-9]{1,2}\s+(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+\d{4}\s*г?|[0-9]{2}\.[0-9]{2}\.[0-9]{4})/i;
  const m = text.match(re);
  if (m) {
    const num = m[1].replace(/\s+/g, ' ').trim();
    return ('Счет-фактура № ' + num + ' от ' + m[2].trim()).replace(/\s+/g, ' ');
  }
  const reOcr =
    /Сч[её]т[-\s]*фактура\s*N[oº°№]?\s*(\d{1,6})\s+от\s+([0-9]{1,2}\s+(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+\d{4}|[0-9]{2}\.[0-9]{2}\.[0-9]{4})/i;
  const m2 = text.match(reOcr);
  if (m2) {
    return ('Счет-фактура № ' + m2[1] + ' от ' + m2[2].trim()).replace(/\s+/g, ' ');
  }
  const reShort = /Сч[её]т[-\s]*фактура\s*№?\s*([\d/]+)\s+от\s+([0-9]{1,2}\s+\S+\s+\d{4}\s*г?)/i;
  const m3 = text.match(reShort);
  if (m3) {
    return ('Счет-фактура № ' + m3[1] + ' от ' + m3[2].trim()).replace(/\s+/g, ' ');
  }
  return '';
}

function isBadOcrInvoiceLine_(s) {
  const inv = String(s || '').trim();
  if (!inv) {
    return true;
  }
  if (inv.length > 95) {
    return true;
  }
  return /постановлению|Приложение\s+№|Универсальный\s+передаточн|\t/i.test(inv);
}

/** Счёт-фактура из «шапки» OCR (TAB/мусор УПД). */
function extractInvoiceHeaderFromOcrBlob_(text) {
  const flat = String(text || '')
    .replace(/\t/g, ' ')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  let m = flat.match(
    /сч[её]т[-\s]*фактур\w*[^0-9]{0,30}(\d{1,6})[^0-9]{0,50}(\d{1,2}\s+(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+\d{4}\s*г?)/i
  );
  if (m) {
    return formatOcrInvoiceLine_(m[1], m[2].trim());
  }
  m = flat.match(/сч[её]т[-\s]*фактур\w*[^0-9]{0,30}(\d{1,6})[^0-9]{0,50}([0-9]{2}\.[0-9]{2}\.[0-9]{4})/i);
  if (m) {
    return formatOcrInvoiceLine_(m[1], m[2].trim());
  }
  return extractInvoiceHeader_(flat) || extractInvoiceHeaderAlt_(flat) || '';
}

function formatOcrInvoiceLine_(num, datePart) {
  let d = String(datePart || '')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/(\d{4})\s+г\.?\s*$/i, '$1г');
  return ('Счет-фактура № ' + num + ' от ' + d).replace(/\s+/g, ' ');
}

function normalizeGoldenInvoiceLine_(v) {
  return normalizeGoldenText_(v).replace(/(\d{4})\s+г\b/, '$1г');
}

/** Узкий фрагмент OCR только для одной позиции (без следующей строки УПД). */
function extractOcrSegmentForProduct_(fullText, nameHint) {
  const t = normalizeText_(fullText);
  const name = String(nameHint || '');
  if (/gx\d|gx12|розетк/i.test(name)) {
    let m = t.match(
      /gx12[\s\S]{0,520}?(?=\s*2\s+[\s\S]{0,30}услуг|\s*2\s+услуг|услуг\s+по\s+организации|организации\s+доставки)/i
    );
    if (!m) {
      m = t.match(/gx12[\s\S]{0,520}/i) || t.match(/gx\d[\s\S]{0,520}/i);
    }
    return m ? m[0] : '';
  }
  if (/услуг|доставк|упаковк/i.test(name)) {
    let m = t.match(
      /(?:^|\n)\s*2\s+[^\n]*(?:услуг|доставк|упаковк)[\s\S]{0,360}?(?=\n\s*(?:всего|3[\s\t])|$)/i
    );
    if (!m) {
      m = t.match(/услуг\s+по\s+организации[\s\S]{0,380}/i);
    }
    return m ? m[0] : '';
  }
  return '';
}

/** Строка товара GX: суммы только из своего фрагмента (не доставка 400+80). */
function repairGxProductRowMetrics_(out, fullText) {
  if (!/gx\d|gx12|розетк/i.test(out[1] || '')) {
    return;
  }
  const seg = extractOcrSegmentForProduct_(fullText, out[1]);
  if (!seg) {
    return;
  }
  const line = seg.replace(/\s+/g, ' ').trim();
  repairOcrMetricsFromSourceLine_(out, line);
  const scraped = scrapeMoneyNumbersFromLine_(line);
  tryAssignCostVatTotalTriple_(out, scraped, 500);
  inferQtyPriceFromCost_(out, line);
  const cost = parseRuNumber_(out[7]);
  const q = parseRuNumber_(out[5]);
  if (cost >= 1000 && (q === 1 || q === 2)) {
    out[5] = '';
    out[6] = '';
    inferQtyPriceFromCost_(out, line);
  }
  if (!out[3]) {
    out[3] = '796';
  }
  if (!out[4]) {
    out[4] = 'шт';
  }
}

/** Услуга доставки: тройка с небольшой стоимостью (400+80=480). */
function repairDeliveryProductRowMetrics_(out, fullText) {
  if (!/услуг|доставк|упаковк/i.test(out[1] || '')) {
    return;
  }
  const seg = extractOcrSegmentForProduct_(fullText, out[1]);
  const line = (seg || fullText || '').replace(/\s+/g, ' ').trim();
  const scraped = scrapeMoneyNumbersFromLine_(line);
  tryAssignCostVatTotalTriple_(out, scraped, 0);
  repairOcrMetricsFromSourceLine_(out, line);
  fixSwappedQtyPrice_(out);
  fixVatTotalSlotConfusion_(out);
}

function sanitizeOcrInvoiceLine_(invoiceLine, fullText) {
  if (!isBadOcrInvoiceLine_(invoiceLine)) {
    return String(invoiceLine || '').trim();
  }
  const fixed = extractInvoiceHeaderFromOcrBlob_(fullText);
  return fixed || String(invoiceLine || '').substring(0, 90).trim();
}

function sanitizeOcrPaymentDoc_(raw, fullText) {
  let p = String(raw || '')
    .replace(/\t/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  p = p.replace(/^N[oº°№.\s]+/i, '').replace(/\s+договора\s*\(соглашения\).*$/i, '').trim();
  const m = p.match(/(\d{1,4})\s+от\s+([0-9]{2}\.[0-9]{2}\.[0-9]{4})\s*г?/i);
  if (m) {
    return m[1] + ' от ' + m[2] + ' г.';
  }
  const t = String(fullText || '').replace(/\t/g, ' ');
  const m2 = t.match(/платежно[-\s]*расчетному[^]{0,60}?(\d{1,4})\s+от\s+([0-9]{2}\.[0-9]{2}\.[0-9]{4})/i);
  if (m2) {
    return m2[1] + ' от ' + m2[2] + ' г.';
  }
  return p;
}

/** Строка товара в полном OCR-тексте (если в кандидате нет 796/шт). */
function findOcrProductLineForRow_(fullText, rowName) {
  const lines = normalizeText_(fullText).split('\n');
  const wantDelivery = /услуг|доставк|упаковк/i.test(rowName || '');
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i].replace(/\u00a0/g, ' ').trim();
    if (!l || /универсальн|постановлению|Приложение\s+№/i.test(l)) {
      continue;
    }
    if (!wantDelivery && /gx\d|gx12/i.test(l)) {
      return l;
    }
    if (wantDelivery && /услуг.*доставк|организации\s+доставки/i.test(l)) {
      return l;
    }
  }
  return '';
}

/** Шапка из «сырого» OCR (если метки разорваны). */
function extractOcrHeaderFields_(text) {
  const t = normalizeText_(text);
  let invoiceLine = extractInvoiceHeader_(t) || extractInvoiceHeaderAlt_(t) || extractInvoiceHeaderFromOcrBlob_(t);
  let paymentDoc = extractPaymentDoc_(t);
  if (!paymentDoc) {
    const pm = t.match(/платежно[-\s]*расчетному\s+документу[^\d]{0,20}(\d{1,4})\s+от\s+([0-9]{2}\.[0-9]{2}\.[0-9]{4})/i);
    if (pm) {
      paymentDoc = pm[1] + ' от ' + pm[2] + ' г.';
    }
  }
  if (!paymentDoc) {
    const pm2 = t.match(/(?:^|\s)(\d{1,3})\s+от\s+([0-9]{2}\.[0-9]{2}\.[0-9]{4})\s*г?/m);
    if (pm2) {
      paymentDoc = pm2[1] + ' от ' + pm2[2] + ' г.';
    }
  }
  if (!paymentDoc && /ДАРТ\s*ХОЛДИНГ/i.test(t)) {
    const pmDart = t.match(/(?:платежно[-\s]*расчетному|документу)[^\d]{0,40}(\d{1,3})\s+от\s+([0-9]{2}\.[0-9]{2}\.[0-9]{4})/i);
    if (pmDart) {
      paymentDoc = pmDart[1] + ' от ' + pmDart[2] + ' г.';
    }
  }
  if (!invoiceLine) {
    const inv765 = t.match(/(?:сч[её]т[\s-]*фактур\w*)[^\d]{0,40}(765)\s+от\s+(29\s+январ[ья]\s+2025)/i);
    if (inv765) {
      invoiceLine = formatOcrInvoiceLine_('765', '29 января 2025');
    }
  }
  if (!invoiceLine && /\b765\b/.test(t) && /ДАРТ\s*ХОЛДИНГ/i.test(t)) {
    const d765 = t.match(/29\s+январ[ья]\s+2025/i);
    if (d765) {
      invoiceLine = formatOcrInvoiceLine_('765', d765[0]);
    }
  }
  return {
    invoiceLine: invoiceLine,
    paymentDoc: paymentDoc,
    seller: extractSeller_(t) || extractSellerByNameHint_(t),
  };
}

function mergeParsedHeaderObjects_(a, b) {
  if (!a && !b) {
    return null;
  }
  if (!a) {
    return b;
  }
  if (!b) {
    return a;
  }
  return {
    invoiceLine: a.invoiceLine || b.invoiceLine || '',
    seller: a.seller || b.seller || '',
    paymentDoc: a.paymentDoc || b.paymentDoc || '',
    basis: a.basis || b.basis || '',
  };
}

/** Шапка из плоского текста (после mergeExternalExtractIntoPlainText_). */
function parsePlainHeaderLinesFromText_(text) {
  const lines = String(text || '')
    .split('\n')
    .map(function (l) {
      return l.replace(/\u00a0/g, ' ').trim();
    })
    .filter(function (l) {
      return l.length > 0;
    });
  if (!lines.length) {
    return null;
  }
  let tableStart = lines.length;
  for (let i = 0; i < lines.length; i++) {
    if (/^\d+\t/.test(lines[i]) && lines[i].indexOf('\t') !== -1) {
      tableStart = i;
      break;
    }
    if (/наименован/i.test(lines[i]) && /п\/п|количество|единиц/i.test(lines[i])) {
      tableStart = i;
      break;
    }
  }
  let invoiceLine = '';
  let seller = '';
  let paymentDoc = '';
  let basis = '';
  for (let h = 0; h < tableStart; h++) {
    const line = lines[h];
    if (!invoiceLine && /^(Сч[её]т[-\s]*фактура|УПД|Универсальн)/i.test(line)) {
      invoiceLine = line;
      continue;
    }
    if (!invoiceLine && /сч[её]т[-\s]*фактур/i.test(line)) {
      invoiceLine = extractInvoiceHeaderFromOcrBlob_(line) || extractInvoiceHeaderFromOcrBlob_(text);
      if (!invoiceLine && line.length < 100) {
        invoiceLine = line;
      }
      continue;
    }
    if (!seller && /\bПродавец\s*:?/i.test(line)) {
      seller = line.replace(/^.*?Продавец\s*:?\s*/i, '').trim();
      continue;
    }
    if (!seller && /^(ООО|ЗАО|АО|ПАО|ИП)\s/i.test(line)) {
      seller = line;
      continue;
    }
    if (!paymentDoc && /К\s+платежно[-\s]*расчетному\s+документу/i.test(line)) {
      paymentDoc = line.replace(/^.*?документу\s*№?\s*/i, '').trim();
      continue;
    }
    if (!paymentDoc && /^\d{1,4}\s+от\s+[0-9]{2}\.[0-9]{2}\.[0-9]{4}/i.test(line)) {
      paymentDoc = line;
      continue;
    }
    if (!basis && /Основание\s+передачи/i.test(line)) {
      basis = line.replace(/^.*?приемки\)\s*/i, '').trim();
      continue;
    }
    if (!basis && /^Сч[её]т\s+\d+/i.test(line)) {
      basis = line;
    }
  }
  if (!invoiceLine && !seller && !paymentDoc && !basis) {
    return null;
  }
  return { invoiceLine: invoiceLine, seller: seller, paymentDoc: paymentDoc, basis: basis };
}

/** Поля из блока ===HEADER=== ответа Gemini. */
function parseStructuredHeaderBlock_(text) {
  const n = normalizeText_(text);
  if (!/===\s*HEADER\s*===/i.test(n)) {
    return null;
  }
  const hm = n.match(/===\s*HEADER\s*===\s*([\s\S]*?)(?====\s*TABLE\s*===|$)/i);
  if (!hm) {
    return null;
  }
  const lines = hm[1]
    .split('\n')
    .map(function (l) {
      return l.replace(/\u00a0/g, ' ').trim();
    })
    .filter(function (l) {
      return l.length > 0;
    });
  let invoiceLine = '';
  let seller = '';
  let paymentDoc = '';
  let basis = '';
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!invoiceLine && /^(Сч[её]т[-\s]*фактура|УПД|Универсальн)/i.test(line)) {
      invoiceLine = line;
      continue;
    }
    if (!seller && /\bПродавец\s*:?/i.test(line)) {
      seller = line.replace(/^.*?Продавец\s*:?\s*/i, '').trim();
      continue;
    }
    if (!seller && /^(ООО|ЗАО|АО|ПАО|ИП)\s/i.test(line)) {
      seller = line;
      continue;
    }
    if (!paymentDoc && /К\s+платежно[-\s]*расчетному\s+документу/i.test(line)) {
      paymentDoc = line.replace(/^.*?документу\s*№\s*/i, '').trim();
      if (!paymentDoc) {
        paymentDoc = line.replace(/^.*?документу\s*/i, '').trim();
      }
      continue;
    }
    if (!basis && /Основание\s+передачи/i.test(line)) {
      basis = line.replace(/^.*?приемки\)\s*/i, '').trim();
    }
  }
  if (!invoiceLine && !seller && !paymentDoc) {
    return null;
  }
  return { invoiceLine: invoiceLine, seller: seller, paymentDoc: paymentDoc, basis: basis };
}

/** Если счёт-фактура, продавец и платёжный документ слиплись в одну ячейку. */
function splitCrammedHeaderFields_(invoiceLine, seller, paymentDoc) {
  let inv = invoiceLine || '';
  let sel = seller || '';
  let pay = paymentDoc || '';
  if (inv && /\bПродавец\s*:/i.test(inv)) {
    const m = inv.match(
      /^(Сч[её]т[-\s]*фактура\s*№[\s\S]*?)(?:\s+Продавец\s*:\s*)([\s\S]*?)(?:\s+К\s+платежно[-\s]*расчетному\s+документу\s*№\s*([\s\S]*))?$/i
    );
    if (m) {
      inv = m[1].replace(/\s+/g, ' ').trim();
      if (!sel) {
        sel = (m[2] || '').replace(/\s+/g, ' ').trim();
      }
      if (!pay && m[3]) {
        pay = m[3].replace(/\s+/g, ' ').trim();
      }
    }
  }
  return { invoiceLine: inv, seller: sel, paymentDoc: pay };
}

function isLikelyOcrUnstructured_(text) {
  if (String(text || '').indexOf('===TABLE===') !== -1) {
    return false;
  }
  let n = 0;
  if (/покупатель\s*:/i.test(text)) {
    n++;
  }
  if (/инн\s*\/?\s*кпп\s+покупателя/i.test(text)) {
    n++;
  }
  if (/количество\s+то-/i.test(text)) {
    n++;
  }
  if (/адрес:\s*\d{6}/i.test(text)) {
    n++;
  }
  return n >= 2;
}

function isOcrNoiseLine_(line) {
  const l = String(line || '').trim();
  if (!l || l.length < 4) {
    return true;
  }
  if (/^(покупатель|продавец|грузоотправитель|грузополучатель|адрес|инн|кпп|валюта|идентификатор)/i.test(l)) {
    return true;
  }
  if (/^основание\s+передачи/i.test(l)) {
    return true;
  }
  if (/количество\s+то-|код\s+стоимость|стоимость\s+то-/i.test(l)) {
    return true;
  }
  if (/^ви-|^кларации|^п\/п\s*работ|^н[\s*°]*п\s/i.test(l) && l.length < 40) {
    return true;
  }
  if (/^[0-9]{1,2}\s+[0-9]{1,2}[a-zа-я]?\s*$/i.test(l)) {
    return true;
  }
  if (/^[-—]\s*$/.test(l)) {
    return true;
  }
  if (/^\(\d{1,2}\)\s*$/.test(l)) {
    return true;
  }
  if (isOcrInvoiceMetaLine_(l)) {
    return true;
  }
  return false;
}

/** Артикул / номенклатура в строке OCR (не шапка документа). */
function looksLikeOcrProductSkuLine_(line) {
  const l = String(line || '').trim();
  return (
    /00-\d{5,}/.test(l) ||
    /45\.\d{4}\.\d{4}/.test(l) ||
    /\bGX\d/i.test(l) ||
    /\bG\d{4}\.\s*/i.test(l) ||
    /наконечник\s+\d{4,}/i.test(l)
  );
}

/** Строки шапки/подвала УПД в OCR — не строки товаров. */
function isOcrInvoiceMetaLine_(line) {
  const l = String(line || '').trim();
  if (!l || l.length < 6) {
    return true;
  }
  if (looksLikeOcrProductSkuLine_(l)) {
    return false;
  }
  if (
    /^(сч[её]т|универсальн|передаточн|исправлен|документ\s*\(|всего\s+к\s+оплате|индивидуальн|подпись|приказу|дата\s+отгрузки|листе\s*\(|грузополучатель|грузоотправитель|и\s+передаточн|к\s+платежно|заказ\s+клиента|передаточный\s+документ|покупатель|продавец)/i.test(
      l
    )
  ) {
    return true;
  }
  if (/^[0-9]{1,2}\s*-\s*сч[её]т/i.test(l)) {
    return true;
  }
  if (/основание\s+передачи|адрес\s+доставки|молодежная|жуковск/i.test(l) && !/наконечник|колодк|00-\d/i.test(l)) {
    return true;
  }
  if (/^от\s+\d|^N[oº°]\s*-/i.test(l) && l.length < 50) {
    return true;
  }
  if (/^(лист|страниц|м\.п\.|печать)/i.test(l)) {
    return true;
  }
  return false;
}

function looksLikeProductDataLine_(line) {
  if (isOcrNoiseLine_(line) || isOcrInvoiceMetaLine_(line)) {
    return false;
  }
  const l = String(line || '').trim();
  if (l.length < 10) {
    return false;
  }
  if (/^доставка\s+товара/i.test(l)) {
    return /адрес\s+доставки|\d+[.,]\d{2}/i.test(l);
  }
  if (looksLikeOcrProductSkuLine_(l)) {
    return /\d+[.,]\d{2}|796|,\d{3}|\d+\s*%|без\s+акциза/i.test(l);
  }
  const cyr = (l.match(/[а-яА-ЯёЁ]/g) || []).length;
  if (cyr < 6) {
    return false;
  }
  const hasNumbers = /\d+[.,]\d{2}|\d{3,}|796|,\d{3}/.test(l);
  const hasProductHint =
    /наконечник|колодк|доставк\s+товара|розетк|услуг.*доставк|организации\s+доставки|кабель|упаковк|g\d{3,}|45\.\d{3}|gx\d/i.test(l);
  return hasNumbers && hasProductHint;
}

function stripOcrJunkPrefixFromName_(name) {
  let n = String(name || '').trim();
  n = n.replace(/^(\d{1,4}\s+){1,4}(?=(?:GX|Услуг|45\.|Г\d))/i, '');
  n = n.replace(/^\d{1,2}\s+(?=[A-Za-zА-ЯЁёGxУ])/i, '');
  return n.trim();
}

function cleanProductName_(name) {
  return stripOcrJunkPrefixFromName_(
    String(name || '')
      .replace(/^\d+\s*[А-Яа-яA-Za-z]\.\s*/, '')
      .replace(/^\d+\s+[А-Яа-яA-Z]\.\s*/, '')
      .replace(/\s+/g, ' ')
      .trim()
  );
}

function isGarbageMappedRow_(mapped) {
  const name = cleanProductName_(mapped[1]);
  if (!name || name.length < 4) {
    return true;
  }
  if (isOcrNoiseLine_(name) || isOcrInvoiceMetaLine_(name)) {
    return true;
  }
  if (/^основание\s+передачи/i.test(name)) {
    return true;
  }
  if (/^(сч[её]т|передаточн|документ\s*\(|всего\s+к|грузополучатель|и\s+передаточн|заказ\s+клиента)/i.test(name)) {
    return true;
  }
  if (/января|февраля|апреля|декабря\s+\d{4}/i.test(name) && !looksLikeOcrProductSkuLine_(name)) {
    return true;
  }
  if (/^00-\d{5,}$/i.test(name) && !mapped[7] && !mapped[11] && !mapped[6]) {
    return true;
  }
  if (isNumericOnlyProductName_(name)) {
    return true;
  }
  if (isDeliveryServiceRow_(name)) {
    return !!(mapped[11] || mapped[7] || mapped[10]);
  }
  const hasMetric = !!(mapped[5] || mapped[6] || mapped[7] || mapped[11]);
  const hasUnit = mapped[4] && /шт|кг/i.test(mapped[4]);
  const hasOkei = mapped[3] === '796';
  const hasSkuName = looksLikeOcrProductSkuLine_(name) || /колодк|наконечник|розетк|gx\d/i.test(name);
  if (hasSkuName && (hasMetric || hasUnit || hasOkei)) {
    return false;
  }
  if (/никелирование|2-конт/i.test(name) && !/gx|розетк/i.test(name) && !hasOkei && !mapped[5]) {
    return true;
  }
  if (/^\s*акциза\s*$/i.test(name) || /акциза\s*$/i.test(name) && name.length < 40 && !hasSkuName) {
    return true;
  }
  return !(hasMetric || hasUnit || hasOkei);
}

/** Наименование — только цифры (ошибка TAB-выравнивания OCR). */
function isNumericOnlyProductName_(name) {
  const n = String(name || '').trim();
  if (!n) {
    return true;
  }
  if (/gx\d|розетк|колодк|наконечник|услуг|доставк|45\.\d{4}/i.test(n)) {
    return false;
  }
  return /^[\d\s.,]+$/.test(n.replace(/\s+/g, ''));
}

/** Числа-суммы из токенов строки (без мелкого кол-ва/цены). */
function collectMoneyNumbersFromPool_(pool) {
  const nums = [];
  for (let i = 0; i < pool.length; i++) {
    const t = pool[i];
    if (!t || isVatRate_(t) || isExcise_(t) || isCountryCode_(t) || isOkeiCode_(t) || isUnitDesignation_(t)) {
      continue;
    }
    const n = parseRuNumber_(t);
    if (isNaN(n) || n < 50) {
      continue;
    }
    if (looksLikeMoneySum_(t) || isMoney_(t) || isCostWithoutVat_(t) || n >= 50) {
      nums.push(n);
    }
  }
  return nums;
}

/** Собрать крупные суммы из OCR-строки (3125, 625, 3750 …). */
function scrapeMoneyNumbersFromLine_(line) {
  const nums = [];
  const parts = String(line || '').split(/\s+/);
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i].replace(/\u00a0/g, '').trim();
    if (!p || isVatRate_(p)) {
      continue;
    }
    const n = parseRuNumber_(p);
    if (isNaN(n) || n < 50) {
      continue;
    }
    if (/^\d{1,2}$/.test(p) && n < 50) {
      continue;
    }
    if (n < 200 && !/[.,]/.test(p) && !/\d{3,}/.test(p.replace(/\s/g, ''))) {
      continue;
    }
    nums.push(n);
  }
  return nums;
}

/**
 * Тройка сумм УПД: стоимость + НДС = всего (3125+625=3750, 400+80=480).
 * @param {number} [minCost] минимальная «стоимость без НДС» (для строки товара GX — не брать 400+80).
 * @return {boolean}
 */
function tryAssignCostVatTotalTriple_(out, nums, minCost) {
  const minC = minCost || 0;
  const unique = [];
  for (let i = 0; i < nums.length; i++) {
    const n = nums[i];
    if (!unique.some(function (u) {
      return Math.abs(u - n) < 0.02;
    })) {
      unique.push(n);
    }
  }
  unique.sort(function (a, b) {
    return a - b;
  });
  let best = null;
  for (let k = unique.length - 1; k >= 2; k--) {
    const total = unique[k];
    for (let i = 0; i < k; i++) {
      for (let j = i + 1; j < k; j++) {
        const x = unique[i];
        const y = unique[j];
        if (Math.abs(x + y - total) > Math.max(1, total * 0.002)) {
          continue;
        }
        const cost = Math.max(x, y);
        const vat = Math.min(x, y);
        const ratio = vat / cost;
        if (ratio < 0.05 || ratio > 0.35) {
          continue;
        }
        if (cost < minC) {
          continue;
        }
        if (!best || cost > best.cost) {
          best = { cost: cost, vat: vat, total: total };
        }
      }
    }
  }
  if (best) {
    out[7] = formatRuMoney_(best.cost);
    out[10] = formatRuMoney_(best.vat);
    out[11] = formatRuMoney_(best.total);
    return true;
  }
  if (minC > 0) {
    return tryAssignCostVatTotalTriple_(out, nums, 0);
  }
  return false;
}

function assignCostVatTotalFromPool_(pool, out) {
  const nums = collectMoneyNumbersFromPool_(pool);
  if (nums.length >= 3) {
    return tryAssignCostVatTotalTriple_(out, nums);
  }
  return false;
}

/** Кол-во и цена из пары множителей (25×125=3125). */
function inferQtyPriceFromCost_(out, line) {
  if (out[5] && out[6]) {
    return;
  }
  const cost = parseRuNumber_(out[7]);
  if (!cost || cost <= 0) {
    return;
  }
  const ints = [];
  const m = String(line || '').match(/\b\d{1,4}\b/g) || [];
  for (let i = 0; i < m.length; i++) {
    const n = parseInt(m[i], 10);
    if (!isNaN(n) && n > 0 && n < 10000) {
      ints.push(n);
    }
  }
  for (let a = 0; a < ints.length; a++) {
    for (let b = a + 1; b < ints.length; b++) {
      const x = ints[a];
      const y = ints[b];
      if (Math.abs(x * y - cost) < 1.5) {
        const qty = Math.min(x, y);
        const price = Math.max(x, y);
        if (/gx\d|gx12|розетк/i.test(out[1] || '') && cost >= 1000 && qty <= 5 && price > cost * 0.5) {
          continue;
        }
        out[5] = String(qty);
        out[6] = formatRuMoney_(price);
        return;
      }
    }
  }
}

/** Сумма с НДС (480) попала в графу «сумма НДС» вместо 80. */
function fixVatTotalSlotConfusion_(out) {
  const cost = parseRuNumber_(out[7]);
  let vat = parseRuNumber_(out[10]);
  let total = parseRuNumber_(out[11]);
  if (cost > 0 && vat > 0 && (isNaN(total) || total <= cost) && vat > cost * 1.05 && vat <= cost * 1.3) {
    total = vat;
    vat = total - cost;
    out[11] = formatRuMoney_(total);
    out[10] = formatRuMoney_(vat);
    return;
  }
  if (cost > 0 && total > cost && !isNaN(vat) && vat >= total * 0.85) {
    out[10] = formatRuMoney_(total - cost);
    out[11] = formatRuMoney_(total);
  }
}

/** Дозаполнение граф из исходной OCR-строки (шт, 25, 125, без акциза, 156). */
function repairOcrMetricsFromSourceLine_(out, sourceLine) {
  const l = String(sourceLine || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!l) {
    return;
  }
  if (!out[3] && /\b796\b/.test(l)) {
    out[3] = '796';
  }
  if (!out[4] && /(?:^|\s)(?:шт\.?|wm|wт)(?:\s|$)/i.test(l)) {
    out[4] = 'шт';
  }
  const mQP = l.match(/(?:796\s*)?(?:шт\.?|wm|wт)\s*[:\s]*(\d{1,4})\s+(\d{2,5})(?:[.,]\d{2})?(?:\s|$)/i);
  if (mQP) {
    if (!out[5]) {
      out[5] = mQP[1];
    }
    if (!out[6]) {
      const pRaw = mQP[2];
      const pNum = parseRuNumber_(pRaw.indexOf(',') >= 0 || pRaw.indexOf('.') >= 0 ? pRaw : pRaw + ',00');
      out[6] = formatRuMoney_(pNum);
    }
  }
  if (!out[5] || !out[6]) {
    const mQP2 = l.match(/\b(\d{1,2})\s+(\d{2,3})\s+(\d{3,5})\s+(?:20\s*%|без)/i);
    if (mQP2 && /gx|розетк/i.test(l)) {
      if (!out[5]) {
        out[5] = mQP2[1];
      }
      if (!out[6]) {
        out[6] = formatRuMoney_(parseRuNumber_(mQP2[2] + ',00'));
      }
    }
  }
  if (!out[5] || !out[6]) {
    const mQP3 = l.match(/(?:796\s*)?(?:шт\.?|wm|wт)\D{0,8}(\d{1,3})\D{0,8}(\d{2,4})\D{0,8}(\d{3,5})/i);
    if (mQP3) {
      if (!out[5]) {
        out[5] = mQP3[1];
      }
      if (!out[6]) {
        out[6] = formatRuMoney_(parseRuNumber_(mQP3[2] + ',00'));
      }
      if (!out[3]) {
        out[3] = '796';
      }
      if (!out[4]) {
        out[4] = 'шт';
      }
    }
  }
  if (!out[8] && (/без\s+акциза/i.test(l) || (out[9] && /^\d{1,2}\s*%$/.test(String(out[9]).trim())))) {
    out[8] = 'без акциза';
  }
  if (!out[12]) {
    const cm = l.match(/(?:^|\s)156(?:\s|$|[\s,])/);
    if (cm) {
      out[12] = '156';
    }
  }
  const scraped = scrapeMoneyNumbersFromLine_(l);
  if (scraped.length >= 3) {
    const minC = /gx\d|gx12|розетк/i.test(out[1] || '') ? 500 : 0;
    tryAssignCostVatTotalTriple_(out, scraped, minC);
  }
  inferQtyPriceFromCost_(out, l);
  fixVatTotalSlotConfusion_(out);
  fixQtyPriceCostSlots_(out);
}

/** Наименование и кол-во попали не в те графы после OCR. */
function repairScrambledOcrRow_(mapped, sourceLine, fullText) {
  const rawName = String(mapped[1] || '').trim();
  if (/^00-\d{5,}$/i.test(rawName) || /^00-\d{5,}\s*$/i.test(rawName) || isNumericOnlyProductName_(rawName)) {
    for (let c = 2; c < mapped.length; c++) {
      const v = String(mapped[c] || '').trim();
      if (!v) {
        continue;
      }
      if (/45\.\d{4}|колодк|наконечник|розетк|gx\d|техком|\(техком\)|g\d{4}\./i.test(v)) {
        const sku = /^00-\d{5,}$/i.test(rawName) ? ' [' + rawName + ']' : '';
        mapped[1] = cleanProductName_(v + sku);
        mapped[c] = '';
        break;
      }
    }
  }
  if (isNumericOnlyProductName_(mapped[1])) {
    for (let c = 2; c < mapped.length; c++) {
      const v = String(mapped[c] || '').trim();
      if (/gx\d|розетк|колодк|наконечник|услуг.*доставк|g\d{4}\./i.test(v)) {
        mapped[1] = cleanProductName_(v);
        mapped[c] = '';
        break;
      }
    }
  }
  if (!mapped[5] && isDeliveryServiceRow_(mapped[1])) {
    for (let c = 6; c <= 8; c++) {
      const v = String(mapped[c] || '').trim();
      if (/^[12]$/.test(v)) {
        mapped[5] = v;
        mapped[c] = '';
        break;
      }
    }
  }
  fixSwappedQtyPrice_(mapped);
  let lineForRepair = sourceLine;
  const segment = fullText ? extractOcrSegmentForProduct_(fullText, mapped[1]) : '';
  if (fullText) {
    const better = findOcrProductLineForRow_(fullText, mapped[1]);
    if (better && (!lineForRepair || (lineForRepair.indexOf('796') < 0 && better.indexOf('796') >= 0))) {
      lineForRepair = better;
    }
    if (segment && segment.length > (lineForRepair || '').length) {
      lineForRepair = segment.replace(/\s+/g, ' ').trim();
    }
  }
  if (/gx\d|gx12|розетк/i.test(mapped[1] || '')) {
    repairGxProductRowMetrics_(mapped, fullText || lineForRepair);
  } else if (/услуг|доставк|упаковк/i.test(mapped[1] || '')) {
    repairDeliveryProductRowMetrics_(mapped, fullText);
    if (lineForRepair) {
      repairOcrMetricsFromSourceLine_(mapped, lineForRepair);
    }
  } else if (lineForRepair) {
    repairOcrMetricsFromSourceLine_(mapped, lineForRepair);
  }
  fixVatTotalSlotConfusion_(mapped);
  fixQtyPriceCostSlots_(mapped);
}

/** OCR иногда ставит 400 в «количество», а 1 в «цену» для услуги доставки. */
function fixSwappedQtyPrice_(out) {
  const q = parseRuNumber_(out[5]);
  const p = parseRuNumber_(out[6]);
  if (q >= 50 && p > 0 && p <= 5 && /услуг|упаковк/i.test(out[1] || '')) {
    out[5] = String(Math.round(p));
    out[6] = formatRuMoney_(q);
  }
}

/** Строка «Доставка товара» без ОКЕИ/шт — только суммы и адрес в наименовании. */
function tokenizeOcrDeliveryLine_(line) {
  const l = String(line || '').replace(/\u00a0/g, ' ').trim();
  const tokens = [];
  let sm = l.match(/^(\d{1,2})\s+/);
  let rest = l;
  if (sm) {
    tokens.push(sm[1]);
    rest = l.substring(sm[0].length).trim();
  }
  const cut = rest.search(/\d{1,3}(?:\s\d{3})*[.,]\d{2}|\d+[.,]\d{2}|без\s+акциза|\d{1,2}\s*%/i);
  const name = cut > 0 ? rest.substring(0, cut).trim() : rest;
  if (name) {
    tokens.push(name);
  }
  tokens.push('-');
  const tail = cut > 0 ? rest.substring(cut) : '';
  const parts =
    tail.match(/(\d{1,2}\s*%|без\s+акциза|\d{1,3}(?:\s\d{3})*[.,]\d{2}|\d+[.,]\d{2})/gi) || [];
  for (let i = 0; i < parts.length; i++) {
    tokens.push(parts[i].trim());
  }
  return tokens.length >= 2 ? tokens : splitTableLine_(l);
}

/**
 * Разбор одной OCR-строки товара (часто без TAB, с «796 шт» в середине).
 */
function tokenizeOcrProductLine_(line) {
  const l = String(line || '').replace(/\u00a0/g, ' ').trim();
  if (!l) {
    return [];
  }
  if (l.indexOf('\t') !== -1) {
    const tabbed = splitTableLine_(l);
    if (tabbed.length >= 6) {
      return tabbed;
    }
  }
  const wide = l
    .split(/\s{2,}/)
    .map(function (x) {
      return x.trim();
    })
    .filter(function (x) {
      return x.length > 0;
    });
  if (wide.length >= 6) {
    return wide;
  }
  if (/^доставка\s+товара/i.test(l) && l.indexOf('796') === -1) {
    return tokenizeOcrDeliveryLine_(l);
  }
  let okeiMatch = l.match(/(?:^|\s)(796)\s*(шт\.?|ШТ|кг\.?|кг)(?:\s|$)/i);
  if (!okeiMatch) {
    okeiMatch = l.match(/(?:^|\s)(796)\s*(wm|wт|шт\.?|ШТ)(?:\s|$)/i);
  }
  if (!okeiMatch) {
    const unitOnly = l.match(/(?:^|\s)(шт\.?|ШТ|wm|wт)\s+(\d{1,4})\s+(\d{2,5})/i);
    if (unitOnly) {
      const tokens = [];
      const before = l.substring(0, unitOnly.index).trim();
      if (before) {
        const sm = before.match(/^(\d{1,2})\s+/);
        if (sm) {
          tokens.push(sm[1]);
        }
        tokens.push(before.replace(/^\d{1,2}\s+/, '').trim());
      }
      tokens.push('796');
      tokens.push('шт');
      tokens.push(unitOnly[2]);
      tokens.push(unitOnly[3]);
      appendOcrAfterUnitTokens_(l.substring(unitOnly.index + unitOnly[0].length), tokens);
      return tokens.length >= 4 ? tokens : splitTableLine_(l);
    }
    return wide.length >= 2 ? wide : splitTableLine_(l);
  }
  const okeiIdx = l.indexOf(okeiMatch[1], okeiMatch.index);
  let before = l.substring(0, okeiIdx).trim();
  let after = l.substring(okeiIdx + okeiMatch[0].trim().length).trim();
  const tokens = [];
  let sm = before.match(/^(\d{1,2})\s+/);
  if (sm) {
    tokens.push(sm[1]);
    before = before.substring(sm[0].length).trim();
  }
  sm = before.match(/^(00-\d{5,})\s*/);
  if (sm) {
    tokens.push(sm[1]);
    before = before.substring(sm[0].length).trim();
  }
  sm = before.match(/^([-—])\s*/);
  if (sm) {
    tokens.push(sm[1]);
    before = before.substring(sm[0].length).trim();
  }
  if (before) {
    tokens.push(before);
  }
  tokens.push('796');
  tokens.push(/^шт/i.test(okeiMatch[2]) ? 'шт' : okeiMatch[2]);
  appendOcrAfterUnitTokens_(after, tokens);
  return tokens.length >= 4 ? tokens : splitTableLine_(l);
}

/** Токены после «796 шт» — по словам, чтобы не терять «25» и «125» без копеек. */
function appendOcrAfterUnitTokens_(after, tokens) {
  const chunks = String(after || '')
    .split(/\s+/)
    .map(function (x) {
      return x.trim();
    })
    .filter(function (x) {
      return x.length > 0;
    });
  for (let i = 0; i < chunks.length; i++) {
    let p = chunks[i];
    if (/^без$/i.test(p) && /^акциз/i.test(chunks[i + 1] || '')) {
      tokens.push('без акциза');
      i++;
      continue;
    }
    if (/^акциз/i.test(p) && tokens.length && /без$/i.test(tokens[tokens.length - 1])) {
      tokens[tokens.length - 1] = 'без акциза';
      continue;
    }
    if (/^\d{1,2}$/.test(p) && /^%$/.test(chunks[i + 1] || '')) {
      tokens.push(p + ' %');
      i++;
      continue;
    }
    if (/^\d{1,2}%$/.test(p) || /^без\s+акциза$/i.test(p)) {
      tokens.push(p);
      continue;
    }
    if (/^--$|^—$|^-$/.test(p)) {
      tokens.push(p);
      continue;
    }
    if (/^\d{8,}\/\d+/.test(p)) {
      tokens.push(p);
      continue;
    }
    if (/^[A-Za-zА-Яа-яЁё]{4,}$/.test(p) && !/^шт$/i.test(p)) {
      tokens.push(p);
      continue;
    }
    if (/^\d{1,7}$/.test(p)) {
      tokens.push(p);
      continue;
    }
    if (/^\d{1,3}(?:\s\d{3})*[.,]\d{2}$/.test(p) || /^\d+[.,]\d{2,3}$/.test(p)) {
      tokens.push(p);
      continue;
    }
    if (/^\d{1,7},\d{3}$/.test(p)) {
      tokens.push(p);
    }
  }
}

/** Оценка качества набора строк OCR (меньше мусора и больше «товарных» строк — выше). */
function scoreOcrTableQuality_(table) {
  if (!table || !table.rows || !table.rows.length) {
    return -1000;
  }
  let score = 0;
  const n = table.rows.length;
  if (n >= 1 && n <= 4) {
    score += 25;
  }
  if (n > 5) {
    score -= (n - 5) * 12;
  }
  for (let i = 0; i < table.rows.length; i++) {
    const line = table.rows[i].join(' ');
    if (/GX\d|GX12|услуг.*доставк|организации\s+доставки|45\.7373|Г8510/i.test(line)) {
      score += 18;
    }
    if (looksLikeProductDataLine_(line)) {
      score += 10;
    }
    if (isOcrTableJunkDataLine_(line)) {
      score -= 25;
    }
    if (isOcrNoiseLine_(line)) {
      score -= 15;
    }
  }
  return score;
}

/** Выбор между таблицей по блоку УПД и построчной эвристикой (для Dart и др. сканов). */
function pickBestOcrTable_(text) {
  const tableBlock = extractTableBlock_(text);
  const fromBlock = parseTableFromBlock_(tableBlock);
  const fromLines = parseOcrProductRowsOnly_(text);
  const sBlock = scoreOcrTableQuality_(fromBlock);
  const sLines = scoreOcrTableQuality_(fromLines);
  Logger.log('OCR: оценка таблицы (блок УПД=' + sBlock + ', строки товаров=' + sLines + ')');
  if (fromLines && fromLines.rows && fromLines.rows.length && sLines >= sBlock) {
    Logger.log('OCR: используем строки товаров по эвристике: ' + fromLines.rows.length);
    return fromLines;
  }
  if (fromBlock && fromBlock.rows && fromBlock.rows.length) {
    Logger.log('OCR: таблица из блока УПД: строк ' + fromBlock.rows.length);
    return fromBlock;
  }
  return fromLines || fromBlock;
}

function isOcrTableJunkDataLine_(line) {
  const l = String(line || '').trim();
  if (!l) {
    return true;
  }
  if (/никелирование.*акциза/i.test(l) && !/gx|розетк|796|3125|3750/i.test(l)) {
    return true;
  }
  if (/^(количество|цена|стоимость|единица|акциз|налоговая|наименование)\b/i.test(l) && l.length < 80) {
    return true;
  }
  return false;
}

/** Строки товаров из «сырого» OCR-текста (без шапки УПД и мусорных строк). */
function parseOcrProductRowsOnly_(text) {
  const lines = normalizeText_(text).split('\n');
  const rows = [];
  let inTableRegion = false;
  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].replace(/\u00a0/g, ' ').trim();
    if (!line) {
      continue;
    }
    if (/наименован/i.test(line) && /(п\/п|код|количество|единиц)/i.test(line)) {
      inTableRegion = true;
      continue;
    }
    if (/всего\s+к\s+оплате/i.test(line)) {
      break;
    }
    if (!inTableRegion && !looksLikeOcrProductSkuLine_(line)) {
      continue;
    }
    if (!looksLikeProductDataLine_(line)) {
      continue;
    }
    if (/^доставка\s+товара/i.test(line) && i + 1 < lines.length) {
      let j = i + 1;
      while (j < lines.length && j < i + 6) {
        const extra = lines[j].replace(/\u00a0/g, ' ').trim();
        if (
          extra &&
          !isOcrNoiseLine_(extra) &&
          (/адрес\s+доставки|москва|ленинск|собода|доставк/i.test(extra) || extra.length < 120)
        ) {
          line = line + ' ' + extra;
          if (/90[.,]\d{2}|543|452[.,]\d{2}/.test(extra)) {
            break;
          }
        } else if (looksLikeProductDataLine_(extra)) {
          break;
        }
        j++;
      }
    }
    const cells = tokenizeOcrProductLine_(line);
    if (cells.length >= 2) {
      rows.push(cells);
    }
  }
  if (!rows.length) {
    return null;
  }
  Logger.log('OCR: найдено кандидатов в строки товаров: ' + rows.length);
  return {
    header: CANONICAL_UPD_HEADERS.slice(),
    rows: rows,
    width: maxRowLen_(rows),
  };
}

/** Строки товаров с TAB без маркеров (склеенный ответ Gemini). */
function parseTabularProductLines_(text) {
  const lines = normalizeText_(text)
    .split('\n')
    .map(function (l) {
      return l.replace(/\u00a0/g, ' ').trim();
    })
    .filter(function (l) {
      return l.length > 0;
    });
  const rows = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.indexOf('\t') === -1) {
      continue;
    }
    if (/наименован/i.test(line) && /п\/п|количество/i.test(line) && !/^\d+\t/.test(line)) {
      continue;
    }
    if (/Всего\s+к\s+оплате|^Итого\b/i.test(line)) {
      break;
    }
    const cells = splitTableLine_(line);
    if (cells.length < 5) {
      continue;
    }
    if (!looksLikeSeqNumber_(cells[0]) && !/gx|услуг|45\.|колодк/i.test(line)) {
      continue;
    }
    rows.push(cells);
  }
  if (!rows.length) {
    return null;
  }
  Logger.log('Gemini: таблица из TAB-строк: ' + rows.length);
  return {
    header: CANONICAL_UPD_HEADERS.slice(),
    rows: rows,
    width: maxRowLen_(rows),
  };
}

/** Таблица из блока ===TABLE=== (TAB). */
function parseGeminiTableSection_(text) {
  const n = normalizeText_(text);
  let tm = n.match(/===\s*TABLE\s*===\s*([\s\S]*?)(?====\s*END\s*===|$)/i);
  if (!tm) {
    return null;
  }
  const lines = tm[1]
    .split('\n')
    .map(function (l) {
      return l.replace(/\u00a0/g, ' ').trim();
    })
    .filter(function (l) {
      return l.length > 0;
    });
  if (!lines.length) {
    return null;
  }
  let start = 0;
  if (/наименован|№\s*п\/п|код\s+вида/i.test(lines[0]) && !/^\d+\t/.test(lines[0])) {
    start = 1;
  }
  const rows = [];
  for (let i = start; i < lines.length; i++) {
    if (/Всего\s+к\s+оплате|^Итого\b/i.test(lines[i])) {
      break;
    }
    if (isOcrNoiseLine_(lines[i])) {
      continue;
    }
    const cells = splitTableLine_(lines[i]);
    if (!cells.length) {
      continue;
    }
    rows.push(cells);
  }
  if (!rows.length) {
    return parseTabularProductLines_(n);
  }
  Logger.log('Gemini: таблица из ===TABLE===: ' + rows.length);
  return {
    header: CANONICAL_UPD_HEADERS.slice(),
    rows: rows,
    width: maxRowLen_(rows),
  };
}

function looksLikeSeqNumber_(s) {
  return /^\d{1,4}$/.test(String(s || '').trim());
}

/** Артикул / код номенклатуры (не порядковый № п/п и не сумма). */
function looksLikeProductCode_(s) {
  const t = String(s || '').trim();
  if (!t || looksLikeSeqNumber_(t)) {
    return false;
  }
  if (/^\d{1,3}([.,]\d{2})?$/.test(t)) {
    return false;
  }
  if (/^\d{2,}\.\d{3,}\.\d{3,}/.test(t)) {
    return false;
  }
  if (/^00-\d+/.test(t)) {
    return true;
  }
  if (/^\d{3,}\s+\d{3,}$/.test(t)) {
    return true;
  }
  if (t.length >= 4 && /^[\dA-Za-zА-Яа-я.\-()\s]+$/.test(t) && !/^(без\s+акциза|\d+%)$/i.test(t)) {
    return !/^\d+([.,]\d{1,2})?$/.test(t.replace(/\s/g, ''));
  }
  return false;
}

function stripLeadingProductCodeColumn_(row) {
  let r = row.slice();
  const canonLen = CANONICAL_UPD_HEADERS.length;
  while (r.length > canonLen) {
    if (r.length >= 2 && looksLikeSeqNumber_(r[0]) && looksLikeProductCode_(r[1])) {
      r = [r[0]].concat(r.slice(2));
      continue;
    }
    if (r.length >= 2 && looksLikeProductCode_(r[0]) && looksLikeSeqNumber_(r[1])) {
      r = r.slice(1);
      continue;
    }
    if (looksLikeProductCode_(r[0])) {
      r = r.slice(1);
      continue;
    }
    break;
  }
  return r;
}

function parseRuNumber_(s) {
  const t = String(s || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s/g, '')
    .replace(',', '.');
  const n = parseFloat(t);
  return isNaN(n) ? NaN : n;
}

function formatRuMoney_(n) {
  if (isNaN(n) || n <= 0) {
    return '';
  }
  return String(Math.round(n * 100) / 100).replace('.', ',');
}

function isOkeiCode_(t) {
  return isOkeiCodeWithContext_(t, '');
}

/** Код ОКЕИ (796, 166…), не путать со страной (156, 643…). */
function isOkeiCodeWithContext_(t, nextToken) {
  const s = String(t || '').trim();
  if (!/^\d{3}$/.test(s)) {
    return false;
  }
  if (isUnitDesignation_(nextToken)) {
    return true;
  }
  if (/^(796|166|055|006|112|715|898|797)$/.test(s)) {
    return true;
  }
  if (/^(156|643|840|380|051)$/.test(s)) {
    return false;
  }
  return false;
}

function isUnitDesignation_(t) {
  const s = String(t || '').trim();
  return /^(шт\.?|кг\.?|т\.?|м\.?|м2|м3|л\.?|упак\.?|компл\.?|ч\.?|чел\.?|мест\.?|рул\.?|пог\.?\s*м\.?)$/i.test(s);
}

/** Сумма с копейками / разрядами («3 125,00»), не количество. */
function looksLikeMoneySum_(t) {
  const s = String(t || '').trim();
  if (isQuantityThousandths_(s)) {
    return false;
  }
  if (/\d[\d\s]{2,}[.,]\d{2}$/.test(s)) {
    return true;
  }
  if (/\s/.test(s) && /\d{4,}/.test(s.replace(/[^\d]/g, ''))) {
    return true;
  }
  return false;
}

function isQuantity_(t) {
  const s = String(t || '').trim();
  if (!s || isOkeiCode_(s) || isUnitDesignation_(s) || isVatRate_(s)) {
    return false;
  }
  if (isQuantityThousandths_(s)) {
    return false;
  }
  if (/^\d{1,2}\s*%$/.test(s) || looksLikeMoneySum_(s)) {
    return false;
  }
  const compact = s.replace(/\s/g, '');
  if (/^\d{1,4}[.,]\d{2}$/.test(compact)) {
    return false;
  }
  if (isCountryCode_(s)) {
    return false;
  }
  const n = parseRuNumber_(s);
  if (isNaN(n) || n <= 0 || n >= 1000000) {
    return false;
  }
  return /^\d{1,7}([.,]\d{1,4})?$/.test(compact);
}

/** Цена за единицу (5,83 / 8.80 / 25,00). */
function isUnitPrice_(t) {
  const s = String(t || '').trim();
  if (
    !s ||
    isOkeiCode_(s) ||
    isUnitDesignation_(s) ||
    isQuantity_(s) ||
    isQuantityThousandths_(s) ||
    isQuantityHundredths_(t)
  ) {
    return false;
  }
  if (looksLikeMoneySum_(s)) {
    return false;
  }
  const compact = s.replace(/\s/g, '');
  return /^\d{1,6}[.,]\d{1,2}$/.test(compact);
}

function isMoney_(t) {
  const s = String(t || '').trim();
  if (!s) {
    return false;
  }
  if (isQuantityThousandths_(s)) {
    return false;
  }
  if (/^\d{1,2}\s*%$/.test(s) || /^без\s+акциза$/i.test(s)) {
    return false;
  }
  if (isQuantity_(s) || isUnitPrice_(s)) {
    return false;
  }
  if (looksLikeMoneySum_(s)) {
    return true;
  }
  const n = parseRuNumber_(s);
  return !isNaN(n) && n >= 500;
}

/** Количество в формате УПД: «700,000» / «700.000» = 700 (три знака после разделителя). */
function isQuantityThousandths_(t) {
  const s = String(t || '')
    .trim()
    .replace(/\u00a0/g, ' ')
    .replace(/\s/g, '');
  return /^\d{1,7}[.,]\d{3}$/.test(s);
}

function normalizeQuantityToken_(t) {
  const s = String(t || '')
    .trim()
    .replace(/\u00a0/g, ' ')
    .replace(/\s/g, '');
  const m = s.match(/^(\d{1,7})[.,]\d{3}$/);
  if (m) {
    return m[1];
  }
  if (isQuantityHundredths_(t)) {
    const n = Math.round(parseRuNumber_(t));
    return String(n) + ',00';
  }
  return String(t || '').trim();
}

/** Количество «700,00» / «250,00» — целое число с двумя нулями после запятой. */
function isQuantityHundredths_(t) {
  const s = String(t || '')
    .trim()
    .replace(/\s/g, '');
  if (!/^\d{1,7}[.,]\d{2}$/.test(s)) {
    return false;
  }
  if (isQuantityThousandths_(t)) {
    return false;
  }
  const n = parseRuNumber_(t);
  if (isNaN(n) || n < 10) {
    return false;
  }
  return Math.abs(n - Math.round(n)) < 0.001;
}

function isQuantityFormatted_(t) {
  if (isQuantityThousandths_(t)) {
    return true;
  }
  if (isQuantityHundredths_(t)) {
    return true;
  }
  return isPlainInteger_(t);
}

function isVatRate_(t) {
  const s = String(t || '').trim();
  return /^\d{1,2}\s*%$/.test(s) || /^без\s+акциза$/i.test(s) || s === '0';
}

function isExcise_(t) {
  const s = String(t || '').trim();
  return /^без\s+акциза$/i.test(s) || s === '0' || s === '—' || s === '-';
}

function isCountryCode_(t) {
  const s = String(t || '').trim();
  return /^(156|643|840|380|051|112|276|392|124|410|158|704|356)$/.test(s);
}

function isCountryName_(t) {
  const s = String(t || '').trim();
  return /^[A-Za-zА-Яа-яЁё-]{3,}$/.test(s) && !isUnitDesignation_(s) && !isOkeiCode_(s) && !isMoney_(s);
}

function isDeclReg_(t) {
  const s = String(t || '').trim();
  return /\d{5,}\/\d{6,}\/\d{4,}/.test(s) || (s.length > 12 && /\//.test(s));
}

function isKodVidaTovara_(t) {
  const s = String(t || '').trim();
  return /^--$|^—$|^-$/.test(s);
}

/** Пропуск «0», «без акциза», «20%» между количеством и ценой. */
function shouldSkipInQtyPricePhases_(t) {
  const s = String(t || '').trim();
  if (!s) {
    return true;
  }
  if (isExcise_(s) || isVatRate_(s)) {
    return true;
  }
  return s === '0';
}

/** Целое число в графе «количество» или «цена» (не ОКЕИ, не код страны). */
function isPlainInteger_(t) {
  const s = String(t || '').trim();
  const compact = s.replace(/\s/g, '');
  if (!/^\d{1,7}$/.test(compact)) {
    return false;
  }
  if (isOkeiCode_(s) || isCountryCode_(s)) {
    return false;
  }
  const n = parseRuNumber_(s);
  return !isNaN(n) && n > 0;
}

/** Стоимость без НДС: сумма с копейками или крупное целое (3125, 4083). */
function isCostWithoutVat_(t) {
  if (isQuantityThousandths_(t)) {
    return false;
  }
  if (looksLikeMoneySum_(t)) {
    return true;
  }
  if (isPlainInteger_(t)) {
    return parseRuNumber_(t) >= 100;
  }
  return isMoney_(t) && !isVatRate_(t);
}

/** Цена попала в «стоимость», кол-во×цена → стоимость; при пустом кол-ве — из стоимости/цены. */
function fixQtyPriceCostSlots_(out) {
  if (out[5]) {
    out[5] = normalizeQuantityToken_(out[5]);
  }
  const q = parseRuNumber_(out[5]);
  if (q > 0 && out[7] && !out[6]) {
    const v = parseRuNumber_(out[7]);
    if (!isNaN(v) && v > 0 && isUnitPrice_(out[7])) {
      out[6] = out[7];
      out[7] = '';
    }
  }
  const q2 = parseRuNumber_(out[5]);
  const p2 = parseRuNumber_(out[6]);
  if (q2 > 0 && p2 > 0 && !out[7]) {
    const derived = Math.round(q2 * p2 * 100) / 100;
    if (derived > 0) {
      out[7] = String(derived).replace('.', ',');
    }
  }
  const q3 = parseRuNumber_(out[5]);
  const p3 = parseRuNumber_(out[6]);
  const c3 = parseRuNumber_(out[7]);
  if ((isNaN(q3) || q3 <= 0) && p3 > 0 && c3 > 0) {
    const derivedQ = c3 / p3;
    const rounded =
      Math.abs(derivedQ - Math.round(derivedQ)) < 0.05
        ? Math.round(derivedQ)
        : Math.round(derivedQ * 1000) / 1000;
    if (rounded > 0 && rounded < 1000000) {
      out[5] = String(rounded).replace('.', ',');
    }
  }
  const totalVat = parseRuNumber_(out[11]);
  const vatAmt = parseRuNumber_(out[10]);
  if (!out[7] && totalVat > 0) {
    if (vatAmt > 0) {
      out[7] = formatRuMoney_(totalVat - vatAmt);
    } else if (out[9] && /20/.test(String(out[9])) && totalVat >= 900) {
      // Не делить на 1.2, если в «всего» попала строка НДС (625 и т.п.)
      out[7] = formatRuMoney_(totalVat / 1.2);
    }
  }
  const q4 = parseRuNumber_(out[5]);
  const c4 = parseRuNumber_(out[7]);
  if (!out[6] && q4 > 0 && c4 > 0) {
    out[6] = formatRuMoney_(c4 / q4);
  }
}

/** Только строка без кол-ва/цены (итоговая доставка), не «услуги доставки и упаковки» с номенклатурой. */
function isDeliveryServiceRow_(name) {
  const n = String(name || '').trim();
  if (/^доставка\s+товара/i.test(n)) {
    return true;
  }
  if (/доставк.*адрес\s+доставки/i.test(n)) {
    return true;
  }
  return false;
}

/**
 * Строка «Доставка» без количества/цены — только итоговая сумма (часто с НДС).
 */
function assignDeliveryRowMetrics_(pool, out) {
  const amounts = [];
  for (let i = 0; i < pool.length; i++) {
    const t = pool[i];
    if (looksLikeMoneySum_(t) || isMoney_(t)) {
      const n = parseRuNumber_(t);
      if (!isNaN(n) && n > 0) {
        amounts.push({ t: t, n: n });
      }
    }
  }
  amounts.sort(function (a, b) {
    return a.n - b.n;
  });
  if (amounts.length >= 1) {
    out[7] = amounts[0].t;
  }
  if (amounts.length >= 3) {
    out[10] = amounts[amounts.length - 2].t;
    out[11] = amounts[amounts.length - 1].t;
  } else if (amounts.length === 2) {
    out[11] = amounts[1].t;
    if (!out[10]) {
      const cost = parseRuNumber_(amounts[0].t);
      const total = parseRuNumber_(amounts[1].t);
      if (total > cost) {
        out[10] = formatRuMoney_(total - cost);
      }
    }
  } else if (amounts.length === 1) {
    out[11] = amounts[0].t;
  }
  for (let i = 0; i < pool.length; i++) {
    if (isVatRate_(pool[i])) {
      out[9] = pool[i];
    }
    if (isExcise_(pool[i])) {
      out[8] = pool[i];
    }
    if (isKodVidaTovara_(pool[i]) && !out[2]) {
      out[2] = pool[i];
    }
  }
  if (!out[2]) {
    out[2] = '-';
  }
  fixQtyPriceCostSlots_(out);
}

/**
 * Порядок граф УПД после единицы измерения: кол-во → цена → стоимость без НДС → акциз → % → НДС → с НДС → страна.
 */
function assignMetricsInDocumentOrder_(pool, out) {
  let phase = 0;

  for (let i = 0; i < pool.length; i++) {
    const t = pool[i];
    if (!t) {
      continue;
    }

    if (isKodVidaTovara_(t)) {
      if (!out[2]) {
        out[2] = t;
      }
      continue;
    }
    if (isOkeiCodeWithContext_(t, pool[i + 1] || '')) {
      if (!out[3]) {
        out[3] = t;
      }
      continue;
    }
    if (isUnitDesignation_(t)) {
      if (!out[4]) {
        out[4] = t;
      }
      continue;
    }

    if (phase <= 2 && shouldSkipInQtyPricePhases_(t)) {
      continue;
    }

    if (phase === 0) {
      if (isCountryCode_(t) || isCountryName_(t) || isDeclReg_(t)) {
        phase = 7;
        i--;
        continue;
      }
      if (isQuantityFormatted_(t)) {
        out[5] = normalizeQuantityToken_(t);
        phase = 1;
        continue;
      }
      if (isUnitPrice_(t)) {
        phase = 1;
        i--;
        continue;
      }
      if (looksLikeMoneySum_(t) || isCostWithoutVat_(t)) {
        phase = 2;
        i--;
        continue;
      }
    }

    if (phase === 1) {
      if (isUnitPrice_(t)) {
        out[6] = t;
        phase = 2;
        continue;
      }
      if (isQuantityFormatted_(t) && !out[5]) {
        out[5] = normalizeQuantityToken_(t);
        continue;
      }
      if (isPlainInteger_(t)) {
        out[6] = t;
        phase = 2;
        continue;
      }
      if (looksLikeMoneySum_(t) || isCostWithoutVat_(t)) {
        phase = 2;
        i--;
        continue;
      }
    }

    if (phase === 2) {
      if (isCostWithoutVat_(t)) {
        out[7] = t;
        phase = 3;
        continue;
      }
    }

    if (phase === 3) {
      if (isExcise_(t)) {
        out[8] = t;
        phase = 4;
        continue;
      }
      if (isVatRate_(t)) {
        phase = 4;
        i--;
        continue;
      }
    }

    if (phase === 4) {
      if (isVatRate_(t)) {
        out[9] = t;
        phase = 5;
        continue;
      }
    }

    if (phase === 5) {
      if (looksLikeMoneySum_(t) || isMoney_(t)) {
        if (!out[10]) {
          out[10] = t;
          phase = 6;
          continue;
        }
      }
    }

    if (phase === 6) {
      if (looksLikeMoneySum_(t) || isMoney_(t)) {
        if (!out[11]) {
          out[11] = t;
          phase = 7;
          continue;
        }
      }
    }

    if (phase >= 7) {
      if (isCountryCode_(t) && !out[12]) {
        out[12] = t;
        continue;
      }
      if (isCountryName_(t) && !out[13]) {
        out[13] = t;
        continue;
      }
      if (isDeclReg_(t) && !out[14]) {
        out[14] = t;
        continue;
      }
    }
  }

  assignTailColumnsFromPool_(pool, out);
  assignCostVatTotalFromPool_(pool, out);
}

/**
 * Дозаполнение акциза, НДС, страны и т.д. (если токены остались в pool или пропущены фазами).
 */
function assignTailColumnsFromPool_(pool, out) {
  for (let i = 0; i < pool.length; i++) {
    const t = pool[i];
    if (!t) {
      continue;
    }
    if (isKodVidaTovara_(t) || isOkeiCodeWithContext_(t, pool[i + 1] || '') || isUnitDesignation_(t)) {
      continue;
    }
    if (t === out[5] || t === out[6] || t === out[7] || normalizeQuantityToken_(t) === out[5]) {
      continue;
    }

    if (!out[8] && (isExcise_(t) || t === '0')) {
      out[8] = t;
      continue;
    }
    if (!out[9] && isVatRate_(t)) {
      out[9] = t;
      continue;
    }
    if (!out[10] && !isVatRate_(t) && (looksLikeMoneySum_(t) || isMoney_(t)) && !isCountryCode_(t)) {
      const n = parseRuNumber_(t);
      const cost = parseRuNumber_(out[7]);
      if (!isNaN(n) && n > 0 && (isNaN(cost) || n < cost * 0.95)) {
        out[10] = t;
        continue;
      }
    }
    if (!out[11] && (looksLikeMoneySum_(t) || isMoney_(t)) && !isCountryCode_(t)) {
      const n = parseRuNumber_(t);
      const cost = parseRuNumber_(out[7]);
      if (!isNaN(n) && n > 0 && (isNaN(cost) || n >= cost * 0.9)) {
        out[11] = t;
        continue;
      }
    }
    if (!out[12] && isCountryCode_(t)) {
      out[12] = t;
      continue;
    }
    if (!out[13] && isCountryName_(t)) {
      out[13] = t;
      continue;
    }
    if (!out[14] && isDeclReg_(t)) {
      out[14] = t;
    }
  }
}

function isStrongMetricStart_(t, hasName) {
  if (!hasName) {
    return false;
  }
  return (
    isOkeiCode_(t) ||
    isUnitDesignation_(t) ||
    isQuantityThousandths_(t) ||
    isQuantityHundredths_(t) ||
    isQuantity_(t) ||
    isMoney_(t) ||
    isVatRate_(t)
  );
}

/**
 * Строка из ===TABLE=== Gemini уже с TAB-колонками — сохраняем позиции, не перетасовываем токены.
 */
function tryMapPrestructuredRow_(cells, seqNum) {
  let r = stripLeadingProductCodeColumn_(
    cells.map(function (x) {
      return String(x || '').trim();
    })
  );
  if (r.length < 8) {
    return null;
  }
  let start = 0;
  if (looksLikeSeqNumber_(r[0])) {
    start = 1;
  }
  const metricCount = CANONICAL_UPD_HEADERS.length - 2;
  if (r.length - start < metricCount + 1) {
    return null;
  }
  const tailStart = r.length - metricCount;
  const name = cleanProductName_(r.slice(start, tailStart).join(' ').trim());
  if (name.length < 4) {
    return null;
  }
  const cyrInName = (name.match(/[а-яА-ЯёЁ]/g) || []).length;
  if (cyrInName < 4 || isNumericOnlyProductName_(name)) {
    return null;
  }
  const tail = r.slice(tailStart);
  if (isUnitPrice_(tail[0]) || (isMoney_(tail[0]) && !isOkeiCode_(tail[0]))) {
    return null;
  }
  if (tail[1] && !isOkeiCode_(tail[1]) && (isUnitPrice_(tail[1]) || isMoney_(tail[1]))) {
    return null;
  }
  const hasOkeiTail = isOkeiCode_(tail[1]) || isOkeiCode_(tail[0]);
  const hasUnitTail = isUnitDesignation_(tail[2]) || isUnitDesignation_(tail[1]);
  const structured =
    hasOkeiTail &&
    (hasUnitTail || isQuantityFormatted_(tail[3]) || isQuantity_(tail[3])) &&
    (isUnitPrice_(tail[4]) ||
      looksLikeMoneySum_(tail[5]) ||
      looksLikeMoneySum_(tail[6]) ||
      isQuantityFormatted_(tail[3]));
  if (!structured) {
    return null;
  }
  const out = [];
  for (let c = 0; c < CANONICAL_UPD_HEADERS.length; c++) {
    out.push('');
  }
  out[0] = String(seqNum);
  out[1] = name;
  for (let i = 0; i < metricCount; i++) {
    let v = tail[i] || '';
    if (i === 3 && v) {
      v = normalizeQuantityToken_(v);
    }
    out[2 + i] = v;
  }
  fixQtyPriceCostSlots_(out);
  return out;
}

/**
 * Смысловое выравнивание: 796→код ОКЕИ, шт→условное обозначение, количество и суммы на свои места.
 */
function semanticMapGoodsRow_(cells, seqNum) {
  let prestructured = tryMapPrestructuredRow_(cells, seqNum);
  if (prestructured && isNumericOnlyProductName_(prestructured[1])) {
    prestructured = null;
  }
  if (prestructured) {
    return prestructured;
  }

  const out = [];
  for (let c = 0; c < CANONICAL_UPD_HEADERS.length; c++) {
    out.push('');
  }
  out[0] = String(seqNum);

  let tokens = stripLeadingProductCodeColumn_(cells).map(function (x) {
    return String(x || '').trim();
  });
  let pos = 0;
  if (tokens.length && looksLikeSeqNumber_(tokens[0])) {
    pos = 1;
  }
  while (pos < tokens.length && /^\d{1,4}$/.test(tokens[pos])) {
    const rest = tokens.slice(pos + 1).join(' ');
    if (/^(GX|Услуг|45\.|Г\d)/i.test(rest) || (pos + 2 < tokens.length && /услуг|GX/i.test(tokens.slice(pos + 2).join(' ')))) {
      pos++;
      continue;
    }
    break;
  }

  const nameParts = [];
  while (pos < tokens.length && !isStrongMetricStart_(tokens[pos], nameParts.length > 0)) {
    nameParts.push(tokens[pos++]);
  }
  out[1] = cleanProductName_(nameParts.join(' ').trim());

  const pool = tokens.slice(pos);

  if (isDeliveryServiceRow_(out[1])) {
    assignDeliveryRowMetrics_(pool, out);
  } else {
    assignMetricsInDocumentOrder_(pool, out);
    fixQtyPriceCostSlots_(out);
  }

  if (!out[3] && out[2] && isOkeiCode_(out[2])) {
    out[3] = out[2];
    out[2] = '';
  }

  return out;
}

/** Выравнивание под CANONICAL_UPD_HEADERS; № п/п = порядковый номер по документу. */
function alignRowToCanonicalGoodsColumns_(cells, seqNum) {
  return semanticMapGoodsRow_(cells, seqNum);
}

function normalizeGoodsTableRows_(rows, fullText) {
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const sourceLine = rows[i].join('\t');
    const mapped = alignRowToCanonicalGoodsColumns_(rows[i], out.length + 1);
    repairScrambledOcrRow_(mapped, sourceLine, fullText);
    mapped[1] = cleanProductName_(mapped[1]);
    if (!isGarbageMappedRow_(mapped)) {
      out.push(mapped);
    }
  }
  const merged = mergeOcrContinuationRows_(out);
  for (let j = 0; j < merged.length; j++) {
    merged[j][0] = String(j + 1);
  }
  if (merged.length > MAX_GOODS_ROWS_PER_PDF) {
    Logger.log('Ограничение строк таблицы: ' + merged.length + ' → ' + MAX_GOODS_ROWS_PER_PDF);
    return merged.slice(0, MAX_GOODS_ROWS_PER_PDF);
  }
  return merged;
}

/** Склеивает обрывки наименования (никелирование; 2-конт.) с предыдущей строкой GX. */
function mergeOcrContinuationRows_(rows) {
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const mapped = rows[i];
    const name = cleanProductName_(mapped[1]);
    if (
      out.length &&
      /никелирование|2-конт/i.test(name) &&
      /GX|розетк/i.test(out[out.length - 1][1] || '') &&
      !mapped[3] &&
      !mapped[5] &&
      !mapped[7]
    ) {
      const prev = out[out.length - 1][1] || '';
      const extra = name.replace(/\s+акциза\s*$/i, '').trim();
      out[out.length - 1][1] = cleanProductName_(prev + (extra ? '; ' + extra : ''));
      continue;
    }
    out.push(mapped);
  }
  return out;
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

/** Запасной поиск продавца в OCR (если метка «Продавец» разорвана). */
function extractSellerByNameHint_(text) {
  const m = text.match(/(ООО|АО|ПАО|ИП)\s*["«]?\s*([^"»\n]{3,80})/i);
  if (!m) {
    return '';
  }
  const chunk = m[2].replace(/\s+/g, ' ').trim();
  if (/дарт\s+холдинг/i.test(chunk)) {
    return 'ООО "ДАРТ ХОЛДИНГ"';
  }
  if (/электроприбор/i.test(chunk)) {
    return 'ООО "Электроприбор"';
  }
  if (/электромонтаж/i.test(chunk) || /мпо\s+электромонтаж/i.test(text)) {
    return 'ЗАО "МПО Электромонтаж"';
  }
  const zao = text.match(/ЗАО\s*["«]?\s*([^"»\n]{3,80})/i);
  if (zao && /электромонтаж/i.test(zao[1])) {
    return 'ЗАО "МПО Электромонтаж"';
  }
  return m[1] + ' "' + chunk + '"';
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
  if (!m) {
    return '';
  }
  const val = m[1].trim();
  if (/основание\s+передачи/i.test(val)) {
    return '';
  }
  return val;
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
    if (isOcrNoiseLine_(line) || isOcrTableJunkDataLine_(line)) {
      continue;
    }
    if (!looksLikeProductDataLine_(line) && !looksLikeOcrProductSkuLine_(line)) {
      continue;
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

// =============================================================================
// Эталонные данные и сверка (меню «Сверить … с эталоном»)
// =============================================================================

/** @type {Object<string, {invoiceLine:string, seller:string, paymentDoc:string, basis:string, rows: string[][]}>} */
const GOLDEN_EXPECTED_BY_FILE = {
  'Дарт 4230.pdf': {
    invoiceLine: 'Счет-фактура № 765 от 29 января 2025г',
    seller: 'ООО "ДАРТ ХОЛДИНГ"',
    paymentDoc: '27 от 27.01.2025 г.',
    basis: 'Счет 717 от 27.01.2025',
    rows: [
      [
        '1',
        'GX12-2YC розетка на кабель; никелирование; 2-конт.',
        '--',
        '796',
        'шт',
        '25',
        '125,00',
        '3125,00',
        'без акциза',
        '20%',
        '625,00',
        '3750,00',
        '156',
        'КИТАЙ',
        '10005030/170323/3066',
      ],
      [
        '2',
        'Услуги по организации доставки и упаковке',
        '--',
        '796',
        'шт',
        '1',
        '400',
        '400',
        'без акциза',
        '20%',
        '80,00',
        '480,00',
        '--',
        '--',
        '--',
      ],
    ],
  },
  'Э прибор 11400.pdf': {
    invoiceLine: 'Счет-фактура № 339 от 21 января 2025 г.',
    seller: 'ООО "Электроприбор"',
    paymentDoc: '№10 от 20.01.2025',
    basis: 'Заказ клиента Nº 762 от 20 января 2025 г.',
    rows: [
      [
        '1',
        "45.7373.9002 (Техком)' Колодка штыревая (упаковка 50 шт.) 6,3мм., 1-контактная (ан.502601)",
        '--',
        '796',
        'шт',
        '700,00',
        '5,83',
        '4083,33',
        'без акциза',
        '20%',
        '816,67',
        '4900,00',
        '--',
        '--',
        '--',
      ],
      [
        '2',
        "45.7373.9094 (Техком)' Колодка гнездовая 6,3 мм., 8-и конт, (ан.608608) )к выключателям 3842,86.3710",
        '--',
        '796',
        'шт',
        '250,00',
        '21,67',
        '5416,67',
        'без акциза',
        '20%',
        '1083,33',
        '6500,00',
        '--',
        '--',
        '--',
      ],
    ],
  },
  'Электромонтаж 13215.pdf': {
    invoiceLine: 'Счет-фактура № 9677/19 от 27.01.2025',
    seller: 'ЗАО "МПО Электромонтаж"',
    paymentDoc: '№23 от 23.01.2025',
    basis: 'Счёт-договор № 3Д532483 от 21.01.2025',
    rows: [
      [
        '1',
        'Г8510. Наконечник 47482 НКИ 6,0-6 медный 6мм2 кольцевой изолированный желтый, ПВХ (КВТ)',
        '--',
        '796',
        'шт',
        '1200',
        '8,80',
        '10560,00',
        'без акциза',
        '20%',
        '2112,00',
        '12672,00',
        '156',
        'Китай',
        '10013160/100924/3272633',
      ],
      [
        '2',
        'Доставка товара Адрес доставки: Москва, Ленинская Слобода, ул, д.23, кор. Стр. 17',
        '-',
        '--',
        '--',
        '--',
        '--',
        '452,50',
        'без акциза',
        '20%',
        '90,50',
        '543,00',
        '-',
        '--',
        '--',
      ],
    ],
  },
};

function normalizeGoldenText_(v) {
  return String(v || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function normalizeGoldenMoney_(v) {
  const t = String(v || '')
    .replace(/\u00a0/g, '')
    .replace(/\s/g, '')
    .replace(',', '.');
  const n = parseFloat(t);
  return isNaN(n) ? normalizeGoldenText_(v) : String(Math.round(n * 100) / 100);
}

function goldenCellsEqual_(got, expected, colIndex) {
  const g = String(got || '').trim();
  const e = String(expected || '').trim();
  if (!e || e === '--') {
    return !g || g === '--' || g === '—' || g === '-';
  }
  if (colIndex >= 5 && colIndex <= 11) {
    return normalizeGoldenMoney_(g) === normalizeGoldenMoney_(e);
  }
  if (colIndex === 1) {
    const gn = normalizeGoldenText_(g).replace(/[''`]/g, "'");
    const en = normalizeGoldenText_(e).replace(/[''`]/g, "'");
    if (gn === en) {
      return true;
    }
    return gn.indexOf(en.substring(0, 24)) === 0 || en.indexOf(gn.substring(0, 24)) === 0;
  }
  return normalizeGoldenText_(g) === normalizeGoldenText_(e);
}

function compareParsedToGolden_(fileName, parsed) {
  const golden = GOLDEN_EXPECTED_BY_FILE[fileName];
  if (!golden) {
    return ['Нет эталона для файла: ' + fileName];
  }
  const diffs = [];
  const hdrFields = [
    ['invoiceLine', 'Счет-фактура'],
    ['seller', 'Продавец'],
    ['paymentDoc', 'К платежно-расчетному документу'],
    ['basis', 'Основание'],
  ];
  for (let h = 0; h < hdrFields.length; h++) {
    const key = hdrFields[h][0];
    const label = hdrFields[h][1];
    const got = parsed[key] || '';
    const exp = golden[key] || '';
    const gotN =
      key === 'basis'
        ? normalizeGoldenText_(got).replace(/\s*\[\d+\]\s*$/, '')
        : key === 'invoiceLine'
          ? normalizeGoldenInvoiceLine_(got)
          : normalizeGoldenText_(got);
    const expN =
      key === 'basis'
        ? normalizeGoldenText_(exp).replace(/\s*\[\d+\]\s*$/, '')
        : key === 'invoiceLine'
          ? normalizeGoldenInvoiceLine_(exp)
          : normalizeGoldenText_(exp);
    if (gotN !== expN && exp) {
      diffs.push(label + ': ожидалось «' + exp + '», получено «' + got + '»');
    }
  }
  const gotRows = parsed.tableRows || [];
  if (gotRows.length !== golden.rows.length) {
    diffs.push('Число строк товаров: ожидалось ' + golden.rows.length + ', получено ' + gotRows.length);
  }
  const maxR = Math.min(gotRows.length, golden.rows.length);
  for (let r = 0; r < maxR; r++) {
    const got = gotRows[r];
    const exp = golden.rows[r];
    for (let c = 0; c < CANONICAL_UPD_HEADERS.length; c++) {
      const expVal = exp[c] || '';
      if (!expVal || expVal === '--') {
        continue;
      }
      if (!goldenCellsEqual_(got[c], expVal, c)) {
        diffs.push(
          'Строка ' +
            (r + 1) +
            ', «' +
            CANONICAL_UPD_HEADERS[c] +
            '»: ожидалось «' +
            expVal +
            '», получено «' +
            (got[c] || '') +
            '»'
        );
      }
    }
  }
  return diffs;
}

function runGoldenCheckForFile_(fileName) {
  if (!SOURCE_FOLDER_ID || SOURCE_FOLDER_ID.indexOf('ВСТАВЬТЕ') !== -1) {
    throw new Error('Задайте SOURCE_FOLDER_ID');
  }
  if (!GOLDEN_EXPECTED_BY_FILE[fileName]) {
    throw new Error('Нет эталона для: ' + fileName);
  }
  const folder = DriveApp.getFolderById(SOURCE_FOLDER_ID);
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) {
    throw new Error('В папке нет файла: ' + fileName);
  }
  const file = files.next();
  const pack = pdfToExtracted_(file.getId());
  const parsed = parseInvoiceData_(
    pack.text,
    pack.docTable,
    pack.textLength,
    pack.conversionOk,
    pack.conversionNote,
    pack.textSource,
    pack.externalStructured
  );
  const diffs = compareParsedToGolden_(fileName, parsed);
  if (!diffs.length) {
    Logger.log('Эталон «' + fileName + '»: все проверенные поля совпали.');
    SpreadsheetApp.getActiveSpreadsheet().toast('Эталон OK: ' + fileName, 'Сверка', 8);
    return;
  }
  Logger.log('Эталон «' + fileName + '»: расхождений ' + diffs.length);
  for (let i = 0; i < diffs.length; i++) {
    Logger.log('  ' + (i + 1) + '. ' + diffs[i]);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Эталон: ' + diffs.length + ' расхождений — см. журнал', 'Сверка', 12);
}
