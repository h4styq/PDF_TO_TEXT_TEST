/**
 * УПД / счета-фактуры: PDF на Drive → OCR.space (сырой текст) → Gemini (колонки) → Google Таблица.
 *
 * Свойства скрипта: OCR_SPACE_API_KEY, GEMINI_API_KEY
 * Сервис: Google Drive API v3 (Расширения → Сервисы).
 */
const SOURCE_FOLDER_ID = 'ВСТАВЬТЕ_ID_ПАПКИ';
const OUTPUT_SHEET_NAME = 'Счета_фактуры';
const BLANK_ROWS_BETWEEN_PDF_FILES = 2;
const SCRIPT_VERSION = '2026-05-21-slim-ocr-gemini';

const GEMINI_MODEL = 'gemini-2.5-flash';
const GEMINI_FALLBACK_MODELS = ['gemini-2.0-flash', 'gemini-2.0-flash-lite'];
const GEMINI_TEXT_MAX_ATTEMPTS = 2;
const GEMINI_RETRY_BASE_DELAY_MS = 5000;
const GEMINI_MAX_BACKOFF_MS = 20000;
const GEMINI_OCR_TEXT_MAX_CHARS = 120000;
const PAUSE_BETWEEN_PDF_MS = 20000;
const OCR_SPACE_MAX_ATTEMPTS = 3;
/** Запись на лист пакетами (меньше 429 от Google Sheets API). */
const SHEETS_FLUSH_EVERY_N_PDF = 25;
// Если PDF > ~1 МБ — грузим не file upload'ом, а через URL Drive.
// Это обычно стабильнее в Apps Script.
const OCR_TRY_DRIVE_URL_FOR_LARGE = true;

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

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Счета-фактуры (PDF)')
    .addItem('Загрузить из папки Drive', 'runProcessFolder')
    .addItem('Как подключить (OCR + Gemini)', 'showRecognitionSetupHelp')
    .addToUi();
}

function runProcessFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Откройте таблицу с привязанным скриптом.');
  }
  processFolderIntoSpreadsheet_(SOURCE_FOLDER_ID, ss.getId());
}

function runProcessFolderForSpreadsheet(spreadsheetId) {
  processFolderIntoSpreadsheet_(SOURCE_FOLDER_ID, spreadsheetId);
}

function showRecognitionSetupHelp() {
  SpreadsheetApp.getUi().alert(
    'PDF → OCR.space → Gemini → лист «' +
      OUTPUT_SHEET_NAME +
      '»\n\n' +
      'Свойства скрипта (Apps Script → Свойства проекта → Свойства скрипта):\n' +
      '• OCR_SPACE_API_KEY — https://ocr.space/ocrapi\n' +
      '• GEMINI_API_KEY — https://aistudio.google.com/apikey\n\n' +
      'Схема: OCR.space (сырой текст) → Gemini (структура, JSON или HEADER/TABLE).\n' +
      'Запись в таблицу — одним батчем setValues (не построчно), пакетами по ' +
      SHEETS_FLUSH_EVERY_N_PDF +
      ' PDF.\n\n' +
      'Пауза между PDF: ' +
      Math.round(PAUSE_BETWEEN_PDF_MS / 1000) +
      ' с (лимит Gemini).\n' +
      'Версия: ' +
      SCRIPT_VERSION
  );
}

function processFolderIntoSpreadsheet_(folderId, spreadsheetId) {
  Logger.log('Старт ' + SCRIPT_VERSION + ', папка=' + folderId);
  if (!folderId || folderId.indexOf('ВСТАВЬТЕ') !== -1) {
    const msg = 'Задайте SOURCE_FOLDER_ID в коде.';
    SpreadsheetApp.openById(spreadsheetId).toast(msg, 'Счета-фактуры', 12);
    throw new Error(msg);
  }

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('OCR_SPACE_API_KEY') || !props.getProperty('GEMINI_API_KEY')) {
    const msg = 'Нужны OCR_SPACE_API_KEY и GEMINI_API_KEY в свойствах скрипта.';
    SpreadsheetApp.openById(spreadsheetId).toast(msg, 'Счета-фактуры', 12);
    throw new Error(msg);
  }

  const folder = DriveApp.getFolderById(folderId);
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(OUTPUT_SHEET_NAME);
  }

  const files = folder.getFilesByType(MimeType.PDF);
  const pending = [];
  let pdfCount = 0;
  let pauseNext = false;
  let sheetRowsWritten = 0;

  sheet.clearContents();

  while (files.hasNext()) {
    if (pauseNext && PAUSE_BETWEEN_PDF_MS > 0) {
      Utilities.sleep(PAUSE_BETWEEN_PDF_MS);
    }
    pauseNext = false;

    const file = files.next();
    Logger.log('PDF: ' + file.getName());
    try {
      const pack = pdfToExtracted_(file.getId());
      pauseNext = true;
      pending.push({ fileName: file.getName(), parsed: parseInvoiceData_(pack) });
    } catch (e) {
      Logger.log('Ошибка: ' + e.message);
      pending.push({
        fileName: file.getName(),
        parsed: {
          invoiceLine: 'ОШИБКА: ' + e.message,
          seller: '',
          paymentDoc: '',
          tableHeader: CANONICAL_UPD_HEADERS.slice(),
          tableRows: [],
          basis: '',
          tableWidth: CANONICAL_UPD_HEADERS.length,
        },
      });
    }
    pdfCount++;
    if (pending.length >= SHEETS_FLUSH_EVERY_N_PDF) {
      sheetRowsWritten = flushParsedItemsToSheet_(sheet, pending, sheetRowsWritten);
    }
  }

  if (pending.length) {
    sheetRowsWritten = flushParsedItemsToSheet_(sheet, pending, sheetRowsWritten);
  } else if (!sheetRowsWritten) {
    writeSheetHeaderOnly_(sheet);
  }

  const n = pdfCount;
  const summary =
    n === 0
      ? 'PDF не найдены.'
      : 'Готово: ' + n + ' PDF (OCR → Gemini). Лист «' + OUTPUT_SHEET_NAME + '».';
  Logger.log(summary);
  ss.toast(summary, 'Счета-фактуры', 12);
}

/**
 * @return {{rawOcr:string, externalStructured:string, conversionOk:boolean, conversionNote:string, textSource:string}}
 */
function pdfToExtracted_(pdfFileId) {
  const step = tryOcrThenGeminiExtract_(pdfFileId);
  if (!step || !step.rawOcr || step.rawOcr.length < 30) {
    return {
      rawOcr: '',
      externalStructured: '',
      conversionOk: false,
      conversionNote: 'OCR.space не вернул текст (проверьте ключ, доступ к файлу, тариф OCR).',
      textSource: 'none',
    };
  }

  const structured = step.structured || '';
  const okStruct = structured ? geminiStructuredResponseIsUsable_(structured) : false;

  return {
    rawOcr: step.rawOcr,
    externalStructured: structured,
    conversionOk: true,
    conversionNote: okStruct
      ? ''
      : 'Gemini не вернул пригодный JSON/HEADER/TABLE — проверьте лог; таблица может быть пустой.',
    textSource: step.source || 'ocr.space',
  };
}

/**
 * @return {{rawOcr:string, structured:string, source:string, rateLimited?:boolean}|null}
 */
function tryOcrThenGeminiExtract_(pdfFileId) {
  const ocrKey = PropertiesService.getScriptProperties().getProperty('OCR_SPACE_API_KEY');
  const geminiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!ocrKey) {
    return null;
  }

  Logger.log('OCR.space: распознавание PDF (полный текст)…');
  const ocr = tryOcrSpacePdfExtract_(pdfFileId, ocrKey);
  if (!ocr || !ocr.text) {
    return null;
  }
  Logger.log('OCR.space: ' + ocr.text.length + ' симв.' + (ocr.viaUrl ? ' (URL Drive)' : ' (upload)'));

  if (!geminiKey) {
    return { rawOcr: ocr.text, structured: '', source: 'ocr.space' };
  }

  Logger.log('Gemini: разбор OCR-текста в колонки…');
  const g = tryGeminiStructureFromOcrTextAllModels_(ocr.text, geminiKey, false);
  if (g && g.rateLimited) {
    return { rawOcr: ocr.text, structured: '', source: 'ocr.space', rateLimited: true };
  }
  if (g && g.text) {
    const normalized = normalizeText_(g.text);
    if (geminiStructuredResponseIsUsable_(normalized)) {
      return {
        rawOcr: ocr.text,
        structured: normalized,
        source: 'gemini-ocr-structure',
      };
    }
    if (geminiResponseLooksLikeJson_(normalized)) {
      Logger.log('Gemini: битый JSON — повтор с форматом HEADER/TABLE (TAB)…');
      const g2 = tryGeminiStructureFromOcrTextAllModels_(ocr.text, geminiKey, true);
      if (g2 && g2.text && geminiStructuredResponseIsUsable_(normalizeText_(g2.text))) {
        return {
          rawOcr: ocr.text,
          structured: normalizeText_(g2.text),
          source: 'gemini-ocr-structure-tab-retry',
        };
      }
    }
  }
  Logger.log('Gemini: нет пригодного JSON/HEADER/TABLE в ответе.');
  return { rawOcr: ocr.text, structured: '', source: 'ocr.space' };
}

/**
 * OCR.space: распознавание PDF.
 * Логика как раньше: если PDF > ~1 МБ — используем Drive URL (стабильнее для Apps Script).
 * @return {{text:string, viaUrl:boolean}|null}
 */
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
      Logger.log('OCR.space: файл больше ~1 МБ (бесплатный тариф). ' + 'Сожмите PDF или задайте OCR_TRY_DRIVE_URL_FOR_LARGE = true.');
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
      return { text: txt, viaUrl: false };
    }
    return null;
  } catch (e) {
    Logger.log('OCR.space: ' + e.message);
    return null;
  }
}

/**
 * OCR.space: для PDF > ~1 МБ — распознавание по URL Drive.
 * @return {{text:string, viaUrl:boolean}|null}
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
        return { text: pr.ParsedText, viaUrl: true };
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

function isDeprecatedGeminiModel_(name) {
  return /^gemini-1\.5-(flash|pro)(-|$)/i.test(name || '') || name === 'gemini-1.5-flash' || name === 'gemini-1.5-pro';
}

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
    if (!m || out.indexOf(m) !== -1 || isDeprecatedGeminiModel_(m)) {
      continue;
    }
    out.push(m);
  }
  return out.length ? out : [GEMINI_MODEL];
}

function geminiBackoffMs_(attempt, resp) {
  try {
    const h = resp.getHeaders();
    const raw = h['Retry-After'] || h['retry-after'];
    if (raw) {
      const sec = parseInt(String(raw).replace(/[^0-9]/g, ''), 10);
      if (!isNaN(sec) && sec > 0) {
        return Math.min(sec * 1000, GEMINI_MAX_BACKOFF_MS);
      }
    }
  } catch (ignore) {}
  return Math.min(GEMINI_RETRY_BASE_DELAY_MS * Math.pow(2, attempt - 1), GEMINI_MAX_BACKOFF_MS);
}

/** Ответ пригоден для parseInvoiceData_: валидный JSON с строками или блок HEADER/TABLE. */
function geminiStructuredResponseIsUsable_(text) {
  if (!text || text.length < 20) {
    return false;
  }
  if (/===\s*HEADER\s*===/i.test(text) && /===\s*TABLE\s*===/i.test(text)) {
    return true;
  }
  const parsed = parseGeminiJsonResponse_(text, true);
  return !!(parsed && parsed.tableRows && parsed.tableRows.length);
}

function geminiResponseLooksLikeJson_(text) {
  return /"rows"\s*:\s*\[/i.test(text) || /^\s*\{/.test(String(text || '').trim());
}

function tryGeminiStructureFromOcrTextAllModels_(ocrText, apiKey, tabFormatOnly) {
  const models = getGeminiModelsToTry_();
  for (let mi = 0; mi < models.length; mi++) {
    const r = tryGeminiTextExtract_(ocrText, apiKey, models[mi], tabFormatOnly);
    if (r && r.rateLimited) {
      return { rateLimited: true };
    }
    if (r && r.text) {
      r.model = models[mi];
      return r;
    }
    if (mi < models.length - 1) {
      Utilities.sleep(2000);
    }
  }
  return null;
}

function getGeminiInvoiceTabPrompt_() {
  return (
    'Ты получаешь СЫРОЙ текст УПД/счёт-фактуры после OCR (таблицы могут быть «кашей» — восстанови структуру).\n' +
    'Верни ТОЛЬКО текст в формате ниже (без markdown, без ```, без JSON):\n\n' +
    '===HEADER===\n' +
    'Счёт-фактура № … от …\n' +
    'Продавец: …\n' +
    'К платежно-расчетному документу № …\n' +
    'Основание передачи …\n' +
    '===TABLE===\n' +
    'Строки таблицы: ровно ' +
    CANONICAL_UPD_HEADERS.length +
    ' колонок в строке, разделитель — символ TAB (\\t).\n' +
    'Первая строка TABLE — заголовки как в УПД (можно опустить, если ясно из OCR).\n' +
    'Далее — ВСЕ позиции товаров по № п/п, без строки «Всего к оплате».\n' +
    '===END===\n' +
    'Не выдумывай суммы. Числа — как в OCR.'
  );
}

function getGeminiInvoicePrompt_() {
  const colKeys = [
    'seq',
    'name',
    'productCode',
    'unitCode',
    'unit',
    'qty',
    'price',
    'costNoVat',
    'excise',
    'vatRate',
    'vatAmount',
    'costWithVat',
    'countryCode',
    'countryName',
    'declaration',
  ];
  return (
    getGeminiInvoiceTabPrompt_() +
    '\n\nЗапасной формат (только если TAB неудобен): валидный JSON без markdown:\n' +
    '{"invoiceLine":"…","seller":"…","paymentDoc":"…","basis":"…","rows":[' +
    '{"seq":"1","name":"…","productCode":"","unitCode":"796","unit":"шт","qty":"…","price":"…",' +
    '"costNoVat":"…","excise":"","vatRate":"22%","vatAmount":"…","costWithVat":"…",' +
    '"countryCode":"","countryName":"","declaration":""}]}\n' +
    'Ключи rows: ' +
    colKeys.join(', ') +
    '. Предпочитай TAB-формат HEADER/TABLE.'
  );
}

function tryGeminiTextExtract_(plainText, apiKey, modelName, tabFormatOnly) {
  const model = modelName || GEMINI_MODEL;
  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    model +
    ':generateContent?key=' +
    encodeURIComponent(apiKey);
  const snippet =
    plainText.length > GEMINI_OCR_TEXT_MAX_CHARS
      ? plainText.substring(0, GEMINI_OCR_TEXT_MAX_CHARS)
      : plainText;
  if (plainText.length > GEMINI_OCR_TEXT_MAX_CHARS) {
    Logger.log(
      'Gemini: OCR-текст обрезан до ' + GEMINI_OCR_TEXT_MAX_CHARS + ' симв. (лимит запроса Apps Script).'
    );
  }
  const prompt =
    (tabFormatOnly ? getGeminiInvoiceTabPrompt_() : getGeminiInvoicePrompt_()) +
    '\n\n--- Полный OCR-текст документа ---\n\n' +
    snippet;

  for (let attempt = 1; attempt <= GEMINI_TEXT_MAX_ATTEMPTS; attempt++) {
    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0, maxOutputTokens: 8192 },
      }),
    });
    const code = resp.getResponseCode();
    if (code === 404) {
      Logger.log('Gemini 404: модель ' + model);
      return null;
    }
    if (code === 429 || code === 503) {
      if (attempt >= GEMINI_TEXT_MAX_ATTEMPTS) {
        return { rateLimited: true };
      }
      Utilities.sleep(geminiBackoffMs_(attempt, resp));
      continue;
    }
    if (code !== 200) {
      Logger.log('Gemini HTTP ' + code);
      return null;
    }
    let json;
    try {
      json = JSON.parse(resp.getContentText());
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
      return { text: out };
    }
  }
  return null;
}

function parseInvoiceData_(pack) {
  if (!pack.conversionOk) {
    return emptyParsedInvoice_(pack.conversionNote || 'Нет текста OCR');
  }

  const structured = pack.externalStructured || '';
  const fromJson = parseGeminiJsonResponse_(structured, false);
  if (fromJson) {
    let note = pack.conversionNote || '';
    if (!fromJson.tableRows.length) {
      note = (note ? note + ' ' : '') + 'JSON без строк товаров.';
    }
    if (note) {
      fromJson.invoiceLine = (fromJson.invoiceLine + ' ' + note).trim();
    }
    return fromJson;
  }

  const hdr = parseStructuredHeaderBlock_(structured) || {};
  let table = parseGeminiTableSection_(structured);
  let rows = table && table.rows ? table.rows : [];
  rows = alignRowsToCanonical_(rows);

  let note = pack.conversionNote || '';
  if (!rows.length) {
    note = (note ? note + ' ' : '') + 'Таблица пуста — проверьте ответ Gemini в логе.';
  }

  return {
    invoiceLine: (hdr.invoiceLine || '') + (note ? ' ' + note : ''),
    seller: hdr.seller || '',
    paymentDoc: hdr.paymentDoc || '',
    tableHeader: CANONICAL_UPD_HEADERS.slice(),
    tableRows: rows,
    basis: hdr.basis || '',
    tableWidth: CANONICAL_UPD_HEADERS.length,
  };
}

function emptyParsedInvoice_(note) {
  return {
    invoiceLine: note,
    seller: '',
    paymentDoc: '',
    tableHeader: CANONICAL_UPD_HEADERS.slice(),
    tableRows: [],
    basis: '',
    tableWidth: CANONICAL_UPD_HEADERS.length,
  };
}

/**
 * @param {string} text
 * @param {boolean} silent — не писать в лог (для проверки usable)
 */
function parseGeminiJsonResponse_(text, silent) {
  let raw = String(text || '').trim();
  if (!raw) {
    return null;
  }
  const fence = raw.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) {
    raw = fence[1].trim();
  }
  const start = raw.indexOf('{');
  const end = raw.lastIndexOf('}');
  if (start === -1) {
    return null;
  }
  const jsonSlice = end > start ? raw.substring(start, end + 1) : raw.substring(start);
  let obj = tryParseGeminiJsonObject_(jsonSlice, silent);
  if (!obj) {
    const rowObjs = extractRowObjectsFromJsonText_(raw);
    if (rowObjs.length) {
      obj = extractScalarFieldsFromJsonText_(raw) || {};
      obj.rows = rowObjs;
      if (!silent) {
        Logger.log('Gemini JSON: извлечено строк из битого ответа: ' + rowObjs.length);
      }
    }
  }
  if (!obj) {
    return null;
  }
  return geminiJsonObjectToParsed_(obj);
}

function tryParseGeminiJsonObject_(jsonSlice, silent) {
  try {
    return JSON.parse(jsonSlice);
  } catch (e) {
    if (!silent) {
      Logger.log('Gemini JSON: ' + e.message);
    }
  }
  const repaired = repairGeminiJsonString_(jsonSlice);
  if (repaired !== jsonSlice) {
    try {
      const obj = JSON.parse(repaired);
      if (!silent) {
        Logger.log('Gemini JSON: восстановлено (repair).');
      }
      return obj;
    } catch (e2) {
      if (!silent) {
        Logger.log('Gemini JSON (repair): ' + e2.message);
      }
    }
  }
  return null;
}

function repairGeminiJsonString_(s) {
  let t = String(s || '').trim();
  t = t.replace(/,\s*([}\]])/g, '$1');
  t = t
    .split('')
    .map(function (ch) {
      const code = ch.charCodeAt(0);
      return code < 32 && code !== 9 && code !== 10 && code !== 13 ? ' ' : ch;
    })
    .join('');
  const openBraces = (t.match(/\{/g) || []).length;
  const closeBraces = (t.match(/\}/g) || []).length;
  const openBrackets = (t.match(/\[/g) || []).length;
  const closeBrackets = (t.match(/\]/g) || []).length;
  for (let i = 0; i < openBrackets - closeBrackets; i++) {
    t += ']';
  }
  for (let j = 0; j < openBraces - closeBraces; j++) {
    t += '}';
  }
  return t;
}

function extractScalarFieldsFromJsonText_(text) {
  const out = {};
  const fields = ['invoiceLine', 'seller', 'paymentDoc', 'basis'];
  for (let i = 0; i < fields.length; i++) {
    const key = fields[i];
    const re = new RegExp('"' + key + '"\\s*:\\s*"((?:[^"\\\\]|\\\\.)*)"', 'i');
    const m = String(text || '').match(re);
    if (m) {
      out[key] = m[1].replace(/\\"/g, '"').replace(/\\n/g, '\n');
    }
  }
  return Object.keys(out).length ? out : null;
}

/** Из битого JSON вытаскивает целые объекты в массиве rows. */
function extractRowObjectsFromJsonText_(text) {
  const slice = String(text || '');
  const rowsKey = slice.search(/"rows"\s*:\s*\[/i);
  const searchFrom = rowsKey >= 0 ? rowsKey : 0;
  const out = [];
  let depth = 0;
  let start = -1;
  for (let i = searchFrom; i < slice.length; i++) {
    const ch = slice.charAt(i);
    if (ch === '{') {
      if (depth === 0) {
        start = i;
      }
      depth++;
    } else if (ch === '}') {
      depth--;
      if (depth === 0 && start >= 0) {
        const chunk = slice.substring(start, i + 1);
        if (/"(name|seq)"\s*:/i.test(chunk)) {
          try {
            const row = JSON.parse(chunk);
            if (row && (row.name || row.seq)) {
              out.push(row);
            }
          } catch (ignore) {
            const fixed = repairGeminiJsonString_(chunk);
            try {
              const row2 = JSON.parse(fixed);
              if (row2 && (row2.name || row2.seq)) {
                out.push(row2);
              }
            } catch (ignore2) {}
          }
        }
        start = -1;
      }
    }
  }
  return out;
}

function geminiJsonObjectToParsed_(obj) {
  if (!obj || typeof obj !== 'object') {
    return null;
  }
  const rowsIn = obj.rows || obj.tableRows || obj.items || [];
  if (!rowsIn || !rowsIn.length) {
    return null;
  }
  const matrix = [];
  for (let i = 0; i < rowsIn.length; i++) {
    const row = rowsIn[i];
    if (Object.prototype.toString.call(row) === '[object Array]') {
      matrix.push(row);
      continue;
    }
    if (row && typeof row === 'object') {
      matrix.push([
        row.seq != null ? String(row.seq) : '',
        row.name != null ? String(row.name) : '',
        row.productCode != null ? String(row.productCode) : '',
        row.unitCode != null ? String(row.unitCode) : '',
        row.unit != null ? String(row.unit) : '',
        row.qty != null ? String(row.qty) : '',
        row.price != null ? String(row.price) : '',
        row.costNoVat != null ? String(row.costNoVat) : '',
        row.excise != null ? String(row.excise) : '',
        row.vatRate != null ? String(row.vatRate) : '',
        row.vatAmount != null ? String(row.vatAmount) : '',
        row.costWithVat != null ? String(row.costWithVat) : '',
        row.countryCode != null ? String(row.countryCode) : '',
        row.countryName != null ? String(row.countryName) : '',
        row.declaration != null ? String(row.declaration) : '',
      ]);
    }
  }
  const rows = alignRowsToCanonical_(matrix);
  return {
    invoiceLine: String(obj.invoiceLine || obj.invoice || ''),
    seller: String(obj.seller || ''),
    paymentDoc: String(obj.paymentDoc || obj.payment || ''),
    basis: String(obj.basis || ''),
    tableHeader: CANONICAL_UPD_HEADERS.slice(),
    tableRows: rows,
    tableWidth: CANONICAL_UPD_HEADERS.length,
  };
}

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
    if (!invoiceLine && /^(Сч[её]т|УПД|Универсальн)/i.test(line)) {
      invoiceLine = line;
    } else if (!seller && /\bПродавец\s*:?/i.test(line)) {
      seller = line.replace(/^.*?Продавец\s*:?\s*/i, '').trim();
    } else if (!seller && /^(ООО|АО|ЗАО|ПАО|ИП)\s/i.test(line)) {
      seller = line;
    } else if (!paymentDoc && /платежно[-\s]*расчетному\s+документу/i.test(line)) {
      paymentDoc = line.replace(/^.*документу\s*№?\s*/i, '').trim();
    } else if (!basis && /Основание\s+передачи/i.test(line)) {
      basis = line.replace(/^.*при[её]мки\)\s*/i, '').trim();
    }
  }
  return { invoiceLine: invoiceLine, seller: seller, paymentDoc: paymentDoc, basis: basis };
}

function parseGeminiTableSection_(text) {
  const n = normalizeText_(text);
  const tm = n.match(/===\s*TABLE\s*===\s*([\s\S]*?)(?====\s*END\s*===|$)/i);
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
  if (/наименован|№\s*п\/п/i.test(lines[0]) && lines[0].indexOf('\t') === -1) {
    start = 1;
  }
  const rows = [];
  for (let i = start; i < lines.length; i++) {
    if (/Всего\s+к\s+оплате|^Итого\b/i.test(lines[i])) {
      break;
    }
    const cells = splitTableLine_(lines[i]);
    if (cells.length) {
      rows.push(cells);
    }
  }
  if (!rows.length) {
    return null;
  }
  return { header: CANONICAL_UPD_HEADERS.slice(), rows: rows, width: CANONICAL_UPD_HEADERS.length };
}

function alignRowsToCanonical_(rows) {
  const w = CANONICAL_UPD_HEADERS.length;
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const cells = padRow_(rows[i], w);
    const name = String(cells[1] || '').trim();
    if (/^(наименование|№\s*п\/п|код\s+вида)/i.test(name)) {
      continue;
    }
    out.push(cells.slice(0, w));
  }
  return out;
}

function normalizeText_(t) {
  return String(t || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n');
}

function splitTableLine_(line) {
  if (line.indexOf('\t') !== -1) {
    return line.split('\t').map(function (c) {
      return c.trim();
    });
  }
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

function getSheetGlobalHeader_() {
  return [
    'Файл',
    'Счет-фактура (№ и дата)',
    'Продавец',
    'К платежно-расчетному документу №',
  ]
    .concat(CANONICAL_UPD_HEADERS.slice())
    .concat(['Основание передачи / счет']);
}

function getSheetTotalCols_() {
  return getSheetGlobalHeader_().length;
}

function buildSheetDataRowsMatrix_(items) {
  const tableCols = CANONICAL_UPD_HEADERS.length;
  const totalCols = getSheetTotalCols_();
  const matrix = [];
  for (let i = 0; i < items.length; i++) {
    if (i > 0) {
      for (let b = 0; b < BLANK_ROWS_BETWEEN_PDF_FILES; b++) {
        matrix.push(padRow_([], totalCols));
      }
    }
    const it = items[i];
    const p = it.parsed;
    if (!p.tableRows.length) {
      matrix.push(
        padRow_(
          [it.fileName, p.invoiceLine, p.seller, p.paymentDoc].concat(padRow_([], tableCols)).concat([p.basis]),
          totalCols
        )
      );
      continue;
    }
    for (let r = 0; r < p.tableRows.length; r++) {
      const tr = padRow_(p.tableRows[r], tableCols);
      matrix.push(
        padRow_(
          [
            r === 0 ? it.fileName : '',
            r === 0 ? p.invoiceLine : '',
            r === 0 ? p.seller : '',
            r === 0 ? p.paymentDoc : '',
          ]
            .concat(tr)
            .concat([r === 0 ? p.basis : '']),
          totalCols
        )
      );
    }
  }
  return matrix;
}

/** @return {number} число строк на листе после записи */
function flushParsedItemsToSheet_(sheet, items, rowsAlreadyOnSheet) {
  if (!items.length) {
    return rowsAlreadyOnSheet;
  }
  const data = buildSheetDataRowsMatrix_(items);
  const totalCols = getSheetTotalCols_();
  if (!rowsAlreadyOnSheet) {
    const block = [getSheetGlobalHeader_()].concat(data);
    sheet.getRange(1, 1, block.length, totalCols).setValues(block);
    Logger.log('Sheets: записан блок ' + block.length + ' строк (шапка + ' + items.length + ' PDF).');
    items.length = 0;
    return block.length;
  }
  sheet.getRange(rowsAlreadyOnSheet + 1, 1, data.length, totalCols).setValues(data);
  Logger.log('Sheets: дописано ' + data.length + ' строк (' + items.length + ' PDF).');
  items.length = 0;
  return rowsAlreadyOnSheet + data.length;
}

function writeSheetHeaderOnly_(sheet) {
  const header = getSheetGlobalHeader_();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
}
