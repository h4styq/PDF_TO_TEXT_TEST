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
/** Порог: выше — сначала upload, при ошибке URL (как в исходном slim). */
const OCR_LARGE_FILE_HINT_BYTES = Math.round(1.5 * 1024 * 1024);

const GEMINI_MODEL = 'gemini-2.5-flash';
const GEMINI_FALLBACK_MODELS = ['gemini-2.0-flash', 'gemini-2.0-flash-lite'];
const GEMINI_TEXT_MAX_ATTEMPTS = 2;
const GEMINI_RETRY_BASE_DELAY_MS = 5000;
const GEMINI_MAX_BACKOFF_MS = 20000;
const GEMINI_OCR_TEXT_MAX_CHARS = 120000;
const PAUSE_BETWEEN_PDF_MS = 20000;
const OCR_SPACE_MAX_ATTEMPTS = 3;
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
      'OCR отдаёт полный сырой текст (без обрезки в скрипте). Gemini раскладывает его в ' +
      CANONICAL_UPD_HEADERS.length +
      ' граф УПД.\n\n' +
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
  const items = [];
  let maxCols = CANONICAL_UPD_HEADERS.length;
  let pauseNext = false;

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
      const parsed = parseInvoiceData_(pack);
      maxCols = Math.max(maxCols, parsed.tableWidth);
      items.push({ fileName: file.getName(), parsed: parsed });
    } catch (e) {
      Logger.log('Ошибка: ' + e.message);
      items.push({
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
  }

  writeParsedRows_(sheet, items, maxCols);
  const n = items.length;
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
    const normalized = normalizeGeminiStructuredText_(g.text);
    if (geminiStructuredResponseIsUsable_(normalized)) {
      return {
        rawOcr: ocr.text,
        structured: normalized,
        source: 'gemini-ocr-structure',
      };
    }
    if (geminiResponseLooksLikeJson_(normalized) || !/===\s*TABLE\s*===/i.test(normalized)) {
      Logger.log('Gemini: ответ не в HEADER/TABLE — повтор (только TAB)…');
      const g2 = tryGeminiStructureFromOcrTextAllModels_(ocr.text, geminiKey, true);
      if (g2 && g2.text) {
        const n2 = normalizeGeminiStructuredText_(g2.text);
        if (geminiStructuredResponseIsUsable_(n2)) {
          return {
            rawOcr: ocr.text,
            structured: n2,
            source: 'gemini-ocr-structure-tab-retry',
          };
        }
      }
    }
    if (normalized.length >= 40) {
      Logger.log('Gemini: сохранён ответ для разбора (' + normalized.length + ' симв.), начало: ' + normalized.substring(0, 280));
      return {
        rawOcr: ocr.text,
        structured: normalized,
        source: 'gemini-ocr-structure-raw',
      };
    }
  }
  Logger.log('Gemini: пустой или слишком короткий ответ.');
  return { rawOcr: ocr.text, structured: '', source: 'ocr.space' };
}

function ocrSpaceCollectText_(json) {
  const parts = [];
  const list = json.ParsedResults || [];
  for (let i = 0; i < list.length; i++) {
    const t = list[i] && list[i].ParsedText;
    if (t) {
      parts.push(String(t));
    }
  }
  return parts.join('\n\n');
}

function ocrSpacePost_(apiKey, payload) {
  for (let attempt = 1; attempt <= OCR_SPACE_MAX_ATTEMPTS; attempt++) {
    if (attempt > 1) {
      Utilities.sleep(1500 * attempt);
    }
    const resp = UrlFetchApp.fetch('https://api.ocr.space/parse/image', {
      method: 'post',
      muteHttpExceptions: true,
      payload: payload,
    });
    const code = resp.getResponseCode();
    const raw = resp.getContentText();
    if (code === 429 || code === 503) {
      Logger.log('OCR.space HTTP ' + code + ', попытка ' + attempt);
      continue;
    }
    if (code !== 200) {
      Logger.log('OCR.space HTTP ' + code + ': ' + raw.substring(0, 350));
      return null;
    }
    let json;
    try {
      json = JSON.parse(raw);
    } catch (e1) {
      continue;
    }
    if (json.IsErroredOnProcessing) {
      Logger.log('OCR.space: ' + (json.ErrorMessage || json.ErrorDetails || ''));
      continue;
    }
    const text = ocrSpaceCollectText_(json);
    if (text.length > 0) {
      return { text: text, json: json };
    }
  }
  return null;
}

/**
 * OCR.space: сначала upload (все страницы ParsedResults), при неудаче — URL Drive.
 * @return {{text:string, viaUrl:boolean}|null}
 */
function tryOcrSpacePdfExtract_(pdfFileId, apiKey) {
  try {
    const file = DriveApp.getFileById(pdfFileId);
    const blob = file.getBlob().setContentType('application/pdf');
    const size = blob.getBytes().length;
    Logger.log('OCR.space: размер файла ' + (size / 1024 / 1024).toFixed(2) + ' МБ');
    const basePayload = {
      apikey: apiKey,
      language: 'rus',
      isOverlayRequired: 'false',
      isTable: 'true',
      detectOrientation: 'true',
      scale: 'true',
      OCREngine: '2',
      file: blob,
    };
    const r = ocrSpacePost_(apiKey, basePayload);
    if (r && r.text && r.text.length > 15) {
      Logger.log('OCR.space upload: ' + r.text.length + ' симв.');
      return { text: r.text, viaUrl: false };
    }
    if (size > OCR_LARGE_FILE_HINT_BYTES) {
      Logger.log('OCR.space: upload не удался для файла > ~1,5 МБ — пробуем URL Drive…');
    } else {
      Logger.log('OCR.space upload не удался — пробуем URL Drive…');
    }
    return tryOcrSpacePdfExtractByDriveUrl_(pdfFileId, apiKey, file);
  } catch (e) {
    Logger.log('OCR.space: ' + e.message);
    return null;
  }
}

/** OCR.space по публичной ссылке Drive (все страницы). */
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
    const r = ocrSpacePost_(apiKey, {
      apikey: apiKey,
      url: driveUrl,
      filetype: 'PDF',
      language: 'rus',
      isOverlayRequired: 'false',
      isTable: 'true',
      detectOrientation: 'true',
      scale: 'true',
      OCREngine: '2',
    });
    if (r && r.text && r.text.length > 15) {
      Logger.log('OCR.space (URL): получен текст (' + r.text.length + ' симв.).');
      return { text: r.text, viaUrl: true };
    }
    Logger.log('OCR.space (URL): не удалось распознать.');
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

function geminiStructuredResponseIsUsable_(text) {
  if (!text || text.length < 20) {
    return false;
  }
  if (/===\s*TABLE\s*===/i.test(text)) {
    const table = parseGeminiTableSection_(text);
    if (table && table.rows && table.rows.length) {
      return true;
    }
  }
  const parsed = parseGeminiJsonResponse_(text, true);
  return !!(parsed && parsed.tableRows && parsed.tableRows.length);
}

function geminiResponseLooksLikeJson_(text) {
  return /"rows"\s*:\s*\[/i.test(text) || /^\s*\{/.test(String(text || '').trim());
}

function normalizeGeminiStructuredText_(text) {
  let t = normalizeText_(text);
  const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) {
    t = fence[1].trim();
  }
  t = t.replace(/\*{1,2}\s*(===\s*HEADER\s*===)\s*\*{0,2}/gi, '$1');
  t = t.replace(/\*{1,2}\s*(===\s*TABLE\s*===)\s*\*{0,2}/gi, '$1');
  t = t.replace(/\*{1,2}\s*(===\s*END\s*===)\s*\*{0,2}/gi, '$1');
  return t;
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
    'По сырому тексту УПД/счёт-фактуры (OCR, возможны ошибки) заполни данные для Google Таблицы.\n' +
    'Ответ только текстом, без markdown, без JSON.\n' +
    '===HEADER===\n' +
    'Счёт-фактура или УПД (№ и дата)\n' +
    'Продавец: …\n' +
    'К платежно-расчетному документу № …\n' +
    'Основание передачи (сдачи) / получения (приемки): …\n' +
    '===TABLE===\n' +
    'Ровно ' +
    CANONICAL_UPD_HEADERS.length +
    ' колонок в строке, разделитель TAB (\\t). Порядок колонок:\n' +
    CANONICAL_UPD_HEADERS.join(' | ') +
    '\n' +
    'Далее все строки товаров из OCR (№ п/п 1, 2, 3…). Не пропускай позиции. Без «Всего к оплате».\n' +
    '===END==='
  );
}

function getGeminiInvoicePrompt_() {
  return (
    getGeminiInvoiceTabPrompt_() +
    '\n\nЕсли TAB неудобен — один валидный JSON: {"invoiceLine":"…","seller":"…","paymentDoc":"…","basis":"…",' +
    '"rows":[{"seq":"1","name":"…","productCode":"","unitCode":"796","unit":"шт","qty":"…","price":"…",' +
    '"costNoVat":"…","excise":"","vatRate":"22%","vatAmount":"…","costWithVat":"…",' +
    '"countryCode":"","countryName":"","declaration":""}]} — предпочитай HEADER/TABLE.'
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
  if (fromJson && fromJson.tableRows.length) {
    let note = pack.conversionNote || '';
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

function parseGeminiJsonResponse_(text, silent) {
  let raw = String(text || '').trim();
  if (!raw) {
    return null;
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
  if (!rows.length) {
    return null;
  }
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

function writeParsedRows_(sheet, items, maxTableCols) {
  sheet.clearContents();
  const tableCols = Math.max(maxTableCols, CANONICAL_UPD_HEADERS.length);
  const globalHeader = [
    'Файл',
    'Счет-фактура (№ и дата)',
    'Продавец',
    'К платежно-расчетному документу №',
  ]
    .concat(CANONICAL_UPD_HEADERS.slice())
    .concat(['Основание передачи / счет']);
  const totalCols = globalHeader.length;

  sheet.getRange(1, 1, 1, totalCols).setValues([globalHeader]);
  let rowPtr = 2;
  for (let i = 0; i < items.length; i++) {
    if (i > 0) {
      rowPtr += BLANK_ROWS_BETWEEN_PDF_FILES;
    }
    const it = items[i];
    const p = it.parsed;
    if (!p.tableRows.length) {
      const single = [it.fileName, p.invoiceLine, p.seller, p.paymentDoc]
        .concat(padRow_([], tableCols))
        .concat([p.basis]);
      sheet.getRange(rowPtr, 1, 1, totalCols).setValues([padRow_(single, totalCols)]);
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
      sheet.getRange(rowPtr, 1, 1, totalCols).setValues([padRow_(row, totalCols)]);
      rowPtr++;
    }
  }
}
