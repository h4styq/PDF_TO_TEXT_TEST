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
  const hasTable = /===\s*TABLE\s*===/i.test(structured);
  const hasHeader = /===\s*HEADER\s*===/i.test(structured);

  return {
    rawOcr: step.rawOcr,
    externalStructured: structured,
    conversionOk: true,
    conversionNote:
      hasTable && hasHeader
        ? ''
        : 'Gemini не вернул HEADER/TABLE — проверьте лог; в листе может быть пустая таблица.',
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
  const g = tryGeminiStructureFromOcrTextAllModels_(ocr.text, geminiKey);
  if (g && g.rateLimited) {
    return { rawOcr: ocr.text, structured: '', source: 'ocr.space', rateLimited: true };
  }
  if (g && g.text && /===\s*HEADER\s*===/i.test(g.text)) {
    return {
      rawOcr: ocr.text,
      structured: normalizeText_(g.text),
      source: 'gemini-ocr-structure',
    };
  }
  Logger.log('Gemini: нет маркеров HEADER/TABLE в ответе.');
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

function tryGeminiStructureFromOcrTextAllModels_(ocrText, apiKey) {
  const models = getGeminiModelsToTry_();
  for (let mi = 0; mi < models.length; mi++) {
    const r = tryGeminiTextExtract_(ocrText, apiKey, models[mi]);
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

function getGeminiInvoicePrompt_() {
  return (
    'По сырому тексту УПД/счёт-фактуры (OCR, возможны ошибки) заполни данные для Google Таблицы.\n' +
    'Ответ только текстом, без markdown.\n' +
    '===HEADER===\n' +
    'Счёт-фактура или УПД (№ и дата)\n' +
    'Продавец: …\n' +
    'К платежно-расчетному документу № …\n' +
    'Основание передачи (сдачи) / получения (приемки): …\n' +
    '===TABLE===\n' +
    'Ровно ' +
    CANONICAL_UPD_HEADERS.length +
    ' колонок (TAB), порядок:\n' +
    CANONICAL_UPD_HEADERS.join(' | ') +
    '\n' +
    'Первая строка — эти заголовки. Далее все строки товаров из OCR (№ п/п 1, 2, 3…). ' +
    'Не пропускай позиции. Без строки «Всего к оплате».\n' +
    '===END==='
  );
}

function tryGeminiTextExtract_(plainText, apiKey, modelName) {
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
    getGeminiInvoicePrompt_() +
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
    return {
      invoiceLine: pack.conversionNote || 'Нет текста OCR',
      seller: '',
      paymentDoc: '',
      tableHeader: CANONICAL_UPD_HEADERS.slice(),
      tableRows: [],
      basis: '',
      tableWidth: CANONICAL_UPD_HEADERS.length,
    };
  }

  const structured = pack.externalStructured || '';
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
        .concat(padRow_(p.tableHeader, tableCols))
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
