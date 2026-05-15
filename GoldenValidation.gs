/* eslint-disable no-undef -- функции и константы из InvoicePdfToSheet.gs в одном проекте Apps Script */
/**
 * Эталонные данные и сверка результата парсинга (ручная выгрузка из PDF).
 * Запуск: меню → «Сверить Дарт 4230 с эталоном» или runGoldenCheckDart4230_() в редакторе.
 */

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
    return normalizeGoldenText_(g) === normalizeGoldenText_(e);
  }
  return normalizeGoldenText_(g) === normalizeGoldenText_(e);
}

/**
 * @param {string} fileName
 * @param {{invoiceLine:string, seller:string, paymentDoc:string, basis:string, tableRows:string[][]}} parsed
 * @return {string[]} список расхождений (пусто = OK)
 */
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
    if (normalizeGoldenText_(got) !== normalizeGoldenText_(exp) && exp) {
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

/** Сверка одного PDF из SOURCE_FOLDER_ID с эталоном Дарт 4230. */
function runGoldenCheckDart4230_() {
  const fileName = 'Дарт 4230.pdf';
  if (!SOURCE_FOLDER_ID || SOURCE_FOLDER_ID.indexOf('ВСТАВЬТЕ') !== -1) {
    throw new Error('Задайте SOURCE_FOLDER_ID');
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
    pack.textSource
  );
  const diffs = compareParsedToGolden_(fileName, parsed);
  if (!diffs.length) {
    Logger.log('Эталон «' + fileName + '»: все проверенные поля совпали.');
    SpreadsheetApp.getActiveSpreadsheet().toast('Эталон Дарт 4230: OK', 'Сверка', 8);
    return;
  }
  Logger.log('Эталон «' + fileName + '»: расхождений ' + diffs.length);
  for (let i = 0; i < diffs.length; i++) {
    Logger.log('  ' + (i + 1) + '. ' + diffs[i]);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Эталон: ' + diffs.length + ' расхождений — см. журнал', 'Сверка', 12);
}
