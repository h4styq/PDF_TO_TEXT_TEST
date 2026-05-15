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

/** Сверка одного PDF из SOURCE_FOLDER_ID с эталоном по имени файла. */
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
    pack.textSource
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

function runGoldenCheckDart4230_() {
  runGoldenCheckForFile_('Дарт 4230.pdf');
}

function runGoldenCheckEpribor11400_() {
  runGoldenCheckForFile_('Э прибор 11400.pdf');
}

function runGoldenCheckElectromontazh13215_() {
  runGoldenCheckForFile_('Электромонтаж 13215.pdf');
}
