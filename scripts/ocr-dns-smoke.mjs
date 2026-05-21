/**
 * Smoke test for DNS-style flat OCR lines (no Apps Script runtime).
 */
function splitLineByOcrRowNumbers(line) {
  const l = String(line || '').trim();
  if (!l || l.length < 50) return [l];
  const starts = [];
  const re = /(?:^|\s)(\d{1,2})\s+(?=[A-Za-zА-ЯёЁ(])/g;
  let m;
  while ((m = re.exec(l)) !== null) {
    const idx = m.index + (m[0].charAt(0) === ' ' ? 1 : 0);
    if (!starts.length || idx > starts[starts.length - 1] + 12) starts.push(idx);
  }
  if (starts.length < 2) return [l];
  const out = [];
  for (let i = 0; i < starts.length; i++) {
    out.push(l.substring(starts[i], i + 1 < starts.length ? starts[i + 1] : l.length).trim());
  }
  return out.length >= 2 ? out : [l];
}

const row =
  '1 Lech G703 Black (910-005644) - CN - 796 шт. 1 4 999,17 4 999,17 акциза 20% 999,83 5 999,00 156 Китай |54020 2 Lech G703 Black (910-005644) - CN - 796 шт. 1 4 999,17 4 999,17 акциза 20% 999,83 5 999,00 156 Китай |54021 3 Lech G703 Black (910-005644) - CN - 796 шт. 1 4 999,17 4 999,17 акциза 20% 999,83 5 999,00 156 Китай |54022 4 Lech G703 Black (910-005644) - CN - 796 шт. 1 4 999,17 4 999,17 акциза 20% 999,83 5 999,00 156 Китай |54023 5 Lech G703 Black (910-005644) - CN - 796 шт. 1 4 999,17 4 999,17 акциза 20% 999,83 5 999,00 156 Китай |54024';

const parts = splitLineByOcrRowNumbers(row);
if (parts.length !== 5) {
  console.error('FAIL: expected 5 rows, got', parts.length);
  process.exit(1);
}
for (let i = 0; i < parts.length; i++) {
  if (!parts[i].startsWith(String(i + 1) + ' ')) {
    console.error('FAIL: row', i + 1, 'prefix', parts[i].slice(0, 20));
    process.exit(1);
  }
}
console.log('OK: split into', parts.length, 'product lines');

function stripProductNameAtOkeiMarker(name) {
  let n = String(name || '').replace(/\s+/g, ' ').trim();
  const okei = n.search(/\s796\s*(?:шт\.?|ШТ|wm|wт)(?:\s|$)/i);
  if (okei > 4) n = n.substring(0, okei).replace(/\s*-\s*$/g, '').trim();
  const money = n.search(/\d{1,3}(?:\s\d{3})*[.,]\d{2}/);
  if (money > 8) n = n.substring(0, money).replace(/\s*-\s*$/g, '').trim();
  return n;
}

const badName =
  'Lech G703 Black (910-005644) - CN - 796 шт. 1 4 999,17 4 999,17 акциза 20% 999,83 5 999,00 156 Китай |54020';
const good = stripProductNameAtOkeiMarker(badName);
if (good !== 'Lech G703 Black (910-005644) - CN -' && good !== 'Lech G703 Black (910-005644) - CN') {
  console.error('FAIL: strip name got', good);
  process.exit(1);
}
console.log('OK: product name strip');

const flat =
  '1 Товар разный 796 шт 1 100,00 3 Товар три 796 шт 2 200,00 4 Мышь беспроводная Logitech G703 Black (910-005644) - CN - 796 шт. 1 4 999,17 4 999,17 5 Мышь беспроводная Logitech G703 Black (910-005644) - CN - 796 шт. 1 4 999,17 всего к оплате';
const esc = '910-005644'.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const re = new RegExp(
  '(?:^|\\s)4\\s+([А-Яа-яЁё][\\s\\S]{8,220}?(?:Logitech|Lech|G703)[\\s\\S]{0,120}?\\(' +
    esc +
    '\\)[\\s\\S]{0,40}?-?\\s*CN)',
  'i'
);
const m4 = flat.match(re);
if (!m4 || !/Мышь беспроводная Logitech G703 Black/.test(m4[1])) {
  console.error('FAIL: row 4 full name in flat', m4 && m4[1]);
  process.exit(1);
}
const re5 = /(?:^|\s)5\s+([\s\S]{20,400}?)(?=\sвсего\s+к\s+оплате|$)/i;
const m5 = flat.match(re5);
if (!m5 || !/910-005644/.test(m5[1])) {
  console.error('FAIL: row 5 supplement', m5 && m5[1]);
  process.exit(1);
}
console.log('OK: row 4 name and row 5 scan');

function stripUpdColumnNumbersFromName(name) {
  const tokens = String(name || '')
    .replace(/\s+/g, ' ')
    .trim()
    .split(/\s+/);
  let start = 0;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (/^\d{1,2}a?$/i.test(t) || (/^\d{3,4}$/.test(t) && t !== '796')) continue;
    if (/^[A-Za-zА-ЯЁа-яё]/.test(t) && !/^(шт|wm|без|акциза)$/i.test(t)) {
      start = i;
      break;
    }
    if (/^[A-Z0-9]{2,}[-/]/.test(t) || /^[A-Z]\d/.test(t)) {
      start = i;
      break;
    }
  }
  let out = tokens.slice(start).join(' ');
  out = out.replace(/^\d{1,2}\s+/, '');
  return out.trim();
}
const junk =
  '2a 3 4 5 6 7 8 9 10 10a 11 0912 1 Кабель 5D-FB CU PVC 006 м';
const clean = stripUpdColumnNumbersFromName(junk);
if (!/^Кабель 5D-FB/.test(clean)) {
  console.error('FAIL: column strip got', clean);
  process.exit(1);
}
console.log('OK: strip UPD column numbers from name');
