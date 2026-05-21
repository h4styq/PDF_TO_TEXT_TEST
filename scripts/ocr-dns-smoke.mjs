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
