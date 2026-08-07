import fs from 'node:fs';

const source = fs.readFileSync(new URL('../PPBJ_Automation.gs', import.meta.url), 'utf8');

const forbidden = [
  ["empty paragraph text insertion", /\.setText\(\s*['\"]\s*['\"]\s*\)/],
  ["unprotected table-cell insertion", /\.appendTableCell\(v\)/],
];

for (const [label, pattern] of forbidden) {
  if (pattern.test(source)) throw new Error(`Unsafe document generation: ${label}`);
}

for (const required of [
  'body.removeChild(paragraph)',
  'appendSafeTableCell_(row, v)',
  'appendSafeTableCell_(totalRow, v)',
  "text === '' ? row.appendTableCell() : row.appendTableCell(text)",
]) {
  if (!source.includes(required)) {
    throw new Error(`Missing empty-text safeguard: ${required}`);
  }
}

console.log('Document-generation empty-text safeguards passed.');
