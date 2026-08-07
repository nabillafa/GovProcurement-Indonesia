import fs from 'node:fs';
import assert from 'node:assert/strict';

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
  'totalRow.getCell(1).merge()',
  'cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER)',
  'setCellHorizontalAlignment_(cell, horizontal)',
  'DocumentApp.HorizontalAlignment.CENTER',
  'DocumentApp.HorizontalAlignment.LEFT',
  'DocumentApp.HorizontalAlignment.RIGHT',
]) {
  if (!source.includes(required)) {
    throw new Error(`Missing empty-text safeguard: ${required}`);
  }
}

const PropertiesService = {
  getScriptProperties() {
    return {getProperty() { return ''; }};
  },
};
const helpers = new Function(
  'PropertiesService',
  `${source}\nreturn {docText_, appendSafeTableCell_};`,
)(PropertiesService);

const calls = [];
const row = {
  appendTableCell(...args) {
    calls.push(args);
    return {args};
  },
};

helpers.appendSafeTableCell_(row, '');
helpers.appendSafeTableCell_(row, null);
helpers.appendSafeTableCell_(row, 0);
helpers.appendSafeTableCell_(row, 'TOTAL');

assert.deepEqual(calls, [[], [], ['0'], ['TOTAL']]);
assert.equal(helpers.docText_(undefined), '');
assert.equal(helpers.docText_('Uraian'), 'Uraian');

console.log('Document-generation empty-text safeguards passed.');
