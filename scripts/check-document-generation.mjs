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
  'mergeLeadingTableCells_(totalRow, 5)',
  'for (let c = count - 1; c >= 1; c--) row.getCell(c).merge()',
  '.setBold(i === 0)',
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
  `${source}\nreturn {docText_, appendSafeTableCell_, mergeLeadingTableCells_};`,
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

// Model the Docs behavior where merged-over cells retain their indexes. A
// right-to-left merge must leave the first cell spanning columns 1–5 while the
// sixth (amount) cell remains independent.
const cells = Array.from({length: 6}, (_, index) => ({index, span: 1}));
const totalRow = {
  getCell(index) {
    const cell = cells[index];
    return {
      merge() {
        const previous = cells[index - 1];
        previous.span += cell.span;
        cell.span = 0;
        return previous;
      },
      getColSpan() {
        return cell.span;
      },
    };
  },
};

helpers.mergeLeadingTableCells_(totalRow, 5);
assert.deepEqual(cells.map(cell => cell.span), [5, 0, 0, 0, 0, 1]);

console.log('Document-generation formatting and merge safeguards passed.');
