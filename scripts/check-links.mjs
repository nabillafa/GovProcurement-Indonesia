import assert from 'node:assert/strict';
import fs from 'node:fs';

const source = fs.readFileSync(new URL('../PPBJ_Web.gs', import.meta.url), 'utf8');
const match = source.match(/function webLinkUrl_\([\s\S]*?\n}\n\nfunction webGetArchives/);
if (!match) throw new Error('webLinkUrl_ was not found');

const functionSource = match[0].replace(/\n\nfunction webGetArchives$/, '');
const webLinkUrl = new Function(`${functionSource}; return webLinkUrl_;`)();

const docsUrl = 'https://docs.google.com/document/d/example-document-id/edit';
const pdfUrl = 'https://drive.google.com/file/d/example-pdf-id/view';

assert.equal(webLinkUrl('Buka Google Docs', {getLinkUrl: () => docsUrl}, ''), docsUrl);
assert.equal(webLinkUrl('Buka PDF', {
  getLinkUrl: () => null,
  getRuns: () => [{getLinkUrl: () => null}, {getLinkUrl: () => pdfUrl}]
}, ''), pdfUrl);
assert.equal(webLinkUrl('Buka Google Docs', null, `=HYPERLINK("${docsUrl}","Buka Google Docs")`), docsUrl);
assert.equal(webLinkUrl(pdfUrl, null, ''), pdfUrl);
assert.equal(webLinkUrl('Buka Google Docs', null, ''), '');
assert.equal(webLinkUrl('https://example.com/file', null, ''), '');
assert.equal(webLinkUrl('https://docs.google.com.evil.example/file', null, ''), '');

console.log('Google Docs/Drive link checks passed.');
