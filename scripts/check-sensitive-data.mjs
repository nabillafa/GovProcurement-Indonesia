import fs from 'node:fs';

const files = ['PPBJ_Automation.gs', 'PPBJ_Web.gs', 'Index.html'];
const rules = [
  ['Google asset/deployment URL', /https:\/\/(?:docs|drive|script)\.google\.com\/(?:[^\s'"<]+\/)?(?:d|s)\/[A-Za-z0-9_-]{20,}/gi],
  ['email address outside the example domain', /\b[A-Z0-9._%+-]+@(?!example\.(?:com|org)\b)[A-Z0-9.-]+\.[A-Z]{2,}\b/gi],
  ['18-digit employee identifier', /\b\d{8}\s?\d{6}\s?\d\s?\d{3}\b/g],
  ['possible Google asset ID literal', /['"][A-Za-z0-9_-]{30,}['"]/g],
  ['organization-specific marker', /BPS Kabupaten Kaur|Peltu M\. Ilyas|Padang Kempas|Susenas|Vseruti|Amerizasni|Mirdiana|Percetakan Fazzy/gi]
];

let failed = false;
for (const file of files) {
  const source = fs.readFileSync(new URL('../' + file, import.meta.url), 'utf8');
  for (const [label, pattern] of rules) {
    pattern.lastIndex = 0;
    const matches = source.match(pattern) || [];
    if (matches.length) {
      failed = true;
      console.error(`${file}: ${label}: ${[...new Set(matches)].join(', ')}`);
    }
  }
}

if (failed) process.exit(1);
console.log('Sensitive-data check passed.');
