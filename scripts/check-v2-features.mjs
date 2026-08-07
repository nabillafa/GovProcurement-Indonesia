import fs from 'node:fs';

const automation = fs.readFileSync(new URL('../PPBJ_Automation.gs', import.meta.url), 'utf8');
const web = fs.readFileSync(new URL('../PPBJ_Web.gs', import.meta.url), 'utf8');
const html = fs.readFileSync(new URL('../Index.html', import.meta.url), 'utf8');
const manifest = JSON.parse(fs.readFileSync(new URL('../appsscript.json', import.meta.url), 'utf8'));

const assertions = [
  ['Gmail OTP request', /function webStartLogin\(/.test(web)],
  ['Gmail OTP verification', /function webVerifyLogin\(/.test(web)],
  ['temporary user key session binding', /Session\.getTemporaryActiveUserKey\(\)/.test(web)],
  ['spreadsheet is not shared with users', !/getFileById\(CFG\.SPREADSHEET_ID\)\.addEditor/.test(web)],
  ['role-based package access', /function webAssertPackageAccess_\(/.test(web)],
  ['calendar-aware schedule', /function webGenerateLetterSchedule\(/.test(web) && /rollToWorkday_/.test(web)],
  ['invoice included in schedule', /'Invoice'/.test(web)],
  ['Excel budget upload', /function webUploadBudgetExcel\(/.test(web)],
  ['budget preview before activation', /function webPreviewBudgetImport\(/.test(web)],
  ['paste-from-Excel UI', /function pasteMatrix\(/.test(html)],
  ['tax and package-cost UI', /function calcCosts\(/.test(html)],
  ['author footer', /Dibuat oleh/.test(html)],
  ['Drive advanced service', manifest.dependencies?.enabledAdvancedServices?.some(x => x.serviceId === 'drive')],
  ['mail send scope', manifest.oauthScopes?.includes('https://www.googleapis.com/auth/script.send_mail')],
  ['Indonesian month formatting', /function monthName_\(/.test(automation) && /function formatDate_\(/.test(automation)]
];

const failed = assertions.filter(([, ok]) => !ok);
if (failed.length) {
  for (const [label] of failed) console.error('Missing safeguard/feature:', label);
  process.exit(1);
}
console.log('V2 access, scheduling, calculator, import, and UI checks passed.');
