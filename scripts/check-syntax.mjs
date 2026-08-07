import fs from 'node:fs';

const serverFiles = ['PPBJ_Automation.gs', 'PPBJ_Web.gs'];
const serverSources = serverFiles.map(file => ({
  file,
  source: fs.readFileSync(new URL('../' + file, import.meta.url), 'utf8'),
}));

for (const { file, source } of serverSources) {
  new Function(source);
  console.log(`Syntax OK: ${file}`);
}

// Apps Script evaluates every .gs file in one global namespace. Parsing the
// concatenated sources catches duplicate top-level declarations that per-file
// checks miss (for example: `Identifier 'CFG' has already been declared`).
new Function(serverSources.map(({ file, source }) =>
  `\n/* ===== ${file} ===== */\n${source}`
).join('\n'));
console.log('Syntax OK: combined Apps Script global scope');

const html = fs.readFileSync(new URL('../Index.html', import.meta.url), 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/i);
if (!script) throw new Error('Inline script was not found in Index.html');
new Function(script[1]);
console.log('Syntax OK: Index.html inline JavaScript');
