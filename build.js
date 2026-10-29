// Builds the app by inlining the libraries, the Coolabah deal and its term sheet into src/app.html.
//   dist/artifact.html : page body only (for publishing as a claude.ai artifact)
//   index.html         : standalone page you can open in a browser or host on GitHub Pages
const fs = require('fs');
const path = require('path');
const dir = __dirname;
const read = (f) => fs.readFileSync(path.join(dir, f), 'utf8');
// Order matters: each library uses the ones before it.
const LIBS = ['engine', 'generator', 'termsheet', 'markdown', 'parser', 'xlsx', 'docx', 'excel'].map((f) => read('src/' + f + '.js')).join('\n');
const deal = read('deal.json').trim();
const termSheet = JSON.stringify(read('docs/coolabah_term_sheet.md'));
const body = read('src/app.html')
  .replace('/*__LIBS__*/', () => LIBS)
  .replace('/*__DEAL__*/', () => deal)
  .replace('/*__COOLABAH_TS__*/', () => termSheet);
fs.mkdirSync(path.join(dir, 'dist'), { recursive: true });
fs.writeFileSync(path.join(dir, 'dist/artifact.html'), body);
fs.writeFileSync(path.join(dir, 'index.html'),
  '<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n</head>\n<body>\n' + body + '\n</body>\n</html>\n');
console.log('Built index.html and dist/artifact.html');
