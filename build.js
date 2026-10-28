// Builds the app by inlining the engine and the deal into src/app.html.
//   dist/artifact.html : page body only (for publishing as a claude.ai artifact)
//   index.html         : standalone page you can open in a browser or host on GitHub Pages
const fs = require('fs');
const path = require('path');
const dir = __dirname;
const tpl = fs.readFileSync(path.join(dir, 'src/app.html'), 'utf8');
const engine = fs.readFileSync(path.join(dir, 'src/engine.js'), 'utf8');
const deal = fs.readFileSync(path.join(dir, 'deal.json'), 'utf8').trim();
const body = tpl.replace('/*__ENGINE__*/', () => engine).replace('/*__DEAL__*/', () => deal);
fs.mkdirSync(path.join(dir, 'dist'), { recursive: true });
fs.writeFileSync(path.join(dir, 'dist/artifact.html'), body);
fs.writeFileSync(path.join(dir, 'index.html'),
  '<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n</head>\n<body>\n' + body + '\n</body>\n</html>\n');
console.log('Built index.html and dist/artifact.html');
