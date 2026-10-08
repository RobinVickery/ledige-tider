// Laver manifest.xml ud fra skabelonen med den adresse, siden ligger på.
//   node tools/lav-manifest.mjs https://robinvickery.github.io/ledige-tider
import { readFileSync, writeFileSync } from 'node:fs';
const vaert = (process.argv[2] || '').replace(/\/+$/, '');
if (!/^https:\/\//.test(vaert)) { console.error('Brug: node tools/lav-manifest.mjs https://din-adresse'); process.exit(1); }
const ud = readFileSync(new URL('../manifest.skabelon.xml', import.meta.url), 'utf8')
  .replace(/^<!-- Skabelon.*-->\n/m, '')
  .replaceAll('{VAERT}', vaert);
writeFileSync(new URL('../manifest.xml', import.meta.url), ud);
console.log('manifest.xml peger nu på ' + vaert);
