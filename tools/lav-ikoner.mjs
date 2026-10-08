// Tegner knappens ikon (en lille kalender med et grønt felt) i de størrelser, Outlook vil have.
import { writeFileSync, mkdirSync } from 'node:fs';
import { deflateSync } from 'node:zlib';

const crcTab = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
const crc = (b) => { let c = 0xffffffff; for (const x of b) c = crcTab[(c ^ x) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
function chunk(type, data) {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const c = Buffer.alloc(4); c.writeUInt32BE(crc(td));
  return Buffer.concat([len, td, c]);
}
function png(n, px) {
  const hdr = Buffer.alloc(13); hdr.writeUInt32BE(n, 0); hdr.writeUInt32BE(n, 4); hdr[8] = 8; hdr[9] = 6;
  const raw = Buffer.alloc(n * (n * 4 + 1));
  for (let y = 0; y < n; y++) { raw[y * (n * 4 + 1)] = 0; px.copy(raw, y * (n * 4 + 1) + 1, y * n * 4, (y + 1) * n * 4); }
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', hdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

const MOERK = [59, 109, 17], LYS = [255, 255, 255], GROEN = [99, 153, 34], STREG = [180, 196, 160];
function farve(x, y) { // x,y i 0..1
  const inde = (x0, y0, x1, y1, r = 0) => {
    if (x < x0 || x > x1 || y < y0 || y > y1) return false;
    const cx = Math.min(Math.max(x, x0 + r), x1 - r), cy = Math.min(Math.max(y, y0 + r), y1 - r);
    return (x - cx) ** 2 + (y - cy) ** 2 <= r * r;
  };
  if (!inde(0.06, 0.1, 0.94, 0.94, 0.12)) return null;
  if (!inde(0.13, 0.32, 0.87, 0.87, 0.05)) return MOERK; // ramme + topbånd
  const cx = (x - 0.13) / 0.74, cy = (y - 0.32) / 0.55;
  const kol = Math.floor(cx * 3), raek = Math.floor(cy * 2);
  if (kol === 1 && raek === 1) return GROEN;
  const fx = cx * 3 - kol, fy = cy * 2 - raek;
  if ((fx < 0.06 && kol > 0) || (fy < 0.08 && raek > 0)) return STREG;
  return LYS;
}
mkdirSync(new URL('../ikoner/', import.meta.url), { recursive: true });
for (const n of [16, 32, 64, 80, 128]) {
  const px = Buffer.alloc(n * n * 4), SS = 4;
  for (let y = 0; y < n; y++) for (let x = 0; x < n; x++) {
    let r = 0, g = 0, b = 0, a = 0;
    for (let sy = 0; sy < SS; sy++) for (let sx = 0; sx < SS; sx++) {
      const f = farve((x + (sx + 0.5) / SS) / n, (y + (sy + 0.5) / SS) / n);
      if (f) { r += f[0]; g += f[1]; b += f[2]; a++; }
    }
    const i = (y * n + x) * 4;
    if (a) { px[i] = r / a; px[i + 1] = g / a; px[i + 2] = b / a; px[i + 3] = 255 * a / (SS * SS); }
  }
  writeFileSync(new URL(`../ikoner/ikon-${n}.png`, import.meta.url), png(n, px));
}
console.log('ikoner/ikon-{16,32,64,80,128}.png');
