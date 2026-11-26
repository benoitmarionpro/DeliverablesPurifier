// ZIP sans dépendance (le .pptx est un zip) : lecture (stocké / deflate) et écriture (deflate natif via CompressionStream).
// Les entrées non modifiées sont réécrites avec leurs octets compressés d'origine (aller-retour fidèle, rapide).
import { t } from './i18n.js';

const CRC = new Uint32Array(256).map((_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
export const crc32 = b => { let c = ~0; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return ~c >>> 0; };

async function pipe(stream, data) {
  const w = stream.writable.getWriter();
  w.write(data).catch(() => {}); w.close().catch(() => {}); // l'erreur éventuelle remonte par la lecture
  return new Uint8Array(await new Response(stream.readable).arrayBuffer());
}
const MAX = 1e9; // octets décompressés au total (bombe zip)

/** Lit un zip : [{ name, data: Uint8Array, orig: { method, raw, crc } }], dans l'ordre du répertoire central. Ni Zip64 ni chiffrement. */
export async function readZip(buf) {
  const b = new Uint8Array(buf), v = new DataView(b.buffer, b.byteOffset, b.byteLength), dec = new TextDecoder();
  let e = b.length - 22;
  while (e >= 0 && v.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw new Error(t('Fichier illisible : ce n’est pas un .pptx valide.'));
  const count = v.getUint16(e + 10, true), cd = v.getUint32(e + 16, true);
  if (count === 0xFFFF || cd === 0xFFFFFFFF) throw new Error(t('Fichier trop volumineux (Zip64 non pris en charge).'));
  const out = [];
  let total = 0;
  for (let i = 0, p = cd; i < count; i++) {
    if (v.getUint32(p, true) !== 0x02014b50) throw new Error(t('Fichier illisible : ce n’est pas un .pptx valide.'));
    const flags = v.getUint16(p + 8, true), method = v.getUint16(p + 10, true), crc = v.getUint32(p + 16, true), size = v.getUint32(p + 20, true), usize = v.getUint32(p + 24, true);
    const nl = v.getUint16(p + 28, true), name = dec.decode(b.subarray(p + 46, p + 46 + nl)), lh = v.getUint32(p + 42, true);
    p += 46 + nl + v.getUint16(p + 30, true) + v.getUint16(p + 32, true);
    if (name.endsWith('/')) continue;
    if (flags & 1) throw new Error(t('Fichier chiffré ou protégé par mot de passe : retire la protection avant de l’anonymiser.'));
    if (method !== 0 && method !== 8) throw new Error(t('Compression non prise en charge ({m}) pour {n}.', { m: method, n: name }));
    if ((total += usize) > MAX) throw new Error(t('Fichier trop volumineux une fois décompressé (> 1 Go).'));
    const start = lh + 30 + v.getUint16(lh + 26, true) + v.getUint16(lh + 28, true), raw = b.subarray(start, start + size);
    const data = method ? await pipe(new DecompressionStream('deflate-raw'), raw) : raw;
    out.push({ name, data, orig: { method, raw, crc, data } });
  }
  return out;
}

const STORE = /\.(png|jpe?g|gif|mp4|m4v|mov|mp3|m4a|wmv|avi|webm|zip|xlsx|docx|pptx|wdp)$/i; // déjà compressés

/** Écrit un zip à partir de [{ name, data }] ; une entrée dont `data` est celle lue garde ses octets d'origine. */
export async function writeZip(entries) {
  const enc = new TextEncoder(), parts = [], central = [];
  let off = 0;
  for (const { name, data, orig } of entries) {
    let method, raw, crc;
    if (orig && orig.data === data) ({ method, raw, crc } = orig);
    else {
      crc = crc32(data);
      method = STORE.test(name) || data.length < 64 ? 0 : 8;
      raw = method ? await pipe(new CompressionStream('deflate-raw'), data) : data;
    }
    const n = enc.encode(name), h = new Uint8Array(30 + n.length), c = new Uint8Array(46 + n.length), hv = new DataView(h.buffer), cv = new DataView(c.buffer);
    [[0, 0x04034b50, 4], [4, 20, 2], [6, 0x0800, 2], [8, method, 2], [14, crc, 4], [18, raw.length, 4], [22, data.length, 4], [26, n.length, 2]]
      .forEach(([o, x, s]) => s === 4 ? hv.setUint32(o, x, true) : hv.setUint16(o, x, true));
    [[0, 0x02014b50, 4], [4, 20, 2], [6, 20, 2], [8, 0x0800, 2], [10, method, 2], [16, crc, 4], [20, raw.length, 4], [24, data.length, 4], [28, n.length, 2], [42, off, 4]]
      .forEach(([o, x, s]) => s === 4 ? cv.setUint32(o, x, true) : cv.setUint16(o, x, true));
    h.set(n, 30); c.set(n, 46);
    parts.push(h, raw); central.push(c);
    off += h.length + raw.length;
    if (off > 0xFFFFFFFF) throw new Error(t('Fichier trop volumineux (Zip64 non pris en charge).'));
  }
  const size = central.reduce((s, c) => s + c.length, 0), e = new Uint8Array(22), ev = new DataView(e.buffer);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, entries.length, true); ev.setUint16(10, entries.length, true);
  ev.setUint32(12, size, true); ev.setUint32(16, off, true);
  return new Blob([...parts, ...central, e], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
}
