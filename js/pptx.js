// Package OOXML en mémoire : parts (octets), XML parsés à la demande, relations, content types.
import { readZip, writeZip } from './zip.js';
import { t } from './i18n.js';

export const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  pr: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  asvg: 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
};
const enc = new TextEncoder(), dec = new TextDecoder();
export const isXml = n => /\.(xml|rels)$/i.test(n);
/** Éléments d'espace de noms ns et de nom local name sous el (descendants). */
export const all = (el, ns, name) => [...el.getElementsByTagNameNS(NS[ns] || ns, name)];
export const kids = (el, ns, name) => [...el.children].filter(c => c.namespaceURI === NS[ns] && (!name || c.localName === name));

/** Ouvre un .pptx (ArrayBuffer). */
export async function openPptx(buf) {
  const entries = await readZip(buf);
  const pkg = { entries: new Map(entries.map(e => [e.name, e])), docs: new Map() };
  if (!pkg.entries.has('[Content_Types].xml') || !pkg.entries.has('ppt/presentation.xml')) throw new Error(t('Ce fichier n’est pas une présentation PowerPoint (.pptx).'));
  return pkg;
}
export const names = pkg => [...pkg.entries.keys()];
export const has = (pkg, n) => pkg.entries.has(n);
export const bytes = (pkg, n) => pkg.entries.get(n)?.data;
export function setBytes(pkg, n, data) { pkg.docs.delete(n); const e = pkg.entries.get(n); e ? (e.data = data) : pkg.entries.set(n, { name: n, data }); }
export function remove(pkg, n) {
  if (!pkg.entries.delete(n)) return false;
  pkg.docs.delete(n);
  const ct = xml(pkg, '[Content_Types].xml');
  for (const o of all(ct, 'ct', 'Override')) if (o.getAttribute('PartName') === '/' + n) o.remove();
  remove(pkg, relsOf(n)); // ses propres relations
  return true;
}
/** XML d'une part (parsé une fois, marqué modifié : resérialisé à l'enregistrement). */
export function xml(pkg, n) {
  if (!pkg.docs.has(n)) {
    const d = pkg.entries.get(n) && new DOMParser().parseFromString(dec.decode(pkg.entries.get(n).data), 'application/xml');
    if (!d || d.getElementsByTagName('parsererror').length) return null;
    pkg.docs.set(n, d);
  }
  return pkg.docs.get(n);
}
export const text = (pkg, n) => pkg.docs.has(n) ? new XMLSerializer().serializeToString(pkg.docs.get(n)) : dec.decode(bytes(pkg, n));

/** Chemin des relations d'une part : ppt/slides/slide1.xml → ppt/slides/_rels/slide1.xml.rels */
export const relsOf = n => n ? n.replace(/(^|\/)([^/]+)$/, '$1_rels/$2.rels') : '_rels/.rels'; // '' = racine du package
/** Résout une cible relative depuis une part. */
export function resolve(from, target) {
  if (target.startsWith('/')) return target.slice(1);
  const p = from.split('/'); p.pop();
  for (const s of target.split('/')) s === '..' ? p.pop() : s !== '.' && p.push(s);
  return p.join('/');
}
/** Relations d'une part : [{ el, id, type, target (chemin résolu ou URL externe), external }] */
export function rels(pkg, n) {
  const d = xml(pkg, relsOf(n));
  if (!d) return [];
  return all(d, 'pr', 'Relationship').map(el => {
    const external = el.getAttribute('TargetMode') === 'External', raw = el.getAttribute('Target') || '';
    return { el, id: el.getAttribute('Id'), type: el.getAttribute('Type') || '', external, target: external ? raw : resolve(n, raw) };
  });
}
/** Toutes les parts .rels et la part source correspondante. */
export const relParts = pkg => names(pkg).filter(n => n.endsWith('.rels')).map(r => [r, r.replace(/(^|\/)_rels\/([^/]+)\.rels$/, '$1$2')]);

/** Slides dans l'ordre de la présentation. */
export function slides(pkg) {
  const pres = xml(pkg, 'ppt/presentation.xml'), byId = Object.fromEntries(rels(pkg, 'ppt/presentation.xml').map(r => [r.id, r.target]));
  return all(pres, 'p', 'sldId').map(s => byId[s.getAttributeNS(NS.r, 'id')]).filter(n => n && has(pkg, n));
}
/** Parts XML par motif. */
export const partsLike = (pkg, re) => names(pkg).filter(n => re.test(n)).sort((x, y) => x.localeCompare(y, 'en', { numeric: true }));

/** Déclare un type par défaut pour une extension (ex. png) s'il manque. */
export function ensureDefault(pkg, ext, type) {
  const ct = xml(pkg, '[Content_Types].xml'), root = ct.documentElement;
  if (all(ct, 'ct', 'Default').some(d => d.getAttribute('Extension').toLowerCase() === ext)) return;
  const d = ct.createElementNS(NS.ct, 'Default'); d.setAttribute('Extension', ext); d.setAttribute('ContentType', type);
  root.insertBefore(d, root.firstElementChild);
}

/** Références r:* vers une relation absente (PowerPoint demanderait une réparation) : [{ part, ids }]. */
export function dangling(pkg) {
  const out = [];
  for (const n of names(pkg).filter(x => x.endsWith('.xml') && !x.startsWith('['))) {
    const d = xml(pkg, n); if (!d) continue;
    const ids = new Set(rels(pkg, n).map(r => r.id)), bad = new Set();
    for (const el of d.getElementsByTagName('*')) for (const a of el.attributes) if (a.namespaceURI === NS.r && a.value && !ids.has(a.value)) bad.add(a.value);
    if (bad.size) out.push({ part: n, ids: [...bad] });
  }
  return out;
}

/** Enregistre : XML resérialisés, ordre des entrées conservé ([Content_Types].xml en premier). */
export async function savePptx(pkg) {
  for (const [n, d] of pkg.docs) {
    let s = new XMLSerializer().serializeToString(d);
    if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s;
    pkg.entries.get(n).data = enc.encode(s);
  }
  pkg.docs.clear();
  const list = [...pkg.entries.values()].sort((x, y) => (y.name === '[Content_Types].xml') - (x.name === '[Content_Types].xml'));
  return writeZip(list);
}
