// Mise en page (portage de LayoutAnonymizer) : masters/layouts purgés, fonds blancs, fonds sombres neutralisés, bordures noires, texte Aptos noir / titres violet Wavestone.
import { NS, all, kids, xml, slides, partsLike } from './pptx.js';

const WHITE = 'FFFFFF', BLACK = '000000', GREY = 'E6E6E6', WAVESTONE = '451DC7';
const el = (doc, name, attrs = {}) => { const e = doc.createElementNS(NS.a, 'a:' + name); for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v); return e; };
const solid = (doc, hex) => { const s = el(doc, 'solidFill'); s.append(el(doc, 'srgbClr', { val: hex })); return s; };
const FILLS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];

/** Fond blanc : p:bg remplacé, en tête de p:cSld. */
function whiteBackground(doc) {
  const cSld = all(doc, 'p', 'cSld')[0]; if (!cSld) return;
  kids(cSld, 'p', 'bg').forEach(b => b.remove());
  const bg = doc.createElementNS(NS.p, 'p:bg'), pr = doc.createElementNS(NS.p, 'p:bgPr');
  pr.append(solid(doc, WHITE), el(doc, 'effectLst')); bg.append(pr);
  cSld.insertBefore(bg, cSld.firstElementChild);
}
/** Placeholder ? (p:nvPr/p:ph dans les propriétés non visuelles de la forme). */
const ph = shape => { const nv = [...shape.getElementsByTagNameNS(NS.p, 'nvPr')][0]; return nv ? kids(nv, 'p', 'ph')[0] : null; };
/** Supprime toutes les formes hors placeholders (logos, images de fond, formes décoratives). */
function stripShapes(doc) {
  const tree = all(doc, 'p', 'spTree')[0]; if (!tree) return;
  for (const s of [...tree.children]) if (!/^(nvGrpSpPr|grpSpPr|extLst)$/.test(s.localName) && !ph(s)) s.remove();
}
const dark = hex => { const n = parseInt(hex, 16); return 0.299 * (n >> 16 & 255) + 0.587 * (n >> 8 & 255) + 0.114 * (n & 255) < 128; };
/** Remplissage sombre, de thème (non résolu) ou complexe (dégradé, motif, image) → gris clair. Remplissage hérité : ignoré. */
function lighten(pr) {
  const f = [...pr.children].find(c => c.namespaceURI === NS.a && FILLS.includes(c.localName));
  if (!f || f.localName === 'noFill' || f.localName === 'grpFill') return;
  if (f.localName === 'solidFill') { const c = f.firstElementChild; if (c?.localName === 'srgbClr' && !dark(c.getAttribute('val') || '000000')) return; }
  pr.replaceChild(solid(pr.ownerDocument, GREY), f);
}
/** Bordure avec couleur explicite → noire. */
function blackBorder(pr) {
  const ln = kids(pr, 'a', 'ln')[0], s = ln && kids(ln, 'a', 'solidFill')[0];
  if (s?.firstElementChild) s.replaceChildren(el(pr.ownerDocument, 'srgbClr', { val: BLACK }));
}
// ordre des enfants de a:rPr imposé par le schéma (sinon PowerPoint demande une réparation)
const RPR = ['ln', ...FILLS, 'effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill', 'latin', 'ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'];
function place(pr, node) {
  const i = RPR.indexOf(node.localName), after = [...pr.children].find(c => RPR.indexOf(c.localName) > i);
  pr.insertBefore(node, after || null);
}
function styleRun(r, font, color) {
  const doc = r.ownerDocument;
  let pr = kids(r, 'a', 'rPr')[0];
  if (!pr) { pr = el(doc, 'rPr'); r.insertBefore(pr, r.firstElementChild); }
  [...pr.children].filter(c => FILLS.includes(c.localName) || c.localName === 'latin').forEach(c => c.remove());
  place(pr, solid(doc, color)); place(pr, el(doc, 'latin', { typeface: font }));
}

export function anonymizeLayout(pkg) {
  for (const p of partsLike(pkg, /^ppt\/(slideMasters|slideLayouts)\/[^/]+\.xml$/)) { const d = xml(pkg, p); if (d) { whiteBackground(d); stripShapes(d); } }
  for (const s of slides(pkg)) {
    const d = xml(pkg, s); if (!d) continue;
    whiteBackground(d);
    for (const sp of all(d, 'p', 'spPr')) { if (sp.parentElement.localName === 'sp') lighten(sp); blackBorder(sp); }
    for (const tc of all(d, 'a', 'tcPr')) lighten(tc);
    for (const shape of all(d, 'p', 'sp')) {
      const title = /^(title|ctrTitle)$/.test(ph(shape)?.getAttribute('type') || '');
      for (const r of all(shape, 'a', 'r')) styleRun(r, title ? 'Aptos SemiBold' : 'Aptos', title ? WAVESTONE : BLACK);
    }
    for (const tbl of all(d, 'a', 'tbl')) for (const r of all(tbl, 'a', 'r')) styleRun(r, 'Aptos', BLACK);
  }
}
