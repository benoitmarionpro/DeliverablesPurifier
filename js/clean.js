// Nettoyages déterministes : notes, commentaires, vidéos/audio, hyperliens, métadonnées, objets embarqués, parts orphelines.
import { NS, all, xml, rels, relParts, partsLike, names, remove, has, isXml } from './pptx.js';

/** Notes du présentateur vidées (parité avec l'app d'origine). Renvoie le nombre de notes non vides. */
export function clearNotes(pkg) {
  let n = 0;
  for (const p of partsLike(pkg, /^ppt\/notesSlides\/[^/]+\.xml$/)) {
    const nodes = all(xml(pkg, p), 'a', 't');
    if (nodes.some(x => x.textContent.trim())) n++;
    nodes.forEach(x => (x.textContent = ''));
  }
  return n;
}

const COMMENT = /(comments|commentAuthors|\/authors)$/i;
/** Commentaires (anciens et modernes) et leurs auteurs. Renvoie le nombre de parts de commentaires supprimées. */
export function removeComments(pkg) {
  const n = partsLike(pkg, /^ppt\/comments\//).length;
  for (const [, src] of relParts(pkg)) for (const r of rels(pkg, src)) if (COMMENT.test(r.type)) {
    r.el.remove(); if (!r.external) remove(pkg, r.target);
    // commentaires modernes : <p:ext><p188:commentRel r:id/></p:ext> dans la slide, retiré avec la relation
    const d = has(pkg, src) && isXml(src) && xml(pkg, src);
    if (d) for (const c of [...d.getElementsByTagName('*')].filter(e => e.localName === 'commentRel' && e.getAttributeNS(NS.r, 'id') === r.id)) {
      const ext = c.parentElement?.localName === 'ext' ? c.parentElement : c, lst = ext.parentElement;
      ext.remove(); if (lst?.localName === 'extLst' && !lst.children.length) lst.remove();
    }
  }
  for (const p of partsLike(pkg, /^ppt\/(comments\/|commentAuthors\.xml$|authors\.xml$)/)) remove(pkg, p);
  return n;
}

const AV = /\.(mp4|mpeg|mpg|avi|mov|wmv|mkv|webm|m4v|3gp|mp3|wav|m4a|wma|aac|ogg|mid|midi)$/i;
const AV_REL = /\/(video|audio)$|2007\/relationships\/media$/;
/** Vidéos et sons : médias, relations et formes qui les portent. Renvoie le nombre de vidéos/sons retirés. */
export function removeMedia(pkg) {
  let n = 0;
  for (const s of partsLike(pkg, /^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/)) {
    const av = rels(pkg, s).filter(r => AV_REL.test(r.type) || AV.test(r.target));
    if (!av.length) continue;
    const ids = new Set(av.map(r => r.id)), doc = xml(pkg, s);
    // forme portant le média (p:pic avec a:videoFile / a:audioFile / p14:media) et ses minutages
    for (const el of [...doc.getElementsByTagName('*')]) {
      if (!/^(videoFile|audioFile|quickTimeFile|media)$/.test(el.localName)) continue;
      const id = el.getAttributeNS(NS.r, 'link') || el.getAttributeNS(NS.r, 'embed');
      if (!ids.has(id)) continue;
      let shape = el; while (shape && shape.localName !== 'pic') shape = shape.parentElement;
      const spid = shape && all(shape, 'p', 'cNvPr')[0]?.getAttribute('id');
      // un mc:AlternateContent englobant part avec la forme
      let top = shape; while (top?.parentElement && top.parentElement.localName !== 'spTree' && top.parentElement.localName !== 'grpSp') top = top.parentElement;
      top?.remove();
      if (spid) for (const tgt of all(doc, 'p', 'spTgt')) if (tgt.getAttribute('spid') === spid) { let par = tgt; while (par && par.localName !== 'par') par = par.parentElement; par?.remove(); }
    }
    for (const r of av) { r.el.remove(); if (!r.external && AV.test(r.target) && remove(pkg, r.target)) n++; }
  }
  for (const p of names(pkg)) if (p.startsWith('ppt/media/') && AV.test(p) && remove(pkg, p)) n++;
  return n;
}

/** Hyperliens (toutes les parts) : lien retiré, texte du lien préfixé « Lien_URL_ » (parité). Renvoie le nombre de liens. */
export function removeLinks(pkg) {
  let n = 0;
  for (const p of names(pkg).filter(x => isXml(x) && !x.endsWith('.rels') && x.startsWith('ppt/'))) {
    const doc = xml(pkg, p); if (!doc) continue;
    const hl = [...all(doc, 'a', 'hlinkClick'), ...all(doc, 'a', 'hlinkMouseOver')];
    if (!hl.length) continue;
    const hrel = new Set(rels(pkg, p).filter(r => r.type.endsWith('/hyperlink')).map(r => r.id));
    for (const h of hl) {
      const id = h.getAttributeNS(NS.r, 'id');
      if (!hrel.has(id) && !(h.getAttribute('action') || '').startsWith('ppaction://hlink')) continue; // actions internes sans URL : conservées
      const run = h.parentElement?.localName === 'rPr' ? h.parentElement.parentElement : null;
      const tn = run && all(run, 'a', 't')[0];
      if (tn && !tn.textContent.startsWith('Lien_URL_') && h.localName === 'hlinkClick') tn.textContent = 'Lien_URL_' + tn.textContent;
      h.remove(); n++;
    }
  }
  for (const [, src] of relParts(pkg)) for (const r of rels(pkg, src)) if (r.type.endsWith('/hyperlink')) r.el.remove();
  return n;
}

/** Métadonnées et objets embarqués : docProps vidés, propriétés personnalisées / customXml / miniature supprimées, classeurs des graphiques retirés. */
export function scrubMeta(pkg) {
  const blank = (p, tags) => { const d = xml(pkg, p); if (d) for (const el of d.getElementsByTagName('*')) if (tags.test(el.localName) && !el.children.length) el.textContent = ''; };
  blank('docProps/core.xml', /^(creator|lastModifiedBy|title|subject|keywords|description|category|contentStatus|identifier|lastPrinted)$/);
  blank('docProps/app.xml', /^(Company|Manager|HyperlinkBase)$/);
  for (const r of rels(pkg, '')) if (/\/(thumbnail|custom-properties)$/.test(r.type)) { r.el.remove(); remove(pkg, r.target); }
  for (const p of ['docProps/thumbnail.jpeg', 'docProps/custom.xml']) remove(pkg, p);
  for (const [, src] of relParts(pkg)) for (const r of rels(pkg, src)) if (/\/customXml$/.test(r.type)) {
    r.el.remove(); remove(pkg, r.target);
    // p:custData (métadonnées de template, ex. Templafy) pointant vers la part supprimée : sinon PowerPoint demande une réparation
    const d = has(pkg, src) && isXml(src) && xml(pkg, src);
    if (d) for (const c of all(d, 'p', 'custData')) if (c.getAttributeNS(NS.r, 'id') === r.id) { const lst = c.parentElement; c.remove(); if (!lst.children.length) lst.remove(); }
  }
  for (const p of partsLike(pkg, /^customXml\//)) remove(pkg, p);
  let embeds = 0;
  for (const ch of partsLike(pkg, /^ppt\/charts\/[^/]+\.xml$/)) {
    const doc = xml(pkg, ch); if (!doc) continue;
    for (const e of all(doc, 'c', 'externalData')) {
      const r = rels(pkg, ch).find(x => x.id === e.getAttributeNS(NS.r, 'id'));
      e.remove(); if (r) { r.el.remove(); if (!r.external && remove(pkg, r.target)) embeds++; }
    }
    // feuilles de style / couleurs du graphique : sans données
  }
  const ole = partsLike(pkg, /^ppt\/embeddings\//).length; // objets OLE restants (Excel, Word… insérés sur une slide) : signalés, pas supprimés
  return { embeds, ole };
}

const MEDIA_REL = /\/(image|video|audio|hdphoto)$|2007\/relationships\/(media|hdphoto)$/;
/** Relations média non référencées par leur part, puis médias plus référencés par aucune relation : supprimés (évite qu'un original subsiste dans le zip). */
export function prune(pkg) {
  let n = 0;
  for (const [, src] of relParts(pkg)) {
    const doc = has(pkg, src) && isXml(src) && xml(pkg, src);
    if (!doc) continue;
    const body = new XMLSerializer().serializeToString(doc);
    for (const r of rels(pkg, src)) if (MEDIA_REL.test(r.type) && !new RegExp(`"${r.id}"`).test(body)) r.el.remove();
  }
  const used = new Set(relParts(pkg).flatMap(([, src]) => rels(pkg, src).filter(r => !r.external).map(r => r.target)));
  for (const p of names(pkg)) if (p.startsWith('ppt/media/') && !used.has(p) && remove(pkg, p)) n++;
  return n;
}
