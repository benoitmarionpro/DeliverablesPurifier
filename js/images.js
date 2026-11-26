// Images : inventaire (ordre des slides, icônes exclues), classification par LLM vision (repli : revue manuelle), floutage / carré gris sur canvas.
import { NS, all, xml, rels, slides, partsLike, bytes, setBytes, remove, has, relParts, ensureDefault } from './pptx.js';
import { chatJson, UNTRUSTED } from './llm.js';
import { t } from './i18n.js';

export const TYPES = ['logo', 'diagram', 'confidential', 'none'];
export const LOGO_LABELS = ['Luxury_Client', 'Bank_Client', 'Insurance_Client', 'Retail_Client', 'Energy_Client', 'Tech_Client', 'Healthcare_Client', 'Public_Sector_Client', 'Partner', 'Client'];
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf' };
const VIEWABLE = /^(png|jpe?g|gif|bmp|webp|svg)$/; // décodables par le navigateur (EMF, WMF, TIFF : non)
const ICON = /icon|icône|icone/i;
const ext = p => p.split('.').pop().toLowerCase();

/** Nom de la forme qui porte l'élément (cNvPr de p:pic / p:sp / p:grpSp ancêtre) : sert à exclure les icônes PowerPoint. */
function shapeNames(el) {
  const out = [];
  for (let e = el.parentElement; e; e = e.parentElement) {
    if (!/^(pic|sp|grpSp|graphicFrame)$/.test(e.localName)) continue;
    const nv = [...e.children].find(c => c.localName.startsWith('nv'));
    const c = nv && [...nv.children].find(x => x.localName === 'cNvPr');
    if (c?.getAttribute('name')) out.push(c.getAttribute('name'));
  }
  return out;
}

/** Images référencées, dans l'ordre d'apparition : [{ path, ext, viewable, slides[], svg[] }]. Avec withMasters, aussi celles des masters/layouts. */
export function inventory(pkg, { withMasters }) {
  const order = slides(pkg), parts = [...order, ...(withMasters ? partsLike(pkg, /^ppt\/(slideLayouts|slideMasters)\/[^/]+\.xml$/) : [])];
  const items = new Map();
  parts.forEach((part, i) => {
    const doc = xml(pkg, part); if (!doc) return;
    const byId = Object.fromEntries(rels(pkg, part).filter(r => !r.external).map(r => [r.id, r.target]));
    for (const blip of all(doc, 'a', 'blip')) {
      const path = byId[blip.getAttributeNS(NS.r, 'embed')];
      if (!path || !has(pkg, path) || !MIME[ext(path)] || shapeNames(blip).some(n => ICON.test(n))) continue;
      const it = items.get(path) || { path, ext: ext(path), viewable: VIEWABLE.test(ext(path)), slides: [], svg: [] };
      const where = i < order.length ? i + 1 : 'M';
      if (!it.slides.includes(where)) it.slides.push(where);
      for (const s of all(blip, 'asvg', 'svgBlip')) { const sp = byId[s.getAttributeNS(NS.r, 'embed')]; if (sp && !it.svg.includes(sp)) it.svg.push(sp); }
      items.set(path, it);
    }
  });
  // une image SVG d'un couple SVG + PNG de repli n'est pas listée deux fois
  const svgs = new Set([...items.values()].flatMap(i => i.svg));
  return [...items.values()].filter(i => !svgs.has(i.path)).map((it, index) => ({ ...it, index }));
}

export const blobOf = (pkg, it) => new Blob([bytes(pkg, it.path)], { type: MIME[it.ext] });
function load(blob) {
  return new Promise((ok, ko) => {
    const u = URL.createObjectURL(blob), img = new Image();
    img.onload = () => { URL.revokeObjectURL(u); ok(img); };
    img.onerror = () => { URL.revokeObjectURL(u); ko(new Error(t('Image illisible'))); };
    img.src = u;
  });
}
const size = img => [img.naturalWidth || 512, img.naturalHeight || 512];

/* ---------- Classification (prompt de l'app d'origine) ---------- */
const PROMPT = `Évalue si cette image contient des données sensibles ou confidentielles à anonymiser.
Classe chaque image détectée comme sensible dans l'un des types suivants :

**TYPE "logo"** — Logos ou marques identifiant un client ou partenaire :
- Logos ou marques de **clients** (banque, luxe, industrie, etc.) ou de **partenaires** dans un contexte commercial
- Tout logo permettant d'identifier un client, un partenaire ou une organisation spécifique
- ⚠️ NE PAS classer ici les logos d'éditeurs logiciels (Microsoft, Google, SAP, Oracle, AWS, Salesforce, Adobe, etc.) utilisés en contexte technique — ceux-ci sont NON SENSIBLES.

**TYPE "diagram"** — Schémas d'architecture, diagrammes de séquence, diagrammes de flux :
- Schémas d'architecture technique (réseau, infra, applicatif)
- Diagrammes de séquence, diagrammes UML, organigrammes techniques
- Flowcharts, diagrammes de flux de données
- Toute représentation visuelle structurée montrant des processus, des connexions ou une architecture
- ⚠️ Classer ici MÊME si le schéma ne contient pas de données confidentielles identifiables (IP, noms de serveurs, etc.) — la structure même doit être protégée.

**TYPE "confidential"** — Autres contenus sensibles :
- Photos de personnes identifiables (visages, équipes, portraits)
- Captures d'écran contenant des données confidentielles : adresses IP, noms de serveurs, URLs internes, noms de projets, données métier
- Filigranes, mentions "Confidentiel" / "Interne"
- Tout autre élément sensible ne rentrant pas dans les catégories "logo" ou "diagram"

**NON SENSIBLE (isSensitive: false)** :
- Marques publiques utilisées comme composants techniques ou logiciels : logos de solutions (Microsoft, Google, SAP, Oracle, AWS, Salesforce, Adobe, etc.) en contexte technique
- Graphiques génériques de données (barres, camemberts, architectures génériques) SANS données financières ni métier identifiables
- Icônes, illustrations décoratives sans lien avec un client
- Images de stock sans lien avec un acteur précis

Retourne UNIQUEMENT un JSON valide avec cette structure exacte :
{
  "isSensitive": true ou false,
  "sensitivityType": "logo" | "diagram" | "confidential" | "none",
  "logoLabel": "Luxury_Client",
  "reason": "Brève raison en une phrase"
}

Règles pour logoLabel (UNIQUEMENT si sensitivityType == "logo") :
- Utilise un label sémantique décrivant le secteur du logo détecté.
- Labels autorisés : Luxury_Client, Bank_Client, Insurance_Client, Retail_Client, Energy_Client, Tech_Client, Healthcare_Client, Public_Sector_Client, Partner, Client
- Si le type n'est pas "logo", mets logoLabel à null.`;

/** Réponse vision bornée aux valeurs connues. */
export function validateVision(r) {
  let type = TYPES.includes(r?.sensitivityType) ? r.sensitivityType : 'none';
  const sensitive = r?.isSensitive === true || (type !== 'none' && r?.isSensitive !== false);
  if (sensitive && type === 'none') type = 'confidential';
  return { sensitive, type: sensitive ? type : 'none', label: type === 'logo' ? (LOGO_LABELS.includes(r?.logoLabel) ? r.logoLabel : 'Client') : null, reason: String(r?.reason ?? '').slice(0, 300) };
}
/** Image réduite (1024 px max, JPEG) pour l'envoi au LLM. */
async function thumbUrl(blob, max = 1024) {
  const img = await load(blob), [w, h] = size(img), k = Math.min(1, max / Math.max(w, h));
  const c = Object.assign(document.createElement('canvas'), { width: Math.max(1, Math.round(w * k)), height: Math.max(1, Math.round(h * k)) }), x = c.getContext('2d');
  x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(img, 0, 0, c.width, c.height);
  return c.toDataURL('image/jpeg', 0.85);
}
const noVision = e => [400, 404, 415, 422, 500].includes(e.status) && /image|vision|multi-?modal|content.?type|image_url|not support|unsupported/i.test(e.message);

/** Classe les images (4 en parallèle). Renvoie { vision: false } si le modèle n'accepte pas les images : revue manuelle. */
export async function analyze(pkg, items, { signal, say }) {
  const todo = items.filter(i => i.viewable);
  let done = 0, vision = true, next = 0;
  const worker = async () => {
    while (vision && next < todo.length) {
      const it = todo[next++];
      try {
        const r = await chatJson({ system: UNTRUSTED, user: PROMPT, images: [await thumbUrl(blobOf(pkg, it))], signal });
        it.ai = validateVision(r);
      } catch (e) {
        if (e.name === 'AbortError') throw e;
        if (noVision(e)) { vision = false; return; }
        it.ai = { sensitive: true, type: 'confidential', label: null, reason: t('Analyse impossible, à revoir : {e}', { e: e.message.slice(0, 160) }), error: true };
      }
      say?.(t('Analyse des images : {i}/{n}', { i: ++done, n: todo.length }));
    }
  };
  await Promise.all(Array.from({ length: Math.min(4, todo.length) }, worker));
  if (!vision) for (const it of items) delete it.ai;
  return { vision };
}

/* ---------- Caviardage ---------- */
let filterOk;
const canFilter = x => filterOk ??= (x.filter = 'blur(2px)', x.filter === 'blur(2px)');
/** Rendu caviardé : 'logo' → carré gris + label ; sinon flou (repli : pixellisation). img null = format non décodable → carré gris. */
export function redactCanvas(img, kind, label, radius) {
  const [w, h] = img ? size(img) : [400, 300];
  const c = Object.assign(document.createElement('canvas'), { width: w, height: h }), x = c.getContext('2d');
  if (kind === 'logo' || !img) {
    x.fillStyle = 'rgb(180,180,180)'; x.fillRect(0, 0, w, h);
    const fs = Math.max(14, Math.floor(Math.min(w, h) / 12)), txt = label || t('Image masquée');
    x.font = `600 ${fs}px Inter, Arial, sans-serif`; x.fillStyle = 'rgb(60,60,60)'; x.textAlign = 'center'; x.textBaseline = 'middle';
    x.fillText(txt, w / 2, h / 2, w * 0.9);
    return c;
  }
  const r = Math.max(2, radius);
  x.fillStyle = 'rgb(128,128,128)'; x.fillRect(0, 0, w, h); // bords du flou : gris, pas l'original
  if (canFilter(x)) { x.filter = `blur(${r}px)`; x.drawImage(img, 0, 0, w, h); x.filter = 'none'; }
  else {
    const k = Math.max(2, r), sw = Math.max(1, Math.round(w / k)), sh = Math.max(1, Math.round(h / k));
    const s = Object.assign(document.createElement('canvas'), { width: sw, height: sh });
    s.getContext('2d').drawImage(img, 0, 0, sw, sh);
    x.imageSmoothingEnabled = false; x.drawImage(s, 0, 0, w, h);
  }
  return c;
}
const toBytes = (c, type) => new Promise(ok => c.toBlob(async b => ok(new Uint8Array(await b.arrayBuffer())), type, 0.9));

/** Remplace une part média par une autre (nouveau chemin) dans toutes les relations. */
function retarget(pkg, from, to) {
  for (const [rp, src] of relParts(pkg)) for (const r of rels(pkg, src)) if (!r.external && r.target === from) {
    r.el.setAttribute('Target', r.el.getAttribute('Target').replace(/[^/]+$/, to.split('/').pop()));
  }
}
/** Retire les extensions svgBlip pointant vers les SVG donnés (PowerPoint affiche alors le PNG caviardé). */
function dropSvg(pkg, svgPaths) {
  if (!svgPaths.length) return;
  for (const [, src] of relParts(pkg)) {
    const ids = new Set(rels(pkg, src).filter(r => svgPaths.includes(r.target)).map(r => r.id));
    if (!ids.size) continue;
    const doc = xml(pkg, src); if (!doc) continue;
    for (const s of all(doc, 'asvg', 'svgBlip')) if (ids.has(s.getAttributeNS(NS.r, 'embed'))) {
      const extEl = s.parentElement, lst = extEl?.parentElement;
      extEl?.remove(); if (lst && !lst.children.length) lst.remove();
    }
  }
}

/** Caviarde les images choisies : choice = Map index → { kind, label }. Renvoie le nombre d'images traitées. */
export async function redact(pkg, items, choice, radius) {
  let n = 0;
  for (const it of items) {
    const ch = choice.get(it.index); if (!ch) continue;
    const img = it.viewable ? await load(blobOf(pkg, it)).catch(() => null) : null;
    const jpeg = /^jpe?g$/.test(it.ext), c = redactCanvas(img, ch.kind, ch.label, radius);
    const data = await toBytes(c, jpeg ? 'image/jpeg' : 'image/png');
    if (jpeg || it.ext === 'png') setBytes(pkg, it.path, data);
    else {
      let to = it.path.replace(/\.[^.]+$/, '.png'); for (let k = 1; has(pkg, to); k++) to = it.path.replace(/\.[^.]+$/, `_${k}.png`);
      setBytes(pkg, to, data); ensureDefault(pkg, 'png', 'image/png'); retarget(pkg, it.path, to); remove(pkg, it.path);
    }
    dropSvg(pkg, it.svg);
    n++;
  }
  return n;
}
