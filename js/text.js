// Texte : extraction (slides, masters, layouts, notes, graphiques, SmartArt, alt text, docProps), détection (regex + LLM), remplacement.
import { NS, all, xml, partsLike, names, isXml, text as partText } from './pptx.js';
import { chatJson, fence, UNTRUSTED } from './llm.js';
import { t } from './i18n.js';

/** Placeholders autorisés (liste fermée du prompt d'origine) ; URL / Redacted ajoutés pour les regex et les sorties hors liste. */
export const PLACEHOLDERS = ['Luxury_Client', 'Bank_Client', 'Insurance_Client', 'Retail_Client', 'Energy_Client', 'Tech_Client', 'Healthcare_Client', 'Public_Sector_Client', 'Client',
  'Company', 'Partner', 'Project', 'Person', 'Executive', 'Location', 'Date', 'Email', 'Phone', 'Financial_Data'];
const EXTRA = ['URL', 'IBAN', 'Redacted'];

/** Parts dont le texte est anonymisé (les notes seulement si elles ne sont pas vidées). */
export const textParts = (pkg, { notes }) => partsLike(pkg, new RegExp('^(ppt/(slides|slideLayouts|slideMasters|notesMasters|handoutMasters|charts|diagrams' + (notes ? '|notesSlides' : '') + ')/[^/]+\\.xml|docProps/(core|app)\\.xml)$'));

/** Unités de texte d'une part : paragraphes DrawingML (a:p), valeurs de graphique, attributs (alt text, noms), propriétés du document. */
export function units(doc) {
  const out = [];
  for (const p of all(doc, 'a', 'p')) {
    const nodes = all(p, 'a', 't');
    if (nodes.length) out.push({ nodes, text: nodes.map(n => n.textContent).join('') });
  }
  for (const v of all(doc, 'c', 'v')) if (v.parentNode?.parentNode?.localName === 'strCache' || v.parentNode?.localName === 'tx') out.push({ nodes: [v], text: v.textContent });
  for (const el of doc.getElementsByTagName('*')) {
    if (el.localName === 'cNvPr') for (const a of ['descr', 'title', 'name']) if (el.getAttribute(a)) out.push({ el, attr: a, text: el.getAttribute(a) });
    // docProps : titre, sujet, mots-clés, auteur, société, titres des slides (app.xml TitlesOfParts)
    if (/^(title|subject|keywords|description|category|creator|lastModifiedBy|Company|Manager|lpstr)$/.test(el.localName) && !el.children.length && el.textContent) out.push({ nodes: [el], text: el.textContent });
  }
  return out.filter(u => u.text.trim());
}

/* ---------- Détection déterministe ---------- */
const RX = [
  ['Email', /[\w.+-]+@[\w-]+(\.[\w-]+)+/g],
  ['URL', /\b(https?:\/\/|www\.)[^\s"<>]+/gi],
  ['IBAN', /\b[A-Z]{2}\d{2}(?: ?[A-Z0-9]{4}){3,7}(?: ?[A-Z0-9]{1,3})?\b/g],
  ['Phone', /(?<![\w+])(?:\+\d{1,3}[ .-]?(?:\(0\))?[ .-]?\d(?:[ .-]?\d{2}){4}|0\d(?:[ .-]?\d{2}){4})(?!\d)/g],
];
export function regexFind(s) {
  const found = new Map();
  for (const [ph, re] of RX) for (const m of s.matchAll(re)) found.set(m[0].replace(/[.,;:)]+$/, ''), ph);
  return found;
}

/* ---------- Détection LLM (prompt de l'app d'origine, découpé par lots) ---------- */
const PROMPT = `Identifie toutes les informations sensibles dans le texte suivant extrait d'une présentation PowerPoint, et remplace-les par des placeholders LOGIQUES selon le type d'entité.

Règles pour les placeholders :
- Utilise UNIQUEMENT les placeholders listés ci-dessous (format exact, sans crochets).
- Choisis le placeholder le plus pertinent selon le contexte (secteur, type d'entité).

Placeholders autorisés :

**Clients (selon le secteur quand c'est identifiable) :**
- Luxury_Client — marques de luxe, retail luxe, joaillerie, haute couture
- Bank_Client — banques, établissements financiers
- Insurance_Client — assurances
- Retail_Client — grande distribution, retail grand public
- Energy_Client — énergie, utilities, industrie lourde
- Tech_Client — tech, éditeurs logiciel, telecoms
- Healthcare_Client — santé, pharma, hôpitaux
- Public_Sector_Client — administration, collectivités, secteur public
- Client — client non identifiable ou générique

**Autres entités :**
- Company — nom d'entreprise (hors client)
- Partner — partenaire, alliance
- Project — nom de projet ou de mission
- Person — nom de personne
- Executive — direction, C-level quand explicite
- Location — lieu, adresse, ville spécifique, bureau
- Date — date précise (jour, mois, année)
- Email — adresse email
- Phone — numéro de téléphone

**Données financières :**
- Financial_Data — toute donnée financière sensible : chiffre d'affaires (CA), revenus, marges, bénéfices, budgets, montants de contrats, TJM, valorisation, etc.
Si la phrase n'a plus de sens après l'anonymisation, remplace-la entièrement par Financial_Data.

Crée une liste de mappings où "original" est la chaîne sensible EXACTE trouvée dans le texte
et "replacement" est UN des placeholders ci-dessus (ex: Bank_Client, Luxury_Client, Project, Financial_Data).

Sois exhaustif. N'inclus PAS les termes génériques non sensibles (ex: "Réunion", "Stratégie", "Trimestriel").
Pour un même acteur (ex: une banque nommée), utilise le MÊME placeholder partout dans le texte.
Ne propose pas comme "original" un nom de l'éditeur du document (Wavestone) ni un placeholder déjà présent.

Retourne UNIQUEMENT un JSON valide avec cette structure :
{
  "entities": [
    {"original": "texte sensible", "replacement": "Placeholder"},
    ...
  ]
}`;

/** Lots de lignes uniques, sous un budget de caractères. */
export function batches(lines, budget) {
  const out = [[]]; let n = 0;
  for (const l of lines) {
    if (n + l.length > budget && out.at(-1).length) { out.push([]); n = 0; }
    out.at(-1).push(l); n += l.length + 1;
  }
  return out.filter(b => b.length);
}
const norm = s => s.replace(/\s+/g, ' ').trim();
/** Ne garde que les entités valides : placeholder autorisé, original présent dans le texte source. */
export function validate(entities, source) {
  const ok = new Map(), lower = Object.fromEntries([...PLACEHOLDERS, ...EXTRA].map(p => [p.toLowerCase(), p]));
  for (const e of Array.isArray(entities) ? entities : []) {
    const o = norm(String(e?.original ?? '')), r = String(e?.replacement ?? '').replace(/[[\]{}<>]/g, '').trim();
    if (o.length < 2 || lower[o.toLowerCase()] || !source.includes(o)) continue;
    ok.set(o, lower[r.toLowerCase()] || 'Redacted');
  }
  return ok;
}

/** Détecte les entités sensibles : regex puis LLM par lots ; renvoie Map original → placeholder. */
export async function detect(lines, { budget, signal, say }) {
  const source = lines.join('\n'), map = regexFind(source), parts = batches(lines, budget);
  for (let i = 0; i < parts.length; i++) {
    say?.(t('Analyse du texte : lot {i}/{n}', { i: i + 1, n: parts.length }));
    const known = [...map].slice(-150).map(([o, r]) => `${o} → ${r}`).join('\n');
    const user = (known ? `Correspondances déjà établies (réutilise le même placeholder pour ces acteurs) :\n${known}\n\n` : '') + 'Texte :\n' + fence('deck', parts[i].join('\n'));
    const r = await chatJson({ system: PROMPT + '\n\n' + UNTRUSTED, user, signal });
    if (!Array.isArray(r?.entities)) throw new Error(t('Réponse du LLM sans liste « entities » (lot {i}) : anonymisation du texte interrompue.', { i: i + 1 }));
    for (const [o, p] of validate(r?.entities, source)) if (!map.has(o)) map.set(o, p);
  }
  return map;
}

/* ---------- Remplacement ---------- */
const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
/** Regex unique (plus longs d'abord), bornée aux mots quand l'original commence / finit par une lettre ou un chiffre. */
export function matcher(map) {
  const keys = [...map.keys()].sort((x, y) => y.length - x.length);
  if (!keys.length) return null;
  const src = keys.map(k => (/^[\p{L}\p{N}]/u.test(k) ? '(?<![\\p{L}\\p{N}])' : '') + esc(k).replace(/ /g, '\\s+') + (/[\p{L}\p{N}]$/u.test(k) ? '(?![\\p{L}\\p{N}])' : '')).join('|');
  return new RegExp(src, 'gu');
}
const lookup = (map, m) => map.get(m) ?? map.get(norm(m)) ?? 'Redacted';

/** Remplace dans une unité ; un original à cheval sur plusieurs runs est écrit dans le premier, les suivants sont rognés. Renvoie le nombre de remplacements. */
export function replaceUnit(u, map, re) {
  const hits = [...u.text.matchAll(re)];
  if (!hits.length) return 0;
  if (u.attr) { u.el.setAttribute(u.attr, u.text.replace(re, m => lookup(map, m))); return hits.length; }
  let pos = 0, h = 0;
  for (const node of u.nodes) {
    const s = node.textContent; let out = '';
    for (let i = 0; i < s.length; i++, pos++) {
      while (h < hits.length && pos >= hits[h].index + hits[h][0].length) h++;
      const m = hits[h];
      if (m && pos >= m.index) { if (pos === m.index) out += lookup(map, m[0]); }
      else out += s[i];
    }
    node.textContent = out;
  }
  return hits.length;
}

/** Lignes uniques à analyser. */
export function collect(pkg, opts) {
  const seen = new Set();
  for (const n of textParts(pkg, opts)) { const d = xml(pkg, n); if (d) for (const u of units(d)) seen.add(norm(u.text)); }
  return [...seen].filter(l => l.length > 1);
}
/** Applique le mapping à toutes les parts concernées ; renvoie le nombre de fragments remplacés. */
export function apply(pkg, map, opts) {
  const re = matcher(map); if (!re) return 0;
  let n = 0;
  for (const p of textParts(pkg, opts)) { const d = xml(pkg, p); if (d) for (const u of units(d)) n += replaceUnit(u, map, re); }
  return n;
}
/** Vérification résiduelle : originaux encore présents dans une part XML du résultat → [{ original, parts[] }]. */
export function residuals(pkg, map) {
  const out = [];
  const xmlEsc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const texts = names(pkg).filter(isXml).map(n => [n, partText(pkg, n)]);
  for (const o of map.keys()) {
    const re = matcher(new Map([[xmlEsc(o), 1]])), parts = texts.filter(([, s]) => (re.lastIndex = 0, re.test(s))).map(([n]) => n);
    if (parts.length) out.push({ original: o, parts });
  }
  return out;
}
