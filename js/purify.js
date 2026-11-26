// Pipeline d'anonymisation : ordre des étapes, statistiques, vérification résiduelle. Sans interface (testable).
import { savePptx, dangling } from './pptx.js';
import { collect, detect, regexFind, apply, residuals, matcher } from './text.js';
import { redact } from './images.js';
import { clearNotes, removeComments, removeMedia, removeLinks, scrubMeta, prune } from './clean.js';
import { anonymizeLayout } from './layout.js';
import { t } from './i18n.js';

export const DEFAULT_OPTIONS = { text: true, images: true, notes: true, comments: true, layout: true, videos: true, links: true, meta: true };

/** Étapes actives, dans l'ordre d'exécution : [clé, libellé]. Les images d'abord (leurs chemins viennent de l'inventaire initial). */
export const steps = o => [
  o.images && ['images', t('Caviardage des images')],
  o.notes && ['notes', t('Notes du présentateur')],
  o.comments && ['comments', t('Commentaires')],
  o.videos && ['videos', t('Vidéos et sons')],
  o.links && ['links', t('Hyperliens')],
  o.layout && ['layout', t('Mise en page')],
  o.text && ['text', t('Texte sensible')],
  o.meta && ['meta', t('Métadonnées et objets embarqués')],
  ['save', t('Vérification et enregistrement')],
].filter(Boolean);

/**
 * Anonymise pkg en place et renvoie { blob, stats, map, residuals, warnings }.
 * items / choice : inventaire et choix de la revue des images ; ai : détection du texte par LLM autorisée.
 * on(step, état, message) : progression.
 */
export async function purify(pkg, o, { items = [], choice = new Map(), radius = 20, ai, budget = 12000, signal, on = () => {} }) {
  const stats = { textFragments: 0, imagesRedacted: 0, videosRemoved: 0, notesCleared: 0, commentsRemoved: 0, linksRemoved: 0, layout: false, embedsRemoved: 0, orphansRemoved: 0 };
  const warnings = [];
  let map = new Map();
  for (const [k] of steps(o)) {
    signal?.throwIfAborted();
    on(k, 'run');
    if (k === 'images') stats.imagesRedacted = await redact(pkg, items, choice, radius);
    if (k === 'notes') stats.notesCleared = clearNotes(pkg);
    if (k === 'comments') stats.commentsRemoved = removeComments(pkg);
    if (k === 'videos') stats.videosRemoved = removeMedia(pkg);
    if (k === 'links') stats.linksRemoved = removeLinks(pkg);
    if (k === 'layout') { anonymizeLayout(pkg); stats.layout = true; }
    if (k === 'text') {
      const lines = collect(pkg, { notes: !o.notes });
      map = ai ? await detect(lines, { budget, signal, say: m => on(k, 'run', m) }) : regexFind(lines.join('\n'));
      if (!ai) warnings.push(t('IA non autorisée : seuls les emails, téléphones, URLs et IBAN ont été remplacés. Les noms de clients, personnes et projets ne sont pas détectés.'));
      stats.textFragments = apply(pkg, map, { notes: !o.notes });
    }
    if (k === 'meta') {
      const m = scrubMeta(pkg); stats.embedsRemoved = m.embeds;
      if (m.ole) warnings.push(t('{n} objet(s) OLE embarqué(s) (Excel, Word…) conservé(s) : vérifie leur contenu ou supprime-les dans PowerPoint.', { n: m.ole }));
    }
    if (k === 'save') {
      stats.orphansRemoved = prune(pkg);
      const res = residuals(pkg, map), bad = dangling(pkg);
      if (bad.length) warnings.push(t('Références internes rompues ({p}) : PowerPoint risque de proposer une réparation du fichier. Merci de signaler ce cas.', { p: bad.map(b => `${b.part} ${b.ids.join(',')}`).join(' ; ') }));
      const blob = await savePptx(pkg);
      on(k, 'ok');
      return { blob, stats, map, residuals: res, warnings };
    }
    on(k, 'ok');
  }
}

/** Nom du fichier produit : mapping appliqué au nom d'origine (un nom de client peut s'y trouver). */
export function outName(name, map) {
  const re = matcher(map), base = name.replace(/\.pptx$/i, '');
  return 'Anonymized_' + (re ? base.replace(re, m => map.get(m) ?? 'Redacted') : base) + '.pptx';
}
