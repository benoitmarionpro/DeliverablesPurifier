// Interface : dépôt → options → (analyse + revue des images) → anonymisation → résumé. Tout reste dans le navigateur.
import { openPptx } from './pptx.js';
import { inventory, analyze, blobOf, LOGO_LABELS } from './images.js';
import { purify, steps, outName, DEFAULT_OPTIONS } from './purify.js';
import { isAuthorized } from './keys.js';
import { state } from './store.js';
import { esc, $ } from './ui.js';
import { t, lang, setLang } from './i18n.js';
import * as settings from './settings.js';

const MAX_SIZE = 300e6;
const OPTS = [
  ['text', t('Texte sensible'), t('Noms, clients, projets, emails, téléphones, données chiffrées (LLM + règles locales)')],
  ['images', t('Images'), t('Logos, photos, schémas : détection par IA puis revue')],
  ['notes', t('Notes présentateur'), t('Notes de chaque slide vidées')],
  ['comments', t('Commentaires'), t('Commentaires et leurs auteurs supprimés')],
  ['layout', t('Mise en page'), t('Masters purgés (logos, fonds), couleurs neutralisées, charte Wavestone')],
  ['videos', t('Vidéos et sons'), t('Suppression complète des médias')],
  ['links', t('Hyperliens'), t('Liens retirés, texte préfixé « Lien_URL_ »')],
  ['meta', t('Métadonnées'), t('Auteur, société, miniature, propriétés, classeurs des graphiques')],
];
const KIND = { logo: t('Logo'), diagram: t('Schéma'), confidential: t('Confidentiel'), none: t('Non sensible') };
const FLOW = [['idle', t('Dépôt')], ['analyzing', t('Analyse')], ['review', t('Revue')], ['running', t('Anonymisation')], ['done', t('Terminé')]];

const S = { step: 'idle', file: null, pkg: null, opts: { ...DEFAULT_OPTIONS }, items: [], intensity: 79, filter: 'all', vision: true, ac: null, result: null, error: null, urls: [], msg: '', prog: {} };
const view = document.getElementById('view'), tabs = document.getElementById('tabs');
const radius = () => Math.round(40 * (S.intensity / 100) ** 3); // même courbe que l'app d'origine (0 à 40 px)

function reset() {
  S.urls.forEach(URL.revokeObjectURL);
  Object.assign(S, { step: 'idle', file: null, pkg: null, items: [], filter: 'all', vision: true, result: null, error: null, urls: [], msg: '', prog: {} });
  render();
}
const fail = e => { if (e.name === 'AbortError') { reset(); return; } console.error(e); Object.assign(S, { step: 'error', error: e }); render(); };

/* ---------- Écrans ---------- */
function stepper() {
  const i = FLOW.findIndex(([k]) => k === S.step);
  return `<ol class="flow">${FLOW.map(([k, l], j) => `<li class="${S.step === 'error' ? '' : j < i ? 'done' : j === i ? 'on' : ''}">${l}</li>`).join('')}</ol>`;
}
function idle() {
  const ai = isAuthorized();
  const opt = ([k, title, desc]) => `<label class="opt ${S.opts[k] ? 'on' : ''}"><input type="checkbox" data-o="${k}" ${S.opts[k] ? 'checked' : ''}><b>${title}</b><span>${desc}</span></label>`;
  return `
  <h2 class="hero">${t('Anonymise tes présentations PowerPoint clients, sans qu’elles quittent ton poste.')}</h2>
  <div class="card warnc">${t('Le fichier est traité dans ce navigateur. Seuls les textes extraits et, si tu l’acceptes, les images sont envoyés au LLM configuré dans Paramètres. Le résultat est à relire avant toute diffusion.')}</div>
  ${ai ? '' : `<div class="card warnc"><b>${t('IA non autorisée.')}</b> ${t('Sans IA : texte limité aux règles locales (emails, téléphones, URLs, IBAN), images à classer à la main.')} <button id="go-set">${t('Ouvrir Paramètres')}</button></div>`}
  <button id="drop" class="drop ${S.file ? 'has' : ''}">${S.file ? `<b>${esc(S.file.name)}</b><span>${(S.file.size / 1e6).toFixed(1)} Mo · ${t('cliquer pour changer de fichier')}</span>` : `<b>${t('Dépose un fichier .pptx ici')}</b><span>${t('ou clique pour le choisir')}</span>`}</button>
  <input id="file" type="file" accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation" hidden>
  ${S.file ? `
  <div class="bar"><h2 style="margin:0">${t('Que faut-il anonymiser ?')}</h2><button id="allopt">${Object.values(S.opts).every(Boolean) ? t('Tout décocher') : t('Tout cocher')}</button></div>
  <div class="opts">${OPTS.map(opt).join('')}</div>
  <div class="bar"><button id="start" class="pri" ${Object.values(S.opts).some(Boolean) ? '' : 'disabled'}>${t('Anonymiser')}</button></div>` : ''}`;
}
function progress() {
  const list = S.step === 'analyzing' ? [['analyze', t('Analyse des images')]] : steps(S.opts);
  return `<div class="card"><ul class="prog">${list.map(([k, l]) => `<li class="${S.prog[k] || ''}">${l}${S.prog[k] === 'run' && S.msg ? ` <span class="mut">— ${esc(S.msg)}</span>` : ''}</li>`).join('')}</ul>
  <button id="cancel">${t('Annuler')}</button></div>`;
}
function review() {
  const r = radius(), cnt = k => S.items.filter(i => k === 'all' || (k === 'check' ? i.check : (i.ai?.type || 'none') === k && !i.check)).length;
  const shown = S.items.filter(i => S.filter === 'all' || (S.filter === 'check' ? i.check : (i.ai?.type || 'none') === S.filter && !i.check));
  const card = it => `
  <div class="img ${it.sel ? 'sel' : ''}">
    <button class="thumb" data-zoom="${it.index}" aria-label="${t('Agrandir')}">${it.url ? `<img src="${it.url}" alt="" style="${it.sel && it.kind !== 'logo' ? `filter:blur(${Math.max(1, r / 4)}px)` : ''}">` : `<span class="noview">.${esc(it.ext)}<br>${t('aperçu impossible')}</span>`}${it.sel && it.kind === 'logo' ? `<i class="grey">${esc(it.label)}</i>` : ''}</button>
    <label class="chk"><input type="checkbox" data-sel="${it.index}" ${it.sel ? 'checked' : ''}> ${t('Caviarder')}</label>
    <div class="row2"><select data-kind="${it.index}" aria-label="${t('Traitement')}">${['logo', 'diagram', 'confidential'].map(k => `<option value="${k}" ${it.kind === k ? 'selected' : ''}>${k === 'logo' ? t('Carré gris + label') : t('Flou') + ' (' + KIND[k] + ')'}</option>`).join('')}</select>
    ${it.kind === 'logo' ? `<select data-label="${it.index}" aria-label="${t('Label')}">${LOGO_LABELS.map(l => `<option ${it.label === l ? 'selected' : ''}>${l}</option>`).join('')}</select>` : ''}</div>
    <small class="${it.check ? 'err' : 'mut'}">${it.ai ? `<b>${KIND[it.ai.type]}</b> · ${esc(it.ai.reason)}` : it.viewable ? t('Non analysée') : t('Format non analysable : à vérifier dans PowerPoint')} · ${t('slide(s)')} ${it.slides.join(', ')}</small>
  </div>`;
  return `
  ${S.vision ? '' : `<div class="card warnc">${isAuthorized() ? t('Le modèle configuré n’accepte pas les images : classe-les à la main (rien n’est pré-coché).') : t('IA non autorisée : classe les images à la main (rien n’est pré-coché).')}</div>`}
  <div class="bar">${[['all', t('Tout')], ['logo', KIND.logo], ['diagram', KIND.diagram], ['confidential', KIND.confidential], ['none', KIND.none], ['check', t('À vérifier')]].map(([k, l]) => `<button class="lg flt ${S.filter === k ? 'on' : ''}" data-f="${k}">${l} (${cnt(k)})</button>`).join('')}
    <button id="selall">${shown.every(i => i.sel) ? t('Tout désélectionner') : t('Tout sélectionner')}</button></div>
  <div class="bar"><label for="intens">${t('Intensité du floutage')}</label><input id="intens" type="range" min="0" max="100" value="${S.intensity}" style="width:14rem"><span class="mut">${r} px</span></div>
  ${S.items.length ? `<div class="grid">${shown.map(card).join('') || `<p class="mut">${t('Aucune image dans ce filtre.')}</p>`}</div>` : `<p class="mut">${t('Aucune image dans la présentation : le texte et la mise en page seront tout de même anonymisés.')}</p>`}
  <div class="bar"><button id="back">${t('Retour')}</button><button id="go" class="pri">${t('Anonymiser la présentation ({n} image(s) caviardée(s))', { n: S.items.filter(i => i.sel).length })}</button></div>
  <dialog id="zoom"><img alt=""><form method="dialog"><button>${t('Fermer')}</button></form></dialog>`;
}
function done() {
  const { stats: s, residuals: res, warnings, map } = S.result;
  const kpi = (n, l) => `<div class="kpi"><b>${n}</b><span>${l}</span></div>`;
  return `
  <div class="kpis">${[
    S.opts.text && kpi(s.textFragments, t('fragments de texte remplacés')),
    S.opts.images && kpi(s.imagesRedacted, t('images caviardées')),
    S.opts.videos && kpi(s.videosRemoved, t('vidéos / sons retirés')),
    S.opts.notes && kpi(s.notesCleared, t('notes vidées')),
    S.opts.comments && kpi(s.commentsRemoved, t('fichiers de commentaires supprimés')),
    S.opts.links && kpi(s.linksRemoved, t('hyperliens retirés')),
    S.opts.layout && kpi(t('OUI'), t('mise en page anonymisée')),
    S.opts.meta && kpi(s.embedsRemoved, t('classeurs de graphiques retirés')),
  ].filter(Boolean).join('')}</div>
  ${res.length ? `<div class="card warnc"><b>${t('{n} terme(s) encore présent(s) dans le fichier', { n: res.length })}</b><ul>${res.map(r => `<li><code>${esc(r.original)}</code> <span class="mut">${esc(r.parts.join(', '))}</span></li>`).join('')}</ul></div>`
    : S.opts.text ? `<div class="card"><span class="ok">✓ ${t('Aucun terme détecté ne subsiste dans le fichier.')}</span></div>` : ''}
  ${warnings.map(w => `<div class="card warnc">${esc(w)}</div>`).join('')}
  ${map.size ? `<details class="card"><summary>${t('Correspondances appliquées ({n})', { n: map.size })}</summary><table><tr><th>${t('Original')}</th><th>${t('Remplacé par')}</th></tr>${[...map].map(([o, p]) => `<tr><td>${esc(o)}</td><td><code>${esc(p)}</code></td></tr>`).join('')}</table></details>` : ''}
  <div class="card"><label for="oname">${t('Nom du fichier')}</label><input id="oname" value="${esc(outName(S.file.name, map))}">
  <div class="bar"><a id="dl" class="btn pri" href="${S.result.url}" download="${esc(outName(S.file.name, map))}">${t('Télécharger')}</a><button id="again">${t('Nouveau document')}</button></div>
  <p class="mut">${t('Vérifie une dernière fois le document dans PowerPoint avant de le partager : les propositions du LLM sont à relire.')}</p></div>`;
}
function error() {
  return `<div class="card warnc"><b>${t('Échec du traitement')}</b><p>${esc(S.error?.message || S.error)}</p><button id="again" class="pri">${t('Réessayer')}</button></div>`;
}

/* ---------- Actions ---------- */
async function pick(file) {
  if (!file) return;
  if (!/\.pptx$/i.test(file.name)) return alert(t('Seuls les fichiers .pptx sont acceptés (enregistre les .ppt au format .pptx).'));
  if (file.size > MAX_SIZE) return alert(t('Fichier trop volumineux (> 300 Mo).'));
  try { const pkg = await openPptx(await file.arrayBuffer()); S.urls.forEach(URL.revokeObjectURL); Object.assign(S, { file, pkg, urls: [] }); render(); }
  catch (e) { alert(e.message); }
}
async function start() {
  S.ac = new AbortController();
  if (!S.opts.images) return run();
  S.items = inventory(S.pkg, { withMasters: !S.opts.layout });
  for (const it of S.items) if (it.viewable) S.urls.push(it.url = URL.createObjectURL(blobOf(S.pkg, it)));
  S.vision = isAuthorized();
  if (S.vision && S.items.some(i => i.viewable)) {
    Object.assign(S, { step: 'analyzing', prog: { analyze: 'run' }, msg: '' }); render();
    try { S.vision = (await analyze(S.pkg, S.items, { signal: S.ac.signal, say: m => { S.msg = m; render(); } })).vision; }
    catch (e) { return fail(e); }
  }
  for (const it of S.items) {
    it.sel = !!it.ai?.sensitive; it.check = !!it.ai?.error || !it.viewable;
    it.kind = it.ai?.type && it.ai.type !== 'none' ? it.ai.type : 'confidential'; it.label = it.ai?.label || 'Client';
  }
  if (!S.items.length) return run();
  S.step = 'review'; render();
}
async function run() {
  Object.assign(S, { step: 'running', prog: {}, msg: '' }); render();
  const choice = new Map(S.items.filter(i => i.sel).map(i => [i.index, { kind: i.kind, label: i.label }]));
  try {
    const r = await purify(S.pkg, S.opts, { items: S.items, choice, radius: radius(), ai: isAuthorized(), budget: state.settings.budget, signal: S.ac.signal, on: (k, st, m) => { S.prog[k] = st; S.msg = m || ''; render(); } });
    S.urls.push(r.url = URL.createObjectURL(r.blob));
    Object.assign(S, { step: 'done', result: r, pkg: null }); render(); // le package est modifié : un nouveau traitement repart du fichier
  } catch (e) { fail(e); }
}

function bind() {
  const on = (sel, ev, fn) => view.querySelectorAll(sel).forEach(e => (e[ev] = fn));
  on('#go-set', 'onclick', () => show('settings'));
  const input = $(view, '#file');
  on('#drop', 'onclick', () => input.click());
  if (input) input.onchange = () => pick(input.files[0]);
  const drop = $(view, '#drop');
  if (drop) {
    drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
    drop.ondragleave = () => drop.classList.remove('over');
    drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); pick(e.dataTransfer.files[0]); };
  }
  on('[data-o]', 'onchange', e => { S.opts[e.target.dataset.o] = e.target.checked; render(); });
  on('#allopt', 'onclick', () => { const v = !Object.values(S.opts).every(Boolean); for (const k in S.opts) S.opts[k] = v; render(); });
  on('#start', 'onclick', start);
  on('#cancel', 'onclick', () => S.ac?.abort());
  on('[data-f]', 'onclick', e => { S.filter = e.target.dataset.f; render(); });
  on('#selall', 'onclick', () => { const sh = S.items.filter(i => S.filter === 'all' || (S.filter === 'check' ? i.check : (i.ai?.type || 'none') === S.filter && !i.check)), v = !sh.every(i => i.sel); sh.forEach(i => (i.sel = v)); render(); });
  on('#intens', 'oninput', e => { S.intensity = +e.target.value; e.target.nextElementSibling.textContent = radius() + ' px'; });
  on('#intens', 'onchange', render);
  on('[data-sel]', 'onchange', e => { S.items[e.target.dataset.sel].sel = e.target.checked; render(); });
  on('[data-kind]', 'onchange', e => { const it = S.items[e.target.dataset.kind]; it.kind = e.target.value; it.sel = true; render(); });
  on('[data-label]', 'onchange', e => { S.items[e.target.dataset.label].label = e.target.value; render(); });
  on('[data-zoom]', 'onclick', e => { const it = S.items[e.currentTarget.dataset.zoom], d = $(view, '#zoom'); if (!it.url) return; d.querySelector('img').src = it.url; d.showModal(); });
  on('#back', 'onclick', () => { S.step = 'idle'; render(); });
  on('#go', 'onclick', run);
  on('#again', 'onclick', reset);
  on('#oname', 'oninput', e => { $(view, '#dl').download = e.target.value.trim() || 'Anonymized.pptx'; });
}
function render() {
  if (location.hash === '#settings') return;
  const focus = document.activeElement?.id;
  view.innerHTML = stepper() + { idle, analyzing: progress, running: progress, review, done, error }[S.step]();
  bind();
  if (focus) document.getElementById(focus)?.focus();
}

/* ---------- Onglets, langue ---------- */
const views = { home: t('Anonymiser'), settings: t('Paramètres') };
function show(id) {
  [...tabs.children].forEach(b => { const on = b.dataset.id === id; b.classList.toggle('on', on); on ? b.setAttribute('aria-current', 'page') : b.removeAttribute('aria-current'); });
  location.hash = id;
  if (id === 'settings') { view.replaceChildren(); settings.render(view); } else render();
}
for (const [id, label] of Object.entries(views)) {
  const b = document.createElement('button'); b.textContent = label; b.dataset.id = id; b.onclick = () => show(id); tabs.append(b);
}
document.documentElement.lang = lang;
tabs.setAttribute('aria-label', t('Onglets'));
document.querySelector('footer').textContent = t('DeliverablesPurifier · Wavestone — Anonymisation locale de livrables PowerPoint · Les propositions du LLM sont à relire');
const lb = document.getElementById('lang');
lb.textContent = lang === 'en' ? 'FR' : 'EN';
lb.title = lang === 'en' ? 'Passer l’interface en français' : 'Switch the interface to English';
lb.onclick = () => { if (S.step === 'idle' || S.step === 'done' || confirm(t('Changer de langue recharge la page et interrompt le traitement en cours. Continuer ?'))) { setLang(lang === 'en' ? 'fr' : 'en'); location.reload(); } };
window.addEventListener('beforeunload', e => { if (S.step === 'running' || S.step === 'analyzing' || S.step === 'review') e.preventDefault(); });
show(location.hash === '#settings' ? 'settings' : 'home');
