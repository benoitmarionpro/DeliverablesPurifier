// Paramètres : autorisation de l'IA, fournisseur, clé, modèle (repris de MITREMapper).
import { state, save } from './store.js';
import { PROVIDERS, llm } from './llm.js';
import { getKey, setKey, clearKeys, isRemembered, mask, isAuthorized, setAuthorized } from './keys.js';
import { esc, $ } from './ui.js';
import { t } from './i18n.js';

export function render(root) {
  const s = state.settings;
  root.innerHTML = `
  <h2>${t('Utilisation de l’IA')}</h2>
  <label class="chk"><input type="checkbox" id="auth"> I confirm that AI use is authorized for this content and complies with the client's rules.</label>
  <p class="mut">${t('Je confirme que l’usage de l’IA est autorisé pour ce contenu et conforme aux règles du client. Aucun appel IA n’est envoyé tant que cette case n’est pas cochée ; elle est à reconfirmer à chaque session. Sans IA, seules les règles locales (emails, téléphones, URLs, IBAN) et la revue manuelle des images sont disponibles.')}</p>
  <h2>${t('Endpoint LLM')}</h2>
  <div class="row">
    <div><label for="prov">${t('Fournisseur')}</label><select id="prov">${Object.entries(PROVIDERS).map(([k, p]) => `<option value="${k}">${esc(t(p.label))}</option>`).join('')}</select></div>
    <div><label for="base" id="base-l"></label><input id="base"></div>
    <div><label for="key" id="key-l"></label><input id="key" type="password" autocomplete="off"></div>
    <div data-hide="azure"><label for="model">${t('Modèle')}</label><input id="model" list="models" autocomplete="off"><datalist id="models"></datalist><small id="mst" class="mst"></small></div>
    <div data-show="azure"><label for="dep">${t('Déploiement')}</label><input id="dep"></div>
    <div data-show="azure"><label for="apiv">${t('Version d’API')}</label><input id="apiv"></div>
    <div><label for="temp">${t('Température')}</label><input id="temp" type="number" step="0.1" min="0" max="2"></div>
    <div><label for="budget">${t('Taille d’un lot de texte (caractères)')}</label><input id="budget" type="number" min="2000" max="200000" step="1000"></div>
  </div>
  <label class="chk"><input type="checkbox" id="rem"> ${t('Mémoriser cette clé sur ce navigateur')}</label>
  <p class="mut">${t('Par défaut la clé reste en mémoire de session (effacée à la fermeture de l’onglet). Elle n’est envoyée qu’à l’URL ci-dessus (HTTPS exigé, sauf en local). Le serveur doit autoriser CORS pour {origin} (fichier ouvert par double-clic : origine {nul}).', { origin: `<code>${esc(location.origin)}</code>`, nul: '<code>null</code>' })}</p>
  <p class="mut">${t('Analyse des images : le modèle doit accepter les images (vision). Sinon, la revue des images se fait manuellement.')}</p>
  <div class="bar"><button id="models-btn">${t('Lister les modèles')}</button><button id="test" class="pri">${t('Tester la connexion')}</button><button id="clear">${t('Effacer les clés API')}</button><span id="pill" class="pill"></span></div>
  <p id="st" class="mut"></p>`;

  const f = { base: 'base', model: 'model', dep: 'deployment', apiv: 'apiVersion', temp: 'temperature', budget: 'budget' };
  const st = $(root, '#st'), pill = $(root, '#pill');
  const showStatus = () => {
    const k = llm.status.state;
    pill.className = 'pill ' + k;
    pill.textContent = { idle: t('Non testé'), busy: t('En cours…'), ok: 'OK', warn: t('Liste de secours'), error: t('Erreur') }[k];
  };
  const applyProvider = () => {
    const p = PROVIDERS[s.provider], az = p.kind === 'azure';
    root.querySelectorAll('[data-show]').forEach(e => (e.hidden = !az));
    root.querySelectorAll('[data-hide]').forEach(e => (e.hidden = az));
    $(root, '#base-l').textContent = az ? t('Endpoint HTTPS Azure (https://<ressource>.openai.azure.com)') : t('URL de base');
    $(root, '#key-l').textContent = p.keyRequired ? t('Clé API (requise)') : t('Clé API (optionnelle)');
    const k = getKey(s.provider);
    $(root, '#key').value = ''; $(root, '#key').placeholder = k ? t('Clé enregistrée {mask}', { mask: mask(k) }) : '';
    $(root, '#rem').checked = isRemembered(s.provider);
  };
  const fill = () => { for (const [id, k] of Object.entries(f)) $(root, '#' + id).value = s[k]; $(root, '#prov').value = s.provider; applyProvider(); };
  const mst = $(root, '#mst'), mb = $(root, '#models-btn');
  const modelStatus = (kind, text) => { mst.className = 'mst ' + kind; mst.textContent = text; };
  const setModels = list => { $(root, '#models').innerHTML = list.map(m => `<option value="${esc(m)}">`).join(''); };
  const defaults = () => {
    setModels(PROVIDERS[s.provider].models);
    modelStatus('', isAuthorized() ? t('{n} modèles suggérés (liste par défaut du fournisseur).', { n: PROVIDERS[s.provider].models.length }) : t('Liste non chargée : cocher l’autorisation d’usage de l’IA ci-dessus.'));
  };
  let loaded = '', busy = false;
  const loadModels = async () => {
    if (busy) return; busy = true;
    modelStatus('busy', t('Chargement des modèles…')); mb.disabled = true;
    try {
      const r = await llm.getAvailableModels();
      setModels(r.models);
      if (r.fallback) { modelStatus('warn', t('Liste de secours ({n}) : chargement impossible, détail ci-dessous.', { n: r.models.length })); st.textContent = r.reason; }
      else { modelStatus('ok', t('{n} modèles disponibles', { n: r.models.length })); loaded = s.provider + s.base; }
    } catch (e) { modelStatus('error', t('Échec du chargement, détail ci-dessous.')); st.textContent = e.message; }
    busy = false; mb.disabled = false; showStatus();
  };
  fill(); showStatus(); st.textContent = llm.status.text; defaults();
  if (isAuthorized()) loadModels();
  // le datalist filtre sur la valeur saisie : champ vidé au focus pour montrer toute la liste, restauré si rien n'est choisi
  const model = $(root, '#model');
  model.onfocus = () => { model.placeholder = model.value; model.value = ''; if (isAuthorized() && loaded !== s.provider + s.base) loadModels(); };
  model.onblur = () => { if (!model.value) model.value = model.placeholder; model.placeholder = ''; };
  $(root, '#auth').checked = isAuthorized();
  $(root, '#auth').onchange = e => { setAuthorized(e.target.checked); e.target.checked ? loadModels() : defaults(); };
  for (const [id, k] of Object.entries(f)) $(root, '#' + id).onchange = e => { s[k] = e.target.type === 'number' ? +e.target.value : e.target.value; save(); };
  $(root, '#prov').onchange = e => {
    const { base, model } = PROVIDERS[e.target.value];
    Object.assign(s, { provider: e.target.value, base, model }); save(); fill(); defaults();
    if (isAuthorized()) loadModels();
  };
  $(root, '#key').onchange = e => {
    if (!e.target.value.trim()) return;
    setKey(s.provider, e.target.value.trim(), $(root, '#rem').checked); applyProvider();
    if (isAuthorized()) loadModels();
  };
  $(root, '#rem').onchange = e => { const k = getKey(s.provider); if (k) setKey(s.provider, k, e.target.checked); };
  $(root, '#clear').onclick = () => { if (confirm(t('Effacer toutes les clés API de ce navigateur (session et mémorisées) ?'))) { clearKeys(); applyProvider(); st.textContent = t('Clés API effacées.'); } };
  mb.onclick = loadModels;
  $(root, '#test').onclick = async () => {
    st.textContent = ''; pill.className = 'pill busy'; pill.textContent = t('En cours…');
    try { st.textContent = t('Connexion OK : {r}', { r: await llm.testConnection() }); } catch (e) { st.textContent = e.message; }
    showStatus();
  };
}
