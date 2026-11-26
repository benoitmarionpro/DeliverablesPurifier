// Réglages LLM (jamais de clé ici : voir keys.js). Aucun contenu de présentation n'est enregistré.
const KEY = 'purifier.settings';
const DEFAULTS = { provider: 'local', base: 'http://127.0.0.1:11434/v1', model: 'auto', deployment: '', apiVersion: '2024-10-21', temperature: 0.1, budget: 12000 };
export const state = { settings: { ...DEFAULTS } };
try { Object.assign(state.settings, JSON.parse(localStorage.getItem(KEY)) || {}); } catch {}
export const save = () => { try { localStorage.setItem(KEY, JSON.stringify(state.settings)); } catch {} };
