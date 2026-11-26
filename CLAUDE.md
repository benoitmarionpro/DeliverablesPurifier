# DeliverablesPurifier

## Objectif
Anonymiser des livrables PowerPoint (.pptx) Wavestone dans le navigateur, sans backend. Portage statique de « Wave Anonymizer » (LAB CTO.ai, `powerpoint-anonymizer-main/`, FastAPI + Gemini sur GCP : référence seulement, non utilisée à l'exécution), avec les fuites de l'original colmatées.

## Contraintes (identiques à MITREMapper)
- Statique pur : pas de backend, pas de dépendance npm au runtime, ES modules natifs, aucune requête externe hors de l'endpoint LLM configuré. Dev : `python tools/serve.py` (sans cache, port 8000). Fichier unique : `node tools/build-standalone.mjs` → `dist/deliverablespurifier.html` (ouvrable en `file://`). **Régénérer `dist/` après toute modification de `js/`, `css/` ou `index.html`.**
- Règles Wavestone : `CYBAIMaker.md` (approbation avant tout appel IA, secrets, `LLMClient` unique, erreurs contextualisées).
- Le .pptx est ouvert et réécrit localement, mais son contenu part au LLM : textes extraits et, si le modèle accepte la vision, images réduites (1024 px, JPEG). Ne jamais écrire dans l'interface ou la doc que le document « ne quitte pas le poste ». Rien du document n'est stocké (`localStorage` : réglages `purifier.settings`, langue `purifier.lang` ; clés : `purifier.keys` en session par défaut ; autorisation IA : `purifier.ai-ok` en session).
- LLM : registre `PROVIDERS` de `js/llm.js` (copié de MITREMapper, `generate({ system, user, images })` ajoute la vision au format OpenAI / Anthropic). Proxy de test : `http://127.0.0.1:11434/v1` (CoPro, **texte seulement** : les images passent en revue manuelle). Pour `file://`, l'origine `null` doit être autorisée côté proxy.
- Mini-bundler : uniquement `export const/let/function/class NOM` en début de ligne, pas de live binding, pas de cycle d'import.
- Code minimal : pas d'abstraction ni de framework sans besoin avéré.

## Structure
| Fichier | Rôle |
|---|---|
| `js/zip.js` | Lecture / écriture zip (deflate natif `DecompressionStream` / `CompressionStream`). Entrée non modifiée = octets d'origine réécrits. Ni Zip64 ni chiffrement |
| `js/pptx.js` | Package OOXML : `openPptx`, `xml(pkg, part)` (DOM parsé une fois, resérialisé à l'enregistrement), `rels`, `relsOf` (`''` = racine), `slides` (ordre de `presentation.xml`), `remove` (part + Override + ses rels), `ensureDefault`, `savePptx` |
| `js/text.js` | `units()` : paragraphes `a:p` (runs concaténés), `c:v` de graphiques, attributs `descr/title/name` de `cNvPr`, champs docProps. `regexFind` (email, URL, IBAN, téléphone), `detect` (prompt d'origine par lots + mapping déjà connu), `validate` (placeholder en liste fermée sinon `Redacted`, original présent dans le texte), `replaceUnit` (remplacement à cheval sur plusieurs runs, borné aux mots), `residuals` |
| `js/images.js` | `inventory` (images des slides dans l'ordre, icônes exclues par nom de forme, couple SVG + PNG fusionné, masters/layouts si la mise en page est désactivée), `analyze` (prompt d'origine, 4 en parallèle, repli manuel si le modèle refuse les images, erreur = « à vérifier » pré-cochée), `redact` (canvas : flou ou carré gris + label ; formats non décodables → PNG gris, rels redirigées) |
| `js/clean.js` | Notes, commentaires (anciens, modernes, auteurs), vidéos/sons, hyperliens, métadonnées (docProps, miniature, custom.xml, customXml, classeurs des graphiques), `prune` (relations média et médias orphelins) |
| `js/layout.js` | Portage de `LayoutAnonymizer` (ordre des enfants de `a:rPr` respecté : sinon réparation PowerPoint) |
| `js/purify.js` | Pipeline (`steps`, `purify`, `outName`) sans interface. Ordre : images d'abord (chemins de l'inventaire), texte après mise en page |
| `js/app.js`, `js/settings.js` | Interface (dépôt → options → analyse/revue des images → traitement → résumé) et paramètres LLM |
| `test.html` | Tests unitaires + parcours d'interface sans IA (`document.title` = `PASS n` / `FAIL n`). Fixtures : `tests/sample.pptx` (app d'origine), `tests/rich.pptx` (graphique, SmartArt, commentaire moderne, notes, lien ; identités fictives) |
| `tests/e2e.html` | Manuel, envoie la fixture au LLM configuré : `tests/e2e.html?run=1&file=rich.pptx[&dump=1]` |

## Conventions
- Tout texte affiché passe par `t('texte français', { var })` avec son entrée dans `EN` (`js/i18n.js`, vérifié par `test.html`). Les prompts restent ceux de l'app d'origine (français).
- Sortie LLM non fiable : `chatJson` puis `validate` / `validateVision` ; contenu du deck entre balises `fence('deck', …)` + `UNTRUSTED`.
- Échec = visible : JSON sans `entities` interrompt le texte avec une erreur ; image en erreur pré-cochée « à vérifier » (l'original laissait passer).
- Toute nouvelle suppression de part passe par `remove()` (Override et rels nettoyés) ; finir par `prune()`. Supprimer une relation impose de retirer aussi les éléments qui la citent (`p:custData`, `p188:commentRel`, `a:hlinkClick`…) : `dangling()` (`pptx.js`) le vérifie en fin de pipeline et dans `test.html`.
- Vérifier : `node --check js/*.js`, `test.html` (Chrome headless : `--virtual-time-budget=300000`, le canvas est lent en temps virtuel), régénérer `dist/` et l'ouvrir en `file://`, ouvrir un résultat dans PowerPoint (aucune réparation demandée).

## Hors périmètre actuel
Objets OLE embarqués (signalés, pas supprimés), `ppt/tags` (données d'add-ins), floutage par zone (image entière seulement), OCR, .ppt / .docx / .xlsx.
