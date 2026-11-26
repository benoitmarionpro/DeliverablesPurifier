# DeliverablesPurifier

Anonymisation de présentations PowerPoint (.pptx) **dans le navigateur, sans serveur** : le fichier est ouvert et réécrit localement, mais **son contenu part au LLM** de ton choix (configuré dans **Paramètres**) : textes extraits et, si le modèle le permet, images. N'utilise qu'un LLM autorisé pour les données du client.

## Utilisation
1. Ouvrir <https://benoitmarionpro.github.io/DeliverablesPurifier/>, ou `dist/deliverablespurifier.html` par double-clic (aucune installation, fonctionne hors ligne).
2. **Paramètres** : cocher l'autorisation d'usage de l'IA, choisir le fournisseur (par défaut le proxy local `http://127.0.0.1:11434/v1`, modèle `auto`), tester la connexion.
3. **Anonymiser** : déposer le .pptx, choisir les traitements, relire les images proposées, télécharger `Anonymized_<nom>.pptx`.
4. Relire le résultat : le résumé liste les correspondances appliquées et tout terme détecté encore présent.

Sans autorisation IA : seuls les emails, téléphones, URLs et IBAN sont remplacés ; les images se classent à la main.

## Traitements
| Option | Effet |
|---|---|
| Texte sensible | Clients, personnes, projets, lieux, dates, montants → placeholders (`Bank_Client`, `Person`, `Financial_Data`…). Slides, masters, graphiques, SmartArt, textes alternatifs, noms de formes, propriétés |
| Images | Classification par LLM vision (logo / schéma / confidentiel) puis revue : carré gris + label ou flou |
| Notes, Commentaires | Notes vidées ; commentaires et auteurs (avec leurs emails) supprimés |
| Mise en page | Masters purgés (logos, fonds), couleurs sombres neutralisées, charte Wavestone |
| Vidéos et sons | Supprimés |
| Hyperliens | Liens retirés, texte préfixé `Lien_URL_` |
| Métadonnées | Auteur, société, miniature de la 1re slide, propriétés personnalisées, classeurs Excel des graphiques |

## Proxy et CORS
Le serveur LLM doit autoriser l'origine de la page : `https://benoitmarionpro.github.io` pour la version en ligne, `null` pour le fichier ouvert en double-clic, `http://localhost:8000` en développement. Le proxy CoPro accepte le texte seulement : la revue des images est alors manuelle.

## Développement
`python tools/serve.py` puis `http://localhost:8000` ; tests : `http://localhost:8000/test.html` ; fichier unique : `node tools/build-standalone.mjs`. Voir `CLAUDE.md`. La CI (`.github/workflows/ci.yml`) vérifie la syntaxe, que `dist/` est à jour et lance `test.html` dans Chrome headless. GitHub Pages publie la branche `main` (racine) à chaque push.
