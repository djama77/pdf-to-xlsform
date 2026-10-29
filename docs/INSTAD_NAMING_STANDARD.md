# Référentiel INSTAD — Nommage et documentation des variables

## Principe directeur

Le nom de variable identifie la donnée ; le label explique la donnée ; le dictionnaire documente la donnée.

## Règles obligatoires pour les propositions générées

1. Utiliser des noms courts mais explicites en `lower_snake_case`.
2. Éliminer espaces, accents et caractères spéciaux ; utiliser des caractères ASCII compatibles.
3. Viser 32 caractères maximum pour favoriser la portabilité vers Stata et les outils analytiques.
4. Assurer l'unicité des noms et ajouter un suffixe contrôlé si nécessaire sans dépasser la longueur cible.
5. Garder les noms stables après validation et déploiement. Tout changement ultérieur doit être documenté et versionné.
6. Conserver le libellé source complet séparément du nom technique.
7. Ne pas utiliser le même suffixe pour des concepts différents sans préfixe ou contexte explicite.
8. Préserver la traçabilité avec la référence de question/module ; ne pas inventer de numéro de page si l'extraction ne le fournit pas.
9. Documenter chaque variable avec type, libellé, modalités/codes, conditions d'affichage, unité si applicable, définition, source et statut de validation.
10. Ne pas interpréter un nom généré comme approuvé par le métier. Le statut initial est `À VALIDER`.

## Exemples de référence

| Question / concept | Nom proposé | Note |
|---|---|---|
| Nombre de chameaux actuellement possédés | `pa20_cham_act` | Exemple du guide, à confirmer selon la question source |
| Situation matrimoniale | `pa09_stat_mat` | Exemple du guide |
| Évolution | `pa21_evol` | Le dictionnaire doit préciser la définition et la période |

Ces exemples illustrent le standard et ne doivent pas être injectés dans un questionnaire sans correspondance réelle avec sa source.

## Dictionnaire minimal

Chaque variable doit contenir :

- `source_ref` : numéro de question/module d'origine ;
- `variable_name` : nom technique proposé ;
- `label_original` : libellé source complet ;
- `data_type` : type XLSForm proposé ;
- `modalities_codes` : modalités et codes si disponibles ;
- `display_condition` : condition documentée ou mention explicite qu'elle n'a pas été extraite ;
- `unit` : unité lorsqu'elle est indiquée dans la source ;
- `definition` : définition et notes ;
- `validation_status` : `PASS`, `À VALIDER`, `BLOQUÉ` ou `NON TESTÉ`.

## Limites de l'automatisation

Le moteur actuel utilise une extraction heuristique. Il ne peut pas garantir la couverture complète, les sauts, les contraintes, les calculs, les groupes/répétitions ni l'interprétation de tableaux complexes. Les feuilles de traçabilité et QA signalent ces limites ; elles ne remplacent pas la comparaison avec le PDF/Word, la validation XLSForm sur la plateforme cible et les tests pilote.
