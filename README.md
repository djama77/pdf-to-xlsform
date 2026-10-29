# PDF to XLSForm

Convertisseur Python local de questionnaires **PDF, DOCX et TXT** en classeurs XLSForm utilisables comme base de travail pour **KoboToolbox** et **KoboCollect**.

## État actuel

Le dépôt contient un premier moteur autonome extrait d'un projet existant, avec des tests de régression. Il s'agit encore d'un prototype : la sortie doit être vérifiée avant toute collecte réelle.

- Extraction de texte sélectionnable depuis les PDF avec PyMuPDF.
- Lecture des paragraphes et tableaux DOCX.
- Lecture des fichiers TXT en UTF-8.
- Détection heuristique de questions et de modalités de réponse.
- Génération d'un fichier Excel comportant les feuilles `survey`, `choices` et `settings`.
- Validation structurelle légère et avertissements pour les questionnaires incomplets.

### Limites connues

- Les PDF numérisés sans couche texte sont détectés, mais l'OCR local n'est pas encore intégré.
- La logique conditionnelle, les groupes/répétitions, les matrices, les calculs et les contraintes ne sont pas reconstruits automatiquement.
- La détection des questions dépend de la mise en page et de la formulation ; une revue humaine est indispensable.
- Avant déploiement, validez le fichier produit avec les outils de validation XLSForm/KoboToolbox.

## Installation

Python 3.10 ou plus récent est recommandé.

```bash
python -m venv .venv
```

Sous Windows PowerShell :

```powershell
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m pip install -e ".[test]"
```

## Exemple d'utilisation

```python
from pathlib import Path
from pdf_to_kobocollect import extract_text, parse_questionnaire, to_xlsform, validate_items

text = extract_text(Path("questionnaire.pdf"))
items = parse_questionnaire(text)

print(validate_items(items))
result = to_xlsform(
    items,
    Path("questionnaire.xlsx"),
    title="Enquête terrain",
    form_id="enquete_terrain",
)
print(result)
```

### Utilisation en ligne de commande

Après installation du package :

```powershell
pdf-to-xlsform questionnaire.pdf --output questionnaire.xlsx --title "Enquête terrain" --form-id enquete_terrain
```

Ou sans installer la commande :

```powershell
python -m pdf_to_kobocollect questionnaire.docx -o questionnaire.xlsx
```

Les formats pris en charge sont PDF, DOCX et TXT. Le moteur travaille localement et n'envoie pas les documents à une API.

## Tests

Une intégration continue GitHub Actions exécute les tests sous Python 3.10, 3.11 et 3.12 à chaque push et pull request.

```bash
pytest -q
```

## Confidentialité et données

Les fichiers de questionnaire et les classeurs générés ne doivent pas être commités dans Git. Le fichier `.gitignore` exclut notamment les fichiers PDF et XLSX. Les dépendances doivent être installées avant une utilisation hors ligne.

## Origine et références

Le moteur a été isolé à partir du dépôt [StatisticiEnStad-Data-Analytics](https://github.com/djama77/StatisticiEnStad-Data-Analytics). Des projets de référence complémentaires sont [danzgabrielgabuat/pdf-to-xlsform](https://github.com/danzgabrielgabuat/pdf-to-xlsform), [Adamact/DTOCR](https://github.com/Adamact/DTOCR) et [XLSForm/pyxform](https://github.com/XLSForm/pyxform). Leur code n'est pas copié dans ce dépôt.

## Licence

Consultez le fichier `LICENSE` pour les conditions de licence du dépôt.


## Validation du fichier généré

Après la conversion, le programme contrôle automatiquement la structure du classeur XLSForm :
- présence des feuilles `survey`, `choices` et `settings` ;
- présence des colonnes obligatoires ;
- champs vides et noms de variables de question en double ;
- doublons de modalités au sein d'une même liste ;
- références `select_one` / `select_multiple` vers des listes de choix inexistantes ;
- présence du titre et de l'identifiant du formulaire.

Cette vérification structurelle ne remplace pas une validation complète par pyxform/KoboToolbox ni la relecture humaine. La logique de saut, les contraintes et la fidélité au document source doivent être vérifiées avant une collecte réelle.


## Référentiel INSTAD XLSForm SUPERSKILL V3 — nomenclature et traçabilité

Le générateur applique une proposition de nomenclature fondée sur les principes INSTAD : noms courts, descriptifs, uniques, stables, traçables et en `lower_snake_case`, sans accents ni espaces, avec un objectif de 32 caractères maximum pour les variables. Les noms réservés/problématiques sont évités autant que possible. Les noms générés restent des **propositions à valider**, jamais une validation métier automatique.

- Le libellé original est conservé dans `survey.label` et dans le dictionnaire.
- Les noms proposés sont normalisés sans accents et désambiguïsés avec un suffixe, dans la limite de longueur.
- Le classeur inclut `variable_dictionary`, `source_traceability`, `decisions`, `logic_register`, `qa_report` et `pilot_tests` en plus des feuilles XLSForm principales.
- Les ambiguïtés, modalités non détectées, logique métier, page/section source et compatibilité plateforme doivent être vérifiées. Le moteur ne prétend pas reconstruire les sauts, contraintes ou calculs absents du texte extrait.
- Une variable validée ne doit pas être renommée après déploiement sans décision documentée et gestion de version.

Les statuts utilisés sont `PASS`, `À VALIDER`, `BLOQUÉ` et `NON TESTÉ`. Une sortie XLSX n'est pas une déclaration de préparation PROD : validation métier, validation XLSForm et tests pilote restent obligatoires.

Voir aussi [docs/INSTAD_NAMING_STANDARD.md](docs/INSTAD_NAMING_STANDARD.md).
