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
