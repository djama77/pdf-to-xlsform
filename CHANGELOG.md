# Changelog

## 0.1.0
- Extraction du moteur XLSForm dans un paquet Python autonome.
- Correction de la détection des options numérotées.
- Tests de régression ajoutés.
- OCR et reconstruction de la logique conditionnelle laissés comme travaux futurs explicites.


## 0.1.1 — Renforcement de la validation
- Validation stricte des noms de variables : format ASCII lower_snake_case, longueur maximale et mots réservés.
- Détection des doublons de noms sans distinction de casse.
- Contrôle des types select_one / select_multiple, des références de listes et des listes inutilisées.
- Vérification du format de form_id et rejet des entrées autres que .xlsx par le validateur.
- Ajout de tests de régression ciblant les cas limites du validateur.
- Aucun téléversement de questionnaire ni de données de collecte.
