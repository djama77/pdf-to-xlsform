# Changelog

## 0.1.0
- Extraction du moteur XLSForm dans un paquet Python autonome.
- Correction de la détection des options numérotées.
- Tests de régression ajoutés.
- OCR et reconstruction de la logique conditionnelle laissés comme travaux futurs explicites.


## 0.1.1 — qualité et traçabilité
- Renforcement de la validation structurelle XLSForm : noms, casse, modalités, types de sélection et listes inutilisées.
- Ajout de tests de non-régression sur les cas limites du validateur.
- Protection supplémentaire des dossiers locaux et fichiers de collecte dans Git.
- Aucune reconstruction automatique de logique métier ; revue humaine toujours requise.
