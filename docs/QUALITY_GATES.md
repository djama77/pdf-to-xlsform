# Contrôles qualité, sauvegarde et traçabilité

## Règle de sauvegarde
- La branche `backup/pre-ecosystem-fix-20261031` conserve l'état initial de `main` avant les corrections.
- Les corrections sont isolées dans `fix/xlsform-quality-traceability-20261031`.
- La branche principale n'est pas modifiée par cette intervention.

## Données
- Le traitement reste local ; aucun questionnaire n'est téléversé automatiquement.
- Les dossiers locaux `input/`, `outputs/`, `data/` et les fichiers DOCX/CSV/XLSX/PDF ne doivent pas être commités.
- Les tests doivent utiliser des données synthétiques ou anonymisées.

## Portes de validation
1. Exécution de toute la suite pytest sur Python 3.10, 3.11 et 3.12.
2. Vérification des feuilles `survey`, `choices`, `settings`.
3. Vérification des colonnes obligatoires, des noms de variables, des doublons et des références de listes.
4. Revue humaine de la fidélité au questionnaire original.
5. Validation sur la plateforme cible et test pilote avant toute collecte réelle.

## Traçabilité
- Chaque correction doit être associée à un commit et à un résultat de test.
- Les noms de variables proposés restent au statut `À VALIDER` jusqu'à approbation métier.
- Les sauts, contraintes, calculs, groupes et répétitions non extraits ne doivent jamais être inventés.
- Aucun merge vers `main` sans revue du diff et réussite des contrôles CI.
