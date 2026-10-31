# Feuille de route

## Phase 1 — socle local
- [x] Extraction du texte PDF, DOCX et TXT.
- [x] Export XLSForm avec les feuilles `survey`, `choices` et `settings`.
- [x] Tests de régression de base.
- [x] Interface en ligne de commande.
- [x] Intégration continue avec pytest.

## Phase 2 — fiabilité de la conversion
- [ ] Ajouter des exemples de questionnaires anonymisés et des tests de référence.
- [ ] Améliorer la détection des questions numérotées sans confondre les modalités.
- [ ] Ajouter une sortie de rapport de conversion lisible par un non-développeur.
- [x] Contrôler les noms, types, références de listes et doublons XLSForm.
- [x] Ajouter des tests ciblés sur les noms, doublons et erreurs structurelles XLSForm.\n- [ ] Ajouter des jeux de référence anonymisés, des tests de fidélité et des cas DOCX/PDF multipages.

## Phase 3 — fonctionnalités avancées
- [ ] OCR local optionnel pour PDF numérisés.
- [ ] Reconnaissance explicite des types de questions (entier, décimal, date, géopoint, etc.).
- [ ] Aide à la reconstruction de `relevant`, groupes et répétitions, avec confirmation humaine.
- [ ] Validation optionnelle avec pyxform/KoboToolbox.
- [ ] Interface graphique locale, après stabilisation du moteur.

## Principes
- Priorité au traitement local et à la confidentialité.
- Ne jamais téléverser automatiquement les questionnaires.
- Toute reconstruction incertaine doit être signalée et vérifiée par un humain.
