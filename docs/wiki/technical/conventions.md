# Conventions de code — VSCode4 (deck OHC)

Rédigé le 2026-07-30 en complément de `CLAUDE.md` (canal, commandes, contraintes
durables du deck) — ce fichier porte le « comment coder ».

## Linting

`pyproject.toml` (ruff, F/I/UP/B, ligne 100) depuis le 2026-07-28. Lancer :
`py -m ruff check .`. Baseline mesurée : 8 points, aucun seuil imposé, jamais
de `--fix` aveugle (leçon VSCode2 2026-07-23). CI en mode informatif
(`.github/workflows/ci.yml`, ajoutée 2026-07-30) — le lint n'y est pas un gate,
voir le commentaire du workflow pour pourquoi la CI ne peut pas exécuter le
test de génération réel (`Exports/` non versionné).

## Nommage

- **Python (`scripts/*.py`)** : `snake_case` pour fonctions/variables,
  fonctions `slide_*` pour chaque slide du deck (`slide_couverture`,
  `slide_sommaire`…), constantes de charte en `MAJUSCULES` (`NAVY`, `CYAN`,
  `SLATE`…).
- **Garde-fous explicites** : préfixe `_exiger_*` pour une garde qui lève
  `SystemExit` avec un message actionnable (`_exiger_source`, ajoutée
  2026-07-30) — convention reprise du projet frère VSCode3
  (`_exiger_template`).
- **Langue** : identifiants et docstrings en français, vocabulaire du domaine
  PowerPoint conservé en anglais (`shape`, `placeholder`, `layout`).

## Duplication assumée

`pptx_deck.py` est une copie explicite du module VSCode2 (voir son propre
docstring) — dette connue, mesurée par la matrice de divergence du hub de
supervision, pas un oubli.

## Git

Ce dépôt est une **cible** du hub de supervision `VScode5` — exclure
systématiquement le churn de données générées avant tout commit (voir
`CLAUDE.md`). `Exports/` (le deck binaire et ses versions) n'est **pas**
versionné : contenu RH interne réel, décision antérieure à ce fichier.
