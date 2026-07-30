# VSCode4 — deck OHC « dispositif d'écoute »

Le livrable de ce dépôt est un deck PowerPoint binaire (`Exports/`, 15 slides)
restituant les plans d'actions du dispositif d'écoute RH (chantiers OHC).

## Utilisation

Depuis `scripts/` :

```bash
# régénérer le deck (source = Exports/…v6.pptx, jamais modifiée ; sortie = …v7-genere.pptx)
py scripts/generate_deck_ohc.py

# vérifier après toute modification du générateur
py scripts/test_generate_deck_ohc.py
```

`test_generate_deck_ohc.py` vérifie structure, géométrie, qualité (aucune ombre
portée, aucune police générique, aucune relation orpheline) et rendu réel
LibreOffice (conversion PDF + comptage de pages) — SKIP propre si LibreOffice
est absent. La vérification de référence reste l'ouverture PowerPoint COM et
le rendu PNG, faits en session.

```bash
py -m ruff check .                                          # linter, baseline mesurée
py -m coverage run --source=scripts scripts/test_generate_deck_ohc.py \
  && py -m coverage report                                   # couverture, première mesure 75 %
```

Détail complet des commandes, contraintes du deck et gotchas connus : voir
[`CLAUDE.md`](CLAUDE.md).

## Documentation

- Règles et contraintes du dépôt : [`CLAUDE.md`](CLAUDE.md).
- Cadrage produit : [`docs/product-brief.md`](docs/product-brief.md).
- Tableau de bord de supervision (synchronisé depuis le hub) :
  [`docs/wiki/index.md`](docs/wiki/index.md) ou `docs/wiki.html`.
