# CLAUDE.md

<!-- TODO: paragraphe de contexte projet — ce que fait l'app une fois le code
     réel démarré, le vocabulaire métier à préserver tel quel (ne pas laisser
     traduire/angliciser des termes établis), pointeur vers un roadmap/docs
     plus profonds (avec un avertissement s'ils peuvent être obsolètes). Ne
     pas deviner à la place de l'équipe — remplir avec des faits une fois le
     premier code écrit. -->

## État actuel (2026-07-16 — terrain préparé avant le code)

Pas encore un projet de code : seulement de la matière de référence.

- `Imports/` — décks/PDF importés (Chantiers OHC "dispositif d'écoute",
  documents RH internes : stress au travail, coaching professionnel,
  personnes ressources) + le mirroir `vscode1-export` habituel
  (`design-system-octo.md`, `optimisation-tokens.md`,
  `points-amelioration-ppt.md`, `ppt-toolkit.md`, `template-octo.md`,
  `claude-code-setup-export.md`).
- `docs/wiki/rh-ecoute.md` — synthèse de la matière RH ci-dessus, produite par
  l'agent `onboarder`. Suggère un thème probable (dispositifs d'écoute RH) —
  **à confirmer par l'équipe**, pas encore une décision de scope.
- `docs/wiki.html` — rendu autonome du wiki.

## Claude Code — préparé en avance (2026-07-16)

- **BMAD-METHOD v6.10.0** installé (`_bmad/`, ~46 skills `bmad-*` →
  `.claude/skills/`). Routeur : `bmad-help`.
- `.gitignore`, `.claude/settings.json` (hook `guard_destructive_git.py` —
  bloque `git push --force`/`git reset --hard`, y compris le cas
  `VAR=value git push --force` — et hook `SessionStart` de rappel), skill
  `revue-increment` (definition-of-done avant commit, délègue à
  `bmad-code-review`/`bmad-retrospective`) posés **avant** le premier code,
  en parité avec les projets frères VSCode/VSCode1/VSCode2/VSCode3.
- ~~Pas encore de `.claude/agents/` custom~~ → voir « Dispositif PPT »
  ci-dessous (2026-07-23).

## Dispositif PPT — sous-agent + skills (greffe VSCode3, 2026-07-23)

Réplication du dispositif deck de VSCode3, adaptée au deck **binaire** de ce
projet (deck OHC « dispositif d'écoute », copies versionnées dans `Exports/` —
cf. mémoire `projet-deck-ohc-ecoute`) :

- `.claude/agents/ppt-designer.md` : sous-agent pilote de l'étape `generation`
  du playbook `export-ppt-verifie` (voie deck unique — `bmad-agent-ux-designer`
  ne double pas ce rôle). Pas de champ `model:` (hérite du thread principal —
  jugement visuel, délibéré, cf. arbitrage VSCode3). Son brief porte les règles
  de sécurité deck binaire : ajouter-avant-supprimer, purge des rels
  orphelines, ouverture PowerPoint COM réelle après chaque save.
- `.claude/skills/` greffées de VSCode3 (tests rejoués après copie : 9/9 et
  9/9) : `pptx-framed-image` (cadres photo du template, used-as-library),
  `slide-text-polish` (lint rédactionnel des slides, used-as-library),
  `deck-design-library` (22 patterns de soutenance OCTO par situation,
  used-as-reference — copie de référence dans VSCode2, resynchroniser
  manuellement). Complètent les skills globales `pptx-deck` / `pptx-verify` /
  `restitution-deck-design`.
- Playbook `export-ppt-verifie` et `catalogue.md` alignés en conséquence
  (génération via sous-agent, étapes conditionnelles toutes routables).

## Générateur PPT versionné (arbitrage superviseur, 2026-07-23)

Le deck OHC était produit par édition in-place via scripts jetables en scratchpad
(pièges pptx rejoués d'un run à l'autre — cf. mémoire
`feedback-suppression-slide-pptx-orphelins`). Diagnostic `agent-supervisor` du
2026-07-23 (`.claude/supervision/diagnostic.json`) + arbitrage utilisateur
(`.claude/supervision/arbitrages.json`) : consolidation d'un outillage versionné,
inspiré des dépôts frères VSCode2 (module helpers) et VSCode3 (forme du
générateur standalone) :

- `scripts/pptx_deck.py` : bibliothèque helpers python-pptx (échelle typo,
  cartes/chips/badges/encarts, double self-check `verifier_geometrie` +
  `verifier_debordements_texte`), reprise de VSCode2 `app/services/pptx_deck.py`
  et complétée d'une section « helpers durcis deck binaire » —
  `trouver_slide_par_titre` (égalité stricte + assertion d'unicité),
  `supprimer_slide`/`clear_slides` (avec `drop_rel`),
  `purger_rels_slides_orphelines`. **Module unique de référence : à importer,
  jamais à redéfinir inline** (le brief `ppt-designer` l'impose).
- `scripts/generate_deck_ohc.py` : générateur standalone (forme VSCode3
  `generate_deck.py` : `slide_*` par slide, `build()`, self-check bloquant) qui
  régénère le deck OHC (15 slides, contenu cartographié sur la v6) en réutilisant
  les layouts natifs du template OCTO. Source = `Exports/… - v6.pptx` (masters/
  layouts/thème/média), jamais modifiée. Sortie : `Exports/… - v7-genere.pptx`.
  Usage : `py scripts/generate_deck_ohc.py`.
- `scripts/test_generate_deck_ohc.py` : 40 assertions (structure, géométrie,
  qualité, rendu réel LibreOffice avec skip propre si absent). Usage :
  `py scripts/test_generate_deck_ohc.py`.
- Linter : `py -m ruff check .` (config `pyproject.toml`, baseline F/I/UP/B).
  Première mesure 2026-07-28 : 8 points, aucun seuil imposé — on mesure d'abord.
  **Jamais `--fix` en aveugle** : sur VSCode2 un `--fix` a supprimé un ré-export et
  cassé un import ; corriger au fil de l'eau puis rejouer le test du générateur.
- Couverture (`requirements-dev.txt` : `coverage` épinglé) — `test_generate_deck_ohc.py`
  est un script autonome (pas des `def test_*` pytest), donc `coverage run`, PAS
  `pytest-cov` (finding VSCode3+VSCode4:exclusion-tests-perimee, l'exclusion « peu de
  code » du 2026-07-23 était déjà démentie côté linter par la mesure du 2026-07-28) :
  `py -m coverage run --source=scripts scripts/test_generate_deck_ohc.py && py -m coverage report`.
  PREMIÈRE MESURE 2026-07-30 : 75 % (generate_deck_ohc.py 96 %, pptx_deck.py 48 % — le
  fork local n'est exercé que par les chemins que le deck OHC emprunte réellement).
  Aucun seuil imposé.

## Agents de pilotage — orchestrateur + superviseur (2026-07-21)

Config d'installation reprise du projet frère VSCode2 (`export/README.md` §7 de ce
dépôt-là), adaptée à l'inventaire réel de VSCode4 (pas de `run-dev-server`, pas d'export
PPT maison, pas de `.opencode/`, BMAD non trié) :

- Skills `agent-orchestrator` (point d'entrée des demandes multi-étapes/multi-agents,
  routé par le hook `UserPromptSubmit`) et `agent-supervisor` (diagnostic qualitatif
  étage 2, sur demande ou signal `SessionStart`) → `.claude/skills/`. Conception :
  `docs/reflexions/agent-orchestrateur.md` / `agent-superviseur.md` (repris tels quels,
  précédents/dates = historique VSCode2, à lire comme référence de méthode).
- `.claude/orchestration/` : `catalogue.md` (adapté — statuts remis à zéro pour ce
  projet), `playbooks/` (`dev-verifie` adapté — `run-dev-server` → skill `run` générique ;
  `export-ppt-verifie` et `revue-design-parallele` repris tels quels ; `cycle-produit-bmad`
  régénéré depuis le CSV BMAD de **ce** projet via `generate_bmad_playbook.py`),
  `log_run.py`, `git_agents_inventory.py`.
- `.claude/supervision/` : `scan_transcripts.py` (étage 1, 0 token, lancé par
  `SessionStart`), `log_usage.py` (`PostToolUse`), `write_diagnostic.py`. Données machine
  (`state.json`, `usage.jsonl`, `diagnostic.json`, `runs.jsonl`, `routing-hints.json`)
  gitignorées, démarrent vides — aucun scan/diagnostic n'a encore tourné sur ce projet.
- 3 hooks ajoutés à `.claude/settings.json` : `SessionStart` (scan superviseur),
  `UserPromptSubmit` (grille de qualification `orchestrator_gate.py`), `PostToolUse`
  matcher `Skill|Agent|Task` (journal d'usage). L'allow-list machine de VSCode2 n'a pas
  été reprise (propre à ce poste-là).
- Gouvernance : le superviseur *propose* (`diagnostic.json`, champ `proposition`),
  l'humain *arbitre* (`.claude/supervision/arbitrages.json`, démarre vide — format
  d'exemple dans `.claude/skills/agent-supervisor/arbitrages.example.json`),
  l'orchestrateur *applique*.

## Discipline de gestion des tokens (cf. `Imports/optimisation-tokens.md`)

Le contexte est un cache actif facturé à chaque tour, pas une mémoire
gratuite (source : OCTO Playbook Agentique, partie « Optimiser la
consommation Tokens »). Règles concrètes, pas de changement de ton/style de
réponse :

- **Ne pas parcourir** `_bmad/`, `_bmad-output/`, `.venv/`, `node_modules/`,
  `__pycache__/`, `.git/` sauf demande explicite.
- **Lire avant d'écrire, grep les appelants avant de modifier** une fois du
  code réel en place.
- **Préférer un grep/read ciblé à un dump récursif** — surtout sur
  `.claude/skills/bmad-*` (~46 skills).
- **Sous-agent pour toute sortie volumineuse** plutôt que de la laisser
  polluer le contexte principal.
- **`/compact` dès ~40 %** de fenêtre de contexte utilisée si la session doit
  continuer longtemps sur le même sujet.
