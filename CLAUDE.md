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
- Pas encore de `.claude/agents/` custom : à créer si le projet en a besoin,
  pas de flotte à dupliquer ici pour l'instant.

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
