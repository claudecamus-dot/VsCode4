---
updated: 2026-07-23
generated-by: .claude/supervision/scan_transcripts.py (superviseur d'agents, étage 1)
---

# Supervision des agents — tableau de bord d'usage

> ⚠️ **Page générée automatiquement** (hook SessionStart → `.claude/supervision/scan_transcripts.py`).
> **Ne pas éditer à la main** — toute modification serait écrasée au prochain scan.
> Conception et phasage : [../../reflexions/agent-superviseur.md](../../reflexions/agent-superviseur.md).

Dernier scan : 2026-07-23T18:06:24+02:00 · **8 sessions** (transcripts) · **12** invocations de skills · **5** lancements de sous-agents.

## Skills — usage réel

| Skill | Famille | Invocations | Première | Dernière |
| --- | --- | --- | --- | --- |
| `pptx-deck` | global | 3 | 2026-07-09 | 2026-07-21 |
| `agent-orchestrator` | projet | 2 | 2026-07-21 | 2026-07-23 |
| `agent-supervisor` | projet | 2 | 2026-07-23 | 2026-07-23 |
| `revue-increment` | projet | 2 | 2026-07-21 | 2026-07-23 |
| `artifact-design` | (builtin/session) | 1 | 2026-07-09 | 2026-07-09 |
| `bmad-correct-course` | BMAD | 1 | 2026-07-23 | 2026-07-23 |
| `pptx-verify` | global | 1 | 2026-07-21 | 2026-07-21 |

## Sous-agents

| Sous-agent | Lancements | Premier | Dernier |
| --- | --- | --- | --- |
| `ppt-designer` | 3 | 2026-07-23 | 2026-07-23 |
| `Explore` | 1 | 2026-07-23 | 2026-07-23 |
| `general-purpose` | 1 | 2026-07-09 | 2026-07-09 |

## Jamais utilisés

**projet** — 3/6 jamais invoqués :

`deck-design-library`, `pptx-framed-image`, `slide-text-polish`

**BMAD** — 45/46 jamais invoqués :

<details><summary>Voir la liste</summary>

`bmad-advanced-elicitation`, `bmad-agent-analyst`, `bmad-agent-architect`, `bmad-agent-dev`, `bmad-agent-pm`, `bmad-agent-tech-writer`, `bmad-agent-ux-designer`, `bmad-architecture`, `bmad-brainstorming`, `bmad-check-implementation-readiness`, `bmad-checkpoint-preview`, `bmad-code-review`, `bmad-create-architecture`, `bmad-create-epics-and-stories`, `bmad-create-prd`, `bmad-create-story`, `bmad-customize`, `bmad-dev-auto`, `bmad-dev-story`, `bmad-document-project`, `bmad-domain-research`, `bmad-edit-prd`, `bmad-editorial-review-prose`, `bmad-editorial-review-structure`, `bmad-forge-idea`, `bmad-generate-project-context`, `bmad-help`, `bmad-index-docs`, `bmad-market-research`, `bmad-party-mode`, `bmad-prd`, `bmad-prfaq`, `bmad-product-brief`, `bmad-qa-generate-e2e-tests`, `bmad-quick-dev`, `bmad-retrospective`, `bmad-review-adversarial-general`, `bmad-review-edge-case-hunter`, `bmad-shard-doc`, `bmad-spec`, `bmad-sprint-planning`, `bmad-sprint-status`, `bmad-technical-research`, `bmad-ux`, `bmad-validate-prd`

</details>

**global** — 3/5 jamais invoqués :

`restitution-deck-design`, `roadmap-keeper`, `skill-creator`

## TODO agents (constats automatiques)

1. **Élaguer les skills BMAD** : 45/46 jamais invoqués — confirmer l'utilité des non-utilisés.
2. **Skills projet sans usage** : `deck-design-library`, `pptx-framed-image`, `slide-text-polish` — vérifier pertinence et déclencheurs.

## Arbitrages enregistrés

_Constats clos par décision humaine (`.claude/supervision/arbitrages.json`) — l'usage réel reste mesuré ci-dessus._

- **`inline python-pptx`** (2026-07-23) : Propositions 1+2 du diagnostic 2026-07-23 retenues : module helpers versionné scripts/pptx_deck.py repris de VSCode2 app/services/pptx_deck.py + générateur standalone régénérable du deck OHC (forme VSCode3 generate_deck.py), helpers durcis (find par titre + assert unicité, clear_slides avec drop_rel, ajouter-avant-supprimer), brief ppt-designer amendé pour importer le module.
- **`pptx-verify`** (2026-07-23) : Proposition 3 du diagnostic 2026-07-23 retenue : test rendu réel versionné (LibreOffice page-count + assertions qualité adaptées OHC, modèle VSCode3 test_generate_deck.py + VSCode2 test_deck_qualite.py) ; l'ouverture COM PowerPoint reste la vérification finale.
- **`export-ppt-verifie`** (2026-07-23) : Proposition 4 du diagnostic 2026-07-23 REPORTÉE : greffe de deck-design-review / priority-matrix (VSCode2) à re-proposer après mise en place du générateur.

## Diagnostic qualitatif (étage 2 — `agent-supervisor`)

_Diagnostic ⚠️ à relancer (> 14 j)._

1. **Contournement du cadre photo des dividers de chapitre jamais re-questionné, malgré l'écart documenté au pattern VSCode3 que le dispositif est censé répliquer** — Le sous-agent a vérifié la propreté géométrique et l'ouverture COM des dividers, mais jamais leur fidélité au pattern de référence qu'il était censé reproduire — un contournement (vider le cadre) a été substitué silencieusement à la résolution du problème réel (collision numéro/logo) sans le signaler comme un écart de design à l'utilisateur. · **Proposition** : Ajouter au brief ppt-designer.md (section Design principles) une règle explicite : quand une étape porte une réplication d'un pattern d'un projet frère (VSCode2/VSCode3), le contrat de vérification doit comparer le résultat au comportement RÉEL du pattern source (pas seulement l'auto-check géométrique du propre résultat) — tout écart délibéré (contournement plutôt que résolution) doit être signalé explicitement à l'utilisateur dans le rapport, jamais appliqué en silence. Correctif déjà en cours ce jour (run explicitement demandé par l'utilisateur).

---

_Étage O-C (croisement modèle × tâche × reprises, exploitation de `runs.jsonl`) : voir `.claude/orchestration/routing-hints.json`, régénéré à chaque session._
