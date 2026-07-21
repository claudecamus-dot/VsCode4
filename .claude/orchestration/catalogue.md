# Catalogue des agents — routage orchestrateur

> Repris du projet frère VSCode2 le 2026-07-21 (config d'installation d'`agent-orchestrator`
> / `agent-supervisor`) et **adapté à l'inventaire réel de ce projet** — voir
> `export/README.md` de VSCode2 §7 : « adapter le catalogue, ses recommandations citent des
> skills spécifiques au projet source ». VSCode4 est encore **sans code produit**
> (cf. CLAUDE.md « État actuel ») : pas de `run-dev-server`, pas de `pptx_export.py`/
> `pptx_deck.py` maison, pas de `.opencode/`. Ce catalogue démarre donc avec les seuls
> agents/skills réellement installés ici ; à étoffer au fil des skills projet créées.
>
> Utilisé par la skill `agent-orchestrator` pour composer ses plans. Descriptions et
> recommandations maintenues à la main ; les **statuts d'usage vivants** (invocations,
> dates, jamais-utilisés) sont dans `routing-hints.json` (généré à chaque session par le
> scan superviseur, avec les stats plan-vs-réel de `runs.jsonl`) et, en version lisible,
> dans `docs/wiki/technical/agents-supervision.md` — toujours les vérifier avant de router
> vers un agent « jamais utilisé ». Statuts ci-dessous : instantané du 2026-07-21, projet
> tout juste initialisé (aucun scan n'a encore tourné).
> Les décisions humaines qui closent un constat d'usage sont dans
> `.claude/supervision/arbitrages.json` (démarre vide ici — voir l'exemple de format dans
> `.claude/skills/agent-supervisor/arbitrages.example.json`).
> Si **aucune entrée ne couvre le besoin** : inventaire git présents + supprimés via
> `py .claude/orchestration/git_agents_inventory.py`, puis proposition de
> restauration/évolution/création (procédure dans la skill, étape 2).
> Conception : `docs/reflexions/agent-orchestrateur.md`.

## Skills projet

| Skill | Quand l'utiliser | Mode typique | Modèle | Statut |
| --- | --- | --- | --- | --- |
| `revue-increment` | Definition-of-done : fin d'incrément, avant commit | Synchrone, étape terminale obligatoire des plans de dev | (session) | Installée, pas encore invoquée dans ce projet |
| `agent-orchestrator` | Point d'entrée des demandes multi-étapes/multi-agents (routé par le hook UserPromptSubmit) | Synchrone | (session) | Installée à l'instant (ce run) |
| `agent-supervisor` | Diagnostic qualitatif des agents (étage 2) — depuis `revue-increment` ou sur signal SessionStart | Synchrone, ≤ 1×/14 j | (session) | Installée à l'instant (ce run) — aucun diagnostic encore produit |

> Pas encore de skills projet spécifiques (ex. lancement d'app, export d'un livrable
> métier) : ce projet est pré-code. Les ajouter ici dès leur création, avec leur
> déclencheur et le playbook qui les invoque.

## Skills globaux clés

| Skill | Quand l'utiliser | Mode typique | Modèle | Statut |
| --- | --- | --- | --- | --- |
| `roadmap-keeper` | Créer/mettre à jour/rendre la roadmap projet | Synchrone | (session) | Disponible |
| `pptx-deck` / `pptx-verify` | Générer un deck PowerPoint / le vérifier en rendu réel — toujours en paire | Synchrone, toujours en paire | (session) | Disponibles |
| `restitution-deck-design` | Deck techniquement correct mais visuellement pauvre | Synchrone | (session) | Disponible |
| `run` | Lancer/screenshoter l'app pour vérifier un changement UI réel (remplace le `run-dev-server` propre à VSCode2 — équivalent générique, cherche d'abord une skill projet dédiée) | Synchrone | (session) | Disponible |
| `dataviz`, `skill-creator`, `update-config`, `code-review` / `verify` / `simplify` | Voir description de chaque skill | Synchrone | (session) | Builtins/globaux disponibles |

> `pptx-framed-image` et `slide-text-polish` (enrichissements PPT conditionnels du
> playbook `export-ppt-verifie` sur VSCode2) **ne sont pas installées dans ce projet** —
> le playbook les référence quand même (repris tel quel, cf. sa note de provenance) mais
> l'orchestrateur ne doit **pas** router vers elles tant qu'elles ne sont pas ajoutées ici ;
> traiter l'étape correspondante comme non applicable jusqu'à leur éventuelle installation.

## Sous-agents (seuls à accepter un choix de modèle)

| Sous-agent | Quand l'utiliser | Mode typique | Modèle conseillé | Statut |
| --- | --- | --- | --- | --- |
| `Explore` | Recherche large en lecture seule, conclusion sans les dumps | Parallèle (fan-out ≤4) ou async | Haiku/Sonnet (mécanique/standard) | Disponible |
| `Plan` | Concevoir une stratégie d'implémentation | Synchrone | Opus/Fable (structurant) | Disponible |
| `general-purpose` | Tâche multi-étapes déléguée, sortie volumineuse | Async ou synchrone | Sonnet ; Opus/Fable si structurant | Disponible |
| `claude-code-guide` | Questions sur Claude Code / SDK / API | Synchrone | (défaut) | Disponible |

## Familles sous condition

| Famille | Règle de routage |
| --- | --- |
| **BMAD (~46 skills `bmad-*`, aucun tri effectué dans ce projet)** | Contrairement à VSCode2 (tri exécuté le 2026-07-18, 39/46 conservés), **aucun tri BMAD n'a été fait ici** — les ~46 skills sont toutes installées et non triées. Ne router que sur demande explicite de l'utilisateur, en passant par `bmad-help`. Un futur tri (s'il a lieu) devra être noté en arbitrage (`famille:BMAD`) comme sur les projets frères. |

> OpenHub (`.opencode/`) : absent de ce projet — pas de ligne de routage, contrairement à
> VSCode2 où l'intégration existe mais est hors périmètre.
>
> Angle mort de mesure (constat superviseur VSCode2, transposable ici) : les sous-skills
> invoquées par un sous-agent via un prompt en langage naturel (pattern utilisé par
> `bmad-code-review` pour lancer `bmad-review-adversarial-general`/
> `bmad-review-edge-case-hunter`) n'apparaissent pas dans `state.json`/`routing-hints.json`
> — seules les invocations directes de la session principale sont trackées. Une absence de
> trace sur ces sous-skills ne signifie donc pas absence d'exécution : ne pas les qualifier
> `agent-mort` sur cette seule base.

## Playbooks

Workflows récurrents pré-composés — la skill cherche un playbook matchant **avant** de
composer à vide. Format : `.claude/orchestration/playbooks/FORMAT.md`.

| Playbook | Quand | Source | Statut |
| --- | --- | --- | --- |
| `dev-verifie` | Dev/correction : tests + vérif réelle (conditionnelle aux fichiers touchés) + `revue-increment` avant commit | Manuel (repris de VSCode2, adapté : `run-dev-server` → `run`) | Repris, jamais rejoué dans ce projet |
| `export-ppt-verifie` | Livrable = le deck : génération (`pptx-deck`) + enrichissements conditionnels + `pptx-verify` obligatoire + `revue-increment` | Manuel (repris tel quel de VSCode2) | Repris, jamais rejoué dans ce projet — 2 des 3 étapes conditionnelles référencent des skills non installées ici (voir note ci-dessus) |
| `revue-design-parallele` | Revue multi-angles en fan-out d'`Explore` (≤4) puis consolidation | Manuel (repris tel quel de VSCode2) | Repris, jamais rejoué dans ce projet |
| `cycle-produit-bmad` | Cycle produit BMAD (brief→PRD→archi→epics→dev→review), clos par `revue-increment` | `generate_bmad_playbook.py` (régénéré depuis le CSV BMAD de **ce** projet) | Jamais joué — sur demande explicite |
