# VSCode4 — Index Wiki

Projet pré-code (cf. `CLAUDE.md` « État actuel ») : pas encore de stack ni d'architecture
à documenter. Contenu existant : [rh-ecoute.md](rh-ecoute.md) (synthèse des documents RH
internes, produite par l'agent `onboarder` — thème probable du projet, à confirmer par
l'équipe). Rendu autonome : [../wiki.html](../wiki.html).

<!-- TODO-AGENTS:START — section générée par .claude/supervision/scan_transcripts.py, ne pas éditer à la main -->
## TODO agents 🤖

Constats automatiques du superviseur d'agents (usage mesuré dans les transcripts de session) :

- **Trier les skills BMAD** : 46 installés, 0 invocation à ce jour — décider lesquels garder, customiser ou désinstaller.
- **`revue-increment` jamais invoquée** malgré le rappel SessionStart à chaque session — revoir son déclencheur (l'ancrer au flux de commit ?) ou la simplifier.
- **Skills projet sans usage** : `agent-orchestrator`, `agent-supervisor` — vérifier pertinence et déclencheurs.

Tableau de bord complet : [technical/agents-supervision.md](technical/agents-supervision.md) — régénéré à chaque session.
<!-- TODO-AGENTS:END -->
