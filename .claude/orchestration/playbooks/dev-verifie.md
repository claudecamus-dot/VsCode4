# Playbook `dev-verifie` — implémentation vérifiée de bout en bout

> Repris du projet frère VSCode2 le 2026-07-21 et adapté : l'agent de vérification UI
> `run-dev-server` (spécifique à la stack FastAPI/Jinja2 de VSCode2) est remplacé par le
> skill générique `run` (disponible dans ce projet) ; les noms de fichiers PPT propres à
> VSCode2 (`pptx_export.py`/`pptx_deck.py`) sont généralisés en « export PPT ». Le
> précédent « pratique effective de tous les incréments livrés » cité ci-dessous est celui
> de VSCode2 — ce projet est pré-code et n'a pas encore son propre historique
> d'incréments.

Le workflow de dev quotidien du projet, rendu structurel : implémenter, tester, **vérifier
en réel** (pas seulement pytest vert), puis boucle de definition-of-done avant tout commit.
Précédent (VSCode2) : c'est la pratique effective de tous les incréments livrés de ce
projet frère (statut `eprouve` là-bas).

Les étapes de vérification réelle sont **conditionnelles au type de fichiers touchés**
(table des vérifications obligatoires de la skill) : ne garder à l'instanciation que
celles dont la condition s'applique, ne jamais retirer `pytest` ni `revue-increment`.

Frontière avec `export-ppt-verifie` : un changement de code qui *touche* l'export PPT au
passage reste ici (l'étape `verification-pptx` couvre) ; quand le **livrable est le deck
lui-même** (layout, contenu, visuel), préférer `export-ppt-verifie` qui déroule la chaîne
PPT complète (cadres photo, polish, passe design).

```json
{
  "nom": "dev-verifie",
  "description": "Implémentation d'une feature/correction avec tests, vérification réelle adaptée aux fichiers touchés, et revue-increment avant commit.",
  "statut": "eprouve",
  "source": "manuel",
  "declencheurs": [
    "implémente/corrige/ajoute une fonctionnalité",
    "changement de template/CSS/JS d'une interface web",
    "changement de l'export PPT (python-pptx)",
    "fin d'incrément, préparation d'un commit de code produit"
  ],
  "etapes": [
    {
      "id": "cadrage",
      "agent": "session principale",
      "mode": "cascade",
      "modele": "(session)",
      "contrat": {
        "type": "deterministe",
        "critere": "fichiers concernés lus, appelants des fonctions/champs partagés grep-és avant modification"
      },
      "checkpoint": false
    },
    {
      "id": "implementation",
      "agent": "session principale",
      "mode": "cascade",
      "modele": "(session)",
      "contrat": {
        "type": "deterministe",
        "critere": "chaque exigence EXPLICITE de la demande (points numérotés, contraintes) cochée une à une contre le diff — pas seulement « ça compile/passe » ; toute exigence réinterprétée ou écartée signalée, jamais silencieuse ; style du fichier environnant respecté (pas de linter configuré)"
      },
      "checkpoint": false
    },
    {
      "id": "tests",
      "agent": "session principale",
      "mode": "cascade",
      "modele": "(session)",
      "contrat": {
        "type": "deterministe",
        "critere": "verdict lu sur la ligne de synthèse RÉELLE de la suite de tests (N passed / 0 failed / 0 error) — jamais sur un résumé filtré ni une sortie tronquée ; en cas de doute, relancer et/ou rediriger toute la sortie dans un fichier",
        "commande": "pytest -q"
      },
      "checkpoint": false
    },
    {
      "id": "verification-ui",
      "agent": "run",
      "mode": "cascade",
      "modele": "(session)",
      "contrat": {
        "type": "reel",
        "critere": "SI template/CSS/JS d'une interface web touché : screenshot de la page modifiée pris et regardé"
      },
      "checkpoint": false
    },
    {
      "id": "verification-pptx",
      "agent": "pptx-verify",
      "mode": "cascade",
      "modele": "(session)",
      "contrat": {
        "type": "reel",
        "critere": "SI un export PPT (python-pptx) est touché : export réel rendu en images et inspecté (python-pptx est un parseur tolérant)"
      },
      "checkpoint": false
    },
    {
      "id": "revue-increment",
      "agent": "revue-increment",
      "mode": "cascade",
      "modele": "(session)",
      "contrat": {
        "type": "reel",
        "critere": "boucle revue + application des correctifs + re-vérification réelle exécutée en entier"
      },
      "checkpoint": "avant tout commit — action difficilement réversible, proposer, ne pas exécuter unilatéralement"
    }
  ],
  "regle_reprise": "une relance ciblée par étape en échec de contrat, puis escalade utilisateur avec l'état réel"
}
```
