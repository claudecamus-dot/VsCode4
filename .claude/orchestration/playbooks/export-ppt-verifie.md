# Playbook `export-ppt-verifie` — travaux sur le deck de restitution, vérifiés au rendu réel

> Repris de VSCode2 le 2026-07-21, aligné sur le dispositif VSCode3 le 2026-07-23 :
> `pptx-framed-image`, `slide-text-polish` et `deck-design-library` sont désormais
> **installées dans ce projet** (greffe VSCode3, tests rejoués 9/9 et 9/9) et l'étape
> `generation` s'instancie via le **sous-agent `ppt-designer`** (`.claude/agents/`).
> `pptx-deck`, `pptx-verify` et `restitution-deck-design` restent des skills globales.
> Ce playbook a déjà été joué plusieurs fois ici (deck OHC, v2→v6 — cf. mémoire projet
> `projet-deck-ohc-ecoute`), en génération inline avant la création du sous-agent.

La chaîne PPT complète du projet source, rendue structurelle : produire ou faire évoluer
le deck de restitution, enrichir si pertinent (cadres photo du template, qualité
rédactionnelle), puis **toujours** vérifier au rendu réel — python-pptx est un parseur
tolérant, un fichier qui parse peut ne pas s'ouvrir dans PowerPoint.

Précédent (VSCode2, statut `eprouve` là-bas) : la colonne vertébrale génération →
vérification rendu est la pratique effective de ce projet frère — paire `pptx-deck` +
`pptx-verify` jouée le 2026-07-03, `pptx-verify` rejoué le 2026-07-18. Les trois étapes
conditionnelles s'appuyaient déjà sur des skills **jamais utilisées à ce jour** sur
VSCode2 (`pptx-framed-image`, `slide-text-polish`, `restitution-deck-design`) —
conservées par arbitrage utilisateur là-bas et reliées ici pour exister dans le routage :
les proposer avec prudence explicite et vérifier leur résultat au rendu.

Routage de l'étape `generation` (aligné sur l'arbitrage VSCode3 du 2026-07-21) : elle
s'instancie via le **sous-agent `ppt-designer`** (outil `Agent`), pas en génération inline
dans la session. Modèle hérité du thread principal (pas de bascule : jugement visuel).
C'est la voie deck unique — `bmad-agent-ux-designer` ne double pas ce rôle. Le brief de
l'agent porte les règles de sécurité **deck binaire** propres à ce projet (ajouter avant
de supprimer, purge des rels orphelines, ouverture COM réelle) — ne pas les court-circuiter.
Une passe de contenu ciblée peut rester inline avec rendu réel, en le notant dans le run
journalisé.

Frontière avec `dev-verifie` : si la demande est un changement de code générique, c'est
`dev-verifie` qui s'applique — ce playbook-ci est la version spécialisée quand le
**livrable est le deck lui-même** (layout, contenu, visuel). Les deux partagent
l'obligation `pptx-verify` et la terminaison `revue-increment`.

```json
{
  "nom": "export-ppt-verifie",
  "description": "Production ou évolution du deck PPT de restitution : génération, enrichissements conditionnels (cadres photo, polish rédactionnel, passe design), vérification au rendu réel obligatoire, revue-increment avant commit.",
  "statut": "eprouve",
  "source": "manuel",
  "declencheurs": [
    "génère/améliore/corrige le deck PPT de restitution d'une mission",
    "changement de layout, de constantes ou de slide dans un export python-pptx",
    "remplir les cadres photo (« ici mettre une Photo ») d'un template client",
    "qualité rédactionnelle / design des slides du deck"
  ],
  "etapes": [
    {
      "id": "cadrage",
      "agent": "session principale",
      "mode": "cascade",
      "modele": "(session)",
      "contrat": {
        "type": "deterministe",
        "critere": "données de la mission identifiées (synthèse globale, axes, recommandations), template client ou deck vierge choisi, constantes de layout relues si elles bougent (parité aperçu web / PPT le cas échéant)"
      },
      "checkpoint": false
    },
    {
      "id": "generation",
      "agent": "ppt-designer",
      "mode": "cascade",
      "modele": "(session)",
      "contrat": {
        "type": "deterministe",
        "critere": "instancié via le sous-agent ppt-designer (outil Agent), pas inline ; .pptx sauvé sans exception sur une copie versionnée (jamais l'original Imports/), auto-check géométrique passé, règles deck binaire respectées (ajouter-avant-supprimer, purge rels orphelines), ouverture PowerPoint COM réelle OK (pas de 0x80CB4404)"
      },
      "checkpoint": false
    },
    {
      "id": "cadres-photo",
      "agent": "pptx-framed-image",
      "mode": "cascade",
      "modele": "(session)",
      "contrat": {
        "type": "deterministe",
        "critere": "SI un cadre photo du template est touché (prstGeom round2DiagRect, « ici mettre une Photo ») : image insérée épousant la forme exacte du cadre, chaque image vérifiée par rendu réel avant d'être gardée — sur ce deck, préférer d'abord le média du pptx original (zipfile → ppt/media) aux fetchs externes, pour la cohérence visuelle"
      },
      "checkpoint": false
    },
    {
      "id": "polish-texte",
      "agent": "slide-text-polish",
      "mode": "cascade",
      "modele": "(session)",
      "contrat": {
        "type": "deterministe",
        "critere": "SI le contenu textuel des slides a été produit ou retouché : slide_lint passé sur {title, bullets}, findings bloquants corrigés (skill greffée le 2026-07-23, jamais encore jouée sur ce dépôt — contrôler à l'étape verification-rendu)"
      },
      "checkpoint": false
    },
    {
      "id": "verification-rendu",
      "agent": "pptx-verify",
      "mode": "cascade",
      "modele": "(session)",
      "contrat": {
        "type": "reel",
        "critere": "export réel rendu en images (PowerPoint COM/LibreOffice) et inspecté visuellement, avec un rendu ZOOMÉ sur chaque NOUVEAU type de slide (valeurs alignées, panneaux ni vides ni sur-étirés, ni contenu centré par slot laissant un grand vide sous l'en-tête — défaut « panneau flottant/étiré » récurrent, invisible au self-check géométrique, cf. arbitrage superviseur VSCode3 2026-07-21 ; pas de collision avec le chrome du template) — jamais retirée à l'instanciation, quelle que soit la taille du changement"
      },
      "checkpoint": false
    },
    {
      "id": "design-review",
      "agent": "restitution-deck-design",
      "mode": "cascade",
      "modele": "(session)",
      "contrat": {
        "type": "reel",
        "critere": "SI le rendu passe la géométrie mais reste visuellement pauvre (mur de boîtes, hiérarchie absente) : passe design appliquée puis retour à verification-rendu (skill disponible mais jamais utilisée dans ce projet à ce jour — prudence)"
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
        "critere": "SI du code produit a été modifié (export PPT, constantes de layout) : boucle revue + correctifs + re-vérification exécutée en entier"
      },
      "checkpoint": "avant tout commit — action difficilement réversible, proposer, ne pas exécuter unilatéralement"
    }
  ],
  "regle_reprise": "une relance ciblée par étape en échec de contrat, puis escalade utilisateur avec l'état réel"
}
```
