# Playbook `export-ppt-verifie` — travaux sur le deck de restitution, vérifiés au rendu réel

> Repris tel quel du projet frère VSCode2 le 2026-07-21 — précédents et dates ci-dessous
> appartiennent à l'historique de VSCode2 ; ce playbook n'a pas encore été rejoué dans ce
> projet. Deux des étapes conditionnelles (`cadres-photo`, `polish-texte`) référencent des
> skills (`pptx-framed-image`, `slide-text-polish`) **non installées dans ce projet** — cf.
> `.claude/orchestration/catalogue.md` ; ne pas router vers elles avant leur éventuelle
> installation. `pptx-deck`, `pptx-verify` et `restitution-deck-design` sont en revanche
> bien disponibles ici (skills globales).

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
      "agent": "pptx-deck",
      "mode": "cascade",
      "modele": "(session)",
      "contrat": {
        "type": "deterministe",
        "critere": "export .pptx produit sans exception, auto-check géométrique passé, pytest -k \"pptx or export\" vert (si applicable)"
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
        "critere": "SI le template client porte des cadres photo (prstGeom round2DiagRect, « ici mettre une Photo ») : image insérée épousant la forme exacte du cadre (skill non installée dans ce projet à ce jour — vérifier sa disponibilité avant de router ici, sinon traiter à la main)"
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
        "critere": "SI le contenu textuel des slides a été produit ou retouché : lint appliqué sur {title, bullets}, findings bloquants corrigés (skill non installée dans ce projet à ce jour — vérifier sa disponibilité avant de router ici, sinon traiter à la main)"
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
        "critere": "export réel rendu en images et inspecté visuellement (valeurs alignées, panneaux ni vides ni étirés, pas de collision avec le chrome du template) — jamais retirée à l'instanciation, quelle que soit la taille du changement"
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
