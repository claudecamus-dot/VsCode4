---
name: deck-design-review
description: Revue de design slide par slide du deck OHC « dispositif d'écoute » de CE projet (généré par scripts/generate_deck_ohc.py, 15 slides) — régénérer le vrai .pptx, le rendre ENTIER en images (LibreOffice), et passer chaque type de slide contre son propre contrat de design (couverture, sommaire, dividers de chapitre, leviers, personas, architecture/priorisation, existant, évaluation, mentorat, arbitrages, séquencement, roadmap). À lancer avant de déclarer un changement de design du deck terminé, quand le deck « n'est pas au niveau », ou comme étape de revue du playbook export-ppt-verifie.
---

# deck-design-review — la revue de design du deck OHC ENTIER

`pptx-verify` dit **comment** regarder (rendre + zoomer + checklist générique) ;
`restitution-deck-design` dit **ce qui fait pro** en général. Ce skill ajoute le
**contrat par slide de CE deck** (OHC RH « dispositif d'écoute »), pour que chaque type
de slide soit revu contre SA définition — pas une impression d'ensemble.

Porté de VSCode2 le 2026-07-23 (finding `pratique-design` du superviseur de flotte) et
**réécrit pour le deck réel de VSCode4** : les types de slides, le générateur et le canal
de rendu sont ceux de ce projet, pas de l'app mission de VSCode2.

## 0. Sur le BON artefact, TOUTES les slides

- Régénérer le vrai deck : `py scripts/generate_deck_ohc.py` — pas un extrait, le fichier
  complet des **15 slides** (`Exports/…v7-genere.pptx`). Ne jamais reviewer une version
  antérieure (`v2`…`v6`) restée dans `Exports/`.
- Rendre **toutes** les slides en images via **LibreOffice** (`soffice --headless
  --convert-to pdf`, la seule voie fiable ici — cf. `test_generate_deck_ohc.py`), pas un
  échantillon. Le test compte 15 pages : un écart de compte = une slide perdue/dupliquée.
- python-pptx est un parseur tolérant : un `.pptx` qui se génère peut mal s'ouvrir. Le
  rendu réel est non négociable.

## 1. Contrat par type de slide

| Slide | Contrat (au rendu) |
| --- | --- |
| Couverture | Layout natif « 40 - Couverture » : photo pleine page, overlay navy, titre = sujet OHC, pas le repli dessiné. |
| Sommaire | Layout natif « 92 - Table des matières » : 4 chapitres listés, parité avec l'ordre réel des dividers. |
| Divider de chapitre | Layout natif « 51 - Chapitre [2] » : numéro présent (exigence PERSISTANTE — a été écarté plusieurs fois par contrainte de gabarit, se dessine, ne s'omet pas), titre + sous-titre, **vraie photo clippée au teardrop** (Openverse CC0), couleur navy uniforme (pas de couleur par chapitre — décision v5→v6). Cadre photo rempli, jamais vide. |
| Leviers | 3 leviers, blobs « pin » cyan positionnés, design hérité du deck original — cohérence des pins. |
| Personas | 8 populations en cartes KPI, grille alignée, pas de carte vide ni débordante. |
| Architecture / priorisation | 4 quadrants + axes DESSINÉS (pas un scatter Excel), libellés d'axes lisibles, bulles dans leur quadrant. |
| Existant | 3 dispositifs RH en cartes + protocole condensé — cartes de hauteur égale, texte non tronqué. |
| Évaluation | 4 mesures + mockup Google Form dessiné (violet) — le mockup lisible, pas un aplat. |
| Mentorat | Format + cadrage — hiérarchie titre/corps, une seule idée directrice. |
| Arbitrages | 3 cartes Enjeu → Reco, pastilles DECISION navy — parallélisme des 3 cartes, pastilles alignées. |
| Séquencement | 3 phases à pastilles statut rouge/ambre/vert — code couleur = sens (pas décoratif), lisible. |
| Roadmap | Timeline 5 jalons datés, tribune 17 sept en vert — jalons alignés sur l'axe, date de tribune mise en avant. |

## 2. Règles transverses (au rendu réel)

- **Une headline par slide** : si deux idées se disputent le haut de slide, en couper une.
- **Couleur = sens** : navy = structure, cyan = accent OHC, statut rouge/ambre/vert =
  état. Pas d'aplat criard là où un accent suffit.
- **Cohérence de composant** : une carte KPI, une pastille, un pin doivent être identiques
  d'une slide à l'autre (même helper). Une variation non intentionnelle est un défaut.
- **Cadres photo remplis** : un cadre de chapitre vide (constaté en v-antérieures) est un
  défaut bloquant — la photo se met, la fidélité au pattern de référence prime.
- **Texte dans sa boîte** : au rendu LibreOffice, vérifier qu'aucun texte ne déborde (le
  français accentué et les mots longs sortent des cadres serrés — invisible au seul
  `verifier_geometrie`).

## 3. Boucle

Rendu → liste de défauts par slide (n° + type + correctif) → correction dans
`generate_deck_ohc.py` → re-génération → re-rendu. Budget 2 itérations au-delà du rendu
initial, puis escalade utilisateur avec les défauts restants en images. Un changement
d'intention design (layout, couleur, présence d'un élément) se fait **valider par
l'utilisateur sur le rendu réel** avant d'être déclaré fait.
