# Figures scientifiques

Quand je demande un **graphical abstract**, un **TOC graphic**, de **générer des graphes / figures / schémas / diagrammes** (ou "figure", "graph", "plot", "flowchart", "workflow figure" en anglais), invoque les skills avant de commencer :

- `academic-figures` : graphical abstract, TOC graphic, figure multi-panneaux, schéma, en SVG éditable (2-3 variantes).
- `scientific-figure-workflow` : figure ou graphical abstract livrés en PPTX/SVG éditables, ou conversion d'une image PNG/JPG existante en version éditable.
- Si le graphe doit tracer des **données réelles** (courbes, barres, nuages de points), utilise aussi `scientific-agent-skills:scientific-visualization` : les deux skills ci-dessus ne laissent que des emplacements vides pour les tracés.

# Présentations

Une figure scientifique livrée en PPTX relève de la section ci-dessus, pas de celle-ci. Pour un deck, invoque un seul de ces skills :

- `presentation-skill` : **créer** un deck PowerPoint (conseil, stratégie, pitch, comité de pilotage, exec readout, deck sur un template `.pptx` fourni, waterfall / matrice 2x2 / gantt), et réviser un deck qu'il a lui-même produit (commentaires de relecture, mise à jour des données).
- `pptx` : **lire, extraire le texte, modifier, fusionner ou découper** un `.pptx` / `.potx` existant qui ne vient pas de `presentation-skill`.
- `slides` : présentation en HTML uniquement.
