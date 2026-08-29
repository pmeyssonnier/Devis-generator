"""
Outil de chiffrage BAG BATTER SRL — bibliothèque de prix unitaires.

Reconstruction du notebook Colab décrit dans CLAUDE.md, sous forme de modules
Python importables (et exécutables en cellule Colab par simple copier-coller).

Découpage volontaire :
  - bibliotheque.py  : les DONNÉES (aucune dépendance externe)
  - moteur.py        : les CALCULS   (aucune dépendance externe)
  - export_xlsx.py   : export de la bibliothèque vers Excel   (openpyxl)
  - gen_metre.py     : génération d'un métré de marché public (openpyxl)
  - metre_io.py      : lecture d'un métré imposé + remplissage (openpyxl)

Les deux premiers modules restent en Python pur : ils sont testables en CI
(tests/test_chiffrage.py) sans installer openpyxl.
"""

from .bibliotheque import (  # noqa: F401
    COMPOSITION,
    LOTS,
    MAPPING,
    METRES_HISTO,
    OUVRAGES,
    PARAMS,
    RESSOURCES,
)
from .moteur import (  # noqa: F401
    calcul_bordereau,
    calibration,
    coefficient_k,
    controle_coherence,
    devis,
    fiche_prix,
)

__all__ = [
    "RESSOURCES",
    "OUVRAGES",
    "COMPOSITION",
    "PARAMS",
    "LOTS",
    "MAPPING",
    "METRES_HISTO",
    "coefficient_k",
    "calcul_bordereau",
    "devis",
    "fiche_prix",
    "calibration",
    "controle_coherence",
]
