"""Rend le package `chiffrage` importable depuis la racine du dépôt.

pytest n'insère dans sys.path que le dossier du conftest (`tests/`), pas la
racine — sans ça, `pytest` (l'exécutable, contrairement à `python -m pytest`)
ne trouve pas `chiffrage`.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
