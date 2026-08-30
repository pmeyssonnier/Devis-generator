"""Rend le package `chiffrage` importable depuis la racine du dépôt.

pytest n'insère dans sys.path que le dossier du conftest (`tests/`), pas la
racine — sans ça, `pytest` (l'exécutable, contrairement à `python -m pytest`)
ne trouve pas `chiffrage`.
"""
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# `.streamlit/config.toml` pose showErrorDetails = "none" pour qu'un client
# final ne reçoive pas une traceback en plein écran. Streamlit lit ce fichier
# aussi quand pytest tourne depuis la racine : sans cette variable, un test
# qui échoue affiche « error message is redacted » au lieu de la cause. La
# variable d'environnement l'emporte sur le fichier.
os.environ.setdefault("STREAMLIT_CLIENT_SHOW_ERROR_DETAILS", "full")
