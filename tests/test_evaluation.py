"""Non-régression de l'appariement, mesurée sur le jeu d'épreuve.

`evaluation/epreuves_appariement.json` porte des libellés écrits en
vocabulaire de cahier spécial des charges. Ces tests figent le niveau
atteint : une modification du lexique ou du score qui ferait retomber
l'appariement casse la CI au lieu de passer inaperçue.

Les seuils sont volontairement SOUS le niveau mesuré : ils protègent
d'une régression, ils ne certifient pas une performance. Le niveau réel
sur un vrai cahier des charges sera plus bas — ces libellés sont écrits
par la même main que la bibliothèque.
"""
import json
import sys
from pathlib import Path

import pytest

RACINE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RACINE / "evaluation"))

from chiffrage.moteur import calcul_bordereau  # noqa: E402


@pytest.fixture(scope="module")
def mesure():
    from mesurer_appariement import charger, mesurer

    return mesurer(charger(), calcul_bordereau())


def test_jeu_d_epreuve_bien_forme():
    donnees = json.loads(
        (RACINE / "evaluation" / "epreuves_appariement.json").read_text("utf-8")
    )
    epreuves = donnees["epreuves"]
    assert len(epreuves) >= 50
    codes = {o["code_ouv"] for o in __import__(
        "chiffrage.bibliotheque", fromlist=["OUVRAGES"]).OUVRAGES}
    for e in epreuves:
        assert e["type"] in ("reformulation", "piege", "absent")
        assert e["unite"]
        # attendu = null est un cas VALIDE : l'outil doit se taire.
        assert e["attendu"] is None or e["attendu"] in codes


def test_appariement_au_premier_rang(mesure):
    assert mesure["rang1"] / mesure["attendus"] >= 0.90


def test_appariement_dans_les_trois_premiers(mesure):
    """C'est ce que voit l'utilisateur dans l'interface, donc le chiffre
    qui compte pour lui."""
    assert mesure["top3"] / mesure["attendus"] >= 0.96


def test_l_outil_se_tait_quand_aucun_ouvrage_ne_convient(mesure):
    """Le plus important des trois. Une suggestion confiante sur un
    poste qu'aucun ouvrage ne couvre fait chiffrer un travail par un
    autre — c'est pire que pas de suggestion du tout."""
    assert mesure["silences_ok"] >= 6


def test_operation_inverse_est_penalisee():
    """« Dépose de X » ne doit pas être apparié à « Pose de X » :
    même vocabulaire, travail opposé, prix du simple au triple."""
    from chiffrage.suggestion import score

    pose = "Faux plafond BA13 sur ossature métallique, jointoyé"
    depose = "Dépose de plafond existant (plâtre ou plaques)"
    assert score(depose, pose) < 0.30


def test_vocabulaire_de_csc_est_traduit():
    """L'exemple du client : « carrelage mural » doit trouver la faïence."""
    from chiffrage.suggestion import suggerer

    b = calcul_bordereau()
    candidats = suggerer(
        {"designation": "Carrelage mural en faïence, profilés compris",
         "unite": "m2"}, b, limite=1)
    assert candidats and candidats[0][0] == "70.60"
