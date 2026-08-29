"""Tests de l'interface Streamlit (`streamlit_app.py`).

Streamlit fournit `AppTest` : il exécute le script comme le ferait le
navigateur, expose les widgets et attrape les exceptions. C'est donc une
vraie exécution, pas une vérification de syntaxe.

Ce que ces tests ont déjà attrapé, et qu'aucune relecture n'aurait vu :
  · un `icon=` refusé par Streamlit (l'app plantait au chargement) ;
  · un `st.stop()` dans l'onglet « métré » qui arrêtait le SCRIPT entier :
    les trois autres onglets ne s'affichaient jamais.

Ils se skippent si streamlit n'est pas installé — le moteur, lui, n'en
dépend pas.
"""
import pathlib

import pytest

APP = pathlib.Path(__file__).resolve().parent.parent / "streamlit_app.py"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@pytest.fixture(scope="module")
def AppTest():
    pytest.importorskip("streamlit")
    from streamlit.testing.v1 import AppTest as _AppTest

    return _AppTest


@pytest.fixture
def app(AppTest):
    at = AppTest.from_file(str(APP), default_timeout=180)
    at.run()
    assert not at.exception, [e.value for e in at.exception]
    return at


def test_app_demarre_sans_erreur(app):
    assert len(app.tabs) == 5


def test_les_cinq_onglets_s_affichent(app):
    """Garde-fou du `st.stop()` : si un onglet arrête le script, les
    suivants n'ont plus ni tableau ni métrique."""
    labels = {m.label for m in app.metric}
    assert "Coefficient K" in labels          # barre latérale
    assert "Total HTVA" in labels             # onglet devis
    assert "Écart moyen absolu" in labels     # onglet calibration
    assert len(app.dataframe) >= 4            # bibliothèque + calibration


def test_avertissement_de_calibration_visible(app):
    """Les prix ne sont pas calibrés : ça doit se lire là où les
    documents sont produits, pas seulement dans le README."""
    assert any("non calibrés" in w.value for w in app.warning)


def test_parcours_complet_metre_vers_offre(app, tmp_path):
    """Dépôt d'un métré -> appariement -> offre téléchargeable."""
    from chiffrage.gen_metre import generer_metre

    metre = tmp_path / "METRE_test.xlsx"
    generer_metre(str(metre))

    app.file_uploader[0].set_value(
        (metre.name, metre.read_bytes(), XLSX)
    ).run()
    assert not app.exception, [e.value for e in app.exception]

    metriques = {m.label: m.value for m in app.metric}
    assert metriques["Postes lus"] == "49"
    # Le métré d'entraînement est intégralement couvert par MAPPING :
    # aucun poste ne doit rester à apparier.
    assert metriques["Appariés"] == "49"
    assert metriques["À revoir"] == "0"

    chiffrer = [b for b in app.button if "Chiffrer" in b.label]
    assert chiffrer, "bouton de chiffrage absent"
    chiffrer[0].click().run()
    assert not app.exception, [e.value for e in app.exception]

    telechargements = [d.label for d in app.get("download_button")]
    assert any("offre" in t.lower() for t in telechargements)


def test_le_total_de_l_interface_egale_celui_du_moteur(app, tmp_path):
    """L'interface ne doit JAMAIS recalculer un prix pour son compte :
    deux vérités sur un prix, c'est une offre fausse tôt ou tard."""
    from chiffrage.gen_metre import generer_metre
    from chiffrage.metre_io import remplir_metre

    metre = tmp_path / "METRE_test.xlsx"
    generer_metre(str(metre))
    attendu = remplir_metre(str(metre), str(tmp_path / "offre.xlsx"))

    app.file_uploader[0].set_value(
        (metre.name, metre.read_bytes(), XLSX)
    ).run()
    total_affiche = {m.label: m.value for m in app.metric}["Total estimé"]

    # '185.308,33 €' -> 185308.33
    valeur = float(
        total_affiche.removesuffix(" €").replace(".", "").replace(",", ".")
    )
    assert valeur == pytest.approx(attendu["total_ht"], abs=0.01)


def test_onglet_lexique_permet_d_ajouter_un_terme(app):
    """Le cycle qui rend le lexique réglable sans écrire de Python :
    on colle un libellé mal apparié, on ajoute le mot manquant, on
    voit le score bouger — et l'app rend le bloc à commiter."""
    from chiffrage.lexique import vider_surcouche

    try:
        champs = {t.label: t for t in app.text_input}
        champs["Terme du cahier des charges"].set_value("sablage").run()
        champs["Terme de la bibliothèque"].set_value("nettoyage").run()

        ajouter = [b for b in app.button if "Ajouter" in b.label]
        assert ajouter and not ajouter[0].disabled
        ajouter[0].click().run()
        assert not app.exception, [e.value for e in app.exception]

        # L'app doit DIRE que ça ne survivra pas au redémarrage.
        assert any("session" in w.value and "redémarrage" in w.value
                    for w in app.warning)
        # …et rendre le bloc à coller dans le dépôt.
        assert any('"sablage": "nettoyage",' in c.value for c in app.code)
    finally:
        vider_surcouche()
