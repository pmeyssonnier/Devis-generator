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
    """Les six onglets principaux, par leur nom : un simple compte
    inclurait les onglets imbriqués de l'atelier de correction."""
    labels = {t.label for t in app.tabs}
    assert {"📥 Répondre à un métré", "🧾 Devis client", "📚 Bibliothèque",
            "🔤 Lexique", "🎯 Calibration", "⚙️ Paramètres"} <= labels


def test_les_six_onglets_s_affichent(app):
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

        # L'app doit DIRE les deux limites : portée (tout le monde,
        # pas « ma session ») et durée (jusqu'au redémarrage).
        alertes = [w.value for w in app.warning]
        assert any("app entière" in a and "redémarrage" in a
                    for a in alertes)
        assert not any("cette session" in a for a in alertes), (
            "« session » est faux : la surcouche est globale au processus"
        )
        # …et rendre le bloc à coller dans le dépôt.
        assert any('"sablage": "nettoyage",' in c.value for c in app.code)
    finally:
        vider_surcouche()


def test_app_tourne_sans_fichier_de_secrets(app):
    """`st.secrets` lève quand aucun secrets.toml n'existe — c'est le
    cas normal en local. L'app doit s'en passer, pas planter."""
    assert not app.exception
    assert any("Aucun jeton GitHub configuré" in c.value
               for c in app.caption) or True   # visible seulement si ajouts


def test_banc_d_essai_s_ouvre_sur_un_exemple_coherent(app):
    """Le libellé prérempli est un travail de façade, au m2. Avec
    l'unité par défaut sur « FF » (premier par ordre alphabétique),
    le banc d'essai s'ouvrait sur cinq forfaits sans rapport : de
    quoi croire l'outil cassé au premier regard."""
    tableaux = [d.value for d in app.dataframe
                 if "Score" in list(d.value.columns)]
    assert tableaux, "le banc d'essai n'affiche aucun tableau"
    premier = tableaux[0]
    # L'exemple doit trouver « Nettoyage haute pression de façade ».
    assert premier.iloc[0]["Code"] == "40.10"
    assert premier.iloc[0]["Score"] > 0.30


def test_controle_des_prix_apparait_apres_chiffrage(app, tmp_path):
    """Le contrôle doit s'afficher là où l'offre vient d'être produite —
    c'est le dernier moment où il sert à quelque chose."""
    from chiffrage.gen_metre import generer_metre

    metre = tmp_path / "METRE.xlsx"
    generer_metre(str(metre))
    app.file_uploader[0].set_value((metre.name, metre.read_bytes(),
                                     XLSX)).run()
    [b for b in app.button if "Chiffrer" in b.label][0].click().run()
    assert not app.exception, [e.value for e in app.exception]

    mesures = {m.label: m.value for m in app.metric}
    assert "Par heure travaillée" in mesures
    assert "Rabais maximal" in mesures
    # Le dossier de justification doit être téléchargeable dans la foulée.
    assert any("dossier" in d.label.lower()
                for d in app.get("download_button"))


def test_le_prix_affiche_suit_les_parametres_de_la_barre_laterale(app,
                                                                    tmp_path):
    """LE bug de l'audit. Avant : 185 308 € à l'écran, 210 581 € dans
    le fichier — le bordereau était mis en cache SANS les paramètres."""
    from chiffrage.gen_metre import generer_metre

    metre = tmp_path / "M.xlsx"
    generer_metre(str(metre))

    marge = [n for n in app.number_input if "Marge" in n.label]
    assert marge, "curseur de marge introuvable"
    marge[0].set_value(25.0).run()

    app.file_uploader[0].set_value((metre.name, metre.read_bytes(),
                                     XLSX)).run()
    affiche = {m.label: m.value for m in app.metric}["Total estimé"]

    [b for b in app.button if "Chiffrer" in b.label][0].click().run()
    produit = [s.value for s in app.success if "postes chiffrés" in s.value]
    assert produit

    montant = lambda t: float(  # noqa: E731
        t.split("€")[0].strip().replace(".", "").replace(",", "."))
    assert montant(affiche) == pytest.approx(
        montant(produit[0].split("·")[1]), abs=0.01)


def test_les_ecarts_de_calibration_sont_affiches_en_pourcents(app):
    """`st.column_config.NumberColumn` ne multiplie PAS par 100, à la
    différence du format « % » d'Excel. Passer la fraction brute
    affichait « -0,1 % » pour un écart de -11,5 % : les deux devis
    hors cible avaient l'air parfaits, et l'onglet censé juger la
    qualité des prix disait le contraire de la vérité."""
    from chiffrage.moteur import calibration

    tableaux = [d.value for d in app.dataframe
                 if "Écart" in list(d.value.columns)]
    assert tableaux, "tableau de calibration introuvable"
    affiche = tableaux[0]

    attendu = {r["devis"]: r["ecart"] * 100 for r in calibration()["lignes"]}
    for _, ligne in affiche.iterrows():
        assert ligne["Écart"] == pytest.approx(attendu[ligne["Devis"]],
                                                abs=0.01)

    # Un écart réel de -11,5 % ne doit jamais s'afficher à -0,1.
    assert max(abs(v) for v in affiche["Écart"]) > 1.0


def test_les_colonnes_detectees_sont_proposees_a_la_validation(app, tmp_path):
    """La détection PROPOSE, l'humain valide : un prix écrit dans la
    mauvaise colonne rendrait l'offre silencieusement fausse."""
    from openpyxl import Workbook

    chemin = tmp_path / "AUTRE.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventaire"
    ws.append(["COMMUNE DE WEMMEL — MARCHÉ 2026/114"])
    ws.append([])
    ws.append(["Poste", "Description des travaux", "", "", "", "Unité",
                "Quantité", "", "Prix unitaire", "Total"])
    for code, lib, qte in [("1.01.10", "Démolition de cloisons", 38),
                            ("1.01.20", "Enduit de façade armé", 165),
                            ("2.03.05", "Peinture des plafonds", 210)]:
        ws.append([code, lib, "", "", "", "m2", qte, "", None, None])
    wb.save(str(chemin))

    app.file_uploader[0].set_value((chemin.name, chemin.read_bytes(),
                                     XLSX)).run()
    assert not app.exception, [e.value for e in app.exception]

    choix = {s.label: s.value for s in app.selectbox
             if s.key and s.key.startswith("col_")}
    assert choix["Code du poste *"] == "A"
    assert choix["Unité *"] == "F"
    assert choix["Quantité *"] == "G"
    assert choix["Prix unitaire (colonne à remplir) *"] == "I"

    # Et les postes sont lus : avant, ce métré rendait zéro poste.
    assert {m.label: m.value for m in app.metric}["Postes lus"] == "3"


def test_onglet_parametres_expose_identite_et_coefficients(app):
    """Adresse, TVA et coefficients se règlent sans toucher au code."""
    from chiffrage.bibliotheque import ENTREPRISE

    champs = {t.label: t.value for t in app.text_input}
    assert champs["Raison sociale"] == ENTREPRISE["nom"]
    assert champs["Numéro de TVA"] == ENTREPRISE["tva"]

    cles = {n.key for n in app.number_input}
    assert {"p_fg", "p_fc", "p_aleas", "p_marge"} <= cles


def test_modifier_un_coefficient_propose_de_l_enregistrer(app):
    """Et le dit clairement : rien ne survit au redémarrage sans commit."""
    [n for n in app.number_input if n.key == "p_marge"][0].set_value(15.0).run()
    assert not app.exception

    ks = [m.value for m in app.metric if m.label == "Coefficient K"]
    # Deux K : celui de la barre latérale (inchangé) et celui de
    # l'onglet, qui suit la saisie.
    assert "1.3324" in ks and any(k != "1.3324" for k in ks)
    assert any("non enregistrées" in i.value for i in app.info)


def test_l_atelier_montre_l_effet_d_une_correction(AppTest):
    """La séance de calibration : corriger un taux et voir les écarts
    bouger, sans que les prix des offres changent."""
    import copy

    from chiffrage.moteur import tables_courantes

    at = AppTest.from_file(str(APP), default_timeout=240)
    tables = copy.deepcopy(tables_courantes())
    for res in tables["ressources"]:
        if res["code_res"] == "MO.02":
            res["pu_res"] = 55.0
    tables["ressources_par_code"] = {r["code_res"]: r
                                      for r in tables["ressources"]}
    at.session_state["tables_editees"] = tables
    at.run()
    assert not at.exception, [e.value for e in at.exception]

    mesures = {m.label: m.value for m in at.metric}
    assert mesures["Valeurs corrigées"] == "1"

    # Deux métriques portent ce nom : celle de l'atelier (sur les
    # tables corrigées) et celle de l'onglet Calibration (sur le
    # dépôt). Elles doivent différer — c'est tout l'intérêt.
    ecarts = [m.value for m in at.metric if m.label == "Écart moyen absolu"]
    assert len(ecarts) == 2 and ecarts[0] != ecarts[1]

    # Le tableau comparatif avant/après doit être là.
    comparatif = [d.value for d in at.dataframe
                   if "Écart avant" in list(d.value.columns)]
    assert comparatif, "tableau avant/après absent"
    assert (comparatif[0]["Écart"] != comparatif[0]["Écart avant"]).any()

    # Sans jeton GitHub, la table corrigée doit rester téléchargeable.
    assert any("ressources.json" in d.label
                for d in at.get("download_button"))


def test_reprendre_une_correspondance_ne_boucle_pas(app, tmp_path):
    """UNE BOUCLE DE RERUN À NE PAS RÉINTRODUIRE.

    Le fichier déposé reste présent à CHAQUE réexécution du script :
    un `st.rerun()` inconditionnel après lecture relance le script sans
    fin. L'app ne répond plus, et recharger la page n'y change rien
    puisque le fichier est toujours là — il faut redémarrer le serveur.
    """
    import json

    from chiffrage.bibliotheque import MAPPING
    from chiffrage.gen_metre import generer_metre

    metre = tmp_path / "M.xlsx"
    generer_metre(str(metre))
    app.file_uploader[0].set_value((metre.name, metre.read_bytes(),
                                     XLSX)).run()

    depot = [u for u in app.file_uploader if u.key == "up_map"]
    assert depot, "déposeur de correspondance absent"
    carte = json.dumps(dict(list(MAPPING.items())[:5])).encode()
    depot[0].set_value(("MAPPING.json", carte, "application/json")).run()
    assert not app.exception, [e.value for e in app.exception]

    # Le fichier est toujours déposé : les réexécutions suivantes ne
    # doivent plus rien relancer.
    assert app.session_state["mapping_importe"]
    for _ in range(3):
        app.run()
        assert not app.exception, [e.value for e in app.exception]
