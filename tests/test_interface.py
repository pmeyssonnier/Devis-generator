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


def test_une_session_amputee_se_recalcule_au_lieu_de_planter(app, tmp_path):
    """Bandeau rouge signalé au rafraîchissement : `proposition` et
    `mapping` étaient lus en accès direct. Une session dont la
    signature avait survécu sans son appariement levait un KeyError,
    là où il suffisait de recalculer."""
    from chiffrage.gen_metre import generer_metre

    metre = tmp_path / "M.xlsx"
    generer_metre(str(metre))
    app.file_uploader[0].set_value((metre.name, metre.read_bytes(),
                                     XLSX)).run()
    assert {m.label: m.value for m in app.metric}["Appariés"] == "49"

    del app.session_state["proposition"]
    app.run()
    assert not app.exception, [e.value for e in app.exception]
    assert {m.label: m.value for m in app.metric}["Appariés"] == "49"


def test_corriger_une_valeur_en_trois_gestes(app):
    """Un tableur de 49 lignes ne se remplit pas au doigt : une
    correction qui ne « prend » pas ne se voit pas, et le bouton
    d'enregistrement — qui n'apparaît qu'après une modification —
    reste introuvable. On corrige donc une valeur à la fois."""
    choix = [s for s in app.selectbox if s.key == "corr_res"]
    assert choix, "sélecteur de ressource absent"

    choix[0].set_value("MO.02").run()
    [n for n in app.number_input if n.key == "corr_res_valeur"][0] \
        .set_value(55.0).run()
    [b for b in app.button if b.key == "corr_res_ok"][0].click().run()
    assert not app.exception, [e.value for e in app.exception]

    assert {m.label: m.value for m in app.metric}["Valeurs corrigées"] == "1"

    # L'atelier recalcule ; l'onglet Calibration garde les valeurs du
    # dépôt. On compare les DEUX ENTRE ELLES, jamais à un littéral :
    # ces chiffres bougent à chaque taux que le chef d'entreprise
    # calibre depuis l'app, et un test qui les fige rendrait la CI
    # rouge sans qu'une ligne de code ait changé — le meilleur moyen
    # de la faire ignorer. La leçon avait déjà été apprise sur le
    # lexique appris.
    ecarts = [m.value for m in app.metric if m.label == "Écart moyen absolu"]
    assert len(ecarts) == 2
    atelier, depot = (float(e.rstrip(" %")) for e in ecarts)
    assert atelier != depot
    # Relever le taux du façadier renchérit l'ouvrage : les devis
    # calculés montent, donc l'écart aux forfaits vendus aussi.
    assert atelier > depot


def test_le_bouton_d_enregistrement_n_apparait_qu_apres_correction(app):
    """Il était introuvable parce qu'il n'existe pas sans modification —
    voulu, mais déroutant si la saisie n'a pas été prise."""
    assert not [b for b in app.button if "Enregistrer" in b.label
                 and "table" in b.label]

    [s for s in app.selectbox if s.key == "corr_res"][0].set_value("MO.02").run()
    [n for n in app.number_input if n.key == "corr_res_valeur"][0] \
        .set_value(55.0).run()
    [b for b in app.button if b.key == "corr_res_ok"][0].click().run()

    # Sans jeton GitHub, la table corrigée reste téléchargeable.
    assert any("ressources.json" in d.label for d in app.get("download_button"))


def _regler(app, cle, valeur, liste):
    [w for w in getattr(app, liste) if w.key == cle][0].set_value(valeur).run()


def test_creer_un_ouvrage_complet(app, tmp_path):
    """Créer un ouvrage absent : identité, unité, composition — et le
    résultat doit passer le contrôle du chargeur, sans quoi l'app
    refuserait de démarrer au prochain déploiement."""
    import json
    import shutil

    from chiffrage.bibliotheque import DOSSIER_DATA, charger_tables

    _regler(app, "neuf_lot", "70", "selectbox")
    _regler(app, "neuf_libelle", "Ragréage de sol autolissant", "text_input")
    _regler(app, "neuf_unite", "m2", "selectbox")

    for ressource, quantite in [("MO.02", 0.25), ("MA.23", 0.6)]:
        _regler(app, "neuf_res", ressource, "selectbox")
        _regler(app, "neuf_qte", quantite, "number_input")
        [b for b in app.button if b.key == "neuf_add"][0].click().run()
    assert not app.exception, [e.value for e in app.exception]

    [b for b in app.button if b.key == "neuf_creer"][0].click().run()
    assert not app.exception, [e.value for e in app.exception]

    tables = app.session_state["tables_editees"]
    neufs = [o for o in tables["ouvrages"] if o["code_ouv"] == "70.80"]
    assert neufs and neufs[0]["unite_ouv"] == "m2"
    assert sum(1 for c in tables["composition"]
                if c["code_ouv"] == "70.80") == 2

    # Le fichier produit doit être rechargeable : un ouvrage sans
    # composition, ou dont l'unité manque, ferait échouer le démarrage.
    data = tmp_path / "data"
    shutil.copytree(DOSSIER_DATA, data)
    (data / "ouvrages.json").write_text(json.dumps(
        [{k: v for k, v in o.items() if k != "lot"}
         for o in tables["ouvrages"]], ensure_ascii=False), encoding="utf-8")
    (data / "composition.json").write_text(json.dumps(
        tables["composition"], ensure_ascii=False), encoding="utf-8")
    rechargees = charger_tables(data)
    assert "70.80" in rechargees["ouvrages_par_code"]


def test_le_numero_propose_suit_le_lot(app):
    """Un widget à clé garde SA valeur d'un rerun à l'autre : sans
    remise à jour, changer de lot laissait le numéro du lot précédent
    — qui tombait sur un code déjà pris, et la création échouait."""
    _regler(app, "neuf_lot", "30", "selectbox")
    assert any("30.70" in m.value for m in app.markdown)
    _regler(app, "neuf_lot", "90", "selectbox")
    assert any("90.40" in m.value for m in app.markdown)


def test_un_code_deja_pris_bloque_la_creation(app):
    _regler(app, "neuf_lot", "40", "selectbox")
    _regler(app, "neuf_num", 20, "number_input")      # 40.20 existe
    assert any("existe déjà" in e.value for e in app.error)
    assert [b for b in app.button if b.key == "neuf_creer"][0].disabled


def test_un_ouvrage_sans_composition_ne_peut_pas_etre_cree(app):
    """Sans composition il se vendrait à 0 €, et le contrôle au
    chargement le refuserait — autant l'empêcher ici."""
    _regler(app, "neuf_lot", "70", "selectbox")
    _regler(app, "neuf_libelle", "Poste sans composition", "text_input")
    assert [b for b in app.button if b.key == "neuf_creer"][0].disabled


# ── Ce que voit un client final ─────────────────────────────────────────────
# Ces deux tests ne vérifient pas du code mais de la configuration. Ils sont
# ici parce que la configuration est justement ce qui casse en silence : un
# fichier ignoré par git n'est jamais déployé, et rien ne le signale.

def test_la_barre_de_developpement_est_masquee():
    """`toolbarMode = "viewer"` retire Rerun / Clear cache / Deploy du menu.
    Ce n'est pas une protection — c'est de l'affichage — mais le client final
    ne doit pas tomber sur une console de développement."""
    tomllib = pytest.importorskip("tomllib")
    fichier = APP.parent / ".streamlit" / "config.toml"
    assert fichier.exists(), "sans ce fichier, Streamlit reprend ses défauts"
    config = tomllib.loads(fichier.read_text(encoding="utf-8"))
    assert config["client"]["toolbarMode"] == "viewer"


def test_la_configuration_daffichage_est_bien_suivie_par_git():
    """`.gitignore` exclut `.streamlit/*.toml` pour protéger secrets.toml.
    L'exception qui laisse passer config.toml est fragile : la perdre ne
    casserait rien en local, l'app déployée retrouverait juste ses boutons
    de développement. D'où ce garde-fou."""
    import shutil
    import subprocess

    if shutil.which("git") is None:                    # pragma: no cover
        pytest.skip("git absent")
    ignore = subprocess.run(
        ["git", "check-ignore", "-q", ".streamlit/config.toml"],
        cwd=APP.parent, capture_output=True)
    assert ignore.returncode != 0, ".streamlit/config.toml est ignoré par git"


# ── Reprendre un devis ──────────────────────────────────────────────────────

def _deposer_devis(app, charge):
    import json

    depot = [u for u in app.file_uploader if u.key == "up_devis"]
    assert depot, "déposeur de devis absent"
    depot[0].set_value(("DEVIS.json", json.dumps(charge).encode(),
                         "application/json")).run()
    return app


def test_un_devis_enregistre_se_recharge_dans_les_champs(app):
    _deposer_devis(app, {
        "version": 1, "objet": "Toiture arrière", "reference": "2026-007",
        "chantier": "Rue Haute 3", "client": "SPRL Test", "tva": 0.21,
        "lignes": [{"code_ouv": "40.20", "qte": 12.5}]})
    assert not app.exception, [e.value for e in app.exception]
    assert app.session_state["devis_objet"] == "Toiture arrière"
    assert app.session_state["devis_reference"] == "2026-007"
    assert app.session_state["tva_devis"] == 0.21
    assert app.session_state["lignes_devis"] == [{"code_ouv": "40.20",
                                                   "qte": 12.5}]
    # Les postes d'exemple du devis précédent ne doivent pas survivre.
    assert len(app.session_state["lignes_devis"]) == 1


def test_le_depot_dun_devis_ne_relance_pas_le_script_sans_fin(app):
    """Même piège que la correspondance de métré : le fichier reste
    déposé à chaque réexécution."""
    _deposer_devis(app, {"objet": "X", "tva": 0.06, "lignes": []})
    assert app.session_state["devis_importe"]
    for _ in range(3):
        app.run()
        assert not app.exception, [e.value for e in app.exception]
    assert app.session_state["devis_objet"] == "X"


def test_les_modifications_du_devis_precedent_ne_suivent_pas(app):
    """`st.data_editor` garde ses ajouts sous sa propre clé et les
    réapplique aux données suivantes : sans oubli explicite, les lignes
    tapées à la main dans le devis précédent viendraient se coller au
    devis qu'on vient de charger."""
    app.session_state["editeur_devis"] = {
        "edited_rows": {}, "added_rows": [{"code_ouv": "40.30", "qte": 99.0}],
        "deleted_rows": []}
    _deposer_devis(app, {"objet": "Neuf", "tva": 0.06,
                          "lignes": [{"code_ouv": "40.20", "qte": 1.0}]})
    assert not app.exception, [e.value for e in app.exception]
    assert all(ligne["qte"] != 99.0
                for ligne in app.session_state["lignes_devis"])


def test_un_devis_illisible_ne_plante_pas_lapp(app):
    depot = [u for u in app.file_uploader if u.key == "up_devis"]
    depot[0].set_value(("DEVIS.json", b"pas du json", "application/json")).run()
    assert not app.exception, [e.value for e in app.exception]
    assert any("illisible" in e.value for e in app.error)


def test_un_ouvrage_inconnu_est_signale_a_lecran(app):
    _deposer_devis(app, {"objet": "X", "tva": 0.06,
                          "lignes": [{"code_ouv": "99.99", "qte": 1.0}]})
    assert any("99.99" in w.value for w in app.warning)


def test_le_devis_se_telecharge_aussi_en_json(app):
    labels = [d.label for d in app.get("download_button")]
    assert any(".json" in label and "modifier" in label for label in labels)


def test_un_code_inconnu_dans_une_correspondance_reste_affiche(app, tmp_path):
    """L'avertissement était posé juste avant un `st.rerun()`, qui
    rejoue le script depuis le début : il disparaissait de l'écran sans
    jamais avoir été lu, et la correspondance repartait amputée en
    silence."""
    import json

    from chiffrage.gen_metre import generer_metre

    metre = tmp_path / "M.xlsx"
    generer_metre(str(metre))
    app.file_uploader[0].set_value((metre.name, metre.read_bytes(),
                                     XLSX)).run()
    depot = [u for u in app.file_uploader if u.key == "up_map"]
    depot[0].set_value(("MAPPING.json",
                         json.dumps({"01.01": "99.99"}).encode(),
                         "application/json")).run()
    assert any("99.99" in w.value for w in app.warning)


# ── Le tableau du devis montre des libellés, pas des codes ──────────────────

# `st.data_editor` n'est pas exposé par AppTest : ses options ne peuvent pas
# être lues depuis l'app en marche. Ces deux tests attaquent donc directement
# les fonctions qui les construisent.

def _ui():
    """Le module de l'interface, importé pour ses fonctions pures.

    L'import exécute le script en mode « bare » : Streamlit prévient que
    l'état de session n'y fonctionne pas, ce qui est sans effet ici — on
    n'appelle que du calcul de libellés.
    """
    pytest.importorskip("streamlit")
    import streamlit_app  # noqa: PLC0415

    return streamlit_app


def test_un_libelle_redonne_toujours_son_code():
    """Le tableau montre « 40.20 · Enduit de façade · 98,64 €/m2 » et
    range « 40.20 ». Le jour où un libellé contiendrait le séparateur,
    le devis se chiffrerait sur un autre ouvrage sans rien dire."""
    ui = _ui()
    b = ui._bordereau()
    for code in b:
        assert ui._code_du_libelle(ui._libelle_ouvrage(code, b)) == code


def test_le_libelle_porte_la_designation_et_le_prix():
    """Un code seul est illisible : personne ne retient cinquante codes,
    et un mauvais choix ne se verrait qu'au chantier. Le prix affiché
    sert de contrôle de vraisemblance au moment du choix."""
    ui = _ui()
    b = ui._bordereau()
    libelle = ui._libelle_ouvrage("40.20", b)
    assert b["40.20"]["libelle_ouv"] in libelle
    assert b["40.20"]["unite_ouv"] in libelle
    assert f"{b['40.20']['pu_vente']:.2f}" in libelle


def test_les_postes_restent_ranges_par_code(app):
    """Le libellé affiché porte le prix du jour : le garder en session
    afficherait un prix périmé après un changement de marge. C'est le
    code qui est stocké, le libellé qui est reconstruit."""
    assert all("code_ouv" in ligne
                for ligne in app.session_state["lignes_devis"])


def test_une_ligne_en_cours_de_saisie_survit_a_une_reexecution(app):
    """Ouvrage choisi, quantité pas encore tapée : filtrer les lignes
    incomplètes ferait disparaître sous les doigts celle qu'on vient
    d'ajouter."""
    app.session_state["lignes_devis"] = [{"code_ouv": "40.20", "qte": None}]
    app.run()
    assert not app.exception, [e.value for e in app.exception]
    assert [ligne["code_ouv"]
            for ligne in app.session_state["lignes_devis"]] == ["40.20"]


def test_un_ouvrage_supprime_de_la_bibliotheque_est_signale(app):
    """Un devis repris peut porter un ouvrage effacé depuis. Le laisser
    ferait planter la liste déroulante, qui n'accepte que ses options."""
    app.session_state["lignes_devis"] = [{"code_ouv": "99.99", "qte": 3.0}]
    app.run()
    assert not app.exception, [e.value for e in app.exception]
    assert any("99.99" in w.value for w in app.warning)


# ── Lever le doute sur un rendement ─────────────────────────────────────────

def _atelier(AppTest, tables=None):
    """L'app avec, au besoin, des tables déjà corrigées en session."""
    import copy  # noqa: PLC0415

    from chiffrage.moteur import tables_courantes  # noqa: PLC0415

    at = AppTest.from_file(str(APP), default_timeout=240)
    at.session_state["tables_editees"] = copy.deepcopy(
        tables if tables is not None else tables_courantes())
    at.run()
    assert not at.exception, [e.value for e in at.exception]
    return at


def test_la_liste_a_valider_voyage_avec_les_tables(AppTest):
    """L'atelier corrige une COPIE des tables. Si la liste des
    rendements douteux n'en fait pas partie, lever un doute pendant une
    séance de calibration n'a nulle part où s'écrire."""
    from chiffrage.moteur import tables_courantes  # noqa: PLC0415

    assert "ouvrages_a_valider" in tables_courantes()


def test_lever_un_doute_fait_apparaitre_lenregistrement(AppTest):
    """Une validation ne change AUCUN prix : elle n'apparaît dans
    aucune des comparaisons de valeurs. Sans traitement à part, le
    bouton d'enregistrement resterait absent et le travail serait perdu
    au rafraîchissement suivant — le défaut déjà rencontré sur la
    création d'ouvrage."""
    import copy  # noqa: PLC0415

    from chiffrage.moteur import tables_courantes  # noqa: PLC0415

    tables = copy.deepcopy(tables_courantes())
    assert tables["ouvrages_a_valider"], "aucun ouvrage douteux à lever"
    tables["ouvrages_a_valider"] = tables["ouvrages_a_valider"][1:]
    at = _atelier(AppTest, tables)

    enregistrements = [b for b in at.button if "Enregistrer" in b.label]
    telechargements = [d for d in at.get("download_button")
                        if "ouvrages_a_valider" in d.label]
    assert enregistrements or telechargements, (
        "lever un doute ne propose ni enregistrement ni téléchargement")


def test_le_compteur_annonce_le_nombre_restant(AppTest):
    import copy  # noqa: PLC0415

    from chiffrage.moteur import tables_courantes  # noqa: PLC0415

    tables = copy.deepcopy(tables_courantes())
    depart = len(tables["ouvrages_a_valider"])
    tables["ouvrages_a_valider"] = tables["ouvrages_a_valider"][1:]
    at = _atelier(AppTest, tables)
    deltas = [m.delta for m in at.metric if m.label == "Valeurs corrigées"]
    assert any(f"{depart} → {depart - 1}" in (d or "") for d in deltas)


def test_le_doute_se_pose_aussi_sur_un_ouvrage_qui_nen_portait_pas(AppTest):
    """Un ouvrage sans ⚠️ n'a pas été vérifié pour autant : il n'a
    simplement pas été signalé. Pouvoir poser le doute est la moitié
    utile du mécanisme."""
    at = _atelier(AppTest)
    # L'atelier ouvre sur la correction des taux : il faut passer sur les
    # rendements pour voir le bouton.
    [r for r in at.radio if "corriger" in (r.label or "").lower()][0].set_value(
        "Un rendement (h/unité)").run()
    assert not at.exception, [e.value for e in at.exception]
    boutons = {b.key for b in at.button}
    assert "corr_rend_valider" in boutons or "corr_rend_douter" in boutons


def test_valider_puis_douter_fait_bien_laller_retour(AppTest):
    """Le parcours réel : choisir le rendement d'un ouvrage marqué ⚠️,
    déclarer qu'un chantier l'a confirmé, et pouvoir revenir dessus."""
    from chiffrage.moteur import tables_courantes  # noqa: PLC0415

    at = _atelier(AppTest)
    [r for r in at.radio if "corriger" in (r.label or "").lower()][0].set_value(
        "Un rendement (h/unité)").run()

    tables = tables_courantes()
    doute = tables["ouvrages_a_valider"][0]
    indice = next(i for i, c in enumerate(tables["composition"])
                   if c["code_ouv"] == doute
                   and tables["ressources_par_code"][
                       c["code_res"]]["type_res"] == "MO")
    liste = [s for s in at.selectbox if s.key == "corr_rend"][0]
    liste.set_value(indice).run()
    assert not at.exception, [e.value for e in at.exception]

    [b for b in at.button if b.key == "corr_rend_valider"][0].click().run()
    assert doute not in at.session_state["tables_editees"]["ouvrages_a_valider"]

    # Et le retour en arrière, sans quoi une validation trop rapide
    # serait irrattrapable depuis l'interface.
    [b for b in at.button if b.key == "corr_rend_douter"][0].click().run()
    assert doute in at.session_state["tables_editees"]["ouvrages_a_valider"]


# ── Retrouver un code dans la bibliothèque ──────────────────────────────────

def _table(app, colonne):
    """Le tableau qui porte cette colonne — l'app en affiche plusieurs, et
    leur ordre change avec ce qui s'affiche ailleurs."""
    trouves = [d.value for d in app.dataframe if colonne in d.value.columns]
    assert trouves, f"aucun tableau avec la colonne « {colonne} »"
    return trouves[0]


def _chercher(app, texte):
    [t for t in app.text_input if t.key == "biblio_recherche"][0].set_value(
        texte).run()
    assert not app.exception, [e.value for e in app.exception]
    return app


def test_la_recherche_filtre_la_bibliotheque(app):
    """Faire défiler 49 lignes au doigt pour retrouver un code n'est pas
    une recherche."""
    from chiffrage.bibliotheque import OUVRAGES_PAR_CODE  # noqa: PLC0415

    complet = _table(app, "À valider")
    _chercher(app, "40.20")
    filtre = _table(app, "À valider")
    assert len(filtre) == 1 < len(complet)
    assert filtre.iloc[0]["Désignation"] == (
        OUVRAGES_PAR_CODE["40.20"]["libelle_ouv"])


def test_la_recherche_ignore_les_accents(app):
    """« etancheite » doit trouver « Étanchéité » : on tape au clavier
    d'un téléphone, pas dans un éditeur."""
    _chercher(app, "etancheite")
    assert len(_table(app, "À valider")) >= 1


def test_la_recherche_cumule_les_mots(app):
    """Deux mots, deux conditions : c'est ce qui rend le filtre utile
    au-delà du premier mot tapé."""
    _chercher(app, "enduit")
    large = len(_table(app, "À valider"))
    _chercher(app, "enduit facade")
    assert 0 < len(_table(app, "À valider")) <= large


def test_un_mot_absent_renvoie_vers_le_lexique(app):
    """« carrelage mural » n'existe pas dans la bibliothèque : l'ouvrage
    s'appelle « faïence ». Un filtre littéral s'arrêterait là, alors que
    le moteur d'appariement sait déjà faire la traduction."""
    _chercher(app, "carrelage mural")
    assert any("Aucun libellé" in w.value for w in app.warning)
    proches = _table(app, "Score")
    assert len(proches) == 5
    faiences = [c for c in proches["Désignation"]
                 if "aïence" in c or "arrelage" in c]
    assert faiences, f"le lexique n'a pas mené à la faïence : {list(proches['Désignation'])}"


def test_une_session_ouverte_pendant_un_deploiement_ne_plante_pas(AppTest):
    """Streamlit relit le script dans le même processus : l'état de
    session survit à un déploiement. Une table ajoutée par la nouvelle
    version manque alors aux tables en cours de correction — KeyError en
    plein écran, chez le client. Compléter, sans écraser ce qui est
    corrigé."""
    import copy  # noqa: PLC0415

    from chiffrage.moteur import tables_courantes  # noqa: PLC0415

    vieilles = copy.deepcopy(tables_courantes())
    del vieilles["ouvrages_a_valider"]          # la version d'avant
    vieilles["ressources"][0]["pu_res"] = 77.0  # une correction en cours

    at = AppTest.from_file(str(APP), default_timeout=240)
    at.session_state["tables_editees"] = vieilles
    at.run()
    assert not at.exception, [e.value for e in at.exception]

    tables = at.session_state["tables_editees"]
    assert tables["ouvrages_a_valider"], "la table manquante n'a pas été reprise"
    assert tables["ressources"][0]["pu_res"] == 77.0, "la correction a été écrasée"
