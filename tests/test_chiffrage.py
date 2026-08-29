"""Tests de l'outil de chiffrage BAG BATTER (`chiffrage/`).

Deux niveaux :
  · le moteur en Python pur — toujours exécuté ;
  · les modules Excel (export, génération de métré, remplissage d'offre) —
    ignorés si openpyxl n'est pas installé (il ne fait pas partie des
    dépendances de l'app, seulement de `requirements-chiffrage.txt`).

Ce que ces tests protègent VRAIMENT, au-delà de la couverture :
  1. le coefficient K et la formule de prix (une dérive silencieuse ici fausse
     toutes les offres) ;
  2. l'intégrité référentielle des trois tables ;
  3. le contrôle d'unité, qui doit refuser de chiffrer plutôt que convertir ;
  4. l'ancrage des formules Excel ligne par ligne — le bug openpyxl
     `insert_rows` décrit dans gen_metre.py.
"""
import pytest

from chiffrage import bibliotheque as biblio
from chiffrage import moteur


# ── 1. Structure de la bibliothèque ────────────────────────────────────────
def test_tailles_des_tables():
    assert len(biblio.RESSOURCES) == 39
    assert len(biblio.OUVRAGES) == 36
    assert len(biblio.LOTS) == 8
    assert len(biblio.NON_COUVERTS) == 13


def test_bibliotheque_coherente():
    """Aucune référence orpheline, aucun ouvrage sans composition ni sans MO."""
    anomalies = moteur.controle_coherence()
    assert anomalies == {
        "res_orphelines": [],
        "ouv_orphelins": [],
        "ouv_sans_compo": [],
        "ouv_sans_mo": [],
        "codes_dupliques": [],
        "lots_inconnus": [],
    }


def test_codes_ouvrages_au_format_lot_numero():
    for ouv in biblio.OUVRAGES:
        lot, numero = ouv["code_ouv"].split(".")
        assert lot in biblio.LOTS and lot == ouv["lot"]
        assert len(numero) == 2 and numero.isdigit()


def test_mapping_ne_pointe_que_des_ouvrages_existants():
    codes = {o["code_ouv"] for o in biblio.OUVRAGES}
    assert set(biblio.MAPPING.values()) <= codes


def test_espaces_de_nommage_disjoints():
    """Un code de poste de métré (NN.NN) ne doit jamais être un code d'ouvrage."""
    postes = set(biblio.MAPPING) | {code for code, _, _ in biblio.NON_COUVERTS}
    ouvrages = {o["code_ouv"] for o in biblio.OUVRAGES}
    assert postes & ouvrages == set()


# ── 2. Formule de prix ─────────────────────────────────────────────────────
def test_coefficient_k():
    assert round(moteur.coefficient_k(), 4) == 1.3324


def test_pu_vente_est_le_debourse_fois_k():
    k = moteur.coefficient_k()
    for ligne in moteur.calcul_bordereau().values():
        assert ligne["pu_vente"] == pytest.approx(ligne["debourse_sec"] * k, abs=0.01)
        assert ligne["debourse_sec"] == pytest.approx(
            ligne["deb_mo"] + ligne["deb_mat"] + ligne["deb_eqp"], abs=0.01
        )
        assert ligne["pu_vente"] > 0


def test_marge_nulle_ramene_le_prix_au_debourse_majore():
    params = dict(biblio.PARAMS, fg=0.0, fc=0.0, aleas=0.0, marge=0.0)
    for ligne in moteur.calcul_bordereau(params).values():
        assert ligne["pu_vente"] == pytest.approx(ligne["debourse_sec"], abs=0.01)


# ── 3. Devis ───────────────────────────────────────────────────────────────
def test_devis_totalise_et_applique_la_tva():
    d = moteur.devis("test", [("70.10", 100.0)], tva=0.21)
    pu = moteur.calcul_bordereau()["70.10"]["pu_vente"]
    assert d["total_ht"] == pytest.approx(pu * 100, abs=0.01)
    assert d["montant_tva"] == pytest.approx(d["total_ht"] * 0.21, abs=0.01)
    assert d["total_ttc"] == pytest.approx(d["total_ht"] * 1.21, abs=0.02)
    assert d["jours_homme"] == pytest.approx(d["heures_mo"] / 8, abs=0.01)


def test_devis_signale_les_codes_inconnus_sans_les_chiffrer():
    """Un code absent ne doit JAMAIS disparaître en silence."""
    d = moteur.devis("test", [("70.10", 10.0), ("99.99", 5.0)])
    assert d["inconnus"] == ["99.99"]
    assert len(d["lignes"]) == 1


def test_tva_marche_public_a_21_pourcent():
    assert biblio.PARAMS["tva_marche_public"] == 0.21


# ── 4. Calibration sur les devis historiques ───────────────────────────────
def test_calibration_couvre_les_six_devis():
    cal = moteur.calibration()
    assert len(cal["lignes"]) == 6
    assert all(not r["inconnus"] for r in cal["lignes"])
    assert all(r["calcule"] > 0 for r in cal["lignes"])


def test_fiche_prix_detaille_toutes_les_ressources():
    texte = moteur.fiche_prix("40.20")
    for comp in biblio.COMPOSITION:
        if comp["code_ouv"] == "40.20":
            assert comp["code_res"] in texte
    assert "DÉBOURSÉ SEC" in texte and "1.3324" in texte


def test_fiche_prix_refuse_un_ouvrage_inconnu():
    with pytest.raises(KeyError):
        moteur.fiche_prix("99.99")


# ── 5. Chaîne Excel (nécessite openpyxl) ───────────────────────────────────
@pytest.fixture(scope="module")
def openpyxl_dispo():
    return pytest.importorskip("openpyxl")


@pytest.fixture(scope="module")
def metre(tmp_path_factory, openpyxl_dispo):
    from chiffrage.gen_metre import generer_metre

    chemin = tmp_path_factory.mktemp("chiffrage") / "metre.xlsx"
    generer_metre(str(chemin))
    return chemin


def test_normalisation_des_unites(openpyxl_dispo):
    from chiffrage.metre_io import normaliser_unite

    assert normaliser_unite("m²") == normaliser_unite("M2") == "m2"
    assert normaliser_unite("PCE") == normaliser_unite("pc") == "pce"
    # Aucune conversion physique : le mètre courant n'est pas le mètre carré.
    assert normaliser_unite("m") != normaliser_unite("m2")


def test_metre_genere_49_postes(metre):
    from chiffrage.metre_io import lire_metre

    postes = lire_metre(str(metre))
    assert len(postes) == 49
    assert len({p["code"] for p in postes}) == 49
    assert all(p["quantite"] > 0 for p in postes)


def test_formules_du_metre_ancrees_sur_leur_propre_ligne(metre):
    """Garde-fou contre le bug openpyxl `insert_rows` (cf. gen_metre.py).

    Une formule de montant doit référencer SA ligne, pas celle où elle a été
    écrite avant un décalage.
    """
    from openpyxl import load_workbook

    ws = load_workbook(str(metre)).active
    montants = 0
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith('=IF($G'):
                montants += 1
                assert v == f'=IF($G{cell.row}="","",$F{cell.row}*$G{cell.row})'
    assert montants == 49


def test_remplissage_de_l_offre(metre, tmp_path):
    from chiffrage.metre_io import remplir_metre

    rapport = remplir_metre(str(metre), str(tmp_path / "offre.xlsx"))
    assert rapport["postes"] == 49
    assert len(rapport["chiffres"]) == 36
    assert len(rapport["non_couverts"]) == 13
    # Contrôle d'unité : aucun écart sur ce métré-ci, et surtout aucun poste
    # chiffré dans une unité différente de celle de l'ouvrage.
    assert rapport["ecarts_unite"] == []
    assert rapport["total_ht"] > 0


def test_remplissage_preserve_les_formules_du_pouvoir_adjudicateur(metre, tmp_path):
    """`data_only=True` détruirait ces formules : le fichier renvoyé au PA ne
    recalculerait plus rien. On vérifie qu'elles survivent au remplissage."""
    from openpyxl import load_workbook

    from chiffrage.metre_io import remplir_metre

    sortie = tmp_path / "offre.xlsx"
    remplir_metre(str(metre), str(sortie))
    ws = load_workbook(str(sortie)).active
    formules = [
        c.value
        for row in ws.iter_rows()
        for c in row
        if isinstance(c.value, str) and c.value.startswith("=")
    ]
    assert sum(1 for f in formules if f.startswith("=IF($G")) == 49
    assert any(f.startswith("=SUM(H") for f in formules)


def test_unite_incompatible_bloque_le_chiffrage(metre, tmp_path):
    """Mapper un poste au m2 vers un ouvrage au mètre courant ne doit PAS
    produire de prix : le poste part en `ecarts_unite`, pas en `chiffres`."""
    from chiffrage.metre_io import remplir_metre

    mapping_faux = dict(biblio.MAPPING)
    mapping_faux["03.02"] = "40.60"  # 03.02 est en m2, 40.60 en m
    rapport = remplir_metre(
        str(metre), str(tmp_path / "offre_ko.xlsx"), mapping=mapping_faux
    )
    assert [e["code"] for e in rapport["ecarts_unite"]] == ["03.02"]
    assert "03.02" not in [c["code"] for c in rapport["chiffres"]]
    assert "03.02" in rapport["vides"]


def test_export_bibliotheque_six_onglets(tmp_path, openpyxl_dispo):
    from openpyxl import load_workbook

    from chiffrage.export_xlsx import exporter_bibliotheque

    chemin = tmp_path / "biblio.xlsx"
    exporter_bibliotheque(str(chemin))
    wb = load_workbook(str(chemin))
    assert wb.sheetnames == [
        "PARAMS", "RESSOURCES", "OUVRAGES", "COMPOSITION", "BORDEREAU", "MAPPING",
    ]
    assert wb["RESSOURCES"].max_row == len(biblio.RESSOURCES) + 1
    assert wb["COMPOSITION"].max_row == len(biblio.COMPOSITION) + 1


def test_formules_excel_du_bordereau_donnent_les_prix_python(tmp_path, openpyxl_dispo):
    """Le classeur exporté calcule tout seul (SUMIFS + K) : on ré-exécute la
    sémantique de ses formules en Python et on la compare à `pu_vente`.

    Sans ça, une erreur de plage dans l'export ne se verrait qu'à l'ouverture
    du fichier par le client — c'est-à-dire jamais avant l'envoi d'une offre.
    """
    from openpyxl import load_workbook

    from chiffrage.export_xlsx import exporter_bibliotheque

    chemin = tmp_path / "biblio.xlsx"
    exporter_bibliotheque(str(chemin))
    wb = load_workbook(str(chemin))
    compo = wb["COMPOSITION"]
    lignes = [
        (
            compo.cell(r, 1).value,   # code_ouv
            compo.cell(r, 5).value,   # type_res
            compo.cell(r, 7).value,   # qte_res
            wb["RESSOURCES"].cell(
                int(compo.cell(r, 8).value.split("!E")[1]), 5
            ).value,                  # pu_res, suivi à travers la formule
        )
        for r in range(2, compo.max_row + 1)
    ]
    k = moteur.coefficient_k()
    bordereau = wb["BORDEREAU"]
    for r in range(2, bordereau.max_row + 1):
        code = bordereau.cell(r, 1).value
        debourse = sum(q * pu for c, _, q, pu in lignes if c == code)
        # colonne M = valeur calculée par Python, écrite comme témoin
        assert bordereau.cell(r, 13).value == pytest.approx(debourse * k, abs=0.01)
