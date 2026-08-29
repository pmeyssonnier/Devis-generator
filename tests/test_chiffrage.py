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
    assert len(biblio.RESSOURCES) == 49
    assert len(biblio.OUVRAGES) == 49
    assert len(biblio.LOTS) == 9
    assert len(biblio.OUVRAGES_A_VALIDER) == 13


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
    ouvrages = {o["code_ouv"] for o in biblio.OUVRAGES}
    assert set(biblio.MAPPING) & ouvrages == set()


def test_ouvrages_a_valider_existent_et_sont_mappes():
    """Les 13 ouvrages créés après coup doivent exister ET couvrir un poste :
    sinon on a fabriqué des prix pour rien."""
    codes = {o["code_ouv"] for o in biblio.OUVRAGES}
    assert set(biblio.OUVRAGES_A_VALIDER) <= codes
    assert set(biblio.OUVRAGES_A_VALIDER) <= set(biblio.MAPPING.values())


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
    # Les 49 postes sont chiffrés : c'est la condition de RÉGULARITÉ de
    # l'offre (art. 76 AR 18/04/2017). Ce test est le garde-fou du seul
    # défaut qui fait rejeter une offre sans discussion.
    assert len(rapport["chiffres"]) == 49
    assert rapport["non_couverts"] == []
    assert rapport["vides"] == []
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


# ── 6. Devis client au format Excel ────────────────────────────────────────
@pytest.fixture
def devis_exemple():
    return moteur.devis(
        "Rénovation façade arrière",
        [("40.20", 26.0), ("40.30", 26.0), ("70.70", 8.0)],
        tva=0.06,
    )


def test_devis_client_ecrit_un_classeur(devis_exemple, tmp_path, openpyxl_dispo):
    from openpyxl import load_workbook

    from chiffrage.devis_xlsx import exporter_devis

    chemin = tmp_path / "devis.xlsx"
    _, nb = exporter_devis(devis_exemple, str(chemin), reference="2026-042",
                           client="M. Dupont", chantier="Av. Renan 62")
    assert nb == 3
    ws = load_workbook(str(chemin))["DEVIS"]
    textes = [
        c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str)
    ]
    assert any(biblio.ENTREPRISE["nom"] in t for t in textes)
    assert any("2026-042" in t for t in textes)
    assert any("TOTAL À PAYER" in t for t in textes)


def test_devis_client_formules_ancrees_sur_leur_ligne(devis_exemple, tmp_path,
                                                      openpyxl_dispo):
    """Même garde-fou que pour le métré : chaque montant multiplie SA ligne."""
    from openpyxl import load_workbook

    from chiffrage.devis_xlsx import exporter_devis

    chemin = tmp_path / "devis.xlsx"
    exporter_devis(devis_exemple, str(chemin))
    ws = load_workbook(str(chemin))["DEVIS"]
    montants = 0
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("=ROUND(D"):
                montants += 1
                assert cell.value == f"=ROUND(D{cell.row}*E{cell.row},2)"
    assert montants == 3


def test_devis_client_totaux_excel_egalent_le_moteur(devis_exemple, tmp_path,
                                                     openpyxl_dispo):
    """Ré-exécution en Python de la sémantique des formules du classeur.

    Une erreur de plage dans un sous-total ne se verrait qu'à l'ouverture du
    fichier par le client — c'est-à-dire après l'envoi.
    """
    import re

    from openpyxl import load_workbook

    from chiffrage.devis_xlsx import exporter_devis

    chemin = tmp_path / "devis.xlsx"
    exporter_devis(devis_exemple, str(chemin))
    ws = load_workbook(str(chemin))["DEVIS"]

    valeurs = {}
    sous_totaux = {}
    for row in ws.iter_rows():
        cell = row[5]                       # colonne F
        if not isinstance(cell.value, str):
            continue
        if cell.value.startswith("=ROUND(D"):
            valeurs[cell.row] = round(ws.cell(cell.row, 4).value
                                      * ws.cell(cell.row, 5).value, 2)
        elif cell.value.startswith("=SUM(F"):
            debut, fin = map(int, re.findall(r"F(\d+)", cell.value))
            sous_totaux[cell.row] = sum(
                valeurs[r] for r in range(debut, fin + 1) if r in valeurs
            )

    total_ht = round(sum(sous_totaux.values()), 2)
    assert total_ht == pytest.approx(devis_exemple["total_ht"], abs=0.01)
    assert round(total_ht * 0.06, 2) == pytest.approx(
        devis_exemple["montant_tva"], abs=0.01
    )


def test_devis_client_refuse_un_devis_incomplet(tmp_path, openpyxl_dispo):
    """Un devis qui a laissé tomber un code inconnu ne doit PAS partir chez le
    client : le poste manquant ne se verrait qu'à la facturation."""
    from chiffrage.devis_xlsx import exporter_devis

    d = moteur.devis("test", [("40.20", 10.0), ("99.99", 5.0)])
    with pytest.raises(ValueError, match="99.99"):
        exporter_devis(d, str(tmp_path / "devis.xlsx"))
    assert not (tmp_path / "devis.xlsx").exists()


def test_devis_client_refuse_un_devis_vide(tmp_path, openpyxl_dispo):
    from chiffrage.devis_xlsx import exporter_devis

    with pytest.raises(ValueError):
        exporter_devis(moteur.devis("vide", []), str(tmp_path / "devis.xlsx"))


# ── 7. Notebook Colab ─────────────────────────────────────────────
#
# Le notebook est le point d'entrée réel du chef d'entreprise (il travaille
# depuis mobile, via Colab). Rien ne l'exécute en CI : sans ces tests, un
# renommage dans `chiffrage/` le casserait en silence et personne ne le
# saurait avant qu'il ouvre Colab pour répondre à un marché.
import json
from pathlib import Path

NOTEBOOK = Path(__file__).resolve().parent.parent / "colab" / "chiffrage_bagbatter.ipynb"


def _cellules_de_code():
    nb = json.loads(NOTEBOOK.read_text(encoding="utf-8"))
    return [
        "".join(c["source"])
        for c in nb["cells"]
        if c["cell_type"] == "code"
    ]


def test_notebook_est_un_json_valide():
    nb = json.loads(NOTEBOOK.read_text(encoding="utf-8"))
    assert nb["nbformat"] == 4
    assert len(_cellules_de_code()) >= 10


def test_cellules_du_notebook_compilent():
    """Syntaxe valide — hors cellules à magies shell (!pip, !git), qui ne
    sont pas du Python et ne compilent que dans IPython."""
    for i, source in enumerate(_cellules_de_code(), start=1):
        if any(ligne.lstrip().startswith("!") for ligne in source.splitlines()):
            continue
        compile(source, f"<cellule {i}>", "exec")


def test_symboles_importes_par_le_notebook_existent():
    """Le vrai garde-fou : chaque `from chiffrage.x import y` du notebook
    doit résoudre. C'est ce qui casse quand on renomme une fonction."""
    import importlib
    import re

    for source in _cellules_de_code():
        for module, noms in re.findall(
            r"from (chiffrage[\w.]*) import ([^\n(]+|\([^)]*\))", source
        ):
            mod = importlib.import_module(module)
            for nom in noms.strip("()").replace("\n", " ").split(","):
                nom = nom.strip().split(" as ")[0].strip()
                if nom:
                    assert hasattr(mod, nom), f"{module}.{nom} n'existe pas"


# ── 8. Appariement automatique ───────────────────────────────
from chiffrage import suggestion  # noqa: E402


def test_normalisation_de_libelle():
    mots = suggestion.normaliser("Étanchéité bitumineuse, relevés compris")
    assert "etancheite" in mots
    assert "bitumineuse" in mots
    # « compris » est un mot vide de métré : le garder ferait remonter
    # n'importe quel poste qui le contient.
    assert "compris" not in mots


def test_score_est_maximal_pour_un_libelle_identique():
    lib = "Enduit de façade minéral armé, deux couches"
    assert suggestion.score(lib, lib) > suggestion.score(lib, "Peinture de façade")


def test_unite_est_eliminatoire_pas_departageante():
    """LE garde-fou du module. Un poste au mètre courant ne doit jamais
    se voir proposer un ouvrage au m2, même si les libellés sont
    identiques : le prix serait faux d'un facteur inconnu."""
    b = moteur.calcul_bordereau()
    poste = {
        "designation": "Enduit de façade minéral armé, deux couches",  # = 40.20
        "unite": "m",                                                # mais en m !
    }
    candidats = suggestion.suggerer(poste, b, limite=10)
    assert "40.20" not in [c for c, _ in candidats]
    assert all(b[c]["unite_ouv"] == "m" for c, _ in candidats)


def test_suggerer_rend_une_liste_vide_si_aucune_unite_ne_correspond():
    b = moteur.calcul_bordereau()
    assert suggestion.suggerer(
        {"designation": "Quoi que ce soit", "unite": "tonne"}, b
    ) == []


def test_proposer_mapping_reprend_les_correspondances_connues():
    """Un choix humain antérieur n'est jamais réécrit par l'algorithme."""
    b = moteur.calcul_bordereau()
    postes = [{"code": "03.02", "designation": "n'importe quoi", "unite": "m2"}]
    prop = suggestion.proposer_mapping(postes, b, mapping_connu=biblio.MAPPING)
    assert prop["03.02"]["origine"] == "connu"
    assert prop["03.02"]["code_ouv"] == biblio.MAPPING["03.02"]


def test_proposer_mapping_signale_ce_qu_il_ne_sait_pas_apparier():
    b = moteur.calcul_bordereau()
    postes = [{"code": "99.99", "designation": "Zzzz qqqq", "unite": "tonne"}]
    prop = suggestion.proposer_mapping(postes, b)
    assert prop["99.99"]["origine"] == "aucun"
    assert prop["99.99"]["code_ouv"] is None


@pytest.mark.parametrize("libelle, unite, attendu", [
    ("Montage, location et démontage d'un échafaudage de pied", "m2", "10.20"),
    ("Membrane d'étanchéité soudée en deux couches", "m2", "40.40"),
    ("Menuiserie extérieure PVC avec double vitrage", "m2", "80.10"),
    ("Nettoyage de parement par projection d'eau sous pression", "m2", "40.10"),
])
def test_appariement_sur_des_libelles_reformules(libelle, unite, attendu):
    """Libellés réécrits « à la manière d'un CSC » — synonymes, ordre
    inversé, jargon administratif — sans mot commun au-delà du sens.

    Sur un échantillon plus large de ce type, l'appariement tombe juste
    du premier coup dans ~60 % des cas et place la bonne réponse dans les
    trois premières dans ~90 %. C'est un dégrossissage, pas une décision :
    l'interface fait trancher l'humain. Ces quatre cas-ci sont ceux qui
    doivent rester au premier rang.
    """
    b = moteur.calcul_bordereau()
    candidats = suggestion.suggerer(
        {"designation": libelle, "unite": unite}, b, limite=1
    )
    assert candidats and candidats[0][0] == attendu


# ── 9. Surcouche de session du lexique ─────────────────
@pytest.fixture(autouse=True)
def _lexique_propre():
    """Isole CHAQUE test des deux couches mutables du lexique.

    La surcouche est un état global de module : sans nettoyage, un
    test qui ajoute un synonyme fausserait les suivants.

    La couche LOCALE, elle, est chargée depuis lexique_local.json —
    un fichier que l'APP écrit en production. Un test qui en dépend
    se met donc à échouer le jour où quelqu'un apprend un terme
    depuis l'interface, sans que le code ait bougé. C'est
    exactement ce qui est arrivé avec « sablage » : la suite est
    devenue rouge à cause d'une donnée, pas d'une régression.

    Les tests qui veulent une couche locale la posent eux-mêmes.
    """
    from chiffrage.lexique import LOCAL, vider_surcouche

    garde = {table: dict(valeurs) for table, valeurs in LOCAL.items()}
    for valeurs in LOCAL.values():
        valeurs.clear()
    vider_surcouche()
    yield
    vider_surcouche()
    for table, valeurs in garde.items():
        LOCAL[table] = valeurs


def test_surcouche_change_l_appariement_immediatement():
    """C'est tout l'intérêt : ajouter un terme dans l'interface et en
    voir l'effet sans redéployer."""
    from chiffrage import lexique

    b = moteur.calcul_bordereau()
    poste = {"designation": "Sablage des maçonneries de façade",
             "unite": "m2"}

    avant = suggestion.suggerer(poste, b, limite=1)[0][1]
    lexique.ajouter_synonyme("sablage", "nettoyage")
    apres = suggestion.suggerer(poste, b, limite=1)[0][1]

    assert avant < 0.40 < apres


def test_surcouche_l_emporte_sur_la_table_du_depot():
    from chiffrage import lexique

    assert lexique.canoniser("crepi") == "enduit"
    lexique.ajouter_synonyme("crepi", "plafonnage")
    assert lexique.canoniser("crepi") == "plafonnage"


def test_surcouche_se_vide():
    from chiffrage import lexique

    lexique.ajouter_synonyme("sablage", "nettoyage")
    lexique.ajouter_expression("mur de refend", "cloison")
    lexique.vider_surcouche()
    assert lexique.canoniser("sablage") == "sablage"
    assert lexique.appliquer_expressions("mur de refend") == "mur de refend"


def test_surcouche_rend_un_bloc_python_collable():
    """L'interface ne peut pas commiter : elle rend le bloc à coller.
    Ce bloc doit être du Python valide, sinon il ne sert à rien."""
    import ast

    from chiffrage import lexique

    lexique.ajouter_synonyme("sablage", "nettoyage")
    lexique.ajouter_expression("mur de refend", "cloison")
    bloc = lexique.surcouche_en_python()

    assert '"sablage": "nettoyage",' in bloc
    assert '"mur de refend": "cloison",' in bloc
    # Collé dans un dict, ça doit se parser.
    ast.parse("SYNONYMES = {\n" + bloc.replace("#", "  #") + "\n}")


# ── 10. Persistance du lexique dans le dépôt ─────────
def test_terme_accentue_est_normalise():
    """Défaut corrigé : les tables sont consultées avec des mots
    DÉPOUILLÉS de leurs accents. Une clé accentuée n'était jamais
    trouvée — l'ajout ne servait à rien, en silence."""
    from chiffrage import lexique

    lexique.ajouter_synonyme("Maçonneries", "maconnerie")
    assert "maconneries" in lexique.SURCOUCHE["synonymes"]
    assert "maconnerie" in suggestion.normaliser("Sablage des maçonneries")


@pytest.mark.parametrize("terme", [
    'x"; import os',          # injection : le terme finissait dans du code
    "trop " * 20,             # longueur
    "",                       # vide
    "<script>",               # balise
    "terme; DROP",           # ponctuation d'injection
])
def test_terme_irrecevable_est_refuse(terme):
    from chiffrage import lexique

    with pytest.raises(ValueError):
        lexique.ajouter_synonyme(terme, "nettoyage")


def test_saut_de_ligne_est_recolle_pas_refuse():
    """Coller un libellé depuis un PDF amène des sauts de ligne. Ils
    deviennent des espaces — une expression multi-mots est valide, et
    un saut de ligne ne peut rien casser dans du JSON."""
    from chiffrage import lexique

    lexique.ajouter_expression("carrelage\nmural", "faience")
    assert "carrelage mural" in lexique.SURCOUCHE["expressions"]


def test_couche_locale_absente_ne_casse_rien(tmp_path):
    """Un lexique_local.json absent est le cas NORMAL : l'outil doit
    chiffrer sans lui."""
    from chiffrage import lexique

    assert lexique.charger_local(tmp_path / "rien.json") == {
        "expressions": {}, "synonymes": {}}


def test_couche_locale_illisible_ne_casse_rien(tmp_path):
    """Un fichier corrompu ne doit pas empêcher de répondre à un
    marché : on revient aux tables du dépôt."""
    from chiffrage import lexique

    fichier = tmp_path / "casse.json"
    fichier.write_text("{ceci n'est pas du json", encoding="utf-8")
    assert lexique.charger_local(fichier) == {
        "expressions": {}, "synonymes": {}}


def test_couche_locale_est_normalisee_au_chargement(tmp_path):
    import json

    from chiffrage import lexique

    fichier = tmp_path / "lex.json"
    fichier.write_text(
        json.dumps({"synonymes": {"Sablage ": "Nettoyage"}}), encoding="utf-8")
    assert lexique.charger_local(fichier)["synonymes"] == {
        "sablage": "nettoyage"}


def test_fusion_respecte_la_precedance():
    """dépôt < appris < à chaud : un essai en cours doit l'emporter
    sur ce qui est commité, sinon on ne peut rien corriger."""
    from chiffrage import lexique

    lexique.LOCAL["synonymes"]["sablage"] = "peinture"
    lexique.ajouter_synonyme("sablage", "nettoyage")
    try:
        fusion = lexique.fusion_a_commiter({"synonymes": {"sablage": "chape"}})
        assert fusion["synonymes"]["sablage"] == "nettoyage"
    finally:
        lexique.LOCAL["synonymes"].clear()


def test_commit_fusionne_avec_le_distant_et_n_ecrase_rien():
    """Deux personnes peuvent régler le lexique le même jour : un
    PUT sans fusion effacerait les termes de l'autre en silence."""
    import json

    from chiffrage import depot_github, lexique

    ecrit = {}

    def _lire(chemin, depot, token, branche):
        return json.dumps({"synonymes": {"pose-existant": "carrelage"}}), "sha1"

    def _ecrire(chemin, contenu, message, depot, token, branche, sha):
        ecrit.update(chemin=chemin, contenu=contenu, sha=sha, branche=branche)
        return "https://github.com/x/y/commit/abc"

    lexique.ajouter_synonyme("sablage", "nettoyage")
    fusion, url = depot_github.commiter_lexique(
        None, "moi/depot", "jeton", _lire=_lire, _ecrire=_ecrire)

    # Le terme de l'autre a survécu, le nôtre est arrivé.
    assert fusion["synonymes"] == {"pose-existant": "carrelage",
                                    "sablage": "nettoyage"}
    assert ecrit["sha"] == "sha1"      # écriture conditionnée à la version lue
    assert ecrit["chemin"].endswith("lexique_local.json")
    assert json.loads(ecrit["contenu"])["synonymes"]["sablage"] == "nettoyage"
    assert url.startswith("https://github.com/")


def test_commit_premier_fichier_sans_sha():
    """Au tout premier commit, le fichier n'existe pas : pas de sha."""
    from chiffrage import depot_github, lexique

    vu = {}

    def _ecrire(chemin, contenu, message, depot, token, branche, sha):
        vu["sha"] = sha
        return "url"

    lexique.ajouter_synonyme("sablage", "nettoyage")
    depot_github.commiter_lexique(
        None, "moi/depot", "jeton",
        _lire=lambda *a, **k: (None, None), _ecrire=_ecrire)
    assert vu["sha"] is None


def test_erreur_github_est_expliquee():
    """Un « 403 » brut n'apprend rien : le message doit dire quoi
    vérifier."""
    import urllib.error

    from chiffrage import depot_github

    def _appel(*a, **k):
        raise urllib.error.HTTPError("u", 403, "Forbidden", {}, None)

    with pytest.raises(depot_github.ErreurDepot) as err:
        depot_github.lire_fichier("f", "moi/depot", "jeton", _appel=_appel)
    assert "Contents: read and write" in str(err.value)


def test_fichier_absent_nest_pas_une_erreur():
    import urllib.error

    from chiffrage import depot_github

    def _appel(*a, **k):
        raise urllib.error.HTTPError("u", 404, "Not Found", {}, None)

    assert depot_github.lire_fichier(
        "f", "moi/depot", "jeton", _appel=_appel) == (None, None)
