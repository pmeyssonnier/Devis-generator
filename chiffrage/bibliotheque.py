"""
╔══════════════════════════════════════════════════════╗
║  BAG BATTER SRL — BIBLIOTHÈQUE DE PRIX UNITAIRES                      ║
║  Chargement et contrôle des tables                                    ║
╚══════════════════════════════════════════════════════╝

Les données vivent dans `chiffrage/data/*.json`, plus dans du Python.

Pourquoi : ce sont des valeurs d'entreprise — prix d'achat, taux
horaires, rendements — que le chef d'entreprise est seul à connaître et
qu'il faut pouvoir corriger sans éditer de code. C'était le principal
obstacle pratique à la calibration : les bons chiffres finissaient dans
un classeur Excel, et il fallait les recopier à la main dans le Python.
Du JSON se relit dans un diff GitHub, se corrige depuis l'interface, et
ne s'exécute pas.

    data/lots.json                un lot = un code à deux chiffres
    data/ressources.json          MO, matériaux, matériel + prix d'achat
    data/ouvrages.json            les postes chiffrables, leur unité
    data/composition.json         ce que consomme chaque ouvrage
                                  ⚠️ les lignes MO portent le RENDEMENT
    data/mapping.json             postes d'un métré -> ouvrages
    data/ouvrages_a_valider.json  les rendements jamais confrontés au réel
    data/releves.json             ce qu'on a VU sur un chantier (optionnelle)
    data/validations.json         un rendement confronté au réel (optionnelle)
    data/metres_histo.json        les six devis vendus, pour la calibration

L'identité de l'entreprise et les coefficients de vente sont ailleurs
encore (`parametres.py`) : ils se règlent depuis l'interface.

──────────────────────────────────────────────────────
⚠️  ÉTAT DES DONNÉES — À LIRE AVANT UTILISATION COMMERCIALE
──────────────────────────────────────────────────────
Le notebook Colab d'origine a été perdu ; ces tables en sont une
RECONSTRUCTION. La structure est fidèle au document de reprise, mais les
VALEURS — prix d'achat et surtout rendements main-d'œuvre — ont été
re-saisies à partir d'ordres de grandeur du marché belge 2026.

Treize ouvrages ont été ajoutés ensuite (voir OUVRAGES_A_VALIDER) pour
couvrir des postes qui restaient sans prix et rendaient toute offre
irrégulière. Leurs rendements sont les moins assis de tous.

Ce sont donc des HYPOTHÈSES DE DÉPART, pas les chiffres calibrés du
client. Point d'entrée pour la relecture : `moteur.calibration()`.

──────────────────────────────────────────────────────
MODÈLE
──────────────────────────────────────────────────────

    VALIDATIONS  (code_ouv, date, rendement, n, quantite, heures, note)
                 Un rendement CONFRONTÉ au réel. Le champ qui compte est
                 `rendement` : une validation porte sur une VALEUR, pas
                 sur un ouvrage. Que la valeur bouge, et la validation ne
                 vaut plus — le doute se relève de lui-même.

    RELEVES      (code_ouv, date, chantier, quantite, heures)
                 Une observation de chantier, pas un paramètre. Le
                 rendement constaté en est le quotient — jamais stocké,
                 pour qu'il n'existe qu'une seule vérité.

    RESSOURCES   (code_res, libelle_res, type_res, unite_res, pu_res,
                  date_prix, source)
                 type_res : MO | MAT | EQP
                 `date_prix` et `source` sont FACULTATIFS : d'où vient ce
                 prix, et de quand il date. Un prix d'achat ne se valide
                 pas comme un rendement — il périme. Absents, l'outil dit
                 qu'il ne sait pas ; il ne dit jamais que le prix est
                 frais.
         |
    COMPOSITION  (code_ouv, code_res, qte_res)
                 qte_res sur les lignes MO = RENDEMENT en h/unité.
                 Seule donnée non achetable : elle vient de l'expérience.
         |
    OUVRAGES     (code_ouv, lot, libelle_ouv, unite_ouv, code_ref)
         v
    BORDEREAU    calculé (moteur.calcul_bordereau)

CODIFICATION — ne jamais fusionner les deux espaces de nommage :
  · ouvrages de la bibliothèque : `LL.NN`  (lot.numéro, ex. 40.20)
  · postes d'un métré imposé    : ceux du pouvoir adjudicateur
Le lien se fait exclusivement par MAPPING (ou la colonne `code_ref`).
"""

import json
import os
from datetime import date
from pathlib import Path

from .parametres import ENTREPRISE, PARAMS  # noqa: F401  (façade historique)

# CHIFFRAGE_DATA permet de charger d'autres tables que celles du dépôt :
# une bibliothèque par entité, un jeu d'essai, ou simplement de vérifier
# qu'une table corrompue est bien refusée. À défaut, `chiffrage/data`.
DOSSIER_DATA = Path(
    os.environ.get("CHIFFRAGE_DATA")
    or Path(__file__).resolve().parent / "data"
)


class BibliothequeInvalide(RuntimeError):
    """Les tables sont inutilisables. On s'arrête au lieu de chiffrer.

    À la différence du lexique ou des paramètres, il n'existe PAS de
    valeurs de repli : une bibliothèque vide ne dégraderait pas le
    résultat, elle rendrait « aucun ouvrage » pour tous les postes —
    c'est-à-dire une offre entièrement vide, présentée comme normale.
    Mieux vaut refuser de démarrer.
    """


def _charger(nom, dossier):
    chemin = Path(dossier) / f"{nom}.json"
    try:
        return json.loads(chemin.read_text(encoding="utf-8"))
    except FileNotFoundError as err:
        raise BibliothequeInvalide(
            f"Table manquante : {chemin}. La bibliothèque ne peut pas "
            f"être chargée."
        ) from err
    except ValueError as err:
        raise BibliothequeInvalide(
            f"Table illisible : {chemin} — {err}"
        ) from err


def _charger_optionnel(nom, dossier, defaut):
    """Comme _charger, mais un fichier ABSENT rend la valeur par défaut.

    Réservé aux tables qui ne portent AUCUN prix. Une bibliothèque plus
    ancienne que l'app — le cas ordinaire sur Streamlit Cloud, où un
    push ne redémarre pas le processus — n'a pas encore le fichier :
    lever refuserait de démarrer pour une table dont rien ne dépend.
    Un fichier PRÉSENT mais illisible reste une faute : c'est une
    corruption, pas une absence.

    Ce raisonnement ne s'étend pas à une table de prix. Une table de
    prix absente ne dégraderait pas le résultat, elle rendrait des
    offres à zéro présentées comme normales.
    """
    if not (Path(dossier) / f"{nom}.json").exists():
        return defaut
    return _charger(nom, dossier)


def _est_une_date(valeur):
    """Une date de calendrier écrite AAAA-MM-JJ, et rien d'autre."""
    try:
        date.fromisoformat(str(valeur))
    except (TypeError, ValueError):
        return False
    return True


def charger_tables(dossier=None):
    """
    Lit et CONTRÔLE les tables d'un dossier `data/`.

    Fonction pure d'un répertoire : c'est ce qui la rend vérifiable
    sans recharger de module ni toucher à l'état global — et ce qui
    permet, accessoirement, de faire tourner l'outil sur une autre
    bibliothèque que celle du dépôt.

    Retourne un dict des tables. Lève BibliothequeInvalide au moindre
    défaut : voir _controler() pour ce qui est vérifié et pourquoi.
    """
    dossier = Path(dossier or DOSSIER_DATA)
    tables = {nom: _charger(nom, dossier) for nom in (
        "lots", "ressources", "ouvrages", "composition", "mapping",
        "ouvrages_a_valider", "metres_histo")}
    tables["releves"] = _charger_optionnel("releves", dossier, [])
    tables["validations"] = _charger_optionnel("validations", dossier, [])

    # Le lot se déduit du code : le stocker deux fois, c'est risquer
    # qu'ils divergent. Les devis historiques reviennent du JSON en
    # listes de listes — on les rétablit en couples.
    tables["ouvrages"] = [dict(o, lot=o["code_ouv"].split(".")[0])
                           for o in tables["ouvrages"]]
    for devis in tables["metres_histo"].values():
        devis["lignes"] = [tuple(ligne) for ligne in devis["lignes"]]

    tables["ressources_par_code"] = {r["code_res"]: r
                                      for r in tables["ressources"]}
    tables["ouvrages_par_code"] = {o["code_ouv"]: o
                                    for o in tables["ouvrages"]}
    _controler(tables)
    return tables


def _controler(t):
    """
    Vérifie les tables AU CHARGEMENT, pas au premier chiffrage.

    Éditable à la main veut dire corrompable à la main. Chacune de ces
    fautes produirait un prix faux sans se voir dans un JSON de
    150 lignes : une ressource orpheline vaut zéro, un ouvrage sans
    composition se vend gratuitement, un type inconnu échappe aux
    trois déboursés.
    """
    fautes = []
    ressources, ouvrages = t["ressources"], t["ouvrages"]
    par_res, par_ouv = t["ressources_par_code"], t["ouvrages_par_code"]

    types_admis = {"MO", "MAT", "EQP"}
    for res in ressources:
        for champ in ("code_res", "libelle_res", "type_res", "unite_res",
                       "pu_res"):
            if champ not in res:
                fautes.append(f"ressource {res.get('code_res', '?')} : "
                               f"champ « {champ} » manquant")
        if res.get("type_res") not in types_admis:
            fautes.append(f"ressource {res.get('code_res')} : type "
                           f"« {res.get('type_res')} » inconnu "
                           f"(attendu {'/'.join(sorted(types_admis))})")
        if not isinstance(res.get("pu_res"), (int, float)) or res["pu_res"] < 0:
            fautes.append(f"ressource {res.get('code_res')} : prix "
                           f"« {res.get('pu_res')} » invalide")
        # La provenance d'un prix est FACULTATIVE : aucune ressource n'en
        # porte aujourd'hui, et l'exiger refuserait de démarrer sur une
        # bibliothèque parfaitement utilisable. Mais une date PRÉSENTE et
        # illisible est une corruption, pas une absence — et une date
        # illisible qu'on laisserait passer se lirait « prix jamais daté »,
        # c'est-à-dire l'inverse de ce qu'elle voulait dire.
        if "date_prix" in res and not _est_une_date(res["date_prix"]):
            fautes.append(f"ressource {res.get('code_res')} : date de prix "
                           f"« {res['date_prix']} » illisible "
                           f"(attendu AAAA-MM-JJ)")
        if "source" in res and not str(res["source"] or "").strip():
            fautes.append(f"ressource {res.get('code_res')} : source vide — "
                           f"une provenance vide ne dit rien, autant "
                           f"l'omettre")

    if len(par_res) != len(ressources):
        fautes.append("codes de ressource en double")
    if len(par_ouv) != len(ouvrages):
        fautes.append("codes d'ouvrage en double")

    for ouv in ouvrages:
        if not ouv.get("unite_ouv"):
            fautes.append(f"ouvrage {ouv.get('code_ouv')} : unité manquante — "
                           f"le contrôle d'unité deviendrait inopérant")
        if ouv["lot"] not in t["lots"]:
            fautes.append(f"ouvrage {ouv['code_ouv']} : lot « {ouv['lot']} » "
                           f"absent de lots.json")

    avec_composition = set()
    for comp in t["composition"]:
        if comp["code_res"] not in par_res:
            fautes.append(f"composition {comp['code_ouv']} : ressource "
                           f"« {comp['code_res']} » inexistante")
        if comp["code_ouv"] not in par_ouv:
            fautes.append(f"composition : ouvrage « {comp['code_ouv']} » "
                           f"inexistant")
        if not isinstance(comp.get("qte_res"), (int, float)) \
                or comp["qte_res"] <= 0:
            fautes.append(f"composition {comp['code_ouv']}/{comp['code_res']} :"
                           f" quantité « {comp.get('qte_res')} » invalide")
        avec_composition.add(comp["code_ouv"])

    for code in sorted(set(par_ouv) - avec_composition):
        fautes.append(f"ouvrage {code} : aucune composition — il se "
                       f"vendrait à 0 €")

    for poste, code in t["mapping"].items():
        if code not in par_ouv:
            fautes.append(f"mapping {poste} -> ouvrage « {code} » inexistant")

    for code in t["ouvrages_a_valider"]:
        if code not in par_ouv:
            fautes.append(f"ouvrages_a_valider : « {code} » inexistant")

    # Les relevés sont des OBSERVATIONS : leur forme se contrôle, mais un
    # relevé qui pointe un ouvrage supprimé depuis n'est pas une faute —
    # c'est une preuve périmée, et la perdre serait pire que la garder.
    # controle_coherence() la signale ; le chargeur, lui, laisse passer.
    # Même traitement que les relevés : la forme se contrôle, la
    # référence à un ouvrage disparu ne bloque rien.
    if not isinstance(t["validations"], list):
        fautes.append("validations : la table doit être une liste")
    else:
        for i, val in enumerate(t["validations"]):
            if not isinstance(val, dict):
                fautes.append(f"validation n°{i + 1} : ce n'est pas un objet")
                continue
            for champ in ("code_ouv", "date"):
                if not str(val.get(champ) or "").strip():
                    fautes.append(f"validation n°{i + 1} : « {champ} » "
                                   f"manquant")
            rendement = val.get("rendement")
            if not isinstance(rendement, (int, float)) or rendement <= 0:
                fautes.append(f"validation n°{i + 1} : rendement "
                               f"« {rendement} » invalide — une validation "
                               f"sans valeur ne prouve rien")

    if not isinstance(t["releves"], list):
        fautes.append("releves : la table doit être une liste")
    else:
        for i, rel in enumerate(t["releves"]):
            if not isinstance(rel, dict):
                fautes.append(f"relevé n°{i + 1} : ce n'est pas un objet")
                continue
            for champ in ("code_ouv", "date", "chantier"):
                if not str(rel.get(champ) or "").strip():
                    fautes.append(f"relevé n°{i + 1} : « {champ} » manquant — "
                                   f"un relevé sans provenance ne prouve rien")
            for champ in ("quantite", "heures"):
                valeur = rel.get(champ)
                if not isinstance(valeur, (int, float)) or valeur <= 0:
                    fautes.append(f"relevé n°{i + 1} : {champ} "
                                   f"« {valeur} » invalide")

    if fautes:
        raise BibliothequeInvalide(
            f"{len(fautes)} incohérence(s) dans {DOSSIER_DATA} :\n  - "
            + "\n  - ".join(fautes[:20])
            + ("\n  …" if len(fautes) > 20 else ""))


_TABLES = charger_tables()

LOTS = _TABLES["lots"]
RESSOURCES = _TABLES["ressources"]
OUVRAGES = _TABLES["ouvrages"]
COMPOSITION = _TABLES["composition"]
MAPPING = _TABLES["mapping"]
OUVRAGES_A_VALIDER = _TABLES["ouvrages_a_valider"]
RELEVES = _TABLES["releves"]
VALIDATIONS = _TABLES["validations"]
METRES_HISTO = _TABLES["metres_histo"]
RESSOURCES_PAR_CODE = _TABLES["ressources_par_code"]
OUVRAGES_PAR_CODE = _TABLES["ouvrages_par_code"]
