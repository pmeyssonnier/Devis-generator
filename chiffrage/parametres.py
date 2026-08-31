"""
╔══════════════════════════════════════════════╗
║  PARAMÈTRES DE L'ENTREPRISE ET DU CHIFFRAGE                          ║
╚══════════════════════════════════════════════╝

L'identité (raison sociale, adresse, TVA) et les coefficients de vente
(frais généraux, frais de chantier, aléas, marge) étaient écrits dans
le code. Changer une adresse ou un point de marge demandait d'éditer
du Python — pour des valeurs qui n'ont rien de technique et que le
chef d'entreprise est seul à connaître.

Même dispositif que le lexique appris, et pour les mêmes raisons :

    parametres_local.json     ← écrit par l'interface, commité
    parametres.py             ← les valeurs de repli, jamais réécrites

**Du JSON, pas du Python.** Le contenu vient de champs de saisie :
écrire du code exécutable à partir d'une saisie serait une injection.

Un fichier absent, illisible ou incomplet ne doit JAMAIS empêcher de
chiffrer : on retombe sur les valeurs ci-dessous, champ par champ.
"""

import json
import os
import re
from pathlib import Path

def _chemin_local(nom):
    """Où lire le fichier local `nom` : dans le dossier de l'entreprise
    si CHIFFRAGE_DATA le désigne, sinon à sa place historique.

    CHIFFRAGE_DATA désigne TOUT ce qui appartient à une entreprise — ses
    tables, son identité, son lexique. C'est ce qui permet de faire
    tourner une instance par entrepreneur sur le même code, chacune ne
    lisant et n'écrivant que son dossier.

    Non défini, rien ne bouge : le déploiement en service garde ses
    fichiers là où ils sont nés.
    """
    dossier = os.environ.get("CHIFFRAGE_DATA")
    if dossier:
        return Path(dossier) / nom
    return Path(__file__).resolve().parent / nom

CHEMIN_LOCAL = _chemin_local("parametres_local.json")

# ── Identité — en-tête des devis et des courriers ────────────────
ENTREPRISE_DEFAUT = {
    "nom": "BAG BATTER SRL",
    "adresse": "Ronkel 18",
    "cp_ville": "1780 Wemmel",
    "pays": "Belgique",
    "tva": "BE 0766.637.025",
    "activite": "Entreprise générale de rénovation — façades, étanchéité, "
                "plafonnage, isolation, peinture, menuiserie extérieure, "
                "sanitaire léger",
}

# ── Coefficients de vente ────────────────────────
#
#   pu_vente = debourse_sec x K
#   K        = (1+FG) x (1+FC) x (1+aleas) x (1+marge)
#
# 0,12 / 0,05 / 0,03 / 0,10  ->  K = 1,3324
PARAMS_DEFAUT = {
    "fg": 0.12,
    "fc": 0.05,
    "aleas": 0.03,
    "marge": 0.10,
    # 6 % : logement privé de plus de 10 ans, facturé au consommateur final.
    # EN MARCHÉ PUBLIC C'EST 21 %.
    "tva": 0.06,
    "tva_marche_public": 0.21,
}

# Un taux hors de cet intervalle relève de la faute de frappe, pas du
# choix commercial : 250 % de marge est une virgule mal placée.
TAUX_MAX = 2.0
LONGUEUR_MAX = 200


def valider_entreprise(valeurs):
    """Rend l'identité nettoyée. Lève ValueError si elle est inutilisable."""
    propre = {}
    for champ, defaut in ENTREPRISE_DEFAUT.items():
        valeur = str(valeurs.get(champ, defaut) or "").strip()
        if len(valeur) > LONGUEUR_MAX:
            raise ValueError(
                f"« {champ} » dépasse {LONGUEUR_MAX} caractères.")
        propre[champ] = valeur
    if not propre["nom"]:
        raise ValueError(
            "La raison sociale ne peut pas être vide : elle figure en "
            "en-tête de chaque devis et de chaque courrier.")
    # Contrôle volontairement souple : on refuse l'absurde, pas les
    # variantes d'écriture (BE0766637025, BE 0766.637.025…).
    if propre["tva"] and not re.search(r"\d{4}", propre["tva"]):
        raise ValueError(
            f"Numéro de TVA douteux : « {propre['tva'] } ». Il figure sur "
            f"les devis et engage l'entreprise.")
    return propre


def valider_params(valeurs):
    """Rend les coefficients nettoyés. Lève ValueError si aberrants."""
    propre = {}
    for champ, defaut in PARAMS_DEFAUT.items():
        brut = valeurs.get(champ, defaut)
        try:
            taux = float(brut)
        except (TypeError, ValueError) as err:
            raise ValueError(f"« {champ} » n'est pas un nombre : {brut!r}"
                              ) from err
        if not 0.0 <= taux <= TAUX_MAX:
            raise ValueError(
                f"« {champ} » vaut {taux:g}, hors de [0 ; {TAUX_MAX:g}]. "
                f"Les taux s'expriment en fraction : 0,10 pour 10 %.")
        propre[champ] = taux
    return propre


def charger_local(chemin=None):
    """
    Lit parametres_local.json et le fusionne avec les valeurs de repli.

    Un fichier absent est le cas NORMAL. Un fichier illisible ou
    partiellement aberrant ne doit pas empêcher de chiffrer : chaque
    bloc retombe indépendamment sur son défaut.
    """
    chemin = Path(chemin or CHEMIN_LOCAL)
    entreprise = dict(ENTREPRISE_DEFAUT)
    params = dict(PARAMS_DEFAUT)
    if not chemin.exists():
        return entreprise, params
    try:
        donnees = json.loads(chemin.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return entreprise, params

    try:
        entreprise = valider_entreprise(
            {**ENTREPRISE_DEFAUT, **(donnees.get("entreprise") or {})})
    except ValueError:
        pass
    try:
        params = valider_params(
            {**PARAMS_DEFAUT, **(donnees.get("params") or {})})
    except ValueError:
        pass
    return entreprise, params


def serialiser(entreprise, params):
    """Le contenu exact de parametres_local.json, validé avant écriture."""
    return json.dumps(
        {"entreprise": valider_entreprise(entreprise),
         "params": valider_params(params)},
        ensure_ascii=False, indent=2, sort_keys=True) + "\n"


# Valeurs effectives, assemblées une fois à l'import.
ENTREPRISE, PARAMS = charger_local()
