"""
╔══════════════════════════════════════════════════════════════╗
║  APPARIEMENT POSTE IMPOSÉ -> OUVRAGE DE LA BIBLIOTHÈQUE                ║
╚══════════════════════════════════════════════════════════════╝

`MAPPING` est à refaire à CHAQUE marché : les codes de poste appartiennent
au pouvoir adjudicateur, pas à nous. C'est le geste le plus fréquent de
tout le processus, et le seul qui obligeait encore à éditer du Python.

Ce module propose des candidats à partir du libellé. Il ne décide rien :
il classe, l'humain tranche. Deux règles qui expliquent tout le reste.

**Le vocabulaire du CSC est traduit** vers celui de la bibliothèque
avant toute comparaison — « carrelage mural » devient « faïence »,
« crépi » devient « enduit ». Voir lexique.py.

**L'unité est éliminatoire, pas départageante.** Un poste imposé au mètre
courant ne peut pas être chiffré par un ouvrage au m², quelle que soit la
ressemblance des libellés — le prix serait faux d'un facteur inconnu, et
ça ne se verrait qu'à la facturation. Les unités incompatibles ne sont
donc pas mal classées : elles sont absentes.

**Le score n'est pas une probabilité.** C'est une ressemblance de
libellés, sur un vocabulaire de bâtiment où « dépose » et « pose » ne
diffèrent que d'une lettre pour un algorithme, et sont opposés pour un
chef de chantier. Un score élevé veut dire « regarde ici d'abord », jamais
« c'est bon ».

Python pur : aucune dépendance, testable en CI.
"""

import re
from difflib import SequenceMatcher

from .lexique import (
    PENALITE_OPERATION_OPPOSEE,
    appliquer_expressions,
    canoniser,
    est_demolition,
    sans_accents,
)

# Mots trop fréquents dans un métré pour distinguer quoi que ce soit.
# Les garder ferait remonter n'importe quel poste contenant « compris ».
_VIDES = frozenset(
    """
    de des du la le les et en au aux sur sous pour par avec sans dans
    a à d l un une compris comprise comprises y compris toute toutes tout
    tous ouvrage ouvrages travaux fourniture fournitures pose mise oeuvre
    ml m m2 m3 pce ff u qf qp
    existant existante existants existantes ancien ancienne neuf neuve
    prealable prealablement necessaire necessaires divers diverses
    seul seule complet complete type suivant selon place
    """.split()
)

# Racines métier qui pèsent double : elles portent le sens du poste, et
# leur présence des deux côtés est un signal bien plus fort qu'un mot
# ordinaire. Liste courte et volontairement conservatrice.
_PIVOTS = frozenset(
    """
    depose demolition piquage evacuation etancheite isolation plafonnage
    enduit peinture carrelage faience chape maconnerie rejointoiement
    linteau seuil chassis porte garde-corps sanitaire electricite terre
    prise wc echafaudage nettoyage cimentage solin membrane epdm
    bitumineuse pir laine platre ba13 siloxane
    """.split()
)


def normaliser(texte, garder_operation=False):
    """
    Libellé -> liste de mots significatifs, traduits vers le vocabulaire
    de la bibliothèque (voir lexique.py).

    Les marqueurs de démolition sont RETIRÉS par défaut : ils sont si
    fréquents dans un métré qu'ils faisaient se ressembler deux
    démolitions sans rapport. L'opération est traitée à part, comme une
    facette. `garder_operation=True` les conserve, pour la détecter.
    """
    texte = sans_accents(str(texte or "")).lower()
    texte = appliquer_expressions(texte)
    mots = [
        canoniser(m)
        for m in re.split(r"[^a-z0-9-]+", texte)
        if m and m not in _VIDES and len(m) > 2
    ]
    mots = [m for m in mots if m and len(m) > 2 and m not in _VIDES]
    if garder_operation:
        return mots
    from .lexique import DEMOLITION

    return [m for m in mots if m not in DEMOLITION]


def score(libelle_poste, libelle_ouvrage):
    """
    Ressemblance entre deux libellés, de 0 à 1.

    Trois composantes, parce qu'aucune ne suffit seule :
      · recouvrement des mots  — insensible à l'ordre et aux tournures ;
      · ressemblance de chaîne — rattrape les variantes d'écriture
        (« châssis » / « chassis PVC ») que le recouvrement rate ;
      · bonus de pivot        — un mot métier partagé (« étanchéité »,
        « dépose ») compte plus qu'un mot ordinaire.
    """
    mots_poste = set(normaliser(libelle_poste))
    mots_ouvrage = set(normaliser(libelle_ouvrage))
    if not mots_poste or not mots_ouvrage:
        return 0.0

    # L'opération (dépose vs mise en œuvre) est une facette, pas un mot :
    # les deux libellés doivent décrire le MÊME sens de travail.
    demolition_poste = est_demolition(
        normaliser(libelle_poste, garder_operation=True))
    demolition_ouvrage = est_demolition(
        normaliser(libelle_ouvrage, garder_operation=True))
    penalite = (
        PENALITE_OPERATION_OPPOSEE
        if demolition_poste != demolition_ouvrage
        else 1.0
    )

    communs = mots_poste & mots_ouvrage
    recouvrement = len(communs) / len(mots_poste | mots_ouvrage)

    chaine = SequenceMatcher(
        None, " ".join(sorted(mots_poste)), " ".join(sorted(mots_ouvrage))
    ).ratio()

    pivots = len(communs & _PIVOTS)
    bonus = min(0.20, 0.10 * pivots)

    brut = min(1.0, 0.5 * recouvrement + 0.5 * chaine + bonus)
    return round(brut * penalite, 4)


def suggerer(poste, bordereau, limite=3, seuil=0.0):
    """
    Classe les ouvrages candidats pour un poste de métré.

    poste     : dict {designation, unite, ...} tel que lu par lire_metre()
    bordereau : sortie de moteur.calcul_bordereau()

    Retourne [(code_ouv, score), ...], au plus `limite`, du meilleur au
    moins bon. **Seuls les ouvrages dont l'unité est compatible sont
    candidats** : chiffrer dans une autre unité est un faux silencieux.

    Liste vide = aucun ouvrage dans cette unité. C'est une information,
    pas un échec : il manque l'ouvrage, il faut le créer.
    """
    from .metre_io import normaliser_unite

    unite_poste = normaliser_unite(poste.get("unite"))
    candidats = [
        (code, score(poste.get("designation", ""), ligne["libelle_ouv"]))
        for code, ligne in bordereau.items()
        if normaliser_unite(ligne["unite_ouv"]) == unite_poste
    ]
    candidats = [(c, s) for c, s in candidats if s >= seuil]
    candidats.sort(key=lambda cs: (-cs[1], cs[0]))
    return candidats[:limite]


def proposer_mapping(postes, bordereau, mapping_connu=None, seuil=0.35):
    """
    Pré-remplit une table de correspondance pour un métré entier.

    Trois origines possibles, dans cet ordre de confiance :
      'connu'    — le code figure déjà dans le MAPPING fourni. Repris tel
                   quel, sans discussion : c'est un choix humain antérieur.
      'suggere'  — meilleur candidat au-dessus du seuil. À RELIRE.
      'aucun'    — rien de convaincant, ou aucun ouvrage dans cette unité.

    Retourne {code_poste: {'code_ouv', 'score', 'origine', 'candidats'}}.
    `candidats` sert à peupler une liste déroulante dans une interface.

    Le seuil ne protège de rien tout seul : il évite d'afficher du bruit,
    il ne dit pas qu'une suggestion au-dessus est juste.
    """
    mapping_connu = mapping_connu or {}
    proposition = {}
    for poste in postes:
        code = poste["code"]
        candidats = suggerer(poste, bordereau, limite=5)
        connu = mapping_connu.get(code)
        if connu and connu in bordereau:
            proposition[code] = {
                "code_ouv": connu,
                "score": 1.0,
                "origine": "connu",
                "candidats": candidats,
            }
            continue
        if candidats and candidats[0][1] >= seuil:
            proposition[code] = {
                "code_ouv": candidats[0][0],
                "score": candidats[0][1],
                "origine": "suggere",
                "candidats": candidats,
            }
        else:
            proposition[code] = {
                "code_ouv": None,
                "score": candidats[0][1] if candidats else 0.0,
                "origine": "aucun",
                "candidats": candidats,
            }
    return proposition
