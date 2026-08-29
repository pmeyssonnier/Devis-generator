"""
╔══════════════════════════════════════════════════════╗
║  LEXIQUE MÉTIER — synonymes de CSC et facette d'opération              ║
╚══════════════════════════════════════════════════════╝

Un pouvoir adjudicateur n'écrit pas « faïence » : il écrit « carrelage
mural ». Ni « PIR » : « mousse polyuréthane ». Ce module traduit son
vocabulaire vers celui de la bibliothèque.

Deux mécanismes, et le second est le plus important.

**1. Synonymes.** Table de traduction, enrichie au fil des marchés reçus.
Déterministe et relisible : là où une similarité cosinus est un nombre
qu'on subit, une entrée de lexique est une décision qu'on relit.

**2. L'opération est une FACETTE, pas un mot.**
« Dépose de carrelage » et « Pose de carrelage » ne diffèrent que par un
mot, sur un vocabulaire où tout le reste est identique — et ce sont deux
travaux opposés, à des prix qui vont du simple au triple. Pire : les
verbes de démolition (« dépose », « démontage », « enlèvement ») sont si
fréquents dans un métré qu'ils faisaient se ressembler DEUX DÉMOLITIONS
SANS RAPPORT : « dépose du carrelage mural » proposait « dépose de
plafond » — deux fois le mot « dépose », et rien d'autre en commun.

On extrait donc l'opération du texte pour en faire une dimension à part,
comme l'unité :
  · les marqueurs de démolition ne comptent plus dans la ressemblance
    des libellés (ils ne distinguent rien) ;
  · une démolition comparée à une mise en œuvre est fortement pénalisée.

**Pénalité et non élimination**, à la différence de l'unité. L'unité est
DÉCLARÉE dans le métré : s'y fier est sûr. L'opération est INFÉRÉE de
mots : une inférence fausse ne doit pas faire disparaître un candidat
valable, seulement le reléguer.
"""

# ── Expressions à traduire AVANT la découpe en mots ──────────────
# (sans accents : la normalisation les a déjà retirés)
EXPRESSIONS = {
    "carrelage mural": "faience",
    "carreaux muraux": "faience",
    "revetement mural ceramique": "faience",
    "revetement de sol": "carrelage",
    "carreaux ceramiques": "carrelage",
    "carreaux de gres": "carrelage",
    "gres cerame": "carrelage",
    "menuiserie exterieure": "chassis",
    "menuiseries exterieures": "chassis",
    "mousse polyurethane": "pir",
    "laine de roche": "laine minerale",
    "laine de verre": "laine minerale",
    "frein vapeur": "pare-vapeur",
    "pare vapeur": "pare-vapeur",
    "frein-vapeur": "pare-vapeur",
    "point de puisance": "prise courant",
    "prise de courant": "prise courant",
    "eaux usees": "evacuation",
    "parois verticales": "murs",
    "parois interieures": "murs",
    "mise en peinture": "peinture",
    "plaques de platre": "ba13",
    "plaque de platre": "ba13",
    "plafond suspendu": "faux plafond",
    "haute pression": "haute-pression",
    "sous pression": "haute-pression",
    "organisme agree": "organisme-agree",
    "bati-support": "bati-support",
    "bati support": "bati-support",
}

# ── Synonymes mot à mot ─────────────────────────────────
SYNONYMES = {
    # façades
    "crepi": "enduit",
    # En Belgique, le cimentage INTÉRIEUR est du plafonnage. La
    # canonisation s'applique aux deux côtés de la comparaison, donc
    # « Cimentage hydrofuge de soubassement » (30.40) reste apparié à un
    # poste de cimentage extérieur : c'est « hydrofuge » et
    # « soubassement » qui les départagent, pas le verbe.
    "cimentage": "plafonnage",
    "cimenter": "plafonnage",
    "parement": "facade",
    "enduire": "enduit",
    # revêtements
    "faiences": "faience",
    "ceramique": "carrelage",
    "chapes": "chape",
    "ravoirage": "chape",
    # plafonnage
    "plafonnage": "plafonnage",
    "ratissage": "lissage",
    "rattrapage": "lissage",
    "lisse": "lissage",
    "lissee": "lissage",
    "lisses": "lissage",
    "platrerie": "plafonnage",
    # menuiseries
    "fenetre": "chassis",
    "fenetres": "chassis",
    "chassis": "chassis",
    "vitrage": "chassis",
    # sanitaire / électricité
    "cuvette": "wc",
    "toilette": "wc",
    "tuyauterie": "tuyau",
    "tube": "tuyau",
    "descente": "evacuation",
    "puisance": "prise",
    "equipotentielles": "equipotentielles",
    # étanchéité
    "natte": "membrane",
    "soudee": "bitumineuse",
    "synthetique": "epdm",
    # divers
    "gravats": "dechets",
    "decombres": "dechets",
    "bachage": "protection",
    "baches": "protection",
    "baguettes": "profiles",
    "cornieres": "profiles",
    "armature": "treillis",
    "armatures": "aciers",
    "preconise": "",
}

# ── Marqueurs d'opération ───────────────────────────────
# Retirés de la comparaison lexicale ET utilisés comme facette.
DEMOLITION = frozenset("""
    depose deposes deposer demolition demolitions demolir demontage
    demonter enlevement enlever arrachage arracher piquage piquer
    decapage decaper curage curer degarnissage suppression supprimer
    retrait evacuation
""".split())

# Une mise en œuvre n'est pas marquée : en français de CSC, l'absence de
# verbe VAUT mise en œuvre (« Enduit de façade minéral armé »). On ne
# détecte donc QUE la démolition, et on traite l'absence comme son
# contraire — d'où la pénalité plutôt que l'élimination.
PENALITE_OPERATION_OPPOSEE = 0.35


# ══════════════════════════════════════════════
#  Surcouche de session
# ══════════════════════════════════════════════
#
# L'interface doit pouvoir ajouter un synonyme et en voir l'effet
# IMMÉDIATEMENT, sans redéploiement : on colle un libellé de CSC qui
# n'a pas été apparié, on ajoute le mot manquant, on regarde le poste
# remonter au premier rang. C'est le seul cycle d'essai qui rende un
# lexique réglable par quelqu'un qui n'écrit pas de Python.
#
# ⚠️ CE N'EST PAS UNE PERSISTANCE. Ces ajouts vivent dans le processus
# et disparaissent au redémarrage — et Streamlit Cloud redémarre
# l'app tout seul. Pour qu'un synonyme survive, il doit finir dans les
# tables ci-dessus, c'est-à-dire dans un commit. L'interface produit
# le bloc à coller ; elle ne peut pas commiter à ta place.
#
# Les tables du module restent donc la source de vérité, et la
# surcouche un brouillon.

SURCOUCHE = {"expressions": {}, "synonymes": {}}


def ajouter_expression(expression, canonique):
    """Ajoute une expression multi-mots à la surcouche de session."""
    SURCOUCHE["expressions"][expression.strip().lower()] = canonique.strip().lower()


def ajouter_synonyme(variante, canonique):
    """Ajoute un synonyme mot à mot à la surcouche de session."""
    SURCOUCHE["synonymes"][variante.strip().lower()] = canonique.strip().lower()


def vider_surcouche():
    """Revient au lexique du dépôt seul."""
    SURCOUCHE["expressions"].clear()
    SURCOUCHE["synonymes"].clear()


def surcouche_en_python():
    """Rend les ajouts de session sous forme de bloc à coller dans ce
    fichier — le seul moyen de les rendre permanents."""
    lignes = []
    if SURCOUCHE["expressions"]:
        lignes.append("# À ajouter dans EXPRESSIONS :")
        lignes += [f'    "{k}": "{v}",'
                    for k, v in sorted(SURCOUCHE["expressions"].items())]
    if SURCOUCHE["synonymes"]:
        if lignes:
            lignes.append("")
        lignes.append("# À ajouter dans SYNONYMES :")
        lignes += [f'    "{k}": "{v}",'
                    for k, v in sorted(SURCOUCHE["synonymes"].items())]
    return "\n".join(lignes)


def appliquer_expressions(texte):
    """Traduit les expressions multi-mots. Texte déjà sans accents."""
    for table in (EXPRESSIONS, SURCOUCHE["expressions"]):
        for expression, canonique in table.items():
            if expression in texte:
                texte = texte.replace(expression, canonique)
    return texte


def canoniser(mot):
    """Ramène un mot à sa forme de bibliothèque ('crepi' -> 'enduit').

    La surcouche de session l'emporte sur la table du dépôt : c'est ce
    qui permet d'essayer une correction avant de la commiter.
    """
    return SURCOUCHE["synonymes"].get(mot, SYNONYMES.get(mot, mot))


def est_demolition(mots):
    """Le libellé décrit-il une dépose plutôt qu'une mise en œuvre ?"""
    return any(mot in DEMOLITION for mot in mots)
