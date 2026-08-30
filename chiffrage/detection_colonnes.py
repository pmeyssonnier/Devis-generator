"""
╔══════════════════════════════════════════════╗
║  QUELLE COLONNE EST QUOI                                             ║
╚══════════════════════════════════════════════╝

Les colonnes étaient codées en dur — B le code, F la quantité, G le
prix unitaire — c'est-à-dire la disposition du métré d'entraînement.
Un autre pouvoir adjudicateur range ses colonnes autrement, et l'outil
ne lisait alors rien du tout, ou pire : il écrivait le prix dans la
mauvaise colonne.

Ce module lit les en-têtes du classeur reçu et propose une
correspondance. **Il propose, l'humain valide** — l'interface affiche
la détection et permet de la corriger avant tout chiffrage.

Deux voies, complémentaires :

  · PAR LES INTITULÉS — « Quantité », « Qté », « Métré » désignent la
    même colonne. Rapide et lisible, mais suppose un en-tête.

  · PAR LE CONTENU — la colonne des codes est celle qui contient le
    plus de cellules ressemblant à un code de poste ; la quantité est
    la première colonne numérique à sa droite. Plus lent à écrire, mais
    c'est la seule voie quand les en-têtes sont absents, fusionnés ou
    rédigés dans une langue qu'on n'a pas prévue.

La détection donne pour chaque champ son ORIGINE (`entete`, `contenu`,
ou rien). Une correspondance déduite du contenu mérite un coup d'œil
de plus qu'une correspondance lue dans un titre, et l'interface le
signale plutôt que d'afficher une confiance uniforme trompeuse.
"""

import re
import unicodedata

# Les quatre premiers sont indispensables : sans eux, on ne sait ni
# quoi chiffrer, ni où écrire, ni comment vérifier l'unité — et le
# contrôle d'unité n'est pas négociable.
CHAMPS_REQUIS = ("code", "unite", "quantite", "pu")
CHAMPS = CHAMPS_REQUIS + ("designation", "nature", "montant")

# Ordre important : les intitulés les plus spécifiques d'abord, sinon
# « prix » attraperait « prix total » aussi bien que « prix unitaire ».
SYNONYMES = {
    "quantite": ["quantite presumee", "quantites presumees", "quantite",
                  "quantites", "qte", "qty", "metre", "metres", "nombre"],
    "pu": ["prix unitaire htva", "prix unitaire hors tva", "prix unitaire",
            "pu htva", "p u", "pu", "prix par unite", "prix"],
    "montant": ["montant htva", "montant total", "montant", "prix total",
                 "total htva", "total"],
    "designation": ["designation des ouvrages", "designation", "description",
                     "libelle", "intitule", "nature des travaux", "objet",
                     "denomination"],
    "unite": ["unite de mesure", "unite", "un", "u", "mesure"],
    "nature": ["nature", "type de poste", "qf qp", "mode"],
    "code": ["code poste", "numero de poste", "no de poste", "poste", "code",
              "article", "reference", "ref", "numero", "no", "n", "index",
              "rubrique"],
}


def _sans_accents(texte):
    decompose = unicodedata.normalize("NFD", str(texte))
    return "".join(c for c in decompose if unicodedata.category(c) != "Mn")


def normaliser_entete(valeur):
    """« Prix unitaire (HTVA) » -> « prix unitaire htva »."""
    texte = _sans_accents(valeur or "").lower()
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", texte)).strip()


def _champ_de(intitule):
    """À quel champ correspond cet intitulé ? None si aucun."""
    normalise = normaliser_entete(intitule)
    if not normalise:
        return None
    # « P.U. HTVA » devient « p u htva » : la ponctuation d'abréviation
    # laisse des lettres isolées qu'aucune liste de synonymes ne peut
    # prévoir. On compare donc aussi la forme sans espaces, où
    # « p u htva » et « pu htva » se rejoignent.
    colle = normalise.replace(" ", "")
    for champ, formes in SYNONYMES.items():
        for forme in formes:
            # Égalité d'abord : « unite » ne doit pas être happé par
            # « quantite » sous prétexte qu'il en est un morceau.
            if normalise == forme or colle == forme.replace(" ", ""):
                return champ
    for champ, formes in SYNONYMES.items():
        for forme in formes:
            if len(forme) >= 4 and forme in normalise:
                return champ
    return None


def detecter_par_entetes(ws, lignes_a_scruter=40):
    """
    Cherche la ligne d'en-tête et ce qu'elle annonce.

    La ligne retenue est celle qui nomme le plus de champs DIFFÉRENTS,
    avec un minimum de trois : un cartouche de marché contient souvent
    un mot isolé comme « Objet » qui ne fait pas un en-tête de tableau.
    """
    meilleure, colonnes_retenues = 0, {}
    ligne_entete = None

    for row in ws.iter_rows(min_row=1, max_row=min(lignes_a_scruter,
                                                     ws.max_row)):
        trouvees = {}
        for cellule in row:
            champ = _champ_de(cellule.value)
            # Premier arrivé, premier servi : deux colonnes « Total »
            # ne doivent pas se voler la place.
            if champ and champ not in trouvees:
                trouvees[champ] = cellule.column
        if len(trouvees) > meilleure:
            meilleure, colonnes_retenues = len(trouvees), trouvees
            ligne_entete = row[0].row

    if meilleure < 3:
        return {}, None
    return colonnes_retenues, ligne_entete


def detecter_par_contenu(ws, ligne_debut=1):
    """
    Déduit les colonnes de ce qu'elles contiennent.

    Sert quand il n'y a pas d'en-tête exploitable — et sert aussi de
    contre-épreuve : une colonne « Quantité » qui ne contient aucun
    nombre est plus probablement mal nommée que mal remplie.
    """
    from .metre_io import est_code_poste

    codes_par_colonne, nombres_par_colonne, textes_par_colonne = {}, {}, {}
    lignes_de_poste = []

    for row in ws.iter_rows(min_row=ligne_debut, max_row=ws.max_row):
        for cellule in row:
            if est_code_poste(cellule.value):
                codes_par_colonne[cellule.column] = (
                    codes_par_colonne.get(cellule.column, 0) + 1)
                lignes_de_poste.append(row)
                break

    if not codes_par_colonne:
        return {}
    colonne_code = max(codes_par_colonne, key=codes_par_colonne.get)
    if codes_par_colonne[colonne_code] < 3:
        return {}

    # On ne compte que sur les lignes qui SONT des postes : le reste du
    # classeur (cartouches, totaux) fausserait le décompte.
    for row in lignes_de_poste:
        for cellule in row:
            if cellule.column <= colonne_code:
                continue
            valeur = cellule.value
            if isinstance(valeur, (int, float)):
                nombres_par_colonne[cellule.column] = (
                    nombres_par_colonne.get(cellule.column, 0) + 1)
            elif isinstance(valeur, str) and valeur.strip():
                textes_par_colonne[cellule.column] = (
                    textes_par_colonne.get(cellule.column, 0) + 1)

    trouvees = {"code": colonne_code}

    # La quantité est remplie ; le prix unitaire, lui, est vide — c'est
    # nous qui devons l'écrire. La première colonne numérique à droite
    # du code est donc la quantité.
    #
    # Le seuil se teste À PART : des colonnes numériques peuvent exister
    # sans qu'aucune l'atteigne — trois nombres épars dans un classeur
    # de dix postes. `min()` recevait alors une séquence vide et levait,
    # ce qui faisait planter le dépôt du métré au lieu de rendre une
    # détection incomplète. Aucune colonne assez remplie : la quantité
    # reste indéterminée, et le mécanisme des champs manquants prend le
    # relais — il est fait pour ça.
    assez_remplies = [c for c, n in nombres_par_colonne.items()
                       if n >= max(3, 0.5 * len(lignes_de_poste))]
    if assez_remplies:
        trouvees["quantite"] = min(assez_remplies)
    if textes_par_colonne:
        # La désignation est la colonne de texte la plus fournie.
        trouvees["designation"] = max(textes_par_colonne,
                                       key=textes_par_colonne.get)
    return trouvees


def detecter(ws):
    """
    Rend la correspondance colonne <-> champ pour une feuille.

    {"champs": {"code": 2, ...}, "origines": {"code": "entete", ...},
     "ligne_entete": 9, "manquants": ["pu"]}

    Les intitulés priment sur le contenu : un pouvoir adjudicateur qui
    nomme ses colonnes sait mieux que nous ce qu'elles contiennent.
    """
    par_entetes, ligne_entete = detecter_par_entetes(ws)
    par_contenu = detecter_par_contenu(
        ws, ligne_debut=(ligne_entete or 0) + 1)

    champs, origines = {}, {}
    for champ in CHAMPS:
        if champ in par_entetes:
            champs[champ] = par_entetes[champ]
            origines[champ] = "entete"
        elif champ in par_contenu:
            champs[champ] = par_contenu[champ]
            origines[champ] = "contenu"

    # EXCEPTION POUR LE CODE : le contenu l'emporte sur l'intitulé.
    # C'est le seul champ qu'on sache vérifier — une cellule ressemble
    # à un code de poste ou non. Et l'intitulé y est trompeur : un
    # métré titre couramment « N° » la colonne du simple compteur de
    # lignes, juste avant la vraie colonne des codes. La détection
    # partait alors sur le compteur, et plus aucun poste n'était lu.
    if "code" in par_contenu and par_contenu["code"] != champs.get("code"):
        champs["code"] = par_contenu["code"]
        origines["code"] = "contenu"

    # Le prix unitaire est presque toujours vide dans le métré reçu :
    # aucune détection par contenu n'est possible. À défaut d'intitulé,
    # on propose la colonne juste après la quantité — proposition
    # explicitement signalée comme telle, car écrire un prix dans la
    # mauvaise colonne rendrait une offre silencieusement fausse.
    if "pu" not in champs and "quantite" in champs:
        champs["pu"] = champs["quantite"] + 1
        origines["pu"] = "position"

    # L'unité précède presque toujours la quantité.
    if "unite" not in champs and "quantite" in champs:
        candidate = champs["quantite"] - 1
        if candidate > champs.get("code", 0):
            champs["unite"] = candidate
            origines["unite"] = "position"

    return {
        "champs": champs,
        "origines": origines,
        "ligne_entete": ligne_entete,
        "manquants": [c for c in CHAMPS_REQUIS if c not in champs],
    }
