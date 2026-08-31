"""
╔══════════════════════════════════════════════════════╗
║  CONTRÔLE DES PRIX — avant de déposer                                 ║
╚══════════════════════════════════════════════════════╝

Deux risques opposés, et l'offre doit être relue contre les deux :

    prix trop bas   -> offre écartée pour prix anormalement bas
                       (art. 36 AR 18/04/2017), ou chantier exécuté
                       à perte ;
    prix trop haut  -> marché perdu.

Ce module ne fait QUE de l'arithmétique. C'est délibéré : un contrôle
de prix doit être vérifiable à la main, pas cru sur parole.

──────────────────────────────────────────────────────
CE QU'IL NE PEUT PAS FAIRE
──────────────────────────────────────────────────────
Le pouvoir adjudicateur juge « anormalement bas » en comparant les
offres entre elles. Cette comparaison-là nous est inaccessible au
moment de déposer : nous n'avons pas les prix des concurrents. Ce
module regarde donc l'offre depuis l'INTÉRIEUR de l'entreprise —
est-ce que ce prix couvre ce que ce travail coûte ? — ce qui est une
question différente, et la seule qu'on puisse trancher seul.

Les seuils ci-dessous sont des repères de relecture, pas des règles
de droit. Le seuil légal, la procédure et le délai de réponse
figurent dans l'AR en vigueur et, le plus souvent, dans le CSC
lui-même.
"""

from .bibliotheque import LOTS, PARAMS, RESSOURCES_PAR_CODE
from .moteur import (
    PRIX_ANCIEN_JOURS,
    calcul_bordereau,
    codes_par_statut,
    coefficient_k,
    materiaux_de,
    mois,
)

# ── Repères de relecture ────────────────────────────
# Aucun n'est une règle : ils décident seulement de ce qui est
# porté à l'attention de l'humain.
SEUILS = {
    # Un poste qui pèse plus que ça engage l'offre à lui seul.
    "poids_poste": 0.15,
    # Au-delà, le prix tient surtout au rendement — donc à une
    # donnée d'expérience, la plus fragile de la bibliothèque.
    "part_mo": 0.70,
    # Au-delà, le prix tient surtout à un prix d'achat : le risque
    # est le fournisseur, pas l'atelier.
    "part_mat": 0.70,
    # Écart au prix pratiqué sur un marché antérieur.
    "ecart_historique": 0.20,
    # Part du montant reposant sur des rendements jamais validés.
    "part_non_validee": 0.20,
    # Part reposant sur des rendements VALIDÉS PUIS MODIFIÉS. Le seuil est
    # bas exprès : un seul poste dans ce cas mérite un coup d'œil.
    "part_a_revalider": 0.05,
    # Part des ACHATS de l'offre dont le prix a plus de six mois, et part
    # dont le prix n'est pas daté du tout. Deux mesures et deux niveaux,
    # comme pour les rendements : « pas daté » est l'état ordinaire
    # aujourd'hui, et une alerte qui se déclenche sur chaque offre ne dit
    # plus rien.
    "part_prix_ancien": 0.20,
    "part_prix_sans_date": 0.20,
    # En deçà, les matériaux ne font pas le prix : les dater n'apprendrait
    # rien, et le dire ferait du bruit pour rien.
    "poids_materiaux": 0.05,
}

CRITIQUE, ATTENTION, INFO = "critique", "attention", "info"


def _alerte(niveau, code, titre, detail, poste=None):
    return {"niveau": niveau, "code": code, "titre": titre,
            "detail": detail, "poste": poste}


# ══════════════════════════════════════════
#  Couverture horaire — le vrai indicateur interne
# ══════════════════════════════════════════


def _taux_horaire_plancher(params):
    """
    Ce qu'une heure d'atelier doit rapporter pour ne pas coûter.

    Coût entreprise complet de la main-d'œuvre, majoré des frais
    généraux et des frais de chantier. Marge et aléas EXCLUS : on
    cherche le point où l'on cesse de gagner, pas celui où l'on vise.
    """
    p = params or PARAMS
    taux = [r["pu_res"] for r in RESSOURCES_PAR_CODE.values()
            if r["type_res"] == "MO"]
    moyen = sum(taux) / len(taux) if taux else 0.0
    return moyen * (1 + p["fg"]) * (1 + p["fc"])


def couverture_horaire(chiffres, bordereau, params=None, rabais=0.0):
    """
    Ce que l'offre laisse par heure de main-d'œuvre, matériaux et
    matériel payés.

        (montant encaissé − matériaux − matériel) / heures

    À comparer au plancher : sous lui, chaque heure travaillée coûte
    de l'argent à l'entreprise. C'est la définition interne d'un prix
    anormalement bas — celle qu'on peut vérifier sans connaître les
    concurrents.

    `rabais` (0,05 = 5 %) simule une remise commerciale sur le total.
    """
    total = sum(x["pu_vente"] * x["qte"] for x in chiffres)
    encaisse = total * (1 - rabais)
    achats = sum((bordereau[x["code_ouv"]]["deb_mat"]
                   + bordereau[x["code_ouv"]]["deb_eqp"]) * x["qte"]
                  for x in chiffres)
    heures = sum(bordereau[x["code_ouv"]]["heures_mo"] * x["qte"]
                  for x in chiffres)
    plancher = _taux_horaire_plancher(params)
    par_heure = (encaisse - achats) / heures if heures else 0.0
    return {
        "total": round(total, 2),
        "encaisse": round(encaisse, 2),
        "achats": round(achats, 2),
        "heures": round(heures, 2),
        "par_heure": round(par_heure, 2),
        "plancher": round(plancher, 2),
        "couvre": par_heure >= plancher,
    }


def rabais_maximal(chiffres, bordereau, params=None):
    """
    La remise au-delà de laquelle l'offre cesse de couvrir ses coûts.

    Le chiffre qu'on veut connaître AVANT de négocier, pas après.
    Rendu en fraction (0,082 = 8,2 %). Zéro si l'offre ne couvre
    déjà pas.
    """
    base = couverture_horaire(chiffres, bordereau, params)
    if not base["total"] or not base["heures"]:
        return 0.0
    # Point d'équilibre : encaissé = achats + heures × plancher
    equilibre = base["achats"] + base["heures"] * base["plancher"]
    marge = (base["total"] - equilibre) / base["total"]
    return round(max(0.0, marge), 4)


# ══════════════════════════════════════════
#  Analyse complète
# ══════════════════════════════════════════


def analyser(chiffres, bordereau=None, params=None, rabais=0.0,
              historique=None, aujourdhui=None):
    """
    Relit une offre chiffrée et remonte ce qui mérite un œil humain.

    chiffres : [{code_ouv, qte, pu_vente, ...}] — la sortie de
               metre_io.remplir_metre() convient après renommage,
               celle de moteur.devis() directement.
    historique : {code_ouv: pu pratiqué ailleurs} — facultatif.
                 Vide tant qu'aucun marché n'a été déposé.

    Retourne {indicateurs, alertes[], postes[], materiaux[]}. Les
    alertes sont triées du plus grave au plus léger.
    """
    b = bordereau if bordereau is not None else calcul_bordereau(params)
    historique = historique or {}
    # Le statut vient du moteur, pas d'une liste lue à part : deux
    # définitions de « validé » finiraient par diverger, et un poste
    # passerait pour sûr ici et douteux à l'écran.
    par_statut = codes_par_statut()
    a_valider = par_statut["a_valider"] | par_statut["a_revalider"]
    a_revalider = par_statut["a_revalider"]

    couverture = couverture_horaire(chiffres, b, params, rabais)
    total = couverture["total"]
    alertes, postes = [], []

    montant_non_valide = montant_a_revalider = 0.0

    for ligne in chiffres:
        code = ligne["code_ouv"]
        ref = b[code]
        montant = ligne["pu_vente"] * ligne["qte"]
        part = montant / total if total else 0.0
        debourse = ref["debourse_sec"] or 1e-9
        part_mo = ref["deb_mo"] / debourse
        part_mat = ref["deb_mat"] / debourse

        if code in a_valider:
            montant_non_valide += montant
        if code in a_revalider:
            montant_a_revalider += montant

        postes.append({
            "code_ouv": code,
            "libelle": ref["libelle_ouv"],
            "lot": ref["lot"],
            "montant": round(montant, 2),
            "part": round(part, 4),
            "part_mo": round(part_mo, 4),
            "part_mat": round(part_mat, 4),
            "heures": round(ref["heures_mo"] * ligne["qte"], 2),
            "rendement_a_valider": code in a_valider,
            "rendement_a_revalider": code in a_revalider,
        })

        if part >= SEUILS["poids_poste"]:
            alertes.append(_alerte(
                ATTENTION, "poids_poste",
                f"{code} pèse {part * 100:.0f} % de l'offre",
                "Une erreur sur ce seul poste déplace le total. "
                "À revérifier en priorité — quantité et prix.",
                code))

        if part >= 0.05 and part_mo >= SEUILS["part_mo"]:
            alertes.append(_alerte(
                INFO, "main_oeuvre_dominante",
                f"{code} : {part_mo * 100:.0f} % de main-d'œuvre",
                "Le prix tient au rendement, donc à une estimation "
                "d'expérience. C'est le poste à confronter au vécu "
                "de chantier avant tout autre.",
                code))
        elif part >= 0.05 and part_mat >= SEUILS["part_mat"]:
            alertes.append(_alerte(
                INFO, "materiaux_dominants",
                f"{code} : {part_mat * 100:.0f} % de matériaux",
                "Le prix tient à un prix d'achat. Le risque est la "
                "tenue de l'offre fournisseur sur la durée du marché.",
                code))

        attendu = historique.get(code)
        if attendu:
            ecart = (ligne["pu_vente"] - attendu) / attendu
            if abs(ecart) >= SEUILS["ecart_historique"]:
                alertes.append(_alerte(
                    ATTENTION, "ecart_historique",
                    f"{code} : {ecart * 100:+.0f} % par rapport à "
                    f"{attendu:.2f} € pratiqué ailleurs",
                    "Soit une décision commerciale, soit une "
                    "dérive. Les deux se défendent — mais elle doit "
                    "être consciente.",
                    code))

    # ── Alertes sur l'offre entière ─────────────────
    if not couverture["couvre"]:
        alertes.append(_alerte(
            CRITIQUE, "couverture",
            f"L'offre ne couvre pas ses coûts : "
            f"{couverture['par_heure']:.2f} €/h contre un plancher de "
            f"{couverture['plancher']:.2f} €/h",
            "Matériaux et matériel payés, chaque heure travaillée "
            "coûte de l'argent à l'entreprise. C'est un prix "
            "anormalement bas au sens propre — celui du bilan."))

    # Deux mesures, et deux niveaux : les confondre ferait crier au feu
    # sur chaque offre. « Jamais confronté » est l'état ORDINAIRE d'une
    # bibliothèque jeune — c'est une information, pas une alerte, et une
    # alerte qui se déclenche toujours ne dit plus rien.
    part_non_validee = montant_non_valide / total if total else 0.0
    if part_non_validee >= SEUILS["part_non_validee"]:
        alertes.append(_alerte(
            INFO, "rendements_non_valides",
            f"{part_non_validee * 100:.0f} % du montant repose sur des "
            f"rendements jamais confrontés à un chantier",
            "C'est l'état ordinaire tant que la calibration n'a pas eu "
            "lieu : la bibliothèque a été reconstruite, pas relevée. "
            "Trois relevés concordants suffisent à en valider un."))

    # Celui-là, en revanche, est anormal : quelqu'un a corrigé un
    # rendement APRÈS l'avoir confirmé sur des chantiers. Soit la
    # correction est juste et il faut revalider, soit elle est fausse.
    part_a_revalider = montant_a_revalider / total if total else 0.0
    if part_a_revalider >= SEUILS["part_a_revalider"]:
        alertes.append(_alerte(
            ATTENTION, "rendements_a_revalider",
            f"{part_a_revalider * 100:.0f} % du montant repose sur des "
            f"rendements validés puis modifiés",
            "Ces rendements avaient été confrontés à des chantiers, et "
            "leur valeur a changé depuis. La validation ne vaut plus : "
            "soit la correction est juste et il faut la reconfronter, "
            "soit elle est de trop."))

    # ── L'âge des prix d'achat ──────────────────────
    # Un rendement se valide, un prix d'achat périme. Et on ne regarde
    # QUE les matériaux que cette offre-ci consomme : signaler les
    # trente-cinq de la bibliothèque, c'est se faire ignorer.
    materiaux = materiaux_de(chiffres, aujourdhui=aujourdhui)
    achats_mat = sum(m["montant"] for m in materiaux)
    part_materiaux = achats_mat / total if total else 0.0
    anciens = [m for m in materiaux if m["ancien"]]
    sans_date = [m for m in materiaux if m["age"] is None]
    part_prix_ancien = (sum(m["montant"] for m in anciens) / achats_mat
                         if achats_mat else 0.0)
    part_prix_sans_date = (sum(m["montant"] for m in sans_date) / achats_mat
                            if achats_mat else 0.0)
    materiaux_comptent = part_materiaux >= SEUILS["poids_materiaux"]

    if materiaux_comptent and part_prix_ancien >= SEUILS["part_prix_ancien"]:
        cites = ", ".join(
            f"{m['code_res']} {m['libelle_res'].lower()} ({mois(m['age'])} mois)"
            for m in anciens[:3])
        alertes.append(_alerte(
            ATTENTION, "prix_anciens",
            f"{len(anciens)} matériau(x) sur {len(materiaux)} ont un prix de "
            f"plus de {mois(PRIX_ANCIEN_JOURS)} mois — "
            f"{part_prix_ancien * 100:.0f} % des achats de l'offre",
            f"{cites}. Ces prix étaient justes le jour où ils ont été "
            f"relevés ; rien ne dit qu'ils tiennent encore, et le "
            f"fournisseur ne prévient pas. À reconfronter à une offre "
            f"avant de déposer — un marché public ne se renégocie pas."))

    # Aucune ressource n'est datée aujourd'hui : cette information doit
    # se dire, mais en INFO. C'est le même raisonnement que pour les
    # rendements jamais confrontés — l'état ordinaire d'un outil jeune
    # n'est pas une alerte.
    if materiaux_comptent \
            and part_prix_sans_date >= SEUILS["part_prix_sans_date"]:
        alertes.append(_alerte(
            INFO, "prix_sans_date",
            f"{part_prix_sans_date * 100:.0f} % des achats de l'offre "
            f"reposent sur des prix dont on ignore la date",
            "Un prix sans date n'est pas un prix frais : c'est un prix "
            "d'origine inconnue. Corriger un prix d'achat depuis "
            "l'atelier le date — les dates viendront d'elles-mêmes, "
            "matériau par matériau, à mesure des offres fournisseurs."))

    tries = sorted(postes, key=lambda p: -p["montant"])
    concentration = sum(p["part"] for p in tries[:3])

    ordre = {CRITIQUE: 0, ATTENTION: 1, INFO: 2}
    alertes.sort(key=lambda a: (ordre[a["niveau"]], a["code"]))

    return {
        "indicateurs": {
            **couverture,
            "coefficient_k": round(coefficient_k(params), 4),
            "rabais_applique": rabais,
            "rabais_maximal": rabais_maximal(chiffres, b, params),
            "part_non_validee": round(part_non_validee, 4),
            "part_a_revalider": round(part_a_revalider, 4),
            "achats_materiaux": round(achats_mat, 2),
            "part_materiaux": round(part_materiaux, 4),
            "part_prix_ancien": round(part_prix_ancien, 4),
            "part_prix_sans_date": round(part_prix_sans_date, 4),
            "nb_materiaux": len(materiaux),
            "concentration_top3": round(concentration, 4),
            "nb_postes": len(postes),
        },
        "alertes": alertes,
        "postes": tries,
        "materiaux": materiaux,
    }


def imprimer(rapport):
    """Rapport lisible en console."""
    i = rapport["indicateurs"]
    out = [
        f"Total                  {i['total']:>12,.2f} €".replace(",", " "),
        f"Heures de main-d'œuvre {i['heures']:>12,.1f} h".replace(",", " "),
        f"Par heure travaillée   {i['par_heure']:>12,.2f} €"
        f"   (plancher {i['plancher']:.2f})".replace(",", " "),
        f"Rabais maximal         {i['rabais_maximal'] * 100:>12.1f} %",
        f"Top 3 des postes       {i['concentration_top3'] * 100:>12.0f} % du total",
        "",
    ]
    if not rapport["alertes"]:
        out.append("✅ Aucun point d'attention.")
    for a in rapport["alertes"]:
        marque = {CRITIQUE: "🛑", ATTENTION: "⚠️ ", INFO: "🛈 "}[a["niveau"]]
        out.append(f"{marque} {a['titre']}")
        out.append(f"    {a['detail']}")
    return "\n".join(out)


def lots_du_rapport(rapport):
    """Répartition du montant par lot — utile en graphique."""
    par_lot = {}
    for poste in rapport["postes"]:
        par_lot.setdefault(poste["lot"], 0.0)
        par_lot[poste["lot"]] += poste["montant"]
    return [
        {"lot": lot, "libelle": LOTS.get(lot, ""), "montant": round(m, 2)}
        for lot, m in sorted(par_lot.items())
    ]
