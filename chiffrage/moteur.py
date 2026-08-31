"""
╔══════════════════════════════════════════════════════════════════════════╗
║  BAG BATTER SRL — MOTEUR DE CALCUL                                       ║
║  Déboursé sec -> prix de vente unitaire -> devis -> calibration          ║
╚══════════════════════════════════════════════════════════════════════════╝

    debourse_sec = somme(qte_res x pu_res) sur toutes les ressources
    pu_vente     = debourse_sec x K
    K            = (1+FG) x (1+FC) x (1+aleas) x (1+marge)

Python pur, aucune dépendance : ce module tourne en CI et se colle tel quel
dans une cellule Colab.
"""

from .bibliotheque import (
    COMPOSITION,
    RESSOURCES,
    LOTS,
    METRES_HISTO,
    OUVRAGES,
    OUVRAGES_A_VALIDER,
    OUVRAGES_PAR_CODE,
    PARAMS,
    RELEVES,
    RESSOURCES_PAR_CODE,
)

# ═══════════════════════════════════════════════════════════════════════════
# Utilitaires de présentation
# ═══════════════════════════════════════════════════════════════════════════


def euro(montant):
    """1234.5 -> '1.234,50 €' (convention belge : point milliers, virgule décimale)."""
    s = f"{montant:,.2f}"
    return s.replace(",", " ").replace(".", ",").replace(" ", ".") + " €"


# ═══════════════════════════════════════════════════════════════════════════
# 1. Coefficient de vente
# ═══════════════════════════════════════════════════════════════════════════


def tables_courantes(tables=None):
    """
    Les tables sur lesquelles calculer : celles du dépôt, ou d'autres.

    Le moteur lisait directement les tables du module. Il ne pouvait
    donc rien calculer sur des valeurs en cours de correction — or
    c'est précisément ce qu'il faut pour une séance de calibration :
    changer un rendement et voir l'écart bouger AVANT d'enregistrer.
    """
    if tables is not None:
        return tables
    return {
        "ressources": RESSOURCES,
        "ouvrages": OUVRAGES,
        "composition": COMPOSITION,
        "lots": LOTS,
        "metres_histo": METRES_HISTO,
        # Aucun calcul ne s'en sert : c'est la liste des rendements
        # jamais confrontés au réel. Elle voyage avec les autres tables
        # parce qu'une séance de calibration la modifie — on lève le
        # doute sur un rendement au moment même où on le corrige.
        "ouvrages_a_valider": OUVRAGES_A_VALIDER,
        # Même raison, et même prudence : le journal des chantiers ne
        # porte aucun prix, mais il voyage avec les tables parce qu'une
        # séance en ajoute — on enregistre le relevé au moment où on
        # s'en sert pour corriger.
        "releves": RELEVES,
        "ressources_par_code": RESSOURCES_PAR_CODE,
        "ouvrages_par_code": OUVRAGES_PAR_CODE,
    }


def coefficient_k(params=None):
    """K = (1+FG)(1+FC)(1+aléas)(1+marge). Avec les PARAMS par défaut : 1,3324."""
    p = params or PARAMS
    return (1 + p["fg"]) * (1 + p["fc"]) * (1 + p["aleas"]) * (1 + p["marge"])


# ═══════════════════════════════════════════════════════════════════════════
# 2. Bordereau — le cœur du calcul
# ═══════════════════════════════════════════════════════════════════════════


def calcul_bordereau(params=None, tables=None):
    """
    Développe chaque ouvrage en déboursés et en prix de vente unitaire.

    Retourne un dict {code_ouv: {...}} avec, par ouvrage :
        lot, libelle_ouv, unite_ouv, code_ref,
        deb_mo, deb_mat, deb_eqp   déboursés par nature (€/unité)
        debourse_sec               somme des trois
        pu_vente                   debourse_sec x K
        heures_mo                  total des heures de main-d'œuvre par unité
        nb_ressources              nombre de lignes de composition

    Un ouvrage sans ligne de composition sort à 0 € : c'est une erreur de
    saisie, pas un ouvrage gratuit. controle_coherence() la signale.
    """
    k = coefficient_k(params)
    t = tables_courantes(tables)
    par_res = t["ressources_par_code"]
    lots = t["lots"]

    bordereau = {}
    for ouv in t["ouvrages"]:
        bordereau[ouv["code_ouv"]] = {
            "code_ouv": ouv["code_ouv"],
            "lot": ouv["lot"],
            "libelle_lot": lots.get(ouv["lot"], ""),
            "libelle_ouv": ouv["libelle_ouv"],
            "unite_ouv": ouv["unite_ouv"],
            "code_ref": ouv["code_ref"],
            "deb_mo": 0.0,
            "deb_mat": 0.0,
            "deb_eqp": 0.0,
            "heures_mo": 0.0,
            "nb_ressources": 0,
        }

    for comp in t["composition"]:
        ligne = bordereau.get(comp["code_ouv"])
        res = par_res.get(comp["code_res"])
        if ligne is None or res is None:
            # Référence orpheline : signalée par controle_coherence(), ignorée
            # ici pour ne pas faire exploser un chiffrage en cours.
            continue
        montant = comp["qte_res"] * res["pu_res"]
        if res["type_res"] == "MO":
            ligne["deb_mo"] += montant
            ligne["heures_mo"] += comp["qte_res"]
        elif res["type_res"] == "MAT":
            ligne["deb_mat"] += montant
        else:
            ligne["deb_eqp"] += montant
        ligne["nb_ressources"] += 1

    for ligne in bordereau.values():
        ligne["debourse_sec"] = round(
            ligne["deb_mo"] + ligne["deb_mat"] + ligne["deb_eqp"], 4
        )
        ligne["pu_vente"] = round(ligne["debourse_sec"] * k, 2)
        for champ in ("deb_mo", "deb_mat", "deb_eqp", "heures_mo"):
            ligne[champ] = round(ligne[champ], 4)

    return bordereau


# ═══════════════════════════════════════════════════════════════════════════
# 3. Devis
# ═══════════════════════════════════════════════════════════════════════════


def devis(nom, lignes, tva=None, params=None, bordereau=None,
           tables=None):
    """
    Chiffre une liste de lignes [(code_ouv, quantite), ...].

    tva : 0.06 par défaut (logement privé > 10 ans, consommateur final).
          EN MARCHÉ PUBLIC : passer tva=0.21.

    Retourne un dict :
        nom, lignes[], total_ht, tva_taux, montant_tva, total_ttc,
        heures_mo, jours_homme, inconnus[]

    `inconnus` liste les codes d'ouvrage absents de la bibliothèque : ils ne
    sont PAS chiffrés et ne sont pas silencieusement ignorés.
    """
    b = (bordereau if bordereau is not None
          else calcul_bordereau(params, tables))
    taux = PARAMS["tva"] if tva is None else tva

    detail, inconnus = [], []
    total_ht = heures = 0.0
    for code_ouv, qte in lignes:
        ref = b.get(code_ouv)
        if ref is None:
            inconnus.append(code_ouv)
            continue
        montant = round(ref["pu_vente"] * qte, 2)
        detail.append(
            {
                "code_ouv": code_ouv,
                "libelle_ouv": ref["libelle_ouv"],
                "unite_ouv": ref["unite_ouv"],
                "qte": qte,
                "pu_vente": ref["pu_vente"],
                "montant": montant,
                "heures_mo": round(ref["heures_mo"] * qte, 2),
                "debourse_sec": round(ref["debourse_sec"] * qte, 2),
            }
        )
        total_ht += montant
        heures += ref["heures_mo"] * qte

    total_ht = round(total_ht, 2)
    montant_tva = round(total_ht * taux, 2)
    return {
        "nom": nom,
        "lignes": detail,
        "total_ht": total_ht,
        "tva_taux": taux,
        "montant_tva": montant_tva,
        "total_ttc": round(total_ht + montant_tva, 2),
        "heures_mo": round(heures, 2),
        "jours_homme": round(heures / 8.0, 2),
        "inconnus": inconnus,
    }


def imprimer_devis(d):
    """Rend un devis lisible en texte (console ou cellule Colab)."""
    out = [f"── {d['nom']} " + "─" * max(0, 66 - len(d["nom"]))]
    out.append(f"{'Code':<8}{'Désignation':<46}{'Un.':<7}{'Qté':>8}"
               f"{'PU':>12}{'Montant':>13}")
    for ligne in d["lignes"]:
        lib = ligne["libelle_ouv"]
        lib = lib if len(lib) <= 44 else lib[:41] + "..."
        out.append(
            f"{ligne['code_ouv']:<8}{lib:<46}{ligne['unite_ouv']:<7}"
            f"{ligne['qte']:>8.2f}{ligne['pu_vente']:>12.2f}"
            f"{ligne['montant']:>13.2f}"
        )
    out.append("─" * 94)
    out.append(f"{'Total HTVA':>81}{d['total_ht']:>13.2f}")
    out.append(f"{'TVA ' + format(d['tva_taux'] * 100, '.0f') + ' %':>81}"
               f"{d['montant_tva']:>13.2f}")
    out.append(f"{'TOTAL TVAC':>81}{d['total_ttc']:>13.2f}")
    out.append(f"Main-d'œuvre : {d['heures_mo']:.1f} h "
               f"({d['jours_homme']:.1f} jours-homme)")
    if d["inconnus"]:
        out.append("⚠️  Codes inconnus, NON chiffrés : " + ", ".join(d["inconnus"]))
    return "\n".join(out)


# ═══════════════════════════════════════════════════════════════════════════
# 4. Fiche de justification de prix
# ═══════════════════════════════════════════════════════════════════════════


def fiche_prix(code_ouv, params=None, bordereau=None, tables=None):
    """
    Décomposition d'un prix unitaire, ligne de composition par ligne de
    composition. C'est la pièce à produire quand le pouvoir adjudicateur
    demande la justification d'un prix jugé anormal (art. 36 AR 18/04/2017).
    """
    t = tables_courantes(tables)
    ouv = t["ouvrages_par_code"].get(code_ouv)
    if ouv is None:
        raise KeyError(f"Ouvrage inconnu : {code_ouv}")
    b = (bordereau if bordereau is not None
          else calcul_bordereau(params, tables))
    ref = b[code_ouv]
    p = params or PARAMS

    out = [
        f"FICHE DE JUSTIFICATION DE PRIX — {code_ouv}",
        f"{ouv['libelle_ouv']}",
        f"Lot {ouv['lot']} · {t['lots'].get(ouv['lot'], '')} "
        f"· unité : {ouv['unite_ouv']}",
        "",
        f"{'Ressource':<10}{'Désignation':<40}{'Type':<6}{'Un.':<8}"
        f"{'Qté':>9}{'PU':>10}{'Montant':>11}",
    ]
    for comp in t["composition"]:
        if comp["code_ouv"] != code_ouv:
            continue
        res = t["ressources_par_code"][comp["code_res"]]
        lib = res["libelle_res"]
        lib = lib if len(lib) <= 38 else lib[:35] + "..."
        out.append(
            f"{res['code_res']:<10}{lib:<40}{res['type_res']:<6}"
            f"{res['unite_res']:<8}{comp['qte_res']:>9.3f}"
            f"{res['pu_res']:>10.2f}{comp['qte_res'] * res['pu_res']:>11.2f}"
        )
    out += [
        "",
        f"  Déboursé main-d'œuvre  {ref['deb_mo']:>10.2f}   "
        f"({ref['heures_mo']:.3f} h/{ouv['unite_ouv']})",
        f"  Déboursé matériaux     {ref['deb_mat']:>10.2f}",
        f"  Déboursé matériel      {ref['deb_eqp']:>10.2f}",
        f"  DÉBOURSÉ SEC           {ref['debourse_sec']:>10.2f}",
        "",
        f"  Frais généraux  {p['fg'] * 100:>5.1f} %",
        f"  Frais chantier  {p['fc'] * 100:>5.1f} %",
        f"  Aléas           {p['aleas'] * 100:>5.1f} %",
        f"  Marge           {p['marge'] * 100:>5.1f} %",
        f"  Coefficient K   {coefficient_k(p):>5.4f}",
        "",
        f"  PRIX DE VENTE UNITAIRE HTVA  {ref['pu_vente']:>10.2f} €"
        f" / {ouv['unite_ouv']}",
    ]
    return "\n".join(out)


# ═══════════════════════════════════════════════════════════════════════════
# 5. Relevé de chantier
# ═══════════════════════════════════════════════════════════════════════════


def releve_rendement(code_ouv, quantite, heures, tables=None):
    """
    Convertit un relevé de chantier en rendement.

    Sur un chantier, PERSONNE ne connaît un rendement. Ce qu'on y sait le
    soir, c'est « on a fait 12 m2, à deux, de 8 h à 11 h 30 ». Le rendement
    en est le quotient — heures / quantité — et le calculer de tête, sur
    place, est exactement ce que cet outil existe pour supprimer.

    `heures` est le total des heures d'HOMME, pas la durée : deux ouvriers
    pendant 3 h 30 font 7 h. La confusion diviserait le rendement par le
    nombre d'ouvriers, sans que rien ne le signale.

    Répartition entre plusieurs lignes de main-d'œuvre — 7 ouvrages sur 49
    en ont deux, un chef et un ouvrier par exemple : le relevé donne un
    total, pas un partage. On garde donc la PROPORTION actuelle, faute de
    mieux, et l'appelant doit le dire à l'écran. Inventer un partage en
    silence serait pire que de le reprendre tel quel.

    Retourne un dict décrivant le relevé — il ne corrige rien : c'est
    l'humain qui applique, ou pas. Voir `calcul_bordereau` pour l'effet.
    """
    t = tables_courantes(tables)
    ouv = t["ouvrages_par_code"].get(code_ouv)
    if ouv is None:
        raise KeyError(f"Ouvrage inconnu : {code_ouv}")
    if quantite <= 0:
        raise ValueError(
            "Quantité réalisée nulle ou négative : des heures sans quantité "
            "ne disent rien d'un rendement.")
    if heures <= 0:
        raise ValueError(
            "Heures relevées nulles ou négatives : un rendement nul rendrait "
            "la main-d'œuvre gratuite.")

    par_res = t["ressources_par_code"]
    lignes_mo = [c for c in t["composition"]
                  if c["code_ouv"] == code_ouv
                  and par_res[c["code_res"]]["type_res"] == "MO"]
    if not lignes_mo:
        raise ValueError(
            f"Ouvrage {code_ouv} : aucune ligne de main-d'œuvre — un relevé "
            f"d'heures n'y a rien à corriger.")

    actuel = sum(c["qte_res"] for c in lignes_mo)
    if actuel <= 0 and len(lignes_mo) > 1:
        # Impossible via le chargeur, qui refuse une quantité <= 0 ; possible
        # dans une table en cours de correction. Aucune proportion ne se
        # déduit de zéro : on refuse plutôt que de partager au hasard.
        raise ValueError(
            f"Ouvrage {code_ouv} : rendement actuel nul sur plusieurs lignes "
            f"de main-d'œuvre — les heures relevées ne peuvent pas être "
            f"réparties.")

    observe = heures / quantite
    return {
        "code_ouv": code_ouv,
        "unite_ouv": ouv["unite_ouv"],
        "quantite": quantite,
        "heures": heures,
        "rendement_observe": round(observe, 4),
        "rendement_actuel": round(actuel, 4),
        "ecart": round((observe - actuel) / actuel, 4) if actuel else None,
        "lignes": [
            {
                "code_res": c["code_res"],
                "libelle_res": par_res[c["code_res"]]["libelle_res"],
                "qte_res": c["qte_res"],
                "part": round(c["qte_res"] / actuel, 4) if actuel else 1.0,
                "propose": round(
                    observe * (c["qte_res"] / actuel if actuel else 1.0), 4),
            }
            for c in lignes_mo
        ],
    }


def releves_de(code_ouv, tables=None):
    """Les relevés de chantier d'un ouvrage, du plus ancien au plus récent.

    Chacun porte son `rendement`, CALCULÉ et non stocké : heures /
    quantité. Le stocker en ferait une seconde vérité, qui finirait par
    diverger de ses deux termes.
    """
    t = tables_courantes(tables)
    retenus = [r for r in (t.get("releves") or [])
                if r.get("code_ouv") == code_ouv
                and isinstance(r.get("quantite"), (int, float))
                and isinstance(r.get("heures"), (int, float))
                and r["quantite"] > 0 and r["heures"] > 0]
    return sorted(
        (dict(r, rendement=round(r["heures"] / r["quantite"], 4))
          for r in retenus),
        key=lambda r: (str(r.get("date") or ""), str(r.get("chantier") or "")))


def rendement_constate(code_ouv, tables=None):
    """Ce que les chantiers disent du rendement d'un ouvrage.

    L'agrégat est Σheures / Σquantités, PAS la moyenne des rendements :
    2 m2 en 3 h et 40 m2 en 20 h ne pèsent pas pareil, et une moyenne
    simple donnerait au tout petit chantier le même poids qu'au grand.

    Rend aussi `n`, `mini` et `maxi` : un seul relevé n'est pas une
    vérité, et deux relevés très écartés ne se résument pas à leur
    milieu. Un nombre seul se prendrait pour une mesure.

    Rend None quand aucun chantier n'a été relevé — c'est le cas
    ordinaire aujourd'hui, et il doit se distinguer d'un rendement de
    zéro.
    """
    releves = releves_de(code_ouv, tables)
    if not releves:
        return None

    t = tables_courantes(tables)
    par_res = t["ressources_par_code"]
    actuel = sum(c["qte_res"] for c in t["composition"]
                  if c["code_ouv"] == code_ouv
                  and par_res[c["code_res"]]["type_res"] == "MO")

    heures = sum(r["heures"] for r in releves)
    quantite = sum(r["quantite"] for r in releves)
    constate = heures / quantite
    rendements = [r["rendement"] for r in releves]
    return {
        "code_ouv": code_ouv,
        "n": len(releves),
        "heures": round(heures, 4),
        "quantite": round(quantite, 4),
        "rendement": round(constate, 4),
        "mini": min(rendements),
        "maxi": max(rendements),
        "rendement_actuel": round(actuel, 4),
        "ecart": round((constate - actuel) / actuel, 4) if actuel else None,
    }


def fusionner_releves(*sources):
    """Réunit plusieurs jeux de relevés sans en perdre ni en doubler.

    Un journal de chantier s'AJOUTE, il ne s'écrase pas : deux téléphones
    peuvent enregistrer le même soir, et la sémantique « le dernier qui
    écrit gagne » — juste pour une table de prix, qui est un tout
    cohérent — perdrait ici l'observation de l'autre.

    Deux relevés identiques sur les cinq champs sont la même observation
    saisie deux fois, pas deux chantiers : on n'en garde qu'un.
    """
    cle = lambda r: (str(r.get("code_ouv")), str(r.get("date")),  # noqa: E731
                      str(r.get("chantier")), r.get("quantite"), r.get("heures"))
    fusion, vus = [], set()
    for source in sources:
        for rel in source or []:
            if cle(rel) in vus:
                continue
            vus.add(cle(rel))
            fusion.append(rel)
    return sorted(fusion, key=lambda r: (str(r.get("date") or ""),
                                          str(r.get("code_ouv") or "")))


# ═══════════════════════════════════════════════════════════════════════════
# 6. Calibration sur les devis historiques
# ═══════════════════════════════════════════════════════════════════════════


def calibration(params=None, tables=None):
    """
    Re-chiffre les 6 devis forfaitaires historiques et compare au montant
    réellement vendu.

    Écart = (calculé − forfait) / forfait.
      · écart NÉGATIF -> la bibliothèque chiffre moins cher que ce qui a été
        vendu : soit les rendements sont optimistes, soit le devis était bien
        margé.
      · écart POSITIF fort -> le chantier a été vendu sous son coût analytique,
        OU les quantités estimées dans METRES_HISTO sont trop élevées. Les
        deux hypothèses restent ouvertes tant que le client n'a pas fourni les
        surfaces réelles (cf. CLAUDE.md §4).

    Objectif fixé : |écart| < 15 % sur chaque ligne.
    """
    t = tables_courantes(tables)
    b = calcul_bordereau(params, tables)
    resultats = []
    for num, info in sorted(t["metres_histo"].items()):
        d = devis(f"Devis {num} — {info['objet']}", info["lignes"], bordereau=b)
        forfait = info["forfait"]
        ecart = (d["total_ht"] - forfait) / forfait if forfait else 0.0
        resultats.append(
            {
                "devis": num,
                "objet": info["objet"],
                "forfait": forfait,
                "calcule": d["total_ht"],
                "ecart": round(ecart, 4),
                "heures_mo": d["heures_mo"],
                "prix_horaire_implicite": (
                    round(forfait / d["heures_mo"], 2) if d["heures_mo"] else None
                ),
                "inconnus": d["inconnus"],
            }
        )
    if resultats:
        moyenne = sum(abs(r["ecart"]) for r in resultats) / len(resultats)
    else:
        moyenne = 0.0
    return {"lignes": resultats, "ecart_moyen_absolu": round(moyenne, 4)}


def analyser_ecart(num, params=None, tables=None):
    """
    Décompose l'écart d'UN devis historique, poste par poste.

    La calibration ne compare que des totaux : elle dit qu'un devis est à
    +42,7 %, pas d'où ça vient. Or deux causes opposées produisent le même
    total, et l'outil ne peut pas trancher entre elles — le chef
    d'entreprise, si, à condition de voir les bons chiffres :

      1. les quantités de METRES_HISTO sont trop élevées — elles viennent
         des descriptifs des devis PDF, pas de relevés ;
      2. le chantier a été vendu sous son coût analytique.

    Trois chiffres les séparent, et cette fonction les rend SANS conclure :

      · `k_implicite` — ce que le forfait vendu représente par rapport au
        déboursé sec. En dessous de 1, le chantier n'a pas couvert ses
        propres achats et heures : l'hypothèse 2 devient difficile à
        écarter. Autour de 1, il a été vendu sans marge ni frais
        généraux. La cible est K (1,3324 aujourd'hui).

      · `facteur_quantites` — le facteur UNIFORME qu'il faudrait appliquer
        à toutes les quantités pour annuler l'écart. « Il faudrait des
        quantités inférieures de 30 % » se discute avec quelqu'un qui
        était sur le chantier : à −5 % c'est crédible, à −30 % ça ne
        l'est plus, et c'est l'hypothèse 2 qui reste.

      · `concentration` — la part des trois plus gros postes. Un écart
        porté par un seul poste se règle en vérifiant SA quantité ; un
        écart réparti sur tous est un biais systématique des rendements
        ou du prix de vente.

    Les postes sortent triés par montant décroissant : on regarde d'abord
    ce qui pèse.
    """
    t = tables_courantes(tables)
    info = t["metres_histo"].get(str(num))
    if info is None:
        raise KeyError(f"Devis historique inconnu : {num}")

    b = calcul_bordereau(params, tables)
    d = devis(f"Devis {num} — {info['objet']}", info["lignes"],
               bordereau=b, tables=tables)
    forfait = info["forfait"]
    calcule = d["total_ht"]
    debourse = round(sum(x["debourse_sec"] for x in d["lignes"]), 2)

    lignes = sorted(d["lignes"], key=lambda x: x["montant"], reverse=True)
    for ligne in lignes:
        ligne["part"] = round(ligne["montant"] / calcule, 4) if calcule else 0.0

    return {
        "devis": str(num),
        "objet": info["objet"],
        "forfait": forfait,
        "calcule": calcule,
        "ecart": round((calcule - forfait) / forfait, 4) if forfait else 0.0,
        "debourse_sec": debourse,
        # Ce que le forfait vendu vaut en coefficient. À comparer au K visé.
        "k_implicite": round(forfait / debourse, 4) if debourse else None,
        "k_vise": coefficient_k(params),
        "couvre_debourse": bool(debourse and forfait >= debourse),
        # < 1 : il aurait fallu MOINS de quantités pour tomber sur le forfait.
        "facteur_quantites": round(forfait / calcule, 4) if calcule else None,
        "heures_mo": d["heures_mo"],
        "prix_horaire_implicite": (round(forfait / d["heures_mo"], 2)
                                    if d["heures_mo"] else None),
        "concentration": round(
            sum(x["montant"] for x in lignes[:3]) / calcule, 4)
        if calcule else 0.0,
        "lignes": lignes,
        "inconnus": d["inconnus"],
    }


def imprimer_calibration(cal=None, params=None, tables=None):
    """Tableau de calibration lisible."""
    cal = cal or calibration(params, tables)
    out = [
        f"{'Devis':<7}{'Objet':<40}{'Forfait':>11}{'Calculé':>11}"
        f"{'Écart':>9}{'h MO':>8}{'€/h vendu':>11}",
        "─" * 97,
    ]
    for r in cal["lignes"]:
        objet = r["objet"] if len(r["objet"]) <= 38 else r["objet"][:35] + "..."
        alerte = "  ⚠️" if abs(r["ecart"]) > 0.15 else ""
        ph = f"{r['prix_horaire_implicite']:.0f}" if r["prix_horaire_implicite"] else "—"
        out.append(
            f"{r['devis']:<7}{objet:<40}{r['forfait']:>11.0f}{r['calcule']:>11.0f}"
            f"{r['ecart'] * 100:>8.1f}%{r['heures_mo']:>8.1f}{ph:>11}{alerte}"
        )
    out.append("─" * 97)
    out.append(f"Écart moyen absolu : {cal['ecart_moyen_absolu'] * 100:.1f} % "
               f"(cible : < 15 % sur chaque ligne)")
    return "\n".join(out)


# ═══════════════════════════════════════════════════════════════════════════
# 7. Contrôle de cohérence de la bibliothèque
# ═══════════════════════════════════════════════════════════════════════════


def controle_coherence():
    """
    Vérifie l'intégrité de la bibliothèque avant tout chiffrage.

    Retourne un dict de listes d'anomalies (toutes vides = bibliothèque saine) :
      · res_orphelines   : composition pointant une ressource inexistante
      · ouv_orphelins    : composition pointant un ouvrage inexistant
      · ouv_sans_compo   : ouvrage sans aucune ressource (prix = 0 €)
      · ouv_sans_mo      : ouvrage sans main-d'œuvre (suspect hors fournitures)
      · codes_dupliques  : code_res ou code_ouv en double
      · lots_inconnus    : ouvrage dont le lot n'est pas dans LOTS
      · releves_orphelins: relevé de chantier sur un ouvrage supprimé depuis
    """
    codes_res = set(RESSOURCES_PAR_CODE)
    codes_ouv = set(OUVRAGES_PAR_CODE)

    compo_par_ouv = {}
    res_orphelines, ouv_orphelins = [], []
    for comp in COMPOSITION:
        if comp["code_res"] not in codes_res:
            res_orphelines.append((comp["code_ouv"], comp["code_res"]))
        if comp["code_ouv"] not in codes_ouv:
            ouv_orphelins.append((comp["code_ouv"], comp["code_res"]))
        compo_par_ouv.setdefault(comp["code_ouv"], []).append(comp)

    ouv_sans_compo, ouv_sans_mo = [], []
    for code in codes_ouv:
        lignes = compo_par_ouv.get(code, [])
        if not lignes:
            ouv_sans_compo.append(code)
            continue
        if not any(
            RESSOURCES_PAR_CODE.get(c["code_res"], {}).get("type_res") == "MO"
            for c in lignes
        ):
            ouv_sans_mo.append(code)

    def _doublons(codes):
        vus, dup = set(), []
        for c in codes:
            if c in vus:
                dup.append(c)
            vus.add(c)
        return dup

    # Un relevé pointant un ouvrage disparu n'empêche aucun chiffrage :
    # il ne porte pas de prix. Mais c'est une preuve devenue muette, et
    # elle doit se voir plutôt que de dormir dans le fichier.
    releves_orphelins = sorted(
        {str(r.get("code_ouv")) for r in RELEVES
          if r.get("code_ouv") not in codes_ouv})

    return {
        "res_orphelines": res_orphelines,
        "ouv_orphelins": ouv_orphelins,
        "releves_orphelins": releves_orphelins,
        "ouv_sans_compo": sorted(ouv_sans_compo),
        "ouv_sans_mo": sorted(ouv_sans_mo),
        "codes_dupliques": (
            _doublons([r["code_res"] for r in RESSOURCES])
            + _doublons([o["code_ouv"] for o in OUVRAGES])
        ),
        "lots_inconnus": sorted(
            {o["lot"] for o in OUVRAGES if o["lot"] not in LOTS}
        ),
    }
