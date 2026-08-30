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
# 5. Calibration sur les devis historiques
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
# 6. Contrôle de cohérence de la bibliothèque
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

    return {
        "res_orphelines": res_orphelines,
        "ouv_orphelins": ouv_orphelins,
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
