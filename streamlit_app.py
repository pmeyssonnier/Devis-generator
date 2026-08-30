"""
╔══════════════════════════════════════════════════════╗
║  BAG BATTER SRL — interface de chiffrage                                  ║
║  Déposer un métré · apparier les postes · télécharger l'offre              ║
╚══════════════════════════════════════════════════════╝

Interface Streamlit au-dessus de `chiffrage/`. Aucune logique de prix ici :
tout passe par le moteur, pour qu'il n'y ait jamais deux vérités sur un prix.

Ce que cette interface apporte et que la ligne de commande ne pouvait pas :
l'appariement des postes à l'écran. `MAPPING` est à refaire à chaque marché
— les codes appartiennent au pouvoir adjudicateur — et c'était jusqu'ici la
seule étape qui obligeait à éditer du Python.

Lancement local :
    streamlit run streamlit_app.py
"""

import copy
import hashlib
import json
from datetime import date, datetime
from pathlib import Path
from tempfile import TemporaryDirectory

import streamlit as st

from chiffrage.bibliotheque import (
    ENTREPRISE,
    LOTS,
    MAPPING,
    OUVRAGES,
    OUVRAGES_A_VALIDER,
    PARAMS,
)
from chiffrage.devis_xlsx import exporter_devis
from chiffrage.controle_prix import analyser
from chiffrage.detection_colonnes import CHAMPS, CHAMPS_REQUIS
from chiffrage.depot_github import (
    ErreurDepot,
    commiter_lexique,
    commiter_parametres,
    commiter_table,
)
from chiffrage.justification_xlsx import exporter_justification
from chiffrage.lexique import (
    DEMOLITION,
    EXPRESSIONS,
    LOCAL,
    SURCOUCHE,
    SYNONYMES,
    adopter_local,
    ajouter_expression,
    ajouter_synonyme,
    est_demolition,
    surcouche_en_python,
    vider_surcouche,
)
from chiffrage.gen_metre import generer_metre
from openpyxl.utils import get_column_letter
from chiffrage.metre_io import (
    feuilles_avec_postes,
    lire_metre_complet,
    normaliser_unite,
    remplir_metre,
)
from chiffrage.moteur import (
    calcul_bordereau,
    calibration,
    coefficient_k,
    controle_coherence,
    devis,
    fiche_prix,
    tables_courantes,
)
from chiffrage.parametres import serialiser
from chiffrage.suggestion import normaliser, proposer_mapping, suggerer

st.set_page_config(page_title="Chiffrage BAG BATTER", page_icon="🧱",
                    layout="wide")

# Un score sous ce seuil n'est pas montré comme une suggestion : afficher du
# bruit ferait perdre plus de temps qu'il n'en fait gagner.
SEUIL_SUGGESTION = 0.35
# Au-dessus de ce score, la suggestion est présentée comme solide — ce qui ne
# veut PAS dire juste. Voir chiffrage/suggestion.py.
SEUIL_CONFIANCE = 0.60

SANS_OUVRAGE = "— ne pas chiffrer —"


def _bordereau(params=None):
    """Le bordereau AUX PARAMÈTRES COURANTS.

    Surtout pas de `@st.cache_data` sans clé ici. Le bordereau
    dépendait « du code seul » tant que les coefficients étaient
    figés ; depuis que la barre latérale les règle, un cache sans
    paramètre sert un prix calculé avec d'AUTRES coefficients que
    ceux affichés.

    Le symptôme mesuré, marge portée à 25 % : 185 308 € affichés à
    l'écran, 210 581 € dans le fichier réellement produit. Deux
    vérités sur le même écran, 13 % d'écart, sur un document qui
    part chez un pouvoir adjudicateur.

    Le calcul prend 0,11 ms pour 49 ouvrages. Le cache économisait
    ça, au prix d'un faux prix affiché.
    """
    return calcul_bordereau(params)


def _libelle_ouvrage(code_ouv, bordereau):
    ligne = bordereau[code_ouv]
    return (f"{code_ouv} · {ligne['libelle_ouv']} "
            f"· {ligne['pu_vente']:.2f} €/{ligne['unite_ouv']}")


def _euro(montant):
    """1234.5 -> '1.234,50 €' (convention belge)."""
    return f"{montant:,.2f}".replace(",", " ").replace(".", ",").replace(" ", ".") + " €"


def _avertissement_calibration():
    """Rappel affiché sur chaque page : les prix ne sont pas calibrés.

    Une interface soignée donne l'impression d'un outil fini. Le dire une
    fois dans un README ne suffit pas : c'est ici que les documents sont
    produits, donc c'est ici que ça doit être écrit.
    """
    st.warning(
        "**Prix non calibrés.** Les taux horaires et surtout les rendements "
        "(h/unité) sont des ordres de grandeur du marché belge, pas les "
        "chiffres de l'entreprise. Tant qu'ils n'ont pas été relus, les "
        "documents produits ici montrent que la mécanique tourne — "
        "ils ne sont pas prêts à partir chez un client ou une commune.",
        icon="⚠️",
    )


# ══════════════════════════════════════════════════
#  Barre latérale — paramètres de prix
# ══════════════════════════════════════════════════

with st.sidebar:
    st.title("🧱 BAG BATTER")
    st.caption(f"{ENTREPRISE['adresse']} · {ENTREPRISE['cp_ville']}\n\n"
               f"TVA {ENTREPRISE['tva']}")

    st.subheader("Coefficient de vente")
    fg = st.number_input("Frais généraux (%)", 0.0, 100.0,
                         PARAMS["fg"] * 100, 0.5)
    fc = st.number_input("Frais de chantier (%)", 0.0, 100.0,
                         PARAMS["fc"] * 100, 0.5)
    aleas = st.number_input("Aléas (%)", 0.0, 100.0,
                            PARAMS["aleas"] * 100, 0.5)
    marge = st.number_input("Marge (%)", 0.0, 100.0,
                            PARAMS["marge"] * 100, 0.5)

    params = dict(PARAMS, fg=fg / 100, fc=fc / 100,
                   aleas=aleas / 100, marge=marge / 100)
    k = coefficient_k(params)
    st.metric("Coefficient K", f"{k:.4f}", f"{(k - 1) * 100:+.1f} %",
              help="pu_vente = déboursé sec × K")

    if abs(k - coefficient_k()) > 1e-9:
        st.info("Coefficient modifié pour cette session seulement. "
                 "Pour changer la référence, voir l'onglet "
                 "« ⚙️ Paramètres ».", icon="💡")

    anomalies = {c: v for c, v in controle_coherence().items() if v}
    if anomalies:
        st.error(f"Bibliothèque incohérente : {anomalies}", icon="🛑")

    st.divider()
    st.caption(
        f"{len(OUVRAGES)} ouvrages · {len(LOTS)} lots · "
        f"{len(OUVRAGES_A_VALIDER)} rendements jamais validés"
    )



# ══════════════════════════════════════════════════
#  1. Répondre à un métré imposé
# ══════════════════════════════════════════════════

(onglet_metre, onglet_devis, onglet_biblio,
 onglet_lexique, onglet_calib, onglet_params) = st.tabs(
    ["📥 Répondre à un métré", "🧾 Devis client", "📚 Bibliothèque",
     "🔤 Lexique", "🎯 Calibration", "⚙️ Paramètres"]
)


# ────────────────────────────────────────────
# L'onglet « métré » est une FONCTION, pas un bloc `with` : il a besoin
# de sorties anticipées (pas de fichier déposé, fichier illisible).
# `st.stop()` arrêterait le script ENTIER — les trois autres onglets
# ne s'afficheraient plus. `return` ne quitte que cet onglet.
# ────────────────────────────────────────────
def _repondre_a_un_metre(params):
    st.header("Répondre à un métré imposé")
    st.markdown(
        "Dépose le fichier Excel reçu du pouvoir adjudicateur. "
        "L'outil lit les postes, propose une correspondance vers les "
        "ouvrages de la bibliothèque, et te rend **son** fichier avec la "
        "seule colonne des prix unitaires remplie."
    )
    _avertissement_calibration()

    fichier = st.file_uploader("Métré imposé (.xlsx)", type=["xlsx", "xlsm"])

    if fichier is None:
        st.info(
            "Pas encore de métré ? Le bouton ci-dessous en fabrique un "
            "fictif (49 postes, 10 lots) pour essayer l'outil.",
            icon="💡",
        )
        if st.button("Générer un métré d'entraînement"):
            with TemporaryDirectory() as tmp:
                chemin = Path(tmp) / "METRE_entrainement.xlsx"
                generer_metre(str(chemin))
                st.download_button(
                    "⬇️ Télécharger le métré d'entraînement",
                    chemin.read_bytes(),
                    file_name="METRE_CSC_2026-TP-0147_Schaerbeek.xlsx",
                    mime="application/vnd.openxmlformats-officedocument."
                          "spreadsheetml.sheet",
                )
        return

    b = _bordereau(params)

    with TemporaryDirectory() as tmp:
        chemin_metre = Path(tmp) / fichier.name
        chemin_metre.write_bytes(fichier.getvalue())

        # ── Choix des feuilles ────────────────────────
        # Un métré réel se répartit souvent en « Lot 01 », « Lot 02 »…
        # plus un « Récapitulatif » qui REPREND les mêmes codes :
        # le traiter ferait compter les postes deux fois.
        try:
            inventaire = [f for f in feuilles_avec_postes(str(chemin_metre))
                           if f["nb_postes"]]
        except Exception as err:
            st.error(f"Lecture impossible : {err}", icon="🛑")
            return

        if not inventaire:
            st.error(
                "Aucune feuille ne contient de postes reconnaissables. "
                "L'outil cherche une colonne de codes (B) et une colonne "
                "de quantités (F).",
                icon="🛑")
            return

        if len(inventaire) > 1:
            st.markdown("**Feuilles à traiter**")
            retenues = []
            for feuille in inventaire:
                marque = " · récapitulatif présumé" if feuille["recapitulatif"] else ""
                if st.checkbox(
                    f"{feuille['nom']} — {feuille['nb_postes']} postes{marque}",
                    value=not feuille["recapitulatif"],
                    key=f"feuille_{feuille['nom']}",
                ):
                    retenues.append(feuille["nom"])
            if not retenues:
                st.warning("Aucune feuille retenue.", icon="⚠️")
                return
            st.caption(
                "Un récapitulatif reprend les codes des lots : le cocher "
                "ferait compter chaque poste deux fois. Les doublons sont "
                "signalés, jamais additionnés."
            )
        else:
            retenues = [inventaire[0]["nom"]]

        # ── Correspondance des colonnes ────────────────
        # Les colonnes étaient codées en dur : chez un autre pouvoir
        # adjudicateur, l'outil ne lisait rien — ou écrivait le prix
        # dans la mauvaise colonne, ce qui est pire. La détection
        # PROPOSE, l'humain valide avant tout chiffrage.
        detectees = next(
            (f["colonnes"] for f in inventaire if f["nom"] == retenues[0]),
            None)
        colonnes = dict(detectees or {})

        incertaines = [c for c in CHAMPS_REQUIS if c not in colonnes]
        libelle_champ = {
            "code": "Code du poste", "designation": "Désignation",
            "nature": "Nature", "unite": "Unité", "quantite": "Quantité",
            "pu": "Prix unitaire (colonne à remplir)", "montant": "Montant",
        }
        resume = " · ".join(
            f"{libelle_champ[c]} = {get_column_letter(colonnes[c])}"
            for c in CHAMPS_REQUIS if c in colonnes)

        with st.expander(
            ("⚠️ Colonnes à confirmer" if incertaines
             else f"🔎 Colonnes détectées — {resume}"),
            expanded=bool(incertaines),
        ):
            st.caption(
                "Corriger ici si la détection s'est trompée. Un prix "
                "écrit dans la mauvaise colonne rendrait l'offre "
                "silencieusement fausse."
            )
            lettres = [get_column_letter(i) for i in range(1, 31)]
            gauche, droite = st.columns(2)
            for i, champ in enumerate(CHAMPS):
                zone = gauche if i % 2 == 0 else droite
                actuelle = colonnes.get(champ)
                options = ["— absente —"] + lettres
                index = (lettres.index(get_column_letter(actuelle)) + 1
                          if actuelle and actuelle <= 30 else 0)
                choix = zone.selectbox(
                    libelle_champ[champ] + (" *" if champ in CHAMPS_REQUIS
                                             else ""),
                    options, index=index, key=f"col_{champ}")
                if choix == "— absente —":
                    colonnes.pop(champ, None)
                else:
                    colonnes[champ] = lettres.index(choix) + 1

            manquants = [libelle_champ[c] for c in CHAMPS_REQUIS
                          if c not in colonnes]
            if manquants:
                st.error(
                    "Champs indispensables non renseignés : "
                    + ", ".join(manquants)
                    + ". Sans eux, impossible de savoir quoi chiffrer, où "
                      "écrire, ni de vérifier les unités.",
                    icon="🛑")

        if any(c not in colonnes for c in CHAMPS_REQUIS):
            return

        try:
            lecture = lire_metre_complet(str(chemin_metre), retenues, colonnes)
        except Exception as err:
            st.error(f"Lecture impossible : {err}", icon="🛑")
            return
        postes = lecture["postes"]

        # Une ligne non lue est un poste ABSENT de l'offre — plus
        # discret qu'un poste sans prix, et tout aussi disqualifiant.
        if lecture["anomalies"]:
            bloquantes = [a for a in lecture["anomalies"]
                           if a["genre"] not in ("quantite_formule",
                                                  "quantite_nulle")]
            texte = "\n".join(
                f"- ligne **{a['ligne']}** · `{a['code']}` — {a['detail']}"
                for a in lecture["anomalies"])
            if bloquantes:
                st.error(
                    f"**{len(bloquantes)} ligne(s) du métré n'ont pas pu "
                    f"être lues — ces postes ne seront PAS chiffrés.**\n\n"
                    + texte, icon="🛑")
            else:
                st.warning(
                    f"**{len(lecture['anomalies'])} ligne(s) à "
                    f"vérifier.**\n\n" + texte, icon="⚠️")

        if not postes:
            st.error(
                "Aucun poste lu. Le fichier doit porter, par feuille, une "
                "colonne de codes au format NN.NN et une colonne de "
                "quantités numériques.",
                icon="🛑",
            )
            return

        # ── Appariement : recalculé si le fichier change ──────────────
        # La version du lexique entre dans la signature : ajouter un
        # synonyme dans l'onglet Lexique doit refaire l'appariement,
        # sinon l'essai n'a aucun effet visible ici.
        # Le contenu, pas le nom ni la taille : deux métrés différents
        # peuvent porter le même nom et peser pareil.
        signature = (
            hashlib.sha256(fichier.getvalue()).hexdigest(),
            tuple(retenues),
            tuple(sorted(colonnes.items())),
            st.session_state.get("lexique_version", 0),
        )
        # L'appariement est refait quand la signature change — mais
        # aussi quand l'un des deux états manque. Ils étaient lus en
        # accès direct : une session dont la signature avait survécu
        # sans son appariement (interruption au milieu du calcul,
        # session restaurée à moitié après un rafraîchissement) levait
        # un KeyError et affichait un bandeau rouge, là où il suffisait
        # de recalculer.
        etat_incomplet = ("proposition" not in st.session_state
                           or "mapping" not in st.session_state)
        if st.session_state.get("signature") != signature or etat_incomplet:
            proposition = proposer_mapping(
                postes, b, mapping_connu=MAPPING, seuil=SEUIL_SUGGESTION)
            st.session_state.signature = signature
            st.session_state.proposition = proposition
            st.session_state.mapping = {
                code: infos["code_ouv"]
                for code, infos in proposition.items()
            }

        proposition = st.session_state.proposition
        mapping = st.session_state.mapping

        # ── Tableau de contrôle ────────────────────────────────
        a_revoir = [
            p for p in postes
            if mapping.get(p["code"]) is None
            or proposition[p["code"]]["origine"] != "connu"
            and proposition[p["code"]]["score"] < SEUIL_CONFIANCE
        ]
        chiffres = [p for p in postes if mapping.get(p["code"])]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Postes lus", len(postes))
        c2.metric("Appariés", len(chiffres))
        c3.metric("À revoir", len(a_revoir),
                    delta=None if not a_revoir else f"{len(a_revoir)} à confirmer",
                    delta_color="off" if not a_revoir else "inverse")
        total = sum(
            b[mapping[p["code"]]]["pu_vente"] * p["quantite"]
            for p in chiffres
        )
        c4.metric("Total estimé", _euro(total))

        # ── Postes à revoir ────────────────────────────────────
        st.subheader("Postes à revoir")
        if a_revoir:
            st.caption(
                "Suggestions issues du seul libellé. Un score élevé veut dire "
                "« regarde ici d'abord », jamais « c'est bon » : "
                "« dépose » et « pose » ne diffèrent que d'une lettre pour un "
                "algorithme, et sont opposés sur un chantier."
            )
        else:
            st.success("Chaque poste porte une correspondance.", icon="✅")

        tout_revoir = st.checkbox(
            "Revoir aussi les postes déjà appariés",
            help="Par défaut, seules postes incertains sont proposésés.",
        )

        a_afficher = postes if tout_revoir else [
            p for p in postes if p["code"] in {q["code"] for q in a_revoir}
        ]

        for poste in a_afficher:
            code = poste["code"]
            infos = proposition[code]
            unite_poste = normaliser_unite(poste["unite"])

            # Unité éliminatoire : un ouvrage dans une autre unité
            # ne peut pas chiffrer ce poste, quelle que soit la ressemblance
            # des libellés. On ne les propose donc pas.
            compatibles = [
                code_ouv for code_ouv, ligne in b.items()
                if normaliser_unite(ligne["unite_ouv"]) == unite_poste
            ]
            options = [SANS_OUVRAGE] + sorted(compatibles)

            actuel = mapping.get(code)
            index = options.index(actuel) if actuel in options else 0

            candidats = ", ".join(
                f"{c} ({s:.2f})" for c, s in infos["candidats"][:3]
            ) or "aucun candidat dans cette unité"

            marque = {"connu": "🔒", "suggere": "🟡", "aucun": "⚪"}[
                infos["origine"]
            ]

            col_gauche, col_droite = st.columns([3, 2])
            with col_gauche:
                st.markdown(
                    f"**{marque} {code} · {poste['designation']}**  \n"
                    f"`{poste['unite']}` × {poste['quantite']:g}"
                )
                st.caption(f"Candidats : {candidats}")
            with col_droite:
                choix = st.selectbox(
                    f"Ouvrage pour {code}",
                    options,
                    index=index,
                    format_func=lambda c: (
                        SANS_OUVRAGE if c == SANS_OUVRAGE
                        else _libelle_ouvrage(c, b)
                    ),
                    key=f"map_{code}",
                    label_visibility="collapsed",
                )
                mapping[code] = None if choix == SANS_OUVRAGE else choix

            if not compatibles:
                st.caption(
                    f"⚠️ Aucun ouvrage en « {poste['unite']} » : il faut le "
                    f"créer dans la bibliothèque, ce poste ne peut pas être "
                    f"chiffré."
                )
            st.divider()

        st.session_state.mapping = mapping

        # ── Production de l'offre ──────────────────────────────
        st.subheader("Produire l'offre")

        mapping_effectif = {poste: ouv for poste, ouv in mapping.items() if ouv}
        vides = [p["code"] for p in postes if not mapping.get(p["code"])]

        if vides:
            st.error(
                f"**{len(vides)} postes sans prix : offre irrégulière.** "
                f"Un seul poste vide suffit à faire rejeter l'offre "
                f"(art. 76 AR 18/04/2017). Postes concernés : "
                + ", ".join(vides),
                icon="🛑",
            )

        col_a, col_b = st.columns(2)
        with col_a:
            tva = st.radio("TVA", [0.21, 0.06],
                            format_func=lambda t: f"{t * 100:.0f} %",
                            horizontal=True,
                            help="21 % en marché public. 6 % : logement privé "
                                  "de plus de 10 ans, usage privé, consommateur final.")
        with col_b:
            st.write("")
            chiffrer = st.button("⚙️ Chiffrer et produire l'offre",
                                  type="primary", width="stretch",
                                  disabled=not mapping_effectif)

        if chiffrer:
            with TemporaryDirectory() as tmp:
                entree = Path(tmp) / fichier.name
                entree.write_bytes(fichier.getvalue())
                sortie = Path(tmp) / f"OFFRE_{Path(fichier.name).stem}.xlsx"
                rapport = remplir_metre(
                    str(entree), str(sortie), mapping=mapping,
                    params=params, tva=tva, feuilles=retenues,
                    colonnes=colonnes,
                )
                octets = sortie.read_bytes()

            st.session_state.offre = {
                "octets": octets,
                "nom": f"OFFRE_{Path(fichier.name).stem}"
                        f"_{datetime.now():%Y%m%d_%H%M}.xlsx",
                "rapport": rapport,
            }

        if st.session_state.get("offre"):
            offre = st.session_state.offre
            rapport = offre["rapport"]

            st.success(
                f"{len(rapport['chiffres'])} postes chiffrés · "
                f"{_euro(rapport['total_ht'])} HTVA · "
                f"{rapport['heures_mo']:.0f} h de main-d'œuvre",
                icon="✅",
            )

            if rapport["ecarts_unite"]:
                st.error(
                    "Écarts d'unité — ces postes ne sont **pas** chiffrés :\n"
                    + "\n".join(
                        f"- `{e['code']}` métré en « {e['unite_metre']} » mais "
                        f"`{e['code_ouv']}` est en « {e['unite_ouvrage']} »"
                        for e in rapport["ecarts_unite"]
                    ),
                    icon="🛑",
                )

            st.download_button(
                "⬇️ Télécharger l'offre à renvoyer",
                offre["octets"],
                file_name=offre["nom"],
                mime="application/vnd.openxmlformats-officedocument."
                      "spreadsheetml.sheet",
                type="primary",
            )
            st.caption(
                "C'est le fichier du pouvoir adjudicateur, avec la seule "
                "colonne des prix unitaires remplie : ses quantités et ses "
                "formules sont intactes."
            )

            # ── Contrôle des prix avant dépôt ──────────────
            st.divider()
            st.subheader("Contrôle des prix")
            st.caption(
                "Deux risques opposés : trop bas, l'offre est écartée pour "
                "prix anormalement bas (art. 36 AR 18/04/2017) ou le "
                "chantier s'exécute à perte ; trop haut, le marché est "
                "perdu. Ces contrôles regardent l'offre depuis l'intérieur "
                "de l'entreprise — les prix des concurrents, personne ne "
                "les a au moment de déposer."
            )

            lignes_controle = [
                {"code_ouv": c["code_ouv"], "qte": c["quantite"],
                 "pu_vente": c["pu"]}
                for c in rapport["chiffres"]
            ]

            rabais = st.slider(
                "Rabais commercial simulé", 0.0, 30.0, 0.0, 0.5,
                format="%.1f %%",
                help="Le seuil au-delà duquel l'offre cesse de couvrir "
                      "ses coûts est calculé juste en dessous.",
            ) / 100.0

            controle = analyser(lignes_controle, b, params=params,
                                 rabais=rabais)
            ind = controle["indicateurs"]

            m1, m2, m3 = st.columns(3)
            m1.metric(
                "Par heure travaillée", f"{ind['par_heure']:.2f} €",
                f"{ind['par_heure'] - ind['plancher']:+.2f} € / plancher",
                delta_color="normal" if ind["couvre"] else "inverse",
                help="Matériaux et matériel payés, ce que l'offre laisse "
                      "par heure de main-d'œuvre. Sous le plancher, chaque "
                      "heure travaillée coûte de l'argent.",
            )
            m2.metric("Rabais maximal",
                       f"{ind['rabais_maximal'] * 100:.1f} %",
                       help="Au-delà, l'offre ne couvre plus ses coûts.")
            m3.metric("Top 3 des postes",
                       f"{ind['concentration_top3'] * 100:.0f} %",
                       help="Part du total portée par trois postes.")

            for alerte in controle["alertes"]:
                texte = f"**{alerte['titre']}**  \n{alerte['detail']}"
                if alerte["niveau"] == "critique":
                    st.error(texte, icon="🛑")
                elif alerte["niveau"] == "attention":
                    st.warning(texte, icon="⚠️")
                else:
                    st.info(texte, icon="💡")
            if not controle["alertes"]:
                st.success("Aucun point d'attention.", icon="✅")

            # ── Dossier de justification ───────────────
            with st.expander("📄 Dossier de justification de prix"):
                st.markdown(
                    "Si le pouvoir adjudicateur conteste un prix, il doit "
                    "demander une justification écrite avant d'écarter "
                    "l'offre — et le délai de réponse est court. Ce "
                    "dossier contient, pour chaque poste visé, la "
                    "décomposition qui a **servi** à établir l'offre : "
                    "ressources, quantités, prix d'achat, déboursés, "
                    "coefficient. C'est ce qui la rend crédible."
                )
                defaut = [p["code_ouv"] for p in controle["postes"][:3]]
                a_justifier = st.multiselect(
                    "Postes à justifier",
                    sorted({p["code_ouv"] for p in controle["postes"]}),
                    default=defaut,
                    format_func=lambda c: _libelle_ouvrage(c, b),
                )
                if a_justifier:
                    with TemporaryDirectory() as tmp:
                        cible = Path(tmp) / "justification.xlsx"
                        exporter_justification(
                            a_justifier, str(cible),
                            marche={"reference": Path(fichier.name).stem},
                            params=params,
                        )
                        octets_just = cible.read_bytes()
                    st.download_button(
                        "⬇️ Télécharger le dossier",
                        octets_just,
                        file_name=f"JUSTIFICATION_{datetime.now():%Y%m%d_%H%M}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument."
                              "spreadsheetml.sheet",
                    )
                    st.caption(
                        "La lettre d'accompagnement est le premier onglet : "
                        "à relire et à signer avant envoi."
                    )

        # ── Réutiliser la correspondance ─────────────────────────────
        with st.expander("♻️ Réutiliser cette correspondance"):
            st.markdown(
                "Une commune réutilise ses propres codes d'un marché à "
                "l'autre. Garde ce fichier : au prochain métré de la même "
                "commune, la correspondance sera déjà faite."
            )
            st.download_button(
                "⬇️ Télécharger la correspondance (.json)",
                json.dumps(mapping_effectif, indent=2, ensure_ascii=False),
                file_name=f"MAPPING_{Path(fichier.name).stem}.json",
                mime="application/json",
            )
            repris = st.file_uploader("Reprendre une correspondance (.json)",
                                        type=["json"], key="up_map")
            # BOUCLE DE RERUN À NE PAS RÉINTRODUIRE.
            # `repris` reste non-None à CHAQUE réexécution du script tant
            # que le fichier est déposé : un st.rerun() inconditionnel
            # ici relance le script sans fin. L'app ne répond plus, et
            # recharger la page n'y change rien puisque le fichier est
            # toujours là — il faut redémarrer le serveur.
            # On ne traite donc le dépôt QU'UNE FOIS, repéré par
            # l'empreinte de son contenu.
            depot_deja_lu = st.session_state.get("mapping_importe")
            empreinte = (hashlib.sha256(repris.getvalue()).hexdigest()
                          if repris is not None else None)
            if repris is not None and empreinte != depot_deja_lu:
                try:
                    charge = json.loads(repris.getvalue().decode("utf-8"))
                    inconnus = set(charge.values()) - set(b) - {None}
                    if inconnus:
                        st.warning(
                            "Codes d'ouvrage inconnus, ignorés : "
                            + ", ".join(sorted(inconnus)),
                            icon="⚠️",
                        )
                    st.session_state.mapping = {
                        poste: ouv for poste, ouv in charge.items()
                        if ouv in b or ouv is None
                    }
                    st.session_state.mapping_importe = empreinte
                    st.rerun()
                except Exception as err:
                    # Marqué lu malgré l'échec : sans ça, un fichier
                    # illisible reposerait la question à chaque rerun.
                    st.session_state.mapping_importe = empreinte
                    st.error(f"Fichier illisible : {err}", icon="🛑")


with onglet_metre:
    _repondre_a_un_metre(params)


# ══════════════════════════════════════════════════
#  2. Devis client
# ══════════════════════════════════════════════════

with onglet_devis:
    st.header("Devis client")
    _avertissement_calibration()

    b = _bordereau(params)
    col_g, col_d = st.columns([2, 1])

    with col_g:
        objet = st.text_input("Objet", "Rénovation de la façade arrière")
        reference = st.text_input("Référence", f"{date.today():%Y}-042")
        chantier = st.text_input("Chantier",
                                   "Avenue Ernest Renan 62, 1030 Schaerbeek")
        client = st.text_area(
            "Client",
            "M. et Mme Dupont\nRue de l'Église 12\n1030 Schaerbeek",
            height=90,
        )
    with col_d:
        tva_devis = st.radio("TVA", [0.06, 0.21],
                              format_func=lambda t: f"{t * 100:.0f} %",
                              key="tva_devis")
        st.caption(
            "6 % : logement privé de plus de dix ans, usage principalement "
            "privé, facturation au consommateur final. Dans le doute, 21 %."
        )

    st.subheader("Postes")
    if "lignes_devis" not in st.session_state:
        st.session_state.lignes_devis = [
            {"code_ouv": "40.20", "qte": 22.0},
            {"code_ouv": "40.30", "qte": 22.0},
        ]

    edite = st.data_editor(
        st.session_state.lignes_devis,
        num_rows="dynamic",
        width="stretch",
        column_config={
            "code_ouv": st.column_config.SelectboxColumn(
                "Ouvrage", options=sorted(b), required=True, width="large"),
            "qte": st.column_config.NumberColumn(
                "Quantité", min_value=0.0, step=0.5, format="%.2f"),
        },
        key="editeur_devis",
    )
    st.session_state.lignes_devis = edite

    lignes = [
        (ligne["code_ouv"], float(ligne["qte"]))
        for ligne in edite
        if ligne.get("code_ouv") and ligne.get("qte")
    ]

    if lignes:
        d = devis(objet or "Devis", lignes, tva=tva_devis, params=params)

        c1, c2, c3 = st.columns(3)
        c1.metric("Total HTVA", _euro(d["total_ht"]))
        c2.metric(f"TVA {tva_devis * 100:.0f} %", _euro(d["montant_tva"]))
        c3.metric("Total TVAC", _euro(d["total_ttc"]))
        st.caption(f"Main-d'œuvre : {d['heures_mo']:.1f} h "
                    f"({d['jours_homme']:.1f} jours-homme)")

        tableau = [
            {
                "Code": poste["code_ouv"],
                "Désignation": poste["libelle_ouv"],
                "Un.": poste["unite_ouv"],
                "Qté": poste["qte"],
                "PU HTVA": poste["pu_vente"],
                "Montant HTVA": poste["montant"],
            }
            for poste in d["lignes"]
        ]
        st.dataframe(tableau, width="stretch", hide_index=True)

        with TemporaryDirectory() as tmp:
            cible = Path(tmp) / f"DEVIS_{reference}.xlsx"
            exporter_devis(d, str(cible), client=client, chantier=chantier,
                            reference=reference)
            octets = cible.read_bytes()

        st.download_button(
            "⬇️ Télécharger le devis", octets,
            file_name=f"DEVIS_{reference}_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument."
                  "spreadsheetml.sheet",
            type="primary",
        )
    else:
        st.info("Ajoute au moins un poste.", icon="💡")

# ══════════════════════════════════════════════════
#  3. Bibliothèque
# ══════════════════════════════════════════════════

with onglet_biblio:
    st.header("Bibliothèque de prix")
    _avertissement_calibration()

    b = _bordereau(params)
    a_valider = set(OUVRAGES_A_VALIDER)

    lignes = [
        {
            "Code": code,
            "Lot": f"{ligne['lot']} — {LOTS.get(ligne['lot'], '')}",
            "Désignation": ligne["libelle_ouv"],
            "Un.": ligne["unite_ouv"],
            "h/unité": ligne["heures_mo"],
            "Déboursé": ligne["debourse_sec"],
            "PU vente": ligne["pu_vente"],
            "À valider": "⚠️" if code in a_valider else "",
        }
        for code, ligne in sorted(b.items())
    ]
    st.dataframe(lignes, width="stretch", hide_index=True,
                  height=420)

    st.subheader("Fiche de justification de prix")
    st.caption(
        "La pièce à produire si un pouvoir adjudicateur conteste un prix "
        "jugé anormal (art. 36 AR 18/04/2017)."
    )
    code_fiche = st.selectbox("Ouvrage", sorted(b),
                                 format_func=lambda c: _libelle_ouvrage(c, b))
    st.code(fiche_prix(code_fiche), language="text")

    # ══════════════════════════════════════
    #  Atelier de correction — la séance de calibration
    # ══════════════════════════════════════
    st.divider()
    st.subheader("✏️ Corriger les prix et les rendements")
    st.markdown(
        "Les deux tables que le chef d'entreprise est seul à pouvoir "
        "corriger : ce que coûte une heure, et combien d'heures prend un "
        "mètre carré. **L'effet sur la calibration s'affiche en direct**, "
        "avant tout enregistrement."
    )

    if "tables_editees" not in st.session_state:
        st.session_state.tables_editees = copy.deepcopy(tables_courantes())
    tables = st.session_state.tables_editees
    origine = tables_courantes()

    onglet_res, onglet_rend = st.tabs(
        ["Taux horaires et prix d'achat", "Rendements (h/unité)"])

    with onglet_res:
        st.caption(
            "Les taux MO doivent être le **coût entreprise complet** — "
            "salaire, ONSS patronale, pécule, jours fériés, assurance loi, "
            "déplacements, EPI. Surtout pas le brut."
        )
        edite_res = st.data_editor(
            [{"Code": r["code_res"], "Désignation": r["libelle_res"],
               "Type": r["type_res"], "Unité": r["unite_res"],
               "Prix": r["pu_res"], "Note": r.get("note", "")}
              for r in tables["ressources"]],
            hide_index=True, width="stretch", height=340,
            disabled=["Code", "Désignation", "Type", "Unité", "Note"],
            column_config={
                "Prix": st.column_config.NumberColumn(
                    "Prix d'achat / taux", min_value=0.0, step=0.5,
                    format="%.2f €", required=True),
                "Note": st.column_config.TextColumn(width="large"),
            },
            key="editeur_ressources")
        prix_saisis = {ligne["Code"]: ligne["Prix"] for ligne in edite_res}
        for res in tables["ressources"]:
            if res["code_res"] in prix_saisis:
                res["pu_res"] = float(prix_saisis[res["code_res"]])
        tables["ressources_par_code"] = {r["code_res"]: r
                                          for r in tables["ressources"]}

    with onglet_rend:
        st.caption(
            "Le rendement est la **seule donnée non achetable** : elle "
            "vient de l'expérience du chantier, et pèse la majeure partie "
            "du déboursé. C'est ici que la calibration se joue."
        )
        a_valider = set(OUVRAGES_A_VALIDER)
        lignes_mo = [
            (i, c) for i, c in enumerate(tables["composition"])
            if tables["ressources_par_code"][c["code_res"]]["type_res"] == "MO"
        ]
        edite_rend = st.data_editor(
            [{"Ouvrage": c["code_ouv"],
               "Désignation": origine["ouvrages_par_code"][c["code_ouv"]]["libelle_ouv"],
               "Un.": origine["ouvrages_par_code"][c["code_ouv"]]["unite_ouv"],
               "Qui": c["code_res"],
               "h/unité": c["qte_res"],
               "": "⚠️" if c["code_ouv"] in a_valider else "",
               "Note": c.get("note", "")}
              for _, c in lignes_mo],
            hide_index=True, width="stretch", height=340,
            disabled=["Ouvrage", "Désignation", "Un.", "Qui", "", "Note"],
            column_config={
                "h/unité": st.column_config.NumberColumn(
                    min_value=0.0, step=0.05, format="%.3f", required=True),
                "Note": st.column_config.TextColumn(width="medium"),
            },
            key="editeur_rendements")
        for (indice, _), ligne in zip(lignes_mo, edite_rend):
            tables["composition"][indice]["qte_res"] = float(ligne["h/unité"])
        st.caption("⚠️ = rendement jamais confronté à un chantier réel.")

    # ── Effet en direct ────────────────────────
    modifs_res = [r for r in tables["ressources"]
                   if r["pu_res"] != origine["ressources_par_code"][
                       r["code_res"]]["pu_res"]]
    modifs_compo = [
        (a, b) for a, b in zip(tables["composition"], origine["composition"])
        if a["qte_res"] != b["qte_res"]]

    avant = calibration(params)
    apres = calibration(params, tables=tables)

    m1, m2, m3 = st.columns(3)
    m1.metric("Valeurs corrigées", len(modifs_res) + len(modifs_compo))
    m2.metric("Écart moyen absolu",
               f"{apres['ecart_moyen_absolu'] * 100:.1f} %",
               f"{(apres['ecart_moyen_absolu'] - avant['ecart_moyen_absolu']) * 100:+.1f} pt"
               if modifs_res or modifs_compo else None,
               delta_color="inverse")
    m3.metric("Devis hors cible",
               sum(1 for r in apres["lignes"] if abs(r["ecart"]) > 0.15),
               help="Cible : moins de 15 % d'écart sur CHAQUE ligne.")

    if modifs_res or modifs_compo:
        st.dataframe(
            [{"Devis": r["devis"], "Objet": r["objet"],
               "Forfait vendu": r["forfait"], "Calculé": r["calcule"],
               "Écart": r["ecart"] * 100,
               "Écart avant": a["ecart"] * 100}
              for r, a in zip(apres["lignes"], avant["lignes"])],
            hide_index=True, width="stretch",
            column_config={
                "Forfait vendu": st.column_config.NumberColumn(format="%.0f €"),
                "Calculé": st.column_config.NumberColumn(format="%.0f €"),
                "Écart": st.column_config.NumberColumn(format="%+.1f %%"),
                "Écart avant": st.column_config.NumberColumn(format="%+.1f %%"),
            })

        # ── Enregistrer ────────────────────────
        a_ecrire = {}
        if modifs_res:
            a_ecrire["ressources"] = tables["ressources"]
        if modifs_compo:
            a_ecrire["composition"] = tables["composition"]

        try:
            github = dict(st.secrets.get("github", {}))
        except Exception:
            github = {}
        depot, jeton = github.get("depot"), github.get("token")

        col_ok, col_raz = st.columns(2)
        with col_ok:
            if depot and jeton:
                if st.button(
                    f"📤 Enregistrer {len(a_ecrire)} table(s) dans le dépôt",
                    type="primary", width="stretch",
                ):
                    with st.spinner("Écriture dans le dépôt…"):
                        try:
                            for nom, contenu in a_ecrire.items():
                                commiter_table(
                                    nom,
                                    json.dumps(contenu, ensure_ascii=False,
                                                indent=2) + "\n",
                                    depot, jeton,
                                    branche=github.get("branche", "main"))
                        except ErreurDepot as err:
                            st.error(str(err), icon="🛑")
                        else:
                            st.success(
                                "Enregistré. L'app se redéploie dans une à "
                                "deux minutes ; ces valeurs deviendront "
                                "celles de la bibliothèque.", icon="✅")
            else:
                for nom, contenu in a_ecrire.items():
                    st.download_button(
                        f"⬇️ {nom}.json",
                        json.dumps(contenu, ensure_ascii=False, indent=2) + "\n",
                        file_name=f"{nom}.json", mime="application/json",
                        width="stretch", key=f"dl_{nom}")
        with col_raz:
            if st.button("↩️ Repartir des valeurs enregistrées",
                          width="stretch"):
                del st.session_state.tables_editees
                st.rerun()

        st.caption(
            "Tant que ce n'est pas enregistré, ces corrections ne valent "
            "que pour cet aperçu : les devis et les offres continuent "
            "d'utiliser les valeurs de la bibliothèque."
        )
    else:
        st.caption("Aucune valeur corrigée pour l'instant.")


# ══════════════════════════════════════════════════
#  4. Lexique métier
# ══════════════════════════════════════════════════

with onglet_lexique:
    st.header("Lexique métier")
    st.markdown(
        "Un pouvoir adjudicateur n'écrit pas « faïence » : il écrit "
        "« carrelage mural ». Ce lexique traduit son vocabulaire vers "
        "celui de la bibliothèque, **avant** toute comparaison de "
        "libellés. C'est lui qui fait la différence entre un poste "
        "apparié et un poste laissé sans prix."
    )

    b = _bordereau(params)

    # ── Banc d'essai ──────────────────────────
    st.subheader("Banc d'essai")
    st.caption(
        "Colle ici un libellé qui n'a pas été apparié, et regarde ce que "
        "l'outil en retient. C'est le cycle qui permet de régler le "
        "lexique sans écrire de Python."
    )

    col_lib, col_unite = st.columns([4, 1])
    with col_lib:
        essai = st.text_input(
            "Libellé du poste",
            "Sablage des maçonneries de façade",
            label_visibility="collapsed",
            placeholder="Libellé tel qu'il figure dans le métré",
        )
    with col_unite:
        # Le m2 par défaut, et non le premier par ordre alphabétique :
        # 29 des 49 ouvrages y sont, et surtout le libellé prérempli
        # est un travail de façade. Avec « FF » en tête, le banc
        # d'essai s'ouvrait sur cinq forfaits sans rapport avec
        # l'exemple — de quoi croire l'outil cassé au premier regard.
        unites = sorted({ligne["unite_ouv"] for ligne in b.values()})
        unite_essai = st.selectbox(
            "Unité",
            unites,
            index=unites.index("m2") if "m2" in unites else 0,
            label_visibility="collapsed",
        )

    if essai.strip():
        mots = normaliser(essai)
        avec_operation = normaliser(essai, garder_operation=True)

        st.markdown(
            "**Ce que l'outil retient :** "
            + (" · ".join(f"`{m}`" for m in mots) or "_rien_")
            + ("  \n**Opération :** dépose"
               if est_demolition(avec_operation)
               else "  \n**Opération :** mise en œuvre")
        )
        if not mots:
            st.warning(
                "Aucun mot significatif : tous ont été écartés comme mots "
                "vides de métré. Aucun appariement n'est possible.",
                icon="⚠️",
            )

        candidats = suggerer(
            {"designation": essai, "unite": unite_essai}, b, limite=5
        )
        if not candidats:
            st.error(
                f"Aucun ouvrage en « {unite_essai} » : ce n'est pas un "
                f"problème de vocabulaire, il manque l'ouvrage dans la "
                f"bibliothèque.",
                icon="🛑",
            )
        else:
            st.dataframe(
                [
                    {
                        "Score": s,
                        "": "✅" if s >= SEUIL_CONFIANCE
                             else ("🟡" if s >= SEUIL_SUGGESTION else "⚪"),
                        "Code": c,
                        "Ouvrage": b[c]["libelle_ouv"],
                        "PU": b[c]["pu_vente"],
                    }
                    for c, s in candidats
                ],
                hide_index=True,
                width="stretch",
                column_config={
                    "Score": st.column_config.ProgressColumn(
                        "Score", min_value=0.0, max_value=1.0, format="%.2f"),
                    "PU": st.column_config.NumberColumn(format="%.2f €"),
                },
            )
            st.caption(
                f"🟡 au-dessus de {SEUIL_SUGGESTION:.2f} : proposé comme "
                f"suggestion · ✅ au-dessus de {SEUIL_CONFIANCE:.2f} : "
                f"présenté comme solide — ce qui ne veut pas dire juste. "
                f"⚪ en dessous : l'outil se tait."
            )

    # ── Enrichir ───────────────────────────────
    st.subheader("Ajouter un terme")
    st.caption(
        "L'ajout prend effet immédiatement — pour **tous** les "
        "utilisateurs de l'app, pas seulement toi. Relance le banc "
        "d'essai ci-dessus pour en voir l'effet ; l'appariement de "
        "l'onglet « métré » est refait lui aussi."
    )

    col_var, col_canon, col_bouton = st.columns([2, 2, 1])
    with col_var:
        variante = st.text_input(
            "Terme du cahier des charges",
            placeholder="sablage  ·  carrelage mural",
        )
    with col_canon:
        canonique = st.text_input(
            "Terme de la bibliothèque",
            placeholder="nettoyage  ·  faience",
        )
    with col_bouton:
        st.write("")
        ajouter = st.button(
            "➕ Ajouter", width="stretch",
            disabled=not (variante.strip() and canonique.strip()),
        )

    if ajouter:
        # Une expression multi-mots se traduit AVANT la découpe en mots,
        # un mot seul APRÈS : ce ne sont pas les mêmes tables.
        try:
            if " " in variante.strip():
                ajouter_expression(variante, canonique)
            else:
                ajouter_synonyme(variante, canonique)
        except ValueError as err:
            st.error(str(err), icon="🛑")
        else:
            st.session_state.lexique_version = (
                st.session_state.get("lexique_version", 0) + 1
            )
            st.rerun()

    message_commit = st.session_state.pop("lexique_commit_texte", None)
    if message_commit:
        st.success(message_commit, icon="✅")

    # ── Ce qui a été ajouté à chaud ────────────────
    ajouts = dict(SURCOUCHE["expressions"], **SURCOUCHE["synonymes"])
    if ajouts:
        st.warning(
            f"**{len(ajouts)} terme(s) ajouté(s) à chaud.** Ils valent "
            "pour **l'app entière et tous ses utilisateurs**, pas pour "
            "toi seul — et **ils ne survivront pas au redémarrage**, "
            "que Streamlit Cloud déclenche tout seul après quelques "
            "heures d'inactivité. Pour les garder, colle le bloc "
            "ci-dessous dans `chiffrage/lexique.py` et commite.",
            icon="⚠️",
        )
        st.code(surcouche_en_python(), language="python")
        # ── Rendre permanent, si un jeton est configuré ─────────
        # `st.secrets` LÈVE quand aucun fichier de secrets n'existe —
        # ce n'est pas un dict vide. C'est le cas normal en local, et
        # un plantage y serait absurde : sans jeton, on n'affiche
        # simplement pas le bouton.
        try:
            github = dict(st.secrets.get("github", {}))
        except Exception:
            github = {}
        depot, jeton = github.get("depot"), github.get("token")

        if depot and jeton:
            st.markdown(
                f"**Commiter dans `{depot}`** — les termes sont écrits "
                f"dans `chiffrage/lexique_local.json`, **du JSON et non "
                f"du code**. L'app se redéploie seule ensuite, et ils "
                f"deviennent définitifs."
            )
            if st.button("📤 Commiter ces termes", type="primary",
                          width="stretch"):
                with st.spinner("Écriture dans le dépôt…"):
                    try:
                        fusion, url = commiter_lexique(
                            None, depot, jeton,
                            branche=github.get("branche", "main"),
                        )
                    except ErreurDepot as err:
                        st.error(str(err), icon="🛑")
                    else:
                        # Ce qui vient d'être commité devient la couche
                        # locale : les ajouts à chaud n'ont plus lieu
                        # d'être, ils sont dans le dépôt.
                        adopter_local(fusion)
                        st.session_state.lexique_version = (
                            st.session_state.get("lexique_version", 0) + 1
                        )
                        # Le message doit survivre au rerun qui suit,
                        # sinon personne ne le voit jamais.
                        st.session_state.lexique_commit_texte = (
                            f"Commité. [Voir le commit]({url}) — l'app se "
                            f"redéploie dans une à deux minutes.")
                        st.rerun()
        else:
            st.caption(
                "Aucun jeton GitHub configuré : les termes ne peuvent "
                "pas être commités depuis ici. Voir README, section "
                "« Rendre les termes permanents »."
            )

        col_dl, col_raz = st.columns(2)
        with col_dl:
            st.download_button(
                "⬇️ Télécharger le bloc",
                surcouche_en_python(),
                file_name="lexique_ajouts.py",
                mime="text/x-python",
                width="stretch",
            )
        with col_raz:
            if st.button("🗑️ Oublier ces ajouts", width="stretch",
                            help="Pour tout le monde, comme l'ajout."):
                vider_surcouche()
                st.session_state.lexique_version = (
                    st.session_state.get("lexique_version", 0) + 1
                )
                st.rerun()

    # ── Le lexique du dépôt ────────────────────
    st.subheader("Le lexique du dépôt")
    filtre = st.text_input("Filtrer", placeholder="faience, enduit…",
                           label_visibility="collapsed")

    entrees = (
        [("expression", k, v) for k, v in EXPRESSIONS.items()]
        + [("mot", k, v) for k, v in SYNONYMES.items() if v]
        + [("appris", k, v) for k, v in LOCAL["expressions"].items()]
        + [("appris", k, v) for k, v in LOCAL["synonymes"].items()]
    )
    if filtre.strip():
        motif = filtre.strip().lower()
        entrees = [e for e in entrees if motif in f"{e[1]} {e[2]}"]

    st.dataframe(
        [{"Type": t, "Terme du CSC": k, "→ Bibliothèque": v}
         for t, k, v in sorted(entrees, key=lambda e: (e[0], e[1]))],
        hide_index=True, width="stretch", height=260,
    )

    with st.expander("Marqueurs de démolition"):
        st.markdown(
            "Ces mots ne sont **pas** comparés comme les autres : ils sont "
            "retirés du libellé et traités comme une dimension à part, "
            "comme l'unité. Sans ça, « dépose du carrelage mural » "
            "s'appariait à « dépose de plafond » — deux fois le mot "
            "« dépose », et rien d'autre en commun. Une dépose comparée à "
            "une pose est fortement pénalisée, mais **pas éliminée** : "
            "l'unité est déclarée dans le métré, l'opération n'est "
            "qu'inférée de mots."
        )
        st.write(" · ".join(f"`{m}`" for m in sorted(DEMOLITION)))


# ══════════════════════════════════════════════════
#  5. Calibration
# ══════════════════════════════════════════════════

with onglet_calib:
    st.header("Calibration sur les devis historiques")
    st.markdown(
        "Les six devis forfaitaires réellement vendus, re-chiffrés avec la "
        "bibliothèque. **C'est ici que se juge la qualité des prix.**"
    )

    cal = calibration(params)

    lignes = [
        {
            "Devis": r["devis"],
            "Objet": r["objet"],
            "Forfait vendu": r["forfait"],
            "Calculé": r["calcule"],
            # ×100 EXPLICITE : contrairement au format « % » d'Excel,
            # st.column_config.NumberColumn ne multiplie pas. Passer la
            # fraction brute affichait « -0,1 % » pour un écart de
            # -11,5 % — trois devis hors cible avaient l'air parfaits.
            "Écart": r["ecart"] * 100,
            "h MO": r["heures_mo"],
            "€/h vendu": r["prix_horaire_implicite"],
        }
        for r in cal["lignes"]
    ]
    st.dataframe(
        lignes,
        width="stretch",
        hide_index=True,
        column_config={
            "Forfait vendu": st.column_config.NumberColumn(format="%.0f €"),
            "Calculé": st.column_config.NumberColumn(format="%.0f €"),
            "Écart": st.column_config.NumberColumn(format="%+.1f %%"),
            "h MO": st.column_config.NumberColumn(format="%.1f"),
            "€/h vendu": st.column_config.NumberColumn(format="%.0f €"),
        },
    )

    moyen = cal["ecart_moyen_absolu"]
    hors_cible = [r for r in cal["lignes"] if abs(r["ecart"]) > 0.15]
    st.metric("Écart moyen absolu", f"{moyen * 100:.1f} %",
              help="Cible : moins de 15 % sur CHAQUE ligne, pas en moyenne")

    if hors_cible:
        st.warning(
            f"**{len(hors_cible)} devis au-delà de 15 % d'écart** : "
            + ", ".join(r["devis"] for r in hors_cible)
            + ".\n\nDeux hypothèses, non tranchées faute de relevés :\n"
            "1. les quantités estimées sont trop élevées — elles viennent "
            "des descriptifs des devis PDF, **pas de relevés** ;\n"
            "2. ces chantiers ont été vendus sous leur coût analytique.",
            icon="⚠️",
        )

    st.subheader("Les 13 rendements jamais validés")
    st.markdown(
        "Ces ouvrages ont été créés pour couvrir des postes qui restaient "
        "sans prix — ce qui rendait toute offre irrégulière. Aucun n'a "
        "jamais été confronté à un chantier réel : **ce sont les premiers "
        "à valider.**"
    )
    st.dataframe(
        [
            {
                "Code": c,
                "Désignation": b[c]["libelle_ouv"],
                "Un.": b[c]["unite_ouv"],
                "h/unité": b[c]["heures_mo"],
                "PU vente": b[c]["pu_vente"],
            }
            for c in OUVRAGES_A_VALIDER
        ],
        width="stretch",
        hide_index=True,
    )


# ══════════════════════════════════════════════════
#  6. Paramètres
# ══════════════════════════════════════════════════

with onglet_params:
    st.header("Paramètres")
    st.markdown(
        "L'identité de l'entreprise et les coefficients de vente. Ce sont "
        "des valeurs **d'entreprise**, pas des constantes techniques : "
        "elles se règlent ici, sans toucher au code."
    )

    st.subheader("Identité")
    st.caption(
        "Figure en en-tête de chaque devis et de chaque courrier de "
        "justification de prix."
    )
    col_g, col_d = st.columns(2)
    saisie_entreprise = {}
    with col_g:
        saisie_entreprise["nom"] = st.text_input(
            "Raison sociale", ENTREPRISE["nom"])
        saisie_entreprise["adresse"] = st.text_input(
            "Adresse", ENTREPRISE["adresse"])
        saisie_entreprise["cp_ville"] = st.text_input(
            "Code postal et localité", ENTREPRISE["cp_ville"])
    with col_d:
        saisie_entreprise["pays"] = st.text_input("Pays", ENTREPRISE["pays"])
        saisie_entreprise["tva"] = st.text_input(
            "Numéro de TVA", ENTREPRISE["tva"])
    saisie_entreprise["activite"] = st.text_area(
        "Activité", ENTREPRISE["activite"], height=80)

    st.subheader("Coefficients de vente")
    st.caption(
        "`pu_vente = déboursé sec × K`, avec "
        "`K = (1+FG)(1+FC)(1+aléas)(1+marge)`. Les valeurs enregistrées "
        "ici deviennent celles de départ ; la barre latérale sert à "
        "simuler autre chose le temps d'une session."
    )
    c1, c2, c3, c4 = st.columns(4)
    saisie_params = dict(PARAMS)
    saisie_params["fg"] = c1.number_input(
        "Frais généraux (%)", 0.0, 100.0, PARAMS["fg"] * 100, 0.5,
        key="p_fg") / 100
    saisie_params["fc"] = c2.number_input(
        "Frais de chantier (%)", 0.0, 100.0, PARAMS["fc"] * 100, 0.5,
        key="p_fc") / 100
    saisie_params["aleas"] = c3.number_input(
        "Aléas (%)", 0.0, 100.0, PARAMS["aleas"] * 100, 0.5,
        key="p_aleas") / 100
    saisie_params["marge"] = c4.number_input(
        "Marge (%)", 0.0, 100.0, PARAMS["marge"] * 100, 0.5,
        key="p_marge") / 100

    t1, t2, t3 = st.columns(3)
    saisie_params["tva"] = t1.number_input(
        "TVA logement privé (%)", 0.0, 100.0, PARAMS["tva"] * 100, 1.0,
        key="p_tva", help="6 % : logement de plus de dix ans, usage "
                           "principalement privé, consommateur final.") / 100
    saisie_params["tva_marche_public"] = t2.number_input(
        "TVA marché public (%)", 0.0, 100.0,
        PARAMS["tva_marche_public"] * 100, 1.0, key="p_tva_mp") / 100
    t3.metric("Coefficient K", f"{coefficient_k(saisie_params):.4f}",
               f"{(coefficient_k(saisie_params) - coefficient_k(PARAMS)):+.4f}"
               if abs(coefficient_k(saisie_params)
                       - coefficient_k(PARAMS)) > 1e-9 else None)

    # ── Enregistrer ────────────────────────────
    try:
        contenu = serialiser(saisie_entreprise, saisie_params)
    except ValueError as err:
        st.error(str(err), icon="🛑")
        contenu = None

    modifie = contenu is not None and contenu != serialiser(ENTREPRISE, PARAMS)

    if not modifie:
        st.caption("Aucune modification par rapport aux valeurs enregistrées.")
    else:
        st.info(
            "Modifications non enregistrées. Elles ne s'appliqueront "
            "qu'une fois commitées — comme pour le lexique, rien ne "
            "survit au redémarrage de l'app.",
            icon="💡",
        )

    try:
        github = dict(st.secrets.get("github", {}))
    except Exception:
        github = {}
    depot, jeton = github.get("depot"), github.get("token")

    if modifie and depot and jeton:
        if st.button("📤 Enregistrer dans le dépôt", type="primary",
                      width="stretch"):
            with st.spinner("Écriture dans le dépôt…"):
                try:
                    url = commiter_parametres(
                        contenu, depot, jeton,
                        branche=github.get("branche", "main"))
                except ErreurDepot as err:
                    st.error(str(err), icon="🛑")
                else:
                    st.success(
                        f"Enregistré. [Voir le commit]({url}) — l'app se "
                        f"redéploie dans une à deux minutes, et ces "
                        f"valeurs deviendront celles de départ.",
                        icon="✅")
    elif modifie:
        st.caption(
            "Aucun jeton GitHub configuré : impossible d'enregistrer "
            "depuis ici. Le fichier à créer est "
            "`chiffrage/parametres_local.json` — son contenu exact est "
            "ci-dessous."
        )

    if contenu:
        with st.expander("Contenu de `chiffrage/parametres_local.json`"):
            st.code(contenu, language="json")
            st.download_button(
                "⬇️ Télécharger", contenu,
                file_name="parametres_local.json",
                mime="application/json", width="stretch")
