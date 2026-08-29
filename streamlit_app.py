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
from chiffrage.depot_github import ErreurDepot, commiter_lexique
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
from chiffrage.metre_io import (
    lire_metre,
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
)
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


@st.cache_data
def _bordereau():
    """Le bordereau ne dépend que du code : calculé une fois par session."""
    return calcul_bordereau()


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
        st.info("Coefficient modifié : il ne vaut que pour cette session, "
                 "le dépôt n'est pas touché.", icon="💡")

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
 onglet_lexique, onglet_calib) = st.tabs(
    ["📥 Répondre à un métré", "🧾 Devis client",
     "📚 Bibliothèque", "🔤 Lexique", "🎯 Calibration"]
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

    b = _bordereau()

    with TemporaryDirectory() as tmp:
        chemin_metre = Path(tmp) / fichier.name
        chemin_metre.write_bytes(fichier.getvalue())

        try:
            postes = lire_metre(str(chemin_metre))
        except Exception as err:
            st.error(f"Lecture impossible : {err}", icon="🛑")
            return

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
        signature = (fichier.name, len(fichier.getvalue()), len(postes),
                      st.session_state.get("lexique_version", 0))
        if st.session_state.get("signature") != signature:
            st.session_state.signature = signature
            st.session_state.mapping = {
                code: infos["code_ouv"]
                for code, infos in proposer_mapping(
                    postes, b, mapping_connu=MAPPING,
                    seuil=SEUIL_SUGGESTION
                ).items()
            }
            st.session_state.proposition = proposer_mapping(
                postes, b, mapping_connu=MAPPING, seuil=SEUIL_SUGGESTION
            )

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
                    params=params, tva=tva,
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
            if repris is not None:
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
                    st.rerun()
                except Exception as err:
                    st.error(f"Fichier illisible : {err}", icon="🛑")


with onglet_metre:
    _repondre_a_un_metre(params)


# ══════════════════════════════════════════════════
#  2. Devis client
# ══════════════════════════════════════════════════

with onglet_devis:
    st.header("Devis client")
    _avertissement_calibration()

    b = _bordereau()
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

    b = _bordereau()
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

    b = _bordereau()

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
        unite_essai = st.selectbox(
            "Unité",
            sorted({ligne["unite_ouv"] for ligne in b.values()}),
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
                        st.success(
                            f"Commité. [Voir le commit]({url}) — l'app se "
                            f"redéploie dans une à deux minutes.",
                            icon="✅",
                        )
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
            "Écart": r["ecart"],
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
