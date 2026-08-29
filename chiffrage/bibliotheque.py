"""
╔══════════════════════════════════════════════════════════════════════════╗
║  BAG BATTER SRL — BIBLIOTHÈQUE DE PRIX UNITAIRES                         ║
║  Les trois tables sources : RESSOURCES · OUVRAGES · COMPOSITION          ║
╚══════════════════════════════════════════════════════════════════════════╝

⚠️  ÉTAT DES DONNÉES — À LIRE AVANT UTILISATION COMMERCIALE

Le notebook Colab d'origine a été perdu ; ce fichier en est une
RECONSTRUCTION à partir du seul document de reprise (CLAUDE.md). La STRUCTURE
est fidèle (mêmes tables, mêmes clés, même formule de prix, même coefficient
K = 1,3324, mêmes 36 ouvrages sur 8 lots, mêmes 13 postes non couverts) mais
les VALEURS NUMÉRIQUES — prix unitaires des ressources et surtout rendements
main-d'œuvre en h/unité — ont dû être re-saisies à partir d'ordres de grandeur
du marché belge 2026.

Ce sont donc des HYPOTHÈSES DE DÉPART, pas les chiffres calibrés du client.
Tant que le chef d'entreprise n'a pas relu les colonnes `pu_res` et les
rendements (`qte_res` des lignes MO), aucune offre ne doit partir sur cette
base. Le point d'entrée pour cette relecture est `moteur.calibration()`.

──────────────────────────────────────────────────────────────────────────
MODÈLE
──────────────────────────────────────────────────────────────────────────

    RESSOURCES   (code_res, libelle_res, type_res, unite_res, pu_res)
                 type_res : MO (main d'œuvre) | MAT (matériaux) | EQP (matériel)
         |
    COMPOSITION  (code_ouv, code_res, qte_res)
                 qte_res sur les lignes MO = RENDEMENT en h/unité d'ouvrage.
                 Seule donnée non achetable : elle vient de l'expérience.
         |
    OUVRAGES     (code_ouv, lot, libelle_ouv, unite_ouv, code_ref)
                 code_ref = référence du référentiel du pouvoir adjudicateur
                 (CCT 2022 Bruxelles / CCT-B Qualiroute / SB 250) — à remplir
                 au premier cahier des charges reçu.
         v
    BORDEREAU    calculé (moteur.calcul_bordereau)

CODIFICATION — ne jamais fusionner les deux espaces de nommage :
  · ouvrages de la bibliothèque   : `LL.NN`  (lot.numéro, ex. 40.20)
  · postes d'un métré imposé      : `NN.NN`  (ex. 03.02)
Le lien se fait exclusivement par MAPPING (ou la colonne `code_ref`).
"""

# ═══════════════════════════════════════════════════════════════════════════
# 0. PARAMÈTRES DE PRIX
# ═══════════════════════════════════════════════════════════════════════════
#
#   pu_vente = debourse_sec x K
#   K        = (1+FG) x (1+FC) x (1+aleas) x (1+marge)
#
# FG    frais généraux (siège, véhicules, assurances, administratif)
# FC    frais de chantier non imputables à un ouvrage précis
# aleas provision pour imprévus
# marge bénéfice visé
#
# 0,12 / 0,05 / 0,03 / 0,10  ->  K = 1,3324  (+33,24 % sur le déboursé sec)

PARAMS = {
    "fg": 0.12,
    "fc": 0.05,
    "aleas": 0.03,
    "marge": 0.10,
    # 6 % : logement privé de plus de 10 ans, facturé au consommateur final.
    # EN MARCHÉ PUBLIC C'EST 21 % — passer tva=0.21 aux fonctions de devis.
    "tva": 0.06,
    "tva_marche_public": 0.21,
}

# ═══════════════════════════════════════════════════════════════════════════
# 1. RESSOURCES — 39 lignes
# ═══════════════════════════════════════════════════════════════════════════
#
# ⚠️ MO : le taux horaire doit être le COÛT ENTREPRISE COMPLET (salaire brut
# + ONSS patronale + pécule + jours fériés + assurance loi + déplacements +
# EPI), PAS le brut. Un taux sous-évalué ici se propage à tous les ouvrages.
#
# (code_res, libelle_res, type_res, unite_res, pu_res)

_RESSOURCES_BRUT = [
    # ── Main d'œuvre (MO) — €/h, coût entreprise complet ──────────────────
    ("MO.01", "Chef de chantier",                        "MO", "h", 58.00),
    ("MO.02", "Ouvrier qualifié maçon / façadier",       "MO", "h", 48.00),
    ("MO.03", "Ouvrier qualifié plafonneur",             "MO", "h", 48.00),
    ("MO.04", "Ouvrier qualifié peintre",                "MO", "h", 45.00),
    ("MO.05", "Ouvrier qualifié menuisier",              "MO", "h", 50.00),
    ("MO.06", "Ouvrier qualifié sanitaire",              "MO", "h", 52.00),
    ("MO.07", "Manœuvre",                                "MO", "h", 38.00),
    ("MO.08", "Étancheur",                               "MO", "h", 52.00),

    # ── Matériaux (MAT) — prix d'achat rendu chantier, HTVA ───────────────
    ("MA.01", "Ciment gris CEM II 32,5 — sac 25 kg",     "MAT", "sac", 7.20),
    ("MA.02", "Sable de rivière 0/2",                    "MAT", "m3", 48.00),
    ("MA.03", "Chaux hydraulique NHL 3,5 — sac 25 kg",   "MAT", "sac", 14.50),
    ("MA.04", "Mortier de réparation fibré — sac 25 kg", "MAT", "sac", 22.00),
    ("MA.05", "Enduit de plafonnage plâtre — sac 25 kg", "MAT", "sac", 9.80),
    ("MA.06", "Enduit de façade minéral — sac 25 kg",    "MAT", "sac", 18.50),
    ("MA.07", "Primaire d'accrochage",                   "MAT", "L", 6.40),
    ("MA.08", "Peinture de façade siloxane",             "MAT", "L", 9.60),
    ("MA.09", "Peinture intérieure mate murs/plafonds",  "MAT", "L", 5.80),
    ("MA.10", "Membrane bitumineuse APP 4 mm",           "MAT", "m2", 12.50),
    ("MA.11", "Membrane EPDM 1,2 mm",                    "MAT", "m2", 16.80),
    ("MA.12", "Isolant PIR 80 mm",                       "MAT", "m2", 21.50),
    ("MA.13", "Laine minérale 100 mm",                   "MAT", "m2", 9.40),
    ("MA.14", "Panneau de plâtre BA13",                  "MAT", "m2", 6.20),
    ("MA.15", "Ossature métallique, suspentes, vis",     "MAT", "m2", 7.50),
    ("MA.16", "Profilé d'angle / cornière de finition",  "MAT", "m", 2.10),
    ("MA.17", "Linteau béton précontraint",              "MAT", "m", 34.00),
    ("MA.18", "Châssis PVC double vitrage",              "MAT", "m2", 320.00),
    ("MA.19", "Seuil en pierre bleue",                   "MAT", "m", 145.00),
    ("MA.20", "Porte d'entrée bois massif",              "MAT", "pce", 1250.00),
    ("MA.21", "Quincaillerie, fixations, consommables",  "MAT", "FF", 45.00),
    ("MA.22", "Carrelage grès cérame",                   "MAT", "m2", 32.00),
    ("MA.23", "Colle à carrelage — sac 25 kg",           "MAT", "sac", 13.50),
    ("MA.24", "Étanchéité liquide sous carrelage",       "MAT", "m2", 11.00),
    ("MA.25", "Garde-corps acier thermolaqué",           "MAT", "m", 185.00),
    ("MA.26", "Tuyau d'évacuation PVC Ø 110",            "MAT", "m", 12.80),
    ("MA.27", "Tuyau multicouche Ø 16 + raccords",       "MAT", "m", 6.50),

    # ── Matériel (EQP) — location / amortissement ─────────────────────────
    ("EQ.01", "Échafaudage de façade (location)",        "EQP", "m2.sem", 4.50),
    ("EQ.02", "Nacelle élévatrice (location)",           "EQP", "j", 190.00),
    ("EQ.03", "Conteneur 10 m3 + évacuation en centre",  "EQP", "pce", 380.00),
    ("EQ.04", "Petit matériel et consommables",          "EQP", "h", 2.20),
]

RESSOURCES = [
    {
        "code_res": code,
        "libelle_res": libelle,
        "type_res": type_res,
        "unite_res": unite,
        "pu_res": pu,
    }
    for code, libelle, type_res, unite, pu in _RESSOURCES_BRUT
]

# ═══════════════════════════════════════════════════════════════════════════
# 2. OUVRAGES — 36 postes, 8 lots
# ═══════════════════════════════════════════════════════════════════════════

LOTS = {
    "10": "Installation de chantier et préparation",
    "20": "Démolitions, déposes et évacuations",
    "30": "Maçonnerie et structure",
    "40": "Façades et étanchéité",
    "50": "Isolation",
    "60": "Plafonnage et plâtrerie",
    "70": "Peintures et revêtements",
    "80": "Menuiseries et sanitaire",
}

# (code_ouv, libelle_ouv, unite_ouv, code_ref)
# code_ref reste vide : à remplir avec les vraies références CCT au premier
# cahier spécial des charges reçu (CCT 2022 à Bruxelles).
_OUVRAGES_BRUT = [
    # ── Lot 10 — Installation de chantier ─────────────────────────────────
    ("10.10", "Installation et repli de chantier, amenée du matériel", "FF", ""),
    ("10.20", "Échafaudage de façade, montage, location 4 sem., démontage", "m2", ""),
    ("10.30", "Protection des ouvrages conservés, bâches et films", "m2", ""),

    # ── Lot 20 — Démolitions et évacuations ───────────────────────────────
    ("20.10", "Piquage d'enduit dégradé sur maçonnerie", "m2", ""),
    ("20.20", "Démolition de cloison légère, évacuation comprise", "m2", ""),
    ("20.30", "Dépose de plafond existant (plâtre ou plaques)", "m2", ""),
    ("20.40", "Dépose de menuiserie extérieure, calfeutrement provisoire", "pce", ""),
    ("20.50", "Évacuation des déchets en conteneur, tri compris", "m3", ""),

    # ── Lot 30 — Maçonnerie et structure ──────────────────────────────────
    ("30.10", "Maçonnerie de briques en rebouchage de baie ou trémie", "m2", ""),
    ("30.20", "Pose de linteau préfabriqué, étançonnement compris", "m", ""),
    ("30.30", "Réparation de béton dégradé, passivation des aciers", "m2", ""),
    ("30.40", "Cimentage hydrofuge de soubassement", "m2", ""),
    ("30.50", "Seuil en pierre bleue, pose et scellement", "m", ""),

    # ── Lot 40 — Façades et étanchéité ────────────────────────────────────
    ("40.10", "Nettoyage haute pression de façade", "m2", ""),
    ("40.20", "Enduit de façade minéral armé, deux couches", "m2", ""),
    ("40.30", "Peinture de façade siloxane, primaire compris", "m2", ""),
    ("40.40", "Étanchéité bitumineuse bicouche sur support préparé", "m2", ""),
    ("40.50", "Étanchéité EPDM collée, relevés compris", "m2", ""),
    ("40.60", "Solin, relevé et couvre-mur, zinguerie comprise", "m", ""),

    # ── Lot 50 — Isolation ────────────────────────────────────────────────
    ("50.10", "Isolation PIR 80 mm sous plafond, fixation mécanique", "m2", ""),
    ("50.20", "Isolation laine minérale 100 mm entre ossature", "m2", ""),
    ("50.30", "Pare-vapeur et adhésifs d'étanchéité à l'air", "m2", ""),

    # ── Lot 60 — Plafonnage et plâtrerie ──────────────────────────────────
    ("60.10", "Plafonnage sur maçonnerie, deux couches, dressé et lissé", "m2", ""),
    ("60.20", "Faux plafond BA13 sur ossature métallique, jointoyé", "m2", ""),
    ("60.30", "Rebouchage et enduit de rattrapage sur linteaux", "m", ""),
    ("60.40", "Cornières et profilés de finition", "m", ""),

    # ── Lot 70 — Peintures et revêtements ─────────────────────────────────
    ("70.10", "Peinture murs intérieurs, deux couches", "m2", ""),
    ("70.20", "Peinture plafonds intérieurs, deux couches", "m2", ""),
    ("70.30", "Enduit de lissage sur murs avant peinture", "m2", ""),
    ("70.40", "Carrelage de sol grès cérame, colle et joints compris", "m2", ""),

    # ── Lot 80 — Menuiseries et sanitaire ─────────────────────────────────
    ("80.10", "Châssis PVC double vitrage, pose et finitions", "m2", ""),
    ("80.20", "Porte d'entrée bois massif, quincaillerie comprise", "pce", ""),
    ("80.30", "Garde-corps acier thermolaqué, scellé", "m", ""),
    ("80.40", "Évacuation PVC Ø 110, pose et raccordement", "m", ""),
    ("80.50", "Alimentation multicouche Ø 16, pose et essais", "m", ""),
    ("80.60", "Étanchéité liquide sous carrelage de salle de bain", "m2", ""),
]

OUVRAGES = [
    {
        "code_ouv": code,
        "lot": code.split(".")[0],
        "libelle_ouv": libelle,
        "unite_ouv": unite,
        "code_ref": code_ref,
    }
    for code, libelle, unite, code_ref in _OUVRAGES_BRUT
]

# ═══════════════════════════════════════════════════════════════════════════
# 3. COMPOSITION — ressources consommées par unité d'ouvrage
# ═══════════════════════════════════════════════════════════════════════════
#
# ⚠️ Sur les lignes MO, qte_res est un RENDEMENT en heures par unité
# d'ouvrage. C'est LA donnée à faire valider par le chef d'entreprise : elle
# ne s'achète nulle part et elle pèse ~60 % du déboursé sec.
#
# (code_ouv, code_res, qte_res)

_COMPOSITION_BRUT = [
    # 10.10 Installation et repli (chantier de maison unifamiliale). Le conteneur
    #       n'est PAS ici : il est facturé au m3 par 20.50, sinon double compte.
    ("10.10", "MO.01", 3.00), ("10.10", "MO.07", 4.00), ("10.10", "MA.21", 1.00),
    # 10.20 Échafaudage (4 semaines de location)
    ("10.20", "MO.07", 0.35), ("10.20", "EQ.01", 4.00),
    # 10.30 Protection
    ("10.30", "MO.07", 0.08), ("10.30", "MA.21", 0.02),

    # 20.10 Piquage d'enduit
    ("20.10", "MO.07", 0.55), ("20.10", "EQ.04", 0.55), ("20.10", "EQ.03", 0.010),
    # 20.20 Démolition de cloison
    ("20.20", "MO.07", 0.70), ("20.20", "EQ.04", 0.70), ("20.20", "EQ.03", 0.015),
    # 20.30 Dépose de plafond
    ("20.30", "MO.07", 0.45), ("20.30", "EQ.04", 0.45), ("20.30", "EQ.03", 0.012),
    # 20.40 Dépose de menuiserie
    ("20.40", "MO.05", 1.20), ("20.40", "MO.07", 1.20), ("20.40", "EQ.03", 0.050),
    # 20.50 Évacuation
    ("20.50", "MO.07", 0.60), ("20.50", "EQ.03", 0.120),

    # 30.10 Maçonnerie de rebouchage
    ("30.10", "MO.02", 1.60), ("30.10", "MA.01", 1.20),
    ("30.10", "MA.02", 0.04), ("30.10", "EQ.04", 1.60),
    # 30.20 Linteau préfabriqué
    ("30.20", "MO.02", 1.40), ("30.20", "MO.07", 0.70), ("30.20", "MA.17", 1.05),
    ("30.20", "MA.01", 0.60), ("30.20", "EQ.04", 2.10),
    # 30.30 Réparation de béton
    ("30.30", "MO.02", 1.80), ("30.30", "MA.04", 1.50),
    ("30.30", "MA.07", 0.30), ("30.30", "EQ.04", 1.80),
    # 30.40 Cimentage hydrofuge
    ("30.40", "MO.02", 0.60), ("30.40", "MA.01", 1.40), ("30.40", "MA.02", 0.03),
    ("30.40", "MA.07", 0.10), ("30.40", "EQ.04", 0.60),
    # 30.50 Seuil en pierre bleue
    ("30.50", "MO.02", 0.90), ("30.50", "MA.19", 1.02), ("30.50", "MA.01", 0.40),

    # 40.10 Nettoyage haute pression
    ("40.10", "MO.07", 0.15), ("40.10", "EQ.04", 0.15),
    # 40.20 Enduit de façade
    ("40.20", "MO.02", 0.85), ("40.20", "MO.07", 0.25), ("40.20", "MA.06", 1.10),
    ("40.20", "MA.07", 0.15), ("40.20", "EQ.04", 1.10),
    # 40.30 Peinture de façade
    ("40.30", "MO.04", 0.32), ("40.30", "MA.08", 0.35), ("40.30", "MA.07", 0.12),
    # 40.40 Étanchéité bitumineuse bicouche (2,2 m2 de membrane par m2 posé,
    #       recouvrements et relevés compris)
    ("40.40", "MO.08", 0.55), ("40.40", "MA.10", 2.20),
    ("40.40", "MA.07", 0.20), ("40.40", "EQ.04", 0.55),
    # 40.50 Étanchéité EPDM
    ("40.50", "MO.08", 0.45), ("40.50", "MA.11", 1.15),
    ("40.50", "MA.21", 0.05), ("40.50", "EQ.04", 0.45),
    # 40.60 Solin et couvre-mur
    ("40.60", "MO.08", 0.40), ("40.60", "MA.04", 0.25), ("40.60", "MA.21", 0.08),

    # 50.10 Isolation PIR
    ("50.10", "MO.03", 0.30), ("50.10", "MA.12", 1.05), ("50.10", "MA.21", 0.06),
    # 50.20 Laine minérale
    ("50.20", "MO.03", 0.25), ("50.20", "MA.13", 1.08),
    # 50.30 Pare-vapeur
    ("50.30", "MO.03", 0.12), ("50.30", "MA.21", 0.04),

    # 60.10 Plafonnage
    ("60.10", "MO.03", 0.55), ("60.10", "MA.05", 0.60),
    ("60.10", "MA.07", 0.10), ("60.10", "EQ.04", 0.55),
    # 60.20 Faux plafond BA13
    ("60.20", "MO.03", 0.65), ("60.20", "MA.14", 1.05),
    ("60.20", "MA.15", 1.00), ("60.20", "MA.05", 0.10),
    # 60.30 Rebouchage sur linteaux
    ("60.30", "MO.03", 0.35), ("60.30", "MA.05", 0.25),
    # 60.40 Cornières
    ("60.40", "MO.03", 0.10), ("60.40", "MA.16", 1.05),

    # 70.10 Peinture murs
    ("70.10", "MO.04", 0.22), ("70.10", "MA.09", 0.30),
    # 70.20 Peinture plafonds
    ("70.20", "MO.04", 0.28), ("70.20", "MA.09", 0.32),
    # 70.30 Enduit de lissage
    ("70.30", "MO.04", 0.30), ("70.30", "MA.05", 0.30),
    # 70.40 Carrelage de sol
    ("70.40", "MO.02", 0.80), ("70.40", "MA.22", 1.08), ("70.40", "MA.23", 0.25),

    # 80.10 Châssis PVC
    ("80.10", "MO.05", 1.10), ("80.10", "MA.18", 1.00), ("80.10", "MA.21", 0.30),
    # 80.20 Porte d'entrée
    ("80.20", "MO.05", 6.00), ("80.20", "MA.20", 1.00), ("80.20", "MA.21", 1.00),
    # 80.30 Garde-corps
    ("80.30", "MO.05", 0.65), ("80.30", "MA.25", 1.00), ("80.30", "MA.21", 0.15),
    # 80.40 Évacuation PVC
    ("80.40", "MO.06", 0.35), ("80.40", "MA.26", 1.05), ("80.40", "MA.21", 0.10),
    # 80.50 Alimentation multicouche
    ("80.50", "MO.06", 0.30), ("80.50", "MA.27", 1.08), ("80.50", "MA.21", 0.12),
    # 80.60 Étanchéité liquide
    ("80.60", "MO.02", 0.30), ("80.60", "MA.24", 1.10), ("80.60", "MA.07", 0.10),
]

COMPOSITION = [
    {"code_ouv": code_ouv, "code_res": code_res, "qte_res": qte}
    for code_ouv, code_res, qte in _COMPOSITION_BRUT
]

# ═══════════════════════════════════════════════════════════════════════════
# 4. MAPPING — postes d'un métré imposé -> ouvrages de la bibliothèque
# ═══════════════════════════════════════════════════════════════════════════
#
# Table de correspondance pour le métré fictif CSC 2026-TP-0147 (49 postes).
# À REFAIRE À CHAQUE MARCHÉ : les codes de poste appartiennent au pouvoir
# adjudicateur, pas à nous. Un poste absent de ce dict est « non couvert » et
# devra être chiffré à la main OU faire l'objet d'un nouvel ouvrage.

MAPPING = {
    "00.01": "10.10",
    "00.02": "10.20",
    "00.04": "10.30",
    "00.05": "20.50",
    "01.01": "20.10",
    "01.02": "20.20",
    "01.03": "20.30",
    "01.04": "20.40",
    "02.01": "30.10",
    "02.03": "30.30",
    "02.04": "30.20",
    "02.05": "30.50",
    "03.01": "40.10",
    "03.02": "40.20",
    "03.03": "40.30",
    "03.04": "40.60",
    "03.05": "30.40",
    "04.01": "40.40",
    "04.02": "40.50",
    "04.03": "80.60",
    "04.04": "80.40",
    "05.01": "50.10",
    "05.02": "50.20",
    "05.03": "50.30",
    "06.02": "60.10",
    "06.03": "60.20",
    "06.05": "70.40",
    "06.06": "60.30",
    "06.07": "60.40",
    "07.01": "70.10",
    "07.02": "70.20",
    "07.04": "70.30",
    "08.01": "80.10",
    "08.02": "80.20",
    "08.03": "80.30",
    "08.06": "80.50",
    # Non couverts (13) — voir NON_COUVERTS ci-dessous :
    # 00.03 · 00.06 · 01.05 · 01.06 · 02.02 · 06.01 · 06.04
    # 07.03 · 08.04 · 08.05 · 09.03 · 09.04 · 09.05
}

# Les 13 postes du métré fictif que la bibliothèque ne sait pas chiffrer.
# Il manque les rendements (h/unité) du client pour créer les ouvrages
# correspondants. Laisser un de ces postes sans prix rend l'offre IRRÉGULIÈRE
# (art. 76 AR 18/04/2017) : il faut soit les chiffrer à la main, soit les
# ajouter à OUVRAGES + COMPOSITION.
NON_COUVERTS = [
    ("00.03", "Signalisation, clôture et sécurisation des accès", "FF"),
    ("00.06", "Dossier as-built, PV de réception, garanties", "FF"),
    ("01.05", "Dépose de revêtements de sol et chape existante", "m2"),
    ("01.06", "Dépose d'appareils sanitaires et tuyauterie", "pce"),
    ("02.02", "Rejointoiement de maçonnerie, mortier de chaux", "m2"),
    ("06.01", "Chape de ravoirage armée 6 cm", "m2"),
    ("06.04", "Faïence murale, profils de finition compris", "m2"),
    ("07.03", "Peinture sur menuiseries bois, ponçage compris", "m2"),
    ("08.04", "WC suspendu complet, bâti-support inclus", "pce"),
    ("08.05", "Essais d'étanchéité, rinçage, mise en service", "FF"),
    ("09.03", "Prise de courant 2P+T encastrée", "pce"),
    ("09.04", "Mise à la terre et liaisons équipotentielles RGIE", "FF"),
    ("09.05", "Contrôle de conformité par organisme agréé", "FF"),
]

# ═══════════════════════════════════════════════════════════════════════════
# 5. METRES_HISTO — re-chiffrage des 6 devis forfaitaires historiques
# ═══════════════════════════════════════════════════════════════════════════
#
# ⚠️ LES QUANTITÉS SONT DES ESTIMATIONS faites à partir des descriptifs des
# devis PDF, PAS DES RELEVÉS. C'est le premier point à corriger avec le
# client (cf. CLAUDE.md §7). Tant qu'elles ne sont pas remplacées par des
# surfaces réelles, les écarts affichés par moteur.calibration() mesurent
# autant l'erreur de métré que l'erreur de prix.
#
# forfait = montant réellement vendu, HTVA.

METRES_HISTO = {
    "07": {
        "objet": "Balcon avant — Av. Ernest Renan 35",
        "date": "11/05/2026",
        "forfait": 2500.00,
        "lignes": [
            ("10.10", 0.30),   # installation partielle, chantier court
            ("10.20", 11.0),   # échafaudage façade avant, m2
            ("20.10", 4.0),    # piquage de l'enduit sous-face
            ("30.30", 3.0),    # réparation du béton du plateau
            ("40.40", 6.0),    # étanchéité du plateau, relevés compris
            ("40.60", 4.0),    # solin et couvre-mur
            ("40.30", 10.0),   # remise en peinture des joues et du dessous
            ("20.50", 1.0),
        ],
    },
    "10": {
        "objet": "Plafond côté route",
        "date": "01/06/2026",
        "forfait": 1500.00,
        "lignes": [
            ("10.10", 0.25),
            ("20.30", 12.0),   # dépose du plafond existant
            ("60.20", 12.0),   # faux plafond BA13
            ("70.20", 12.0),   # mise en peinture
            ("20.50", 1.2),
        ],
    },
    "11": {
        "objet": "Jardin arrière + fenêtre 1er étage",
        "date": "01/06/2026",
        "forfait": 1500.00,
        "lignes": [
            ("10.10", 0.25),
            ("10.30", 5.0),
            ("20.40", 1.0),    # dépose de l'ancien châssis
            ("80.10", 1.5),    # châssis PVC, m2
            ("60.30", 4.0),    # rebouchage du pourtour
            ("60.40", 2.0),
            ("70.10", 10.0),   # reprise peinture intérieure
            ("30.40", 4.0),    # cimentage du muret côté jardin
            ("20.50", 0.8),
        ],
    },
    "13": {
        "objet": "Cave + porte d'entrée",
        "date": "03/06/2026",
        "forfait": 930.00,
        # ⚠️ La partie « porte d'entrée » est en réalité une remise en peinture
        # de menuiserie : l'ouvrage correspondant N'EXISTE PAS ENCORE dans la
        # bibliothèque (c'est le poste 07.03 de la liste des non couverts).
        # Elle est ici approchée par 70.30 (enduit de lissage). À reprendre dès
        # que le client aura donné le rendement de la peinture sur bois.
        "lignes": [
            ("10.10", 0.15),
            ("70.10", 30.0),   # murs de cave
            ("70.20", 8.0),    # plafond de cave
            ("70.30", 11.0),   # préparation + approximation de la porte
        ],
    },
    "15": {
        "objet": "Façade salle de bain",
        "date": "25/06/2026",
        "forfait": 3650.00,
        "lignes": [
            ("10.10", 0.40),
            ("10.20", 20.0),   # échafaudage
            ("20.10", 20.0),   # piquage de l'enduit dégradé
            ("40.10", 20.0),   # nettoyage haute pression
            ("40.20", 20.0),   # enduit de façade armé
            ("40.30", 20.0),   # peinture siloxane
            ("20.50", 1.0),
        ],
    },
    "16": {
        "objet": "Plafond + linteaux + isolation",
        "date": "03/07/2026",
        "forfait": 2400.00,
        "lignes": [
            ("10.10", 0.30),
            ("10.30", 15.0),
            ("20.30", 15.0),   # dépose du plafond
            ("30.20", 3.0),    # linteaux, m
            ("50.10", 15.0),   # isolation PIR sous plafond
            ("60.20", 15.0),   # faux plafond BA13
            ("60.30", 3.0),    # rebouchage sur linteaux
            ("70.20", 15.0),   # peinture
            ("20.50", 1.2),
        ],
    },
}

# ── Index pratiques (construits une fois à l'import) ──────────────────────
RESSOURCES_PAR_CODE = {r["code_res"]: r for r in RESSOURCES}
OUVRAGES_PAR_CODE = {o["code_ouv"]: o for o in OUVRAGES}
