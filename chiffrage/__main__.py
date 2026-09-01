"""
Ligne de commande de l'outil de chiffrage.

    python -m chiffrage controle                    contrôle de cohérence
    python -m chiffrage bordereau                   les 36 prix unitaires
    python -m chiffrage calibration                 comparaison aux 6 devis vendus
    python -m chiffrage fiche 40.20                 justification d'un prix
    python -m chiffrage devis 40.20:26 70.10:120    devis à la volée
    python -m chiffrage export [fichier.xlsx]       bibliothèque -> Excel
    python -m chiffrage metre  [fichier.xlsx]       métré de marché public fictif
    python -m chiffrage metre2 [fichier.xlsx]       un second, d'un autre pouvoir
                                                   adjudicateur (autres colonnes,
                                                   une feuille par lot)
    python -m chiffrage offre  metre.xlsx offre.xlsx  remplit un métré imposé

`devis` accepte --tva=21 (défaut 6) et sait produire un vrai devis client :

    python -m chiffrage devis 40.20:26 40.30:26 --sortie=devis.xlsx \
        --nom="Rénovation façade arrière" --reference=2026-042 \
        --client="M. Dupont, Rue de l'Église 12, 1030 Schaerbeek" \
        --chantier="Av. Ernest Renan 62, 1030 Schaerbeek"

Les commandes export/metre/offre et l'option --sortie demandent openpyxl ;
tout le reste tourne en Python nu.
"""

import sys

from .bibliotheque import LOTS
from .moteur import (
    calcul_bordereau,
    coefficient_k,
    controle_coherence,
    devis,
    fiche_prix,
    imprimer_calibration,
    imprimer_devis,
)


def _cmd_controle():
    anomalies = controle_coherence()
    total = sum(len(v) for v in anomalies.values())
    if total == 0:
        b = calcul_bordereau()
        print("✅ Bibliothèque cohérente.")
        print(f"   {len(b)} ouvrages sur {len(LOTS)} lots · K = {coefficient_k():.4f}")
        return 0
    print(f"⚠️  {total} anomalie(s) :")
    for nom, liste in anomalies.items():
        if liste:
            print(f"   {nom} : {liste}")
    return 1


def _cmd_bordereau():
    b = calcul_bordereau()
    lot_courant = None
    for code in sorted(b):
        ligne = b[code]
        if ligne["lot"] != lot_courant:
            lot_courant = ligne["lot"]
            print(f"\nLOT {lot_courant} — {LOTS.get(lot_courant, '')}")
        print(
            f"  {code}  {ligne['pu_vente']:>9.2f} €/{ligne['unite_ouv']:<5}"
            f" {ligne['heures_mo']:>6.3f} h  {ligne['libelle_ouv']}"
        )
    return 0


_OPTIONS_DEVIS = ("nom", "sortie", "client", "chantier", "reference")


def _cmd_devis(args):
    tva = 0.06
    opts = {}
    lignes = []
    for arg in args:
        if arg.startswith("--tva="):
            tva = float(arg.split("=", 1)[1]) / 100.0
            continue
        if arg.startswith("--"):
            cle, _, valeur = arg[2:].partition("=")
            if cle in _OPTIONS_DEVIS:
                opts[cle] = valeur
            else:
                print(f"Option inconnue, ignorée : --{cle}")
            continue
        if ":" not in arg:
            print(f"Argument ignoré (attendu code:quantité) : {arg}")
            continue
        code, qte = arg.split(":", 1)
        lignes.append((code.strip(), float(qte.replace(",", "."))))

    if not lignes:
        print("Rien à chiffrer. Exemple : python -m chiffrage devis 40.20:26 --tva=21")
        return 2

    d = devis(opts.get("nom") or "Devis", lignes, tva=tva)
    print(imprimer_devis(d))

    sortie = opts.get("sortie")
    if not sortie:
        return 0
    try:
        from .devis_xlsx import exporter_devis
    except ImportError:
        print("\nopenpyxl est requis pour --sortie : pip install openpyxl")
        return 2
    try:
        chemin, nb = exporter_devis(
            d,
            sortie,
            client=opts.get("client"),
            chantier=opts.get("chantier"),
            reference=opts.get("reference"),
        )
    except ValueError as err:
        # Devis incomplet : mieux vaut pas de fichier qu'un fichier amputé.
        print(f"\n⚠️  {err}")
        return 1
    print(f"\nDevis client écrit : {chemin} ({nb} postes)")
    return 0


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv[0] in ("-h", "--help", "aide"):
        print(__doc__)
        return 0
    cmd, args = argv[0], argv[1:]

    if cmd == "controle":
        return _cmd_controle()
    if cmd == "bordereau":
        return _cmd_bordereau()
    if cmd == "calibration":
        print(imprimer_calibration())
        return 0
    if cmd == "fiche":
        if not args:
            print("Usage : python -m chiffrage fiche 40.20")
            return 2
        print(fiche_prix(args[0]))
        return 0
    if cmd == "devis":
        return _cmd_devis(args)

    # ── Commandes nécessitant openpyxl ───────────────────────────────────
    if cmd in ("export", "metre", "metre2", "offre"):
        try:
            if cmd == "export":
                from .export_xlsx import NOM_FICHIER_DEFAUT, exporter_bibliotheque

                chemin, nb = exporter_bibliotheque(
                    args[0] if args else NOM_FICHIER_DEFAUT
                )
                print(f"Bibliothèque exportée : {chemin} ({nb} onglets)")
                return 0
            if cmd == "metre":
                from .gen_metre import NOM_FICHIER_DEFAUT, generer_metre

                chemin, nb_postes, nb_lots = generer_metre(
                    args[0] if args else NOM_FICHIER_DEFAUT
                )
                print(f"Métré généré : {chemin} ({nb_postes} postes, {nb_lots} lots)")
                return 0
            if cmd == "metre2":
                from .gen_metre_b import NOM_FICHIER_DEFAUT, generer_metre

                chemin, nb_postes, nb_lots = generer_metre(
                    args[0] if args else NOM_FICHIER_DEFAUT
                )
                print(f"Métré généré : {chemin} ({nb_postes} postes, "
                       f"{nb_lots} feuilles de lot + récapitulatif)")
                return 0
            if len(args) < 2:
                print("Usage : python -m chiffrage offre metre.xlsx offre.xlsx")
                return 2
            from .metre_io import imprimer_rapport, remplir_metre

            rapport = remplir_metre(args[0], args[1])
            print(imprimer_rapport(rapport))
            return 1 if rapport["vides"] else 0
        except ImportError:
            print("openpyxl est requis pour cette commande : pip install openpyxl")
            return 2

    print(f"Commande inconnue : {cmd}")
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main())
