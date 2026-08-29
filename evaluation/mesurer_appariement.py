"""
Mesure l'appariement poste imposé -> ouvrage sur le jeu d'épreuve.

    python evaluation/mesurer_appariement.py

Trois chiffres, et le troisième est le plus important :

  rang 1   la bonne réponse arrive en tête ;
  top 3    elle est dans les trois premières — ce que voit l'utilisateur
           dans l'interface, donc ce qui compte vraiment pour lui ;
  silence  sur les postes qu'AUCUN ouvrage ne couvre, l'outil se tait
           (meilleur score sous le seuil de suggestion). Une suggestion
           confiante sur un poste inexistant est PIRE que pas de
           suggestion : elle fait chiffrer un travail par un autre.

Un gain sur « rang 1 » payé par une perte sur « silence » n'est pas un
gain.
"""

import json
import sys
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RACINE))

from chiffrage.moteur import calcul_bordereau  # noqa: E402
from chiffrage.suggestion import suggerer  # noqa: E402

SEUIL = 0.35        # identique à SEUIL_SUGGESTION de l'interface


def charger():
    donnees = json.loads(
        (RACINE / "evaluation" / "epreuves_appariement.json").read_text("utf-8")
    )
    return donnees["epreuves"]


def mesurer(epreuves, bordereau, verbeux=False):
    rang1 = top3 = 0
    attendus = silences_ok = silences = 0
    echecs = []

    for e in epreuves:
        poste = {"designation": e["libelle"], "unite": e["unite"]}
        candidats = suggerer(poste, bordereau, limite=3)
        codes = [c for c, _ in candidats]
        meilleur = candidats[0][1] if candidats else 0.0

        if e["attendu"] is None:
            silences += 1
            if meilleur < SEUIL:
                silences_ok += 1
            else:
                echecs.append(
                    (e, f"propose {codes[0]} à {meilleur:.2f} alors "
                        f"qu'aucun ouvrage ne convient")
                )
            continue

        attendus += 1
        if codes and codes[0] == e["attendu"]:
            rang1 += 1
            top3 += 1
        elif e["attendu"] in codes:
            top3 += 1
            if verbeux:
                echecs.append((e, f"rang {codes.index(e['attendu']) + 1} "
                                   f"— tête : {codes[0]}"))
        else:
            echecs.append((e, f"absent des 3 — proposé : "
                               f"{', '.join(codes) or 'rien'}"))

    return {
        "rang1": rang1, "top3": top3, "attendus": attendus,
        "silences_ok": silences_ok, "silences": silences,
        "echecs": echecs,
    }


def _pct(n, d):
    return f"{100 * n / d:.0f} %" if d else "—"


def imprimer(resultat, titre):
    r = resultat
    print(f"\n── {titre} " + "─" * max(0, 60 - len(titre)))
    print(f"  rang 1  {r['rang1']:>2}/{r['attendus']}  {_pct(r['rang1'], r['attendus'])}")
    print(f"  top 3   {r['top3']:>2}/{r['attendus']}  {_pct(r['top3'], r['attendus'])}")
    print(f"  silence {r['silences_ok']:>2}/{r['silences']}  "
          f"{_pct(r['silences_ok'], r['silences'])}"
          f"   (postes qu'aucun ouvrage ne couvre)")
    if r["echecs"]:
        print(f"\n  {len(r['echecs'])} cas à regarder :")
        for e, motif in r["echecs"]:
            print(f"    [{e['type']:<12}] {e['libelle'][:52]:<52} → {motif}")
            if e.get("attendu"):
                print(f"      attendu {e['attendu']}")


if __name__ == "__main__":
    epreuves = charger()
    bordereau = calcul_bordereau()
    imprimer(mesurer(epreuves, bordereau, verbeux="-v" in sys.argv),
             "appariement actuel")
