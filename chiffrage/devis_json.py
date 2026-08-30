"""Enregistrer un devis pour le reprendre plus tard.

Le `.xlsx` produit est un document destiné au client : mise en page,
formules, conditions. Le relire pour en retrouver les intentions
serait de la rétro-ingénierie sur un document de sortie — et il n'y a
rien à en tirer que l'app ne sache déjà recalculer. On enregistre donc
à côté ce qui a réellement été SAISI : l'en-tête, le taux de TVA et la
liste (ouvrage, quantité). Les prix, eux, ne sont pas dans le fichier :
ils viennent de la bibliothèque au moment de la relecture. Un devis
repris six mois plus tard est donc chiffré aux prix d'aujourd'hui, ce
qui est le comportement voulu — un devis rouvert est un devis à
réémettre, pas une archive.

Même choix que pour le lexique et la correspondance de métré : du JSON,
jamais du Python. Le contenu vient de champs de saisie, et il est relu
par une app qui tourne ensuite.
"""
import json

VERSION = 1

TAUX_TVA = (0.06, 0.21)
# Dans le doute, 21 % : sous-facturer la TVA se paie au contrôle, la
# sur-facturer se corrige par une note de crédit.
TVA_PAR_DEFAUT = 0.21

CHAMPS_TEXTE = ("objet", "reference", "chantier", "client")


def serialiser(objet, reference, chantier, client, tva, lignes):
    """Rend le texte JSON d'un devis. `lignes` : [(code_ouv, qte)]."""
    return json.dumps(
        {
            "version": VERSION,
            "objet": objet,
            "reference": reference,
            "chantier": chantier,
            "client": client,
            "tva": tva,
            "lignes": [{"code_ouv": code, "qte": float(qte)}
                        for code, qte in lignes],
        },
        indent=2, ensure_ascii=False,
    ) + "\n"


def lire(octets, codes_connus):
    """Relit un devis enregistré. Rend `(devis, anomalies)`.

    Rien de ce qui arrive ici n'est digne de confiance : le fichier a
    pu être édité à la main, ou venir d'une version de la bibliothèque
    où un ouvrage existait encore. Une ligne douteuse est écartée et
    signalée — jamais devinée, jamais laissée passer en silence : un
    code inconnu ferait planter le tableau d'édition, et une quantité
    négative produirait un devis à montant négatif sans prévenir.
    """
    charge = json.loads(octets.decode("utf-8") if isinstance(octets, bytes)
                        else octets)
    if not isinstance(charge, dict):
        raise ValueError("le fichier ne contient pas un devis "
                          f"(type {type(charge).__name__})")

    anomalies = []
    devis = {champ: str(charge.get(champ) or "") for champ in CHAMPS_TEXTE}

    tva = charge.get("tva")
    if tva in TAUX_TVA:
        devis["tva"] = tva
    else:
        if tva is not None:
            anomalies.append(f"taux de TVA inattendu ({tva}) — remis à "
                              f"{TVA_PAR_DEFAUT * 100:.0f} %")
        devis["tva"] = TVA_PAR_DEFAUT

    brutes = charge.get("lignes")
    if not isinstance(brutes, list):
        raise ValueError("le fichier ne contient aucune liste de postes")

    lignes = []
    for rang, ligne in enumerate(brutes, start=1):
        if not isinstance(ligne, dict):
            anomalies.append(f"poste {rang} : ligne illisible, ignorée")
            continue
        code = ligne.get("code_ouv")
        if code not in codes_connus:
            anomalies.append(f"poste {rang} : ouvrage « {code} » absent de "
                              "la bibliothèque, ignoré")
            continue
        try:
            qte = float(ligne.get("qte"))
        except (TypeError, ValueError):
            anomalies.append(f"poste {rang} ({code}) : quantité illisible, "
                              "ignoré")
            continue
        if qte < 0:
            anomalies.append(f"poste {rang} ({code}) : quantité négative, "
                              "ignoré")
            continue
        lignes.append({"code_ouv": code, "qte": qte})

    devis["lignes"] = lignes
    return devis, anomalies
