# Devis generator — chiffrage de marchés publics

**BAG BATTER SRL** — Ronkel 18, 1780 Wemmel · TVA BE0766637025
Entreprise générale de rénovation : façades, étanchéité, plafonnage, isolation,
peinture, menuiserie extérieure, sanitaire léger.

Bibliothèque de prix unitaires réutilisable, pour répondre à un **métré imposé**
par un pouvoir adjudicateur sans re-chiffrer à zéro à chaque marché.

---

## Le problème

L'entreprise établit tous ses devis en **forfait global** : « FF 1,00 », un seul montant pour dix pages de
prestations. C'est irrecevable en marché public — le pouvoir adjudicateur
impose un métré où **chaque poste** porte une unité, une quantité et un prix
unitaire.

## Le modèle

Trois tables sources, une table calculée.

```
RESSOURCES   (code_res, libelle_res, type_res, unite_res, pu_res)
             type_res : MO (main d'œuvre) | MAT (matériaux) | EQP (matériel)
     |
COMPOSITION  (code_ouv, code_res, qte_res)
             qte_res sur les lignes MO = RENDEMENT en h/unité.
             Seule donnée non achetable : elle vient de l'expérience.
     |
OUVRAGES     (code_ouv, lot, libelle_ouv, unite_ouv, code_ref)
     v
BORDEREAU    calculé : deb_mo, deb_mat, deb_eqp, debourse_sec,
             pu_vente, heures_mo
```

```
debourse_sec = Σ (qte_res × pu_res)
pu_vente     = debourse_sec × K
K            = (1+FG)(1+FC)(1+aléas)(1+marge)
```

Avec les paramètres actuels — FG 12 %, FC 5 %, aléas 3 %, marge 10 % —
**K = 1,3324** (+33,24 % sur le déboursé sec).

---

## ⚠️ État des données : reconstruction, pas calibration

Le notebook Colab d'origine a été perdu. Ce code en est une **reconstruction à
partir du seul document de reprise**.

| | |
|---|---|
| **Fidèle** | structure des tables, formule de prix, K = 1,3324, 39 ressources, 36 ouvrages sur 8 lots, 13 postes non couverts, pièges openpyxl documentés |
| **Re-saisi** | tous les prix d'achat (`pu_res`), tous les rendements MO (`qte_res`), toutes les quantités des devis historiques |

Les valeurs numériques sont des **ordres de grandeur du marché belge 2026**, pas
les chiffres calibrés du chef d'entreprise.

**Aucune offre ne doit partir sur cette base avant relecture** des taux horaires
et des rendements. Point d'entrée : `python -m chiffrage calibration`.

---

## Utilisation

```bash
pip install -r requirements.txt                  # openpyxl, pour la chaîne Excel

python -m chiffrage controle                     # intégrité des trois tables
python -m chiffrage bordereau                    # les 36 prix unitaires
python -m chiffrage calibration                  # comparaison aux 6 devis vendus
python -m chiffrage fiche 40.20                  # justification d'un prix
python -m chiffrage devis 40.20:26 70.10:120 --tva=21
python -m chiffrage export  bibliotheque.xlsx    # bibliothèque -> Excel, 6 onglets
python -m chiffrage metre   metre.xlsx           # métré de marché public fictif
python -m chiffrage offre   metre.xlsx offre.xlsx
```

En Python (ou dans une cellule Colab) :

```python
from chiffrage import devis, fiche_prix, calcul_bordereau

d = devis("Rénovation façade arrière", [("40.20", 26), ("40.30", 26)], tva=0.21)
print(d["total_ht"], d["heures_mo"])
print(fiche_prix("40.20"))
```

### Modules

| Fichier | Rôle | Dépendance |
|---|---|---|
| `bibliotheque.py` | les données : RESSOURCES · OUVRAGES · COMPOSITION · MAPPING · PARAMS · METRES_HISTO | — |
| `moteur.py` | calcul du bordereau, devis, fiche de prix, calibration, contrôle de cohérence | — |
| `export_xlsx.py` | bibliothèque → Excel 6 onglets, **avec vraies formules** | openpyxl |
| `gen_metre.py` | métré de marché public fictif (49 postes, 10 lots) pour s'entraîner | openpyxl |
| `metre_io.py` | lecture d'un métré imposé + remplissage de l'offre | openpyxl |
| `__main__.py` | ligne de commande | — |

Le moteur reste en **Python pur** : il tourne en CI et se colle tel quel dans
une cellule Colab.

### Structure du dépôt

```
.
├── chiffrage/                → le paquet, appelable par `python -m chiffrage`
│   ├── bibliotheque.py       → RESSOURCES · OUVRAGES · COMPOSITION · MAPPING ·
│   │                            PARAMS · METRES_HISTO            (Python pur)
│   ├── moteur.py             → bordereau · devis · fiche de prix ·
│   │                            calibration · contrôle de cohérence (Python pur)
│   ├── gen_metre.py          → métré de marché public fictif       (openpyxl)
│   ├── metre_io.py           → lecture d'un métré imposé + offre   (openpyxl)
│   ├── export_xlsx.py        → bibliothèque -> Excel 6 onglets     (openpyxl)
│   └── __main__.py           → ligne de commande
├── tests/test_chiffrage.py   → 22 tests (8 se skippent sans openpyxl)
├── requirements.txt          → openpyxl (chaîne Excel uniquement)
├── requirements-dev.txt      → + pytest, ruff
├── ruff.toml · pytest.ini    → lint + config de tests
└── .github/workflows/ci.yml  → CI : ruff, pytest, contrôle de cohérence
```

---

## Arborescence Drive du client

```
MyDrive/BAG_BATTER/Chiffrage/
    01_bibliotheque/     bibliotheque_prix_bagbatter.xlsx (6 onglets)
    02_metres_recus/     métrés reçus des pouvoirs adjudicateurs
    03_offres_remises/   métrés complétés, horodatés
    04_archives/         versions successives de la bibliothèque
```

---

## Conventions et pièges à respecter

**Codification.** Ouvrages en `LL.NN` (lot.numéro, ex. `40.20`). Postes de métré
imposé en `NN.NN` (ex. `03.02`). **Ne jamais fusionner les deux espaces de
nommage** : le lien se fait exclusivement par `MAPPING` ou la colonne
`code_ref`. Un test le vérifie.

**openpyxl — écriture dans un métré imposé.** Toujours ouvrir **sans**
`data_only=True`. Avec ce paramètre, les formules du pouvoir adjudicateur sont
définitivement remplacées par des valeurs figées à la sauvegarde, et le fichier
renvoyé ne recalcule plus rien.

**openpyxl — `insert_rows` ne décale pas les formules.** Une formule écrite
`=IF($G20="","",$F20*$G20)` reste littéralement attachée à la ligne 20 même si
la cellule passe en ligne 21 : le fichier s'ouvre sans erreur et calcule faux.
`gen_metre.py` écrit donc les lignes de sous-total **au fil de la boucle**,
jamais insérées après coup — et n'appelle jamais `insert_rows`.

**Contrôle des unités.** Chiffrer au m² un poste imposé au mètre courant ne se
voit qu'au moment de facturer. Quand les unités divergent, `remplir_metre()`
**n'écrit pas de prix** et remonte le poste dans `ecarts_unite` : c'est un
arbitrage humain, pas une conversion automatique. **Ne jamais désactiver ce
contrôle.**

**Nature des postes.** `QF` = quantité forfaitaire garantie · `QP` = quantité
présumée, payée au métré réellement exécuté · `FF` = forfait global, quantité 1.

**Irrégularité.** Un poste laissé sans prix rend l'offre irrégulière et entraîne
son rejet (art. 76 AR 18/04/2017). `remplir_metre()` liste explicitement les
postes restés vides ; cette liste doit être à **zéro** avant envoi.

**TVA.** 6 % uniquement si logement de plus de dix ans, usage principalement
privé, facturation au consommateur final. **En marché public : 21 %.**

---

## État de la calibration

Les six devis forfaitaires historiques re-chiffrés avec la bibliothèque
(`python -m chiffrage calibration`) :

| Devis | Objet | Forfait vendu | Calculé | Écart |
|---|---|---:|---:|---:|
| 07 | Balcon avant (Av. E. Renan 35) | 2 500 € | 2 212 € | −11,5 % |
| 10 | Plafond côté route | 1 500 € | 1 565 € | +4,3 % |
| 11 | Jardin arrière + fenêtre 1er | 1 500 € | 1 667 € | +11,1 % |
| 13 | Cave + porte d'entrée | 930 € | 934 € | +0,5 % |
| 15 | Façade salle de bain | 3 650 € | 4 440 € | **+21,7 %** |
| 16 | Plafond + linteaux + isolation | 2 400 € | 3 424 € | **+42,7 %** |

Écart moyen absolu **15,3 %**. Cible : < 15 % sur **chaque** ligne.

Les devis 15 et 16 ressortent au-dessus du prix vendu. Deux hypothèses, non
tranchées faute de relevés :

1. les quantités de `METRES_HISTO` sont surestimées — elles viennent des
   descriptifs des devis PDF, **pas de relevés** ;
2. ces deux chantiers ont effectivement été vendus sous leur coût analytique.

Le devis 16 demande 37 h de main-d'œuvre au chiffrage analytique. À 2 400 €
HTVA matériaux compris, cela fait ~66 €/h : les frais généraux ne sont pas
couverts.

## Test d'intégration

`python -m chiffrage metre m.xlsx && python -m chiffrage offre m.xlsx o.xlsx` :

- 49 postes lus, **36 chiffrés automatiquement**, 13 non couverts
- **0 écart d'unité**
- 72 formules dans le fichier d'offre, toutes ancrées sur leur propre ligne
- les formules du pouvoir adjudicateur survivent au remplissage

Les 13 postes non couverts (il manque les rendements du client) :

```
00.03  Signalisation, clôture et sécurisation des accès      FF
00.06  Dossier as-built, PV de réception, garanties          FF
01.05  Dépose de revêtements de sol et chape existante       m2
01.06  Dépose d'appareils sanitaires et tuyauterie           pce
02.02  Rejointoiement de maçonnerie, mortier de chaux        m2
06.01  Chape de ravoirage armée 6 cm                         m2
06.04  Faïence murale, profils de finition compris           m2
07.03  Peinture sur menuiseries bois, ponçage compris        m2
08.04  WC suspendu complet, bâti-support inclus              pce
08.05  Essais d'étanchéité, rinçage, mise en service         FF
09.03  Prise de courant 2P+T encastrée                       pce
09.04  Mise à la terre et liaisons équipotentielles RGIE     FF
09.05  Contrôle de conformité par organisme agréé            FF
```

---

## À faire

- [ ] **Relire les taux horaires** de `RESSOURCES` : coût entreprise complet
      (salaire + ONSS + congés + assurances + déplacements), pas le brut
- [ ] **Relire les rendements** (lignes MO de `COMPOSITION`) — ils pèsent ~60 %
      du déboursé sec et sont ici des ordres de grandeur, pas du vécu
- [ ] Obtenir les **surfaces réelles** des 6 chantiers historiques et corriger
      `METRES_HISTO` ; viser un écart < 15 % sur chaque ligne
- [ ] Trancher l'hypothèse sur les devis 15 et 16
      (sous-tarification ou surestimation des quantités)
- [ ] Collecter les rendements des **13 ouvrages manquants** et les intégrer
- [ ] Remplir `code_ref` avec les vraies références CCT (CCT 2022 Bruxelles /
      CCT-B Qualiroute / SB 250) au premier cahier des charges reçu
- [ ] Envisager une bascule vers une base (Oracle/PL-SQL) si la bibliothèque
      dépasse quelques centaines d'ouvrages
- [ ] Envisager un tableau de bord Qlik Sense sur l'historique des offres
      (taux de succès par lot, marge réalisée vs marge chiffrée)

## Tests

```bash
pip install -r requirements-dev.txt
pytest
```

22 tests. Les 8 tests de la chaîne Excel se skippent proprement si openpyxl
n'est pas installé.
