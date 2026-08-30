# Devis generator — chiffrage de marchés publics

**BAG BATTER SRL** — Ronkel 18, 1780 Wemmel · TVA BE0766637025
Entreprise générale de rénovation : façades, étanchéité, plafonnage, isolation,
peinture, menuiserie extérieure, sanitaire léger.

Bibliothèque de prix unitaires réutilisable, pour répondre à un **métré imposé**
par un pouvoir adjudicateur sans re-chiffrer à zéro à chaque marché.

**➜ [devis-generator.streamlit.app](https://devis-generator.streamlit.app/)** —
déposer le métré, relire les correspondances, télécharger l'offre.
Aucune installation.

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
| **Fidèle** | structure des tables, formule de prix, K = 1,3324, les 36 ouvrages d'origine sur 8 lots, pièges openpyxl documentés |
| **Re-saisi** | tous les prix d'achat (`pu_res`), tous les rendements MO (`qte_res`), toutes les quantités des devis historiques |
| **Ajouté** | 13 ouvrages (lot 90 compris) pour couvrir les postes qui restaient sans prix — voir `OUVRAGES_A_VALIDER` : ce sont les rendements les moins assis de tous |

Les valeurs numériques sont des **ordres de grandeur du marché belge 2026**, pas
les chiffres calibrés du chef d'entreprise.

**Aucune offre ne doit partir sur cette base avant relecture** des taux horaires
et des rendements. Point d'entrée : `python -m chiffrage calibration`.

---

## Utilisation

### L'interface web (le plus simple)

**[devis-generator.streamlit.app](https://devis-generator.streamlit.app/)**

Déposer le métré Excel reçu, relire les correspondances proposées,
télécharger l'offre. Aucune commande Python.

Déployée sur Streamlit Community Cloud depuis la branche `main`, fichier
principal `streamlit_app.py` : **chaque push met l'app à jour**. Pour la
faire tourner en local à la place :

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

> ⚠️ L'app est publique. Tant que les prix ne sont pas calibrés, c'est
> sans enjeu — mais le jour où ce seront les vrais taux horaires et
> rendements de l'entreprise, l'URL exposera sa structure de coûts à qui
> la connaît. L'accès se restreint à une liste d'adresses e-mail dans
> *Settings → Sharing*, côté Streamlit Cloud.

L'interface apporte ce que la ligne de commande ne pouvait pas :
**l'appariement des postes à l'écran**. `MAPPING` est à refaire à chaque
marché — les codes appartiennent au pouvoir adjudicateur — et c'était la
seule étape qui obligeait encore à éditer du Python. L'outil propose une
correspondance par ouvrage, à partir du libellé ; l'humain tranche.

Deux règles y sont gravées :

- **l'unité est éliminatoire, pas départageante** — un poste imposé au
  mètre courant ne se voit jamais proposer un ouvrage au m², quelle que
  soit la ressemblance des libellés. Le prix serait faux d'un facteur
  inconnu, et ça ne se verrait qu'à la facturation ;
- **le score n'est pas une probabilité** — c'est une ressemblance de
  libellés, sur un vocabulaire où « dépose » et « pose » ne diffèrent que
  d'une lettre. Un score élevé dit « regarde ici d'abord », jamais
  « c'est bon ».

Le vocabulaire du CSC est traduit vers celui de la bibliothèque avant
toute comparaison (`chiffrage/lexique.py`) : « carrelage mural » devient
« faïence », « crépi » devient « enduit », « mousse polyuréthane »
devient « PIR ». Déterministe et relisible — là où une similarité
cosinus est un nombre qu'on subit, une entrée de lexique est une
décision qu'on relit.

**L'onglet « Lexique » de l'interface le rend réglable sans Python.**
On colle un libellé qui n'a pas été apparié, l'app montre ce qu'elle
en a retenu (les mots gardés, l'opération détectée) et les cinq
candidats avec leur score ; on ajoute le terme manquant et on voit le
poste remonter au premier rang, immédiatement.

Deux limites, que l'app affiche plutôt que de les laisser découvrir :
l'ajout vaut pour **l'app entière et tous ses utilisateurs** (la
surcouche est globale au processus, et un serveur Streamlit sert tous
ses visiteurs depuis un seul processus), et il **ne survit pas au
redémarrage** — sauf à le commiter, voir juste en dessous.

### Régler l'entreprise et les coefficients

L'onglet **⚙️ Paramètres** porte la raison sociale, l'adresse, le
numéro de TVA et les coefficients de vente (FG, FC, aléas, marge, taux
de TVA). Ce sont des valeurs **d'entreprise**, pas des constantes
techniques : changer une adresse ou un point de marge ne devrait pas
demander d'éditer du Python.

Même dispositif que le lexique — `chiffrage/parametres_local.json`,
**du JSON et non du code**, écrit par le même bouton et le même jeton.
Un fichier absent, illisible ou partiellement aberrant ne bloque rien :
chaque bloc retombe indépendamment sur ses valeurs de repli, si bien
qu'une marge mal saisie ne fait pas perdre l'adresse.

À la différence du lexique, **rien ne se fusionne** : deux adresses ne
s'additionnent pas. Le dernier qui écrit gagne, mais pas en aveugle —
le `sha` relu juste avant fait échouer l'écriture si quelqu'un est
passé entre-temps.

Les coefficients de la **barre latérale** restent une simulation de
session : ils servent à essayer un autre K sans engager la référence.
L'onglet Paramètres, lui, fixe le point de départ.

### Rendre les termes permanents

Avec un jeton GitHub configuré, l'onglet Lexique affiche un bouton
**« Commiter ces termes »** : les termes appris sont écrits dans
`chiffrage/lexique_local.json`, l'app se redéploie seule, et ils
deviennent définitifs. Plus aucun aller-retour par un éditeur.

**C'est du JSON, pas du Python, et c'est délibéré.** Le contenu vient
d'un champ de saisie : écrire du code exécutable à partir d'une saisie
serait une injection — un terme contenant un guillemet ou un saut de
ligne deviendrait du code au prochain déploiement. Du JSON ne
s'exécute pas ; le pire cas est un synonyme absurde. Les termes sont
en plus bornés à l'entrée (lettres, chiffres, espaces, apostrophes,
tirets, 49 caractères).

Le commit **relit le fichier distant et fusionne** avant d'écrire :
deux personnes peuvent régler le lexique le même jour sans que l'une
efface les termes de l'autre.

**Configuration** — Streamlit Cloud, *Settings → Secrets* :

```toml
[github]
token = "github_pat_..."
depot = "pmeyssonnier/Devis-generator"
branche = "main"
```

⚠️ **Un PAT fine-grained**, limité au seul dépôt `Devis-generator`,
permission *Contents: read and write*, avec une date d'expiration.
**Pas** un jeton classique à portée `repo` : celui-là donne
l'écriture sur **tous** les dépôts du compte, pour une fonction qui
n'a besoin que d'un fichier.

Sans jeton, rien ne casse : le bouton n'apparaît pas, et l'app
continue de rendre le bloc à coller à la main.

**L'opération est une facette, pas un mot.** « Dépose de carrelage » et
« Pose de carrelage » ne diffèrent que par un mot, et ce sont deux
travaux opposés à des prix qui vont du simple au triple. Les verbes de
démolition sont donc retirés de la comparaison lexicale et traités comme
une dimension à part, comme l'unité — sinon deux démolitions **sans
rapport** se ressemblent (« dépose du carrelage mural » proposait
« dépose de plafond » : deux fois le mot « dépose », et rien d'autre en
commun). Pénalité et non élimination, à la différence de l'unité :
l'unité est déclarée dans le métré, l'opération est inférée de mots, et
une inférence fausse ne doit pas faire disparaître un candidat valable.

Mesuré sur `evaluation/` (voir plus bas) : **98 % au premier rang, 100 %
dans les trois premiers**, et l'outil se tait sur 7 des 8 postes
qu'aucun ouvrage ne couvre.

La correspondance obtenue se télécharge en `.json` et se recharge au
marché suivant : une commune réutilise ses propres codes.

### Depuis le téléphone — Colab

Ouvre **[`colab/chiffrage_bagbatter.ipynb`](colab/chiffrage_bagbatter.ipynb)**
dans Google Colab : il installe l'outil, monte ton Drive, crée l'arborescence
`BAG_BATTER/Chiffrage/` et **y écrit directement les fichiers produits**,
horodatés. Neuf sections : contrôle et calibration · export de la
bibliothèque · devis client · réponse à un métré imposé · métré
d'entraînement · consultation des prix.

C'est la seule voie où les documents atterrissent dans Drive sans manipulation.

### En ligne de commande

Les fichiers sortent **dans le dossier où tu lances la commande** — rien
n'est envoyé vers Drive.

```bash
pip install -r requirements.txt                  # openpyxl, pour la chaîne Excel

python -m chiffrage controle                     # intégrité des trois tables
python -m chiffrage bordereau                    # les 36 prix unitaires
python -m chiffrage calibration                  # comparaison aux 6 devis vendus
python -m chiffrage fiche 40.20                  # justification d'un prix
python -m chiffrage devis 40.20:26 70.10:120 --tva=21

# Devis client prêt à envoyer (Excel, avec formules vivantes)
python -m chiffrage devis 40.20:26 40.30:26 --sortie=devis.xlsx \
    --nom="Rénovation façade arrière" --reference=2026-042 \
    --client="M. Dupont, Rue de l'Église 12, 1030 Schaerbeek" \
    --chantier="Av. Ernest Renan 62, 1030 Schaerbeek"

python -m chiffrage export  bibliotheque.xlsx    # bibliothèque -> Excel, 6 onglets
python -m chiffrage metre   metre.xlsx           # métré de marché public fictif
python -m chiffrage offre   metre.xlsx offre.xlsx
```

En Python (ou dans une cellule Colab) :

```python
from chiffrage import devis, fiche_prix
from chiffrage.devis_xlsx import exporter_devis

d = devis("Rénovation façade arrière", [("40.20", 26), ("40.30", 26)], tva=0.06)
print(d["total_ht"], d["heures_mo"])
print(fiche_prix("40.20"))

exporter_devis(d, "devis_2026-042.xlsx",
               client="M. et Mme Dupont\nRue de l'Église 12\n1030 Schaerbeek",
               chantier="Avenue Ernest Renan 62, 1030 Schaerbeek",
               reference="2026-042")
```

### Modules

| Fichier | Rôle | Dépendance |
|---|---|---|
| `bibliotheque.py` | les données : RESSOURCES · OUVRAGES · COMPOSITION · MAPPING · PARAMS · METRES_HISTO | — |
| `moteur.py` | calcul du bordereau, devis, fiche de prix, calibration, contrôle de cohérence | — |
| `export_xlsx.py` | bibliothèque → Excel 6 onglets, **avec vraies formules** | openpyxl |
| `devis_xlsx.py` | devis client prêt à envoyer : en-tête, postes par lot, TVA, conditions, signature | openpyxl |
| `suggestion.py` | appariement poste imposé → ouvrage, à partir du libellé | — |
| `lexique.py` | vocabulaire de CSC → vocabulaire de la bibliothèque, facette d'opération | — |
| `parametres.py` | identité de l'entreprise et coefficients de vente, réglables depuis l'app | — |
| `data/*.json` | les tables elles-mêmes : ressources, ouvrages, composition, lots, mapping | — |
| `detection_colonnes.py` | quelle colonne est le code, la quantité, le prix | — |
| `depot_github.py` | écrit le lexique appris dans le dépôt (API GitHub, urllib) | — |
| `controle_prix.py` | relit une offre avant dépôt : couverture, rabais maximal, alertes | — |
| `justification_xlsx.py` | dossier de justification de prix (art. 36) | openpyxl |
| `gen_metre.py` | métré de marché public fictif (49 postes, 10 lots) pour s'entraîner | openpyxl |
| `metre_io.py` | lecture d'un métré imposé + remplissage de l'offre | openpyxl |
| `__main__.py` | ligne de commande | — |

Le moteur reste en **Python pur** : il tourne en CI et se colle tel quel dans
une cellule Colab.

### Structure du dépôt

```
.
├── chiffrage/                → le paquet, appelable par `python -m chiffrage`
│   ├── data/*.json           → LES DONNÉES : ressources, ouvrages,
│   │                            composition, lots, mapping, historique
│   ├── bibliotheque.py       → chargement et CONTRÔLE des tables  (Python pur)
│   ├── moteur.py             → bordereau · devis · fiche de prix ·
│   │                            calibration · contrôle de cohérence (Python pur)
│   ├── gen_metre.py          → métré de marché public fictif       (openpyxl)
│   ├── metre_io.py           → lecture d'un métré imposé + offre   (openpyxl)
│   ├── export_xlsx.py        → bibliothèque -> Excel 6 onglets     (openpyxl)
│   ├── devis_xlsx.py         → devis client prêt à envoyer         (openpyxl)
│   └── __main__.py           → ligne de commande
├── streamlit_app.py        → interface web : dépôt du métré, appariement
│                               des postes à l'écran, réglage du lexique,
│                               téléchargement
├── evaluation/              → jeu d'épreuve de l'appariement + mesure
├── colab/                   → notebook Colab : monte Drive, range les
│                               fichiers dans BAG_BATTER/Chiffrage/
├── tests/                   → 42 tests (se skippent sans openpyxl/streamlit)
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

**Les colonnes sont détectées, pas supposées.** `B = code`,
`F = quantité`, `G = prix` était la disposition du métré
d'entraînement — chez une autre commune, l'outil ne lisait rien, ou
pire, écrivait le prix dans la mauvaise colonne. `detection_colonnes.py`
lit les intitulés (« Qté », « Métré », « P.U. HTVA » désignent la même
chose) et, à défaut, le contenu.

Une exception éclaire le reste : **pour la colonne des codes, le
contenu l'emporte sur l'intitulé.** Un métré titre couramment « N° »
son simple compteur de lignes, juste avant la vraie colonne des codes ;
se fier au titre partait sur le compteur et ne lisait plus aucun poste.
C'est le seul champ qu'on sache vérifier — une cellule ressemble à un
code, ou non.

La détection **propose**, l'interface affiche la correspondance et
permet de la corriger avant tout chiffrage.

**Les codes appartiennent au pouvoir adjudicateur.** Le lecteur accepte
`03.02`, `3.2`, `01.02.03`, `03.02.A`, `1.01.10`, `A.1.2`, `03-02`,
`03/02` — auparavant seul `NN.NN` passait, et un cahier des charges
numéroté `01.02.03` rendait **zéro** poste, avec pour tout message
« aucun poste lu ». Élargir un motif risque de prendre pour un poste ce
qui n'en est pas : le garde-fou n'est pas la sévérité du motif mais la
**seconde condition** — une ligne n'est un poste que si elle porte
aussi une quantité lisible.

**Un classeur, plusieurs feuilles.** Un métré réel se répartit souvent
en « Lot 01 », « Lot 02 »… plus un « Récapitulatif » qui **reprend les
mêmes codes**. L'interface liste les feuilles avec leur nombre de
postes et décoche d'office celles qui ressemblent à un récapitulatif —
une présomption tirée du nom, jamais une décision. Chaque poste retient
sa feuille, et le prix y retourne : tout écrire sur la première
rendrait au pouvoir adjudicateur un classeur incohérent sans qu'aucune
erreur ne soit levée. Un code vu deux fois est signalé, jamais
additionné.

**Rien n'est écarté en silence.** Une ligne du métré qu'on ne sait pas
lire doit **apparaître**, pas disparaître. La règle a été apprise à la
dure : une quantité écrite `=12.5*3` — ce qu'un pouvoir adjudicateur
fait couramment — était ignorée, le poste sortait du décompte, et le
rapport annonçait « tous les postes portent un prix » sur une offre
amputée de trois lignes. Le garde-fou certifiait l'inverse de la vérité.

D'où la **double ouverture du classeur** : sans `data_only`, une
cellule de quantité rend la formule et non son résultat ; avec, on
perdrait les formules du pouvoir adjudicateur à la sauvegarde. Les deux
besoins s'opposent, donc on lit deux fois — les formules pour la
structure, les valeurs pour les quantités calculées — et on n'écrit que
sur le premier classeur.

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

## Les données ne sont pas du code

Les tables vivent dans **`chiffrage/data/*.json`** : ressources,
ouvrages, composition, lots, mapping, devis historiques. Ce sont des
valeurs d'entreprise — prix d'achat, taux horaires, **rendements** —
que le chef d'entreprise est seul à connaître.

C'était le principal obstacle pratique à la calibration : les bons
chiffres finissaient dans un classeur Excel, et il fallait les recopier
à la main dans le Python. Du JSON se relit dans un diff GitHub, se
corrige sans éditeur de code, et ne s'exécute pas.

**Les commentaires qui portaient le raisonnement sont devenus des
données.** Un JSON est muet ; le *pourquoi* d'un chiffre serait mort
avec les commentaires Python. Il vit maintenant dans un champ `note` :

```json
{ "code_ouv": "40.40", "code_res": "MA.10", "qte_res": 2.2,
  "note": "2,2 m2 de membrane par m2 posé : bicouche,
           recouvrements et relevés compris." }
```

**Les tables sont contrôlées au chargement, pas au premier chiffrage.**
Éditable à la main veut dire corrompable à la main, et chacune de ces
fautes produirait un prix faux sans se voir dans un fichier de
150 lignes : une ressource orpheline vaut zéro, un ouvrage sans
composition se vend gratuitement, un type de ressource inconnu échappe
aux trois déboursés. Sept familles d'incohérence sont refusées, avec un
message qui nomme la faute.

**Et il n'y a pas de valeurs de repli**, à la différence du lexique ou
des paramètres : une bibliothèque vide ne dégraderait pas le résultat,
elle rendrait « aucun ouvrage » pour tous les postes — une offre
entièrement vide, présentée comme normale. L'outil refuse de démarrer.

`CHIFFRAGE_DATA` permet de charger d'autres tables que celles du dépôt.

---

## Contrôle des prix avant dépôt

Deux risques opposés : **trop bas**, l'offre est écartée pour prix
anormalement bas (art. 36 AR 18/04/2017) ou le chantier s'exécute à
perte ; **trop haut**, le marché est perdu.

Le pouvoir adjudicateur juge « anormalement bas » en comparant les
offres entre elles — comparaison inaccessible au moment de déposer,
puisque personne n'a les prix des concurrents. `controle_prix.py`
regarde donc l'offre **depuis l'intérieur de l'entreprise** : est-ce
que ce prix couvre ce que ce travail coûte ? C'est une autre
question, et la seule qu'on puisse trancher seul.

**L'indicateur central** est ce que l'offre laisse par heure de
main-d'œuvre, matériaux et matériel payés :

```
(montant encaissé − matériaux − matériel) / heures de main-d'œuvre
```

à comparer au plancher — coût horaire complet majoré des frais
généraux et de chantier, **marge et aléas exclus** : on cherche le
point où l'on cesse de gagner, pas celui où l'on vise.

D'où le **rabais maximal**, le chiffre qu'on veut connaître *avant* de
négocier et non après : la remise au-delà de laquelle l'offre ne
couvre plus ses coûts.

S'y ajoutent les alertes par poste — poids dans l'offre, prix dominé
par le rendement (donc par une estimation) ou par un prix d'achat
(donc par un fournisseur), écart à un marché antérieur, et part du
montant reposant sur des rendements jamais validés.

Le module ne fait **que de l'arithmétique**, délibérément : un
contrôle de prix doit être vérifiable à la main, pas cru sur parole.
Les seuils sont des repères de relecture, pas des règles de droit —
le seuil légal, la procédure et le délai de réponse figurent dans
l'AR en vigueur et le plus souvent dans le CSC lui-même.

### Dossier de justification

Si le pouvoir adjudicateur conteste un prix, il doit demander une
justification écrite avant d'écarter l'offre, et le délai est court.
L'interface produit le dossier : une lettre d'accompagnement à
relire et signer, puis un onglet par poste avec la décomposition qui
a **servi** à établir l'offre — ressources, quantités, prix d'achat,
déboursés par nature, coefficient. Ce ne sont pas des chiffres
reconstitués pour l'occasion, et c'est ce qui les rend crédibles.

---

## Mesurer l'appariement

```bash
python evaluation/mesurer_appariement.py       # -v pour le détail
```

`evaluation/epreuves_appariement.json` contient 58 libellés écrits en
vocabulaire de cahier spécial des charges, avec la bonne réponse.
Trois chiffres, et le troisième est le plus important :

| | |
|---|---:|
| **rang 1** — la bonne réponse arrive en tête | 49/50 · 98 % |
| **top 3** — elle est dans les trois premières, donc visible dans l'interface | 50/50 · 100 % |
| **silence** — sur les postes qu'aucun ouvrage ne couvre, l'outil se tait | 7/8 · 88 % |

Une suggestion confiante sur un poste qu'aucun ouvrage ne couvre est
**pire** que pas de suggestion : elle fait chiffrer un travail par un
autre. Un gain sur « rang 1 » payé par une perte sur « silence » n'est
donc pas un gain.

**Ces chiffres sont optimistes et il faut le savoir.** Les libellés
d'épreuve sont écrits par la même main que la bibliothèque, et le
lexique a été réglé en les regardant. Sur un échantillon témoin de 12
libellés écrits avant que le lexique n'existe et qui n'ont servi à aucun
réglage, l'appariement passe de 7/12 à 12/12 au premier rang — c'est le
gain réel du lexique, mesuré hors réglage. Le niveau sur un vrai cahier
des charges sera plus bas.

**Le jeu d'épreuve est fait pour être enrichi** : à chaque marché reçu,
y ajouter les libellés réels et leur bonne réponse. C'est ce qui rend la
mesure représentative, et ce qui permettra un jour de trancher
objectivement l'ajout d'embeddings sémantiques — en comparant sur le
même jeu, plutôt qu'au jugé.

Les seuils de `tests/test_evaluation.py` sont volontairement **sous** le
niveau mesuré : ils protègent d'une régression, ils ne certifient pas
une performance.

---

## État de la calibration

Les six devis forfaitaires historiques re-chiffrés avec la bibliothèque
(`python -m chiffrage calibration`) :

| Devis | Objet | Forfait vendu | Calculé | Écart |
|---|---|---:|---:|---:|
| 07 | Balcon avant (Av. E. Renan 35) | 2 500 € | 2 212 € | −11,5 % |
| 10 | Plafond côté route | 1 500 € | 1 565 € | +4,3 % |
| 11 | Jardin arrière + fenêtre 1er | 1 500 € | 1 667 € | +11,1 % |
| 13 | Cave + porte d'entrée | 930 € | 932 € | +0,2 % |
| 15 | Façade salle de bain | 3 650 € | 4 440 € | **+21,7 %** |
| 16 | Plafond + linteaux + isolation | 2 400 € | 3 424 € | **+42,7 %** |

Écart moyen absolu **15,2 %**. Cible : < 15 % sur **chaque** ligne.

Les devis 15 et 16 ressortent au-dessus du prix vendu. Deux hypothèses, non
tranchées faute de relevés :

1. les quantités de `METRES_HISTO` sont surestimées — elles viennent des
   descriptifs des devis PDF, **pas de relevés** ;
2. ces deux chantiers ont effectivement été vendus sous leur coût analytique.

Le devis 16 demande 37 h de main-d'œuvre au chiffrage analytique. À 2 400 €
HTVA matériaux compris, cela fait ~66 €/h : les frais généraux ne sont pas
couverts.

---

## Test d'intégration

`python -m chiffrage metre m.xlsx && python -m chiffrage offre m.xlsx o.xlsx` :

- 49 postes lus, **49 chiffrés automatiquement**, 0 sans prix
- **0 écart d'unité**
- 72 formules dans le fichier d'offre, toutes ancrées sur leur propre ligne
- les formules du pouvoir adjudicateur survivent au remplissage

L'offre est donc **régulière** au sens de l'art. 76 : plus aucun poste vide.
Les 13 postes qui restaient sans prix sont couverts depuis la création des
ouvrages `OUVRAGES_A_VALIDER` — dont les rendements restent à confirmer.

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
- [ ] Valider les rendements des **13 ouvrages de `OUVRAGES_A_VALIDER`** —
      ils comblent un trou qui rendait l'offre irrégulière, mais aucun n'a
      jamais été confronté à un chantier réel
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

145 tests. Ceux de la chaîne Excel se skippent sans openpyxl, ceux de
l'interface sans streamlit. L'écriture GitHub est testée avec un
dépôt simulé — la suite ne touche jamais au réseau.

L'interface est testée par `AppTest` de Streamlit, qui exécute vraiment le
script : dépôt d'un métré, appariement, clic sur « Chiffrer », offre
téléchargeable. Il a déjà attrapé deux défauts qu'aucune relecture n'aurait
vus — un `icon=` refusé par Streamlit qui faisait planter l'app au
chargement, et un `st.stop()` qui arrêtait le script **entier**, si bien
que trois onglets sur quatre ne s'affichaient jamais. Un test vérifie
aussi que le total affiché à l'écran égale celui du moteur : deux vérités
sur un prix, c'est une offre fausse tôt ou tard.

Trois d'entre eux portent sur le notebook Colab : JSON valide, cellules qui
compilent, et surtout **chaque `from chiffrage.x import y` du notebook doit
résoudre**. Rien n'exécute le notebook en CI — sans ce dernier test, un
renommage dans `chiffrage/` le casserait en silence, et ça ne se verrait
qu'au moment d'ouvrir Colab pour répondre à un marché.
