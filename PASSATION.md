# Passation — outil de chiffrage BAG BATTER

Ce fichier existe pour **reprendre le projet ailleurs** : autre conversation,
autre assistant, autre développeur. Il donne le contexte, l'état réel, les
décisions structurantes et — surtout — les pièges déjà payés. Le `README.md`
documente l'outil ; celui-ci documente *le travail*.

Dernière mise à jour : 30 août 2026 · 46 commits · 214 tests au vert.

---

## 1. Le besoin

**BAG BATTER SRL** — Ronkel 18, 1780 Wemmel, TVA BE 0766.637.025. Entreprise
générale de rénovation : façades, étanchéité, plafonnage, isolation, peinture,
menuiserie extérieure, sanitaire léger.

L'entreprise répond à des **marchés publics belges**. Un pouvoir adjudicateur
impose un métré — une liste de postes numérotés, avec leurs quantités — et
attend un **prix unitaire pour chaque poste**. Pas un forfait global.

Deux conséquences juridiques qui commandent tout le reste :

- **AR 18/04/2017, art. 76** — un poste sans prix rend l'offre **irrégulière**.
  Elle est écartée sans examen. D'où l'obsession de la couverture : mieux vaut
  un prix estimé et signalé comme tel qu'une case vide.
- **AR 18/04/2017, art. 36** — un prix jugé anormalement bas doit pouvoir être
  **justifié**. D'où la décomposition conservée poste par poste, et le dossier
  de justification exportable.

Avant cet outil, les prix vivaient dans la tête du chef d'entreprise et dans
des classeurs Excel divergents.

---

## 2. Où vit le projet

| Quoi | Où |
|---|---|
| Dépôt | `github.com/pmeyssonnier/Devis-generator`, branche `main` |
| App en ligne | `https://devis-generator.streamlit.app/` |
| Hébergement | Streamlit Community Cloud, Python 3.14.7, redéploiement auto au push |
| Accès | liste d'e-mails autorisés (*Settings → Sharing*) |
| Secrets | *Settings → Secrets*, jamais dans le dépôt |

Le dépôt est **public**. Ce qui ferme l'accès à l'app, c'est la liste
d'e-mails — pas la configuration d'affichage, qui n'est que cosmétique.

---

## 3. Le modèle de prix

```
debourse_sec = Σ (qte_res × pu_res)        pour chaque ressource de l'ouvrage
pu_vente     = debourse_sec × K
K            = (1+FG) (1+FC) (1+aléas) (1+marge)
```

Valeurs en vigueur (`chiffrage/parametres.py`, réglables depuis l'app) :

| | |
|---|---|
| FG — frais généraux | 0,12 |
| FC — frais de chantier | 0,05 |
| aléas | 0,03 |
| marge | 0,10 |
| **K** | **1,3324** |
| TVA privé (logement > 10 ans) | 6 % |
| TVA marché public | 21 % |

Trois natures de poste dans un métré : **QF** (quantité forfaitaire), **QP**
(quantité présumée), **FF** (forfait).

---

## 4. État réel des données — à lire avant de faire confiance à un prix

| Table | Volume |
|---|---|
| Ressources | 49 |
| Ouvrages | 49 |
| Lignes de composition | 153 |
| Lots | 9 |
| Devis historiques (calibration) | 6 |
| Rendements jamais validés | 13 |

**La bibliothèque a été reconstruite à partir d'un document texte, pas
relevée sur chantier.** C'est le fait le plus important de ce projet.

Calibration actuelle — écart entre le forfait réellement vendu et ce que la
bibliothèque recalcule, écart moyen absolu **15,3 %** :

| Devis | Objet | Écart |
|---|---|---|
| 07 | Balcon avant — Av. Ernest Renan 35 | −11,5 % |
| 10 | Plafond côté route | +4,4 % |
| 11 | Jardin arrière + fenêtre 1er étage | +11,2 % |
| 13 | Cave + porte d'entrée | +0,3 % |
| 15 | Façade salle de bain | **+21,7 %** |
| 16 | Plafond + linteaux + isolation | **+42,7 %** |

Les devis 15 et 16 sont hors cible et l'hypothèse n'est pas tranchée :
sous-tarification à l'époque, ou surestimation des quantités aujourd'hui.

Le ⚠️ « à valider » sur 13 ouvrages ne dit **pas** « prix faux » : il dit
« rendement jamais confronté à un chantier réel ». Ils ont été inventés pour
couvrir des postes qui seraient restés sans prix (art. 76). **L'absence de ⚠️
ne veut pas dire validé — seulement non signalé** : les 36 autres viennent de
la même documentation reconstruite.

**Le vrai blocage du projet n'est pas technique.** C'est une séance avec le
chef d'entreprise pour relever les rendements et les taux horaires réels.
Tout le reste est prêt à les recevoir.

---

## 5. Architecture

Moteur en **Python pur** (aucune dépendance), interface et Excel au-dessus.

| Module | Rôle |
|---|---|
| `bibliotheque.py` | chargeur validant des tables JSON — refuse de démarrer au moindre défaut |
| `moteur.py` | bordereau, devis, fiche de prix, calibration, cohérence |
| `parametres.py` | identité entreprise + coefficients, réglables depuis l'app |
| `suggestion.py` | appariement poste imposé → ouvrage, par le libellé |
| `lexique.py` | vocabulaire de CSC → vocabulaire de la bibliothèque |
| `metre_io.py` | lecture d'un métré reçu, remplissage de l'offre |
| `detection_colonnes.py` | quelle colonne est le code, la quantité, le prix |
| `controle_prix.py` | relecture avant dépôt : couverture, rabais maximal |
| `justification_xlsx.py` | dossier de justification de prix (art. 36) |
| `devis_xlsx.py` · `export_xlsx.py` | classeurs, **avec formules vivantes** |
| `devis_json.py` | enregistrer un devis et le reprendre pour le modifier |
| `depot_github.py` | écriture des tables corrigées dans le dépôt (API GitHub, urllib) |
| `gen_metre.py` | métré fictif de 49 postes pour s'entraîner |
| `data/*.json` | les tables elles-mêmes |

`streamlit_app.py` (~1 500 lignes) — six onglets :

1. **📥 Répondre à un métré** — dépôt du fichier, appariement, offre remplie
2. **🧾 Devis client** — devis privé, TVA 6/21 %, export `.xlsx` + `.json`
3. **📚 Bibliothèque** — recherche, fiche de justification, atelier de
   correction (avec la calculette de rendement : quantité + heures → h/unité)
4. **🔤 Lexique** — banc d'essai de l'appariement, ajout de termes
5. **🎯 Calibration** — les 6 devis historiques et leurs écarts
6. **⚙️ Paramètres** — entreprise et coefficients

Qualité d'appariement mesurée sur `evaluation/` (58 libellés) :
**98 % au rang 1 · 100 % dans les trois premiers · silence sur 7 des 8 postes
qu'aucun ouvrage ne couvre.** Le jeu d'épreuves a été écrit après coup :
il y a du surapprentissage, le chiffre est un garde-fou contre les
régressions, pas une mesure de généralisation.

---

## 6. Décisions structurantes — le *pourquoi*

**Les données sont du JSON, jamais du Python généré.** Le contenu vient de
champs de saisie, et l'app exécute ce qu'elle commite : écrire du code depuis
une saisie serait une injection. Un JSON absurde reste un JSON.

**Les commentaires qui portaient le raisonnement sont devenus un champ
`note`.** Un JSON est muet ; sans ça, le *pourquoi* d'un chiffre mourait avec
les commentaires Python.

**Le chargeur n'a aucune valeur de repli.** Une bibliothèque vide produirait
silencieusement des offres à zéro. Il lève, l'app ne démarre pas. Exception
assumée : `ouvrages_a_valider`, qui ne porte aucun prix — voir §7.

**L'unité est éliminatoire, l'opération est pénalisée.** Un poste au mètre
courant ne se voit jamais proposer un ouvrage au m² : le prix serait faux
d'un facteur inconnu et ça ne se verrait qu'à la facturation. En revanche
« dépose » vs « pose » est *pénalisé* (×0,35) et non éliminatoire : l'unité
est déclarée dans le métré, l'opération est inférée de mots, et une inférence
fausse ne doit pas faire disparaître un candidat valable.

**Un score n'est pas une probabilité.** C'est une ressemblance de libellés.
Il dit « regarde ici d'abord », jamais « c'est celui-là ».

**Les codes `xx.xx` sont internes.** `lot . position`, par pas de dix pour
laisser intercaler. **Rien n'est standardisé** : le pouvoir adjudicateur
numérote son métré comme il veut, d'où une correspondance refaite à chaque
marché. La colonne `code_ref` de `OUVRAGES` est prévue pour une vraie
référence (CCT 2022 Bruxelles / CCT-B Qualiroute / SB 250) — elle est vide.

**On stocke un code, on affiche un libellé.** Le libellé porte le prix du
jour et change ; le code est la clé.

**L'interface ne recalcule jamais un prix pour son compte.** Deux vérités sur
un prix, c'est une offre fausse tôt ou tard. Un test compare le total à
l'écran avec celui du moteur.

---

## 7. Pièges déjà payés — ne pas les repayer

**openpyxl**
- ne jamais ouvrir en `data_only=True` pour écrire : ça détruit les formules
- `insert_rows` ne décale pas les formules — ancrer ligne par ligne
- lire une quantité issue d'une formule demande **deux ouvertures** du
  classeur ; sans ça des postes disparaissent en silence
- et la seconde ouverture doit viser la colonne **détectée**, pas celle de
  la disposition par défaut : la quantité trouvée en G, son cache cherché
  en F, et le poste ressortait « quantité illisible » alors que la valeur
  était là. Un poste sans prix rend l'offre irrégulière (art. 76)

**Streamlit**
- `st.cache_data` sans clé sur une fonction qui dépend des paramètres : le
  bug le plus cher du projet — **185 308 € affichés à l'écran contre
  210 581 € dans le fichier produit**, soit 13,6 %. Le calcul coûte 0,11 ms :
  il n'y avait rien à cacher.
- `st.stop()` dans un onglet arrête le **script entier** — les onglets
  suivants ne s'affichent jamais. Extraire en fonction et `return`.
- un `st.rerun()` inconditionnel après un `file_uploader` **boucle sans fin** :
  le fichier reste déposé à chaque réexécution. Repérer le dépôt par
  l'empreinte SHA-256 de son contenu.
- un avertissement affiché juste avant un `st.rerun()` est **effacé avant
  d'être lu**. Le ranger dans l'état, l'afficher après.
- `st.column_config.NumberColumn(format="%.1f %%")` ne multiplie pas par 100 :
  des écarts s'affichaient à −0,1 % au lieu de −11,5 %.
- un widget à clé garde SA valeur d'une réexécution à l'autre. Remettre la
  valeur dans l'état **avant** d'instancier le widget, et ne pas passer
  `value=` en plus d'une clé. Le prix à payer, mesuré : dans l'atelier,
  choisir MO.04 (45 €) après MO.01 (59 €) laissait 59 € dans la case, et
  « Appliquer » sans rien taper **écrivait 59 € sur MO.04** — un prix faux
  en un geste, sans rien à l'écran pour le dire. Repère de remise à jour :
  le couple (ce qui est choisi, sa valeur enregistrée), pour qu'un retour
  aux valeurs d'origine remette la case d'aplomb sans effacer une saisie
  en cours.
- `st.secrets` **lève** s'il n'y a pas de fichier de secrets. Envelopper.
- `st.data_editor` n'est **pas exposé par `AppTest`** : ses options ne se
  testent pas depuis l'app en marche, il faut attaquer les fonctions.
- le menu d'un `SelectboxColumn` est une grille : il **ne s'ajuste pas à la
  largeur de l'écran**. Des libellés de 80 caractères débordaient à droite et
  se retrouvaient rognés *à gauche* sur un téléphone — le code, seule partie
  non ambiguë, disparaissait. Plafond : 42 caractères.

**Streamlit Cloud — le piège le plus coûteux**

Un `git push` déclenche `🔄 Updated app!` **sans redémarrer le processus**
(pas de `Shutting down` dans les logs). Survivent alors à la mise à jour :

- l'`st.session_state` des sessions ouvertes ;
- **les modules déjà importés**.

L'app peut donc tourner sur du code plus récent que les données qu'elle lit.
Ajouter une clé à une table est un **changement de schéma**, pas un ajout
anodin. Ça a mis un `KeyError` en plein écran, chez l'utilisateur, huit fois
de suite. Parade en deux moitiés, les deux nécessaires :

1. les tables en session sont **complétées** depuis la bibliothèque au
   démarrage de l'atelier — sans écraser les corrections en cours ;
2. les lectures de la table optionnelle passent par un accès qui **ne peut pas
   lever**. Elle ne porte aucun prix : liste vide = repli honnête. **Ce
   raisonnement ne s'étend pas à une table de prix.**

Même exposition pour une **fonction ajoutée à un module** : le script
réexécuté fait `from chiffrage.moteur import ...` sur l'objet module déjà
en mémoire, qui ne la connaît pas — `ImportError` en plein écran. Il n'y a
pas de parade en code qui vaille : un repli silencieux ferait tourner
l'app sur l'ancien comportement sans le dire.

En cas d'écran rouge après un push : **Manage app → Reboot app**.

**Ce qui s'écrit et ce qui se lit ne se replient pas pareil.** Un champ
qu'on LIT peut se replier sur une valeur par défaut : si elle est fausse,
la lecture produit une anomalie visible. Le champ qu'on ÉCRIT — la
colonne du PU, la seule — ne le peut pas : un repli faux met le prix
par-dessus une autre colonne du pouvoir adjudicateur, et personne ne le
voit. D'où la règle : **PU inconnu, on n'écrit pas**, les postes
ressortent sans prix avec leur montant à porter à la main. Refuser se
voit, se tromper de colonne, non.

**Ergonomie mobile** — le chef d'entreprise travaille au téléphone.
Un `st.data_editor` de 49 lignes est inutilisable au doigt : taper dans une
cellule, valider, en sortir — chaque geste rate une fois sur deux, et une
correction qui ne « prend » pas ne se voit pas. Remplacé partout par
**trois gestes sûrs : choisir, saisir, appliquer**.

**Tests** — deux fois j'ai figé dans un test une valeur que l'app écrit en
production (un terme de lexique, puis des écarts de calibration) : le test
cassait dès que l'utilisateur s'en servait. **Asserter des relations, pas des
littéraux.**

**Tester les combinaisons, pas seulement les cas** — les deux défauts
ci-dessus ont survécu à une suite qui couvrait chacune de leurs moitiés :
« colonne déplacée » passait, « quantité en formule » passait, les deux
ensemble perdaient le poste. De même, des colonnes numériques existaient
sans qu'aucune n'atteigne le seuil, et `min()` levait sur une séquence
vide. D'où un **classeur de torture** dans la suite : plusieurs feuilles
+ colonnes déplacées + quantité en formule + récapitulatif, d'un seul
tenant. Un audit du dépôt les a trouvés ; les tests isolés, non.

---

## 8. Sécurité — contraintes en vigueur

- Le jeton d'écriture est un **PAT fine-grained limité au seul dépôt
  `Devis-generator`**, permission *Contents: read and write*, avec expiration.
  **Jamais** un jeton classique à portée `repo` : celui-là donne l'écriture
  sur tous les dépôts du compte.
- Les secrets vivent dans *Settings → Secrets* de Streamlit Cloud.
  `.gitignore` exclut `.streamlit/secrets.toml` ; seul le `.example` est suivi.
- Les termes de lexique saisis sont bornés à l'entrée :
  `^[a-z0-9][a-z0-9 '-]{0,48}$`.
- Le commit de lexique **relit et fusionne** le fichier distant ; celui d'une
  table écrase, mais après relecture du `sha` — l'écriture échoue si quelqu'un
  est passé entre temps.

---

## 9. Tests

```bash
pip install -r requirements-dev.txt
pytest
```

**214 tests.** Excel se skippe sans openpyxl, l'interface sans streamlit.
L'écriture GitHub est testée avec un dépôt simulé : la suite ne touche jamais
au réseau. L'interface est testée par `AppTest`, qui **exécute vraiment le
script**. Trois tests portent sur le notebook Colab, dont un qui vérifie que
chaque `from chiffrage.x import y` résout — rien n'exécute le notebook en CI.

`ruff` doit passer. `conftest.py` force `STREAMLIT_CLIENT_SHOW_ERROR_DETAILS=full`,
sinon la configuration client masque la cause des échecs.

---

## 10. Ce qui reste, par ordre d'importance

1. **Séance de calibration avec le chef d'entreprise** — taux horaires et
   rendements réels. Tout le reste en dépend. L'atelier sait maintenant
   recevoir un relevé brut (quantité faite, personnes, durée) et non plus
   seulement un rendement déjà calculé ; ce qui manque encore, c'est de
   **garder le relevé** — aujourd'hui lever un ⚠️ n'enregistre ni la date,
   ni le chantier, ni les heures. Une table `releves.json` est le pas
   suivant, et elle bute sur une question qui n'est pas technique : le
   statut devient calculable (0 relevé = jamais confronté), ce qui ferait
   passer les ⚠️ de 13 à 49. Plus vrai, peut-être intenable
   commercialement — au chef d'entreprise de trancher.
2. Surfaces réelles des 6 chantiers historiques, puis viser < 15 % d'écart
   sur **chaque** ligne.
3. Trancher l'hypothèse sur les devis 15 et 16.
4. Valider les 13 rendements marqués ⚠️, un chantier à la fois.
5. Remplir `code_ref` au premier cahier des charges reçu, et en faire un
   second axe d'appariement : une référence CCT donne une correspondance
   **exacte**, là où le libellé ne donne qu'un score.
6. Différé volontairement : extraction du CSC par un LLM, recoupement
   clause × composition, embeddings (seulement si l'évaluation plafonne),
   découpage UI/données, `Decimal` à la place des flottants.

---

## 11. Pour reprendre la conversation ailleurs

Fournir ce fichier, puis :

> Voici le contexte d'un outil de chiffrage pour une entreprise de rénovation
> belge qui répond à des marchés publics. Le code est sur
> `github.com/pmeyssonnier/Devis-generator`, l'app sur
> `devis-generator.streamlit.app`. Lis `PASSATION.md` puis `README.md`.

Ce qu'il faut garder à l'esprit en reprenant :

- **les prix ne sont pas calibrés** — le dire chaque fois qu'un montant est
  montré, l'app le fait déjà ;
- **ne jamais inventer un chiffre en silence** — si une valeur manque, le
  signaler plutôt que de la deviner ; c'est la règle qui a tenu tout du long ;
- **l'utilisateur travaille au téléphone** — toute interface se juge au doigt ;
- **un push touche des sessions déjà ouvertes** — voir §7.
