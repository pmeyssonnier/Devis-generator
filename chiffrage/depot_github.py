"""
╔══════════════════════════════════════════════╗
║  ÉCRITURE D'UN FICHIER DANS LE DÉPÔT GITHUB                          ║
╚══════════════════════════════════════════════╝

Sert au bouton « commiter » du lexique : rendre permanent un terme
appris, sans quitter l'interface ni écrire de Python.

Des fichiers de DONNÉES uniquement — le lexique, les paramètres, les
tables — et **du JSON, jamais du code** :
le contenu vient de champs de saisie, et écrire du Python à partir
d'une saisie serait une injection.

──────────────────────────────────────────────
LE JETON
──────────────────────────────────────────────
Un PAT **fine-grained**, limité au SEUL dépôt Devis-generator,
permission « Contents: read and write », avec une date d'expiration.

**Pas** un jeton classique à portée `repo` : celui-là donne l'écriture
sur TOUS les dépôts du compte, pour une fonction qui n'a besoin que
d'un fichier. Si ce jeton fuite, ce qu'il permet doit rester borné à
ce dépôt-ci.

Il vit dans les secrets Streamlit (Settings -> Secrets), jamais dans
le dépôt :

    [github]
    token = "github_pat_..."
    depot = "pmeyssonnier/Devis-generator"
    branche = "main"

Aucune dépendance : urllib de la bibliothèque standard suffit, et
n'ajoute rien à installer sur l'hébergement.
"""

import base64
import json
import urllib.error
import urllib.request

API = "https://api.github.com"
DELAI = 20

# Disposition historique : une seule entreprise, ses fichiers là où ils
# sont nés. Elle reste le défaut — la déplacer casserait le déploiement
# en service.
DOSSIER_TABLES = "chiffrage/data"
CHEMIN_PARAMETRES = "chiffrage/parametres_local.json"
CHEMIN_LEXIQUE = "chiffrage/lexique_local.json"


def chemins_entreprise(dossier=None):
    """
    Où vivent, DANS LE DÉPÔT, les fichiers propres à une entreprise.

    Sans `dossier`, la disposition historique. Avec, tout est réuni au
    même endroit — et c'est la condition d'une instance par entrepreneur :
    ses prix, son identité et son lexique dans SON dossier, écrits par un
    jeton limité à SON dépôt.

    Le pendant côté lecture est la variable d'environnement
    CHIFFRAGE_DATA, qui doit désigner le même dossier. Les deux se
    règlent ensemble : écrire dans un dossier que l'app ne lit pas
    perdrait les corrections sans rien dire.
    """
    if not dossier:
        return {"tables": DOSSIER_TABLES,
                 "parametres": CHEMIN_PARAMETRES,
                 "lexique": CHEMIN_LEXIQUE}
    d = str(dossier).strip("/")
    return {"tables": d,
             "parametres": f"{d}/parametres_local.json",
             "lexique": f"{d}/lexique_local.json"}


class ErreurDepot(RuntimeError):
    """Échec d'une opération GitHub, avec un message lisible."""


def _appel(methode, url, token, corps=None, delai=DELAI):
    donnees = json.dumps(corps).encode("utf-8") if corps is not None else None
    requete = urllib.request.Request(
        url,
        data=donnees,
        method=methode,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "Content-Type": "application/json",
            "User-Agent": "devis-generator",
        },
    )
    with urllib.request.urlopen(requete, timeout=delai) as reponse:
        return json.loads(reponse.read().decode("utf-8") or "{}")


def lire_fichier(chemin, depot, token, branche="main", _appel=_appel):
    """
    Lit un fichier du dépôt.

    Retourne (contenu_texte, sha) — et (None, None) si le fichier
    n'existe pas encore, ce qui est un cas normal au premier commit.
    """
    url = f"{API}/repos/{depot}/contents/{chemin}?ref={branche}"
    try:
        reponse = _appel("GET", url, token)
    except urllib.error.HTTPError as err:
        if err.code == 404:
            return None, None
        raise ErreurDepot(_expliquer(err, depot)) from err
    except urllib.error.URLError as err:
        raise ErreurDepot(f"GitHub injoignable : {err.reason}") from err

    contenu = base64.b64decode(reponse.get("content", "")).decode("utf-8")
    return contenu, reponse.get("sha")


def ecrire_fichier(chemin, contenu, message, depot, token,
                    branche="main", sha=None, _appel=_appel):
    """
    Crée ou met à jour un fichier, et rend l'URL du commit.

    `sha` est celui de la version qu'on croit remplacer : GitHub
    refuse l'écriture s'il ne correspond plus. C'est ce qui empêche
    d'écraser en silence l'ajout de quelqu'un d'autre — d'où la
    relecture juste avant l'écriture, côté appelant.
    """
    corps = {
        "message": message,
        "content": base64.b64encode(contenu.encode("utf-8")).decode("ascii"),
        "branch": branche,
    }
    if sha:
        corps["sha"] = sha

    url = f"{API}/repos/{depot}/contents/{chemin}"
    try:
        reponse = _appel("PUT", url, token, corps)
    except urllib.error.HTTPError as err:
        raise ErreurDepot(_expliquer(err, depot)) from err
    except urllib.error.URLError as err:
        raise ErreurDepot(f"GitHub injoignable : {err.reason}") from err

    return (reponse.get("commit") or {}).get("html_url", "")


def _expliquer(err, depot):
    """Traduit un code HTTP en une phrase qui dit quoi faire.

    Un « 403 » brut dans l'interface n'apprend rien à personne ; la
    cause est presque toujours la portée du jeton.
    """
    causes = {
        401: "jeton invalide ou expiré — en régénérer un dans "
              "GitHub, Settings > Developer settings > "
              "Personal access tokens.",
        403: f"jeton refusé sur {depot} — vérifier qu'il est "
              f"fine-grained, qu'il vise CE dépôt et qu'il a la "
              f"permission « Contents: read and write ».",
        404: f"dépôt {depot} introuvable — nom exact attendu sous la "
              f"forme « proprietaire/depot », et jeton ayant accès à "
              f"ce dépôt (un dépôt privé invisible du jeton répond 404, "
              f"pas 403).",
        409: "conflit : le fichier a changé entre-temps. Recharger "
              "la page et refaire l'ajout.",
        422: "requête refusée par GitHub — la branche existe-t-elle ?",
    }
    detail = causes.get(err.code, f"HTTP {err.code}")
    return f"Écriture impossible : {detail}"


def commiter_lexique(contenu, depot, token, branche="main",
                      chemin=None,
                      message="feat(lexique): termes appris depuis l'interface",
                      _lire=lire_fichier, _ecrire=ecrire_fichier):
    """
    Écrit le lexique appris, en fusionnant avec ce qui est déjà commité.

    Retourne (contenu_fusionne, url_du_commit).

    La relecture juste avant l'écriture n'est pas une précaution de
    style : deux personnes peuvent régler le lexique le même jour, et
    un PUT sans fusion effacerait les termes de l'autre sans rien dire.
    """
    from .lexique import fusion_a_commiter

    chemin = chemin or CHEMIN_LEXIQUE
    texte_distant, sha = _lire(chemin, depot, token, branche)
    distant = {}
    if texte_distant:
        try:
            distant = json.loads(texte_distant)
        except ValueError as err:
            raise ErreurDepot(
                f"{chemin} est présent dans le dépôt mais illisible : "
                f"{err}. Le corriger à la main avant de commiter d'ici."
            ) from err

    fusion = fusion_a_commiter(distant)
    texte = json.dumps(fusion, ensure_ascii=False, indent=2,
                        sort_keys=True) + "\n"
    url = _ecrire(chemin, texte, message, depot, token, branche, sha)
    return fusion, url


def commiter_parametres(contenu, depot, token, branche="main",
                         chemin=None,
                         message="chore(parametres): réglés depuis l'interface",
                         _lire=lire_fichier, _ecrire=ecrire_fichier):
    """
    Écrit l'identité et les coefficients réglés dans l'interface.

    Retourne l'URL du commit.

    À la différence du lexique, il n'y a RIEN à fusionner : les
    paramètres forment un jeu cohérent, et deux versions ne
    s'additionnent pas — une adresse ne se mélange pas à une autre. Le
    dernier qui écrit a raison, mais pas en aveugle : le `sha` relu
    juste avant fait échouer l'écriture si quelqu'un est passé entre
    temps, plutôt que d'écraser en silence.
    """
    chemin = chemin or CHEMIN_PARAMETRES
    _, sha = _lire(chemin, depot, token, branche)
    return _ecrire(chemin, contenu, message, depot, token, branche, sha)


def commiter_validations(validations, depot, token, branche="main",
                          chemin=None, dossier=None,
                          message="data(validations): rendements validés "
                                   "depuis l'interface",
                          _lire=lire_fichier, _ecrire=ecrire_fichier):
    """
    Ajoute des validations de rendement, en FUSIONNANT avec le distant.

    Même sémantique que le journal des relevés, et pour la même raison :
    une validation s'ajoute à l'histoire d'un rendement, elle ne la
    remplace pas. Un rendement validé en mars puis revalidé en septembre
    garde ses deux bulletins — c'est la seule trace de ce qui a bougé.
    """
    from .moteur import fusionner_validations

    chemin = chemin or f"{chemins_entreprise(dossier)['tables']}/validations.json"
    texte_distant, sha = _lire(chemin, depot, token, branche)
    distant = []
    if texte_distant:
        try:
            distant = json.loads(texte_distant)
        except ValueError as err:
            raise ErreurDepot(
                f"{chemin} est présent dans le dépôt mais illisible : "
                f"{err}. Le corriger à la main avant de commiter d'ici."
            ) from err
    if not isinstance(distant, list):
        raise ErreurDepot(
            f"{chemin} n'est pas une liste de validations. Le corriger à "
            f"la main avant de commiter d'ici.")

    fusion = fusionner_validations(distant, validations)
    texte = json.dumps(fusion, ensure_ascii=False, indent=2) + "\n"
    url = _ecrire(chemin, texte, message, depot, token, branche, sha)
    return fusion, url


def commiter_releves(releves, depot, token, branche="main",
                      chemin=None, dossier=None,
                      message="data(releves): relevés de chantier ajoutés "
                               "depuis l'interface",
                      _lire=lire_fichier, _ecrire=ecrire_fichier):
    """
    Ajoute des relevés de chantier, en FUSIONNANT avec le journal distant.

    Retourne (journal_fusionne, url_du_commit).

    C'est la seule table qui ne s'écrase pas. Une table de prix est un
    tout cohérent — deux versions ne s'additionnent pas, et le dernier
    qui écrit gagne. Un journal de chantier est l'inverse : il
    s'AJOUTE. Deux téléphones peuvent enregistrer le même soir, et un
    PUT sans fusion effacerait l'observation de l'autre sans rien dire
    — une demi-journée de chantier perdue, invisible.
    """
    from .moteur import fusionner_releves

    chemin = chemin or f"{chemins_entreprise(dossier)['tables']}/releves.json"
    texte_distant, sha = _lire(chemin, depot, token, branche)
    distant = []
    if texte_distant:
        try:
            distant = json.loads(texte_distant)
        except ValueError as err:
            raise ErreurDepot(
                f"{chemin} est présent dans le dépôt mais illisible : "
                f"{err}. Le corriger à la main avant de commiter d'ici."
            ) from err
    if not isinstance(distant, list):
        raise ErreurDepot(
            f"{chemin} n'est pas une liste de relevés. Le corriger à la "
            f"main avant de commiter d'ici.")

    fusion = fusionner_releves(distant, releves)
    texte = json.dumps(fusion, ensure_ascii=False, indent=2) + "\n"
    url = _ecrire(chemin, texte, message, depot, token, branche, sha)
    return fusion, url


def commiter_table(nom, contenu, depot, token, branche="main",
                    message=None, dossier=None,
                    _lire=lire_fichier, _ecrire=ecrire_fichier):
    """
    Écrit une table de `chiffrage/data/` corrigée depuis l'interface.

    Comme pour les paramètres, RIEN NE SE FUSIONNE : une table est un
    tout cohérent, et deux versions ne s'additionnent pas — un prix
    corrigé ne se mélange pas à un autre prix corrigé. Le dernier qui
    écrit gagne, mais pas en aveugle : le `sha` relu juste avant fait
    échouer l'écriture si quelqu'un est passé entre temps.
    """
    chemin = f"{chemins_entreprise(dossier)['tables']}/{nom}.json"
    _, sha = _lire(chemin, depot, token, branche)
    return _ecrire(
        chemin, contenu,
        message or f"data({nom}): corrigé depuis l'interface",
        depot, token, branche, sha)
