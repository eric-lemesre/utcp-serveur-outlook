# Fournisseur UTCP `outlook`

Service FastAPI qui publie sur `GET /utcp` un manuel UTCP décrivant la partie
**Outlook de Microsoft Graph** : courrier, dossiers, agenda, contacts, personnes
et paramètres de boîte, pour l'utilisateur connecté (`/me`). Un agent qui charge
ce manuel appelle ensuite Graph **directement**, sans serveur MCP intermédiaire.

Le détail des opérations couvertes, des portées requises et des écarts est
dans `utcp/COUVERTURE.md`.

## Forme « manuel seul »

Le service n'a **aucune route métier** : il sert `GET /utcp` et `GET /health`,
rien d'autre. Les URL d'outils du manuel visent Microsoft Graph
(`${GRAPH_BASE}`, rendu au démarrage avec `OUTLOOK_GRAPH_BASE`), pas ce service.

### Pourquoi Graph impose OAuth

- Microsoft Graph n'accepte **ni clé d'API ni auth basic** : chaque appel porte
  un jeton d'accès OAuth 2.0 émis par Microsoft Entra ID
  (`Authorization: Bearer <jeton>`).
- Pour lire la boîte d'**une personne**, le jeton doit être **délégué** : il
  est obtenu par cette personne elle-même (connexion interactive) et ne donne
  que les droits qu'elle a consentis (scopes) et qu'elle possède déjà.
- Les jetons **d'application** (client credentials) donnent accès à toutes
  les boîtes du locataire : hors périmètre, volontairement.

### Pourquoi le serveur ne porte pas le jeton

- Le manuel est public par nature (il décrit l'API publique de Graph) et ne
  contient aucun secret : le jeton y reste une variable,
  `Bearer ${OUTLOOK_GRAPH_TOKEN}` dans `auth.api_key` (en-tête `Authorization`).
- C'est le **client UTCP** qui substitue la variable et appelle Graph : le jeton
  ne transite **jamais** par ce service. Il ne le lit pas, même s'il traîne
  dans son environnement (aucun réglage ne le déclare).
- Un même service peut donc servir plusieurs utilisateurs : chacun garde son
  propre jeton délégué chez lui.
- Pas de clé d'API sur ce service, faute de route métier à protéger (comme
  `atlassian`).

## Fichiers

| Fichier | Rôle |
|---|---|
| `utcp/manual.json` | manuel source ; `${GRAPH_BASE}` (et `${BASE_URL}` s'il apparaît) rendus au démarrage |
| `utcp/COUVERTURE.md` | opérations couvertes, portées et écarts |
| `utcp/validation.vars` | variables rendues avant validation par `tools/valider_tous.sh` du parent |
| `serveur_utcp_outlook/__main__.py` | seul endroit qui configure `logging`, lit les réglages, lance uvicorn ; refuse de démarrer si réglages ou manuel invalides |
| `serveur_utcp_outlook/main.py` | `creer_application()` : avertissement de lecture seule, `X-Request-Id` lu ou généré et renvoyé, une ligne de journal par requête sur la route générique |
| `serveur_utcp_outlook/config.py` | réglages `OUTLOOK_*` (pydantic-settings) ; URL HTTPS (http seulement vers loopback), sans identifiant ni slash final |
| `serveur_utcp_outlook/utcp.py` | `rendre()` : rendu, contrôles, filtre lecture seule ; routeur `/utcp` |
| `serveur_utcp_outlook/routes.py` | `/health` |
| `outils/obtenir_jeton.py` | jeton délégué Graph par device code (MSAL), au format `client/.env` |
| `tests/test_utcp_manual.py` | contrat : `/utcp` public, JSON, cache ; marqueurs rendus ; tags ; lecture seule ; aucun jeton ; corrélation ; refus au démarrage |
| `tests/test_obtenir_jeton.py` | arguments et scopes du script de jeton, sans réseau |

## Réglages du service

| Variable | Défaut | Rôle |
|---|---|---|
| `OUTLOOK_GRAPH_BASE` | `https://graph.microsoft.com/v1.0` | racine de Graph rendue à la place de `${GRAPH_BASE}` ; HTTPS, ou http vers loopback (tests) |
| `OUTLOOK_URL_PUBLIQUE` | — | URL publique de ce service ; exigée seulement si le manuel cite `${BASE_URL}` |
| `OUTLOOK_LECTURE_SEULE` | `false` | `true` : outils `ecriture` retirés du manuel servi |
| `OUTLOOK_ECOUTE_HOTE` | `127.0.0.1` | adresse d'écoute |
| `OUTLOOK_ECOUTE_PORT` | `8082` | port d'écoute |
| `OUTLOOK_JOURNAL` | `INFO` | niveau du journal |
| `OUTLOOK_JOURNAL_FORMAT` | `texte` | `texte` ou `json` (stderr) |

Au démarrage, `rendre()` remplace `${GRAPH_BASE}` et `${BASE_URL}`, puis
refuse de démarrer, avec un message clair, si :

- `utcp_version` manque ou `tools` est vide ;
- il reste un `${…}` autre que `${OUTLOOK_GRAPH_TOKEN}` et `${OUTLOOK_GRAPH_AUTHORIZATION}` ;
- une URL d'outil n'est ni HTTPS ni loopback, ou contient un identifiant ;
- un champ secret d'`auth` (`api_key`, `password`, `token`…) ou un en-tête
  `Authorization` n'est pas une variable (`${VAR}` ou `Bearer ${VAR}`) ;
- un outil n'a ni tag `lecture` ni tag `ecriture`, ou a les deux ;
- en lecture seule, il ne reste aucun outil.

## Mode lecture seule

`OUTLOOK_LECTURE_SEULE=true` retire du manuel servi tous les outils tagués
`ecriture` (création, modification, suppression, envoi). Le démarrage le
signale en `warning` avec le nombre d'outils retirés. L'agent ne voit alors
que les outils de lecture.

C'est une restriction de **découverte**, pas une barrière de sécurité : un
client qui connaît l'API peut toujours écrire avec un jeton qui le permet.
Pour une vraie garantie, obtenir aussi un jeton limité aux scopes de lecture
(`outils/obtenir_jeton.py --lecture-seule`) : Graph refusera alors toute
écriture.

## Inscrire l'application dans Entra ID (une fois)

Le flux device code exige une application « client public » dans le
locataire (ou une application multi-locataire pour des comptes personnels).

1. Portail Entra (https://entra.microsoft.com) → *Identité* → *Applications* →
   *Inscriptions d'applications* → *Nouvelle inscription*.
2. Nom : `utcp-outlook` ; types de comptes : ce locataire seulement, ou
   « tout annuaire + comptes Microsoft personnels » pour outlook.com.
3. URI de redirection : plateforme *Applications mobiles et de bureau*,
   `https://login.microsoftonline.com/common/oauth2/nativeclient`.
4. *Authentification* → *Paramètres avancés* → **Autoriser les flux de clients
   publics** (« Allow public client flows ») : **Oui**. Aucun secret client
   n'est créé.
5. *Autorisations d'API* → *Microsoft Graph* → *Autorisations déléguées* :
   - `User.Read` ;
   - `Mail.ReadWrite`, `Mail.Send` ;
   - `MailboxSettings.Read` ;
   - `Calendars.ReadWrite` ;
   - `Contacts.ReadWrite` ;
   - `People.Read` ;
   - `offline_access` (jeton de rafraîchissement du cache local) ;
   - pour un usage lecture seule : `Mail.Read`, `Calendars.Read`,
     `Contacts.Read` à la place des `ReadWrite`, sans `Mail.Send`.
   Selon la politique du locataire, un administrateur doit accorder le
   consentement.
6. Noter l'**ID d'application (client)** et l'**ID de l'annuaire (locataire)**.

## Obtenir un jeton délégué

```bash
uv pip install --python .venv/bin/python -e ".[jeton]"     # msal
.venv/bin/python outils/obtenir_jeton.py --client-id <id-application> --tenant <id-locataire>
#  → message Microsoft sur stderr : ouvrir https://microsoft.com/devicelogin, saisir le code
#  → sur stdout : outlook_OUTLOOK_GRAPH_TOKEN=<jeton>
```

- `--client-id` / `OUTLOOK_CLIENT_ID` (obligatoire), `--tenant` /
  `OUTLOOK_TENANT` (défaut `common`).
- `--scopes "…"` remplace la liste par défaut ; `offline_access` est toujours
  ajouté ; `--lecture-seule` ramène chaque `X.ReadWrite` à `X.Read` et retire
  `Mail.Send`.
- `--cache` (défaut `~/.cache/utcp-outlook/jetons.json`, fichier 0600 dans un
  répertoire 0700) : le jeton de rafraîchissement y est gardé ; les appels
  suivants renvoient un jeton neuf **sans** reconnexion tant qu'il est valide.
  Ce fichier est un secret : ne jamais le copier ni le versionner.
- `--format env` (défaut), `json` ou `brut`. Le jeton n'apparaît que sur la
  sortie standard, jamais dans les messages.

Mettre la ligne obtenue dans `client/.env` du dépôt parent (voir
`client.env.example`), en remplaçant l'ancienne.

## Lancer

```bash
just installer                                     # uv venv + pip install -e .[dev,jeton]
.venv/bin/python -m serveur_utcp_outlook           # Graph v1.0 par défaut, port 8082
OUTLOOK_LECTURE_SEULE=true .venv/bin/python -m serveur_utcp_outlook
curl -s http://127.0.0.1:8082/utcp | jq '.tools | length'
```

## Consommer le manuel avec python-utcp

Template de manuel côté client (le nom `outlook` fixe le préfixe des variables) :

```json
{
  "manual_call_templates": [
    {"name": "outlook", "call_template_type": "http",
     "url": "http://127.0.0.1:8082/utcp", "http_method": "GET"}
  ],
  "load_variables_from": [{"variable_loader_type": "dotenv", "env_file_path": ".env"}]
}
```

Variable à fournir au client : `outlook_OUTLOOK_GRAPH_TOKEN`. La découverte
(`search_tools`) ne demande aucun jeton ; l'appel d'un outil en demande un.
Depuis le dépôt parent : passer l'entrée `outlook` du catalogue à
`actif: true`, lancer `tools/generer_config_client.py`, puis
`client/verifier_tous.py`.

## Limites

- **Jeton d'environ 1 h, sans rafraîchissement automatique** côté client
  UTCP : le client substitue une valeur fixe ; à expiration, Graph répond 401
  et il faut relancer `outils/obtenir_jeton.py` (silencieux tant que le cache
  est valide). Le type d'auth `oauth2` d'UTCP ne fait que le flux client
  credentials, inadapté à un jeton délégué.
- **Boîte de l'utilisateur connecté seulement** (`/me`) : boîtes partagées,
  délégations et `/users/{id}` ne sont pas couverts.
- **Exchange sur site** sans Graph : hors périmètre.
- Un outil = un appel HTTP JSON : pas d'agrégation ni de pagination
  automatique (détail des écarts dans `utcp/COUVERTURE.md`).

## Vérifier

```bash
just verifier        # ruff, pytest, validation du manuel (validateur du parent)
```

`OUTLOOK_MANUEL_TEST=/chemin/autre.json .venv/bin/pytest -q` teste un autre
manuel que `utcp/manual.json`.

À tout changement de contrat, incrémenter ensemble `version` de
`pyproject.toml` et `manual_version` du manuel (un test le vérifie).
