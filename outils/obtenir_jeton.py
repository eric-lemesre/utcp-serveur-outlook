#!/usr/bin/env python3
"""Obtient un jeton DÉLÉGUÉ Microsoft Graph pour le client UTCP (flux device code, MSAL).

Usage :
    python outils/obtenir_jeton.py --client-id <id-application> [--tenant common]
        [--scopes "User.Read Mail.ReadWrite …"] [--lecture-seule]
        [--cache ~/.cache/utcp-outlook/jetons.json] [--format env|json|brut]

Le jeton est imprimé sur la SORTIE STANDARD seulement, au format demandé (par
défaut ``outlook_OUTLOOK_GRAPH_TOKEN=<jeton>``, prêt pour ``client/.env``) ;
les messages (code à saisir, erreurs) vont sur la sortie d'erreur et ne citent
jamais le jeton. Le cache MSAL (jeton de rafraîchissement compris) est un
fichier 0600 dans un répertoire 0700 : tant qu'il est valide, le script renvoie
un jeton neuf sans redemander de connexion.

Dépendance optionnelle : ``pip install -e '.[jeton]'`` (msal).
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

journal = logging.getLogger("obtenir_jeton")

SCOPES_COMPLETS = (
    "User.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "MailboxSettings.Read",
    "Calendars.ReadWrite",
    "Contacts.ReadWrite",
    "People.Read",
)
# Demandé pour obtenir un jeton de rafraîchissement (cache réutilisable).
SCOPE_HORS_LIGNE = "offline_access"
# Scopes OpenID que MSAL ajoute lui-même et refuse qu'on lui passe.
SCOPES_RESERVES = frozenset({"offline_access", "openid", "profile"})
CACHE_DEFAUT = Path("~/.cache/utcp-outlook/jetons.json")
VARIABLE_CLIENT = "outlook_OUTLOOK_GRAPH_TOKEN"
_SCOPE = re.compile(r"^[A-Za-z][A-Za-z0-9.\-_]*$")
_GUID = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
_TENANT = re.compile(r"^[A-Za-z0-9][A-Za-z0-9.\-]{0,252}$")


class ErreurJeton(Exception):
    """Erreur présentée telle quelle à l'utilisateur (jamais de jeton dedans)."""


# --- Fonctions pures (testées sans réseau) -------------------------------------------


def restreindre_lecture(scopes: Sequence[str]) -> list[str]:
    """Version lecture seule d'une liste de scopes.

    ``X.ReadWrite[.Y]`` devient ``X.Read[.Y]`` ; les scopes d'envoi (``*.Send``)
    et toute autre écriture restante sont retirés. Ordre conservé, sans doublon.
    """
    resultat: list[str] = []
    for scope in scopes:
        lu = scope.replace(".ReadWrite", ".Read")
        if lu.endswith(".Send") or ".Write" in lu:
            continue
        if lu not in resultat:
            resultat.append(lu)
    return resultat


def calculer_scopes(scopes: str | None, lecture_seule: bool) -> list[str]:
    """Scopes demandés : défaut ou ``--scopes`` (espaces ou virgules), restreints si lecture
    seule, ``offline_access`` toujours présent (une seule fois, en dernier)."""
    if scopes is None:
        liste = list(SCOPES_COMPLETS)
    else:
        liste = [s for s in re.split(r"[\s,]+", scopes.strip()) if s]
    for s in liste:
        if not _SCOPE.match(s):
            raise ErreurJeton(f"scope invalide : {s!r}")
    liste = [s for s in liste if s != SCOPE_HORS_LIGNE]
    if lecture_seule:
        liste = restreindre_lecture(liste)
    else:
        liste = list(dict.fromkeys(liste))
    if not liste:
        raise ErreurJeton("aucun scope à demander")
    return [*liste, SCOPE_HORS_LIGNE]


def scopes_pour_msal(scopes: Sequence[str]) -> list[str]:
    """MSAL refuse les scopes réservés (il ajoute ``offline_access`` lui-même)."""
    return [s for s in scopes if s not in SCOPES_RESERVES]


def analyser_arguments(
    argv: Sequence[str], environ: Mapping[str, str] | None = None
) -> argparse.Namespace:
    """Analyse la ligne de commande ; ``OUTLOOK_CLIENT_ID`` et ``OUTLOOK_TENANT`` en défauts."""
    env = os.environ if environ is None else environ
    p = argparse.ArgumentParser(
        prog="obtenir_jeton.py",
        description="Jeton délégué Microsoft Graph par device code, pour le client UTCP outlook.",
    )
    p.add_argument(
        "--client-id",
        default=env.get("OUTLOOK_CLIENT_ID"),
        help="ID d'application (client) Entra ID ; défaut : $OUTLOOK_CLIENT_ID",
    )
    p.add_argument(
        "--tenant",
        default=env.get("OUTLOOK_TENANT") or "common",
        help="locataire : common, organizations, consumers, GUID ou domaine (défaut : common)",
    )
    p.add_argument(
        "--scopes",
        default=None,
        help="scopes délégués séparés par des espaces ; défaut : " + " ".join(SCOPES_COMPLETS),
    )
    p.add_argument(
        "--cache",
        type=Path,
        default=CACHE_DEFAUT,
        help=f"cache de jetons MSAL (défaut : {CACHE_DEFAUT})",
    )
    p.add_argument("--lecture-seule", action="store_true", help="scopes de lecture seulement")
    p.add_argument("--format", choices=("env", "json", "brut"), default="env")
    a = p.parse_args(list(argv))
    if not a.client_id:
        p.error("--client-id obligatoire (ou variable OUTLOOK_CLIENT_ID)")
    if not _GUID.match(a.client_id):
        p.error("--client-id : GUID attendu (ID d'application Entra ID)")
    if not _TENANT.match(a.tenant):
        p.error("--tenant invalide")
    try:
        a.liste_scopes = calculer_scopes(a.scopes, a.lecture_seule)
    except ErreurJeton as e:
        p.error(str(e))
    a.cache = a.cache.expanduser()
    return a


def formater(resultat: Mapping[str, Any], format_: str) -> str:
    """Texte à imprimer sur la sortie standard (seul endroit où le jeton apparaît)."""
    jeton = resultat["access_token"]
    if format_ == "brut":
        return jeton
    if format_ == "json":
        return json.dumps(
            {
                "access_token": jeton,
                "expires_in": resultat.get("expires_in"),
                "scope": resultat.get("scope"),
            },
            ensure_ascii=False,
        )
    return f"{VARIABLE_CLIENT}={jeton}"


# --- Flux réseau (MSAL) --------------------------------------------------------------


def _charger_cache(msal: Any, chemin: Path) -> Any:
    cache = msal.SerializableTokenCache()
    if chemin.is_file():
        cache.deserialize(chemin.read_text(encoding="utf-8"))
    return cache


def _enregistrer_cache(cache: Any, chemin: Path) -> None:
    if not cache.has_state_changed:
        return
    chemin.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
    os.chmod(chemin.parent, 0o700)
    fd = os.open(chemin, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(cache.serialize())
    os.chmod(chemin, 0o600)


def obtenir(a: argparse.Namespace) -> dict[str, Any]:
    """Cache d'abord (acquire_token_silent), device code sinon ; lève ``ErreurJeton``."""
    try:
        import msal
    except ImportError as e:
        raise ErreurJeton(
            "bibliothèque msal absente : pip install -e '.[jeton]' (ou pip install 'msal>=1.28')"
        ) from e
    scopes = scopes_pour_msal(a.liste_scopes)
    cache = _charger_cache(msal, a.cache)
    app = msal.PublicClientApplication(
        a.client_id,
        authority=f"https://login.microsoftonline.com/{a.tenant}",
        token_cache=cache,
    )
    resultat = None
    comptes = app.get_accounts()
    if comptes:
        resultat = app.acquire_token_silent(scopes, account=comptes[0])
        if resultat and "access_token" in resultat:
            journal.info("jeton renouvelé depuis le cache (%s)", comptes[0].get("username", "?"))
    if not resultat or "access_token" not in resultat:
        flux = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flux:
            raise ErreurJeton(
                "device code refusé : "
                f"{flux.get('error', '?')} — {flux.get('error_description', '')}".strip()
            )
        # Message Microsoft : URL de connexion et code à saisir.
        print(flux["message"], file=sys.stderr, flush=True)
        resultat = app.acquire_token_by_device_flow(flux)
    _enregistrer_cache(cache, a.cache)
    if "access_token" not in resultat:
        raise ErreurJeton(
            f"jeton refusé : {resultat.get('error', '?')} — {resultat.get('error_description', '')}"
        )
    return resultat


def main(argv: Sequence[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s", stream=sys.stderr)
    a = analyser_arguments(sys.argv[1:] if argv is None else argv)
    journal.info("scopes demandés : %s", " ".join(a.liste_scopes))
    try:
        resultat = obtenir(a)
    except ErreurJeton as e:
        journal.error("%s", e)
        return 2
    print(formater(resultat, a.format))
    journal.info("jeton délégué obtenu, valable %s s", resultat.get("expires_in", "?"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
