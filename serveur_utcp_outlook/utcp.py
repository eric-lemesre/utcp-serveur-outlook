"""Manuel UTCP : rendu au démarrage depuis ``utcp/manual.json``, servi sur ``GET /utcp``.

Le service rend ``${GRAPH_BASE}`` (racine de Microsoft Graph) et ``${BASE_URL}``
(sa propre URL publique, si le manuel la cite). Le jeton Graph reste une
variable, ``${OUTLOOK_GRAPH_TOKEN}`` ou ``${OUTLOOK_GRAPH_AUTHORIZATION}``, que
le client UTCP fournit de son côté : il ne transite jamais par ici.

Chaque outil porte exactement un des tags ``lecture`` ou ``ecriture`` ; en mode
lecture seule, les outils ``ecriture`` sont retirés du manuel servi.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit

from fastapi import APIRouter, Response

CHEMIN_SOURCE = Path(__file__).resolve().parents[1] / "utcp" / "manual.json"
MARQUEUR_GRAPH_BASE = "${GRAPH_BASE}"
MARQUEUR_BASE_URL = "${BASE_URL}"
# Seules variables laissées au client après rendu (le manuel en utilise une).
VARIABLES_CLIENT = frozenset({"OUTLOOK_GRAPH_TOKEN", "OUTLOOK_GRAPH_AUTHORIZATION"})
TAG_LECTURE = "lecture"
TAG_ECRITURE = "ecriture"
_MARQUEUR = re.compile(r"\$\{([^}]*)\}")
# Valeur d'un champ secret : une variable seule, éventuellement précédée du schéma « Bearer ».
_VALEUR_VARIABLE = re.compile(r"^(?:Bearer )?\$\{[A-Za-z0-9_]+\}$")
# Champs d'``auth`` qui portent un secret ou un identifiant : jamais en clair dans le manuel.
_CHAMPS_SECRETS = ("username", "password", "api_key", "client_id", "client_secret", "token")
_HOTES_LOOPBACK = {"localhost", "127.0.0.1", "::1"}


class ErreurManuel(ValueError):
    """Le manuel source ou les URL configurées sont inexploitables ; arrête le service au démarrage."""


@dataclass(frozen=True)
class ManuelRendu:
    """Manuel sérialisé prêt à servir, et nombre d'outils retirés par le mode lecture seule."""

    contenu: bytes
    nb_outils: int
    nb_retires: int


def url_outil_admise(url: str) -> bool:
    """HTTPS, ou http vers loopback (tests) ; jamais d'identifiant dans l'URL."""
    m = urlsplit(url)
    hote = (m.hostname or "").lower()
    if m.username or m.password or not hote:
        return False
    return m.scheme == "https" or (
        m.scheme == "http" and (hote in _HOTES_LOOPBACK or hote.startswith("127."))
    )


def _est_variable(valeur: Any) -> bool:
    return isinstance(valeur, str) and bool(_VALEUR_VARIABLE.match(valeur))


def _verifier_outil(outil: Any) -> None:
    nom = outil.get("name", "?") if isinstance(outil, dict) else "?"
    modele = outil.get("tool_call_template") if isinstance(outil, dict) else None
    if not isinstance(modele, dict):
        raise ErreurManuel(f"outil {nom} : `tool_call_template` absent")
    url = modele.get("url")
    if not isinstance(url, str) or not url_outil_admise(url):
        raise ErreurManuel(f"outil {nom} : URL refusée (https ou loopback obligatoire) : {url}")
    auth = modele.get("auth") or {}
    if not isinstance(auth, dict):
        raise ErreurManuel(f"outil {nom} : `auth` doit être un objet")
    for champ in _CHAMPS_SECRETS:
        valeur = auth.get(champ)
        if valeur is not None and not _est_variable(valeur):
            # La valeur n'est pas une variable : c'est peut-être un jeton en clair.
            raise ErreurManuel(f"outil {nom} : `auth.{champ}` doit être une variable ${{…}}")
    entetes = modele.get("headers") or {}
    if isinstance(entetes, dict):
        for cle, valeur in entetes.items():
            if cle.lower() == "authorization" and not _est_variable(valeur):
                raise ErreurManuel(f"outil {nom} : en-tête `{cle}` doit être une variable ${{…}}")
    tags = outil.get("tags") or []
    if not isinstance(tags, list):
        raise ErreurManuel(f"outil {nom} : `tags` doit être une liste")
    lecture, ecriture = TAG_LECTURE in tags, TAG_ECRITURE in tags
    if lecture and ecriture:
        raise ErreurManuel(f"outil {nom} : tags `lecture` et `ecriture` à la fois")
    if not (lecture or ecriture):
        raise ErreurManuel(f"outil {nom} : ni tag `lecture` ni tag `ecriture`")


def _verifier(manuel: Any) -> list[dict]:
    if not isinstance(manuel, dict):
        raise ErreurManuel("manuel source : un objet JSON est attendu")
    if not isinstance(manuel.get("utcp_version"), str):
        raise ErreurManuel("manuel source : champ `utcp_version` absent")
    outils = manuel.get("tools")
    if not isinstance(outils, list) or not outils:
        raise ErreurManuel("manuel source : champ `tools` absent ou vide")
    for outil in outils:
        _verifier_outil(outil)
    return outils


def rendre(
    graph_base: str,
    url_publique: str | None = None,
    lecture_seule: bool = False,
    source: Path = CHEMIN_SOURCE,
) -> ManuelRendu:
    """Rend ``${GRAPH_BASE}`` et ``${BASE_URL}``, vérifie, filtre, sérialise une fois.

    ``graph_base`` et ``url_publique`` arrivent validés par ``Reglages`` (HTTPS
    ou loopback, sans slash final). Toute anomalie lève ``ErreurManuel``.
    """
    try:
        texte = source.read_text(encoding="utf-8")
    except OSError as e:
        raise ErreurManuel(f"manuel source illisible : {source} ({e.strerror})") from e
    texte = texte.replace(MARQUEUR_GRAPH_BASE, graph_base.rstrip("/"))
    if MARQUEUR_BASE_URL in texte:
        if not url_publique:
            raise ErreurManuel("le manuel cite ${BASE_URL} : définir OUTLOOK_URL_PUBLIQUE")
        texte = texte.replace(MARQUEUR_BASE_URL, url_publique.rstrip("/"))
    restants = sorted(set(_MARQUEUR.findall(texte)) - VARIABLES_CLIENT)
    if restants:
        raise ErreurManuel(
            "marqueur(s) non rendu(s) dans le manuel : " + ", ".join(f"${{{v}}}" for v in restants)
        )
    try:
        manuel = json.loads(texte)
    except json.JSONDecodeError as e:
        raise ErreurManuel(f"manuel source invalide : {e}") from e
    outils = _verifier(manuel)
    retenus = [o for o in outils if not (lecture_seule and TAG_ECRITURE in o["tags"])]
    if not retenus:
        raise ErreurManuel("mode lecture seule : aucun outil `lecture` dans le manuel")
    manuel["tools"] = retenus
    contenu = json.dumps(manuel, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return ManuelRendu(contenu, len(retenus), len(outils) - len(retenus))


def routeur(manuel: bytes) -> APIRouter:
    """Routeur autonome pour ``GET /utcp`` (public : le manuel ne contient aucun secret)."""
    r = APIRouter()

    @r.get("/utcp", include_in_schema=False)
    def servir() -> Response:
        return Response(
            content=manuel,
            media_type="application/json",
            # Le manuel change avec le déploiement, pas avec les données : cache court.
            headers={"Cache-Control": "public, max-age=300"},
        )

    return r
