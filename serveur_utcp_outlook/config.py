"""Réglages lus dans l'environnement, préfixe ``OUTLOOK_``."""

from __future__ import annotations

from urllib.parse import urlsplit

from pydantic import Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

GRAPH_BASE_DEFAUT = "https://graph.microsoft.com/v1.0"
_HOTES_LOOPBACK = {"localhost", "127.0.0.1", "::1"}


def valider_url_https(url: str, quoi: str) -> str:
    """URL HTTPS (ou http vers loopback), sans identifiants, requête ni fragment ; slash final retiré.

    Les deux URL finissent dans un manuel public : un identifiant glissé dans
    l'URL (``https://moi:jeton@…``) y serait publié, d'où le refus.
    """
    url = url.strip().rstrip("/")
    morceaux = urlsplit(url)
    hote = (morceaux.hostname or "").lower()
    loopback = hote in _HOTES_LOOPBACK or hote.startswith("127.")
    if morceaux.scheme == "https" or (morceaux.scheme == "http" and loopback):
        pass
    else:
        raise ValueError(f"{quoi} refusée : https obligatoire, http seulement vers loopback")
    if not hote:
        raise ValueError(f"{quoi} refusée : hôte absent")
    if morceaux.username or morceaux.password:
        raise ValueError(f"{quoi} refusée : pas d'identifiant dans l'URL")
    if morceaux.query or morceaux.fragment:
        raise ValueError(f"{quoi} refusée : ni requête ni fragment")
    return url


class Reglages(BaseSettings):
    """Réglages du service.

    Aucun secret ici : le service n'a pas de route métier à protéger, et le
    jeton Microsoft Graph (``OUTLOOK_GRAPH_TOKEN``) est fourni par le client
    UTCP, jamais par ce service. Une variable ``OUTLOOK_GRAPH_TOKEN`` présente
    par erreur dans l'environnement du service est ignorée (aucun champ ne la lit).
    """

    model_config = SettingsConfigDict(env_prefix="OUTLOOK_", extra="ignore")

    graph_base: str = Field(
        default=GRAPH_BASE_DEFAUT,
        description="Racine de l'API Microsoft Graph, rendue à la place de ${GRAPH_BASE}",
    )
    url_publique: str | None = Field(
        default=None,
        description="URL publique de ce service, rendue à la place de ${BASE_URL} s'il apparaît",
    )
    lecture_seule: bool = Field(
        default=False,
        description="true : les outils marqués `ecriture` sont retirés du manuel servi",
    )
    ecoute_hote: str = "127.0.0.1"
    ecoute_port: int = Field(default=8082, ge=1, le=65535)
    journal: str = Field(default="INFO", description="Niveau du journal racine")
    journal_format: str = Field(default="texte", description="texte ou json")

    @field_validator("graph_base")
    @classmethod
    def _graph_base(cls, v: str) -> str:
        return valider_url_https(v, "OUTLOOK_GRAPH_BASE")

    @field_validator("url_publique")
    @classmethod
    def _url_publique(cls, v: str | None) -> str | None:
        return None if v is None else valider_url_https(v, "OUTLOOK_URL_PUBLIQUE")
