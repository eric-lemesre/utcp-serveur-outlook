"""Point d'entrée : seul endroit qui configure ``logging``, lit les réglages et sert.

python -m serveur_utcp_outlook
"""

from __future__ import annotations

import json
import logging
import sys

import uvicorn
from pydantic import ValidationError

from serveur_utcp_outlook.config import Reglages
from serveur_utcp_outlook.main import creer_application
from serveur_utcp_outlook.utcp import ErreurManuel

journal = logging.getLogger("serveur_utcp_outlook")


class FormateurJson(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        champs = {
            "niveau": record.levelname,
            "cible": record.name,
            "message": record.getMessage(),
        }
        for cle in ("id_requete", "methode", "route", "statut"):
            if hasattr(record, cle):
                champs[cle] = getattr(record, cle)
        return json.dumps(champs, ensure_ascii=False)


def init_journal(niveau: str, format_: str) -> None:
    gestionnaire = logging.StreamHandler(sys.stderr)
    if format_.lower() == "json":
        gestionnaire.setFormatter(FormateurJson())
    else:
        gestionnaire.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s")
        )
    logging.basicConfig(level=niveau.upper(), handlers=[gestionnaire], force=True)
    # Les lignes d'accès d'uvicorn citent le chemin brut (données dedans) : on garde la nôtre.
    logging.getLogger("uvicorn.access").disabled = True


def main() -> int:
    try:
        # OUTLOOK_GRAPH_BASE non HTTPS, port invalide… : erreur claire, pas de démarrage.
        reglages = Reglages()
    except ValidationError as e:
        print(f"réglages invalides, service non démarré :\n{e}", file=sys.stderr)
        return 2
    init_journal(reglages.journal, reglages.journal_format)
    try:
        app = creer_application(reglages)
    except ErreurManuel as e:
        journal.error("manuel non rendu, service non démarré : %s", e)
        return 2
    journal.info(
        "service démarré, manuel servi sur /utcp",
        extra={"route": "/utcp", "methode": "GET", "statut": 0, "id_requete": "-"},
    )
    uvicorn.run(app, host=reglages.ecoute_hote, port=reglages.ecoute_port, log_config=None)
    return 0


if __name__ == "__main__":
    sys.exit(main())
