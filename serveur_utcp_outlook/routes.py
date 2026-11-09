"""Routes du service : ``/health`` seulement (``/utcp`` est dans ``utcp.py``).

Pas de route métier : les outils du manuel visent directement Microsoft Graph.
"""

from __future__ import annotations

from fastapi import APIRouter
from pydantic import BaseModel

public = APIRouter()


class Sante(BaseModel):
    statut: str = "ok"


@public.get("/health")
def sante() -> Sante:
    return Sante()
