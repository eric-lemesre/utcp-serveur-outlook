"""Fournisseur UTCP ``outlook`` : sert le manuel de Microsoft Graph (Outlook).

Forme « manuel seul » : le service publie ``GET /utcp`` et ``GET /health``,
sans route métier ; les outils du manuel appellent Microsoft Graph directement
avec le jeton délégué que fournit le client UTCP.
Le paquet émet des journaux mais ne configure jamais ``logging`` : c'est le
rôle du seul point d'entrée, ``serveur_utcp_outlook/__main__.py``.
"""
