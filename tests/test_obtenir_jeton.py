"""``outils/obtenir_jeton.py`` : analyse des arguments et scopes, sans aucun appel réseau."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest

_CHEMIN = Path(__file__).resolve().parents[1] / "outils" / "obtenir_jeton.py"
_spec = importlib.util.spec_from_file_location("obtenir_jeton", _CHEMIN)
assert _spec and _spec.loader
oj = importlib.util.module_from_spec(_spec)
sys.modules["obtenir_jeton"] = oj
_spec.loader.exec_module(oj)

CLIENT_ID = "00000000-1111-2222-3333-444444444444"
ECRITURE = ("ReadWrite", ".Send", ".Write")


def test_scopes_par_defaut_complets_avec_offline_access() -> None:
    scopes = oj.calculer_scopes(None, lecture_seule=False)
    assert scopes == [*oj.SCOPES_COMPLETS, "offline_access"]


def test_scopes_lecture_seule_sans_ecriture_avec_offline_access() -> None:
    scopes = oj.calculer_scopes(None, lecture_seule=True)
    assert scopes[-1] == "offline_access"
    assert scopes.count("offline_access") == 1
    assert not [s for s in scopes if any(m in s for m in ECRITURE)]
    assert scopes == [
        "User.Read",
        "Mail.Read",
        "MailboxSettings.Read",
        "Calendars.Read",
        "Contacts.Read",
        "People.Read",
        "offline_access",
    ]


def test_scopes_lecture_seule_inclus_dans_le_complet() -> None:
    """Chaque droit de lecture est accordé par le jeu complet (X.Read ⊂ X.ReadWrite)."""
    complet = set(oj.calculer_scopes(None, lecture_seule=False))
    lecture = oj.calculer_scopes(None, lecture_seule=True)
    assert len(lecture) < len(complet)
    for s in lecture:
        assert s in complet or s.replace(".Read", ".ReadWrite", 1) in complet, s


def test_scopes_explicites_sans_doublon_d_offline_access() -> None:
    scopes = oj.calculer_scopes("User.Read, Mail.Read offline_access Mail.Read", False)
    assert scopes == ["User.Read", "Mail.Read", "offline_access"]


def test_scopes_invalides_refuses() -> None:
    with pytest.raises(oj.ErreurJeton):
        oj.calculer_scopes("User.Read ;rm", False)
    with pytest.raises(oj.ErreurJeton):
        oj.calculer_scopes("Mail.Send", True)


def test_msal_ne_recoit_pas_les_scopes_reserves() -> None:
    scopes = oj.scopes_pour_msal(oj.calculer_scopes(None, False))
    assert "offline_access" not in scopes
    assert scopes == list(oj.SCOPES_COMPLETS)


def test_arguments_par_defaut() -> None:
    a = oj.analyser_arguments(["--client-id", CLIENT_ID], environ={})
    assert a.tenant == "common"
    assert a.format == "env"
    assert a.lecture_seule is False
    assert a.cache == Path("~/.cache/utcp-outlook/jetons.json").expanduser()
    assert a.liste_scopes[-1] == "offline_access"


def test_arguments_lus_dans_l_environnement() -> None:
    env = {"OUTLOOK_CLIENT_ID": CLIENT_ID, "OUTLOOK_TENANT": "contoso.onmicrosoft.com"}
    a = oj.analyser_arguments(["--lecture-seule", "--format", "json"], environ=env)
    assert (a.client_id, a.tenant, a.format) == (CLIENT_ID, "contoso.onmicrosoft.com", "json")
    assert "Mail.Send" not in a.liste_scopes


def test_client_id_obligatoire() -> None:
    with pytest.raises(SystemExit):
        oj.analyser_arguments([], environ={})


def test_client_id_et_tenant_invalides_refuses() -> None:
    with pytest.raises(SystemExit):
        oj.analyser_arguments(["--client-id", "pas-un-guid"], environ={})
    with pytest.raises(SystemExit):
        oj.analyser_arguments(["--client-id", CLIENT_ID, "--tenant", "a/b"], environ={})


def test_format_inconnu_refuse() -> None:
    with pytest.raises(SystemExit):
        oj.analyser_arguments(["--client-id", CLIENT_ID, "--format", "yaml"], environ={})


def test_formats_de_sortie() -> None:
    r = {"access_token": "abc", "expires_in": 3599, "scope": "Mail.Read"}
    assert oj.formater(r, "env") == "outlook_OUTLOOK_GRAPH_TOKEN=abc"
    assert oj.formater(r, "brut") == "abc"
    assert '"access_token": "abc"' in oj.formater(r, "json")
