"""Le contrat du sous-module, vérifié sur l'application assemblée.

Le manuel testé est ``utcp/manual.json`` ; la variable d'environnement
``OUTLOOK_MANUEL_TEST`` permet d'en injecter un autre (manuel en cours de
rédaction, substitut).
"""

from __future__ import annotations

import json
import logging
import os
import re
import tomllib
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from pydantic import ValidationError

from serveur_utcp_outlook import utcp
from serveur_utcp_outlook.config import GRAPH_BASE_DEFAUT, Reglages
from serveur_utcp_outlook.main import creer_application

GRAPH = GRAPH_BASE_DEFAUT
JETON = "eyJ0eXAiOiJKV1QiLCJqZXRvbi1kZS10ZXN0LW5lLWRvaXQtamFtYWlzLXNvcnRpciJ9"
AUTH = {
    "auth_type": "api_key",
    "api_key": "Bearer ${OUTLOOK_GRAPH_TOKEN}",
    "var_name": "Authorization",
    "location": "header",
}
PARAM_URL = re.compile(r"(?<!\$)\{([A-Za-z0-9_]+)\}")


@pytest.fixture
def source() -> Path:
    chemin = Path(os.environ.get("OUTLOOK_MANUEL_TEST", utcp.CHEMIN_SOURCE))
    if not chemin.is_file():
        pytest.skip(f"manuel absent : {chemin}")
    return chemin


@pytest.fixture(autouse=True)
def environnement_propre(monkeypatch: pytest.MonkeyPatch) -> None:
    # Le poste de développement ne doit pas changer le résultat des tests.
    for cle in list(os.environ):
        if cle.startswith("OUTLOOK_") and cle != "OUTLOOK_MANUEL_TEST":
            monkeypatch.delenv(cle)


@pytest.fixture
def client(source: Path, monkeypatch: pytest.MonkeyPatch) -> TestClient:
    # Un jeton présent par erreur dans l'environnement du service ne doit rien changer.
    monkeypatch.setenv("OUTLOOK_GRAPH_TOKEN", JETON)
    monkeypatch.setenv("OUTLOOK_GRAPH_AUTHORIZATION", "Bearer " + JETON)
    return TestClient(creer_application(Reglages(), source))


def ecrire(tmp_path: Path, manuel: dict) -> Path:
    chemin = tmp_path / "manual.json"
    chemin.write_text(json.dumps(manuel), encoding="utf-8")
    return chemin


def outil(nom: str, url: str, tags: list[str], auth: dict | None = AUTH) -> dict:
    modele = {"call_template_type": "http", "url": url, "http_method": "GET"}
    if auth is not None:
        modele["auth"] = auth
    return {
        "name": nom,
        "description": "d",
        "tags": tags,
        "inputs": {},
        "tool_call_template": modele,
    }


def manuel_minimal(*outils: dict) -> dict:
    return {
        "manual_version": "0.1.0",
        "utcp_version": "1.0.2",
        "tools": list(outils) or [outil("t", "${GRAPH_BASE}/me", ["lecture"])],
    }


# --- Contrat de /utcp sur le manuel testé ---------------------------------------------


def test_utcp_repond_200_sans_authentification(client: TestClient) -> None:
    assert client.get("/utcp").status_code == 200


def test_utcp_est_servi_en_json_avec_cache_court(client: TestClient) -> None:
    r = client.get("/utcp")
    assert r.headers["content-type"].startswith("application/json")
    assert r.headers["cache-control"] == "public, max-age=300"


def test_manuel_rendu_porte_utcp_version_et_des_outils(client: TestClient) -> None:
    m = client.get("/utcp").json()
    assert isinstance(m["utcp_version"], str)
    assert m["tools"]


def test_manuel_rendu_ne_contient_plus_les_marqueurs_d_url(client: TestClient) -> None:
    texte = client.get("/utcp").text
    assert "${GRAPH_BASE}" not in texte
    assert "${BASE_URL}" not in texte


def test_seules_les_variables_client_restent(client: TestClient) -> None:
    restants = set(re.findall(r"\$\{([^}]*)\}", client.get("/utcp").text))
    assert restants
    assert restants <= {"OUTLOOK_GRAPH_TOKEN", "OUTLOOK_GRAPH_AUTHORIZATION"}


def test_chaque_url_d_outil_vise_graph(client: TestClient) -> None:
    for o in client.get("/utcp").json()["tools"]:
        assert o["tool_call_template"]["url"].startswith(GRAPH + "/"), o["name"]


def test_chaque_parametre_d_url_est_declare_dans_inputs(client: TestClient) -> None:
    for o in client.get("/utcp").json()["tools"]:
        props = o["inputs"].get("properties", {})
        for p in PARAM_URL.findall(o["tool_call_template"]["url"]):
            assert p in props, f"{o['name']} : {{{p}}} absent de inputs"


def test_chaque_outil_porte_exactement_un_tag_lecture_ou_ecriture(client: TestClient) -> None:
    for o in client.get("/utcp").json()["tools"]:
        assert len({"lecture", "ecriture"} & set(o["tags"])) == 1, o["name"]


def test_aucun_jeton_en_clair_dans_le_rendu(client: TestClient) -> None:
    r = client.get("/utcp")
    assert JETON not in r.text
    assert "OUTLOOK_GRAPH_TOKEN=" not in r.text
    for o in r.json()["tools"]:
        auth = o["tool_call_template"].get("auth")
        assert auth, f"{o['name']} : auth absente"
        assert auth["api_key"] in {
            "Bearer ${OUTLOOK_GRAPH_TOKEN}",
            "${OUTLOOK_GRAPH_AUTHORIZATION}",
        }


def test_health_est_public(client: TestClient) -> None:
    assert client.get("/health").json() == {"statut": "ok"}


def test_la_reponse_porte_un_x_request_id(client: TestClient) -> None:
    assert "x-request-id" in client.get("/utcp").headers


def test_un_x_request_id_entrant_est_renvoye_tel_quel(client: TestClient) -> None:
    r = client.get("/utcp", headers={"X-Request-Id": "corr-123"})
    assert r.headers["x-request-id"] == "corr-123"


def test_le_journal_d_acces_cite_la_route_generique(
    client: TestClient, caplog: pytest.LogCaptureFixture
) -> None:
    with caplog.at_level(logging.INFO, logger="serveur_utcp_outlook"):
        client.get("/utcp", headers={"X-Request-Id": "corr-9"})
        client.get("/inconnu/donnee-privee")
    lignes = [r for r in caplog.records if r.getMessage() == "requête servie"]
    assert [(r.route, r.statut) for r in lignes] == [("/utcp", 200), ("-", 404)]
    assert lignes[0].id_requete == "corr-9"
    assert "donnee-privee" not in caplog.text


# --- Mode lecture seule ---------------------------------------------------------------


def test_lecture_seule_retire_exactement_les_outils_ecriture(
    source: Path, monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    complet = json.loads(utcp.rendre(GRAPH, source=source).contenu)["tools"]
    monkeypatch.setenv("OUTLOOK_LECTURE_SEULE", "true")
    with caplog.at_level(logging.WARNING, logger="serveur_utcp_outlook"):
        app = creer_application(Reglages(), source)
    servis = TestClient(app).get("/utcp").json()["tools"]
    attendus = [o["name"] for o in complet if "ecriture" not in o["tags"]]
    assert [o["name"] for o in servis] == attendus
    assert all("lecture" in o["tags"] for o in servis)
    retires = len(complet) - len(servis)
    assert retires > 0, "le manuel testé devrait contenir des outils `ecriture`"
    avertissements = [r for r in caplog.records if r.levelno == logging.WARNING]
    assert len(avertissements) == 1
    assert f"{retires} outil(s)" in avertissements[0].getMessage()


def test_sans_lecture_seule_tous_les_outils_sont_servis(tmp_path: Path) -> None:
    chemin = ecrire(
        tmp_path,
        manuel_minimal(
            outil("lire", "${GRAPH_BASE}/me/messages", ["courrier", "lecture"]),
            outil("ecrire", "${GRAPH_BASE}/me/sendMail", ["courrier", "ecriture"]),
        ),
    )
    rendu = utcp.rendre(GRAPH, source=chemin)
    assert (rendu.nb_outils, rendu.nb_retires) == (2, 0)
    rendu = utcp.rendre(GRAPH, lecture_seule=True, source=chemin)
    assert (rendu.nb_outils, rendu.nb_retires) == (1, 1)
    assert [o["name"] for o in json.loads(rendu.contenu)["tools"]] == ["lire"]


def test_lecture_seule_refuse_un_manuel_sans_outil_lecture(tmp_path: Path) -> None:
    chemin = ecrire(tmp_path, manuel_minimal(outil("e", "${GRAPH_BASE}/me", ["ecriture"])))
    with pytest.raises(utcp.ErreurManuel, match="lecture seule"):
        utcp.rendre(GRAPH, lecture_seule=True, source=chemin)


def test_rendre_refuse_un_outil_sans_tag_lecture_ni_ecriture(tmp_path: Path) -> None:
    chemin = ecrire(tmp_path, manuel_minimal(outil("muet", "${GRAPH_BASE}/me", ["courrier"])))
    with pytest.raises(utcp.ErreurManuel, match="muet : ni tag"):
        utcp.rendre(GRAPH, source=chemin)


def test_rendre_refuse_un_outil_avec_les_deux_tags(tmp_path: Path) -> None:
    chemin = ecrire(
        tmp_path, manuel_minimal(outil("double", "${GRAPH_BASE}/me", ["lecture", "ecriture"]))
    )
    with pytest.raises(utcp.ErreurManuel, match="double : tags"):
        utcp.rendre(GRAPH, source=chemin)


def test_reglages_lisent_lecture_seule(monkeypatch: pytest.MonkeyPatch) -> None:
    assert Reglages().lecture_seule is False
    monkeypatch.setenv("OUTLOOK_LECTURE_SEULE", "true")
    assert Reglages().lecture_seule is True


# --- Réglages et refus au démarrage ---------------------------------------------------


def test_reglages_par_defaut() -> None:
    r = Reglages()
    assert r.graph_base == "https://graph.microsoft.com/v1.0"
    assert r.ecoute_port == 8082


def test_reglages_refusent_une_graph_base_http_distante() -> None:
    with pytest.raises(ValidationError):
        Reglages(graph_base="http://graph.microsoft.com/v1.0")


def test_reglages_refusent_des_identifiants_dans_graph_base() -> None:
    with pytest.raises(ValidationError):
        Reglages(graph_base="https://moi:secret@graph.microsoft.com/v1.0")


def test_reglages_acceptent_loopback_en_http_et_retirent_le_slash_final() -> None:
    assert Reglages(graph_base="http://127.0.0.1:9000/v1.0/").graph_base == (
        "http://127.0.0.1:9000/v1.0"
    )
    assert Reglages(graph_base=GRAPH + "/").graph_base == GRAPH


def test_rendre_refuse_un_marqueur_non_rendu(tmp_path: Path) -> None:
    chemin = ecrire(
        tmp_path, manuel_minimal(outil("t", "${GRAPH_BASE}/x?cle=${AUTRE_SECRET}", ["lecture"]))
    )
    with pytest.raises(utcp.ErreurManuel, match="AUTRE_SECRET"):
        utcp.rendre(GRAPH, source=chemin)


def test_rendre_exige_l_url_publique_si_le_manuel_cite_base_url(tmp_path: Path) -> None:
    chemin = ecrire(tmp_path, manuel_minimal(outil("t", "${BASE_URL}/x", ["lecture"])))
    with pytest.raises(utcp.ErreurManuel, match="OUTLOOK_URL_PUBLIQUE"):
        utcp.rendre(GRAPH, source=chemin)
    rendu = json.loads(utcp.rendre(GRAPH, "https://utcp.exemple.invalid", source=chemin).contenu)
    assert rendu["tools"][0]["tool_call_template"]["url"] == "https://utcp.exemple.invalid/x"


def test_rendre_refuse_une_url_d_outil_http_distante(tmp_path: Path) -> None:
    chemin = ecrire(
        tmp_path, manuel_minimal(outil("t", "http://ailleurs.exemple.invalid/x", ["lecture"]))
    )
    with pytest.raises(utcp.ErreurManuel, match="URL refusée"):
        utcp.rendre(GRAPH, source=chemin)


def test_rendre_refuse_un_jeton_en_clair_dans_auth(tmp_path: Path) -> None:
    auth = dict(AUTH, api_key="Bearer " + JETON)
    chemin = ecrire(tmp_path, manuel_minimal(outil("t", "${GRAPH_BASE}/me", ["lecture"], auth)))
    with pytest.raises(utcp.ErreurManuel, match="auth.api_key"):
        utcp.rendre(GRAPH, source=chemin)


def test_rendre_refuse_un_jeton_en_clair_dans_l_en_tete_authorization(tmp_path: Path) -> None:
    o = outil("t", "${GRAPH_BASE}/me", ["lecture"], auth=None)
    o["tool_call_template"]["headers"] = {"Authorization": "Bearer " + JETON}
    chemin = ecrire(tmp_path, manuel_minimal(o))
    with pytest.raises(utcp.ErreurManuel, match="Authorization"):
        utcp.rendre(GRAPH, source=chemin)


def test_rendre_accepte_les_deux_formes_de_variable_de_jeton(tmp_path: Path) -> None:
    auth = dict(AUTH, api_key="${OUTLOOK_GRAPH_AUTHORIZATION}")
    chemin = ecrire(
        tmp_path,
        manuel_minimal(
            outil("a", "${GRAPH_BASE}/me", ["lecture"]),
            outil("b", "${GRAPH_BASE}/me", ["lecture"], auth),
        ),
    )
    assert utcp.rendre(GRAPH, source=chemin).nb_outils == 2


def test_manual_version_suit_la_version_du_paquet() -> None:
    if not utcp.CHEMIN_SOURCE.is_file():
        pytest.skip("utcp/manual.json absent")
    pyproject = tomllib.loads((utcp.CHEMIN_SOURCE.parents[1] / "pyproject.toml").read_text())
    manuel = json.loads(utcp.CHEMIN_SOURCE.read_text(encoding="utf-8"))
    assert manuel["manual_version"] == pyproject["project"]["version"]
