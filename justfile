# Commandes du fournisseur outlook. `just` : https://just.systems

validateur := env_var_or_default("UTCP_VALIDATEUR", "../../tools/valider_manuel.py")

# Tout ce que la CI vérifie
verifier: lint test valider

installer:
    uv venv .venv && uv pip install --python .venv/bin/python -e ".[dev,jeton]"

lint:
    .venv/bin/ruff check . && .venv/bin/ruff format --check .

test:
    .venv/bin/pytest -q

# Valide le manuel source (${GRAPH_BASE} et ${BASE_URL} rendus)
valider:
    python3 -I {{validateur}} utcp/manual.json --var GRAPH_BASE=https://graph.microsoft.com/v1.0 --var BASE_URL=https://validation.invalid --strict

# Lance en local ; le manuel est alors servi sur http://127.0.0.1:8082/utcp
run lecture_seule="false":
    OUTLOOK_LECTURE_SEULE={{lecture_seule}} OUTLOOK_ECOUTE_PORT=8082 .venv/bin/python -m serveur_utcp_outlook

# Jeton délégué Graph pour client/.env (device code ; voir README)
jeton client_id *args:
    .venv/bin/python outils/obtenir_jeton.py --client-id {{client_id}} {{args}}
