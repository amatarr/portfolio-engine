"""
Standalone script -- run this locally (NOT through Streamlit) while CQGXL
is live-updating CQG LINKS SHEET_.xlsx. It pushes the parsed quotes to a
GitHub Gist every PUSH_INTERVAL_SECONDS, so the deployed Streamlit Cloud
app (which can't see this machine's filesystem) can read them instead.

Setup:
  1. Generate a GitHub Personal Access Token scoped to just "gist"
     (github.com -> Settings -> Developer settings -> Personal access
     tokens -> Fine-grained or classic, scope: gist).
  2. Add it to .streamlit/secrets.toml:
         github_token = "..."
  3. Run: py push_quotes.py
     First run creates the gist and appends its id to secrets.toml as
     gist_id. Add that same gist_id to the Streamlit Cloud app's secrets
     too (Settings -> Secrets), then redeploy -- it's not sensitive, just
     an identifier.
  4. Leave this running. Ctrl+C to stop -- the deployed app just keeps
     showing the last pushed values until you restart it.
"""

import json
import time
import tomllib

import requests

from vol_calculator import import_cqg_quotes

SECRETS_PATH = ".streamlit/secrets.toml"
QUOTES_PATH = "CQG LINKS SHEET_.xlsx"
PUSH_INTERVAL_SECONDS = 30


def _load_secrets():
    with open(SECRETS_PATH, "rb") as f:
        return tomllib.load(f)


def _save_gist_id(gist_id):
    with open(SECRETS_PATH, "a", encoding="utf-8") as f:
        f.write(f'\ngist_id = "{gist_id}"\n')


def _create_gist(token, content):
    resp = requests.post(
        "https://api.github.com/gists",
        headers={"Authorization": f"token {token}"},
        json={
            "description": "Vol Calculator -- live CQGXL quotes (auto-updated)",
            "public": False,
            "files": {"quotes.json": {"content": content}},
        },
        timeout=10,
    )
    resp.raise_for_status()
    return resp.json()["id"]


def _update_gist(token, gist_id, content):
    resp = requests.patch(
        f"https://api.github.com/gists/{gist_id}",
        headers={"Authorization": f"token {token}"},
        json={"files": {"quotes.json": {"content": content}}},
        timeout=10,
    )
    resp.raise_for_status()


def main():
    secrets = _load_secrets()
    token = secrets.get("github_token")

    if not token:
        raise SystemExit(
            "Missing github_token in .streamlit/secrets.toml -- generate a "
            "Personal Access Token (scope: gist) and add it there first."
        )

    gist_id = secrets.get("gist_id")

    print(f"Pushing CQGXL quotes to GitHub Gist every {PUSH_INTERVAL_SECONDS}s. Ctrl+C to stop.")

    while True:
        quotes = import_cqg_quotes(QUOTES_PATH)
        content = json.dumps(quotes, indent=2)

        if gist_id:
            _update_gist(token, gist_id, content)
            print(f"Pushed {len(quotes)} quote(s) to gist {gist_id}")
        else:
            gist_id = _create_gist(token, content)
            _save_gist_id(gist_id)
            print(
                f"Created gist {gist_id} -- add gist_id to Streamlit Cloud "
                "secrets too, then redeploy."
            )

        time.sleep(PUSH_INTERVAL_SECONDS)


if __name__ == "__main__":
    main()
