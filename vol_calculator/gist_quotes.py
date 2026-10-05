import json

import requests

_GIST_API = "https://api.github.com/gists/{gist_id}"


def fetch_gist_quotes(gist_id):
    """
    Reads the quotes.json file out of a GitHub Gist -- no auth needed for
    reads, even on a secret (unlisted) gist, as long as you have its ID.
    Returns the same [{"commodity", "contract", "price"}, ...] shape
    import_cqg_quotes produces.
    """

    resp = requests.get(_GIST_API.format(gist_id=gist_id), timeout=5)
    resp.raise_for_status()

    files = resp.json()["files"]
    content = files["quotes.json"]["content"]

    return json.loads(content)
