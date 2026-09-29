import requests

SOFR_API_URL = "https://markets.newyorkfed.org/api/rates/secured/sofr/last/1.json"
SOFR_PAGE_URL = "https://www.newyorkfed.org/markets/reference-rates/sofr"


def fetch_latest_sofr(timeout=5):
    """
    Latest published SOFR from the NY Fed's public Markets Data API.

    Returns {"rate": decimal (e.g. 0.0390), "effective_date": "YYYY-MM-DD",
    "source_url": SOFR_PAGE_URL}. Raises on any network/parsing failure --
    callers decide the fallback (e.g. keep the last known/manual value).
    """

    response = requests.get(SOFR_API_URL, timeout=timeout)
    response.raise_for_status()

    payload = response.json()
    latest = payload["refRates"][0]

    return {
        "rate": latest["percentRate"] / 100,
        "effective_date": latest["effectiveDate"],
        "source_url": SOFR_PAGE_URL,
    }
