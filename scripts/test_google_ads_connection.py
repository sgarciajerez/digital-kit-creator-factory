from __future__ import annotations

"""Safely verify local Google Ads OAuth and list accessible customer IDs (redacted)."""

import re
import sys
from pathlib import Path

import requests
from dotenv import dotenv_values, set_key
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"
TOKEN_URI = "https://oauth2.googleapis.com/token"
API_URL = "https://googleads.googleapis.com/v25/customers:listAccessibleCustomers"
SCOPE = "https://www.googleapis.com/auth/adwords"


def _masked_customer(resource_name: str) -> str:
    match = re.search(r"customers/(\d+)", resource_name)
    if not match:
        return "unknown"
    customer_id = match.group(1)
    return f"customers/{customer_id[:2]}…{customer_id[-2:]}"


def main() -> int:
    values = dotenv_values(ENV_PATH)
    required = ("GOOGLE_ADS_CLIENT_ID", "GOOGLE_ADS_CLIENT_SECRET", "GOOGLE_ADS_REFRESH_TOKEN")
    missing = [key for key in required if not str(values.get(key) or "").strip()]
    if missing:
        print("OAuth config incomplete; missing: " + ", ".join(missing))
        return 2

    credentials = Credentials(
        token=None,
        refresh_token=values["GOOGLE_ADS_REFRESH_TOKEN"],
        token_uri=TOKEN_URI,
        client_id=values["GOOGLE_ADS_CLIENT_ID"],
        client_secret=values["GOOGLE_ADS_CLIENT_SECRET"],
        scopes=[SCOPE],
    )
    try:
        credentials.refresh(Request())
    except Exception as error:
        print(f"OAuth refresh failed: {type(error).__name__}: {error}")
        return 3
    print("OAuth refresh: OK (access token not displayed)")

    try:
        response = requests.get(
            API_URL,
            headers={"Authorization": f"Bearer {credentials.token}"},
            timeout=30,
        )
    except requests.RequestException as error:
        print(f"Google Ads API network call failed: {type(error).__name__}: {error}")
        return 4

    if response.ok:
        resources = response.json().get("resourceNames", [])
        print(f"Google Ads API: OK; accessible accounts={len(resources)}")
        for resource_name in resources:
            print(_masked_customer(str(resource_name)))
        if len(resources) == 1:
            match = re.search(r"customers/(\d+)", str(resources[0]))
            if match:
                customer_id = match.group(1)
                configured_id = str(values.get("GOOGLE_ADS_CUSTOMER_ID") or "").replace("-", "").strip()
                if configured_id and configured_id != customer_id:
                    print("Did not change GOOGLE_ADS_CUSTOMER_ID: a different customer ID is already configured.")
                    return 6
                set_key(str(ENV_PATH), "GOOGLE_ADS_CUSTOMER_ID", customer_id, quote_mode="never")
                print("Saved the sole accessible customer ID to the local .env file (value not displayed).")
        return 0

    try:
        payload = response.json()
        error = payload.get("error", {})
        status = str(error.get("status") or "unknown")
        message = str(error.get("message") or "No error message returned")
    except ValueError:
        status, message = "unknown", "Non-JSON error response"
    print(f"Google Ads API: HTTP {response.status_code}; {status}: {message}")
    return 5


if __name__ == "__main__":
    raise SystemExit(main())
