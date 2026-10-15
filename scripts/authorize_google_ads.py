from __future__ import annotations

"""Authorize this local factory for Google Ads and store OAuth secrets in .env."""

import argparse
import json
from pathlib import Path

from dotenv import dotenv_values, set_key
from google_auth_oauthlib.flow import InstalledAppFlow

PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"
SCOPE = "https://www.googleapis.com/auth/adwords"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--credentials", type=Path, required=True, help="Ruta local al JSON OAuth de app de escritorio")
    args = parser.parse_args()
    credentials_path = args.credentials.expanduser().resolve()
    if not credentials_path.is_file():
        parser.error(f"No existe el JSON de credenciales: {credentials_path}")

    oauth_config = json.loads(credentials_path.read_text(encoding="utf-8"))
    client_config = oauth_config.get("installed")
    if not client_config or not client_config.get("client_id") or not client_config.get("client_secret"):
        parser.error("El JSON no parece ser una credencial OAuth de tipo app de escritorio (installed).")

    current_values = dotenv_values(ENV_PATH)
    existing_client_id = str(current_values.get("GOOGLE_ADS_CLIENT_ID") or "").strip()
    if existing_client_id and existing_client_id != client_config["client_id"]:
        parser.error(".env ya contiene otro GOOGLE_ADS_CLIENT_ID; no lo he sobrescrito.")

    flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), scopes=[SCOPE])
    credentials = flow.run_local_server(
        host="localhost",
        port=0,
        open_browser=False,
        authorization_prompt_message="Abre esta URL en tu navegador e inicia sesión con la cuenta de Google Ads:\n{url}\n",
        success_message="Autorización recibida. Puedes cerrar esta pestaña.",
        prompt="consent",
    )
    if not credentials.refresh_token:
        raise RuntimeError("Google no devolvió refresh_token. Revoca el acceso de esta app y vuelve a autorizarla con consentimiento.")

    ENV_PATH.touch(exist_ok=True)
    for key, value in (
        ("GOOGLE_ADS_CLIENT_ID", client_config["client_id"]),
        ("GOOGLE_ADS_CLIENT_SECRET", client_config["client_secret"]),
        ("GOOGLE_ADS_REFRESH_TOKEN", credentials.refresh_token),
    ):
        set_key(str(ENV_PATH), key, value, quote_mode="always")
    print("OAuth completado: client ID, client secret y refresh token guardados en el .env local. No se han mostrado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
