from __future__ import annotations

import argparse
import base64
import hashlib
import http.server
import secrets
import threading
import webbrowser
from urllib.parse import parse_qs, urlencode, urlparse

from etsy_client import ENV_PATH, load_env, required, request_json, update_env_value


def make_verifier() -> str:
    return base64.urlsafe_b64encode(secrets.token_bytes(32)).rstrip(b"=").decode()


def make_challenge(verifier: str) -> str:
    digest = hashlib.sha256(verifier.encode()).digest()
    return base64.urlsafe_b64encode(digest).rstrip(b"=").decode()


def main() -> int:
    parser = argparse.ArgumentParser(description="Autoriza Etsy mediante OAuth 2.0 + PKCE")
    parser.add_argument("--no-browser", action="store_true", help="No abrir el navegador")
    args = parser.parse_args()

    values = load_env()
    client_id = required(values, "ETSY_API_KEY")
    redirect_uri = required(values, "ETSY_REDIRECT_URI")
    parsed = urlparse(redirect_uri)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise SystemExit("ETSY_REDIRECT_URI no es una URL válida")

    verifier = make_verifier()
    state = secrets.token_urlsafe(32)
    query = urlencode(
        {
            "response_type": "code",
            "client_id": client_id,
            "redirect_uri": redirect_uri,
            "scope": "listings_r listings_w shops_r shops_w",
            "state": state,
            "code_challenge": make_challenge(verifier),
            "code_challenge_method": "S256",
        }
    )
    auth_url = f"https://www.etsy.com/oauth/connect?{query}"
    result: dict[str, str] = {}
    event = threading.Event()

    class CallbackHandler(http.server.BaseHTTPRequestHandler):
        def do_GET(self) -> None:  # noqa: N802
            request_url = urlparse(self.path)
            if request_url.path != parsed.path:
                self.send_error(404)
                return
            params = parse_qs(request_url.query)
            result.update({key: values[0] for key, values in params.items() if values})
            body = b"Autorizacion recibida. Puedes cerrar esta ventana y volver a la terminal."
            self.send_response(200 if result.get("code") else 400)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write(body)
            event.set()

        def log_message(self, *_args: object) -> None:
            return

    server = http.server.ThreadingHTTPServer((parsed.hostname, parsed.port or 80), CallbackHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    print(f"Abriendo autorización de Etsy. Si no se abre, copia esta URL:\n{auth_url}")
    if not args.no_browser:
        webbrowser.open(auth_url)
    try:
        if not event.wait(timeout=300):
            raise SystemExit("Tiempo agotado esperando el callback de Etsy")
    finally:
        server.shutdown()
        server.server_close()

    if result.get("state") != state:
        raise SystemExit("El estado OAuth no coincide; autorización cancelada por seguridad")
    if result.get("error"):
        raise SystemExit(f"Etsy rechazó la autorización: {result['error']}")
    code = result.get("code")
    if not code:
        raise SystemExit("Etsy no devolvió un código de autorización")

    token_data = request_json(
        "POST",
        "https://api.etsy.com/v3/public/oauth/token",
        values,
        form=[
            ("grant_type", "authorization_code"),
            ("client_id", client_id),
            ("redirect_uri", redirect_uri),
            ("code", code),
            ("code_verifier", verifier),
        ],
    )
    refresh_token = str(token_data.get("refresh_token", ""))
    if not refresh_token:
        raise SystemExit("Etsy no devolvió refresh_token")
    update_env_value("ETSY_REFRESH_TOKEN", refresh_token)
    print(f"Refresh token guardado en {ENV_PATH} (valor omitido).")
    print(f"Scopes concedidos: {token_data.get('scope', '<no indicado>')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

