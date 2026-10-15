from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from dotenv import dotenv_values


PROJECT_ROOT = Path(__file__).resolve().parents[2]
ENV_PATH = PROJECT_ROOT / ".env"
API_BASE = "https://api.etsy.com/v3"
TOKEN_URL = f"{API_BASE}/public/oauth/token"


class EtsyApiError(RuntimeError):
    def __init__(self, method: str, url: str, status: int | None, body: str):
        self.method = method
        self.url = url
        self.status = status
        self.body = body
        suffix = f" HTTP {status}" if status else ""
        super().__init__(f"Etsy API error: {method} {url}{suffix}: {body[:500]}")


def load_env() -> dict[str, str]:
    values = {key: value or "" for key, value in dotenv_values(ENV_PATH).items()}
    values.update({key: value for key, value in os.environ.items() if key.startswith("ETSY_")})
    return values


def required(values: dict[str, str], key: str) -> str:
    value = values.get(key, "").strip()
    if not value:
        raise RuntimeError(f"Falta {key} en {ENV_PATH}")
    return value


def api_key(values: dict[str, str]) -> str:
    return f"{required(values, 'ETSY_API_KEY')}:{required(values, 'ETSY_API_SECRET')}"


def request_bytes(
    method: str,
    url: str,
    *,
    headers: dict[str, str] | None = None,
    body: bytes | None = None,
    timeout: int = 30,
) -> tuple[int, bytes]:
    request = Request(url, data=body, headers=headers or {}, method=method)
    try:
        with urlopen(request, timeout=timeout) as response:
            return int(response.status), response.read()
    except HTTPError as error:
        payload = error.read().decode("utf-8", errors="replace")
        raise EtsyApiError(method, url, error.code, payload) from error
    except URLError as error:
        raise EtsyApiError(method, url, None, str(error.reason)) from error


def request_json(
    method: str,
    path: str,
    values: dict[str, str],
    *,
    oauth_token: str | None = None,
    form: list[tuple[str, str]] | None = None,
    json_body: dict[str, Any] | None = None,
) -> dict[str, Any]:
    url = path if path.startswith("http") else f"{API_BASE}{path}"
    headers = {
        "Accept": "application/json",
        "x-api-key": api_key(values),
    }
    body: bytes | None = None
    if oauth_token:
        headers["Authorization"] = f"Bearer {oauth_token}"
    if form is not None:
        headers["Content-Type"] = "application/x-www-form-urlencoded"
        body = urlencode(form).encode("utf-8")
    elif json_body is not None:
        headers["Content-Type"] = "application/json"
        body = json.dumps(json_body).encode("utf-8")
    _, raw = request_bytes(method, url, headers=headers, body=body)
    if not raw:
        return {}
    return json.loads(raw.decode("utf-8"))


def refresh_access_token(values: dict[str, str]) -> dict[str, Any]:
    refresh_token = required(values, "ETSY_REFRESH_TOKEN")
    return request_json(
        "POST",
        TOKEN_URL,
        values,
        form=[
            ("grant_type", "refresh_token"),
            ("client_id", required(values, "ETSY_API_KEY")),
            ("refresh_token", refresh_token),
        ],
    )


def update_env_value(key: str, value: str) -> None:
    lines = ENV_PATH.read_text(encoding="utf-8").splitlines()
    replacement = f"{key}={value}"
    found = False
    output: list[str] = []
    for line in lines:
        if line.startswith(f"{key}="):
            output.append(replacement)
            found = True
        else:
            output.append(line)
    if not found:
        output.append(replacement)
    ENV_PATH.write_text("\n".join(output) + "\n", encoding="utf-8")


def multipart_body(
    fields: list[tuple[str, str]],
    file_field: str,
    file_path: Path,
    boundary: str,
    content_type: str = "application/octet-stream",
) -> bytes:
    chunks: list[bytes] = []
    marker = boundary.encode("ascii")
    for name, value in fields:
        chunks.extend(
            [
                b"--" + marker + b"\r\n",
                f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode(),
                value.encode("utf-8"),
                b"\r\n",
            ]
        )
    filename = file_path.name
    chunks.extend(
        [
            b"--" + marker + b"\r\n",
            (
                f'Content-Disposition: form-data; name="{file_field}"; '
                f'filename="{filename}"\r\n'
            ).encode(),
            f"Content-Type: {content_type}\r\n\r\n".encode("ascii"),
            file_path.read_bytes(),
            b"\r\n",
            b"--" + marker + b"--\r\n",
        ]
    )
    return b"".join(chunks)

