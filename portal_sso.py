"""Contrato de SSO do Portal Operacional para o módulo de Cadastros."""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import time
from urllib.parse import urlencode, urlsplit


def enabled() -> bool:
    return os.environ.get("ERP_PORTAL_SSO_ENABLED", "").strip().casefold() in {"1", "true", "yes", "sim", "on"}


def normalize_next(value: str | None, default: str = "/") -> str:
    candidate = str(value or "").strip()
    parsed = urlsplit(candidate)
    if not candidate.startswith("/") or candidate.startswith("//") or parsed.scheme or parsed.netloc:
        return default
    return candidate


def _portal_url() -> str:
    return os.environ.get("ERP_PORTAL_URL", "https://ji-portal-operacional.onrender.com").strip().rstrip("/")


def portal_login_url(app_code: str, next_path: str | None) -> str:
    return f"{_portal_url()}/login?{urlencode({'app': app_code, 'next': normalize_next(next_path)})}"


def portal_logout_url() -> str:
    return f"{_portal_url()}/logout"


def _b64decode(value: str) -> bytes:
    return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))


def consume_ticket(ticket: str | None, expected_app: str) -> dict:
    secret = os.environ.get("ERP_PORTAL_SSO_SECRET", "").encode("utf-8")
    raw_ticket = str(ticket or "")
    if not secret or not raw_ticket or len(raw_ticket) > 4096 or raw_ticket.count(".") != 1:
        raise ValueError("Comprovante de acesso inválido.")
    encoded, signature = raw_ticket.rsplit(".", 1)
    expected = hmac.new(secret, encoded.encode("ascii"), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(signature, expected):
        raise ValueError("Comprovante de acesso inválido.")
    try:
        claims = json.loads(_b64decode(encoded).decode("utf-8"))
        issued_at, expires_at = int(claims["iat"]), int(claims["exp"])
        claims["uid"] = int(claims["uid"])
        claims["auth_version"] = int(claims["auth_version"])
    except (ValueError, TypeError, KeyError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("Comprovante de acesso incompleto.") from exc
    now = int(time.time())
    if str(claims.get("app") or "").upper() != expected_app.upper() or claims["uid"] <= 0 or claims["auth_version"] < 0 or issued_at > now + 30 or expires_at < now or expires_at - issued_at > 180:
        raise ValueError("Comprovante de acesso expirado.")
    username = str(claims.get("username") or "").strip()
    if not username or len(username) > 128:
        raise ValueError("Comprovante de acesso inválido.")
    claims["username"] = username
    claims["next"] = normalize_next(claims.get("next"), "/cadastro/bancos")
    return claims
