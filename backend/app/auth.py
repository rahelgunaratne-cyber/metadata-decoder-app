"""
Authentication: verify Google Identity ID tokens and enforce the allowed
email domain.

The frontend signs the user in with Google Identity Services and sends the
resulting ID token as `Authorization: Bearer <token>` on every API call. Here
we verify the token's signature/audience against our OAuth client ID and check
that the email belongs to the allowed company domain (or the explicit
allowlist). When AUTH_ENABLED is false (local dev) everything is waved through
as a synthetic local user.
"""
from __future__ import annotations

from functools import lru_cache

from fastapi import Depends, Header, HTTPException, status

from .config import Settings, get_settings


class User(dict):
    @property
    def email(self) -> str:
        return self.get("email", "")


@lru_cache
def _google_request():
    import google.auth.transport.requests as gar

    return gar.Request()


def _verify_google_token(token: str, settings: Settings) -> User:
    from google.oauth2 import id_token as google_id_token

    try:
        claims = google_id_token.verify_oauth2_token(
            token, _google_request(), settings.oauth_client_id or None
        )
    except Exception as exc:  # invalid signature, expired, wrong audience, etc.
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"Invalid sign-in token: {exc}",
        )

    email = (claims.get("email") or "").lower()
    if not email or not claims.get("email_verified", False):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token has no verified email.",
        )

    domain = (claims.get("hd") or email.split("@")[-1]).lower()
    allowed = (
        email in settings.allowed_emails
        or (settings.allowed_email_domain and domain == settings.allowed_email_domain)
    )
    if not allowed:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail=(
                f"Access is restricted to @{settings.allowed_email_domain} accounts."
            ),
        )

    return User(
        email=email,
        name=claims.get("name", email),
        picture=claims.get("picture", ""),
        domain=domain,
    )


async def require_user(
    authorization: str | None = Header(default=None),
    settings: Settings = Depends(get_settings),
) -> User:
    if not settings.auth_enabled:
        return User(email="local@dev", name="Local Dev", picture="", domain="dev")

    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Missing bearer token.",
            headers={"WWW-Authenticate": "Bearer"},
        )
    token = authorization.split(" ", 1)[1].strip()
    return _verify_google_token(token, settings)
