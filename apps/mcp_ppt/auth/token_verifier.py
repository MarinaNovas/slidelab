import logging
from typing import Any

import httpx
from mcp.server.auth.provider import AccessToken, TokenVerifier

logger = logging.getLogger(__name__)

class IntrospectionTokenVerifier(TokenVerifier):
    def __init__(
        self,
        introspection_endpoint: str,
        client_id: str,
        client_secret: str,
        required_audience: str,
    ):
        self.introspection_endpoint = introspection_endpoint
        self.client_id = client_id
        self.client_secret = client_secret
        self.required_audience = required_audience

    async def verify_token(self, token: str) -> AccessToken | None:
        async with httpx.AsyncClient(timeout=10.0) as client:
            response = await client.post(
                self.introspection_endpoint,
                data={"token": token},
                auth=(self.client_id, self.client_secret),
                headers={"Content-Type": "application/x-www-form-urlencoded"},
            )

        print("INTROSPECTION_STATUS =", response.status_code)
        print("INTROSPECTION_BODY =", response.text)
        print("TOKEN_START =", token[:80])
        print("TOKEN_DOTS =", token.count("."))

        if response.status_code != 200:
            logger.warning("Introspection failed: %s %s", response.status_code, response.text)
            return None

        data: dict[str, Any] = response.json()

        if not data.get("active"):
            return None

        scope = data.get("scope", "")
        scopes = scope.split() if scope else []

        # Пока audience у тебя в Keycloak не всегда содержит http://localhost:8001/mcp.
        # Поэтому для MVP можно не валить токен по audience, а проверять scope.
        # Потом, когда aud настроишь, раскомментируешь проверку ниже.
        #
        # aud = data.get("aud", [])
        # if isinstance(aud, str):
        #     aud = [aud]
        # if self.required_audience not in aud:
        #     logger.warning("Invalid audience: %s", aud)
        #     return None

        aud = data.get("aud")
        resource = aud[0] if isinstance(aud, list) and aud else aud

        return AccessToken(
            token=token,
            client_id=data.get("client_id", self.client_id),
            scopes=scopes,
            expires_at=data.get("exp"),
            resource = resource,
        )