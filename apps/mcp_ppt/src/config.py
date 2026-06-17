from pydantic_settings import BaseSettings, SettingsConfigDict
from pathlib import Path

# определяем абсолютный путь до корня проекта
BASE_DIR = Path(__file__).resolve().parents[1]
TEMPLATES_DIR = BASE_DIR / "templates"
DEFAULT_TEMPLATE = TEMPLATES_DIR / "template_axenix.pptx"
MEDIA_DIR = BASE_DIR / "media"


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file = BASE_DIR / ".env",
        env_file_encoding = "utf-8",
        extra = "ignore",
    )

    LOCAL_HOST: str
    MCP_PPT_PORT: int
    YANDEX_API_KEY: str
    YANDEX_FOLDER_ID: str
    PUBLIC_BASE_URL: str
    API_PUBLIC_BASE_URL: str
    PLANTUML_SERVER_URL: str

    OAUTH_ENABLED: bool = False
    OAUTH_CLIENT_ID: str
    OAUTH_CLIENT_SECRET: str
    OAUTH_AUTH_BASE_URL: str
    OAUTH_REQUIRED_SCOPE: str = "mcp:tools"
    OAUTH_REQUIRED_AUDIENCE: str
    OAUTH_RESOURCE_METADATA_URL: str

    @property
    def MCP_PPT_URL(self):
        return f"http://{self.LOCAL_HOST}:{self.MCP_PPT_PORT}/mcp"

    @property
    def YANDEX_ART_CONFIG(self):
        return {
        "folder_id": self.YANDEX_FOLDER_ID,
        "auth": self.YANDEX_API_KEY,
        "images_dir": MEDIA_DIR,
        "public_base_url": self.PUBLIC_BASE_URL,
        }

    @property
    def UML_CONFIG(self):
        return {
            "plantuml_server_url": self.PLANTUML_SERVER_URL,
            "images_dir": MEDIA_DIR,
            "public_base_url": self.PUBLIC_BASE_URL,
        }

    @property
    def OAUTH_INTROSPECTION_URL(self) -> str:
        return f"{self.OAUTH_AUTH_BASE_URL}/protocol/openid-connect/token/introspect"

    @property
    def OAUTH_PROTECTED_RESOURCE_METADATA(self) -> dict:
        return {
            "resource": self.OAUTH_REQUIRED_AUDIENCE,
            "authorization_servers": [self.OAUTH_AUTH_BASE_URL],
            "scopes_supported": [self.OAUTH_REQUIRED_SCOPE],
        }

    @property
    def OAUTH_WWW_AUTHENTICATE(self) -> str:
        return f'Bearer resource_metadata="{self.OAUTH_RESOURCE_METADATA_URL}"'

settings = Settings()