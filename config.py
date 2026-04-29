from pydantic_settings import BaseSettings
from pathlib import Path

# определяем абсолютный путь до корня проекта
BASE_DIR = Path(__file__).resolve().parent.parent

class Settings(BaseSettings):
    @property
    def MCP_PPT_URL(self):
        return f"http://{self.LOCAL_HOST}:{self.MCP_PPT_PORT}"

settings = Settings()