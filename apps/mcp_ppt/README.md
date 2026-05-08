Установка и запуск проекта
1. Клонировать репозиторий
git clone <repo_url>
cd slidelab
2. Создать виртуальное окружение
uv venv
3. Активировать окружение
Windows PowerShell
.venv\Scripts\activate
Linux / macOS
source .venv/bin/activate
4. Установить зависимости
uv sync
Запуск проекта

Перейти в папку:

cd apps/mcp_ppt
Запуск FastAPI

В первом терминале:

uv run uvicorn main:app --reload

FastAPI будет доступен по адресу:

http://127.0.0.1:8000

Swagger UI:

http://127.0.0.1:8000/docs
Запуск MCP сервера

Во втором терминале:

uv run --with mcp main_mcp.py

MCP endpoint:

http://127.0.0.1:8001/mcp

DOCKER
5. собрать docker образ
docker compose build --progress=plain
6. Запустить Docker контейнер 
docker compose up