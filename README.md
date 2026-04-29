uv init weather
cd weather

uv venv
.venv\Scripts\activate 

uv add "mcp[cli]" 
uv install "mcp[cli]"

new item server.py

uv run --with mcp server.py

npx -y @modelcontextprotocol/inspector