uv init slidelab
cd slidelab

uv venv
.venv\Scripts\activate 

uv add "mcp[cli]" 
uv install "mcp[cli]"

new item server.py

uv run --with mcp main_mcp.py


npx -y @modelcontextprotocol/inspector

uv run uvicorn main:app --reload
taskkill /F /IM uv.exe   