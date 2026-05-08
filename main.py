# This is a sample Python script.
from client.client_demo import demo

# Запустите веб-интерфейс Gradio и сервер MCP
if __name__ == "__main__":
 demo.launch(mcp_server=True)
