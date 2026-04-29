import gradio as gr
import requests

from config import settings

MCP_PPT = settings.MCP_PPT_URL

def get_id(query):
    response = requests.post(
        f"{MCP_PPT}/",
        json={
            "arguments": {"query": query}
        }
    )
    return response.json()

def chat(message):
    result = get_id(message)
    return str(result)

gr.Interface(
    fn=chat,
    inputs="text",
    outputs="text"
).launch()