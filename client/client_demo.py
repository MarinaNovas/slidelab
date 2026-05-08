from config import settings

MCP_PPT = settings.MCP_PPT_URL

import gradio as gr
import asyncio

from mcp import ClientSession
from mcp.client.streamable_http import streamable_http_client

from config import settings

MCP_PPT = settings.MCP_PPT_URL  # например: http://127.0.0.1:8001/mcp

async def call_create_presentation(title: str):
    async with streamable_http_client(MCP_PPT) as (read, write, _):
        async with ClientSession(read, write) as session:
            await session.initialize()

            result = await session.call_tool(
                "create_presentation",
                arguments={
                    "title": title
                }
            )

            return result.content[0].text


def chat(message):
    return asyncio.run(call_create_presentation(message))


demo = gr.Interface(
    fn=chat,
    inputs="text",
    outputs="text"
)