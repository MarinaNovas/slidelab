from fastapi import FastAPI, Request
from fastapi.responses import FileResponse
from pathlib import Path

EXPORTS_DIR = Path("exports")
EXPORTS_DIR.mkdir(exist_ok=True)

app = FastAPI()


@app.get("/exports/{file_name}")
def download_export(file_name: str):
    path = EXPORTS_DIR / file_name

    if not path.exists():
        return {"error": "File not found"}

    return FileResponse(
        path,
        filename=file_name,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

@app.middleware("http")
async def debug_headers(request: Request, call_next):
    print("HEADERS:", dict(request.headers))
    response = await call_next(request)
    return response