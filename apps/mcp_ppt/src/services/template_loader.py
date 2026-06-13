from pathlib import Path
from urllib.parse import urlparse
from datetime import datetime
import uuid

import requests


class TemplateLoader:
    def __init__(self, templates_dir):
        self.templates_dir = templates_dir

    def _get_file_name_from_url(self, url: str) -> str | None:
        path = urlparse(url).path
        file_name = Path(path).name

        return file_name or None

    # примитивная нормализация и защита имени файла
    def _safe_file_name(self, file_name: str) -> str:
        safe_name = (
            file_name
            .replace("/", "_")
            .replace("\\", "_")
            .replace("..", "_")
            .strip()
        )

        return safe_name or "uploaded_template.pptx"

    def _make_unique_file_name(self, file_name: str) -> str:
        stem = Path(file_name).stem

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        short_id = uuid.uuid4().hex[:8]

        return f"{stem}_{timestamp}_{short_id}.pptx"

    def download(self, template_url: str) -> dict:
        if not template_url:
            return {
                "status": "error",
                "message": "template_url is required",
            }

        if not self.templates_dir.exists():
            return {
                "status": "error",
                "message": f"Templates directory does not exist: {self.templates_dir}",
            }

        safe_original_name = self._safe_file_name(self._get_file_name_from_url(template_url))


        if not safe_original_name.lower().endswith(".pptx"):
            return {
                "status": "error",
                "message": "Only .pptx templates are supported",
                "file_name": safe_original_name,
                "source_url": template_url,
            }

        file_name = self._make_unique_file_name(safe_original_name)
        template_path = self.templates_dir / file_name

        try:
            response = requests.get(
                template_url,
                stream = True,
                timeout = 30,
                headers = {
                    "User-Agent": "Mozilla/5.0",
                    "Accept": "application/octet-stream,*/*",
                }
            )
            response.raise_for_status()
            content_type = response.headers.get("content-type", "")

            with open(template_path, "wb") as file:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        file.write(chunk)


            if template_path.stat().st_size == 0:
                template_path.unlink(missing_ok=True)
                return {
                    "status": "error",
                    "message": "Downloaded template is empty",
                    "source_url": template_url,
                }

            return {
                "status": "ok",
                "template_name": file_name,
                "original_file_name": safe_original_name,
                "template_path": str(template_path),
                "size_bytes": template_path.stat().st_size,
                "source_url": template_url,
                "content_type": content_type,
            }

        except requests.RequestException as e:
            template_path.unlink(missing_ok=True)

            return {
                "status": "error",
                "message": f"Failed to download template: {str(e)}",
                "source_url": template_url,
            }

        except Exception as e:
            template_path.unlink(missing_ok=True)

            return {
                "status": "error",
                "message": str(e),
                "source_url": template_url,
            }



