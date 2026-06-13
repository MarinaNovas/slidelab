import uuid
import zlib
import base64
import requests
from pathlib import Path


# Кастомный алфавит для PlantUML
STANDARD_B64 = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
PLANTUML_B64 = b"0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-_"
ENCODE_TABLE = bytes.maketrans(STANDARD_B64, PLANTUML_B64)


class PlantUMLImageGenerator:
    def __init__(
        self,
        plantuml_server_url: str,
        images_dir: Path,
        public_base_url: str,
    ):
        self.plantuml_server_url = plantuml_server_url.rstrip("/")
        self.images_dir = images_dir
        self.images_dir.mkdir(parents=True, exist_ok=True)
        self.public_base_url = public_base_url.rstrip("/")

    @staticmethod
    def _encode_plantuml(text: str) -> str:
        """Сжимает UTF-8 текст методом DEFLATE и применяет PlantUML Base64 формат."""
        utf8_bytes = text.encode("utf-8")
        
        # 2. Сжатие методом DEFLATE (wbits=-15 удаляет zlib/gzip заголовки)
        compressor = zlib.compressobj(level=9, method=zlib.DEFLATED, wbits=-15)
        deflated_bytes = compressor.compress(utf8_bytes) + compressor.flush()
        
        # 3. Обработка 3-байтовых блоков для соответсвия PlantUML's Base64 формату
        b64_bytes = base64.b64encode(deflated_bytes)
        
        # 4. Преобразование симовлов Base64 в набор для PlantUML
        puml_bytes = b64_bytes.translate(ENCODE_TABLE)
        return puml_bytes.decode("ascii")

    def generate(
        self,
        plantuml_code: str,
        format: str = "png",
        timeout: int = 30,
    ) -> dict:
        image_id = uuid.uuid4().hex
        file_name = f"{image_id}.{format}"
        image_path = self.images_dir / file_name

        encoded = self._encode_plantuml(plantuml_code)

         # Пробуем получить TXT (ошибки будут текстом)
        txt_url = f"{self.plantuml_server_url}/plantuml/txt/{encoded}"
        
        try:
            response = requests.get(txt_url, timeout=30)
            
            # Если вернулся стутус 400 — это ошибка
            if response.status_code != 200 and "text/plain" in response.headers.get("content-type", ""):
                error_text = response.text.strip()
                if error_text and not error_text.startswith("@"):
                    return {
                        "status": "error",
                        "message": error_text,
                    }
        except Exception as e:
            return {"status": "error", "message": str(e)}

        url = f"{self.plantuml_server_url}/plantuml/{format}/{encoded}"

        response = requests.get(url, timeout=timeout)
        response.raise_for_status()

        image_path.write_bytes(response.content)

        return {
            "status": "ok",
            "image_id": image_id,
            "image_path": str(image_path),
            "image_url": f"{self.public_base_url}/media/{file_name}",
        }