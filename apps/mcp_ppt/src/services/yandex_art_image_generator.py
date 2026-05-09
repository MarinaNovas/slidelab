import uuid
from pathlib import Path
from typing import Optional
from yandex_ai_studio_sdk import AIStudio


class YandexArtImageGenerator:
    def __init__(
        self,
        folder_id: str,
        auth: str,
        images_dir: Path,
        public_base_url: str,
    ):
        self.images_dir = images_dir
        self.images_dir.mkdir(parents=True, exist_ok=True)

        self.public_base_url = public_base_url.rstrip("/")

        self.sdk = AIStudio(
            folder_id=folder_id,
            auth=auth,
        )

        self.model = self.sdk.models.image_generation("yandex-art")

    def generate(
        self,
        prompt: str,
        style: Optional[str] = None,
        width_ratio: int = 1,
        height_ratio: int = 1,
        seed: Optional[int] = None,
    ) -> dict:
        image_id = uuid.uuid4().hex
        file_name = f"{image_id}.jpeg"
        image_path = self.images_dir / file_name

        model = self.model.configure(
            width_ratio=width_ratio,
            height_ratio=height_ratio,
            seed=seed,
        )

        messages: list[str] = [prompt]

        if style:
            messages.append(style)

        operation = model.run_deferred(messages)
        result = operation.wait()

        image_path.write_bytes(result.image_bytes)

        return {
            "status": "ok",
            "image_id": image_id,
            "image_path": str(image_path),
            "image_url": f"{self.public_base_url}/media/{file_name}",
            "prompt": prompt,
            "style": style,
        }