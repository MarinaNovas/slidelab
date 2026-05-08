import base64
import io
import os
import sys
import uuid
import oss2

from src.services.presentation_store import PresentationStore


class PresentationExportService:
    def __init__(self, store: PresentationStore) -> None:
        self.store = store

    def save_to_oss(self, prs_id: str) -> str:
        presentation = self.store.get(prs_id)
        prs = presentation.prs

        tmp_path = None

        try:
            file_name = f"{prs_id}_{uuid.uuid4().hex[:8]}.pptx"
            tmp_path = os.path.join("/tmp", file_name)

            prs.save(tmp_path)

            endpoint = os.getenv("OSS_ENDPOINT")
            bucket_name = os.getenv("OSS_BUCKET_NAME")
            access_key = os.getenv("OSS_ACCESS_KEY")
            secret_key = os.getenv("OSS_SECRET_KEY")

            if not all([endpoint, bucket_name, access_key, secret_key]):
                raise ValueError("OSS configuration missing")

            auth = oss2.Auth(access_key, secret_key)
            bucket = oss2.Bucket(auth, endpoint, bucket_name)

            oss_object_name = f"presentations/{file_name}"
            bucket.put_object_from_file(oss_object_name, tmp_path)

            download_url = bucket.sign_url("GET", oss_object_name, 3600)

            return download_url

        except Exception as e:
            print(f"Error saving presentation: {e}", file=sys.stderr)
            raise

        finally:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)

    def export_to_base64(self, prs_id: str) -> dict:
        presentation = self.store.get(prs_id)
        prs = presentation.prs

        try:
            buffer = io.BytesIO()
            prs.save(buffer)
            buffer.seek(0)

            file_data = buffer.read()
            base64_data = base64.b64encode(file_data).decode("utf-8")

            return {
                "file_name": f"{prs_id}.pptx",
                "file_size_kb": round(len(file_data) / 1024, 2),
                "base64_data": base64_data,
            }

        except Exception as e:
            print(f"Error exporting presentation to base64: {e}", file=sys.stderr)
            raise