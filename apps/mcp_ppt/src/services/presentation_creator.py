import re
import uuid

class PresentationCreator:
    def generate_id(title: str) -> str:
        """
        Generate a unique presentation ID using cleaned title + UUID.

        Args:
            title: The presentation title

        Returns:
            A unique, clean presentation ID
        """
        cleaned_title = re.sub(r'[^\w\s\-]', '', title)
        cleaned_title = re.sub(r'[\s\-]+', '_', cleaned_title)
        cleaned_title = cleaned_title.strip('_').lower()
        if len(cleaned_title) > 20:
            cleaned_title = cleaned_title[:20].rstrip('_')
        short_uuid = uuid.uuid4().hex[:8]

        if cleaned_title:
            return f"{cleaned_title}_{short_uuid}"
        else:
            return f"presentation_{short_uuid}"
