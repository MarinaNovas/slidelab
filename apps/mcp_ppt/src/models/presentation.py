from dataclasses import dataclass
from pptx import Presentation

@dataclass
class PresentationModel:
    id: str
    title: str
    prs: Presentation
    template_name: str | None = None