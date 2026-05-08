from dataclasses import dataclass, field
from typing import Optional


@dataclass
class TitleSlideData:
    title: str
    subtitle: Optional[str] = None


@dataclass
class ContentSlideData:
    title: str
    content: list[str]

@dataclass
class TableSlideData:
    title: str
    headers: list[str]
    rows: list[list[str]]


@dataclass
class SectionSlideData:
    section_title: str
    background_color: Optional[str] = None

@dataclass
class ImageSlideData:
    title: str
    content: list[str]
    image_path: str
    subtitle: Optional[str] = None

@dataclass
class AgendaItemData:
    number: str
    title: str

@dataclass
class AgendaSlideData:
    title: str = "AGENDA"
    items: list[AgendaItemData] = field(default_factory=list)

@dataclass
class ComparisonTableSlideData:
    title: str
    sidebar_items: list[str] = field(default_factory=list)
    table_title: str = ""
    headers: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)

    def __post_init__(self):
        if not self.title:
            raise ValueError("title is required")

        if not self.headers:
            raise ValueError("headers must not be empty")

        if not self.rows:
            raise ValueError("rows must not be empty")

        for row in self.rows:
            if len(row) != len(self.headers):
                raise ValueError("Each row must have the same number of values as headers")
