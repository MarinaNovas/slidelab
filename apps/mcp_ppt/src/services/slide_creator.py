from datetime import date
from pathlib import Path

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

from src.models.slide import (
    AgendaSlideData, ComparisonTableSlideData, ImageSlideData, TitleSlideData,
    ContentSlideData,
    SectionSlideData,
    TableSlideData,
)
from src.services.presentation_store import PresentationStore


class SlideCreator:
    def __init__(self, store: PresentationStore) -> None:
        self.store = store

    def add_title_slide(self, prs_id: str, data: TitleSlideData) -> int:
        presentation = self.store.get(prs_id)
        prs = presentation.prs

        master = prs.slide_masters[0]
        slide = prs.slides.add_slide(master.slide_layouts[0])
        title_ph = self._get_placeholder(slide, 0)
        body_1_ph = self._get_placeholder(slide, 14)
        body_2_ph = self._get_placeholder(slide, 15)

        title_ph.text = data.title
        if data.subtitle:
            body_1_ph.text = data.subtitle
        body_2_ph.text = str(date.today())
        #slide.shapes.title.text = data.title

        #if data.subtitle and len(slide.placeholders) > 1:
        #    slide.placeholders[1].text = data.subtitle

        return len(prs.slides)

    def add_content_slide(self, prs_id: str, data: ContentSlideData) -> int:
        presentation = self.store.get(prs_id)
        prs = presentation.prs

        master = prs.slide_masters[1]
        slide = prs.slides.add_slide(master.slide_layouts[0])

        title_ph = self._get_placeholder(slide, 0)
        body_1_ph = self._get_placeholder(slide, 14)

        title_ph.text = data.title
        body_1_ph.text = "\n".join(data.content)


        return len(prs.slides)

    def add_image_content_slide(self, prs_id: str, data: ImageSlideData) -> int:
        presentation = self.store.get(prs_id)
        prs = presentation.prs

        master = prs.slide_masters[2]
        slide = prs.slides.add_slide(master.slide_layouts[4])

        title_ph = self._get_placeholder(slide, 0)
        body_ph = self._get_placeholder(slide, 10)
        object_ph = self._get_placeholder(slide, 2)
        picture_ph = self._get_placeholder(slide, 13)

        title_ph.text = data.title
        body_ph.text = data.subtitle
        print(f"{data.content}=")
        tf = object_ph.text_frame
        tf.clear()  # очищаем дефолтный текст

        for i, item in enumerate(data.content):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            p.text = item
            p.level = 0  # уровень bullet (0 = основной)

        if data.image_path:
            image_path = Path(data.image_path) #2bf37c60d1ae44c6b9f495035cc40eb0

            if not image_path.exists():
                raise ValueError(f"Image not found: {image_path}")

            picture_ph.insert_picture(str(image_path))

        return len(prs.slides)

    def add_section_slide(self, prs_id: str, data: SectionSlideData) -> int:
        presentation = self.store.get(prs_id)
        prs = presentation.prs

        master = prs.slide_masters[3]
        slide = prs.slides.add_slide(master.slide_layouts[1])

        title_ph = self._get_placeholder(slide, 0)

        title_ph.text = data.section_title

        return len(prs.slides)

    def add_table_slide(self, prs_id: str, data: TableSlideData) -> int:
        presentation = self.store.get(prs_id)
        prs = presentation.prs

        slide = prs.slides.add_slide(prs.slide_layouts[13])

        title_shape = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.5), Inches(9), Inches(1)
        )
        title_shape.text = data.title
        title_shape.text_frame.paragraphs[0].font.size = Pt(28)
        title_shape.text_frame.paragraphs[0].font.bold = True

        rows_count = len(data.rows) + 1
        cols_count = len(data.headers)

        table = slide.shapes.add_table(
            rows_count,
            cols_count,
            Inches(1),
            Inches(2),
            Inches(8),
            Inches(0.5 * rows_count),
        ).table

        for index, header in enumerate(data.headers):
            cell = table.cell(0, index)
            cell.text = header
            cell.text_frame.paragraphs[0].font.bold = True
            cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        for row_index, row in enumerate(data.rows, start=1):
            for col_index, value in enumerate(row):
                if col_index < cols_count:
                    table.cell(row_index, col_index).text = str(value)

        return len(prs.slides)

    def add_thank_you_slide(self, prs_id: str) -> int:
        presentation = self.store.get(prs_id)

        if presentation is None:
            raise ValueError(f"Presentation '{prs_id}' not found")

        prs = presentation.prs

        master = prs.slide_masters[4]
        slide = prs.slides.add_slide(master.slide_layouts[0])
        return len(prs.slides)

    def add_agenda_slide(self, prs_id: str, data: AgendaSlideData) -> int:
        presentation = self.store.get(prs_id)

        if presentation is None:
            raise ValueError(f"Presentation '{prs_id}' not found")

        prs = presentation.prs

        master = prs.slide_masters[0]
        slide = prs.slides.add_slide(master.slide_layouts[1])

        title_ph = self._get_placeholder(slide, 10)
        title_ph.text = data.title or "AGENDA"

        number_placeholder_ids = [34, 35, 36, 44, 45, 46]
        title_placeholder_ids = [54, 55, 56, 59, 60, 61]

        for i, item in enumerate(data.items[:6]):
            number_ph = self._get_placeholder(slide, number_placeholder_ids[i])
            item_title_ph = self._get_placeholder(slide, title_placeholder_ids[i])

            number_ph.text = str(item.number or "")
            item_title_ph.text = str(item.title or "")

        return len(prs.slides)

    def add_comparison_table_slide(
            self,
            prs_id: str,
            data: ComparisonTableSlideData,
    ) -> int:
        presentation = self.store.get(prs_id)

        if presentation is None:
            raise ValueError(f"Presentation '{prs_id}' not found")

        prs = presentation.prs

        # MASTER 0, layout 8: 2_Comparison
        master = prs.slide_masters[2]
        slide = prs.slides.add_slide(master.slide_layouts[8])

        table_title_ph = self._get_placeholder(slide, 1)
        sidebar_ph = self._get_placeholder(slide, 10)
        title_ph = self._get_placeholder(slide, 0)
        object_ph = self._get_placeholder(slide, 2)

        title_ph.text = data.title or ""
        table_title_ph.text = data.table_title or ""

        sidebar_ph.text = "\n".join(str(item) for item in data.sidebar_items if item is not None)

        rows_count = len(data.rows) + 1
        cols_count = len(data.headers)

        graphic_frame = slide.shapes.add_table(
            rows_count,
            cols_count,
            object_ph.left,
            object_ph.top,
            object_ph.width,
            object_ph.height,
        )

        table = graphic_frame.table

        # header row
        for col_idx, header in enumerate(data.headers):
            cell = table.cell(0, col_idx)
            cell.text = str(header)

            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.size = Pt(12)

        # body rows
        for row_idx, row in enumerate(data.rows, start = 1):
            for col_idx, value in enumerate(row):
                cell = table.cell(row_idx, col_idx)
                cell.text = str(value)

                for paragraph in cell.text_frame.paragraphs:
                    paragraph.alignment = PP_ALIGN.LEFT
                    for run in paragraph.runs:
                        run.font.size = Pt(10)

        return len(prs.slides)

    def _get_placeholder(self, slide, idx: int):
        for placeholder in slide.placeholders:
            if placeholder.placeholder_format.idx == idx:
                return placeholder

        raise ValueError(f"Placeholder with idx={idx} not found")

    def _set_background_color(self, slide, color: str) -> None:
        hex_color = color.lstrip("#")
        r, g, b = tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))

        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(r, g, b)