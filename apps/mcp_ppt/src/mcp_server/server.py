from mcp.server.fastmcp import FastMCP
from typing import Optional

from src.api.app import EXPORTS_DIR
from src.config import MEDIA_DIR, settings
from src.models.slide import AgendaItemData, AgendaSlideData, ComparisonTableSlideData, ContentSlideData, \
    ImageSlideData, SectionSlideData, \
    TableSlideData, TitleSlideData
from src.services.presentation_creator import PresentationCreator
from src.services.presentation_store import PresentationStore
from src.services.slide_creator import SlideCreator
from src.services.yandex_art_image_generator import YandexArtImageGenerator

# Create an MCP server
mcp = FastMCP("PowerPoint Creator",  dependencies=["python-pptx","requests"], port=8001)

store = PresentationStore()
presentation_service = PresentationCreator(store)
slide_service = SlideCreator(store)
image_service = YandexArtImageGenerator(**settings.YANDEX_ART_CONFIG)
print(settings.YANDEX_ART_CONFIG)

@mcp.tool()
def create_presentation(title: str,  template_name: Optional[str] = None,) -> str:
    """
    Create a new PowerPoint presentation.
    Use this tool FIRST before adding any slides.

    Workflow:
    1. Call this tool to create a presentation
    2. Use returned prs_id in add_* tools
    3. Call save_presentation at the end

    Args:
        title: Presentation title
        template_name: Optional template file name from templates directory.
                       If not provided, default.pptx is used.


    Returns:
        prs_id: unique identifier for the presentation
    """
    try:
        prs_id = presentation_service.create(title, template_name)
        return prs_id
    except Exception as e:
        return f"Error creating presentation: {str(e)}"

@mcp.tool()
def add_title_slide(prs_id: str, title: str, subtitle: Optional[str] = None) -> str:
    """
       Add a title slide (first slide) to an existing presentation.

       Use this tool AFTER create_presentation.
       Typically this should be the first slide in the presentation.

       Args:
           prs_id: Presentation ID returned by create_presentation
           title: Main title text
           subtitle: Optional subtitle text

       Returns:
           Slide number (position in the presentation).

       Important:
           - Always pass the correct prs_id
           - Usually called once at the beginning
           - After this, use add_content_slide or other slide tools
    """
    try:
        slide_number = slide_service.add_title_slide(
            prs_id,
            TitleSlideData(title = title, subtitle = subtitle),
        )
        return f"Added title slide at position {slide_number}"
    except Exception as e:
        return f"Error adding title slide: {str(e)}"

@mcp.tool()
def add_table_slide(
    prs_id: str,
    title: str,
    headers: list[str],
    rows: list[list[str]],
) -> str:
    """
        Add a table slide to an existing PowerPoint presentation.

        Use this tool when the slide content is structured data that should be
        displayed as a table: metrics, comparisons, financial results, plans,
        timelines, KPIs, product features, or any row/column-based information.

        Use this tool AFTER create_presentation and after receiving a valid prs_id.

        Args:
            prs_id: Presentation ID returned by create_presentation.
            title: Slide title displayed above the table.
            headers: List of column names for the table header row.
            rows: Table data. Each inner list is one table row.
                  The number of values in each row should match the number of headers.

        Returns:
            Slide number / position of the newly added table slide.

        Important:
            - Use this tool only for tabular data.
            - Do not use it for normal bullet-point content; use add_content_slide instead.
            - headers must not be empty.
            - rows should contain at least one row.
            - Keep table text short so it fits on the slide.
            - After adding all slides, call save_presentation to export the file.

        Example:
            headers = ["Metric", "2024", "2025"]
            rows = [
                ["Revenue", "120M", "150M"],
                ["Profit", "18M", "25M"]
            ]
    """
    try:
        slide_number = slide_service.add_table_slide(
            prs_id,
            TableSlideData(
                title=title,
                headers=headers,
                rows=rows,
            ),
        )
        return f"Added table slide at position {slide_number}"
    except Exception as e:
        return f"Error adding table slide: {str(e)}"


@mcp.tool()
def get_presentation_info(prs_id: str) -> str:
    """
       Get basic information about an existing presentation.

       Use this tool to check the current state of a presentation,
       such as how many slides it contains and verify that it exists.

       This is useful when:
           - You are unsure whether a presentation was already created
           - You want to check progress before adding more slides
           - You need to confirm the presentation structure

       Args:
           prs_id: Presentation ID returned by create_presentation

       Returns:
           Text information including:
               - presentation ID
               - title
               - number of slides

       Important:
           - Use only with a valid prs_id
           - Do not use this tool to modify the presentation
           - This tool is for inspection only (read-only)
           - After checking info, you can continue with add_* tools

       Workflow:
           create_presentation → get_presentation_info (optional) → add slides → save_presentation
    """
    try:
        return presentation_service.get_info(prs_id)
    except Exception as e:
        return f"Error getting presentation info: {str(e)}"

@mcp.tool()
def save_presentation(prs_id: str) -> str:
    """
       Save a completed presentation as a .pptx file and return a download URL.

       Use this tool ONLY after all slides have been added and the presentation is complete.

       Args:
           prs_id: The presentation ID returned by create_presentation.

       Returns:
           A direct download URL to the generated .pptx file.

       Behavior:
           - Saves the presentation file locally on the server.
           - Makes the file доступным через HTTP endpoint (/exports/{file_name}).
           - Returns a URL that can be opened in a browser to download the file.

       Important:
           - This is the FINAL step in the presentation workflow.
           - Do NOT call this before adding slides (unless an empty presentation is explicitly requested).
           - Always call this after finishing slide creation.
           - The returned link is immediately usable for download.

       Example workflow:
           1. create_presentation
           2. add_slide / add_content / add_image
           3. save_presentation
    """
    current_prs = store.get(prs_id)
    if current_prs is None:
        return f"Error: Presentation '{prs_id}' not found"

    file_name = f"{prs_id}.pptx"
    path = EXPORTS_DIR / file_name
    current_prs.prs.save(path)
    return f"Presentation saved. Download URL: http://127.0.0.1:8000/exports/{file_name}"

@mcp.tool()
def add_content_slide(
    prs_id: str,
    title: str,
    content: list[str],
) -> str:
    """
    Add a normal content slide with bullet points.

    Use this tool for text slides: key ideas, arguments, conclusions,
    explanations, agenda items, recommendations.

    Args:
        prs_id: Presentation ID returned by create_presentation.
        title: Slide title.
        content: List of bullet points.

    Returns:
        Slide number / position of the newly added content slide.

    Important:
        - Use this tool for normal text/bullet slides.
        - Do not use it for tables; use add_table_slide instead.
        - Keep bullet points short so they fit on the slide.
        - After adding all slides, call save_presentation.
    """
    try:
        slide_number = slide_service.add_content_slide(
            prs_id,
            ContentSlideData(
                title=title,
                content=content,
            ),
        )
        return f"Added content slide at position {slide_number}"
    except Exception as e:
        return f"Error adding content slide: {str(e)}"

@mcp.tool()
def add_section_slide(
    prs_id: str,
    section_title: str,
    background_color: Optional[str] = None,
) -> str:
    """
    Add a section divider slide to an existing presentation.

    Use this tool to separate major parts of the presentation:
    introduction, market analysis, financial results, roadmap, conclusions.

    Args:
        prs_id: Presentation ID returned by create_presentation.
        section_title: Main section title displayed in the center.
        background_color: Optional HEX background color, for example "#000080".

    Returns:
        Slide number / position of the newly added section slide.

    Important:
        - Use this tool only for section divider slides.
        - background_color should be HEX format.
        - After adding all slides, call save_presentation.
    """
    try:
        slide_number = slide_service.add_section_slide(
            prs_id,
            SectionSlideData(
                section_title=section_title,
                background_color=background_color,
            ),
        )
        return f"Added section slide at position {slide_number}"
    except Exception as e:
        return f"Error adding section slide: {str(e)}"

@mcp.tool()
def generate_yandex_art_image(
    prompt: str,
    style: Optional[str] = None,
    width_ratio: int = 1,
    height_ratio: int = 1,
    seed: Optional[int] = None,
) -> dict:
    """
    Generate an image using YandexART and save it locally.

    Use this tool when a presentation needs a generated image,
    illustration, concept visual, background, or slide picture.

    Args:
        prompt: Main image description.
        style: Optional style description, for example "realistic", "business illustration", "Miyazaki style".
        width_ratio: Image width ratio.
        height_ratio: Image height ratio.
        seed: Optional seed for reproducible generation.

    Returns:
        Dict with:
        - status
        - image_id
        - image_path
        - image_url
        - prompt
        - style

    Workflow:
        1. Call create_presentation
        2. Call generate_yandex_art_image
        3. Pass returned image_id to add_image_content_slide
        4. Call save_presentation
    """
    try:
        return image_service.generate(
            prompt=prompt,
            style=style,
            width_ratio=width_ratio,
            height_ratio=height_ratio,
            seed=seed,
        )

    except Exception as e:
        return {
            "status": "error",
            "message": str(e),
        }

@mcp.tool()
def add_image_content_slide(
    prs_id: str,
    title: str,
    subtitle: str,
    content: list[str],
    image_id: str,
) -> str:
    """
    Add a slide with title, bullet points and previously generated image.

    Use this tool after generate_yandex_art_image.
    Pass image_id returned by generate_yandex_art_image.
    """
    try:
        image_path = MEDIA_DIR / f"{image_id}.jpeg"

        if not image_path.exists():
            return f"Error: image '{image_id}' not found"

        slide_number = slide_service.add_image_content_slide(
            prs_id,
            ImageSlideData(
            title=title,
            subtitle=subtitle,
            content=content,
            image_path=str(image_path))
        )

        return f"Added image content slide at position {slide_number}"

    except Exception as e:
        return f"Error adding image content slide: {str(e)}"

@mcp.tool()
def add_comparison_table_slide(
    prs_id: str,
    title: str,
    headers: list[str],
    rows: list[list[str]],
    sidebar_items: list[str] | None = None,
    table_title: str = "",
) -> str:
    """
    Add a comparison slide with a left sidebar and a table in the main content area.

    Use this tool when the presentation needs to compare several options,
    approaches, products, technologies, metrics, risks, features, or scenarios.

    This slide layout is useful for:
    - feature comparison
    - technology comparison
    - pros and cons
    - vendor comparison
    - project alternatives
    - decision matrices
    - structured analytical summaries

    Args:
        prs_id: Presentation ID returned by create_presentation.
        title: Main slide headline.
        sidebar_items: Optional list of short supporting points shown in the orange left sidebar.
        table_title: Optional title above the table.
        headers: Table column headers.
        rows: Table rows. Each row must have the same number of values as headers.

    Returns:
        Slide number / position of the newly added comparison table slide.

    Important:
        - Use this tool only for comparison or structured table content.
        - Keep headers short.
        - Keep cell text short so the table fits the slide.
        - Best table size: 3–5 columns and 3–7 rows.
        - Do not use this tool for ordinary bullet slides.
        - After adding all slides, call save_presentation.
    """
    try:
        slide_number = slide_service.add_comparison_table_slide(
            prs_id,
            ComparisonTableSlideData(
                title=title,
                sidebar_items=sidebar_items or [],
                table_title=table_title,
                headers=headers,
                rows=rows,
            ),
        )

        return f"Added comparison table slide at position {slide_number}"

    except Exception as e:
        return f"Error adding comparison table slide: {str(e)}"

@mcp.tool()
def add_agenda_slide(
    prs_id: str,
    items: list[str],
    title: str = "AGENDA",
) -> str:
    """
    Add an agenda slide with up to 6 numbered agenda items.

    Use this tool near the beginning of the presentation, usually after
    the title slide and before the main content sections.

    This slide is intended for:
    - presentation structure
    - lesson plan
    - meeting agenda
    - roadmap overview
    - training module outline

    Args:
        prs_id: Presentation ID returned by create_presentation.
        items: List of agenda item titles. Maximum 6 items.
               Each item will be automatically numbered.
        title: Main agenda slide title. Usually "AGENDA" or "ПОВЕСТКА".

    Returns:
        Slide number / position of the newly added agenda slide.

    Important:
        - Use only for agenda / structure overview slides.
        - Maximum 6 agenda items.
        - Keep item titles short, ideally 2-5 words.
        - Do not use this tool for normal bullet-point content.
    """
    try:
        agenda_items = [
            AgendaItemData(
                number=str(i + 1),
                title=item,
            )
            for i, item in enumerate(items[:6])
        ]

        slide_number = slide_service.add_agenda_slide(
            prs_id,
            AgendaSlideData(
                title=title,
                items=agenda_items,
            ),
        )

        return f"Added agenda slide at position {slide_number}"

    except Exception as e:
        return f"Error adding agenda slide: {str(e)}"

@mcp.tool()
def add_thank_you_slide(
    prs_id: str,
) -> str:
    """
    Add a final thank-you / closing slide to the presentation.

    Use this tool as the LAST content slide before save_presentation.

    This slide is intended for:
    - closing the presentation

    Args:
        prs_id: Presentation ID returned by create_presentation.

    Returns:
        Slide number / position of the newly added thank-you slide.

    Important:
        - Use this tool only for the final closing slide.
        - Do not use it for normal content slides.
        - After this slide, usually call save_presentation.
    """
    try:
        slide_number = slide_service.add_thank_you_slide(
            prs_id)

        return f"Added thank-you slide at position {slide_number}"

    except Exception as e:
        return f"Error adding thank-you slide: {str(e)}"