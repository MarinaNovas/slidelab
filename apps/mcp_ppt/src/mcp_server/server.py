from mcp.server.fastmcp import FastMCP
from typing import Optional

from typing_extensions import Any

from src.api.app import EXPORTS_DIR
from src.config import MEDIA_DIR, settings
from src.models.presentation import PresentationPlan
from src.models.slide import AgendaItemData, AgendaSlideData, ComparisonTableSlideData, ContentSlideData, \
    ImageSlideData, SectionSlideData, \
    TableSlideData, TitleSlideData
from src.services.presentation_creator import PresentationCreator
from src.services.presentation_store import PresentationStore
from src.services.slide_creator import SlideCreator
from src.services.yandex_art_image_generator import YandexArtImageGenerator
from src.services.plantuml_generator import PlantUMLImageGenerator

# Create an MCP server
mcp = FastMCP("PowerPoint Creator",  dependencies=["python-pptx","requests"], host="0.0.0.0", port=8001)

store = PresentationStore()
presentation_service = PresentationCreator(store)
slide_service = SlideCreator(store)
image_service = YandexArtImageGenerator(**settings.YANDEX_ART_CONFIG)
uml_service = PlantUMLImageGenerator(**settings.UML_CONFIG)
print(settings.YANDEX_ART_CONFIG)
print(settings.UML_CONFIG)

@mcp.tool()
def generate_presentation_from_json(payload: dict[str, Any]) -> dict:
    """
        Generate a complete PowerPoint presentation from one structured JSON payload.

        Use this tool when the presentation plan is already prepared.

        IMPORTANT:
        - Do NOT call create_presentation manually.
        - Do NOT call add_title_slide manually.
        - Do NOT call add_agenda_slide manually.
        - Do NOT call add_section_slide manually.
        - Do NOT call add_content_slide manually.
        - Do NOT call generate_yandex_art_image manually.
        - Do NOT call add_image_content_slide manually.
        - Do NOT call add_comparison_table_slide manually.
        - Do NOT call add_thank_you_slide manually.
        - Do NOT call save_presentation manually.

        This tool performs the full workflow internally:
        1. creates the presentation;
        2. adds title slide;
        3. adds agenda slide if agenda is provided;
        4. adds section slides;
        5. adds content slides;
        6. generates images for image_content slides;
        7. adds image_content slides;
        8. adds comparison table slides;
        9. adds final thank-you slide if add_thank_you=true;
        10. saves presentation if save=true;
        11. returns download_url.

        Payload structure:

        {
          "metadata": {
            "title": "Presentation title",
            "subtitle": "Optional subtitle or short description",
            "template_name": null
          },
          "agenda": [
            "Agenda item 1",
            "Agenda item 2",
            "Agenda item 3"
          ],
          "slides": [
            {
              "type": "section",
              "title": "Section title"
            },
            {
              "type": "content",
              "title": "Content slide title",
              "content": [
                "Short bullet point 1",
                "Short bullet point 2",
                "Short bullet point 3"
              ]
            },
            {
              "type": "image_content",
              "title": "Image slide title",
              "section_title": "Current section title",
              "subtitle": "Short subtitle",
              "content": [
                "Short bullet point 1",
                "Short bullet point 2",
                "Short bullet point 3"
              ],
              "image": {
                "prompt": "Detailed image generation prompt",
                "style": "business illustration",
                "width_ratio": 1,
                "height_ratio": 1,
                "seed": null
              }
            },
            {
              "type": "comparison_table",
              "title": "Comparison slide title",
              "sidebar_items": [
                "Short point for sidebar",
                "Short point for sidebar"
              ],
              "table_title": "Table title",
              "headers": ["Option", "Pros", "Cons"],
              "rows": [
                ["Option A", "Good for speed", "Less flexible"],
                ["Option B", "More flexible", "More complex"]
              ]
            }
          ],
          "add_thank_you": true,
          "save": true
        }

        Required fields:
        - metadata.title
        - slides
        - every slide.type
        - every slide.title

        Slide types:

        1. section
        Use for section divider slides.
        Required:
        - type = "section"
        - title

        Example:
        {
          "type": "section",
          "title": "Market Overview"
        }

        2. content
        Use for normal text slides with bullet points.
        Required:
        - type = "content"
        - title
        - content

        Rules:
        - content must be a list of short bullet points.
        - Use 3–6 bullet points.
        - Avoid long paragraphs.

        Example:
        {
          "type": "content",
          "title": "Key Challenges",
          "content": [
            "Manual slide preparation takes too much time",
            "Visual consistency is difficult to maintain",
            "Content often lacks structure"
          ]
        }

        3. image_content
        Use for slides with generated image and bullet points.
        Required:
        - type = "image_content"
        - title
        - content
        - image.prompt

        Optional:
        - section_title
        - subtitle
        - image.style
        - image.width_ratio
        - image.height_ratio
        - image.seed

        Rules:
        - Use this slide type for approximately 30–50% of main slides.
        - The image prompt must describe the visual scene clearly.
        - The image must support the slide content.
        - Do not pass image_id manually.
        - The tool generates the image internally.

        Example:
        {
          "type": "image_content",
          "title": "AI-assisted Presentation Workflow",
          "section_title": "Solution Architecture",
          "subtitle": "From user request to final .pptx",
          "content": [
            "Agent collects the requirements",
            "MCP receives a structured plan",
            "PowerPoint file is generated automatically"
          ],
          "image": {
            "prompt": "modern business workflow diagram, AI assistant creating presentation slides, clean corporate style",
            "style": "business illustration",
            "width_ratio": 1,
            "height_ratio": 1,
            "seed": null
          }
        }

        4. comparison_table
        Use only for comparisons, metrics, structured data, KPIs, risks, features or scenarios.
        Required:
        - type = "comparison_table"
        - title
        - headers
        - rows

        Optional:
        - sidebar_items
        - table_title

        Rules:
        - headers should contain 3–5 columns.
        - rows should contain 3–7 rows.
        - Every row must have the same number of cells as headers.
        - Keep cell text short.

        Example:
        {
          "type": "comparison_table",
          "title": "Approach Comparison",
          "sidebar_items": [
            "JSON mode is more stable",
            "Step-by-step mode is more flexible"
          ],
          "table_title": "Generation modes",
          "headers": ["Mode", "Best for", "Risk"],
          "rows": [
            ["JSON", "Stable generation", "Strict schema required"],
            ["Step-by-step", "Flexible editing", "Wrong call order"],
            ["Hybrid", "Production use", "More code"]
          ]
        }

        Global rules:
        - Start each major section with a section slide.
        - After section slide, add 1–3 main slides.
        - Use image_content for visual explanation.
        - Use content for ideas, conclusions and explanations.
        - Use comparison_table only when table format is really needed.
        - Keep all text short enough to fit on slides.
        - Do not include markdown in slide text.
        - Do not include HTML in slide text.
        - Do not include speaker notes in this payload unless schema supports them.
        - Do not invent unsupported slide types.
        - Supported slide types are only:
          section, content, image_content, comparison_table.

        Returns:
        {
          "status": "ok",
          "prs_id": "...",
          "download_url": "...",
          "slides_count": 10,
          "log": [...]
        }

        On error returns:
        {
          "status": "error",
          "message": "Error description",
          "log": [...]
        }
    """
    log: list[str] = []

    try:
        plan = PresentationPlan.model_validate(payload)

        prs_id = presentation_service.create(
            title=plan.metadata.title,
            template_name=plan.metadata.template_name,
        )
        log.append(f"Created presentation: {prs_id}")

        title_slide_number = slide_service.add_title_slide(
            prs_id,
            TitleSlideData(
                title=plan.metadata.title,
                subtitle=plan.metadata.subtitle,
            ),
        )
        log.append(f"Added title slide: {title_slide_number}")

        if plan.agenda:
            agenda_items = [
                AgendaItemData(
                    number=str(i + 1),
                    title=item,
                )
                for i, item in enumerate(plan.agenda[:6])
            ]

            agenda_slide_number = slide_service.add_agenda_slide(
                prs_id,
                AgendaSlideData(
                    title="AGENDA",
                    items=agenda_items,
                ),
            )
            log.append(f"Added agenda slide: {agenda_slide_number}")

        for slide in plan.slides:
            if slide.type == "section":
                slide_number = slide_service.add_section_slide(
                    prs_id,
                    SectionSlideData(
                        section_title=slide.title,
                    ),
                )
                log.append(f"Added section slide: {slide_number}")

            elif slide.type == "content":
                slide_number = slide_service.add_content_slide(
                    prs_id,
                    ContentSlideData(
                        title=slide.title,
                        content=slide.content,
                    ),
                )
                log.append(f"Added content slide: {slide_number}")

            elif slide.type == "image_content":
                image_result = image_service.generate(
                    prompt=slide.image.prompt,
                    style=slide.image.style,
                    width_ratio=slide.image.width_ratio,
                    height_ratio=slide.image.height_ratio,
                    seed=slide.image.seed,
                )

                if image_result.get("status") != "ok":
                    raise ValueError(
                        f"Image generation failed for slide '{slide.title}': {image_result}"
                    )

                image_id = image_result["image_id"]
                image_path = MEDIA_DIR / f"{image_id}.jpeg"

                slide_number = slide_service.add_image_content_slide(
                    prs_id,
                    ImageSlideData(
                        title=slide.title,
                        subtitle=slide.subtitle or slide.section_title or "",
                        content=slide.content,
                        image_path=str(image_path),
                    ),
                )

                log.append(
                    f"Added image content slide: {slide_number}, image_id={image_id}"
                )

            elif slide.type == "comparison_table":
                slide_number = slide_service.add_comparison_table_slide(
                    prs_id,
                    ComparisonTableSlideData(
                        title=slide.title,
                        sidebar_items=slide.sidebar_items,
                        table_title=slide.table_title or "",
                        headers=slide.headers,
                        rows=slide.rows,
                    ),
                )
                log.append(f"Added comparison table slide: {slide_number}")

            else:
                raise ValueError(f"Unsupported slide type: {slide.type}")

        if plan.add_thank_you:
            slide_number = slide_service.add_thank_you_slide(prs_id)
            log.append(f"Added thank-you slide: {slide_number}")

        download_url = None

        if plan.save:
            current_prs = store.get(prs_id)

            if current_prs is None:
                raise ValueError(f"Presentation '{prs_id}' not found")

            file_name = f"{prs_id}.pptx"
            path = EXPORTS_DIR / file_name
            current_prs.prs.save(path)

            download_url = f"{settings.PUBLIC_BASE_URL}/exports/{file_name}"
            log.append(f"Saved presentation: {download_url}")

        return {
            "status": "ok",
            "prs_id": prs_id,
            "download_url": download_url,
            "slides_count": len(store.get(prs_id).prs.slides),
            "log": log,
        }

    except Exception as e:
        return {
            "status": "error",
            "message": str(e),
            "log": log,
        }

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
    return f"Presentation saved. Download URL: {settings.PUBLIC_BASE_URL}/exports/{file_name}"

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
            image_path = MEDIA_DIR / f"{image_id}.png"

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
def generate_uml_diagram(
    plantuml_code: str,
) -> dict:
    """
    Generate uml diagram as image and save it locally.

    Use this tool when a presentation needs a generated uml diagram such as:
    sequence, class, use case, activity, component, state, deplyment, object or thinking.

    Args:
        plantuml_code: The PlantUML definition code.
                       Example:
                       @startuml
                       Alice -> Bob : Hello
                       @enduml.

    Returns:
        Dict with:
        - status
        - image_id
        - image_path
        - image_url

    Workflow:
        1. Call create_presentation
        2. Call generate_uml_image
        3. Pass returned image_id to add_image_content_slide
        4. Call save_presentation
    """
    try:
        return uml_service.generate(
        plantuml_code=plantuml_code,
        )

    except Exception as e:
        return {
            "status": "error",
            "message": str(e),
        }

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