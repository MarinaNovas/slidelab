from mcp.server.fastmcp import FastMCP

from src.services.presentation_creator import PresentationCreator

# Import pptx after declaring dependencies
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
except ImportError:
    # Print detailed error to stderr for debugging
    print("Failed to import pptx modules. Make sure python-pptx is installed.")
    raise

# Store presentations in memory during a session
presentations = {}

def register_tools(mcp: FastMCP) -> None:
    @mcp.tool()
    def create_presentation(title: str) -> str:
        """
        Create a new empty PowerPoint presentation with a title.

        Args:
            title: The title/name for the presentation

        Returns:
            Presentation ID to use in subsequent operations
        """
        prs = Presentation()
        prs_id = (title)

        # Ensure uniqueness (though UUID makes collision extremely unlikely)
        while prs_id in presentations:
            prs_id =  PresentationCreator.generate_id(title)

        presentations[prs_id] = prs
        return f"Created presentation: {prs_id}"

