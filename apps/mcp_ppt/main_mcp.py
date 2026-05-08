from pptx import Presentation

from src.mcp_server.server import mcp

def main():
   prs = Presentation("templates/template_axenix.pptx")

   for master_index, master in enumerate(prs.slide_masters):
      print(f"\nMASTER {master_index}")

      for layout_index, layout in enumerate(master.slide_layouts):
         print(f"  {layout_index}: {layout.name}")

         for ph in layout.placeholders:
            print(
               "     ",
               ph.placeholder_format.idx,
               ph.placeholder_format.type,
               ph.name,
            )

main()

if __name__ == "__main__":
   mcp.run(transport = "streamable-http")
