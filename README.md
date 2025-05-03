# PPTX MCP Server

A FastMCP-powered server for programmatically creating, editing, and rendering PowerPoint (PPTX) presentations. Supports slide creation, text and shape insertion, image embedding, and slide rendering to PNG (with LibreOffice).

## Features

- **Create/Clear Presentations:** Start new or reset existing PPTX files.
- **Add Slides:** Insert slides with customizable layouts.
- **Text & Content:** Add titles, content, and custom textboxes to slides.
- **Shapes:** Insert a wide variety of PowerPoint shapes (including flowchart elements).
- **Images:** Embed images into slides.
- **Slide Description:** Get a textual summary of slide contents for verification.
- **Slide Rendering:** Render slides as PNG images (requires LibreOffice).
- **Download PPTX:** Download the generated presentation file.

## Requirements

- Python 3.12+
- [python-pptx](https://python-pptx.readthedocs.io/)
- [Pillow](https://python-pillow.org/)
- [FastMCP](https://github.com/ContextualAI/fastmcp)
- **LibreOffice** (for slide image rendering; must be installed separately and available in your system PATH)

## Installation

1. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   *(Or see `pyproject.toml` for dependencies)*

2. **Install LibreOffice** (for image rendering):
   - Linux: `sudo pacman -S libreoffice-fresh` or `sudo apt install libreoffice`
   - macOS: `brew install --cask libreoffice`
   - Windows: [Download from libreoffice.org](https://www.libreoffice.org/download/download/)

## Usage

Start the server:
```bash
python server.py
```
or (for development with FastMCP):
```bash
fastmcp dev server.py
```

## API Overview

The server exposes tools and resources via FastMCP, including:

- `create_or_clear_presentation(filename)`
- `add_slide(filename, layout_index)`
- `add_title_and_content(filename, slide_index, title, content)`
- `add_textbox(filename, slide_index, text, left_inches, top_inches, width_inches, height_inches, font_size_pt, bold)`
- `add_shape(filename, slide_index, shape_type_name, left_inches, top_inches, width_inches, height_inches, text)`
- `add_picture(filename, slide_index, image, left_inches, top_inches, width_inches, height_inches)`
- `get_slide_content_description(filename, slide_index)`
- `get_slide_image(filename, slide_index)` *(requires LibreOffice)*
- `get_pptx_file(filename)`

See the code for full parameter details and available shape types.

## Presentations & Templates

- Presentations are saved in the `presentations/` directory.
- You can add your own templates in `presentations/templates/`.

## License

See [LICENSE](LICENSE) for details.


