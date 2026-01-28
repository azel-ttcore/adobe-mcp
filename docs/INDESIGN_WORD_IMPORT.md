# InDesign Word Document Import

This guide explains how to use the InDesign MCP server to automatically layout documents from Word files with style mapping.

## Overview

The InDesign MCP server provides tools to:

1. **Parse Word documents** - Extract text, styles, formatting, and images from DOCX files
2. **Map styles** - Map Word paragraph styles to InDesign paragraph styles
3. **Auto-layout content** - Automatically create text frames, flow content, and add pages
4. **Place images** - Insert images from the Word document into the InDesign layout

## Prerequisites

- InDesign 2024 or later with UXP support
- The InDesign UXP plugin loaded and connected to the proxy server
- Python environment with the adobe-mcp package installed

## Quick Start

### 1. Start the Services

```bash
# Start the proxy server
adobe-proxy

# In InDesign, load the UXP plugin and click "Connect"
```

### 2. Basic Word Import

Use the `import_word_to_indesign` tool for a complete automated workflow:

```python
# Example prompt to an LLM using the MCP server:
"Import the Word document at C:/Documents/report.docx into InDesign 
using the template at C:/Templates/report-template.indd"
```

The tool will:
1. Parse the Word document
2. Open the InDesign template
3. Create/map paragraph styles
4. Flow content into text frames
5. Place images

### 3. Step-by-Step Workflow

For more control, use individual tools:

#### Parse the Word Document First

```
parse_word_document(docx_path="C:/Documents/report.docx")
```

Returns:
- List of styles found in the document
- Content blocks with paragraph text and style info
- Extracted images with temporary file paths
- Suggested style mapping

#### Preview Style Mapping

```
get_style_mapping_suggestions(docx_path="C:/Documents/report.docx")
```

Returns a mapping like:
```json
{
  "Heading 1": "Heading 1",
  "Normal": "Body Text",
  "Title": "Title",
  "Quote": "Block Quote"
}
```

#### Customize and Import

```
import_word_to_indesign(
    docx_path="C:/Documents/report.docx",
    template_path="C:/Templates/my-template.indd",
    style_mapping={
        "Heading 1": "Chapter Title",
        "Normal": "Body Copy",
        "Quote": "Pull Quote"
    },
    auto_flow=True,
    place_images=True
)
```

## Available Tools

### Document Management

| Tool | Description |
|------|-------------|
| `create_document` | Create a new InDesign document |
| `open_document` | Open an existing .indd file |
| `save_document` | Save the active document |
| `export_pdf` | Export to PDF |

### Paragraph Styles

| Tool | Description |
|------|-------------|
| `get_paragraph_styles` | List all paragraph styles |
| `create_paragraph_style` | Create a new paragraph style |
| `apply_paragraph_style` | Apply a style to text |

### Text Frames

| Tool | Description |
|------|-------------|
| `create_text_frame` | Create a text frame on a page |
| `get_text_frames` | Get info about text frames |
| `link_text_frames` | Link frames for text flow |
| `insert_text` | Insert text into a frame |

### Images

| Tool | Description |
|------|-------------|
| `place_image` | Place an image on a page |
| `get_images` | List all placed images |

### Pages

| Tool | Description |
|------|-------------|
| `add_page` | Add a new page |
| `get_pages` | Get page information |
| `get_master_pages` | List master pages |

### Word Import

| Tool | Description |
|------|-------------|
| `parse_word_document` | Parse DOCX and extract content |
| `get_style_mapping_suggestions` | Get suggested style mappings |
| `import_word_to_indesign` | Full automated import workflow |

## Style Mapping

### Default Mappings

The system provides sensible defaults for common Word styles:

| Word Style | InDesign Style |
|------------|----------------|
| Normal | Body Text |
| Heading 1 | Heading 1 |
| Heading 2 | Heading 2 |
| Heading 3 | Heading 3 |
| Title | Title |
| Subtitle | Subtitle |
| Quote | Block Quote |
| List Paragraph | List Item |
| Caption | Caption |

### Custom Mappings

Provide a custom mapping dictionary to override defaults:

```python
style_mapping = {
    "Normal": "My Body Style",
    "Heading 1": "Chapter Header",
    "Custom Word Style": "Custom InDesign Style"
}
```

If a mapped InDesign style doesn't exist, it will be created automatically.

## Working with Templates

For best results, prepare an InDesign template with:

1. **Paragraph styles** matching your expected Word styles
2. **Master pages** with text frame placeholders
3. **Document settings** (page size, margins, columns)

The import process will use existing styles when available and create new ones as needed.

## Image Handling

Images embedded in Word documents are:

1. Extracted to a temporary directory
2. Placed in the InDesign document
3. Positioned automatically (or at specified locations)

For precise image placement, use the `place_image` tool directly after parsing.

## Troubleshooting

### "Style not found" warnings
- The system creates missing styles automatically
- Check that your template has the expected styles defined

### Text overflow
- Enable `auto_flow=True` to automatically create pages
- Or manually create linked text frames first

### Images not appearing
- Verify image paths are accessible
- Check that `place_images=True` is set

### Connection issues
- Ensure the proxy server is running on port 3001
- Verify the UXP plugin shows "Connected" status

## Example Workflow

```
1. "Open the template at C:/Templates/newsletter.indd"

2. "Parse the Word document at C:/Content/article.docx and show me the styles"

3. "Import the content using this style mapping:
    - 'Article Title' -> 'Newsletter Headline'
    - 'Body' -> 'Article Body'
    - 'Byline' -> 'Author Credit'"

4. "Place the first image at position (72, 200) on page 1 with width 300"

5. "Export as PDF to C:/Output/newsletter.pdf"
```

## Limitations

- Complex Word formatting (columns, text boxes) may not transfer perfectly
- Table support is basic (converted to tab-separated text)
- Some advanced InDesign features require manual adjustment
- Large documents may take time to process

## Future Enhancements

Planned improvements include:
- Full table support with InDesign table creation
- Character style mapping
- Anchored image positioning
- Multi-column layout detection
- IDML/ICML import support
