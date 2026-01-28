# MIT License
#
# Copyright (c) 2025 Mike Chambers
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

from mcp.server.fastmcp import FastMCP
from ..shared import init, sendCommand, createCommand, socket_client
from .docx_parser import parse_docx, get_style_mapping_template, DocxParser
import sys
import os
import json

#logger.log(f"Python path: {sys.executable}")
#logger.log(f"PYTHONPATH: {os.environ.get('PYTHONPATH')}")
#logger.log(f"Current working directory: {os.getcwd()}")
#logger.log(f"Sys.path: {sys.path}")


# Create an MCP server
mcp_name = "Adobe InDesign MCP Server"
mcp = FastMCP(mcp_name, log_level="ERROR")
print(f"{mcp_name} running on stdio", file=sys.stderr)

APPLICATION = "indesign"
PROXY_URL = 'http://localhost:3001'
PROXY_TIMEOUT = 20

socket_client.configure(
    app=APPLICATION, 
    url=PROXY_URL,
    timeout=PROXY_TIMEOUT
)

init(APPLICATION, socket_client)

@mcp.tool()
def create_document(
    width:int, height:int, pages:int = 0,
    pages_facing:bool = False,
    columns:dict = {"count":1, "gutter":12},
    margins:dict = {"top":36, "bottom":36, "left":36, "right":36}):

    command = createCommand("createDocument", {
        "intent":"WEB_INTENT",
        "pageWidth":width,
        "pageHeight":height,
        "margins":margins,
        "columns":columns,
        "pagesPerDocument":pages,
        "pagesFacing":pages_facing
    })

    return sendCommand(command)

@mcp.resource("config://get_instructions")
def get_instructions() -> str:
    """Read this first! Returns information and instructions on how to use Photoshop and this API"""

    return f"""
    You are an InDesign and design expert who is creative and loves to help other people learn to use InDesign and create.

    Rules to follow:

    1. Think deeply about how to solve the task
    2. Always check your work
    3. Read the info for the API calls to make sure you understand the requirements and arguments

    ## Word Document Import Workflow

    To import a Word document into InDesign with automatic layout:

    1. Use `parse_word_document` to analyze the DOCX file and see its styles
    2. Use `get_style_mapping_suggestions` to preview how Word styles map to InDesign
    3. Optionally open an InDesign template with `open_document`
    4. Use `import_word_to_indesign` for the full automated workflow, or:
       - Create paragraph styles with `create_paragraph_style`
       - Create text frames with `create_text_frame`
       - Insert content with `insert_text`
       - Place images with `place_image`

    The `import_word_to_indesign` tool handles the complete workflow automatically:
    - Parses the Word document
    - Opens a template (if provided)
    - Maps Word styles to InDesign paragraph styles
    - Creates text frames and flows content
    - Places images from the document
    - Adds pages as needed for overflow text
    """


# =============================================================================
# PARAGRAPH STYLE TOOLS
# =============================================================================

@mcp.tool()
def get_paragraph_styles():
    """
    Returns a list of all paragraph styles defined in the active InDesign document.
    
    Returns:
        list: Array of style objects with name, id, and formatting properties.
    """
    command = createCommand("getParagraphStyles", {})
    return sendCommand(command)


@mcp.tool()
def create_paragraph_style(
    style_name: str,
    based_on: str = None,
    font_family: str = None,
    font_size: float = None,
    leading: float = None,
    alignment: str = None,
    space_before: float = None,
    space_after: float = None,
    first_line_indent: float = None,
    left_indent: float = None,
    right_indent: float = None,
    color: dict = None
):
    """
    Creates a new paragraph style in the active InDesign document.
    
    Args:
        style_name (str): Name for the new paragraph style.
        based_on (str): Name of existing style to base this style on.
        font_family (str): Font family name (e.g., "Minion Pro").
        font_size (float): Font size in points.
        leading (float): Line spacing in points (use 0 for auto).
        alignment (str): Text alignment - LEFT, CENTER, RIGHT, JUSTIFY, FULLY_JUSTIFIED.
        space_before (float): Space before paragraph in points.
        space_after (float): Space after paragraph in points.
        first_line_indent (float): First line indent in points.
        left_indent (float): Left indent in points.
        right_indent (float): Right indent in points.
        color (dict): Text color as RGB dict {"red": 0-255, "green": 0-255, "blue": 0-255}.
    """
    options = {"styleName": style_name}
    
    if based_on:
        options["basedOn"] = based_on
    if font_family:
        options["fontFamily"] = font_family
    if font_size is not None:
        options["fontSize"] = font_size
    if leading is not None:
        options["leading"] = leading
    if alignment:
        options["alignment"] = alignment
    if space_before is not None:
        options["spaceBefore"] = space_before
    if space_after is not None:
        options["spaceAfter"] = space_after
    if first_line_indent is not None:
        options["firstLineIndent"] = first_line_indent
    if left_indent is not None:
        options["leftIndent"] = left_indent
    if right_indent is not None:
        options["rightIndent"] = right_indent
    if color:
        options["color"] = color
    
    command = createCommand("createParagraphStyle", options)
    return sendCommand(command)


@mcp.tool()
def apply_paragraph_style(style_name: str, story_index: int = 0, paragraph_index: int = None):
    """
    Applies a paragraph style to text in the document.
    
    Args:
        style_name (str): Name of the paragraph style to apply.
        story_index (int): Index of the story (text flow) to target. Default 0.
        paragraph_index (int): Index of specific paragraph, or None for all paragraphs.
    """
    options = {"styleName": style_name, "storyIndex": story_index}
    if paragraph_index is not None:
        options["paragraphIndex"] = paragraph_index
    
    command = createCommand("applyParagraphStyle", options)
    return sendCommand(command)


# =============================================================================
# TEXT FRAME TOOLS
# =============================================================================

@mcp.tool()
def create_text_frame(
    page_index: int,
    x: float,
    y: float,
    width: float,
    height: float,
    content: str = None
):
    """
    Creates a new text frame on the specified page.
    
    Args:
        page_index (int): Zero-based index of the page to add the frame to.
        x (float): X position of the frame's top-left corner in points.
        y (float): Y position of the frame's top-left corner in points.
        width (float): Width of the frame in points.
        height (float): Height of the frame in points.
        content (str): Optional initial text content for the frame.
    
    Returns:
        dict: Information about the created frame including its ID.
    """
    options = {
        "pageIndex": page_index,
        "x": x,
        "y": y,
        "width": width,
        "height": height
    }
    if content:
        options["content"] = content
    
    command = createCommand("createTextFrame", options)
    return sendCommand(command)


@mcp.tool()
def get_text_frames(page_index: int = None):
    """
    Returns information about text frames in the document.
    
    Args:
        page_index (int): Optional page index to filter frames. None returns all frames.
    
    Returns:
        list: Array of text frame objects with id, bounds, and content info.
    """
    options = {}
    if page_index is not None:
        options["pageIndex"] = page_index
    
    command = createCommand("getTextFrames", options)
    return sendCommand(command)


@mcp.tool()
def link_text_frames(source_frame_id: str, target_frame_id: str):
    """
    Links two text frames so text flows from source to target.
    
    Args:
        source_frame_id (str): ID of the source text frame.
        target_frame_id (str): ID of the target text frame to flow into.
    """
    command = createCommand("linkTextFrames", {
        "sourceFrameId": source_frame_id,
        "targetFrameId": target_frame_id
    })
    return sendCommand(command)


@mcp.tool()
def insert_text(
    frame_id: str,
    text: str,
    position: str = "END",
    style_name: str = None
):
    """
    Inserts text into a text frame.
    
    Args:
        frame_id (str): ID of the text frame to insert into.
        text (str): The text content to insert.
        position (str): Where to insert - START, END, or REPLACE.
        style_name (str): Optional paragraph style to apply to inserted text.
    """
    options = {
        "frameId": frame_id,
        "text": text,
        "position": position
    }
    if style_name:
        options["styleName"] = style_name
    
    command = createCommand("insertText", options)
    return sendCommand(command)


# =============================================================================
# IMAGE PLACEMENT TOOLS
# =============================================================================

@mcp.tool()
def place_image(
    image_path: str,
    page_index: int,
    x: float,
    y: float,
    width: float = None,
    height: float = None,
    fit_option: str = "PROPORTIONALLY"
):
    """
    Places an image on a page in the InDesign document.
    
    Args:
        image_path (str): Absolute file path to the image file.
        page_index (int): Zero-based index of the page to place the image on.
        x (float): X position of the image frame's top-left corner in points.
        y (float): Y position of the image frame's top-left corner in points.
        width (float): Optional width of the image frame in points.
        height (float): Optional height of the image frame in points.
        fit_option (str): How to fit the image - PROPORTIONALLY, FILL_PROPORTIONALLY, 
                         FIT_CONTENT, CENTER_CONTENT, FRAME_TO_CONTENT.
    
    Returns:
        dict: Information about the placed image including frame ID.
    """
    options = {
        "imagePath": image_path,
        "pageIndex": page_index,
        "x": x,
        "y": y,
        "fitOption": fit_option
    }
    if width is not None:
        options["width"] = width
    if height is not None:
        options["height"] = height
    
    command = createCommand("placeImage", options)
    return sendCommand(command)


@mcp.tool()
def get_images():
    """
    Returns information about all placed images in the document.
    
    Returns:
        list: Array of image objects with id, path, bounds, and page info.
    """
    command = createCommand("getImages", {})
    return sendCommand(command)


# =============================================================================
# PAGE MANAGEMENT TOOLS
# =============================================================================

@mcp.tool()
def add_page(at_index: int = None, master_page: str = None):
    """
    Adds a new page to the document.
    
    Args:
        at_index (int): Position to insert the page. None adds at end.
        master_page (str): Name of master page to apply. None uses default.
    
    Returns:
        dict: Information about the new page including its index.
    """
    options = {}
    if at_index is not None:
        options["atIndex"] = at_index
    if master_page:
        options["masterPage"] = master_page
    
    command = createCommand("addPage", options)
    return sendCommand(command)


@mcp.tool()
def get_pages():
    """
    Returns information about all pages in the document.
    
    Returns:
        list: Array of page objects with index, size, margins, and master page info.
    """
    command = createCommand("getPages", {})
    return sendCommand(command)


@mcp.tool()
def get_master_pages():
    """
    Returns information about all master pages in the document.
    
    Returns:
        list: Array of master page objects with name and applied items.
    """
    command = createCommand("getMasterPages", {})
    return sendCommand(command)


# =============================================================================
# DOCUMENT TOOLS
# =============================================================================

@mcp.tool()
def open_document(file_path: str):
    """
    Opens an existing InDesign document.
    
    Args:
        file_path (str): Absolute path to the .indd file.
    
    Returns:
        dict: Information about the opened document.
    """
    command = createCommand("openDocument", {"filePath": file_path})
    return sendCommand(command)


@mcp.tool()
def save_document(file_path: str = None):
    """
    Saves the active document.
    
    Args:
        file_path (str): Optional path to save to. None saves to current location.
    """
    options = {}
    if file_path:
        options["filePath"] = file_path
    
    command = createCommand("saveDocument", options)
    return sendCommand(command)


@mcp.tool()
def export_pdf(file_path: str, preset: str = None):
    """
    Exports the document as a PDF.
    
    Args:
        file_path (str): Absolute path for the output PDF file.
        preset (str): Optional PDF preset name (e.g., "High Quality Print").
    """
    options = {"filePath": file_path}
    if preset:
        options["preset"] = preset
    
    command = createCommand("exportPdf", options)
    return sendCommand(command)


# =============================================================================
# WORD DOCUMENT IMPORT TOOLS
# =============================================================================

@mcp.tool()
def parse_word_document(docx_path: str):
    """
    Parses a Word document and returns its structured content.
    
    This extracts paragraphs, styles, formatting, and images from the DOCX file
    for use in subsequent layout operations.
    
    Args:
        docx_path (str): Absolute path to the .docx file.
    
    Returns:
        dict: Structured content including:
            - styles: List of Word styles found
            - content: List of paragraphs with style and formatting info
            - images: List of extracted images with paths
            - style_mapping: Suggested Word-to-InDesign style mapping
    """
    if not os.path.exists(docx_path):
        return {"status": "error", "message": f"File not found: {docx_path}"}
    
    try:
        parsed = parse_docx(docx_path)
        style_mapping = get_style_mapping_template(parsed)
        parsed["style_mapping"] = style_mapping
        
        return {
            "status": "success",
            "source_file": docx_path,
            "styles": parsed["styles"],
            "content_blocks": len(parsed["content"]),
            "images_found": len(parsed["images"]),
            "style_mapping": style_mapping,
            "content": parsed["content"],
            "images": parsed["images"],
            "image_temp_dir": parsed["image_temp_dir"]
        }
    except Exception as e:
        return {"status": "error", "message": str(e)}


@mcp.tool()
def import_word_to_indesign(
    docx_path: str,
    template_path: str = None,
    style_mapping: dict = None,
    auto_flow: bool = True,
    place_images: bool = True,
    start_page: int = 0
):
    """
    High-level tool to import a Word document into InDesign with automatic layout.
    
    This orchestrates the full workflow:
    1. Parse the Word document
    2. Open or create an InDesign document (optionally from template)
    3. Create/map paragraph styles
    4. Create text frames and flow content
    5. Place images at appropriate locations
    
    Args:
        docx_path (str): Absolute path to the .docx file.
        template_path (str): Optional path to an .indd template file.
        style_mapping (dict): Optional dict mapping Word style names to InDesign style names.
                             If not provided, uses automatic mapping.
        auto_flow (bool): If True, automatically creates pages and flows text.
        place_images (bool): If True, places images from the Word document.
        start_page (int): Page index to start placing content.
    
    Returns:
        dict: Status and summary of the import operation.
    """
    results = {
        "status": "in_progress",
        "steps_completed": [],
        "errors": []
    }
    
    # Step 1: Parse Word document
    try:
        parsed = parse_docx(docx_path)
        results["steps_completed"].append("parsed_word_document")
        results["content_blocks"] = len(parsed["content"])
        results["images_found"] = len(parsed["images"])
    except Exception as e:
        results["status"] = "error"
        results["errors"].append(f"Failed to parse Word document: {e}")
        return results
    
    # Step 2: Open template or use active document
    if template_path:
        try:
            open_result = sendCommand(createCommand("openDocument", {"filePath": template_path}))
            if open_result.get("status") == "FAILURE":
                results["errors"].append(f"Failed to open template: {open_result.get('message')}")
            else:
                results["steps_completed"].append("opened_template")
        except Exception as e:
            results["errors"].append(f"Failed to open template: {e}")
    
    # Step 3: Build style mapping
    if style_mapping is None:
        style_mapping = get_style_mapping_template(parsed)
    results["style_mapping"] = style_mapping
    
    # Step 4: Send content to InDesign for layout
    content_payload = {
        "content": parsed["content"],
        "styleMapping": style_mapping,
        "autoFlow": auto_flow,
        "startPage": start_page
    }
    
    try:
        layout_result = sendCommand(createCommand("layoutWordContent", content_payload))
        if layout_result.get("status") == "SUCCESS":
            results["steps_completed"].append("laid_out_content")
            results["pages_created"] = layout_result.get("response", {}).get("pagesCreated", 0)
            results["paragraphs_placed"] = layout_result.get("response", {}).get("paragraphsPlaced", 0)
        else:
            results["errors"].append(f"Layout failed: {layout_result.get('message')}")
    except Exception as e:
        results["errors"].append(f"Layout error: {e}")
    
    # Step 5: Place images
    if place_images and parsed["images"]:
        images_placed = 0
        for img in parsed["images"]:
            try:
                img_result = sendCommand(createCommand("placeImageFromWord", {
                    "imagePath": img["extracted_path"],
                    "imageId": img["id"],
                    "autoPosition": True
                }))
                if img_result.get("status") == "SUCCESS":
                    images_placed += 1
            except Exception as e:
                results["errors"].append(f"Failed to place image {img['id']}: {e}")
        
        results["images_placed"] = images_placed
        if images_placed > 0:
            results["steps_completed"].append("placed_images")
    
    # Final status
    if not results["errors"]:
        results["status"] = "success"
    elif results["steps_completed"]:
        results["status"] = "partial_success"
    else:
        results["status"] = "error"
    
    return results


@mcp.tool()
def get_style_mapping_suggestions(docx_path: str):
    """
    Analyzes a Word document and suggests style mappings to InDesign.
    
    Use this to preview and customize the style mapping before import.
    
    Args:
        docx_path (str): Absolute path to the .docx file.
    
    Returns:
        dict: Suggested mappings and style details from the Word document.
    """
    if not os.path.exists(docx_path):
        return {"status": "error", "message": f"File not found: {docx_path}"}
    
    try:
        parsed = parse_docx(docx_path)
        mapping = get_style_mapping_template(parsed)
        
        # Get styles actually used in content
        styles_used = {}
        for block in parsed["content"]:
            style = block.get("style")
            if style:
                if style not in styles_used:
                    styles_used[style] = {"count": 0, "suggested_indesign_style": mapping.get(style, style)}
                styles_used[style]["count"] += 1
        
        return {
            "status": "success",
            "styles_in_document": parsed["styles"],
            "styles_used_in_content": styles_used,
            "suggested_mapping": mapping
        }
    except Exception as e:
        return {"status": "error", "message": str(e)}


def main():
    """Run the InDesign MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()