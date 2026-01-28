# MIT License
#
# Copyright (c) 2025 Adobe MCP Contributors
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

"""
DOCX Parser for InDesign MCP Server.

Extracts structured content from Word documents including:
- Paragraphs with style names
- Inline formatting (bold, italic, underline)
- Images (extracted to temp files)
- Tables (basic support)
"""

import os
import zipfile
import xml.etree.ElementTree as ET
from typing import List, Dict, Any, Optional
import base64
import tempfile
import re


# Word XML namespaces
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}


class DocxParser:
    """Parse DOCX files and extract structured content for InDesign."""
    
    def __init__(self, docx_path: str):
        self.docx_path = docx_path
        self.temp_dir = tempfile.mkdtemp(prefix="docx_images_")
        self.relationships = {}
        self.styles = {}
        self.content_blocks = []
        self.extracted_images = []
        
    def parse(self) -> Dict[str, Any]:
        """
        Parse the DOCX file and return structured content.
        
        Returns:
            dict: {
                "styles": [{"name": str, "basedOn": str, "formatting": dict}, ...],
                "content": [{"type": str, "style": str, "runs": [...], ...}, ...],
                "images": [{"id": str, "path": str, "width": int, "height": int}, ...]
            }
        """
        if not os.path.exists(self.docx_path):
            raise FileNotFoundError(f"DOCX file not found: {self.docx_path}")
        
        with zipfile.ZipFile(self.docx_path, 'r') as docx:
            # Parse relationships first (for images)
            self._parse_relationships(docx)
            
            # Parse styles
            self._parse_styles(docx)
            
            # Parse main document content
            self._parse_document(docx)
            
            # Extract images
            self._extract_images(docx)
        
        return {
            "source_file": self.docx_path,
            "styles": list(self.styles.values()),
            "content": self.content_blocks,
            "images": self.extracted_images,
            "image_temp_dir": self.temp_dir
        }
    
    def _parse_relationships(self, docx: zipfile.ZipFile):
        """Parse document relationships to map image IDs to file paths."""
        try:
            rels_xml = docx.read('word/_rels/document.xml.rels')
            root = ET.fromstring(rels_xml)
            
            for rel in root.findall('.//{%s}Relationship' % NAMESPACES['rel']):
                rel_id = rel.get('Id')
                target = rel.get('Target')
                rel_type = rel.get('Type', '')
                
                self.relationships[rel_id] = {
                    'target': target,
                    'type': rel_type
                }
        except KeyError:
            pass  # No relationships file
    
    def _parse_styles(self, docx: zipfile.ZipFile):
        """Parse document styles."""
        try:
            styles_xml = docx.read('word/styles.xml')
            root = ET.fromstring(styles_xml)
            
            for style in root.findall('.//w:style', NAMESPACES):
                style_id = style.get('{%s}styleId' % NAMESPACES['w'])
                style_type = style.get('{%s}type' % NAMESPACES['w'])
                
                # Get style name
                name_elem = style.find('w:name', NAMESPACES)
                style_name = name_elem.get('{%s}val' % NAMESPACES['w']) if name_elem is not None else style_id
                
                # Get basedOn
                based_on_elem = style.find('w:basedOn', NAMESPACES)
                based_on = based_on_elem.get('{%s}val' % NAMESPACES['w']) if based_on_elem is not None else None
                
                # Extract formatting properties
                formatting = self._extract_style_formatting(style)
                
                self.styles[style_id] = {
                    'id': style_id,
                    'name': style_name,
                    'type': style_type,
                    'basedOn': based_on,
                    'formatting': formatting
                }
        except KeyError:
            pass  # No styles file
    
    def _extract_style_formatting(self, style_elem) -> Dict[str, Any]:
        """Extract formatting properties from a style element."""
        formatting = {}
        
        # Paragraph properties
        pPr = style_elem.find('w:pPr', NAMESPACES)
        if pPr is not None:
            # Alignment
            jc = pPr.find('w:jc', NAMESPACES)
            if jc is not None:
                formatting['alignment'] = jc.get('{%s}val' % NAMESPACES['w'])
            
            # Spacing
            spacing = pPr.find('w:spacing', NAMESPACES)
            if spacing is not None:
                if spacing.get('{%s}before' % NAMESPACES['w']):
                    formatting['spaceBefore'] = int(spacing.get('{%s}before' % NAMESPACES['w'])) / 20  # twips to points
                if spacing.get('{%s}after' % NAMESPACES['w']):
                    formatting['spaceAfter'] = int(spacing.get('{%s}after' % NAMESPACES['w'])) / 20
                if spacing.get('{%s}line' % NAMESPACES['w']):
                    formatting['lineSpacing'] = int(spacing.get('{%s}line' % NAMESPACES['w'])) / 240  # to lines
            
            # Indentation
            ind = pPr.find('w:ind', NAMESPACES)
            if ind is not None:
                if ind.get('{%s}left' % NAMESPACES['w']):
                    formatting['leftIndent'] = int(ind.get('{%s}left' % NAMESPACES['w'])) / 20
                if ind.get('{%s}right' % NAMESPACES['w']):
                    formatting['rightIndent'] = int(ind.get('{%s}right' % NAMESPACES['w'])) / 20
                if ind.get('{%s}firstLine' % NAMESPACES['w']):
                    formatting['firstLineIndent'] = int(ind.get('{%s}firstLine' % NAMESPACES['w'])) / 20
        
        # Run (character) properties
        rPr = style_elem.find('w:rPr', NAMESPACES)
        if rPr is not None:
            formatting.update(self._extract_run_formatting(rPr))
        
        return formatting
    
    def _extract_run_formatting(self, rPr) -> Dict[str, Any]:
        """Extract character formatting from run properties."""
        formatting = {}
        
        # Font
        rFonts = rPr.find('w:rFonts', NAMESPACES)
        if rFonts is not None:
            formatting['fontFamily'] = rFonts.get('{%s}ascii' % NAMESPACES['w']) or rFonts.get('{%s}hAnsi' % NAMESPACES['w'])
        
        # Font size (half-points to points)
        sz = rPr.find('w:sz', NAMESPACES)
        if sz is not None:
            formatting['fontSize'] = int(sz.get('{%s}val' % NAMESPACES['w'])) / 2
        
        # Bold
        b = rPr.find('w:b', NAMESPACES)
        if b is not None:
            val = b.get('{%s}val' % NAMESPACES['w'])
            formatting['bold'] = val != '0' and val != 'false'
        
        # Italic
        i = rPr.find('w:i', NAMESPACES)
        if i is not None:
            val = i.get('{%s}val' % NAMESPACES['w'])
            formatting['italic'] = val != '0' and val != 'false'
        
        # Underline
        u = rPr.find('w:u', NAMESPACES)
        if u is not None:
            formatting['underline'] = u.get('{%s}val' % NAMESPACES['w']) != 'none'
        
        # Color
        color = rPr.find('w:color', NAMESPACES)
        if color is not None:
            formatting['color'] = color.get('{%s}val' % NAMESPACES['w'])
        
        return formatting
    
    def _parse_document(self, docx: zipfile.ZipFile):
        """Parse the main document content."""
        doc_xml = docx.read('word/document.xml')
        root = ET.fromstring(doc_xml)
        
        body = root.find('.//w:body', NAMESPACES)
        if body is None:
            return
        
        for child in body:
            tag = child.tag.split('}')[-1]
            
            if tag == 'p':
                block = self._parse_paragraph(child)
                if block:
                    self.content_blocks.append(block)
            elif tag == 'tbl':
                block = self._parse_table(child)
                if block:
                    self.content_blocks.append(block)
    
    def _parse_paragraph(self, p_elem) -> Optional[Dict[str, Any]]:
        """Parse a paragraph element."""
        block = {
            'type': 'paragraph',
            'style': None,
            'runs': [],
            'images': []
        }
        
        # Get paragraph style
        pPr = p_elem.find('w:pPr', NAMESPACES)
        if pPr is not None:
            pStyle = pPr.find('w:pStyle', NAMESPACES)
            if pStyle is not None:
                style_id = pStyle.get('{%s}val' % NAMESPACES['w'])
                # Map to style name if available
                if style_id in self.styles:
                    block['style'] = self.styles[style_id]['name']
                else:
                    block['style'] = style_id
        
        # Parse runs (text segments)
        for r in p_elem.findall('.//w:r', NAMESPACES):
            run = self._parse_run(r)
            if run:
                block['runs'].append(run)
        
        # Parse inline images
        for drawing in p_elem.findall('.//w:drawing', NAMESPACES):
            img_info = self._parse_drawing(drawing)
            if img_info:
                block['images'].append(img_info)
        
        # Skip empty paragraphs (no text, no images)
        if not block['runs'] and not block['images']:
            # Still include if it has a style (might be intentional spacing)
            if not block['style']:
                return None
        
        return block
    
    def _parse_run(self, r_elem) -> Optional[Dict[str, Any]]:
        """Parse a run (text segment) element."""
        text_elem = r_elem.find('w:t', NAMESPACES)
        if text_elem is None or text_elem.text is None:
            return None
        
        run = {
            'text': text_elem.text,
            'formatting': {}
        }
        
        # Get run properties
        rPr = r_elem.find('w:rPr', NAMESPACES)
        if rPr is not None:
            run['formatting'] = self._extract_run_formatting(rPr)
        
        return run
    
    def _parse_drawing(self, drawing_elem) -> Optional[Dict[str, Any]]:
        """Parse a drawing (image) element."""
        # Look for blip (image reference)
        blip = drawing_elem.find('.//a:blip', NAMESPACES)
        if blip is None:
            return None
        
        embed_id = blip.get('{%s}embed' % NAMESPACES['r'])
        if not embed_id or embed_id not in self.relationships:
            return None
        
        rel = self.relationships[embed_id]
        
        # Get image dimensions
        extent = drawing_elem.find('.//wp:extent', NAMESPACES)
        width = height = None
        if extent is not None:
            # EMUs to pixels (914400 EMUs per inch, assume 96 DPI)
            cx = extent.get('cx')
            cy = extent.get('cy')
            if cx:
                width = int(int(cx) / 914400 * 96)
            if cy:
                height = int(int(cy) / 914400 * 96)
        
        return {
            'id': embed_id,
            'target': rel['target'],
            'width': width,
            'height': height
        }
    
    def _parse_table(self, tbl_elem) -> Dict[str, Any]:
        """Parse a table element."""
        table = {
            'type': 'table',
            'rows': []
        }
        
        for tr in tbl_elem.findall('.//w:tr', NAMESPACES):
            row = []
            for tc in tr.findall('.//w:tc', NAMESPACES):
                cell_content = []
                for p in tc.findall('.//w:p', NAMESPACES):
                    para = self._parse_paragraph(p)
                    if para:
                        cell_content.append(para)
                row.append(cell_content)
            table['rows'].append(row)
        
        return table
    
    def _extract_images(self, docx: zipfile.ZipFile):
        """Extract embedded images to temp directory."""
        for rel_id, rel in self.relationships.items():
            if 'image' in rel.get('type', '').lower():
                target = rel['target']
                # Handle relative paths
                if target.startswith('/'):
                    img_path = target[1:]
                else:
                    img_path = 'word/' + target
                
                try:
                    img_data = docx.read(img_path)
                    
                    # Determine filename
                    filename = os.path.basename(target)
                    output_path = os.path.join(self.temp_dir, filename)
                    
                    with open(output_path, 'wb') as f:
                        f.write(img_data)
                    
                    self.extracted_images.append({
                        'id': rel_id,
                        'original_path': target,
                        'extracted_path': output_path,
                        'size_bytes': len(img_data)
                    })
                except KeyError:
                    pass  # Image not found in archive
    
    def cleanup(self):
        """Remove temporary image files."""
        import shutil
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)


def parse_docx(docx_path: str) -> Dict[str, Any]:
    """
    Convenience function to parse a DOCX file.
    
    Args:
        docx_path: Path to the DOCX file
        
    Returns:
        Structured content dictionary
    """
    parser = DocxParser(docx_path)
    return parser.parse()


def get_style_mapping_template(parsed_content: Dict[str, Any]) -> Dict[str, str]:
    """
    Generate a style mapping template from parsed content.
    
    Returns a dict mapping Word style names to suggested InDesign style names.
    Users can modify this mapping before import.
    """
    mapping = {}
    
    # Get unique styles from content
    styles_used = set()
    for block in parsed_content.get('content', []):
        if block.get('style'):
            styles_used.add(block['style'])
    
    # Create default mapping (Word style -> InDesign style)
    # Common mappings
    common_mappings = {
        'Normal': 'Body Text',
        'Heading 1': 'Heading 1',
        'Heading 2': 'Heading 2',
        'Heading 3': 'Heading 3',
        'Heading 4': 'Heading 4',
        'Title': 'Title',
        'Subtitle': 'Subtitle',
        'Quote': 'Block Quote',
        'List Paragraph': 'List Item',
        'Caption': 'Caption',
    }
    
    for style in styles_used:
        if style in common_mappings:
            mapping[style] = common_mappings[style]
        else:
            # Keep the same name by default
            mapping[style] = style
    
    return mapping
