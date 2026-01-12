"""
PowerPoint generation utilities using python-pptx.
Generates .pptx files from structured JSON slide data.
"""

import io
import base64
import logging
from typing import Dict, List, Any, Optional
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

from open_webui.env import SRC_LOG_LEVELS

log = logging.getLogger(__name__)
log.setLevel(SRC_LOG_LEVELS["MAIN"])


def hex_to_rgb(hex_color: str) -> RGBColor:
    """Convert hex color string to RGBColor."""
    hex_color = hex_color.lstrip('#')
    if len(hex_color) != 6:
        return RGBColor(0, 0, 0)  # Default to black
    try:
        return RGBColor(
            int(hex_color[0:2], 16),
            int(hex_color[2:4], 16),
            int(hex_color[4:6], 16)
        )
    except ValueError:
        return RGBColor(0, 0, 0)


def format_paragraph(paragraph, size: int = 18, bold: bool = False, font_name: str = 'Calibri'):
    """Apply formatting to a paragraph."""
    paragraph.font.size = Pt(size)
    paragraph.font.bold = bold
    paragraph.font.name = font_name


def format_text_frame(text_frame, size: int = 18, bold: bool = False, font_name: str = 'Calibri'):
    """Apply formatting to all paragraphs in a text frame."""
    for paragraph in text_frame.paragraphs:
        format_paragraph(paragraph, size, bold, font_name)


def format_cell(cell, bold: bool = False, bg_color: Optional[str] = None, font_size: int = 14):
    """Format a table cell."""
    text_frame = cell.text_frame
    if text_frame.paragraphs:
        text_frame.paragraphs[0].font.size = Pt(font_size)
        text_frame.paragraphs[0].font.bold = bold
        text_frame.paragraphs[0].font.name = 'Calibri'
    
    if bg_color:
        fill = cell.fill
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(bg_color)


def add_content_to_slide(slide, content_items: List[Dict[str, Any]], prs: Presentation):
    """Add various content types to a slide."""
    # Starting position for content (below title area)
    top = Inches(1.8)
    left = Inches(0.75)
    width = Inches(8.5)
    slide_height = prs.slide_height
    
    for item in content_items:
        item_type = item.get('type', '')
        
        if item_type == 'text':
            text = item.get('text', '')
            textbox = slide.shapes.add_textbox(left, top, width, Inches(1))
            text_frame = textbox.text_frame
            text_frame.word_wrap = True
            text_frame.text = text
            format_text_frame(text_frame, size=18)
            top += Inches(0.8)
        
        elif item_type == 'bullet':
            items = item.get('items', [])
            if not items:
                continue
            
            # Calculate height based on number of items
            height = Inches(len(items) * 0.45 + 0.2)
            textbox = slide.shapes.add_textbox(left, top, width, height)
            text_frame = textbox.text_frame
            text_frame.word_wrap = True
            
            for i, bullet_text in enumerate(items):
                if i == 0:
                    p = text_frame.paragraphs[0]
                else:
                    p = text_frame.add_paragraph()
                
                p.text = f"• {bullet_text}"
                p.level = 0
                format_paragraph(p, size=16)
                p.space_after = Pt(6)
            
            top += height + Inches(0.2)
        
        elif item_type == 'image':
            src = item.get('src', '')
            if not src:
                continue
            
            try:
                if src.startswith('data:'):
                    # Base64 encoded image
                    # Format: data:image/png;base64,XXXX
                    header, data = src.split(',', 1)
                    image_data = base64.b64decode(data)
                    image_stream = io.BytesIO(image_data)
                    
                    # Add image with max width of 6 inches, maintaining aspect ratio
                    pic = slide.shapes.add_picture(image_stream, left, top, width=Inches(6))
                    
                    # Calculate height based on aspect ratio
                    image_height = pic.height
                    top += Inches(image_height.inches + 0.3)
                else:
                    # URL-based image - skip for now (would need async download)
                    log.warning(f"URL-based images not yet supported: {src[:50]}...")
                    continue
            except Exception as e:
                log.error(f"Error adding image: {e}")
                continue
        
        elif item_type == 'table':
            headers = item.get('headers', [])
            rows = item.get('rows', [])
            
            if not headers:
                continue
            
            num_rows = len(rows) + 1  # +1 for header
            num_cols = len(headers)
            
            # Calculate table dimensions
            table_height = Inches(num_rows * 0.4 + 0.2)
            
            table_shape = slide.shapes.add_table(
                num_rows, num_cols,
                left, top,
                width, table_height
            )
            table = table_shape.table
            
            # Style the table
            # Add headers
            for col_idx, header in enumerate(headers):
                cell = table.cell(0, col_idx)
                cell.text = str(header)
                format_cell(cell, bold=True, bg_color='4472C4', font_size=12)
                # White text on blue background
                if cell.text_frame.paragraphs:
                    cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
            
            # Add data rows
            for row_idx, row_data in enumerate(rows):
                for col_idx, cell_data in enumerate(row_data):
                    if col_idx < num_cols:  # Ensure we don't exceed column count
                        cell = table.cell(row_idx + 1, col_idx)
                        cell.text = str(cell_data) if cell_data is not None else ''
                        format_cell(cell, font_size=11)
                        # Alternating row colors
                        if row_idx % 2 == 1:
                            format_cell(cell, bg_color='E8EEF7', font_size=11)
            
            top += table_height + Inches(0.3)
        
        # Safety check - don't overflow the slide
        if top > Inches(6.5):
            log.warning("Content may overflow slide boundaries")
            break


def create_slide(prs: Presentation, slide_info: Dict[str, Any]) -> Any:
    """Create a single slide based on slide_info structure."""
    
    title = slide_info.get('title', '')
    content = slide_info.get('content', [])
    background_color = slide_info.get('backgroundColor')
    notes = slide_info.get('notes', '')
    
    # Determine layout based on content
    # Layout 0 = Title Slide
    # Layout 1 = Title and Content
    # Layout 5 = Title Only
    # Layout 6 = Blank
    
    if title and not content:
        # Title-only slide
        layout_index = 5
    elif title and content:
        # Title and content
        layout_index = 5  # Use Title Only and add content manually for better control
    else:
        # Blank
        layout_index = 6
    
    # Ensure layout index is valid
    if layout_index >= len(prs.slide_layouts):
        layout_index = 6 if len(prs.slide_layouts) > 6 else 0
    
    slide_layout = prs.slide_layouts[layout_index]
    slide = prs.slides.add_slide(slide_layout)
    
    # Set background color if specified
    if background_color:
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(background_color)
    
    # Add title
    if title:
        if slide.shapes.title:
            title_shape = slide.shapes.title
            title_shape.text = title
            format_text_frame(title_shape.text_frame, size=32, bold=True)
        else:
            # Create title textbox manually
            title_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(0.3),
                Inches(9), Inches(1)
            )
            title_frame = title_box.text_frame
            title_frame.text = title
            format_text_frame(title_frame, size=32, bold=True)
    
    # Add content
    if content:
        add_content_to_slide(slide, content, prs)
    
    # Add speaker notes
    if notes:
        notes_slide = slide.notes_slide
        text_frame = notes_slide.notes_text_frame
        text_frame.text = notes
    
    return slide


def generate_pptx_from_data(
    slide_data: Dict[str, Any],
    template_path: Optional[str] = None
) -> bytes:
    """
    Generate a PowerPoint file from structured slide data.
    
    Args:
        slide_data: Dictionary containing presentation structure:
            {
                "title": "Presentation Title",
                "slides": [
                    {
                        "title": "Slide Title",
                        "backgroundColor": "#FFFFFF",
                        "content": [
                            {"type": "text", "text": "Some text"},
                            {"type": "bullet", "items": ["Item 1", "Item 2"]},
                            {"type": "table", "headers": [...], "rows": [...]},
                            {"type": "image", "src": "data:image/png;base64,..."}
                        ],
                        "notes": "Speaker notes"
                    }
                ]
            }
        template_path: Optional path to a .pptx template file
        
    Returns:
        bytes: PPTX file content as bytes
    """
    # Create presentation from template or blank
    if template_path and Path(template_path).exists():
        try:
            prs = Presentation(template_path)
            log.info(f"Using template: {template_path}")
            # Clear existing slides if template has sample content
            # (Keep this commented - templates should be clean)
            # while len(prs.slides) > 0:
            #     rId = prs.slides._sldIdLst[0].rId
            #     prs.part.drop_rel(rId)
            #     del prs.slides._sldIdLst[0]
        except Exception as e:
            log.warning(f"Failed to load template: {e}. Using blank presentation.")
            prs = Presentation()
    else:
        prs = Presentation()
    
    # Set slide dimensions (16:9 aspect ratio)
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Get slides from data
    slides = slide_data.get('slides', [])
    
    if not slides:
        # Create at least one title slide if no slides provided
        title = slide_data.get('title', 'Untitled Presentation')
        slides = [{'title': title, 'content': []}]
    
    # Create each slide
    for slide_info in slides:
        create_slide(prs, slide_info)
    
    # Save to bytes buffer
    pptx_bytes = io.BytesIO()
    prs.save(pptx_bytes)
    pptx_bytes.seek(0)
    
    return pptx_bytes.getvalue()


def get_presentation_info(pptx_path: str) -> Dict[str, Any]:
    """
    Extract metadata from an existing presentation.
    
    Args:
        pptx_path: Path to the .pptx file
        
    Returns:
        Dictionary with presentation metadata
    """
    try:
        prs = Presentation(pptx_path)
        
        layouts = []
        for i, layout in enumerate(prs.slide_layouts):
            layouts.append({
                'index': i,
                'name': layout.name
            })
        
        return {
            'slide_count': len(prs.slides),
            'width_inches': prs.slide_width.inches,
            'height_inches': prs.slide_height.inches,
            'layouts': layouts
        }
    except Exception as e:
        log.error(f"Error reading presentation info: {e}")
        return {'error': str(e)}