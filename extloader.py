import fitz  # PyMuPDF
import base64
import logging
import os
import tempfile
import time
import hashlib
import json
import re
import unicodedata
from typing import Dict, List, Tuple, Any, Optional
from fastapi import FastAPI, Request, HTTPException, status
from fastapi.responses import JSONResponse
from dataclasses import dataclass, asdict
from datetime import datetime

# Fast native document processors
try:
    from docx import Document as DocxDocument
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    from pptx import Presentation
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

try:
    from striprtf.striprtf import rtf_to_text
    RTF_AVAILABLE = True
except ImportError:
    RTF_AVAILABLE = False

# Image processing
try:
    from PIL import Image
    import io
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False

# Optional: pymupdf4llm for enhanced markdown extraction
try:
    import pymupdf4llm
    PYMUPDF4LLM_AVAILABLE = True
except ImportError:
    PYMUPDF4LLM_AVAILABLE = False

# Azure Document Intelligence compatible output formats
AZURE_DOC_INTEL_COMPATIBLE = True
DEFAULT_OUTPUT_FORMAT = "json"  # or "markdown"

# Image processing efficiency settings
AUTO_COMPRESS_IMAGES = os.getenv("AUTO_COMPRESS_IMAGES", "true").lower() == "true"
DEFAULT_COMPRESSION_WIDTH = int(os.getenv("FILE_IMAGE_COMPRESSION_WIDTH", "1024"))
DEFAULT_COMPRESSION_HEIGHT = int(os.getenv("FILE_IMAGE_COMPRESSION_HEIGHT", "1024"))
COMPRESSION_QUALITY = int(os.getenv("IMAGE_COMPRESSION_QUALITY", "85"))

# Image filtering thresholds
MIN_IMAGE_WIDTH = int(os.getenv("MIN_IMAGE_WIDTH", "32"))
MIN_IMAGE_HEIGHT = int(os.getenv("MIN_IMAGE_HEIGHT", "32"))
MIN_IMAGE_SIZE_BYTES = int(os.getenv("MIN_IMAGE_SIZE_BYTES", "1024"))

# Layout detection settings
ENABLE_MULTI_COLUMN_DETECTION = os.getenv("ENABLE_MULTI_COLUMN_DETECTION", "true").lower() == "true"
ENABLE_HEADER_FOOTER_DETECTION = os.getenv("ENABLE_HEADER_FOOTER_DETECTION", "true").lower() == "true"
HEADER_MARGIN_RATIO = float(os.getenv("HEADER_MARGIN_RATIO", "0.08"))  # Top 8% of page
FOOTER_MARGIN_RATIO = float(os.getenv("FOOTER_MARGIN_RATIO", "0.08"))  # Bottom 8% of page

# Debug output settings
EXTLOADER_DEBUG_OUTPUT = os.getenv("EXTLOADER_DEBUG_OUTPUT", "false").lower() == "true"
EXTLOADER_DEBUG_PATH = os.getenv("EXTLOADER_DEBUG_PATH", "/app/backend/data/extloader_debug")

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize the FastAPI application
app = FastAPI(
    title="Enhanced Content Processing Engine",
    description="An API to extract text, tables, layout, and metadata from documents for OpenWebUI.",
    version="2.0.0"
)

# =============================================================================
# DATA CLASSES FOR STRUCTURED OUTPUT
# =============================================================================

@dataclass
class DocumentMetadata:
    """Document-level metadata extracted from PDF."""
    title: Optional[str] = None
    author: Optional[str] = None
    subject: Optional[str] = None
    keywords: Optional[str] = None
    creator: Optional[str] = None
    producer: Optional[str] = None
    creation_date: Optional[str] = None
    modification_date: Optional[str] = None
    format: Optional[str] = None
    encryption: Optional[str] = None
    page_count: int = 0
    
    def to_dict(self) -> dict:
        return {k: v for k, v in asdict(self).items() if v is not None}

@dataclass
class TableCell:
    """A single cell in an extracted table."""
    row_index: int
    col_index: int
    content: str
    row_span: int = 1
    col_span: int = 1
    is_header: bool = False

@dataclass
class ExtractedTable:
    """A table extracted from a document page."""
    page_number: int
    table_index: int
    row_count: int
    col_count: int
    cells: List[TableCell]
    bbox: Tuple[float, float, float, float]  # x0, y0, x1, y1
    markdown: str = ""
    header_external: bool = False
    
    def to_dict(self) -> dict:
        return {
            "page_number": self.page_number,
            "table_index": self.table_index,
            "row_count": self.row_count,
            "col_count": self.col_count,
            "cells": [asdict(c) for c in self.cells],
            "bbox": self.bbox,
            "markdown": self.markdown,
            "header_external": self.header_external
        }

@dataclass
class PageRegion:
    """A detected region on a page (header, footer, column, etc.)."""
    region_type: str  # "header", "footer", "column", "body"
    bbox: Tuple[float, float, float, float]
    content: str = ""
    page_number: int = 0

@dataclass 
class LayoutInfo:
    """Layout information for a page."""
    page_number: int
    width: float
    height: float
    column_count: int = 1
    has_header: bool = False
    has_footer: bool = False
    header_content: str = ""
    footer_content: str = ""
    columns: List[Tuple[float, float, float, float]] = None  # List of column bboxes
    
    def __post_init__(self):
        if self.columns is None:
            self.columns = []

# =============================================================================
# TEXT NORMALIZATION
# =============================================================================

def normalize_text_encoding(text: str) -> str:
    """
    Normalize text encoding to handle unicode issues comprehensively.
    """
    if not text:
        return ""
    
    try:
        normalized = unicodedata.normalize('NFKC', text)
        
        replacements = {
            '\u00a0': ' ',
            '\xa0': ' ',
            '\u2013': '-',
            '\u2014': '--',
            '\u2018': "'",
            '\u2019': "'",
            '\u201c': '"',
            '\u201d': '"',
            '\u2022': '*',
            '\u2026': '...',
        }
        
        for unicode_char, replacement in replacements.items():
            normalized = normalized.replace(unicode_char, replacement)
        
        normalized = ''.join(
            char for char in normalized 
            if char == '\n' or char == '\t' or not unicodedata.category(char).startswith('C')
        )
        
        normalized = normalized.encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
        
        return normalized
        
    except Exception as e:
        logger.warning(f"Text normalization failed, returning original text: {e}")
        try:
            return text.replace('\u00a0', ' ').replace('\xa0', ' ')
        except:
            return text

# =============================================================================
# DEBUG OUTPUT
# =============================================================================

def write_debug_output(filename: str, page_content: str, full_result: dict) -> None:
    """
    Write extraction output to debug files for inspection.
    Only runs if EXTLOADER_DEBUG_OUTPUT is enabled.
    """
    if not EXTLOADER_DEBUG_OUTPUT:
        return
    
    try:
        # Ensure debug directory exists
        os.makedirs(EXTLOADER_DEBUG_PATH, exist_ok=True)
        
        # Create unique filename prefix using timestamp and source filename hash
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename_hash = hashlib.md5(filename.encode()).hexdigest()[:8]
        safe_filename = re.sub(r'[^\w\-.]', '_', filename)[:50]
        prefix = f"{timestamp}_{safe_filename}_{filename_hash}"
        
        # Write markdown content
        md_path = os.path.join(EXTLOADER_DEBUG_PATH, f"{prefix}_content.md")
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(f"# Extraction Debug Output\n")
            f.write(f"**Source:** {filename}\n")
            f.write(f"**Timestamp:** {datetime.now().isoformat()}\n\n")
            f.write("---\n\n")
            f.write(page_content)
        
        # Write full JSON result
        json_path = os.path.join(EXTLOADER_DEBUG_PATH, f"{prefix}_full.json")
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(full_result, f, indent=2, ensure_ascii=False, default=str)
        
        logger.info(f"Debug output written: {md_path}, {json_path}")
        
    except Exception as e:
        logger.warning(f"Failed to write debug output: {e}")

# =============================================================================
# METADATA EXTRACTION
# =============================================================================

def extract_pdf_metadata(doc: fitz.Document) -> DocumentMetadata:
    """
    Extract comprehensive metadata from a PDF document.
    """
    metadata = DocumentMetadata()
    
    try:
        # Get standard metadata dictionary
        pdf_metadata = doc.metadata
        
        if pdf_metadata:
            metadata.title = pdf_metadata.get('title') or None
            metadata.author = pdf_metadata.get('author') or None
            metadata.subject = pdf_metadata.get('subject') or None
            metadata.keywords = pdf_metadata.get('keywords') or None
            metadata.creator = pdf_metadata.get('creator') or None
            metadata.producer = pdf_metadata.get('producer') or None
            metadata.format = pdf_metadata.get('format') or None
            metadata.encryption = pdf_metadata.get('encryption') or None
            
            # Parse creation and modification dates
            creation_date = pdf_metadata.get('creationDate')
            if creation_date:
                metadata.creation_date = parse_pdf_date(creation_date)
            
            mod_date = pdf_metadata.get('modDate')
            if mod_date:
                metadata.modification_date = parse_pdf_date(mod_date)
        
        metadata.page_count = doc.page_count
        
        logger.info(f"Extracted metadata: title='{metadata.title}', author='{metadata.author}', pages={metadata.page_count}")
        
    except Exception as e:
        logger.warning(f"Error extracting metadata: {e}")
        metadata.page_count = doc.page_count if doc else 0
    
    return metadata

def parse_pdf_date(date_string: str) -> Optional[str]:
    """
    Parse PDF date format (D:YYYYMMDDHHmmSSOHH'mm') to ISO format.
    """
    if not date_string:
        return None
    
    try:
        # Remove the "D:" prefix if present
        if date_string.startswith("D:"):
            date_string = date_string[2:]
        
        # Try to parse the core date/time part (YYYYMMDDHHmmSS)
        # Handle various lengths of date strings
        year = date_string[0:4] if len(date_string) >= 4 else None
        month = date_string[4:6] if len(date_string) >= 6 else "01"
        day = date_string[6:8] if len(date_string) >= 8 else "01"
        hour = date_string[8:10] if len(date_string) >= 10 else "00"
        minute = date_string[10:12] if len(date_string) >= 12 else "00"
        second = date_string[12:14] if len(date_string) >= 14 else "00"
        
        if year:
            return f"{year}-{month}-{day}T{hour}:{minute}:{second}"
        
    except Exception as e:
        logger.debug(f"Could not parse PDF date '{date_string}': {e}")
    
    return date_string  # Return original if parsing fails

def get_table_of_contents(doc: fitz.Document) -> List[Dict]:
    """
    Extract the table of contents (bookmarks/outlines) from a PDF.
    """
    toc = []
    try:
        raw_toc = doc.get_toc()
        for entry in raw_toc:
            if len(entry) >= 3:
                toc.append({
                    "level": entry[0],
                    "title": normalize_text_encoding(entry[1]),
                    "page": entry[2]
                })
        logger.info(f"Extracted {len(toc)} TOC entries")
    except Exception as e:
        logger.warning(f"Error extracting TOC: {e}")
    
    return toc

# =============================================================================
# TABLE EXTRACTION
# =============================================================================

def extract_tables_from_page(page: fitz.Page, page_num: int) -> List[ExtractedTable]:
    """
    Extract tables from a PDF page using PyMuPDF's find_tables() method.
    """
    extracted_tables = []
    
    try:
        # Use PyMuPDF's table finder (available since v1.23.0)
        tables = page.find_tables()
        
        for table_idx, table in enumerate(tables.tables):
            try:
                # Get table data
                table_data = table.extract()
                
                if not table_data or len(table_data) == 0:
                    continue
                
                # Determine dimensions
                row_count = len(table_data)
                col_count = max(len(row) for row in table_data) if table_data else 0
                
                # Build cells list
                cells = []
                for row_idx, row in enumerate(table_data):
                    for col_idx, cell_content in enumerate(row):
                        if cell_content is not None:
                            cell_text = normalize_text_encoding(str(cell_content))
                            cells.append(TableCell(
                                row_index=row_idx,
                                col_index=col_idx,
                                content=cell_text,
                                is_header=(row_idx == 0)  # First row as header by default
                            ))
                
                # Generate markdown representation
                markdown = table_to_markdown(table_data)
                
                # Get bounding box
                bbox = table.bbox if hasattr(table, 'bbox') else (0, 0, 0, 0)
                
                # Check if header is external
                header_external = False
                if hasattr(table, 'header') and table.header:
                    header_external = getattr(table.header, 'external', False)
                
                extracted_table = ExtractedTable(
                    page_number=page_num,
                    table_index=table_idx,
                    row_count=row_count,
                    col_count=col_count,
                    cells=cells,
                    bbox=bbox,
                    markdown=markdown,
                    header_external=header_external
                )
                
                extracted_tables.append(extracted_table)
                logger.debug(f"Extracted table {table_idx} from page {page_num}: {row_count}x{col_count}")
                
            except Exception as e:
                logger.warning(f"Error extracting table {table_idx} from page {page_num}: {e}")
                continue
        
        logger.info(f"Extracted {len(extracted_tables)} tables from page {page_num}")
        
    except Exception as e:
        logger.warning(f"Error in table extraction for page {page_num}: {e}")
    
    return extracted_tables

def table_to_markdown(table_data: List[List]) -> str:
    """
    Convert table data to markdown format.
    """
    if not table_data or len(table_data) == 0:
        return ""
    
    # Normalize all cells to strings
    normalized = []
    for row in table_data:
        normalized_row = []
        for cell in row:
            cell_str = normalize_text_encoding(str(cell)) if cell is not None else ""
            # Escape pipe characters in cell content
            cell_str = cell_str.replace("|", "\\|")
            normalized_row.append(cell_str)
        normalized.append(normalized_row)
    
    # Ensure all rows have same number of columns
    max_cols = max(len(row) for row in normalized)
    for row in normalized:
        while len(row) < max_cols:
            row.append("")
    
    # Build markdown
    lines = []
    
    # Header row
    if normalized:
        lines.append("| " + " | ".join(normalized[0]) + " |")
        # Separator row
        lines.append("|" + "|".join(["---"] * max_cols) + "|")
        
        # Data rows
        for row in normalized[1:]:
            lines.append("| " + " | ".join(row) + " |")
    
    return "\n".join(lines)

# =============================================================================
# LAYOUT DETECTION (Multi-column, Headers, Footers)
# =============================================================================

def detect_page_layout(page: fitz.Page, page_num: int) -> LayoutInfo:
    """
    Detect page layout including columns, headers, and footers.
    """
    rect = page.rect
    width = rect.width
    height = rect.height
    
    layout = LayoutInfo(
        page_number=page_num,
        width=width,
        height=height
    )
    
    if not ENABLE_HEADER_FOOTER_DETECTION and not ENABLE_MULTI_COLUMN_DETECTION:
        return layout
    
    try:
        # Get text blocks with position information
        blocks = page.get_text("dict", sort=True)["blocks"]
        text_blocks = [b for b in blocks if b.get("type") == 0]  # Type 0 = text blocks
        
        if not text_blocks:
            return layout
        
        # Calculate header/footer boundaries
        header_boundary = height * HEADER_MARGIN_RATIO
        footer_boundary = height * (1 - FOOTER_MARGIN_RATIO)
        
        # Detect headers and footers
        if ENABLE_HEADER_FOOTER_DETECTION:
            header_blocks = []
            footer_blocks = []
            body_blocks = []
            
            for block in text_blocks:
                bbox = block.get("bbox", (0, 0, 0, 0))
                block_top = bbox[1]
                block_bottom = bbox[3]
                
                if block_bottom < header_boundary:
                    header_blocks.append(block)
                elif block_top > footer_boundary:
                    footer_blocks.append(block)
                else:
                    body_blocks.append(block)
            
            if header_blocks:
                layout.has_header = True
                layout.header_content = extract_text_from_blocks(header_blocks)
            
            if footer_blocks:
                layout.has_footer = True
                layout.footer_content = extract_text_from_blocks(footer_blocks)
        
        # Detect columns
        if ENABLE_MULTI_COLUMN_DETECTION:
            columns = detect_columns(page, text_blocks)
            layout.column_count = len(columns) if columns else 1
            layout.columns = columns
        
    except Exception as e:
        logger.warning(f"Error detecting layout for page {page_num}: {e}")
    
    return layout

def detect_columns(page: fitz.Page, text_blocks: List[Dict]) -> List[Tuple[float, float, float, float]]:
    """
    Detect text columns on a page by analyzing text block positions.
    """
    if not text_blocks:
        return []
    
    rect = page.rect
    page_width = rect.width
    page_height = rect.height
    
    # Filter to body area (exclude likely headers/footers)
    header_boundary = page_height * HEADER_MARGIN_RATIO
    footer_boundary = page_height * (1 - FOOTER_MARGIN_RATIO)
    
    body_blocks = []
    for block in text_blocks:
        bbox = block.get("bbox", (0, 0, 0, 0))
        block_center_y = (bbox[1] + bbox[3]) / 2
        if header_boundary < block_center_y < footer_boundary:
            body_blocks.append(block)
    
    if not body_blocks:
        return [(0, 0, page_width, page_height)]
    
    # Collect all x-coordinates of block edges
    # Find gaps in x-coordinates that could be column separators
    # A gap is significant if it's wider than 5% of page width
    min_gap = page_width * 0.05
    
    # Track coverage intervals
    intervals = []
    for block in body_blocks:
        bbox = block.get("bbox", (0, 0, 0, 0))
        intervals.append((bbox[0], bbox[2]))
    
    # Merge overlapping intervals
    intervals.sort()
    merged = []
    for start, end in intervals:
        if merged and start <= merged[-1][1] + min_gap:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    
    # Each merged interval represents a column
    columns = []
    for x0, x1 in merged:
        # Find y extent for this column
        col_blocks = [b for b in body_blocks 
                      if b["bbox"][0] >= x0 - min_gap and b["bbox"][2] <= x1 + min_gap]
        if col_blocks:
            y0 = min(b["bbox"][1] for b in col_blocks)
            y1 = max(b["bbox"][3] for b in col_blocks)
            columns.append((x0, y0, x1, y1))
    
    # If we only detected one column spanning most of the page, treat as single column
    if len(columns) == 1:
        col = columns[0]
        if (col[2] - col[0]) > page_width * 0.7:
            return [(0, 0, page_width, page_height)]
    
    return columns if columns else [(0, 0, page_width, page_height)]

def extract_text_from_blocks(blocks: List[Dict]) -> str:
    """
    Extract and concatenate text from a list of text blocks.
    """
    text_parts = []
    
    for block in blocks:
        if block.get("type") != 0:  # Skip non-text blocks
            continue
        
        for line in block.get("lines", []):
            line_text = ""
            for span in line.get("spans", []):
                span_text = span.get("text", "")
                line_text += span_text
            if line_text.strip():
                text_parts.append(normalize_text_encoding(line_text.strip()))
    
    return " | ".join(text_parts)

def is_block_in_table(block_bbox: Tuple[float, float, float, float], table_bboxes: List[Tuple[float, float, float, float]]) -> bool:
    """
    Check if a text block is located inside any detected table.
    
    Args:
        block_bbox: (x0, y0, x1, y1) of the text block
        table_bboxes: List of table bounding boxes
        
    Returns:
        True if the block overlaps significantly with a table
    """
    if not table_bboxes:
        return False
        
    bx0, by0, bx1, by1 = block_bbox
    block_area = (bx1 - bx0) * (by1 - by0)
    
    if block_area <= 0:
        return False
    
    block_rect = fitz.Rect(block_bbox)
    
    for table_bbox in table_bboxes:
        table_rect = fitz.Rect(table_bbox)
        # Calculate intersection
        intersect = block_rect & table_rect # Intersection rectangle
        
        if intersect.is_valid and not intersect.is_empty:
            intersect_area = intersect.get_area()
            # If >50% of the text block is inside the table, consider it part of the table
            if intersect_area / block_area > 0.5:
                return True
                
    return False

def extract_text_with_layout(page: fitz.Page, layout: LayoutInfo, exclude_headers_footers: bool = False, table_bboxes: List[Tuple] = None) -> str:
    """
    Extract text from a page respecting layout and excluding specific regions (tables, headers, footers).
    
    Uses PyMuPDF's 'blocks' output (sort=True) to ensure reading order (columns) is respected,
    while allowing for coordinate-based filtering of tables and artifacts.
    
    Args:
        page: PyMuPDF Page object
        layout: LayoutInfo with detected columns/regions
        exclude_headers_footers: Whether to exclude header/footer text
        table_bboxes: List of table bounding boxes to exclude
        
    Returns:
        Text extracted in reading order
    """
    table_bboxes = table_bboxes or []
    
    # Get all text blocks. sort=True organizes them in reading order (top-left to bottom-right),
    # effectively handling standard multi-column layouts automatically.
    blocks = page.get_text("blocks", sort=True)
    
    cleaned_text_parts = []
    
    for block in blocks:
        # block format: (x0, y0, x1, y1, "text", block_no, block_type)
        if block[6] != 0: # Skip non-text blocks (images, graphics)
            continue
            
        bbox = block[:4]
        text = block[4]
        
        if not text.strip():
            continue
            
        # 1. Check Table Exclusion (Critical to prevent data duplication)
        if is_block_in_table(bbox, table_bboxes):
            continue
            
        # 2. Check Header/Footer Exclusion
        if exclude_headers_footers:
            y_center = (bbox[1] + bbox[3]) / 2
            
            # Filter Header
            if layout.has_header and y_center < (layout.height * HEADER_MARGIN_RATIO):
                continue
                
            # Filter Footer
            if layout.has_footer and y_center > (layout.height * (1 - FOOTER_MARGIN_RATIO)):
                continue

        # 3. Add Valid Text
        # Normalize encoding to fix common PDF font issues
        normalized_text = normalize_text_encoding(text.strip())
        if normalized_text:
            cleaned_text_parts.append(normalized_text)
            
    return "\n\n".join(cleaned_text_parts)

# =============================================================================
# ENHANCED PDF PROCESSING
# =============================================================================

def process_pdf_enhanced(
    file_bytes: bytes, 
    filename: str,
    extract_images_flag: str = "false",
    extract_tables: bool = True,
    detect_layout: bool = True,
    extract_metadata: bool = True,
    exclude_headers_footers: bool = False,
    output_format: str = "json"
) -> Dict[str, Any]:
    """
    Enhanced PDF processing with tables, layout detection, and metadata.
    """
    start_time = time.time()
    
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    
    result = {
        "page_content": "",
        "pages": [],
        "tables": [],
        "metadata": {
            "source": filename,
            "page_count": doc.page_count,
            "processing_status": "completed"
        },
        "images": [],
        "document_metadata": {},
        "table_of_contents": [],
        "layout_info": []
    }
    
    # Extract document metadata
    if extract_metadata:
        doc_metadata = extract_pdf_metadata(doc)
        result["document_metadata"] = doc_metadata.to_dict()
        result["table_of_contents"] = get_table_of_contents(doc)
    
    # Process each page
    all_page_texts = []
    all_tables = []
    all_layouts = []
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        page_idx = page_num + 1  # 1-indexed
        
        # 1. Detect layout
        layout = None
        if detect_layout:
            layout = detect_page_layout(page, page_idx)
            all_layouts.append({
                "page_number": page_idx,
                "width": layout.width,
                "height": layout.height,
                "column_count": layout.column_count,
                "has_header": layout.has_header,
                "has_footer": layout.has_footer,
                "header_content": layout.header_content if layout.has_header else None,
                "footer_content": layout.footer_content if layout.has_footer else None
            })
        else:
            # Create a dummy layout object for passing to extractor
            layout = LayoutInfo(page_idx, page.rect.width, page.rect.height)
        
        # 2. Extract tables FIRST (to get bboxes)
        page_tables = []
        table_bboxes = []
        if extract_tables:
            page_tables = extract_tables_from_page(page, page_idx)
            for table in page_tables:
                all_tables.append(table.to_dict())
                table_bboxes.append(table.bbox) # Collect bboxes for masking
        
        # 3. Extract text (passing layout and table bboxes for exclusion)
        page_text = extract_text_with_layout(
            page, 
            layout, 
            exclude_headers_footers, 
            table_bboxes=table_bboxes
        )

        # 4. Re-inject tables as clean Markdown (ensures anchor verification works)
        if page_tables:
            table_sections = []
            num_tables = len(page_tables)
            for idx, t in enumerate(page_tables, start=1):
                if t.markdown:
                    if num_tables > 1:
                        table_sections.append(f"[Table {idx} of {num_tables}]\n{t.markdown}")
                    else:
                        table_sections.append(t.markdown)
            
            if table_sections:
                table_block = "\n\n--- TABLES ---\n\n" + "\n\n".join(table_sections)
                page_text = (page_text.strip() + table_block) if page_text.strip() else table_block.strip()
        
        # 5. Collect non-empty pages (check AFTER table injection)
        if page_text.strip():
            all_page_texts.append(page_text.strip())
    
    # Extract images
    if extract_images_flag.lower() == "true":
        result["images"] = process_images_efficiently(doc, extract_images_flag, filename)
    
    # Build full text with page delimiters
    full_text_parts = []
    total_pages = len(all_page_texts)
    
    for idx, page_text in enumerate(all_page_texts):
        page_num = idx + 1
        if idx > 0:
            full_text_parts.append(f"\n\nPage {page_num} of {total_pages}\n\n")
        full_text_parts.append(page_text)
    
    result["page_content"] = "".join(full_text_parts)
    result["pages"] = all_page_texts
    result["tables"] = all_tables
    result["layout_info"] = all_layouts
    
    # Update metadata
    result["metadata"]["total_pages"] = doc.page_count
    result["metadata"]["non_empty_pages"] = len(all_page_texts)
    result["metadata"]["tables_extracted"] = len(all_tables)
    result["metadata"]["images_extracted"] = len(result["images"])
    result["metadata"]["processing_time_ms"] = int((time.time() - start_time) * 1000)
    result["metadata"]["layout_detection_enabled"] = detect_layout
    result["metadata"]["multi_column_pages"] = sum(1 for l in all_layouts if l.get("column_count", 1) > 1)
    result["metadata"]["pages_with_headers"] = sum(1 for l in all_layouts if l.get("has_header", False))
    result["metadata"]["pages_with_footers"] = sum(1 for l in all_layouts if l.get("has_footer", False))
    
    doc.close()
    
    # Write debug output if enabled
    write_debug_output(filename, result["page_content"], result)
    
    logger.info(f"Enhanced PDF processing complete: {len(all_page_texts)} pages, {len(all_tables)} tables, {len(result['images'])} images")
    
    return result

# =============================================================================
# IMAGE PROCESSING (Unchanged from original)
# =============================================================================

def compress_image_for_efficiency(image_bytes: bytes, image_ext: str) -> tuple[bytes, bool, tuple[int, int]]:
    """Compress image for processing efficiency if auto-compression is enabled."""
    if not PILLOW_AVAILABLE:
        return image_bytes, False, (0, 0)
    
    try:
        image = Image.open(io.BytesIO(image_bytes))
        original_dimensions = (image.width, image.height)
        
        if not AUTO_COMPRESS_IMAGES:
            return image_bytes, False, original_dimensions
        
        if image.mode in ('RGBA', 'LA', 'P'):
            background = Image.new('RGB', image.size, (255, 255, 255))
            if image.mode == 'P':
                image = image.convert('RGBA')
            background.paste(image, mask=image.split()[-1] if image.mode == 'RGBA' else None)
            image = background
        
        if image.width > DEFAULT_COMPRESSION_WIDTH or image.height > DEFAULT_COMPRESSION_HEIGHT:
            image.thumbnail((DEFAULT_COMPRESSION_WIDTH, DEFAULT_COMPRESSION_HEIGHT), Image.Resampling.LANCZOS)
        
        output = io.BytesIO()
        image.save(output, format='JPEG', quality=COMPRESSION_QUALITY, optimize=True)
        compressed_bytes = output.getvalue()
        
        if len(compressed_bytes) < len(image_bytes):
            return compressed_bytes, True, original_dimensions
        else:
            return image_bytes, False, original_dimensions
            
    except Exception as e:
        logger.warning(f"Image compression failed, using original: {e}")
        return image_bytes, False, (0, 0)

def process_images_efficiently(doc, extract_images_flag: str, filename: str) -> list:
    """Process PDF images with efficiency optimizations, filtering, and deduplication."""
    base64_images = []
    
    if extract_images_flag != 'true':
        return base64_images
    
    logger.info("Starting efficient image extraction...")
    
    seen_hashes = set()
    seen_xrefs = set()
    images_processed = 0
    images_filtered = 0
    images_duplicated = 0
    
    try:
        for page_num in range(len(doc)):
            image_list = doc.get_page_images(page_num, full=True)
            
            for img_index, img in enumerate(image_list):
                try:
                    xref = img[0]
                    
                    if xref in seen_xrefs:
                        images_duplicated += 1
                        continue
                    
                    seen_xrefs.add(xref)
                    
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    
                    if len(image_bytes) < MIN_IMAGE_SIZE_BYTES:
                        images_filtered += 1
                        continue
                    
                    processed_bytes, was_compressed, dimensions = compress_image_for_efficiency(image_bytes, image_ext)
                    width, height = dimensions
                    
                    if width > 0 and height > 0:
                        if width < MIN_IMAGE_WIDTH or height < MIN_IMAGE_HEIGHT:
                            images_filtered += 1
                            continue
                    
                    image_hash = hashlib.md5(processed_bytes).hexdigest()
                    if image_hash in seen_hashes:
                        images_duplicated += 1
                        continue
                    
                    seen_hashes.add(image_hash)
                    
                    encoded_string = base64.b64encode(processed_bytes).decode("utf-8")
                    
                    if was_compressed:
                        data_uri = f"data:image/jpeg;base64,{encoded_string}"
                    else:
                        data_uri = f"data:image/{image_ext};base64,{encoded_string}"
                    
                    base64_images.append(data_uri)
                    images_processed += 1
                    
                except Exception as e:
                    logger.warning(f"Failed to process image on page {page_num + 1}: {e}")
                    continue
        
        logger.info(f"Image processing: {images_processed} extracted, {images_filtered} filtered, {images_duplicated} duplicates")
        
    except Exception as e:
        logger.error(f"Error during image processing: {e}")
    
    return base64_images

# =============================================================================
# OFFICE DOCUMENT PROCESSORS (DOCX, XLSX, PPTX - from original)
# =============================================================================

def extract_docx_structure_azure_format(doc, filename: str, output_format: str = "json") -> str:
    """Extract structured content from DOCX in Azure Document Intelligence format."""
    
    full_content = ""
    content_elements = []
    element_id = 0
    
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            start_offset = len(full_content)
            text_content = normalize_text_encoding(paragraph.text.strip())
            
            style_name = paragraph.style.name if paragraph.style else ""
            if "Heading" in style_name:
                level = 1
                if "Heading 2" in style_name:
                    level = 2
                elif "Heading 3" in style_name:
                    level = 3
                elif "Heading 4" in style_name:
                    level = 4
                elif "Heading 5" in style_name:
                    level = 5
                elif "Heading 6" in style_name:
                    level = 6
                
                element_type = "title"
                role = f"heading{level}"
            else:
                element_type = "paragraph"
                role = "paragraph"
            
            content_elements.append({
                "id": f"element_{element_id}",
                "kind": element_type,
                "role": role,
                "content": text_content,
                "boundingRegions": [{"pageNumber": 1}],
                "spans": [{"offset": start_offset, "length": len(text_content)}]
            })
            
            full_content += text_content + "\n\n"
            element_id += 1
    
    # Process tables
    table_id = 0
    for table in doc.tables:
        start_offset = len(full_content)
        table_content = ""
        
        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                cell_text = normalize_text_encoding(cell.text.strip())
                row_data.append(cell_text)
            table_data.append(row_data)
        
        for row in table_data:
            table_content += " | ".join(row) + "\n"
        
        content_elements.append({
            "id": f"table_{table_id}",
            "kind": "table",
            "rowCount": len(table_data),
            "columnCount": len(table_data[0]) if table_data else 0,
            "cells": [
                {
                    "kind": "content",
                    "rowIndex": row_idx,
                    "columnIndex": col_idx,
                    "content": cell_content,
                }
                for row_idx, row in enumerate(table_data)
                for col_idx, cell_content in enumerate(row)
                if cell_content.strip()
            ],
            "boundingRegions": [{"pageNumber": 1}],
            "spans": [{"offset": start_offset, "length": len(table_content)}]
        })
        
        full_content += table_content + "\n\n"
        table_id += 1
    
    if output_format.lower() == "markdown":
        markdown_content = ""
        for element in content_elements:
            if element["kind"] == "title":
                level = int(element["role"].replace("heading", ""))
                markdown_content += "#" * level + " " + element["content"] + "\n\n"
            elif element["kind"] == "paragraph":
                markdown_content += element["content"] + "\n\n"
            elif element["kind"] == "table":
                # Convert to markdown table
                table_rows = []
                max_cols = element.get("columnCount", 0)
                
                for cell in element["cells"]:
                    row_idx = cell["rowIndex"]
                    col_idx = cell["columnIndex"]
                    while row_idx >= len(table_rows):
                        table_rows.append([""] * max_cols)
                    if col_idx < max_cols:
                        table_rows[row_idx][col_idx] = cell["content"]
                
                if table_rows:
                    markdown_content += "| " + " | ".join(table_rows[0]) + " |\n"
                    markdown_content += "|" + "---|" * len(table_rows[0]) + "\n"
                    for row in table_rows[1:]:
                        markdown_content += "| " + " | ".join(row) + " |\n"
                    markdown_content += "\n"
        
        return markdown_content.strip()
    
    else:
        azure_response = {
            "apiVersion": "2024-11-30",
            "modelId": "prebuilt-layout",
            "content": full_content.strip(),
            "pages": [{"pageNumber": 1, "width": 8.5, "height": 11, "unit": "inch"}],
            "paragraphs": [e for e in content_elements if e["kind"] in ["paragraph", "title"]],
            "tables": [e for e in content_elements if e["kind"] == "table"],
        }
        
        return json.dumps(azure_response, indent=2)

def process_docx(file_bytes: bytes, filename: str) -> tuple[str, int]:
    """Process DOCX files using python-docx."""
    if not DOCX_AVAILABLE:
        raise HTTPException(status_code=500, detail="python-docx not available")
    
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_file:
            temp_file.write(file_bytes)
            temp_file_path = temp_file.name
        
        doc = DocxDocument(temp_file_path)
        structured_content = extract_docx_structure_azure_format(doc, filename, "json")
        
        os.unlink(temp_file_path)
        
        page_count = max(1, len(structured_content) // 3000)
        
        return structured_content, page_count
        
    except Exception as e:
        if 'temp_file_path' in locals():
            try:
                os.unlink(temp_file_path)
            except:
                pass
        raise HTTPException(status_code=500, detail=f"Failed to process DOCX: {e}")

def process_xlsx(file_bytes: bytes, filename: str) -> tuple[str, int]:
    """Process XLSX/XLSM files using openpyxl with enhanced extraction."""
    if not OPENPYXL_AVAILABLE:
        raise HTTPException(status_code=500, detail="openpyxl not available")
    
    try:
        file_ext = os.path.splitext(filename)[1].lower() or '.xlsx'
        if file_ext not in ['.xlsx', '.xlsm', '.xls']:
            file_ext = '.xlsx'
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as temp_file:
            temp_file.write(file_bytes)
            temp_file_path = temp_file.name
        
        workbook = openpyxl.load_workbook(temp_file_path, read_only=False, data_only=False)
        
        full_text = ""
        sheet_count = 0
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            sheet_count += 1
            
            full_text += f"\n=== Sheet: {sheet_name} ===\n\n"
            
            row_num = 0
            for row in sheet.iter_rows():
                row_num += 1
                row_values = []
                has_content = False
                
                for cell in row:
                    if cell.value is not None:
                        has_content = True
                        if isinstance(cell.value, str) and cell.value.startswith('='):
                            cell_str = f"{normalize_text_encoding(cell.value)} (formula)"
                        else:
                            cell_str = normalize_text_encoding(str(cell.value))
                        row_values.append(cell_str)
                    else:
                        row_values.append("")
                
                if has_content:
                    while row_values and row_values[-1] == "":
                        row_values.pop()
                    if row_values:
                        full_text += f"[R{row_num}] " + "\t".join(row_values) + "\n"
            
            full_text += "\n"
        
        workbook.close()
        os.unlink(temp_file_path)
        
        return full_text.strip(), sheet_count
        
    except Exception as e:
        if 'temp_file_path' in locals():
            try:
                os.unlink(temp_file_path)
            except:
                pass
        raise HTTPException(status_code=500, detail=f"Failed to process XLSX: {e}")

def process_pptx(file_bytes: bytes, filename: str) -> tuple[str, int]:
    """Process PPTX files using python-pptx."""
    if not PPTX_AVAILABLE:
        raise HTTPException(status_code=500, detail="python-pptx not available")
    
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as temp_file:
            temp_file.write(file_bytes)
            temp_file_path = temp_file.name
        
        prs = Presentation(temp_file_path)
        full_text = ""
        slide_count = 0
        
        for slide in prs.slides:
            slide_count += 1
            full_text += f"\n=== Slide {slide_count} ===\n"
            
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    shape_text = normalize_text_encoding(shape.text)
                    full_text += shape_text + "\n"
                
                if shape.has_table:
                    table = shape.table
                    for row in table.rows:
                        for cell in row.cells:
                            cell_text = normalize_text_encoding(cell.text)
                            full_text += cell_text + " "
                        full_text += "\n"
        
        os.unlink(temp_file_path)
        
        return full_text.strip(), slide_count
        
    except Exception as e:
        if 'temp_file_path' in locals():
            try:
                os.unlink(temp_file_path)
            except:
                pass
        raise HTTPException(status_code=500, detail=f"Failed to process PPTX: {e}")

def process_rtf(file_bytes: bytes, filename: str) -> tuple[str, int]:
    """Process RTF files using striprtf."""
    if not RTF_AVAILABLE:
        raise HTTPException(status_code=500, detail="striprtf not available")
    
    try:
        rtf_content = file_bytes.decode('utf-8', errors='ignore')
        plain_text = rtf_to_text(rtf_content)
        plain_text = normalize_text_encoding(plain_text)
        page_count = max(1, len(plain_text) // 3000)
        
        return plain_text.strip(), page_count
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to process RTF: {e}")

def process_non_pdf_fast(file_bytes: bytes, filename: str, file_ext: str) -> tuple[str, int]:
    """Process non-PDF documents using fast native Python processors."""
    logger.info(f"Processing {file_ext} file: {filename}")
    
    if file_ext == ".docx":
        return process_docx(file_bytes, filename)
    elif file_ext in [".xlsx", ".xlsm"]:
        return process_xlsx(file_bytes, filename)
    elif file_ext == ".pptx":
        return process_pptx(file_bytes, filename)
    elif file_ext == ".rtf":
        return process_rtf(file_bytes, filename)
    else:
        raise HTTPException(
            status_code=400, 
            detail=f"Unsupported file type: {file_ext}"
        )

# =============================================================================
# API ENDPOINTS
# =============================================================================

@app.put("/process")
async def process_document(request: Request):
    """
    Processes an uploaded document to extract text, tables, layout, and images.
    
    Headers:
        X-Filename: Original filename
        X-Extract-Images: "true" to extract images
        X-Extract-Tables: "true" (default) to extract tables
        X-Detect-Layout: "true" (default) to detect multi-column layouts
        X-Extract-Metadata: "true" (default) to extract document metadata
        X-Exclude-Headers-Footers: "true" to exclude headers/footers from text
        outputContentFormat: "json" (default) or "markdown"
    """
    start_time = time.time()
    
    file_bytes = await request.body()
    if not file_bytes:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="No file content received.")

    filename = request.headers.get("x-filename", "unknown_file")
    file_ext = os.path.splitext(filename)[1].lower()
    
    # Extract processing options from headers
    extract_images_flag = request.headers.get("x-extract-images", "false").lower()
    extract_tables = request.headers.get("x-extract-tables", "true").lower() == "true"
    detect_layout = request.headers.get("x-detect-layout", "true").lower() == "true"
    extract_metadata = request.headers.get("x-extract-metadata", "true").lower() == "true"
    exclude_headers_footers = request.headers.get("x-exclude-headers-footers", "false").lower() == "true"
    output_format = request.headers.get("outputContentFormat", DEFAULT_OUTPUT_FORMAT).lower()
    
    logger.info(f"Processing: {filename}, ext={file_ext}, tables={extract_tables}, layout={detect_layout}")

    # PDF processing with enhanced features
    if file_ext == ".pdf":
        result = process_pdf_enhanced(
            file_bytes=file_bytes,
            filename=filename,
            extract_images_flag=extract_images_flag,
            extract_tables=extract_tables,
            detect_layout=detect_layout,
            extract_metadata=extract_metadata,
            exclude_headers_footers=exclude_headers_footers,
            output_format=output_format
        )
        
        return JSONResponse(content=result)
    
    # Non-PDF processing
    else:
        try:
            if file_ext == ".docx" and DOCX_AVAILABLE:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_file:
                    temp_file.write(file_bytes)
                    temp_file_path = temp_file.name
                
                doc = DocxDocument(temp_file_path)
                full_text = extract_docx_structure_azure_format(doc, filename, output_format)
                page_count = max(1, len(full_text) // 3000)
                os.unlink(temp_file_path)
            else:
                full_text, page_count = process_non_pdf_fast(file_bytes, filename, file_ext)
            
            response_payload = {
                "page_content": full_text.strip(),
                "metadata": {
                    "source": filename,
                    "page_count": page_count,
                    "images_extracted": 0,
                    "processing_status": "completed",
                    "processing_time_ms": int((time.time() - start_time) * 1000),
                    "output_format": output_format,
                },
                "images": [],
                "tables": [],
            }
            
            # Write debug output if enabled
            write_debug_output(filename, response_payload["page_content"], response_payload)
            
            return JSONResponse(content=response_payload)
            
        except HTTPException:
            raise
        except Exception as e:
            logger.error(f"Error processing {file_ext} document: {e}")
            raise HTTPException(status_code=500, detail=f"Failed to process document: {e}")

@app.get("/")
def read_root():
    """Root endpoint with API information."""
    return {
        "status": "ok",
        "message": "Enhanced Content Processing Engine is running.",
        "version": "2.0.0",
        "features": {
            "pdf_processing": "PyMuPDF with enhanced extraction",
            "table_extraction": "Structured tables with markdown output",
            "layout_detection": "Multi-column and header/footer detection",
            "metadata_extraction": "Full document metadata and TOC",
            "image_processing": "Filtered and deduplicated with compression",
        },
        "supported_formats": ["PDF", "DOCX", "XLSX", "PPTX", "RTF"],
        "processors_available": {
            "docx": DOCX_AVAILABLE,
            "xlsx": OPENPYXL_AVAILABLE,
            "pptx": PPTX_AVAILABLE,
            "rtf": RTF_AVAILABLE,
            "pymupdf4llm": PYMUPDF4LLM_AVAILABLE,
        },
        "settings": {
            "multi_column_detection": ENABLE_MULTI_COLUMN_DETECTION,
            "header_footer_detection": ENABLE_HEADER_FOOTER_DETECTION,
            "image_compression": AUTO_COMPRESS_IMAGES,
        }
    }

@app.get("/health")
def health_check():
    """Health check endpoint."""
    return {
        "status": "healthy",
        "version": "2.0.0",
        "processors": {
            "pdf": "PyMuPDF - Available (Enhanced)",
            "docx": f"python-docx - {'Available' if DOCX_AVAILABLE else 'Not Available'}",
            "xlsx": f"openpyxl - {'Available' if OPENPYXL_AVAILABLE else 'Not Available'}",
            "pptx": f"python-pptx - {'Available' if PPTX_AVAILABLE else 'Not Available'}",
            "rtf": f"striprtf - {'Available' if RTF_AVAILABLE else 'Not Available'}",
            "pymupdf4llm": f"{'Available' if PYMUPDF4LLM_AVAILABLE else 'Not Available'}",
        },
        "features": {
            "table_extraction": True,
            "layout_detection": ENABLE_MULTI_COLUMN_DETECTION,
            "header_footer_detection": ENABLE_HEADER_FOOTER_DETECTION,
            "image_compression": AUTO_COMPRESS_IMAGES,
        }
    }

@app.get("/capabilities")
def get_capabilities():
    """Return detailed capability information."""
    return {
        "pdf_features": {
            "text_extraction": "Page-by-page with unicode normalization",
            "table_extraction": "PyMuPDF find_tables() with markdown output",
            "layout_detection": {
                "multi_column": ENABLE_MULTI_COLUMN_DETECTION,
                "header_detection": ENABLE_HEADER_FOOTER_DETECTION,
                "footer_detection": ENABLE_HEADER_FOOTER_DETECTION,
                "header_margin": f"{HEADER_MARGIN_RATIO * 100}% of page height",
                "footer_margin": f"{FOOTER_MARGIN_RATIO * 100}% of page height",
            },
            "metadata_extraction": [
                "title", "author", "subject", "keywords", 
                "creator", "producer", "creation_date", "modification_date",
                "table_of_contents"
            ],
            "image_extraction": {
                "enabled": True,
                "compression": AUTO_COMPRESS_IMAGES,
                "deduplication": True,
                "min_dimensions": f"{MIN_IMAGE_WIDTH}x{MIN_IMAGE_HEIGHT}",
                "min_size": f"{MIN_IMAGE_SIZE_BYTES} bytes",
            }
        },
        "output_formats": ["json", "markdown"],
        "request_headers": {
            "X-Filename": "Original filename (required)",
            "X-Extract-Images": "true/false (default: false)",
            "X-Extract-Tables": "true/false (default: true)",
            "X-Detect-Layout": "true/false (default: true)",
            "X-Extract-Metadata": "true/false (default: true)",
            "X-Exclude-Headers-Footers": "true/false (default: false)",
            "outputContentFormat": "json/markdown (default: json)",
        }
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)