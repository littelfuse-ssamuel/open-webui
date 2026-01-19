"""
PowerPoint file generation and management API endpoints.
Provides endpoints for generating .pptx files from JSON slide data.

Uses the existing Storage provider for file persistence (supports Azure, S3, GCS, local).
"""

import os
import uuid
import logging
import io
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Any, Optional

from fastapi import APIRouter, Depends, HTTPException, Request, status, BackgroundTasks, UploadFile
from fastapi.responses import Response
from pydantic import BaseModel, Field

from open_webui.constants import ERROR_MESSAGES
from open_webui.utils.auth import get_verified_user
from open_webui.utils.pptx_generator import generate_pptx_from_data, get_presentation_info
from open_webui.storage.provider import Storage
from open_webui.models.files import Files, FileForm, FileModel
from open_webui.routers.files import upload_file_handler
from open_webui.config import DATA_DIR, UPLOAD_DIR
from open_webui.env import SRC_LOG_LEVELS

log = logging.getLogger(__name__)
log.setLevel(SRC_LOG_LEVELS["MAIN"])

router = APIRouter()

# ================================
# Configuration
# ================================

# Template directory within DATA_DIR (persistent storage)
PPTX_TEMPLATE_DIR = DATA_DIR / "templates"
PPTX_TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)

# Default template filename
DEFAULT_TEMPLATE_FILENAME = os.environ.get("PPTX_TEMPLATE_FILENAME", "littelfuse_template.pptx")

# Full path to template
PPTX_TEMPLATE_PATH = PPTX_TEMPLATE_DIR / DEFAULT_TEMPLATE_FILENAME


# ================================
# Pydantic Models
# ================================

class PptxContentItem(BaseModel):
    """A single content item on a slide."""
    type: str = Field(..., description="Content type: 'text', 'bullet', 'table', or 'image'")
    text: Optional[str] = Field(None, description="Text content (for type='text')")
    items: Optional[List[str]] = Field(None, description="Bullet items (for type='bullet')")
    headers: Optional[List[str]] = Field(None, description="Table headers (for type='table')")
    rows: Optional[List[List[Any]]] = Field(None, description="Table rows (for type='table')")
    src: Optional[str] = Field(None, description="Image source - base64 or URL (for type='image')")
    alt: Optional[str] = Field(None, description="Image alt text (for type='image')")


class PptxSlide(BaseModel):
    """A single slide in the presentation."""
    title: Optional[str] = Field(None, description="Slide title")
    backgroundColor: Optional[str] = Field(None, description="Background color in hex format")
    content: Optional[List[PptxContentItem]] = Field(default_factory=list, description="Slide content items")
    notes: Optional[str] = Field(None, description="Speaker notes")


class PptxGenerateRequest(BaseModel):
    """Request body for generating a PPTX file."""
    title: str = Field(..., description="Presentation title")
    slides: List[PptxSlide] = Field(..., description="List of slides")
    use_template: Optional[bool] = Field(True, description="Whether to use the company template")


class PptxGenerateResponse(BaseModel):
    """Response from PPTX generation."""
    success: bool
    file_id: str
    download_url: str
    slide_count: int
    filename: str
    message: Optional[str] = None


class PptxTemplateInfoResponse(BaseModel):
    """Response with template information."""
    available: bool
    path: Optional[str] = None
    filename: Optional[str] = None
    slide_count: Optional[int] = None
    layouts: Optional[List[Dict[str, Any]]] = None
    error: Optional[str] = None


class PptxFileArtifact(BaseModel):
    """PPTX file artifact structure (matches Excel artifact pattern)."""
    type: str = "pptx"
    url: str
    name: str
    fileId: str
    meta: Optional[Dict[str, Any]] = None


# ================================
# Helper Functions
# ================================

def get_template_path() -> Optional[str]:
    """Get the path to the company template if it exists."""
    if PPTX_TEMPLATE_PATH.exists():
        return str(PPTX_TEMPLATE_PATH)
    
    # Also check for template in UPLOAD_DIR (legacy support)
    legacy_path = UPLOAD_DIR / DEFAULT_TEMPLATE_FILENAME
    if legacy_path.exists():
        return str(legacy_path)
    
    return None


def create_pptx_file_record(
    request: Request,
    pptx_bytes: bytes,
    filename: str,
    user,
    metadata: Optional[Dict] = None
) -> Optional[Dict]:
    """
    Create a file record for a generated PPTX file using the standard upload handler.
    
    This follows the same pattern as Excel artifact creation.
    
    Args:
        request: FastAPI request object
        pptx_bytes: Generated PPTX file as bytes
        filename: Filename for the PPTX file
        user: User object
        metadata: Optional metadata dict
        
    Returns:
        Dict with file artifact structure or None if failed
    """
    try:
        # Create UploadFile object from bytes
        file = UploadFile(
            file=io.BytesIO(pptx_bytes),
            filename=filename,
            headers={
                "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            },
        )
        
        # Build metadata
        if metadata is None:
            metadata = {}
        
        metadata.update({
            "name": filename,
            "content_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "generated": True,
            "generator": "pptx_artifact",
        })
        
        # Upload using standard handler (this uses the Storage provider)
        file_item = upload_file_handler(
            request,
            file=file,
            metadata=metadata,
            process=False,  # Don't process for RAG
            user=user,
        )
        
        if not file_item:
            log.error("upload_file_handler returned None")
            return None
        
        # Build file artifact dict (matches Excel artifact structure)
        url = request.app.url_path_for("get_file_content_by_id", id=file_item.id)
        
        artifact = {
            "type": "pptx",
            "url": url,
            "name": filename,
            "fileId": str(file_item.id),
            "meta": metadata,
        }
        
        log.info(f"Created PPTX file artifact: {filename} -> {file_item.id}")
        return artifact
        
    except Exception as e:
        log.error(f"Error creating PPTX file record: {e}")
        return None


# ================================
# API Endpoints
# ================================

@router.post("/generate", response_model=PptxGenerateResponse)
async def generate_pptx(
    request: Request,
    form_data: PptxGenerateRequest,
    user=Depends(get_verified_user)
):
    """
    Generate a PPTX file from structured slide data.
    
    The slide data should include a title and array of slides, where each slide
    can contain text, bullet points, tables, and images.
    
    The generated file is stored using the configured storage provider (Azure, S3, etc.)
    and can be downloaded via the standard file content endpoint.
    """
    try:
        # Convert request to dict for generator
        slide_data = {
            'title': form_data.title,
            'slides': [slide.model_dump() for slide in form_data.slides]
        }
        
        # Determine template path
        template_path = None
        if form_data.use_template:
            template_path = get_template_path()
            if template_path:
                log.info(f"Using template: {template_path}")
            else:
                log.warning("Template requested but not found, using blank presentation")
        
        # Generate PPTX bytes
        pptx_bytes = generate_pptx_from_data(slide_data, template_path)
        
        # Create filename
        safe_title = "".join(c for c in form_data.title if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_title = safe_title[:50] if safe_title else "presentation"
        filename = f"{safe_title}.pptx"
        
        # Create file record using storage provider
        artifact = create_pptx_file_record(
            request=request,
            pptx_bytes=pptx_bytes,
            filename=filename,
            user=user,
            metadata={
                "slide_count": len(form_data.slides),
                "title": form_data.title,
                "used_template": template_path is not None,
            }
        )
        
        if not artifact:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail="Failed to store generated presentation"
            )
        
        log.info(f"Generated PPTX: {filename} ({len(pptx_bytes)} bytes) -> {artifact['fileId']}")
        
        return PptxGenerateResponse(
            success=True,
            file_id=artifact['fileId'],
            download_url=artifact['url'],
            slide_count=len(form_data.slides),
            filename=filename,
            message="Presentation generated successfully"
        )
        
    except HTTPException:
        raise
    except Exception as e:
        log.error(f"Error generating PPTX: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to generate presentation: {str(e)}"
        )


@router.get("/template/info", response_model=PptxTemplateInfoResponse)
async def get_template_info(user=Depends(get_verified_user)):
    """
    Get information about the configured company template.
    
    Returns template availability, slide count, and available layouts.
    """
    template_path = get_template_path()
    
    if not template_path:
        return PptxTemplateInfoResponse(
            available=False,
            path=str(PPTX_TEMPLATE_DIR),
            filename=DEFAULT_TEMPLATE_FILENAME,
            error=f"Template not found. Please upload '{DEFAULT_TEMPLATE_FILENAME}' to {PPTX_TEMPLATE_DIR}"
        )
    
    try:
        info = get_presentation_info(template_path)
        
        if 'error' in info:
            return PptxTemplateInfoResponse(
                available=False,
                path=template_path,
                filename=DEFAULT_TEMPLATE_FILENAME,
                error=info['error']
            )
        
        return PptxTemplateInfoResponse(
            available=True,
            path=template_path,
            filename=DEFAULT_TEMPLATE_FILENAME,
            slide_count=info.get('slide_count', 0),
            layouts=info.get('layouts', [])
        )
        
    except Exception as e:
        log.error(f"Error reading template info: {e}")
        return PptxTemplateInfoResponse(
            available=False,
            path=template_path,
            filename=DEFAULT_TEMPLATE_FILENAME,
            error=str(e)
        )


@router.post("/template/upload")
async def upload_template(
    request: Request,
    file: UploadFile,
    user=Depends(get_verified_user)
):
    """
    Upload a new company template.
    
    The template will be saved to the templates directory and used for
    future presentation generation.
    """
    # Verify file is a PPTX
    if not file.filename.endswith('.pptx'):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File must be a .pptx PowerPoint file"
        )
    
    try:
        # Read file contents
        contents = await file.read()
        
        if not contents:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Empty file"
            )
        
        # Ensure template directory exists
        PPTX_TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
        
        # Save template
        template_path = PPTX_TEMPLATE_DIR / DEFAULT_TEMPLATE_FILENAME
        with open(template_path, 'wb') as f:
            f.write(contents)
        
        # Get template info
        info = get_presentation_info(str(template_path))
        
        log.info(f"Uploaded template: {DEFAULT_TEMPLATE_FILENAME} ({len(contents)} bytes)")
        
        return {
            "success": True,
            "message": f"Template uploaded successfully",
            "path": str(template_path),
            "size_bytes": len(contents),
            "slide_count": info.get('slide_count', 0),
            "layouts": info.get('layouts', [])
        }
        
    except HTTPException:
        raise
    except Exception as e:
        log.error(f"Error uploading template: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to upload template: {str(e)}"
        )


@router.get("/files", response_model=List[Dict])
async def list_pptx_files(
    request: Request,
    user=Depends(get_verified_user),
    limit: int = 20
):
    """
    List recent PPTX files generated by the current user.
    """
    try:
        # Get user's files and filter for PPTX
        all_files = Files.get_files_by_user_id(user.id)
        
        pptx_files = []
        for file in all_files:
            # Check if it's a PPTX file by content type or extension
            is_pptx = (
                file.filename.endswith('.pptx') or
                (file.meta and file.meta.get('content_type') == 
                 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
            )
            
            if is_pptx:
                url = request.app.url_path_for("get_file_content_by_id", id=file.id)
                pptx_files.append({
                    "id": file.id,
                    "filename": file.filename,
                    "url": url,
                    "created_at": file.created_at,
                    "meta": file.meta
                })
        
        # Sort by created_at descending and limit
        pptx_files.sort(key=lambda x: x['created_at'], reverse=True)
        return pptx_files[:limit]
        
    except Exception as e:
        log.error(f"Error listing PPTX files: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to list files: {str(e)}"
        )


@router.delete("/files/{file_id}")
async def delete_pptx_file(
    file_id: str,
    user=Depends(get_verified_user)
):
    """
    Delete a generated PPTX file.
    """
    try:
        # Verify file exists and belongs to user
        file = Files.get_file_by_id(file_id)
        
        if not file:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found"
            )
        
        if file.user_id != user.id:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Not authorized to delete this file"
            )
        
        # Delete from storage
        try:
            Storage.delete_file(file.path)
        except Exception as e:
            log.warning(f"Error deleting file from storage: {e}")
        
        # Delete from database
        Files.delete_file_by_id(file_id)
        
        return {"success": True, "message": "File deleted"}
        
    except HTTPException:
        raise
    except Exception as e:
        log.error(f"Error deleting PPTX file: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to delete file: {str(e)}"
        )