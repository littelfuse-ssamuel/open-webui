"""
Excel file editing API endpoints
"""

import logging
from pathlib import Path
from typing import List, Optional
from fastapi import APIRouter, Depends, HTTPException, status
from pydantic import BaseModel
import openpyxl
from openpyxl.utils import get_column_letter

from open_webui.constants import ERROR_MESSAGES
from open_webui.models.files import Files, FileModel
from open_webui.storage.provider import Storage
from open_webui.utils.auth import get_verified_user

log = logging.getLogger(__name__)

router = APIRouter()


class CellChange(BaseModel):
    """Represents a single cell change"""
    row: int  # 1-based row number
    col: int  # 1-based column number
    value: Optional[str | int | float | bool] = None
    isFormula: Optional[bool] = False  # Whether the value is a formula (starts with =)


class ExcelUpdateRequest(BaseModel):
    """Request to update cells in an Excel file"""
    fileId: str
    sheet: str
    changes: List[CellChange]


class ExcelUpdateResponse(BaseModel):
    """Response from Excel update operation"""
    status: str
    message: Optional[str] = None


class ExcelMetadataResponse(BaseModel):
    """Response with Excel file metadata"""
    fileId: str
    sheetNames: List[str]
    activeSheet: Optional[str] = None


@router.post("/update")
async def update_excel_file(
    request: ExcelUpdateRequest,
    user=Depends(get_verified_user),
) -> ExcelUpdateResponse:
    """
    Update cells in an Excel workbook.

    Args:
        request: The update request with file ID, sheet name, and cell changes
        user: The authenticated user

    Returns:
        ExcelUpdateResponse with status

    Raises:
        HTTPException: If file not found, access denied, or update fails
    """
    try:
        # Get the file from database
        file = Files.get_file_by_id(request.fileId)

        if not file:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=ERROR_MESSAGES.NOT_FOUND,
            )

        # Check user permissions
        if file.user_id != user.id and user.role != "admin":
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail=ERROR_MESSAGES.ACCESS_PROHIBITED,
            )

        # Get the file path from storage
        file_path = Storage.get_file(file.path)
        file_path = Path(file_path)

        if not file_path.is_file():
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found in storage",
            )

        # Load the workbook
        try:
            wb = openpyxl.load_workbook(file_path, keep_vba=True)
        except Exception as e:
            log.error(f"Error loading workbook {request.fileId}: {e}")
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Invalid Excel file: {str(e)}",
            )

        # Check if sheet exists
        if request.sheet not in wb.sheetnames:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Sheet '{request.sheet}' not found in workbook. Available sheets: {', '.join(wb.sheetnames)}",
            )

        # Get the worksheet
        ws = wb[request.sheet]

        # Apply changes
        changes_applied = 0
        for change in request.changes:
            try:
                # Validate row and column are positive
                if change.row < 1 or change.col < 1:
                    log.warning(f"Invalid cell coordinates: row={change.row}, col={change.col}")
                    continue

                # Get the cell and set its value
                cell = ws.cell(row=change.row, column=change.col)

                # Handle formulas vs regular values
                if change.isFormula and isinstance(change.value, str) and change.value.startswith('='):
                    # Set as formula - openpyxl will evaluate it
                    cell.value = change.value
                    log.debug(f"Updated cell {get_column_letter(change.col)}{change.row} with formula: {change.value}")
                else:
                    # Set as regular value
                    # Try to convert numeric strings to numbers for proper Excel formatting
                    value = change.value
                    if isinstance(value, str):
                        try:
                            if '.' in value:
                                value = float(value)
                            else:
                                value = int(value)
                        except ValueError:
                            pass  # Keep as string
                    cell.value = value
                    log.debug(f"Updated cell {get_column_letter(change.col)}{change.row} = {value}")

                changes_applied += 1

            except Exception as e:
                log.error(f"Error updating cell at row={change.row}, col={change.col}: {e}")
                continue

        # Save the workbook
        try:
            wb.save(file_path)
            wb.close()
            log.info(f"Successfully updated {changes_applied} cells in file {request.fileId}")
        except Exception as e:
            log.error(f"Error saving workbook {request.fileId}: {e}")
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Failed to save changes: {str(e)}",
            )

        return ExcelUpdateResponse(
            status="ok",
            message=f"Successfully updated {changes_applied} cells",
        )

    except HTTPException:
        raise
    except Exception as e:
        log.error(f"Unexpected error updating Excel file: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Internal server error: {str(e)}",
        )


@router.get("/{file_id}/metadata")
async def get_excel_metadata(
    file_id: str,
    user=Depends(get_verified_user),
) -> ExcelMetadataResponse:
    """
    Get metadata about an Excel file (sheet names, etc.).

    Args:
        file_id: The ID of the Excel file
        user: The authenticated user

    Returns:
        ExcelMetadataResponse with sheet names and active sheet

    Raises:
        HTTPException: If file not found or access denied
    """
    try:
        # Get the file from database
        file = Files.get_file_by_id(file_id)

        if not file:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=ERROR_MESSAGES.NOT_FOUND,
            )

        # Check user permissions
        if file.user_id != user.id and user.role != "admin":
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail=ERROR_MESSAGES.ACCESS_PROHIBITED,
            )

        # Get the file path from storage
        file_path = Storage.get_file(file.path)
        file_path = Path(file_path)

        if not file_path.is_file():
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found in storage",
            )

        # Load the workbook (data_only=True for faster loading)
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            sheet_names = wb.sheetnames
            active_sheet = wb.active.title if wb.active else None
            wb.close()

            return ExcelMetadataResponse(
                fileId=file_id,
                sheetNames=sheet_names,
                activeSheet=active_sheet,
            )

        except Exception as e:
            log.error(f"Error reading workbook {file_id}: {e}")
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Invalid Excel file: {str(e)}",
            )

    except HTTPException:
        raise
    except Exception as e:
        log.error(f"Unexpected error getting Excel metadata: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Internal server error: {str(e)}",
        )

class CellReference(BaseModel):
    """A cell reference with sheet, row, col"""
    sheet: str
    row: int
    col: int
    address: str  # e.g., "A1"


class CellWarning(BaseModel):
    """Warning about a cell that may be affected by changes"""
    cell: CellReference
    warning_type: str  # "contains_formula", "referenced_by_formula", "in_named_range"
    message: str
    details: Optional[dict] = None


class ExcelValidationRequest(BaseModel):
    """Request to validate proposed changes before applying"""
    fileId: str
    sheet: str
    changes: List[CellChange]


class ExcelValidationResponse(BaseModel):
    """Response with validation warnings"""
    status: str
    warnings: List[CellWarning]
    safe_to_apply: bool
    message: Optional[str] = None


def _parse_cell_references_from_formula(formula: str) -> List[str]:
    """
    Extract cell references from a formula string.
    Returns list of cell addresses like ['A1', 'B2:B10', 'Sheet2!C3']
    """
    import re
    # Match cell references: A1, $A$1, A1:B10, Sheet1!A1, 'Sheet Name'!A1
    pattern = r"(?:'[^']+'!|\w+!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?"
    matches = re.findall(pattern, formula.upper())
    return matches


def _expand_range(range_str: str) -> List[str]:
    """
    Expand a range like 'A1:A10' into individual cell addresses.
    Returns list of cell addresses.
    """
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
    
    if ':' not in range_str:
        return [range_str]
    
    # Remove sheet reference if present
    if '!' in range_str:
        range_str = range_str.split('!')[-1]
    
    # Remove $ signs
    range_str = range_str.replace('$', '')
    
    try:
        start, end = range_str.split(':')
        start_col, start_row = coordinate_from_string(start)
        end_col, end_row = coordinate_from_string(end)
        
        start_col_idx = column_index_from_string(start_col)
        end_col_idx = column_index_from_string(end_col)
        
        cells = []
        for row in range(start_row, end_row + 1):
            for col in range(start_col_idx, end_col_idx + 1):
                cells.append(f"{get_column_letter(col)}{row}")
        return cells
    except Exception:
        return [range_str]


@router.post("/validate")
async def validate_excel_changes(
    request: ExcelValidationRequest,
    user=Depends(get_verified_user),
) -> ExcelValidationResponse:
    """
    Validate proposed changes before applying them.
    
    Returns warnings about:
    - Cells that contain formulas (will be overwritten)
    - Cells that are referenced by other formulas (may break calculations)
    - Cells within named ranges
    
    This endpoint is advisory - changes can still be applied even with warnings.
    """
    try:
        # Get the file from database
        file = Files.get_file_by_id(request.fileId)

        if not file:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=ERROR_MESSAGES.NOT_FOUND,
            )

        # Check user permissions
        if file.user_id != user.id and user.role != "admin":
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail=ERROR_MESSAGES.ACCESS_PROHIBITED,
            )

        # Get the file path from storage
        file_path = Storage.get_file(file.path)
        file_path = Path(file_path)

        if not file_path.is_file():
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found in storage",
            )

        # Load the workbook (need formulas, not values)
        try:
            wb = openpyxl.load_workbook(file_path, data_only=False)
        except Exception as e:
            log.error(f"Error loading workbook for validation {request.fileId}: {e}")
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Invalid Excel file: {str(e)}",
            )

        # Check if sheet exists
        if request.sheet not in wb.sheetnames:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Sheet '{request.sheet}' not found. Available: {', '.join(wb.sheetnames)}",
            )

        ws = wb[request.sheet]
        warnings: List[CellWarning] = []
        
        # Build set of cells being changed for quick lookup
        changing_cells = set()
        for change in request.changes:
            addr = f"{get_column_letter(change.col)}{change.row}"
            changing_cells.add(addr)

        # Check 1: Are any of the changing cells formulas?
        for change in request.changes:
            cell = ws.cell(row=change.row, column=change.col)
            addr = f"{get_column_letter(change.col)}{change.row}"
            
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                warnings.append(CellWarning(
                    cell=CellReference(
                        sheet=request.sheet,
                        row=change.row,
                        col=change.col,
                        address=addr
                    ),
                    warning_type="contains_formula",
                    message=f"Cell {addr} contains a formula that will be overwritten",
                    details={"current_formula": cell.value}
                ))

        # Check 2: Are any changing cells referenced by formulas elsewhere?
        # Scan all cells in the sheet for formulas that reference changing cells
        referenced_by: dict[str, List[str]] = {addr: [] for addr in changing_cells}
        
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    formula = cell.value
                    refs = _parse_cell_references_from_formula(formula)
                    
                    for ref in refs:
                        # Expand ranges and check each cell
                        expanded = _expand_range(ref)
                        for expanded_addr in expanded:
                            # Normalize address (remove $ and sheet refs)
                            clean_addr = expanded_addr.replace('$', '')
                            if '!' in clean_addr:
                                clean_addr = clean_addr.split('!')[-1]
                            
                            if clean_addr in changing_cells:
                                source_addr = f"{get_column_letter(cell.column)}{cell.row}"
                                referenced_by[clean_addr].append(source_addr)

        for addr, referencing_cells in referenced_by.items():
            if referencing_cells:
                # Find the change object for this address
                for change in request.changes:
                    if f"{get_column_letter(change.col)}{change.row}" == addr:
                        warnings.append(CellWarning(
                            cell=CellReference(
                                sheet=request.sheet,
                                row=change.row,
                                col=change.col,
                                address=addr
                            ),
                            warning_type="referenced_by_formula",
                            message=f"Cell {addr} is referenced by {len(referencing_cells)} formula(s)",
                            details={"referenced_by": referencing_cells[:10]}  # Limit to 10
                        ))
                        break

        # Check 3: Named ranges (check workbook-level defined names)
        try:
            for defined_name in wb.defined_names.definedName:
                if defined_name.value:
                    # Parse the range
                    refs = _parse_cell_references_from_formula(defined_name.value)
                    for ref in refs:
                        expanded = _expand_range(ref)
                        for expanded_addr in expanded:
                            clean_addr = expanded_addr.replace('$', '')
                            if '!' in clean_addr:
                                clean_addr = clean_addr.split('!')[-1]
                            
                            if clean_addr in changing_cells:
                                for change in request.changes:
                                    if f"{get_column_letter(change.col)}{change.row}" == clean_addr:
                                        warnings.append(CellWarning(
                                            cell=CellReference(
                                                sheet=request.sheet,
                                                row=change.row,
                                                col=change.col,
                                                address=clean_addr
                                            ),
                                            warning_type="in_named_range",
                                            message=f"Cell {clean_addr} is part of named range '{defined_name.name}'",
                                            details={"named_range": defined_name.name}
                                        ))
                                        break
        except Exception as e:
            log.warning(f"Could not check named ranges: {e}")

        wb.close()

        # Determine if safe to apply
        has_critical_warnings = any(
            w.warning_type == "referenced_by_formula" for w in warnings
        )

        return ExcelValidationResponse(
            status="ok",
            warnings=warnings,
            safe_to_apply=not has_critical_warnings,
            message=f"Found {len(warnings)} warning(s)" if warnings else "No warnings - safe to apply"
        )

    except HTTPException:
        raise
    except Exception as e:
        log.error(f"Unexpected error validating Excel changes: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Internal server error: {str(e)}",
        )