"""
Regression test suite for Excel artifact emission and file handling.

Phase 1 - Stabilize & Baseline:
These tests ensure the event contract is preserved and no breaking changes
are introduced to Excel artifact flows.

Event Contract (FROZEN):
- Event type: "files"
- File type: "excel"
- Required: url, name
- Optional: fileId, meta.sheetNames, meta.activeSheet
"""

import pytest
import asyncio
from unittest.mock import MagicMock, AsyncMock, patch
from pathlib import Path
import tempfile
import os
import uuid

# Import the modules under test
from open_webui.utils.artifacts import (
    emit_excel_artifact,
    emit_file_artifacts,
    create_excel_file_record,
)


class TestExcelArtifactEventContract:
    """Tests to verify the event contract is preserved."""

    @pytest.mark.asyncio
    async def test_emit_excel_artifact_event_structure(self):
        """Verify emit_excel_artifact produces correct event structure."""
        # Setup
        mock_emitter = AsyncMock()
        mock_file_model = MagicMock()
        mock_file_model.id = "test-file-id-123"
        mock_file_model.filename = "test_report.xlsx"
        
        # Execute
        await emit_excel_artifact(
            event_emitter=mock_emitter,
            file_model=mock_file_model,
            webui_url="http://localhost:3000",
            sheet_names=["Sheet1", "Sheet2"],
            active_sheet="Sheet1",
        )
        
        # Verify event structure (FROZEN CONTRACT)
        mock_emitter.assert_called_once()
        event = mock_emitter.call_args[0][0]
        
        # Event type must be "files"
        assert event["type"] == "files"
        
        # Data structure
        assert "data" in event
        assert "files" in event["data"]
        assert isinstance(event["data"]["files"], list)
        assert len(event["data"]["files"]) == 1
        
        # File artifact structure
        file_artifact = event["data"]["files"][0]
        assert file_artifact["type"] == "excel"
        assert file_artifact["url"] == "http://localhost:3000/api/v1/files/test-file-id-123/content"
        assert file_artifact["name"] == "test_report.xlsx"
        assert file_artifact["fileId"] == "test-file-id-123"
        
        # Metadata structure
        assert "meta" in file_artifact
        assert file_artifact["meta"]["sheetNames"] == ["Sheet1", "Sheet2"]
        assert file_artifact["meta"]["activeSheet"] == "Sheet1"

    @pytest.mark.asyncio
    async def test_emit_excel_artifact_defaults_active_sheet(self):
        """Verify activeSheet defaults to first sheet if not specified."""
        mock_emitter = AsyncMock()
        mock_file_model = MagicMock()
        mock_file_model.id = "test-id"
        mock_file_model.filename = "test.xlsx"
        
        await emit_excel_artifact(
            event_emitter=mock_emitter,
            file_model=mock_file_model,
            webui_url="http://localhost:3000",
            sheet_names=["FirstSheet", "SecondSheet"],
            active_sheet=None,  # Not specified
        )
        
        event = mock_emitter.call_args[0][0]
        file_artifact = event["data"]["files"][0]
        
        # Should default to first sheet
        assert file_artifact["meta"]["activeSheet"] == "FirstSheet"

    @pytest.mark.asyncio
    async def test_emit_excel_artifact_handles_empty_sheets(self):
        """Verify graceful handling when no sheet names provided."""
        mock_emitter = AsyncMock()
        mock_file_model = MagicMock()
        mock_file_model.id = "test-id"
        mock_file_model.filename = "test.xlsx"
        
        await emit_excel_artifact(
            event_emitter=mock_emitter,
            file_model=mock_file_model,
            webui_url="http://localhost:3000",
            sheet_names=None,
            active_sheet=None,
        )
        
        event = mock_emitter.call_args[0][0]
        file_artifact = event["data"]["files"][0]
        
        # Meta should still be present but empty
        assert "meta" in file_artifact
        # Should not fail, just have empty/minimal meta

    @pytest.mark.asyncio
    async def test_emit_file_artifacts_detects_excel_type(self):
        """Verify emit_file_artifacts correctly identifies Excel files."""
        mock_emitter = AsyncMock()
        mock_file_model = MagicMock()
        mock_file_model.id = "excel-file-id"
        mock_file_model.filename = "data.xlsx"
        mock_file_model.meta = {
            "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
        
        await emit_file_artifacts(
            event_emitter=mock_emitter,
            file_models=[mock_file_model],
            webui_url="http://localhost:3000",
        )
        
        event = mock_emitter.call_args[0][0]
        file_artifact = event["data"]["files"][0]
        
        assert file_artifact["type"] == "excel"


class TestExcelFileRecord:
    """Tests for Excel file record creation."""

    def test_create_excel_file_record_metadata_structure(self):
        """Verify file record metadata matches expected structure."""
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            f.write(b"test content")
            temp_path = f.name
        
        try:
            with patch('open_webui.utils.artifacts.Files') as mock_files:
                mock_file_model = MagicMock()
                mock_file_model.id = "new-file-id"
                mock_files.insert_new_file.return_value = mock_file_model
                
                result = create_excel_file_record(
                    user_id="user-123",
                    file_path=temp_path,
                    filename="report.xlsx",
                    sheet_names=["Summary", "Details"],
                )
                
                # Verify the file form was created correctly
                call_args = mock_files.insert_new_file.call_args
                file_form = call_args[0][1]  # Second positional arg
                
                assert file_form.filename == "report.xlsx"
                assert file_form.meta["content_type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                assert file_form.meta["sheetNames"] == ["Summary", "Details"]
                assert file_form.meta["size"] > 0
        finally:
            os.unlink(temp_path)

    def test_create_excel_file_record_handles_missing_file(self):
        """Verify graceful handling when file doesn't exist."""
        with patch('open_webui.utils.artifacts.Files') as mock_files:
            mock_files.insert_new_file.return_value = None
            
            result = create_excel_file_record(
                user_id="user-123",
                file_path="/nonexistent/path.xlsx",
                filename="missing.xlsx",
            )
            
            # Should not raise, but return None or handle gracefully
            # The actual behavior depends on implementation


class TestEventContractRegression:
    """
    Regression tests to catch any breaking changes to the event contract.
    These tests should FAIL if anyone modifies the event structure.
    """

    @pytest.mark.asyncio
    async def test_event_type_must_be_files(self):
        """EVENT CONTRACT: type must be 'files' (not 'excel', 'file', etc.)"""
        mock_emitter = AsyncMock()
        mock_file_model = MagicMock()
        mock_file_model.id = "id"
        mock_file_model.filename = "test.xlsx"
        
        await emit_excel_artifact(
            event_emitter=mock_emitter,
            file_model=mock_file_model,
            webui_url="http://localhost:3000",
        )
        
        event = mock_emitter.call_args[0][0]
        assert event["type"] == "files", "EVENT CONTRACT VIOLATION: type must be 'files'"

    @pytest.mark.asyncio
    async def test_file_type_must_be_excel(self):
        """EVENT CONTRACT: file.type must be 'excel' for Excel artifacts."""
        mock_emitter = AsyncMock()
        mock_file_model = MagicMock()
        mock_file_model.id = "id"
        mock_file_model.filename = "test.xlsx"
        
        await emit_excel_artifact(
            event_emitter=mock_emitter,
            file_model=mock_file_model,
            webui_url="http://localhost:3000",
        )
        
        event = mock_emitter.call_args[0][0]
        file_artifact = event["data"]["files"][0]
        assert file_artifact["type"] == "excel", "EVENT CONTRACT VIOLATION: file type must be 'excel'"

    @pytest.mark.asyncio
    async def test_url_format_must_match_api_pattern(self):
        """EVENT CONTRACT: URL must follow /api/v1/files/{id}/content pattern."""
        mock_emitter = AsyncMock()
        mock_file_model = MagicMock()
        mock_file_model.id = "file-id-123"
        mock_file_model.filename = "test.xlsx"
        
        await emit_excel_artifact(
            event_emitter=mock_emitter,
            file_model=mock_file_model,
            webui_url="http://localhost:3000",
        )
        
        event = mock_emitter.call_args[0][0]
        url = event["data"]["files"][0]["url"]
        
        assert "/api/v1/files/" in url, "EVENT CONTRACT VIOLATION: URL must contain /api/v1/files/"
        assert "/content" in url, "EVENT CONTRACT VIOLATION: URL must end with /content"
        assert "file-id-123" in url, "EVENT CONTRACT VIOLATION: URL must contain file ID"
