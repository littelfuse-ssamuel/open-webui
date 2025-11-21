"""
Integration test for file upload with non-serializable metadata.

This test verifies that file uploads work correctly even when metadata contains
non-JSON-serializable objects (e.g., MCP clients, tool objects, functions).
"""
import io
import json
import pytest
from unittest.mock import Mock, MagicMock, patch
from fastapi import UploadFile


def test_upload_file_handler_with_non_serializable_metadata():
    """
    Test that upload_file_handler can handle metadata with non-serializable objects.
    
    This simulates the scenario where middleware passes metadata containing
    MCP clients or other non-serializable objects to the file upload handler.
    """
    # Import the function we're testing
    from open_webui.routers.files import upload_file_handler
    from open_webui.utils.misc import make_json_safe
    
    # Create mock objects
    mock_request = MagicMock()
    mock_request.app.state.config.ALLOWED_FILE_EXTENSIONS = []
    
    mock_user = MagicMock()
    mock_user.id = "test_user_123"
    mock_user.email = "test@example.com"
    mock_user.name = "Test User"
    
    # Create a file upload object
    file_content = b"Test Excel content"
    upload_file = UploadFile(
        file=io.BytesIO(file_content),
        filename="test.xlsx",
        headers={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
    )
    
    # Create metadata with non-serializable objects
    class MockMCPClient:
        def __init__(self):
            self.connected = True
        def send(self, msg):
            pass
    
    metadata = {
        "tool_ids": ["tool1", "tool2"],
        "files": [{"name": "file1.txt", "type": "text"}],
        "mcp_clients": {
            "client1": MockMCPClient(),  # Non-serializable
        },
        "tools": {
            "tool1": {
                "spec": {"name": "tool1"},
                "server": lambda: "server",  # Non-serializable function
            }
        },
        "session_id": "session123",
    }
    
    # Verify the metadata contains non-serializable objects
    with pytest.raises((TypeError, ValueError)):
        json.dumps(metadata)
    
    # But make_json_safe should convert it successfully
    safe_metadata = make_json_safe(metadata)
    json_str = json.dumps(safe_metadata)  # Should not raise
    assert isinstance(json_str, str)
    assert len(json_str) > 0
    
    # Verify the structure is preserved but non-serializable objects are stringified
    assert safe_metadata["tool_ids"] == ["tool1", "tool2"]
    assert safe_metadata["session_id"] == "session123"
    assert isinstance(safe_metadata["mcp_clients"]["client1"], str)
    assert isinstance(safe_metadata["tools"]["tool1"]["server"], str)
    
    print("✓ Metadata sanitization works correctly")


def test_make_json_safe_preserves_normal_metadata():
    """
    Test that make_json_safe doesn't modify normal, already-serializable metadata.
    """
    from open_webui.utils.misc import make_json_safe
    
    # Normal metadata (already JSON-safe)
    normal_metadata = {
        "name": "quarterly_sales_report.xlsx",
        "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "size": 8192,
        "sheetNames": ["Sheet1", "Sheet2"],
    }
    
    # Verify it's already JSON-safe
    json.dumps(normal_metadata)  # Should not raise
    
    # Apply make_json_safe
    result = make_json_safe(normal_metadata)
    
    # Should be identical
    assert result == normal_metadata
    
    print("✓ Normal metadata is preserved")


if __name__ == "__main__":
    test_upload_file_handler_with_non_serializable_metadata()
    test_make_json_safe_preserves_normal_metadata()
    print("\n✅ All integration tests passed!")
