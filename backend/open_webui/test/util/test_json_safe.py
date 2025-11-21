import pytest
from open_webui.utils.misc import make_json_safe


def test_make_json_safe_primitives():
    """Test that primitives are left as-is"""
    assert make_json_safe(None) is None
    assert make_json_safe("string") == "string"
    assert make_json_safe(42) == 42
    assert make_json_safe(3.14) == 3.14
    assert make_json_safe(True) is True
    assert make_json_safe(False) is False


def test_make_json_safe_lists():
    """Test that lists are converted properly"""
    assert make_json_safe([1, 2, 3]) == [1, 2, 3]
    assert make_json_safe(["a", "b", "c"]) == ["a", "b", "c"]
    assert make_json_safe([1, "two", 3.0, True, None]) == [1, "two", 3.0, True, None]


def test_make_json_safe_dicts():
    """Test that dicts are converted properly"""
    assert make_json_safe({"key": "value"}) == {"key": "value"}
    assert make_json_safe({"a": 1, "b": 2}) == {"a": 1, "b": 2}
    assert make_json_safe({"nested": {"key": "value"}}) == {"nested": {"key": "value"}}


def test_make_json_safe_nested_structures():
    """Test nested lists and dicts"""
    data = {
        "list": [1, 2, 3],
        "dict": {"nested": "value"},
        "mixed": [{"a": 1}, {"b": 2}],
    }
    expected = {
        "list": [1, 2, 3],
        "dict": {"nested": "value"},
        "mixed": [{"a": 1}, {"b": 2}],
    }
    assert make_json_safe(data) == expected


def test_make_json_safe_non_serializable_function():
    """Test that functions are converted to string representation"""
    def my_function():
        pass
    
    result = make_json_safe(my_function)
    assert isinstance(result, str)
    assert "function" in result.lower()


def test_make_json_safe_non_serializable_class():
    """Test that class instances are converted to string representation"""
    class MyClass:
        def __init__(self):
            self.value = 42
    
    obj = MyClass()
    result = make_json_safe(obj)
    assert isinstance(result, str)
    assert "MyClass" in result


def test_make_json_safe_dict_with_non_serializable():
    """Test dict containing non-serializable objects"""
    def my_function():
        return "test"
    
    data = {
        "normal_key": "normal_value",
        "function_key": my_function,
        "nested": {
            "another_function": lambda x: x,
        },
    }
    
    result = make_json_safe(data)
    assert isinstance(result, dict)
    assert result["normal_key"] == "normal_value"
    assert isinstance(result["function_key"], str)
    assert "function" in result["function_key"].lower()
    assert isinstance(result["nested"]["another_function"], str)


def test_make_json_safe_list_with_non_serializable():
    """Test list containing non-serializable objects"""
    def my_function():
        return "test"
    
    data = [1, "string", my_function, {"key": lambda: None}]
    
    result = make_json_safe(data)
    assert isinstance(result, list)
    assert result[0] == 1
    assert result[1] == "string"
    assert isinstance(result[2], str)
    assert isinstance(result[3], dict)
    assert isinstance(result[3]["key"], str)


def test_make_json_safe_tuple_to_list():
    """Test that tuples are converted to lists"""
    result = make_json_safe((1, 2, 3))
    assert isinstance(result, list)
    assert result == [1, 2, 3]


def test_make_json_safe_set_to_list():
    """Test that sets are converted to lists"""
    result = make_json_safe({1, 2, 3})
    assert isinstance(result, list)
    assert set(result) == {1, 2, 3}


def test_make_json_safe_max_depth():
    """Test that max depth prevents infinite recursion"""
    # Create a deeply nested structure
    data = {"level": 0}
    current = data
    for i in range(20):
        current["nested"] = {"level": i + 1}
        current = current["nested"]
    
    result = make_json_safe(data, max_depth=5)
    # Should handle this gracefully without error
    assert isinstance(result, dict)


def test_make_json_safe_metadata_with_mcp_client():
    """Test metadata structure similar to what causes the bug"""
    # Simulate metadata with non-serializable objects like MCP clients
    class MockMCPClient:
        def __init__(self):
            self.connected = True
        
        def send(self, msg):
            pass
    
    metadata = {
        "tool_ids": ["tool1", "tool2"],
        "files": [{"name": "file1.txt", "type": "text"}],
        "mcp_clients": {
            "client1": MockMCPClient(),
            "client2": MockMCPClient(),
        },
        "tools": {
            "tool1": {
                "spec": {"name": "tool1"},
                "direct": True,
                "server": lambda: "server",  # Non-serializable
            }
        },
        "session_id": "session123",
        "chat_id": "chat456",
    }
    
    result = make_json_safe(metadata)
    
    # Verify structure is maintained
    assert isinstance(result, dict)
    assert result["tool_ids"] == ["tool1", "tool2"]
    assert result["files"] == [{"name": "file1.txt", "type": "text"}]
    assert result["session_id"] == "session123"
    assert result["chat_id"] == "chat456"
    
    # Verify non-serializable objects are converted to strings
    assert isinstance(result["mcp_clients"]["client1"], str)
    assert isinstance(result["mcp_clients"]["client2"], str)
    assert isinstance(result["tools"]["tool1"]["server"], str)
    
    # Should still have the serializable parts
    assert result["tools"]["tool1"]["spec"] == {"name": "tool1"}
    assert result["tools"]["tool1"]["direct"] is True


def test_make_json_safe_excel_metadata():
    """Test metadata structure used for Excel artifacts"""
    metadata = {
        "name": "quarterly_sales_report.xlsx",
        "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "size": 8192,
        "sheetNames": ["Sheet1", "Sheet2"],
        # Simulate a non-serializable object that might be added
        "callback": lambda: "done",
    }
    
    result = make_json_safe(metadata)
    
    assert result["name"] == "quarterly_sales_report.xlsx"
    assert result["content_type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert result["size"] == 8192
    assert result["sheetNames"] == ["Sheet1", "Sheet2"]
    assert isinstance(result["callback"], str)
