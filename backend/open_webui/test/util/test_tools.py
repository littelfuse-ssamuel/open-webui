from open_webui.utils.tools import get_functions_from_tool


class _DummyTool:
    def public_method(self):
        return "ok"

    def _private_helper(self):
        return "hidden"

    def __dunder__(self):
        return "hidden"


def test_get_functions_from_tool_exposes_only_public_methods():
    functions = get_functions_from_tool(_DummyTool())
    names = {func.__name__ for func in functions}

    assert "public_method" in names
    assert "_private_helper" not in names
    assert all(not name.startswith("_") for name in names)
