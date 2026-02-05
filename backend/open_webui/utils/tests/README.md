# Excel Artifact Tests

This directory contains regression tests for the Excel artifact emission and file handling system.

## Phase 1 - Stabilize & Baseline

These tests ensure the event contract is preserved and no breaking changes are introduced to Excel artifact flows.

## Running Tests

The tests require the full backend environment to be set up:

```bash
cd backend

# Install dependencies
pip install -r requirements.txt

# Run the tests
pytest open_webui/utils/tests/test_excel_artifacts.py -v
```

## Test Coverage

### Event Contract Tests (`TestExcelArtifactEventContract`)
- ✅ Verifies `emit_excel_artifact` produces correct event structure
- ✅ Verifies `activeSheet` defaults to first sheet if not specified
- ✅ Verifies graceful handling when no sheet names provided
- ✅ Verifies `emit_file_artifacts` correctly identifies Excel files

### File Record Tests (`TestExcelFileRecord`)
- ✅ Verifies file record metadata matches expected structure
- ✅ Verifies graceful handling when file doesn't exist

### Regression Tests (`TestEventContractRegression`)
- ✅ Ensures event type is 'files' (frozen contract)
- ✅ Ensures file type is 'excel' for Excel artifacts (frozen contract)
- ✅ Ensures URL follows `/api/v1/files/{id}/content` pattern (frozen contract)

## Event Contract (FROZEN)

**DO NOT MODIFY WITHOUT MIGRATION PLAN**

```
Event Type: "files"
Event Structure:
{
  type: "files",
  data: {
    files: ExcelArtifact[]
  }
}

Required fields for Excel artifacts:
- type: "excel" (literal)
- url: string (download URL from /api/v1/files/{id}/content)
- name: string (filename with .xlsx extension)

Optional fields:
- fileId: string (UUID for save operations)
- meta.sheetNames: string[] (list of sheet names)
- meta.activeSheet: string (default sheet to display)
```

## Adding New Tests

When adding new tests, ensure they:
1. Use mocks to avoid requiring actual files or database
2. Follow the existing test structure
3. Mark async tests with `@pytest.mark.asyncio`
4. Include clear docstrings explaining what is being tested
