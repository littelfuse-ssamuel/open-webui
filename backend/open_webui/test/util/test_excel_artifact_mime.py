import asyncio
from pathlib import Path
from types import SimpleNamespace

from open_webui.utils import artifacts as artifacts_utils
from open_webui.utils import files as files_utils


def test_excel_content_type_from_filename_supports_all_excel_extensions():
    assert (
        artifacts_utils._excel_content_type_from_filename("report.xlsx")
        == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert (
        artifacts_utils._excel_content_type_from_filename("macro.xlsm")
        == "application/vnd.ms-excel.sheet.macroEnabled.12"
    )
    assert (
        artifacts_utils._excel_content_type_from_filename("legacy.xls")
        == "application/vnd.ms-excel"
    )


def test_emit_file_artifacts_classifies_xls_and_xlsm_as_excel():
    events: list[dict] = []

    async def _event_emitter(payload):
        events.append(payload)

    files = [
        SimpleNamespace(
            id="file-xls",
            filename="legacy.xls",
            meta={"content_type": "application/vnd.ms-excel"},
        ),
        SimpleNamespace(
            id="file-xlsm",
            filename="macro.xlsm",
            meta={"content_type": "application/vnd.ms-excel.sheet.macroEnabled.12"},
        ),
    ]

    asyncio.run(
        artifacts_utils.emit_file_artifacts(
            event_emitter=_event_emitter,
            file_models=files,
            webui_url="http://localhost:3000",
        )
    )

    emitted_files = events[0]["data"]["files"]
    assert emitted_files[0]["type"] == "excel"
    assert emitted_files[1]["type"] == "excel"


def test_create_excel_file_record_sets_content_type_from_extension(monkeypatch, tmp_path: Path):
    captured = {}
    file_path = tmp_path / "macro.xlsm"
    file_path.write_bytes(b"dummy")

    def _fake_insert_new_file(user_id, file_form):
        captured["meta"] = file_form.meta
        return SimpleNamespace(id=file_form.id, filename=file_form.filename)

    monkeypatch.setattr(artifacts_utils.Files, "insert_new_file", _fake_insert_new_file)

    result = artifacts_utils.create_excel_file_record(
        user_id="user-1",
        file_path=str(file_path),
        filename="macro.xlsm",
    )

    assert result is not None
    assert captured["meta"]["content_type"] == (
        "application/vnd.ms-excel.sheet.macroEnabled.12"
    )


def test_parse_excel_data_uri_accepts_xlsx_xlsm_and_xls():
    xlsx_content_type, xlsx_payload = files_utils._parse_excel_data_uri(
        "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,QUJD"
    )
    xlsm_content_type, xlsm_payload = files_utils._parse_excel_data_uri(
        "data:application/vnd.ms-excel.sheet.macroEnabled.12;base64,REVG"
    )
    xls_content_type, xls_payload = files_utils._parse_excel_data_uri(
        "data:application/vnd.ms-excel;base64,R0hJ"
    )

    assert xlsx_content_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert xlsx_payload == "QUJD"
    assert xlsm_content_type == "application/vnd.ms-excel.sheet.macroenabled.12"
    assert xlsm_payload == "REVG"
    assert xls_content_type == "application/vnd.ms-excel"
    assert xls_payload == "R0hJ"
