import asyncio
import json
from types import SimpleNamespace

from open_webui.tools import builtin


def _make_request(
    *,
    excel_enabled: bool = True,
    dfmea_enabled: bool = True,
    pptx_enabled: bool = True,
):
    config = SimpleNamespace(
        ENABLE_EXCEL_ARTIFACT_TOOLS=excel_enabled,
        ENABLE_DFMEA_ARTIFACT_TOOLS=dfmea_enabled,
        ENABLE_PPTX_ARTIFACT_TOOLS=pptx_enabled,
    )
    return SimpleNamespace(app=SimpleNamespace(state=SimpleNamespace(config=config)))


def _user_payload():
    return {
        "id": "user-1",
        "email": "user@example.com",
        "role": "user",
        "name": "Test User",
        "profile_image_url": "",
        "last_active_at": 0,
        "updated_at": 0,
        "created_at": 0,
    }


def test_generate_excel_artifact_emits_files_event(monkeypatch):
    async def _fake_generate_excel_file(request, user):
        return SimpleNamespace(
            status="ok",
            message="Workbook generated successfully",
            fileId="file-1",
            downloadUrl="/api/v1/files/file-1/content",
            artifact={
                "type": "excel",
                "url": "/api/v1/files/file-1/content",
                "name": "artifact.xlsx",
                "fileId": "file-1",
                "meta": {},
            },
        )

    monkeypatch.setattr(builtin, "generate_excel_file", _fake_generate_excel_file)
    monkeypatch.setattr(
        builtin.Chats,
        "add_message_files_by_id_and_message_id",
        lambda chat_id, message_id, files: files,
    )

    events = []

    async def _emit(event):
        events.append(event)

    response = asyncio.run(
        builtin.generate_excel_artifact(
            prompt="Create an engineering workbook",
            __request__=_make_request(excel_enabled=True),
            __user__=_user_payload(),
            __event_emitter__=_emit,
            __chat_id__="chat-1",
            __message_id__="msg-1",
        )
    )
    payload = json.loads(response)

    assert payload["status"] == "success"
    assert payload["fileId"] == "file-1"
    assert events
    assert events[0]["type"] == "files"
    assert events[0]["data"]["files"][0]["fileId"] == "file-1"


def test_generate_excel_artifact_respects_disabled_flag(monkeypatch):
    called = {"value": False}

    async def _fake_generate_excel_file(request, user):
        called["value"] = True
        return SimpleNamespace(status="ok", message="ok", fileId="unused", downloadUrl="")

    monkeypatch.setattr(builtin, "generate_excel_file", _fake_generate_excel_file)

    response = asyncio.run(
        builtin.generate_excel_artifact(
            prompt="Should not run",
            __request__=_make_request(excel_enabled=False),
            __user__=_user_payload(),
        )
    )
    payload = json.loads(response)

    assert "disabled" in payload["error"].lower()
    assert called["value"] is False


def test_generate_pptx_artifact_emits_files_event(monkeypatch):
    async def _fake_generate_pptx(request, form_data, user):
        return SimpleNamespace(
            success=True,
            file_id="pptx-file-1",
            download_url="/api/v1/files/pptx-file-1/content",
            slide_count=1,
            filename="deck.pptx",
            message="ok",
        )

    monkeypatch.setattr(builtin, "generate_pptx", _fake_generate_pptx)
    monkeypatch.setattr(
        builtin.Chats,
        "add_message_files_by_id_and_message_id",
        lambda chat_id, message_id, files: files,
    )

    events = []

    async def _emit(event):
        events.append(event)

    response = asyncio.run(
        builtin.generate_pptx_artifact(
            title="Roadmap",
            slides=[{"title": "Slide 1", "content": [{"type": "text", "text": "Hello"}]}],
            __request__=_make_request(pptx_enabled=True),
            __user__=_user_payload(),
            __event_emitter__=_emit,
            __chat_id__="chat-1",
            __message_id__="msg-1",
        )
    )
    payload = json.loads(response)

    assert payload["status"] == "success"
    assert payload["fileId"] == "pptx-file-1"
    assert events
    assert events[0]["type"] == "files"
    assert events[0]["data"]["files"][0]["type"] == "pptx"


def test_generate_dfmea_artifact_persists_and_emits(monkeypatch):
    monkeypatch.setattr(
        builtin.Storage,
        "upload_file",
        lambda file_obj, filename, metadata: (b"", f"uploads/{filename}"),
    )

    inserted = {}

    def _fake_insert_new_file(user_id, form_data):
        inserted["user_id"] = user_id
        inserted["form_data"] = form_data
        return SimpleNamespace(id=form_data.id, filename=form_data.filename)

    monkeypatch.setattr(builtin.Files, "insert_new_file", _fake_insert_new_file)
    monkeypatch.setattr(
        builtin.Chats,
        "add_message_files_by_id_and_message_id",
        lambda chat_id, message_id, files: files,
    )

    events = []

    async def _emit(event):
        events.append(event)

    response = asyncio.run(
        builtin.generate_dfmea_artifact(
            prompt="Power distribution shall detect overload\nSwitch shall fail safe",
            template_name="littelfuse",
            __request__=_make_request(dfmea_enabled=True),
            __user__=_user_payload(),
            __event_emitter__=_emit,
            __chat_id__="chat-1",
            __message_id__="msg-1",
        )
    )
    payload = json.loads(response)

    assert payload["status"] == "success"
    assert payload["fileId"]
    assert inserted["user_id"] == "user-1"
    assert inserted["form_data"].meta["dfmea_template"] == "littelfuse"
    assert events
    assert events[0]["type"] == "files"
    assert events[0]["data"]["files"][0]["type"] == "excel"
