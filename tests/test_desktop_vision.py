import json
import sys
import types
from pathlib import Path
from types import SimpleNamespace

from Mudabbir.tools.builtin.desktop import DesktopTool


def test_describe_screen_is_multimodal_first(tmp_path, monkeypatch) -> None:
    settings = SimpleNamespace(
        vision_provider="openai",
        vision_model="gpt-4o",
        vision_fallback_ocr_enabled=True,
        openai_api_key="test-key",
        google_api_key="",
    )
    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.get_settings", lambda: settings)
    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.get_media_dir", lambda: Path(tmp_path))

    class _FakeImage:
        def save(self, path):
            Path(path).write_bytes(b"fake-image")

    fake_pil = types.ModuleType("PIL")
    fake_image_grab = types.ModuleType("PIL.ImageGrab")
    fake_image_grab.grab = lambda *args, **kwargs: _FakeImage()
    fake_pil.ImageGrab = fake_image_grab
    monkeypatch.setitem(sys.modules, "PIL", fake_pil)
    monkeypatch.setitem(sys.modules, "PIL.ImageGrab", fake_image_grab)

    class _FakeResponse:
        def raise_for_status(self) -> None:
            return None

        def json(self):
            return {
                "choices": [
                    {
                        "message": {
                            "content": json.dumps(
                                {
                                    "ui_summary": "Task Manager is open and sorted by memory.",
                                    "top_app": "Cursor",
                                    "detected_elements": ["Task Manager", "Memory column", "Cursor (24)"],
                                    "confidence": 0.94,
                                }
                            )
                        }
                    }
                ]
            }

    class _FakeClient:
        def __init__(self, *args, **kwargs):
            pass

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

        def post(self, *args, **kwargs):
            return _FakeResponse()

    import httpx

    monkeypatch.setattr(httpx, "Client", _FakeClient)

    raw = DesktopTool()._vision_tools(mode="describe_screen")
    parsed = json.loads(raw)
    assert parsed["ok"] is True
    assert parsed["source"] == "vision"
    assert parsed["top_app"] == "Cursor"
