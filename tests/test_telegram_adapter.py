from types import SimpleNamespace
from unittest.mock import AsyncMock

import pytest

from Mudabbir.bus.adapters.telegram_adapter import TelegramAdapter


@pytest.mark.asyncio
async def test_handle_start_uses_arabic_display_name(monkeypatch: pytest.MonkeyPatch) -> None:
    adapter = TelegramAdapter(token="dummy-token")

    mock_message = SimpleNamespace(reply_text=AsyncMock())
    update = SimpleNamespace(
        effective_user=SimpleNamespace(id=123, username="ahmed"),
        message=mock_message,
    )

    settings = SimpleNamespace(assistant_display_name_ar="مُدَبِّر")
    monkeypatch.setattr("Mudabbir.bus.adapters.telegram_adapter.get_settings", lambda: settings)

    await adapter._handle_start(update, context=None)

    assert mock_message.reply_text.await_count == 1
    await_args = mock_message.reply_text.await_args
    sent_text = await_args.kwargs.get("text")
    if sent_text is None and await_args.args:
        sent_text = await_args.args[0]
    assert "مُدَبِّر" in str(sent_text)
