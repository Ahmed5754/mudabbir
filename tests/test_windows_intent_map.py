from Mudabbir.tools.capabilities.windows_intent_map import (
    is_confirmation_message,
    resolve_windows_intent,
)


def test_resolve_arabic_volume_set() -> None:
    result = resolve_windows_intent("خلي الصوت 33%")
    assert result.matched is True
    assert result.action == "volume"
    assert result.params.get("mode") == "set"
    assert result.params.get("level") == 33


def test_resolve_shutdown_is_destructive() -> None:
    result = resolve_windows_intent("shutdown the pc now")
    assert result.matched is True
    assert result.action == "system_power"
    assert result.params.get("mode") == "shutdown"
    assert result.risk_level == "destructive"


def test_resolve_unsupported_audio_output() -> None:
    result = resolve_windows_intent("change audio output to headset")
    assert result.matched is True
    assert result.unsupported is True
    assert "not implemented" in result.unsupported_reason.lower()


def test_confirmation_message_detection() -> None:
    assert is_confirmation_message("yes") is True
    assert is_confirmation_message("نعم نفذ") is True
    assert is_confirmation_message("cancel") is False
