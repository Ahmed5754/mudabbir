"""Quota/rate-limit handling tests for Open Interpreter backend."""

from __future__ import annotations

from Mudabbir.agents.open_interpreter import (
    _build_quota_error_message,
    _is_google_quota_signal,
    _is_quota_or_rate_limit_message,
    _should_retry_with_gemini_fallback,
)


def test_quota_detection_matches_known_openai_phrase() -> None:
    text = (
        "You ran out of current quota for OpenAI's API, please check your plan and billing details."
    )
    assert _is_quota_or_rate_limit_message(text) is True


def test_google_quota_detection_matches_resource_exhausted() -> None:
    text = "RESOURCE_EXHAUSTED: Quota exceeded for generativelanguage.googleapis.com"
    assert _is_google_quota_signal(text) is True
    assert _is_quota_or_rate_limit_message(text) is True


def test_quota_detection_does_not_match_normal_reply() -> None:
    text = "Hello, I can help you with your request."
    assert _is_google_quota_signal(text) is False
    assert _is_quota_or_rate_limit_message(text) is False


def test_gemini_quota_message_is_provider_specific() -> None:
    msg = _build_quota_error_message(
        "gemini",
        "gemini-3-flash-preview",
        "gemini-2.5-flash",
        arabic=False,
        fallback_attempted=True,
    )
    assert "Gemini" in msg
    assert "OpenAI billing page" not in msg
    assert "gemini-2.5-flash" in msg


def test_fallback_policy_retries_once_for_gemini_quota() -> None:
    quota_text = "RESOURCE_EXHAUSTED: Quota exceeded"
    assert _should_retry_with_gemini_fallback("gemini", False, quota_text) is True
    assert _should_retry_with_gemini_fallback("gemini", True, quota_text) is False
    assert _should_retry_with_gemini_fallback("openai", False, quota_text) is False
