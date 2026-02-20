from Mudabbir.agents.executor import (
    ExecutionResult,
    _extract_percent_value,
    _format_execution_summary,
    _looks_like_planning_text,
)


def test_executor_detects_planning_text() -> None:
    assert _looks_like_planning_text("Here's my plan: 1) check volume 2) set volume")
    assert _looks_like_planning_text("لنبدأ خطوة بخطوة")
    assert not _looks_like_planning_text("System volume successfully increased to 37%")


def test_executor_extract_percent_value() -> None:
    assert _extract_percent_value("Current volume: 48%") == "48%"
    assert _extract_percent_value("Volume is 130%") == ""
    assert _extract_percent_value("No percent here") == ""


def test_executor_format_execution_summary() -> None:
    ok = ExecutionResult(status="ok", action="computer task", evidence="done", final_value="62%")
    err = ExecutionResult(
        status="error",
        action="computer task",
        error="quota/rate limit reached",
        error_code="quota_exhausted",
    )
    assert "succeeded" in _format_execution_summary(ok)
    assert "[quota_exhausted]" in _format_execution_summary(err)
