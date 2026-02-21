# Mudabbir 0.5.3

Release date: 2026-02-21

## Highlights

- Added deterministic WhatsApp latest-message routing in `open_interpreter`:
  - Requests like "who messaged me last on WhatsApp" now force `vision_tools` with `describe_screen`.
  - Prevents falling through to raw model-generated `pyautogui` scripts for this scenario.
- Improved response behavior when image vision is not configured:
  - Returns a concise, explicit message to enable OpenAI/Gemini vision provider.
  - Avoids OCR-text style over-explanations for this flow.
- Version bump:
  - `pyproject.toml` -> `0.5.3`
  - `src/Mudabbir/__init__.py` -> `0.5.3`

## Tests

- Added routing coverage:
  - `tests/test_open_interpreter_deterministic_routing.py::test_whatsapp_latest_message_uses_vision_describe_screen`
