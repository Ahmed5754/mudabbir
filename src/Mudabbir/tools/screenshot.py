"""Screenshot tool."""

import io
from typing import Optional

try:
    import pyautogui

    PYAUTOGUI_AVAILABLE = True
except Exception:
    PYAUTOGUI_AVAILABLE = False


def take_screenshot() -> Optional[bytes]:
    """Take a screenshot and return as bytes."""
    if not PYAUTOGUI_AVAILABLE:
        return None

    try:
        # Take screenshot
        screenshot = pyautogui.screenshot()

        # Convert to bytes
        buffer = io.BytesIO()
        screenshot.save(buffer, format="PNG")
        buffer.seek(0)

        return buffer.getvalue()
    except Exception as e:
        # Common on headless servers or when display is not available
        return None
