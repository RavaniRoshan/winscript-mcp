import re
import difflib
from dataclasses import dataclass
from typing import Optional, Any

@dataclass
class CoordinateElement:
    """Pseudo-element for coordinate-based interaction."""
    x: int
    y: int
    label: str = ""
    layer_used: str = "coordinates"

    def click_input(self):
        try:
            import pyautogui
        except ImportError:
            pyautogui = None
        if pyautogui is None:
            raise ImportError("pyautogui is not available")
        pyautogui.click(self.x, self.y)

    def window_text(self) -> str:
        return self.label

@dataclass
class SelectorResult:
    element: Any  # pywinauto element OR CoordinateElement
    layer_used: str  # "uia_name" | "uia_auto_id" | "uia_role" | "ocr" | "coordinates"
    confidence: float  # 0.0–1.0
    fallback_used: bool

def find_element(
    window,
    name: str,
    ocr_enabled: bool = True
) -> Optional[SelectorResult]:
    """
    Find a UI element using a 5-layer fallback chain.
    Returns SelectorResult with element + metadata, or None if all layers fail.
    """

    # Layer 1: UIA exact name
    try:
        el = window.child_window(title=name, found_index=0)
        el.wait("exists", timeout=2)
        return SelectorResult(el, "uia_name", 1.0, False)
    except Exception:
        pass

    # Layer 2: UIA automation_id
    try:
        el = window.child_window(auto_id=name, found_index=0)
        el.wait("exists", timeout=2)
        return SelectorResult(el, "uia_auto_id", 1.0, False)
    except Exception:
        pass

    # Layer 3: UIA role + partial name fuzzy
    try:
        best_match = None
        best_ratio = 0.0
        for child in window.descendants():
            try:
                title = child.window_text() or ""
                ratio = difflib.SequenceMatcher(
                    None, name.lower(), title.lower()
                ).ratio()
                if ratio > best_ratio and ratio > 0.6:
                    best_ratio = ratio
                    best_match = child
            except Exception:
                continue
        if best_match:
            return SelectorResult(best_match, "uia_role", best_ratio, True)
    except Exception:
        pass

    # Layer 4: OCR scan
    if ocr_enabled:
        try:
            import pytesseract
            from PIL import Image
            import mss
            import io

            with mss.mss() as sct:
                # Grab window region
                rect = window.rectangle()
                monitor = {
                    "top": rect.top, "left": rect.left,
                    "width": rect.width(), "height": rect.height()
                }
                shot = sct.grab(monitor)
                img = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")

            # Get OCR data with bounding boxes
            data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
            best_conf = 0.0
            best_x, best_y = None, None
            best_text = ""

            for i, text in enumerate(data["text"]):
                if not text.strip():
                    continue
                ratio = difflib.SequenceMatcher(
                    None, name.lower(), text.lower()
                ).ratio()
                conf = float(data["conf"][i]) / 100.0
                score = ratio * conf
                if score > best_conf and ratio > 0.7:
                    best_conf = score
                    x = rect.left + data["left"][i] + data["width"][i] // 2
                    y = rect.top + data["top"][i] + data["height"][i] // 2
                    best_x, best_y = x, y
                    best_text = text

            if best_x is not None:
                el = CoordinateElement(best_x, best_y, best_text, "ocr")
                return SelectorResult(el, "ocr", best_conf, True)
        except Exception:
            pass

    # Layer 5: Raw coordinates
    coord_match = re.match(r"x=(\d+),y=(\d+)", name.strip())
    if not coord_match:
        coord_match = re.match(r"\(\s*(\d+)\s*,\s*(\d+)\s*\)", name.strip())
    if coord_match:
        x, y = int(coord_match.group(1)), int(coord_match.group(2))
        el = CoordinateElement(x, y, name, "coordinates")
        return SelectorResult(el, "coordinates", 1.0, True)

    if name.lower() == "center":
        try:
            rect = window.rectangle()
            x = rect.left + rect.width() // 2
            y = rect.top + rect.height() // 2
            el = CoordinateElement(x, y, "center", "coordinates")
            return SelectorResult(el, "coordinates", 1.0, True)
        except Exception:
            pass

    return None