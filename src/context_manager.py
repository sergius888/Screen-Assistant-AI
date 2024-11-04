from datetime import datetime
import pytesseract
import cv2
import numpy as np
from pynput import mouse, keyboard
from dataclasses import dataclass
from typing import List, Dict, Set, Optional
import win32gui
import win32process
import psutil

# Add this if Tesseract isn't in your PATH
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

@dataclass
class Region:
    x: int
    y: int
    width: int
    height: int
    content_type: str
    confidence: float
    content: Optional[str] = None

@dataclass
class UserAction:
    timestamp: datetime
    action_type: str
    window_title: str
    process_name: str
    extra_data: Dict = None

class ContextManager:
    def __init__(self):
        self.action_history: List[UserAction] = []
        self.active_regions: List[Region] = []
        self.current_window = None
        self.setup_tracking()
        
    def setup_tracking(self):
        """Initialize input tracking"""
        self.keyboard_listener = keyboard.Listener(on_press=self._on_key_press)
        self.mouse_listener = mouse.Listener(on_click=self._on_mouse_click)
        self.keyboard_listener.start()
        self.mouse_listener.start()

    def _on_key_press(self, key):
        """Track keyboard events"""
        try:
            window = win32gui.GetForegroundWindow()
            window_title = win32gui.GetWindowText(window)
            pid = win32process.GetWindowThreadProcessId(window)[1]
            process_name = psutil.Process(pid).name()
            
            action = UserAction(
                timestamp=datetime.now(),
                action_type="keyboard",
                window_title=window_title,
                process_name=process_name,
                extra_data={"key": str(key)}
            )
            self.action_history.append(action)
            
            # Keep only last 100 actions
            if len(self.action_history) > 100:
                self.action_history.pop(0)
        except Exception as e:
            print(f"Error tracking keyboard event: {e}")

    def _on_mouse_click(self, x, y, button, pressed):
        """Track mouse events"""
        if pressed:
            try:
                window = win32gui.GetForegroundWindow()
                window_title = win32gui.GetWindowText(window)
                pid = win32process.GetWindowThreadProcessId(window)[1]
                process_name = psutil.Process(pid).name()
                
                action = UserAction(
                    timestamp=datetime.now(),
                    action_type="mouse",
                    window_title=window_title,
                    process_name=process_name,
                    extra_data={"position": (x, y), "button": str(button)}
                )
                self.action_history.append(action)
            except Exception as e:
                print(f"Error tracking mouse event: {e}")

    def process_screenshot(self, image) -> List[Region]:
        """Process screenshot to detect regions of interest"""
        # Convert PIL image to OpenCV format
        cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        
        regions = []
        
        # Get text regions
        text_regions = self._detect_text_regions(cv_image)
        regions.extend(text_regions)
        
        # Get UI element regions
        ui_regions = self._detect_ui_regions(cv_image)
        regions.extend(ui_regions)
        
        self.active_regions = regions
        return regions

    def _detect_text_regions(self, image) -> List[Region]:
        """Detect regions containing text"""
        regions = []
        
        # Convert to grayscale
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # Improve image quality for OCR
        gray = cv2.medianBlur(gray, 3)
        gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        
        # Get OCR data with bounding boxes
        custom_config = r'--oem 3 --psm 11'
        ocr_data = pytesseract.image_to_data(gray, output_type=pytesseract.Output.DICT, config=custom_config)
        
        n_boxes = len(ocr_data['text'])
        for i in range(n_boxes):
            # Filter empty results and low confidence
            if float(ocr_data['conf'][i]) > 60 and ocr_data['text'][i].strip():
                region = Region(
                    x=ocr_data['left'][i],
                    y=ocr_data['top'][i],
                    width=ocr_data['width'][i],
                    height=ocr_data['height'][i],
                    content_type="text",
                    confidence=float(ocr_data['conf'][i]),
                    content=ocr_data['text'][i]
                )
                regions.append(region)
        
        return regions
    def _detect_ui_regions(self, image) -> List[Region]:
        """Detect UI elements like buttons, input boxes, etc."""
        regions = []
        
        # Convert to grayscale
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # Apply threshold to get binary image
        _, binary = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY_INV)
        
        # Find contours
        contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        for contour in contours:
            # Get bounding rectangle
            x, y, w, h = cv2.boundingRect(contour)
            
            # Filter small regions
            if w > 20 and h > 20:
                region = Region(
                    x=x,
                    y=y,
                    width=w,
                    height=h,
                    content_type="ui_element",
                    confidence=0.8  # Default confidence for UI elements
                )
                regions.append(region)
        
        return regions

    def get_recent_context(self, limit: int = 5) -> Dict:
        """Get recent context information"""
        return {
            "recent_actions": self.action_history[-limit:],
            "active_regions": self.active_regions,
            "current_window": self.current_window
        }
    

    def _filter_and_organize_text(self, regions: List[Region]) -> dict:
        """Organize detected text by active windows and content types"""
        organized = {
            "active_window": [],
            "window_title": "",
            "recent_text": []
        }
        
        try:
            # Get active window info
            window = win32gui.GetForegroundWindow()
            rect = win32gui.GetWindowRect(window)
            title = win32gui.GetWindowText(window)
            organized["window_title"] = title
            
            # Get window client area (excludes title bar and borders)
            client_rect = win32gui.GetClientRect(window)
            left, top, right, bottom = client_rect
            
            # Convert client coordinates to screen coordinates
            pt_left_top = win32gui.ClientToScreen(window, (left, top))
            pt_right_bottom = win32gui.ClientToScreen(window, (right, bottom))
            
            client_area = (
                pt_left_top[0],     # left
                pt_left_top[1],     # top
                pt_right_bottom[0], # right
                pt_right_bottom[1]  # bottom
            )
            
            print(f"Window: {title}")
            print(f"Client area: {client_area}")
            
            # Collect all text within the active window
            window_text = []
            for region in regions:
                if region.content_type == "text" and region.content:
                    # Skip common UI text and low confidence results
                    if (region.confidence < 70 or 
                        "confidence:" in region.content.lower() or
                        region.content in ["File", "Edit", "View", "Help"]):
                        continue
                        
                    # Check if text is within client area
                    if (client_area[0] <= region.x <= client_area[2] and 
                        client_area[1] <= region.y <= client_area[3]):
                        window_text.append({
                            "text": region.content.strip(),
                            "confidence": region.confidence,
                            "y": region.y
                        })
            
            # Sort text by vertical position
            window_text.sort(key=lambda x: x["y"])
            
            # Remove duplicates and merge nearby text
            merged_text = []
            prev_y = None
            current_line = []
            
            for text in window_text:
                if prev_y is None or abs(text["y"] - prev_y) > 10:
                    if current_line:
                        merged_text.append(" ".join(current_line))
                    current_line = [text["text"]]
                else:
                    current_line.append(text["text"])
                prev_y = text["y"]
                
            if current_line:
                merged_text.append(" ".join(current_line))
                
            organized["active_window"] = merged_text
                
            return organized
            
        except Exception as e:
            print(f"Error organizing text: {e}")
            return organized