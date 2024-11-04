import sys
import keyboard
from PySide6.QtWidgets import (QApplication, QMainWindow, QSystemTrayIcon, 
                              QMenu, QStyle, QVBoxLayout, QWidget, 
                              QTextEdit, QPushButton, QLabel, QHBoxLayout,
                              QComboBox) 
from PySide6.QtGui import QIcon, QPixmap, QImage
from PySide6.QtCore import Qt, QTimer
from PIL import ImageGrab, Image
import win32gui
import win32process
import os
import numpy as np
import cv2  # Add this
import pytesseract  # Add this
from datetime import datetime  # Add this
from context_manager import ContextManager

# Set Tesseract path - adjust this path to match your installation
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Debug print
print("Tesseract Version:", pytesseract.get_tesseract_version())
print("Tesseract Path:", pytesseract.pytesseract.tesseract_cmd)

class ScreenAssistant(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.setupHotkeys()
        self.setupScreenCapture()
        self.context_manager = ContextManager()
        self.update_window_list()
        
    def initUI(self):
        # Create main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # Add UI elements
        self.status_label = QLabel("Screen Assistant Active")
        self.status_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        
        selector_layout = QHBoxLayout()
        self.window_selector = QComboBox()
        self.window_selector.setMinimumWidth(200)
        self.refresh_windows_button = QPushButton("Refresh Windows")
        self.refresh_windows_button.clicked.connect(self.update_window_list)
        selector_layout.addWidget(QLabel("Select Window:"))
        selector_layout.addWidget(self.window_selector)
        selector_layout.addWidget(self.refresh_windows_button)
        selector_layout.addStretch()

        # Create horizontal layout for preview and input
        h_layout = QHBoxLayout()
        
        # Left side - Preview section
        preview_layout = QVBoxLayout()
        self.preview_label = QLabel("Screen Preview")
        self.preview_label.setStyleSheet("font-weight: bold;")
        self.image_label = QLabel()
        self.image_label.setMinimumSize(400, 300)
        self.image_label.setStyleSheet("border: 1px solid gray;")
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview_layout.addWidget(self.preview_label)
        preview_layout.addWidget(self.image_label)
        
        # Right side - Input/Output section
        input_layout = QVBoxLayout()
        self.input_box = QTextEdit()
        self.input_box.setPlaceholderText("Type your question here...")
        self.input_box.setMinimumHeight(100)
        
        self.analyze_button = QPushButton("Analyze (Mock Response)")
        self.analyze_button.setStyleSheet("font-size: 12px; padding: 5px;")
        
        self.output_box = QTextEdit()
        self.output_box.setReadOnly(True)
        self.output_box.setMinimumHeight(200)
        self.output_box.setPlaceholderText("Analysis results will appear here...")
        
        input_layout.addWidget(self.input_box)
        input_layout.addWidget(self.analyze_button)
        input_layout.addWidget(self.output_box)
        

        layout.addLayout(selector_layout)
        layout.addLayout(h_layout)
        # Add layouts to horizontal layout
        h_layout.addLayout(preview_layout)
        h_layout.addLayout(input_layout)
        
        # Add all to main layout
        layout.addWidget(self.status_label)
        layout.addLayout(h_layout)
        
        # Connect button
        self.analyze_button.clicked.connect(self.mock_analysis)
        
        # Create system tray icon
        self.tray_icon = QSystemTrayIcon(self)
        icon = QIcon(self.style().standardPixmap(QStyle.StandardPixmap.SP_ComputerIcon))
        self.tray_icon.setIcon(icon)
        
        # Create tray menu
        tray_menu = QMenu()
        show_action = tray_menu.addAction("Show Window")
        show_action.triggered.connect(self.show_window)
        self.toggle_action = tray_menu.addAction("Pause")
        self.toggle_action.triggered.connect(self.toggleService)
        quit_action = tray_menu.addAction("Quit")
        quit_action.triggered.connect(QApplication.instance().quit)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_icon_clicked)
        self.tray_icon.show()
        
        # Set window properties
        self.setWindowTitle('Screen Assistant')
        self.setGeometry(100, 100, 1200, 800)  # Made window larger for preview
        self.setWindowFlags(Qt.WindowType.Window)
        print("UI initialized successfully")

        
        
    def update_preview(self):
        """Update the preview image and return PIL image for analysis"""
        try:
            # Capture screen
            screenshot = ImageGrab.grab()
            
            # Convert PIL image to QPixmap for display
            screenshot_rgb = screenshot.convert('RGB')
            image = QImage(screenshot_rgb.tobytes(), screenshot_rgb.width, 
                        screenshot_rgb.height, 3 * screenshot_rgb.width, 
                        QImage.Format.Format_RGB888)
            
            # Scale to fit preview maintaining aspect ratio
            pixmap = QPixmap.fromImage(image)
            scaled_pixmap = pixmap.scaled(self.image_label.size(), 
                                        Qt.AspectRatioMode.KeepAspectRatio,
                                        Qt.TransformationMode.SmoothTransformation)
            
            self.image_label.setPixmap(scaled_pixmap)
            return screenshot  # Return the original PIL Image
        except Exception as e:
            print(f"Error updating preview: {e}")
            return None
        
    def show_window(self):
        self.show()
        self.setWindowState(self.windowState() & ~Qt.WindowState.WindowMinimized | Qt.WindowState.WindowActive)
        self.activateWindow()
        self.update_preview()  # Update preview when window is shown
        print("Showing main window")
        
    def tray_icon_clicked(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.show_window()
            
    def setupHotkeys(self):
        try:
            keyboard.add_hotkey('alt+shift+a', self.show_window)
            keyboard.add_hotkey('ctrl+q', self.close)
            self.active = True
            print("Hotkeys registered (Alt+Shift+A for assistant, Ctrl+Q to quit)")
        except Exception as e:
            print(f"Error setting up hotkey: {e}")
        
    def setupScreenCapture(self):
        self.capture_timer = QTimer()
        self.capture_timer.timeout.connect(self.updateContext)
        self.capture_timer.start(1000)  # Update every second
        print("Screen capture timer started")
        
    def mock_analysis(self):
        """Analyze current context"""
        try:
            # Get selected window
            hwnd = self.get_selected_window()
            if not hwnd:
                self.output_box.setText("Please select a window to analyze")
                return
                
            # Get window rectangle
            rect = win32gui.GetWindowRect(hwnd)
            title = win32gui.GetWindowText(hwnd)
            
            # Take screenshot of just the selected window
            screenshot = ImageGrab.grab(bbox=rect)
            
            # Update preview with full screen
            self.update_preview()
            
            # Process screenshot for OCR
            cv_image = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(cv_image, cv2.COLOR_BGR2GRAY)
            gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
            
            # OCR with custom config
            custom_config = r'--oem 3 --psm 6'
            text = pytesseract.image_to_string(gray, config=custom_config)
            
            # Create analysis output
            output = "Context Analysis:\n\n"
            output += f"Active Window: {title}\n\n"
            output += "Detected Text:\n"
            output += text.strip() or "No text detected"
            
            # Add recent actions
            output += "\n\nRecent Actions:\n"
            context = self.context_manager.get_recent_context(limit=5)
            for action in context["recent_actions"]:
                output += f"- {action.timestamp.strftime('%H:%M:%S')}: {action.action_type} in {action.window_title}\n"
            
            # Debug info
            output += "\n\nDebug Info:\n"
            output += f"Window Rectangle: {rect}\n"
            output += f"Screenshot Size: {screenshot.size}\n"
            
            self.output_box.setText(output)
            
        except Exception as e:
            self.output_box.setText(f"Error during analysis: {str(e)}")
            print(f"Error: {e}")
        
    def toggleService(self):
        self.active = not self.active
        self.toggle_action.setText("Resume" if not self.active else "Pause")
        self.status_label.setText(
            "Screen Assistant Paused" if not self.active else "Screen Assistant Active"
        )
        print(f"Service {'paused' if not self.active else 'resumed'}")
        
    def updateContext(self):
        if not self.active:
            return
        if self.isVisible():
            self.update_preview()  # Update preview only when window is visible


    def closeEvent(self, event):
        """Clean up resources before closing"""
        try:
            # Stop context manager listeners
            self.context_manager.keyboard_listener.stop()
            self.context_manager.mouse_listener.stop()
            event.accept()
        except Exception as e:
            print(f"Error during cleanup: {e}")
            event.accept()



    def update_window_list(self):
        """Update the list of available windows"""
        self.window_selector.clear()
        self.windows_list = []
        
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title and title != 'Screen Assistant':  # Skip our own window
                    windows.append((title, hwnd))
            return True
        
        win32gui.EnumWindows(enum_windows_callback, self.windows_list)
        
        # Sort by window title
        self.windows_list.sort(key=lambda x: x[0].lower())
        
        # Add to combo box
        for title, _ in self.windows_list:
            self.window_selector.addItem(title)

    def get_selected_window(self):
        """Get the currently selected window handle"""
        current_index = self.window_selector.currentIndex()
        if current_index >= 0 and current_index < len(self.windows_list):
            return self.windows_list[current_index][1]
        return None

if __name__ == '__main__':
    try:
        app = QApplication(sys.argv)
        app.setQuitOnLastWindowClosed(False)
        screen_assistant = ScreenAssistant()
        
        # Show initial message
        if QSystemTrayIcon.isSystemTrayAvailable():
            screen_assistant.tray_icon.showMessage(
                "Screen Assistant",
                "Application is running. Press Alt+Shift+A to show window.",
                QSystemTrayIcon.MessageIcon.Information,
                3000
            )
        
        print("Application initialized successfully")
        print("Press Ctrl+Q to quit or use the system tray menu")
        sys.exit(app.exec())
    except Exception as e:
        print(f"Error starting application: {e}")