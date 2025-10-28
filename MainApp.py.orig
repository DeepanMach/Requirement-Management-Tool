from __future__ import annotations
import os
import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QStackedWidget

# Adjust import paths
# Ensure Main_page.py is accessible when this file runs as `python -m requirement_tool.MainApp`
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

# Import UI pages (these are inside requirement_tool/)
from Main_page import MainFrontUI
from requirement_tool.ui.main_window import MainWindow


class MainApp(QMainWindow):
    """Unified launcher: smooth transition between Main Page and Project Window"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("MACH Requirement Management Tool")
        self.resize(1366, 860)

        # --- QStackedWidget holds both pages ---
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        # --- Page 1: Main front (project list) ---
        self.main_front = MainFrontUI()
        self.stack.addWidget(self.main_front)

        # --- Page 2: Project main window ---
        self.main_window = MainWindow()
        self.stack.addWidget(self.main_window)

        # --- Link the navigation ---
        # Replace open_project handler from MainFrontUI
        self.main_front.open_project = self.open_project

        # Make Home button switch back to front page
        if hasattr(self.main_window, "home_btn"):
            self.main_window.home_btn.clicked.connect(self.return_to_main_page)

        # Start on MainFrontUI
        self.stack.setCurrentWidget(self.main_front)

        # Fix missing asset paths for QPixmap warnings
        self._fix_asset_paths()

    # -------------------------------------------------------
    def _fix_asset_paths(self):
        """Ensure Main_page assets and logo paths resolve correctly."""
        base_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(base_dir)
        assets_dir = os.path.join(project_root, "assets_png")
        logo_path = os.path.join(project_root, "logo.png")

        if not os.path.exists(assets_dir):
            print(f"[Warning] Missing assets folder: {assets_dir}")
        else:
            os.environ["ASSETS_PNG_PATH"] = assets_dir

        if not os.path.exists(logo_path):
            print(f"[Warning] Missing logo.png in {project_root}")

    # -------------------------------------------------------
    def open_project(self, project_name: str):
        """Triggered when double-clicking a project card."""
        try:
            self.main_window.set_project_name(project_name)
            self.stack.setCurrentWidget(self.main_window)
        except Exception as e:
            print(f"[Error] Failed to open project '{project_name}': {e}")

    # -------------------------------------------------------
    def return_to_main_page(self):
        """Triggered by Home button to go back to MainFrontUI."""
        try:
            self.stack.setCurrentWidget(self.main_front)
        except Exception as e:
            print(f"[Error] Failed to return to main page: {e}")


# -------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec())
