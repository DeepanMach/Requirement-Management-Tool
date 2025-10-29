import sys, math, random, os, json, datetime, shutil
from PyQt6.QtCore import Qt, QTimer, QRectF, QPointF
from PyQt6.QtGui import QPixmap, QPainter
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout, QFrame,
    QPushButton, QGraphicsView, QGraphicsScene, QGraphicsPixmapItem,
    QScrollArea, QGraphicsDropShadowEffect, QInputDialog, QMessageBox, QMenu
)
from requirement_tool.ui.main_window import MainWindow

NAVY = "#022f57"
YELLOW = "#66b2ff"

from pathlib import Path

def resource_path(relative: str) -> str:
    """Return an absolute path to a bundled resource.

    - In a PyInstaller frozen app, resources are under `sys._MEIPASS`.
    - In dev mode, resolve relative to the repo root (this file's parent).
    """
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)

def canonical_projects_root(file_path: str | Path) -> Path:
    """
    Always return the OUTER `<.../Requirement-Management-Tool>/projects` folder,
    even if this file lives under the inner `<.../Requirement-Management-Tool/Requirement-Management-Tool/...>`.
    """
    # In frozen apps, prefer a user-writable location (LOCALAPPDATA on Windows)
    if getattr(sys, "frozen", False):
        if os.name == "nt":
            base = os.environ.get("LOCALAPPDATA") or os.path.join(Path.home(), "AppData", "Local")
            root = Path(base) / "Requirement-Management-Tool" / "projects"
        else:
            base = os.environ.get("XDG_DATA_HOME") or os.path.join(Path.home(), ".local", "share")
            root = Path(base) / "Requirement-Management-Tool" / "projects"
        root = root.resolve()
    else:
        here = Path(file_path).resolve()
        outer = None
        for p in here.parents:
            if p.name == "Requirement-Management-Tool":
                outer = p  # this will end up the *highest* one
        if outer is None:
            # Fallback if the folder name is different on your machine
            outer = here.parents[2]
        root = (outer / "projects").resolve()
    root.mkdir(parents=True, exist_ok=True)
    return root

def both_project_dirs(project_name: str, file_path: str | Path) -> list[Path]:
    """Return [canonical, legacy] project folders (legacy only if it exists)."""
    canon = canonical_projects_root(file_path) / project_name
    legacy = (canonical_projects_root(file_path).parent / "Requirement-Management-Tool" / "projects" / project_name)
    return [canon, legacy]

# ---------------- Sky animation ---------------- #
class SkyView(QGraphicsView):
    def __init__(self, assets_path=None, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True)
        self.scene = QGraphicsScene(self)
        self.setScene(self.scene)
        self.setRenderHints(QPainter.RenderHint.Antialiasing)
        self.setStyleSheet("background: transparent; border: none;")

        # Resolve assets path (env var > bundled resource > default)
        assets_path = assets_path or os.environ.get("ASSETS_PNG_PATH") or resource_path("assets_png")

        self.sky_bg = QGraphicsPixmapItem(QPixmap(f"{assets_path}/sky_bg.png"))
        self.sky_bg.setZValue(-10)
        self.scene.addItem(self.sky_bg)

        self.clouds = []
        for i, name in enumerate(["cloud1.png", "cloud2.png", "cloud3.png"]):
            for _ in range(2):
                pix = QPixmap(f"{assets_path}/{name}")
                scale = random.uniform(0.55, 0.95)
                item = QGraphicsPixmapItem(
                    pix.scaled(
                        int(pix.width() * scale),
                        int(pix.height() * scale),
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation,
                    )
                )
                item.setZValue(-5 if i < 2 else -4)
                item.speed = 0.30 + 0.15 * i
                item.layer = 1 if i < 2 else 2
                item.setPos(random.randint(-200, 1200), random.randint(-120, 180))
                self.scene.addItem(item)
                self.clouds.append(item)

        self.helis = []
        for path, pos, dir_sign, w in [
            (f"{assets_path}/uh60x.png", QPointF(260, 120), 0.45, 420),
            (f"{assets_path}/ch47.png", QPointF(960, 200), -0.35, 460),
        ]:
            pm = QPixmap(path).scaledToWidth(w, Qt.TransformationMode.SmoothTransformation)
            h = QGraphicsPixmapItem(pm)
            h.setTransformOriginPoint(pm.width() / 2, pm.height() / 2)
            h.setPos(pos)
            h.float_t = random.random() * math.tau
            h.dir = dir_sign
            h.base_speed = 0.5
            h.prox = 1.0
            h.setZValue(1)
            self.scene.addItem(h)
            self.helis.append(h)

        self.cursor_scene_pos = QPointF(self.width() / 2, self.height() / 2)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.tick)
        self.timer.start(16)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.setSceneRect(QRectF(self.rect()))
        pm = self.sky_bg.pixmap()
        if not pm.isNull():
            self.sky_bg.setScale(max(self.width() / pm.width(), self.height() / pm.height()))
        self.sky_bg.setPos(0, 0)

    def mouseMoveEvent(self, e):
        self.cursor_scene_pos = self.mapToScene(e.pos())
        cx, cy = self.width() / 2, self.height() / 2
        dx, dy = (e.pos().x() - cx), (e.pos().y() - cy)
        self.sky_bg.setOffset(-dx * 0.02, -dy * 0.015)
        for c in self.clouds:
            depth = 0.01 if c.layer == 1 else 0.02
            c.setOffset(-dx * depth, -dy * depth)
        super().mouseMoveEvent(e)

    def tick(self):
        rect = self.sceneRect()
        for c in self.clouds:
            c.setX(c.x() + c.speed * (0.6 if c.layer == 2 else 1.0))
            if c.x() > rect.right() + 240:
                c.setX(rect.left() - 260)
        for i, h in enumerate(self.helis):
            h.float_t += 0.02
            bob = math.sin(h.float_t) * 0.6
            target = self.cursor_scene_pos + QPointF(i * 80 - 100, -60 - i * 20)
            vx = (target.x() - (h.x() + h.pixmap().width() / 2)) * 0.02
            vy = (target.y() - (h.y() + h.pixmap().height() / 2)) * 0.02 + bob
            speed = h.base_speed * (0.6 + h.prox * 0.6)
            h.moveBy(vx + speed * h.dir, vy)
            h.setRotation(max(-10, min(10, -vx * 6)))


# ---------------- Main front page ---------------- #
class MainFrontUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MACH Requirement Management Tool")
        self.resize(1366, 860)
        self.setStyleSheet("QWidget { background:white; }")
        self.projects = []
        self.project_cards = []

        # Persistent storage
        from pathlib import Path
        here = Path(__file__).resolve()
        self.project_root = str((here.parent.parent / "projects").resolve())
        Path(self.project_root).mkdir(parents=True, exist_ok=True)
        self.project_list_file = os.path.join(self.project_root, "project_list.json")

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Header
        header = QFrame()
        header.setStyleSheet(f"background:{NAVY};")
        header.setFixedHeight(72)
        hl = QHBoxLayout(header)
        hl.setContentsMargins(20, 0, 20, 0)
        logo_path = resource_path(os.path.join("requirement_tool", "ui", "logo.png"))
        logo = QLabel()
        if os.path.exists(logo_path):
            logo.setPixmap(QPixmap(logo_path).scaled(150, 150, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        title = QLabel("MACH REQUIREMENT MANAGEMENT TOOL")
        title.setStyleSheet("color:white; font-weight:800; font-size:20px; margin-left:10px;")
        hl.addWidget(logo)
        hl.addWidget(title)
        hl.addStretch(1)
        root.addWidget(header)

        # Animated sky
        sky_box = QFrame()
        sky_box.setFixedHeight(360)
        sky_box.setStyleSheet("background:transparent;")
        sky_l = QVBoxLayout(sky_box)
        sky_l.setContentsMargins(0, 0, 0, 0)
        self.sky = SkyView()
        sky_l.addWidget(self.sky)
        root.addWidget(sky_box)

        # --- Tool badge overlay (fully floating overlay above layout) --- #
        from PyQt6.QtWidgets import QGraphicsOpacityEffect

        badge_path = resource_path(os.path.join("requirement_tool", "ui", "tool_badge.png"))
        if os.path.exists(badge_path):
            self.badge_label = QLabel(self)
            self.badge_pix = QPixmap(badge_path)

            # initial scaled pixmap (will be re-scaled inside reposition_badge)
            init_w = min(int(self.width() * 0.3), 460) or 420
            pix = self.badge_pix.scaledToWidth(init_w, Qt.TransformationMode.SmoothTransformation)
            self.badge_label.setPixmap(pix)

            # Make the label exactly the pixmap size so it is never clipped
            self.badge_label.setFixedSize(pix.size())
            self.badge_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.badge_label.setStyleSheet("background: transparent;")
            self.badge_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)

            # Parent it to the main widget (top-level overlay) and show
            self.badge_label.setParent(self)
            self.badge_label.show()

            # optional: add smooth fade effect
            fade = QGraphicsOpacityEffect(self.badge_label)
            fade.setOpacity(0.95)
            self.badge_label.setGraphicsEffect(fade)

            def reposition_badge():
                # Rescale pixmap to a fraction of window width (keeps it readable on resize)
                scaled_w = min(int(self.width() * 0.3), 460)
                pix = self.badge_pix.scaledToWidth(scaled_w, Qt.TransformationMode.SmoothTransformation)

                # Apply pixmap and make label match its size
                self.badge_label.setPixmap(pix)
                self.badge_label.setFixedSize(pix.size())

                # Center horizontally using the pixmap width (not the label old width)
                x = (self.width() - pix.width()) // 2

                # Anchor at bottom of sky_box and overlap by ~1/3 of badge height
                overlap_frac = 1/3

                # <-- tweak this to move the badge up (increase value to move higher)
                badge_vert_offset = 90 # pixels to nudge the badge upwards

                y = sky_box.geometry().bottom() - int(overlap_frac * pix.height()) - badge_vert_offset

                # Keep badge fully inside the window vertically (clamp)
                y = max(0, min(y, self.height() - pix.height()))

                self.badge_label.move(x, y)
                self.badge_label.raise_()


            # Override instance resizeEvent to reposition badge on main window resize
            old_resize = getattr(self, "resizeEvent", None)

            def _resizeEvent(event):
                # Call previous resizeEvent if any, then reposition badge
                if old_resize is not None:
                    try:
                        old_resize(event)
                    except Exception:
                        super(MainFrontUI, self).resizeEvent(event)
                else:
                    super(MainFrontUI, self).resizeEvent(event)
                reposition_badge()

            self.resizeEvent = _resizeEvent

            # initial placement
            reposition_badge()
        else:
            print(f"[Warning] tool_badge.png not found at {badge_path}")


        # Glossy section (no unsupported CSS)
        create_wrap = QFrame()
        create_wrap.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 rgba(210,210,210,180),
                    stop:1 rgba(140,140,140,210));
                border-radius: 24px;
                border: 1px solid rgba(255,255,255,0.25);
            }
        """)
        main_layout = QVBoxLayout(create_wrap)
        main_layout.setContentsMargins(40, 40, 40, 40)
        main_layout.setSpacing(20)

        # Create button
        self.create_btn = QPushButton("＋ Create Project")
        self.create_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.create_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #3a77d1, stop:1 #174f9e);
                color: white;
                border: none;
                border-radius: 26px;
                font-weight: 800;
                font-size: 17px;
                padding: 12px 30px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #5291ee, stop:1 #1d5fbb);
            }
        """)
        self.create_btn.clicked.connect(self.on_create_clicked)
        btn_box = QHBoxLayout()
        btn_box.addStretch(1)
        btn_box.addWidget(self.create_btn)
        btn_box.addStretch(1)
        main_layout.addLayout(btn_box)

        # Project cards scroll
        self.scroll_container = QWidget()
        self.scroll_layout = QHBoxLayout(self.scroll_container)
        self.scroll_layout.setContentsMargins(20, 20, 20, 20)
        self.scroll_layout.setSpacing(20)
        self.scroll_layout.addStretch(1)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setWidget(self.scroll_container)
        main_layout.addWidget(scroll)
        root.addWidget(create_wrap)

        # Footer
        footer = QFrame()
        footer.setFixedHeight(44)
        footer.setStyleSheet(f"background:{NAVY};")
        fl = QHBoxLayout(footer)
        fl.setContentsMargins(12, 0, 12, 0)
        foot = QLabel("© MACH Global Technologies — Version 1.0")
        foot.setStyleSheet("color:#e3f2fd;")
        fl.addStretch(1)
        fl.addWidget(foot)
        fl.addStretch(1)
        root.addWidget(footer)

        self.load_projects_from_storage()
    

    # ---------------- Existing logic unchanged below ---------------- #
    def load_projects_from_storage(self):
        if not os.path.exists(self.project_list_file):
            return
        try:
            with open(self.project_list_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            for name in data.keys():
                self.add_project_card(name)
        except Exception as e:
            print(f"Error loading projects: {e}")

    def save_project_to_storage(self, name):
        """Save project info."""
        try:
            if os.path.exists(self.project_list_file):
                with open(self.project_list_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = {}
            path = os.path.join(self.project_root, name)
            os.makedirs(path, exist_ok=True)
            data[name] = {
                "path": path,
                "created_at": datetime.datetime.now().isoformat(),
            }
            with open(self.project_list_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            print(f"Error saving project {name}: {e}")

    # ----------------- Existing UI Logic ---------------- #
    def on_create_clicked(self):
        name, ok = QInputDialog.getText(self, "Create Project", "Enter Project Name:")
        if ok and name.strip():
            name = name.strip()
            self.add_project_card(name)
            self.save_project_to_storage(name)

    # ... keep your existing add_project_card(), rename_project(), etc ...


    def add_project_card(self, name):
        # (same as your version)
        card = QFrame()
        card.setFixedSize(260, 140)
        card.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 rgba(34,78,155,0.85),
                    stop:1 rgba(70,130,200,0.85));
                border-radius: 18px;
                border: 1px solid rgba(255,255,255,0.25);
            }
            QFrame:hover { border: 1px solid rgba(255,255,255,0.45); }
        """)
        label = QLabel(name)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("color:white; font-weight:800; font-size:15px;")
        vbox = QVBoxLayout(card)
        vbox.addStretch(1)
        vbox.addWidget(label)
        vbox.addStretch(1)
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(24)
        shadow.setOffset(0, 8)
        shadow.setColor(Qt.GlobalColor.black)
        card.setGraphicsEffect(shadow)
        card.mouseDoubleClickEvent = lambda e, n=name: self.open_project(n)
        stretch = self.scroll_layout.takeAt(self.scroll_layout.count() - 1)
        self.scroll_layout.addWidget(card)
        self.scroll_layout.addItem(stretch)
        self.projects.append(name)
        self.project_cards.append(card)
        # Enable right-click menu on the card itself
        card.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        card.customContextMenuRequested.connect(
            lambda pos, n=name, c=card: self.show_project_context_menu(n, c, pos)
        )

        # Also enable it on the label (so right-clicking text works too)
        label.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        label.customContextMenuRequested.connect(
            lambda pos, n=name, c=card, w=label: self.show_project_context_menu(n, c, c.mapFromGlobal(w.mapToGlobal(pos)))
        )

    def open_project(self, name):
        """Open MainWindow for selected project."""
        try:
            # Keep a reference so it isn’t garbage-collected
            self.main_window = MainWindow()
            self.clear_all_tabs(delete_files=False)
            self.main_window.set_project_name(name)
            self.main_window.show()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open project '{name}': {e}")


    def show_project_context_menu(self, project_name, card, pos):
        menu = QMenu()
        act_rename = menu.addAction("Rename Project")
        act_delete = menu.addAction("Delete Project")
        act_clear = menu.addAction("Clear Contents")
        action = menu.exec(card.mapToGlobal(pos))
        if action == act_rename:
            self.rename_project(project_name, card)
        elif action == act_delete:
            self.delete_project(project_name, card)
        elif action == act_clear:
            self.clear_project_contents(project_name)

    def rename_project(self, old_name, card):
        new_name, ok = QInputDialog.getText(self, "Rename Project", "Enter new project name:", text=old_name)
        if not ok or not new_name.strip():
            return
        new_name = new_name.strip()
        old_path = os.path.join(self.project_root, old_name)
        new_path = os.path.join(self.project_root, new_name)
        try:
            os.rename(old_path, new_path)
            label = card.findChild(QLabel)
            if label:
                label.setText(new_name)

            # Update JSON
            if os.path.exists(self.project_list_file):
                with open(self.project_list_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                data[new_name] = data.pop(old_name)
                data[new_name]["path"] = new_path
                with open(self.project_list_file, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=4)

            # --- Update event bindings so double-click and right-click work instantly ---
            card.mouseDoubleClickEvent = lambda e, n=new_name: self.open_project(n)
            card.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            card.customContextMenuRequested.connect(
                lambda pos, n=new_name, c=card: self.show_project_context_menu(n, c, pos)
            )

            QMessageBox.information(self, "Renamed", f"Project renamed to {new_name}.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Rename failed: {e}")


    def delete_project(self, name, card):
        reply = QMessageBox.question(
            self, "Delete Project",
            f"Are you sure you want to permanently delete project '{name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            shutil.rmtree(os.path.join(self.project_root, name), ignore_errors=True)
            if os.path.exists(self.project_list_file):
                with open(self.project_list_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                data.pop(name, None)
                with open(self.project_list_file, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=4)
            card.setParent(None)
            self.scroll_layout.removeWidget(card)
            QMessageBox.information(self, "Deleted", f"Project '{name}' deleted successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Deletion failed: {e}")

    def clear_project_contents(self, name):
        """Right-click → Clear Contents.
        Deletes files from BOTH possible /projects folders, then resets the open window UI if needed.
        """
        from PyQt6.QtWidgets import QMessageBox
        import json, os, glob

        reply = QMessageBox.question(
            self, "Clear Contents",
            f"Clear all saved files and tabs for '{name}'?\n(This will not delete the project folder.)",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        try:
            # 1) Delete in both roots
            for proj_dir in both_project_dirs(name, __file__):
                if not proj_dir.exists():
                    continue

                # read metadata for copied sources
                recorded = set()
                meta_path = proj_dir / "project.json"
                try:
                    if meta_path.exists():
                        with open(meta_path, "r", encoding="utf-8") as f:
                            meta = json.load(f)
                        for k in ("excel_files", "word_files"):
                            for f in (meta.get(k) or []):
                                from pathlib import Path
                                recorded.add(Path(str(f)).name)
                except Exception as e:
                    print(f"[Warning] Could not parse {meta_path}: {e}")

                fixed = {"data.json", "data.xlsx", "output.docx", "project.json"}
                for fname in sorted(fixed | recorded):
                    p = proj_dir / fname
                    if p.exists():
                        try:
                            os.remove(str(p))
                        except Exception as e:
                            print(f"[Warning] Failed to remove {p}: {e}")

                # sweep any leftover Office docs
                try:
                    for pat in ("*.xls", "*.xlsx", "*.xlsm", "*.xlsb", "*.docx", "*.docm"):
                        for p in glob.glob(str(proj_dir / pat)):
                            try:
                                os.remove(p)
                            except Exception as e:
                                print(f"[Warning] Failed to remove {p}: {e}")
                except Exception as e:
                    print(f"[Warning] Fallback cleanup failed: {e}")

            # 2) If that project is open, reset its UI (no extra disk deletes here)
            try:
                if hasattr(self, "main_window") and getattr(self.main_window, "project_name", None) == name:
                    try:
                        self.main_window.clear_all_tabs(delete_files=False)
                    except TypeError:
                        self.main_window.clear_all_tabs()
            except Exception:
                pass

            QMessageBox.information(self, "Cleared", f"Contents for '{name}' removed.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to clear contents: {e}")



# ---------------- Run ---------------- #
if __name__ == "__main__":
    app = QApplication(sys.argv)
    front = MainFrontUI()
    front.show()
    sys.exit(app.exec())
