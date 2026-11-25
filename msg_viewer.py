"""
MSG Viewer - Local offline viewer and converter for Outlook .msg files.
"""

import base64
import hashlib
import html as html_lib
import json
import os
import re
import shutil
import sys
import threading
from datetime import datetime
from pathlib import Path

try:
    import extract_msg
except ImportError:
    print("Error: Missing 'extract-msg' package.")
    print("Install dependencies with: pip install extract-msg PyQt6")
    sys.exit(1)

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QAction, QColor, QDragEnterEvent, QDropEvent
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QStatusBar,
    QTabWidget,
    QTextBrowser,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

DARK_BG = "#1a1d23"
PANEL_BG = "#21252e"
SIDEBAR_BG = "#191c23"
ACCENT = "#4f8ef7"
ACCENT_HOVER = "#6ba3ff"
TEXT_PRIMARY = "#ffffff"
TEXT_SECONDARY = "#8892a4"
BORDER = "#2d3240"
SUCCESS = "#4caf7d"
DANGER = "#e05c5c"
ITEM_HOVER = "#2a2f3d"
ITEM_SELECT = "#2e3650"

STYLESHEET = f"""
QMainWindow, QWidget {{
    background-color: {DARK_BG};
    color: {TEXT_PRIMARY};
    font-family: 'Segoe UI', sans-serif;
    font-size: 13px;
}}

#sidebar {{
    background-color: {SIDEBAR_BG};
    border-right: 1px solid {BORDER};
}}

QListWidget {{
    background-color: transparent;
    border: none;
    outline: none;
    padding: 4px;
}}

QListWidget::item {{
    background-color: transparent;
    border-radius: 6px;
    padding: 10px 12px;
    margin: 1px 4px;
    color: {TEXT_PRIMARY};
}}

QListWidget::item:hover {{
    background-color: {ITEM_HOVER};
}}

QListWidget::item:selected {{
    background-color: {ITEM_SELECT};
    color: white;
    border-left: 3px solid {ACCENT};
}}

QPushButton {{
    background-color: {PANEL_BG};
    color: {TEXT_PRIMARY};
    border: 1px solid {BORDER};
    border-radius: 6px;
    padding: 6px 12px;
    font-weight: 500;
}}

QPushButton:hover {{
    background-color: {ITEM_HOVER};
    border-color: {ACCENT};
}}

QPushButton:disabled {{
    color: {TEXT_SECONDARY};
    border-color: {BORDER};
}}

QPushButton#primaryBtn {{
    background-color: {ACCENT};
    color: white;
    border: none;
    font-weight: 600;
}}

QPushButton#primaryBtn:hover {{
    background-color: {ACCENT_HOVER};
}}

QPushButton#primaryBtn:disabled {{
    background-color: {BORDER};
    color: {TEXT_SECONDARY};
}}

QPushButton#dangerBtn {{
    background-color: transparent;
    color: {DANGER};
    border: 1px solid {DANGER};
}}

QPushButton#dangerBtn:hover {{
    background-color: {DANGER};
    color: white;
}}

QPushButton#successBtn {{
    background-color: transparent;
    color: {SUCCESS};
    border: 1px solid {SUCCESS};
}}

QPushButton#successBtn:hover {{
    background-color: {SUCCESS};
    color: white;
}}

QPushButton#successBtn:disabled {{
    background-color: transparent;
    color: {TEXT_SECONDARY};
    border-color: {BORDER};
}}

QLineEdit {{
    background-color: {PANEL_BG};
    border: 1px solid {BORDER};
    border-radius: 6px;
    padding: 8px 12px;
    color: {TEXT_PRIMARY};
}}

QLineEdit:focus {{
    border-color: {ACCENT};
}}

QTabWidget::pane {{
    border: 1px solid {BORDER};
    background-color: {PANEL_BG};
}}

QTabBar::tab {{
    background-color: {DARK_BG};
    color: {TEXT_SECONDARY};
    padding: 8px 20px;
    border: 1px solid {BORDER};
    border-bottom: none;
    margin-right: 2px;
    border-radius: 6px 6px 0 0;
}}

QTabBar::tab:selected {{
    background-color: {PANEL_BG};
    color: {TEXT_PRIMARY};
    border-bottom: 2px solid {ACCENT};
}}

QTabBar::tab:hover {{
    background-color: {ITEM_HOVER};
    color: {TEXT_PRIMARY};
}}

QTextBrowser {{
    background-color: {PANEL_BG};
    border: none;
    color: {TEXT_PRIMARY};
    padding: 16px 20px;
}}

QTreeWidget {{
    background-color: transparent;
    border: none;
    outline: none;
}}

QTreeWidget::item {{
    padding: 6px 8px;
    border-radius: 4px;
}}

QTreeWidget::item:hover {{
    background-color: {ITEM_HOVER};
}}

QTreeWidget::item:selected {{
    background-color: {ITEM_SELECT};
    color: white;
}}

QTreeWidget QHeaderView::section {{
    background-color: {SIDEBAR_BG};
    color: {TEXT_SECONDARY};
    border: none;
    border-bottom: 1px solid {BORDER};
    padding: 8px;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 1px;
    text-transform: uppercase;
}}

QProgressBar {{
    background-color: {PANEL_BG};
    border: 1px solid {BORDER};
    border-radius: 4px;
    height: 6px;
}}

QProgressBar::chunk {{
    background-color: {ACCENT};
    border-radius: 4px;
}}

QStatusBar {{
    background-color: {SIDEBAR_BG};
    border-top: 1px solid {BORDER};
    color: {TEXT_SECONDARY};
    font-size: 12px;
}}

QMenu {{
    background-color: {PANEL_BG};
    border: 1px solid {BORDER};
    padding: 4px;
}}

QMenu::item {{
    padding: 6px 20px;
    border-radius: 4px;
}}

QMenu::item:selected {{
    background-color: {ITEM_SELECT};
}}

QScrollBar:vertical {{
    background-color: transparent;
    width: 8px;
    margin: 0;
}}

QScrollBar::handle:vertical {{
    background-color: {BORDER};
    border-radius: 4px;
    min-height: 30px;
}}

QScrollBar::handle:vertical:hover {{
    background-color: {TEXT_SECONDARY};
}}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}

QScrollBar:horizontal {{
    background-color: transparent;
    height: 8px;
}}

QScrollBar::handle:horizontal {{
    background-color: {BORDER};
    border-radius: 4px;
}}
"""

SYSTEM_DARK_COLORS = {
    "windowtext",
    "buttontext",
    "captiontext",
    "infotext",
    "menutext",
    "threaddarkshadow",
    "initial",
    "inherit",
    "currentcolor",
}


class MessageCache:
    """Persists opened .msg files locally so they remain accessible even if originals are deleted."""

    def __init__(self, cache_root: Path | None = None):
        if cache_root is not None:
            self.root = Path(cache_root)
        else:
            base_dir = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / ".msg_viewer"))
            self.root = base_dir / "MSGViewer" / "cache"
        self.files_dir = self.root / "files"
        self.index_path = self.root / "index.json"
        self._lock = threading.Lock()
        self.files_dir.mkdir(parents=True, exist_ok=True)

    def _read_index(self) -> dict:
        if not self.index_path.exists():
            return {}
        try:
            with open(self.index_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def _write_index(self, index: dict):
        tmp_path = self.index_path.with_suffix(".tmp")
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(index, f, indent=2, ensure_ascii=False)
        os.replace(tmp_path, self.index_path)

    @staticmethod
    def compute_hash(filepath: str) -> str:
        hasher = hashlib.sha256()
        with open(filepath, "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                hasher.update(chunk)
        return hasher.hexdigest()

    def store(self, filepath: str) -> dict:
        abs_path = os.path.abspath(filepath)
        file_hash = self.compute_hash(abs_path)
        cached_file = self.files_dir / f"{file_hash}.msg"

        with self._lock:
            if not cached_file.exists():
                shutil.copy2(abs_path, cached_file)

            index = self._read_index()
            existing = index.get(file_hash, {})
            entry = {
                "hash": file_hash,
                "filename": os.path.basename(abs_path),
                "original_path": abs_path,
                "cached_path": str(cached_file),
                "cached_at": existing.get("cached_at") or datetime.now().strftime("%Y-%m-%d %H:%M"),
            }
            index[file_hash] = entry
            self._write_index(index)
            return entry

    def list_entries(self) -> list[dict]:
        with self._lock:
            index = self._read_index()
            valid_entries = []
            dirty = False
            for file_hash, entry in list(index.items()):
                cached_path = entry.get("cached_path") or str(self.files_dir / f"{file_hash}.msg")
                if os.path.exists(cached_path):
                    entry["hash"] = file_hash
                    entry["cached_path"] = cached_path
                    valid_entries.append(entry)
                else:
                    index.pop(file_hash, None)
                    dirty = True
            if dirty:
                self._write_index(index)
            return valid_entries

    def remove(self, file_hash: str):
        with self._lock:
            index = self._read_index()
            entry = index.pop(file_hash, None)
            self._write_index(index)
            cached_file = Path(entry["cached_path"]) if entry and entry.get("cached_path") else (self.files_dir / f"{file_hash}.msg")
            if cached_file.exists():
                try:
                    cached_file.unlink()
                except OSError:
                    pass

    def clear(self):
        with self._lock:
            for f in self.files_dir.glob("*.msg"):
                try:
                    f.unlink()
                except OSError:
                    pass
            if self.index_path.exists():
                try:
                    self.index_path.unlink()
                except OSError:
                    pass


def _decode_bytes(data: bytes) -> str:
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16", errors="replace")
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig", errors="replace")

    meta_charset = None
    match = re.search(rb'charset\s*=\s*["\']?([a-zA-Z0-9._-]+)', data[:2048], flags=re.IGNORECASE)
    if match:
        meta_charset = match.group(1).decode("ascii", errors="ignore")

    for enc in ("utf-8", meta_charset, "cp1254", "cp1252"):
        if not enc:
            continue
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return data.decode("utf-8", errors="replace")


def _parse_css_rgb(raw_color: str) -> tuple[int, int, int] | None:
    val = re.sub(r"!important", "", raw_color, flags=re.IGNORECASE).strip().lower()
    if not val or val in ("transparent", "none"):
        return None
    if val in SYSTEM_DARK_COLORS:
        return (0, 0, 0)

    rgb_match = re.match(r"rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})", val)
    if rgb_match:
        return (
            min(255, int(rgb_match.group(1))),
            min(255, int(rgb_match.group(2))),
            min(255, int(rgb_match.group(3))),
        )

    qcolor = QColor(val)
    if qcolor.isValid():
        return (qcolor.red(), qcolor.green(), qcolor.blue())
    return None


def _ensure_light_text_color(match: re.Match) -> str:
    raw_value = match.group(1)
    rgb = _parse_css_rgb(raw_value)
    if rgb is None:
        return "color: #ffffff"
    luminance = (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) / 255.0
    if luminance < 0.62:
        return "color: #ffffff"
    return match.group(0)


def adapt_html_for_dark_mode(raw_html: str) -> str:
    """Normalizes dark/black fonts and light table backgrounds in MSG HTML for dark UI readability."""
    html = re.sub(r"<xml[^>]*>.*?</xml>", "", raw_html, flags=re.DOTALL | re.IGNORECASE)

    # Strip legacy HTML color/background attributes on layout and text elements
    html = re.sub(
        r"<(?:body|table|tr|td|th|div|p|span)\b[^>]*>",
        lambda m: re.sub(
            r"""\s+\b(?:bgcolor|text|link|vlink|alink)\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)""",
            "",
            m.group(0),
            flags=re.IGNORECASE,
        ),
        html,
        flags=re.IGNORECASE,
    )
    html = re.sub(
        r"(<font\b[^>]*?)\bcolor\s*=\s*(?:\"[^\"]*\"|'[^']*'|[^\s>]+)",
        r'\1color="#ffffff"',
        html,
        flags=re.IGNORECASE,
    )

    def _transform_css(css_text: str) -> str:
        css_text = re.sub(
            r"(?<![-\w])background(?:-color)?\s*:\s*[^;}\"']+",
            "background-color: transparent",
            css_text,
            flags=re.IGNORECASE,
        )
        css_text = re.sub(
            r"(?<![-\w])color\s*:\s*([^;}\"']+)",
            _ensure_light_text_color,
            css_text,
            flags=re.IGNORECASE,
        )
        return css_text

    html = re.sub(
        r"(<style[^>]*>)(.*?)(</style>)",
        lambda m: m.group(1) + _transform_css(m.group(2)) + m.group(3),
        html,
        flags=re.DOTALL | re.IGNORECASE,
    )
    html = re.sub(
        r'(\bstyle\s*=\s*)(["\'])(.*?)\2',
        lambda m: m.group(1) + m.group(2) + _transform_css(m.group(3)) + m.group(2),
        html,
        flags=re.DOTALL | re.IGNORECASE,
    )

    dark_override_css = (
        f"<style>"
        f"body, table, tr, td, th, div, p, span, li, ul, ol, font {{"
        f" background-color: transparent; color: {TEXT_PRIMARY};"
        f" font-family: 'Segoe UI', sans-serif; font-size: 13px; }}"
        f"a, a:link, a:visited, span.MsoHyperlink, span.MsoHyperlinkFollowed {{"
        f" color: {ACCENT_HOVER}; }}"
        f"</style>"
    )

    if re.search(r"</head>", html, flags=re.IGNORECASE):
        html = re.sub(r"</head>", dark_override_css + "</head>", html, count=1, flags=re.IGNORECASE)
    else:
        html = dark_override_css + html

    return html


def parse_msg(source_path: str, cache_entry: dict | None = None) -> dict:
    """Parses an Outlook .msg file and returns its structured representation."""
    original_path = cache_entry["original_path"] if cache_entry else source_path
    filename = cache_entry["filename"] if cache_entry else os.path.basename(source_path)
    cached_path = cache_entry["cached_path"] if cache_entry else source_path
    file_hash = cache_entry["hash"] if cache_entry else ""

    try:
        with extract_msg.openMsg(cached_path) as msg:
            body_html = getattr(msg, "htmlBody", None)
            body_plain = getattr(msg, "body", None) or ""

            if isinstance(body_html, bytes):
                body_html = _decode_bytes(body_html)
            if isinstance(body_plain, bytes):
                body_plain = _decode_bytes(body_plain)

            recipients = []
            try:
                for r in msg.recipients or []:
                    name = getattr(r, "name", "") or ""
                    addr = getattr(r, "email", "") or ""
                    recipients.append(f"{name} <{addr}>".strip(" <>") if name else addr)
            except Exception:
                pass

            to_str = getattr(msg, "to", "") or ", ".join(recipients)
            cc_str = getattr(msg, "cc", "") or ""
            bcc_str = getattr(msg, "bcc", "") or ""

            attachments = []
            try:
                for att in msg.attachments or []:
                    name = getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or "attachment"
                    data = getattr(att, "data", None) or b""
                    if isinstance(data, str):
                        data = data.encode("utf-8")
                    attachments.append({"name": name, "data": data, "size": len(data)})
            except Exception:
                pass

            date_val = getattr(msg, "date", None)
            date_str = ""
            if date_val:
                try:
                    date_str = date_val.strftime("%Y-%m-%d %H:%M") if hasattr(date_val, "strftime") else str(date_val)
                except Exception:
                    date_str = str(date_val)

            return {
                "hash": file_hash,
                "filepath": original_path,
                "cached_path": cached_path,
                "filename": filename,
                "subject": getattr(msg, "subject", "") or "(No Subject)",
                "sender": getattr(msg, "sender", "") or "",
                "to": to_str or "",
                "cc": cc_str,
                "bcc": bcc_str,
                "date": date_str,
                "body_html": body_html,
                "body_plain": body_plain,
                "attachments": attachments,
                "error": None,
            }
    except Exception as exc:
        return {
            "hash": file_hash,
            "filepath": original_path,
            "cached_path": cached_path,
            "filename": filename,
            "subject": filename,
            "sender": "",
            "to": "",
            "cc": "",
            "bcc": "",
            "date": "",
            "body_html": None,
            "body_plain": "",
            "attachments": [],
            "error": str(exc),
        }


def msg_to_eml(msg_data: dict) -> str:
    """Converts parsed MSG dictionary into an RFC 5322 EML string."""
    lines = [
        f"From: {msg_data['sender']}",
        f"To: {msg_data['to']}",
    ]
    if msg_data["cc"]:
        lines.append(f"Cc: {msg_data['cc']}")
    lines.extend([
        f"Subject: {msg_data['subject']}",
        f"Date: {msg_data['date']}",
        "MIME-Version: 1.0",
    ])

    atts = msg_data["attachments"]
    boundary = f"----=_Part_{hashlib.md5(msg_data['subject'].encode('utf-8', errors='replace')).hexdigest()}"

    if atts:
        lines.append(f'Content-Type: multipart/mixed; boundary="{boundary}"')
        lines.append("")
        lines.append(f"--{boundary}")

    if msg_data["body_html"]:
        lines.append('Content-Type: text/html; charset="utf-8"')
        lines.append("Content-Transfer-Encoding: 8bit")
        lines.append("")
        lines.append(msg_data["body_html"])
    else:
        lines.append('Content-Type: text/plain; charset="utf-8"')
        lines.append("Content-Transfer-Encoding: 8bit")
        lines.append("")
        lines.append(msg_data["body_plain"])

    if atts:
        for att in atts:
            lines.append(f"--{boundary}")
            lines.append(f'Content-Type: application/octet-stream; name="{att["name"]}"')
            lines.append(f'Content-Disposition: attachment; filename="{att["name"]}"')
            lines.append("Content-Transfer-Encoding: base64")
            lines.append("")
            if att["data"]:
                lines.append(base64.b64encode(att["data"]).decode("ascii"))
        lines.append(f"--{boundary}--")

    return "\n".join(lines)


def format_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024 * 1024):.1f} MB"


class LoadThread(QThread):
    progress = pyqtSignal(int, int)
    loaded = pyqtSignal(dict)
    finished_all = pyqtSignal()

    def __init__(self, items: list, cache: MessageCache, from_cache: bool = False):
        super().__init__()
        self.items = items
        self.cache = cache
        self.from_cache = from_cache

    def run(self):
        total = len(self.items)
        for idx, item in enumerate(self.items, 1):
            if self.from_cache:
                data = parse_msg(item["cached_path"], cache_entry=item)
            else:
                try:
                    entry = self.cache.store(item)
                    data = parse_msg(item, cache_entry=entry)
                except Exception as exc:
                    data = {
                        "hash": "",
                        "filepath": item,
                        "cached_path": item,
                        "filename": os.path.basename(item),
                        "subject": os.path.basename(item),
                        "sender": "",
                        "to": "",
                        "cc": "",
                        "bcc": "",
                        "date": "",
                        "body_html": None,
                        "body_plain": "",
                        "attachments": [],
                        "error": str(exc),
                    }
            self.loaded.emit(data)
            self.progress.emit(idx, total)
        self.finished_all.emit()


class DropListWidget(QListWidget):
    files_dropped = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setDragDropMode(QListWidget.DragDropMode.NoDragDrop)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        paths = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isdir(path):
                paths.extend(str(f) for f in Path(path).rglob("*.msg"))
            elif path.lower().endswith(".msg"):
                paths.append(path)
        if paths:
            self.files_dropped.emit(paths)


class MSGViewer(QMainWindow):
    def __init__(self, initial_files: list[str] | None = None, cache_root: Path | None = None):
        super().__init__()
        self.cache = MessageCache(cache_root=cache_root)
        self.messages: list[dict] = []
        self.current: dict | None = None
        self._pending_files = initial_files or []
        self._thread: LoadThread | None = None

        self.setWindowTitle("MSG Viewer")
        self.setMinimumSize(1100, 700)
        self.resize(1280, 780)
        self._build_ui()
        self._connect_signals()
        self._load_initial_state()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Left sidebar
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(310)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0)

        header = QWidget()
        header.setStyleSheet(f"background-color: {SIDEBAR_BG};")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(12, 12, 12, 8)
        header_layout.setSpacing(8)

        title_lbl = QLabel("MSG VIEWER")
        title_lbl.setStyleSheet(f"color: {ACCENT}; font-size: 13px; font-weight: 700; letter-spacing: 3px;")
        header_layout.addWidget(title_lbl)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(6)
        self.btn_add_files = QPushButton("＋ Files")
        self.btn_add_folder = QPushButton("📁 Folder")
        self.btn_clear = QPushButton("✕ Clear")
        self.btn_clear.setObjectName("dangerBtn")
        for btn in (self.btn_add_files, self.btn_add_folder, self.btn_clear):
            btn.setFixedHeight(32)
            btn_row.addWidget(btn)
        header_layout.addLayout(btn_row)

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("🔍  Search messages...")
        self.search_box.setFixedHeight(34)
        header_layout.addWidget(self.search_box)

        sidebar_layout.addWidget(header)

        self.stat_lbl = QLabel("No messages")
        self.stat_lbl.setStyleSheet(f"color: {TEXT_SECONDARY}; font-size: 11px; padding: 4px 16px;")
        sidebar_layout.addWidget(self.stat_lbl)

        self.msg_list = DropListWidget()
        self.msg_list.setStyleSheet("QListWidget { font-size: 12px; }")
        self.msg_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        sidebar_layout.addWidget(self.msg_list)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(4)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.hide()
        sidebar_layout.addWidget(self.progress_bar)

        root.addWidget(sidebar)

        # Right panel
        content = QWidget()
        content.setStyleSheet(f"background-color: {PANEL_BG};")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        toolbar = QWidget()
        toolbar.setStyleSheet(f"background-color: {DARK_BG}; border-bottom: 1px solid {BORDER};")
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(16, 8, 16, 8)
        toolbar_layout.setSpacing(8)

        self.btn_save_msg = QPushButton("📥 Save .MSG")
        self.btn_eml = QPushButton("💾 Export EML")
        self.btn_save_atts = QPushButton("📎 Save Attachments")
        self.btn_eml_all = QPushButton("📦 Export All EML")
        self.btn_open_cache = QPushButton("📂 Cache Folder")

        self.btn_eml.setObjectName("successBtn")
        self.btn_save_atts.setObjectName("primaryBtn")

        for btn in (self.btn_save_msg, self.btn_eml, self.btn_save_atts, self.btn_eml_all):
            btn.setFixedHeight(34)
            btn.setEnabled(False)
            toolbar_layout.addWidget(btn)

        self.btn_open_cache.setFixedHeight(34)
        toolbar_layout.addWidget(self.btn_open_cache)

        toolbar_layout.addStretch()

        self.info_lbl = QLabel("")
        self.info_lbl.setStyleSheet(f"color: {TEXT_SECONDARY}; font-size: 11px;")
        toolbar_layout.addWidget(self.info_lbl)

        content_layout.addWidget(toolbar)

        # Tabs (Content + Attachments only)
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        content_layout.addWidget(self.tabs)

        # Tab 1: Content (General Information + Message Body)
        content_tab = QWidget()
        content_tab_layout = QVBoxLayout(content_tab)
        content_tab_layout.setContentsMargins(0, 0, 0, 0)
        content_tab_layout.setSpacing(0)

        self.info_panel = QWidget()
        self.info_panel.setStyleSheet(f"background-color: {DARK_BG}; border-bottom: 1px solid {BORDER};")
        info_panel_layout = QVBoxLayout(self.info_panel)
        info_panel_layout.setContentsMargins(20, 14, 20, 14)
        info_panel_layout.setSpacing(8)

        self.lbl_subject = QLabel("")
        self.lbl_subject.setStyleSheet(f"color: {TEXT_PRIMARY}; font-size: 16px; font-weight: 600;")
        self.lbl_subject.setWordWrap(True)
        self.lbl_subject.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        info_panel_layout.addWidget(self.lbl_subject)

        self.meta_form = QFormLayout()
        self.meta_form.setContentsMargins(0, 0, 0, 0)
        self.meta_form.setHorizontalSpacing(14)
        self.meta_form.setVerticalSpacing(4)
        self.meta_form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)

        self.lbl_from = self._create_meta_value_label()
        self.lbl_to = self._create_meta_value_label()
        self.lbl_cc = self._create_meta_value_label()
        self.lbl_bcc = self._create_meta_value_label()
        self.lbl_date = self._create_meta_value_label()
        self.lbl_file = self._create_meta_value_label()
        self.lbl_att_summary = self._create_meta_value_label()

        self._meta_rows = {
            "from": (self._create_meta_key_label("From:"), self.lbl_from),
            "to": (self._create_meta_key_label("To:"), self.lbl_to),
            "cc": (self._create_meta_key_label("CC:"), self.lbl_cc),
            "bcc": (self._create_meta_key_label("BCC:"), self.lbl_bcc),
            "date": (self._create_meta_key_label("Date:"), self.lbl_date),
            "file": (self._create_meta_key_label("File:"), self.lbl_file),
            "atts": (self._create_meta_key_label("Attachments:"), self.lbl_att_summary),
        }
        for key_lbl, val_lbl in self._meta_rows.values():
            self.meta_form.addRow(key_lbl, val_lbl)

        info_panel_layout.addLayout(self.meta_form)
        content_tab_layout.addWidget(self.info_panel)

        self.body_browser = QTextBrowser()
        self.body_browser.setOpenExternalLinks(True)
        content_tab_layout.addWidget(self.body_browser, stretch=1)

        self.tabs.addTab(content_tab, "📧 Content")

        # Tab 2: Attachments
        att_widget = QWidget()
        att_layout = QVBoxLayout(att_widget)
        att_layout.setContentsMargins(16, 16, 16, 16)

        self.att_tree = QTreeWidget()
        self.att_tree.setHeaderLabels(["FILE NAME", "SIZE"])
        self.att_tree.setColumnWidth(0, 440)
        self.att_tree.setRootIsDecorated(False)
        att_layout.addWidget(self.att_tree)

        self.btn_save_sel = QPushButton("Save Selected Attachment")
        self.btn_save_sel.setObjectName("primaryBtn")
        self.btn_save_sel.setFixedHeight(36)
        self.btn_save_sel.setEnabled(False)
        att_layout.addWidget(self.btn_save_sel)

        self.tabs.addTab(att_widget, "📎 Attachments (0)")

        # Welcome view
        self.welcome = QWidget()
        welcome_layout = QVBoxLayout(self.welcome)
        welcome_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        w_icon = QLabel("📨")
        w_icon.setStyleSheet("font-size: 56px;")
        w_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        w_text = QLabel("Click '+ Files' or drag and drop .msg files here\nOpened messages are automatically cached for offline access")
        w_text.setStyleSheet(f"color: {TEXT_SECONDARY}; font-size: 14px;")
        w_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        welcome_layout.addWidget(w_icon)
        welcome_layout.addWidget(w_text)
        content_layout.addWidget(self.welcome)

        self.tabs.hide()
        root.addWidget(content)

        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("Ready — Drag & drop .msg files or click '+ Files'")

    @staticmethod
    def _create_meta_key_label(text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setStyleSheet(f"color: {TEXT_SECONDARY}; font-size: 12px; font-weight: 600;")
        return lbl

    @staticmethod
    def _create_meta_value_label() -> QLabel:
        lbl = QLabel("")
        lbl.setStyleSheet(f"color: {TEXT_PRIMARY}; font-size: 12px;")
        lbl.setWordWrap(True)
        lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        return lbl

    def _set_meta_row(self, key: str, value: str, always_show: bool = False):
        key_lbl, val_lbl = self._meta_rows[key]
        visible = bool(value) or always_show
        val_lbl.setText(value or "—")
        key_lbl.setVisible(visible)
        val_lbl.setVisible(visible)

    def _connect_signals(self):
        self.btn_add_files.clicked.connect(self._add_files)
        self.btn_add_folder.clicked.connect(self._add_folder)
        self.btn_clear.clicked.connect(self._clear_all)
        self.msg_list.currentItemChanged.connect(self._on_select)
        self.msg_list.files_dropped.connect(self._load_paths)
        self.msg_list.customContextMenuRequested.connect(self._show_list_context_menu)
        self.search_box.textChanged.connect(self._filter_list)
        self.btn_save_msg.clicked.connect(self._save_msg_current)
        self.btn_eml.clicked.connect(self._export_eml_current)
        self.btn_save_atts.clicked.connect(lambda: self._save_attachments(self.current))
        self.btn_eml_all.clicked.connect(self._export_eml_all)
        self.btn_open_cache.clicked.connect(self._open_cache_folder)
        self.btn_save_sel.clicked.connect(self._save_selected_att)
        self.att_tree.itemDoubleClicked.connect(lambda *_: self._save_selected_att())
        self.att_tree.itemSelectionChanged.connect(
            lambda: self.btn_save_sel.setEnabled(bool(self.att_tree.selectedItems()))
        )

    def _load_initial_state(self):
        cached_entries = self.cache.list_entries()
        if cached_entries:
            self._start_loader(cached_entries, from_cache=True)
        elif self._pending_files:
            files = self._pending_files
            self._pending_files = []
            self._load_paths(files)

    def _add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select MSG Files", "", "Outlook MSG (*.msg);;All Files (*)"
        )
        if paths:
            self._load_paths(paths)

    def _add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            paths = [str(p) for p in Path(folder).rglob("*.msg")]
            if paths:
                self._load_paths(paths)
            else:
                self.status.showMessage("No .msg files found in the selected folder.")

    def _load_paths(self, paths: list[str]):
        existing_hashes = {m["hash"] for m in self.messages if m.get("hash")}
        existing_paths = {os.path.abspath(m["filepath"]) for m in self.messages}
        new_paths = []
        first_existing_idx = None

        for path in paths:
            if not os.path.exists(path):
                continue
            abs_path = os.path.abspath(path)
            if abs_path in existing_paths:
                first_existing_idx = next(
                    (i for i, m in enumerate(self.messages) if os.path.abspath(m["filepath"]) == abs_path),
                    first_existing_idx,
                )
                continue
            try:
                file_hash = MessageCache.compute_hash(abs_path)
            except OSError:
                file_hash = ""
            if file_hash and file_hash in existing_hashes:
                first_existing_idx = next(
                    (i for i, m in enumerate(self.messages) if m.get("hash") == file_hash),
                    first_existing_idx,
                )
                continue
            if file_hash:
                existing_hashes.add(file_hash)
            new_paths.append(abs_path)

        if not new_paths:
            if first_existing_idx is not None:
                self._select_message_by_index(first_existing_idx)
                self.status.showMessage("Selected file is already loaded from cache.")
            return

        self._start_loader(new_paths, from_cache=False)

    def _start_loader(self, items: list, from_cache: bool):
        self.progress_bar.setMaximum(len(items))
        self.progress_bar.setValue(0)
        self.progress_bar.show()
        self.btn_add_files.setEnabled(False)
        self.btn_add_folder.setEnabled(False)

        self._thread = LoadThread(items, self.cache, from_cache=from_cache)
        self._thread.loaded.connect(self._on_msg_loaded)
        self._thread.progress.connect(lambda c, _t: self.progress_bar.setValue(c))
        self._thread.finished_all.connect(self._on_load_done)
        self._thread.start()

    def _on_msg_loaded(self, data: dict):
        if data.get("hash") and any(m.get("hash") == data["hash"] for m in self.messages):
            return
        self.messages.append(data)
        item = QListWidgetItem()
        subject = data["subject"] or data["filename"]
        sender = data["sender"] or "Unknown sender"
        date = data["date"] or ""
        suffix = f"  ·  {date}" if date else ""
        item.setText(f"{subject}\n{sender}{suffix}")
        item.setData(Qt.ItemDataRole.UserRole, len(self.messages) - 1)
        if data["error"]:
            item.setForeground(QColor(DANGER))
        self.msg_list.addItem(item)
        self._update_stats()

    def _on_load_done(self):
        self.progress_bar.hide()
        self.btn_add_files.setEnabled(True)
        self.btn_add_folder.setEnabled(True)
        count = self.msg_list.count()
        self.btn_eml_all.setEnabled(count > 0)
        self.status.showMessage(f"Loaded {count} message{'s' if count != 1 else ''} (cached locally).")

        if self._pending_files:
            pending = self._pending_files
            self._pending_files = []
            self._load_paths(pending)
            return

        if count > 0 and self.msg_list.currentRow() < 0:
            self.msg_list.setCurrentRow(count - 1)

    def _select_message_by_index(self, msg_index: int):
        for row in range(self.msg_list.count()):
            item = self.msg_list.item(row)
            if item.data(Qt.ItemDataRole.UserRole) == msg_index:
                self.msg_list.setCurrentRow(row)
                break

    def _show_list_context_menu(self, pos):
        item = self.msg_list.itemAt(pos)
        if not item:
            return
        self.msg_list.setCurrentItem(item)
        menu = QMenu(self)

        act_save_msg = QAction("Save Copy as .MSG...", self)
        act_export_eml = QAction("Export as .EML...", self)
        act_remove = QAction("Remove from Cache", self)

        act_save_msg.triggered.connect(self._save_msg_current)
        act_export_eml.triggered.connect(self._export_eml_current)
        act_remove.triggered.connect(lambda: self._remove_item(item))

        menu.addAction(act_save_msg)
        menu.addAction(act_export_eml)
        menu.addSeparator()
        menu.addAction(act_remove)
        menu.exec(self.msg_list.mapToGlobal(pos))

    def _remove_item(self, item: QListWidgetItem):
        idx = item.data(Qt.ItemDataRole.UserRole)
        msg = self.messages[idx]
        if msg.get("hash"):
            self.cache.remove(msg["hash"])

        del self.messages[idx]
        self.msg_list.takeItem(self.msg_list.row(item))

        # Re-index remaining list items
        for row in range(self.msg_list.count()):
            list_item = self.msg_list.item(row)
            old_idx = list_item.data(Qt.ItemDataRole.UserRole)
            if old_idx > idx:
                list_item.setData(Qt.ItemDataRole.UserRole, old_idx - 1)

        if not self.messages:
            self.current = None
            self._show_welcome()
        self._update_stats()
        self.status.showMessage(f"Removed '{msg['filename']}' from cache.")

    def _clear_all(self):
        if not self.messages:
            return
        reply = QMessageBox.question(
            self,
            "Clear Cache",
            "Remove all messages from the list and delete their cached copies?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.cache.clear()
            self.messages.clear()
            self.msg_list.clear()
            self.current = None
            self._show_welcome()
            self._update_stats()
            self.status.showMessage("All messages and cached copies cleared.")

    def _open_cache_folder(self):
        cache_dir = str(self.cache.files_dir)
        try:
            if sys.platform == "win32":
                os.startfile(cache_dir)
            else:
                QMessageBox.information(self, "Cache Location", cache_dir)
        except Exception as exc:
            QMessageBox.warning(self, "Cache Folder", f"{cache_dir}\n\n{exc}")

    def _update_stats(self):
        n = len(self.messages)
        atts = sum(len(m["attachments"]) for m in self.messages)
        self.stat_lbl.setText(f"{n} message{'s' if n != 1 else ''}  ·  {atts} attachment{'s' if atts != 1 else ''}" if n else "No messages")
        self.btn_eml_all.setEnabled(n > 0)

    def _filter_list(self, query: str):
        query = query.lower()
        for i in range(self.msg_list.count()):
            item = self.msg_list.item(i)
            idx = item.data(Qt.ItemDataRole.UserRole)
            msg = self.messages[idx]
            match = (
                query in (msg["subject"] or "").lower()
                or query in (msg["sender"] or "").lower()
                or query in (msg["to"] or "").lower()
                or query in (msg["body_plain"] or "").lower()
                or query in (msg["filename"] or "").lower()
            )
            item.setHidden(not match)

    def _on_select(self, current, _previous):
        if not current:
            return
        idx = current.data(Qt.ItemDataRole.UserRole)
        self.current = self.messages[idx]
        self._display(self.current)

    def _display(self, msg: dict):
        self.welcome.hide()
        self.tabs.show()

        atts = msg["attachments"]
        total_att_size = sum(a["size"] for a in atts)
        original_exists = os.path.exists(msg["filepath"])
        file_info = (
            f"{msg['filepath']}  (Cached)"
            if original_exists
            else f"{msg['filename']}  (Loaded from cache — original file removed)"
        )
        att_summary = f"{len(atts)} file{'s' if len(atts) != 1 else ''} ({format_size(total_att_size)})" if atts else "None"

        self.lbl_subject.setText(msg["subject"] or "(No Subject)")
        self._set_meta_row("from", msg["sender"], always_show=True)
        self._set_meta_row("to", msg["to"], always_show=True)
        self._set_meta_row("cc", msg["cc"], always_show=False)
        self._set_meta_row("bcc", msg["bcc"], always_show=False)
        self._set_meta_row("date", msg["date"], always_show=True)
        self._set_meta_row("file", file_info, always_show=True)
        self._set_meta_row("atts", att_summary, always_show=True)

        can_save_msg = os.path.exists(msg.get("cached_path", ""))
        self.btn_save_msg.setEnabled(can_save_msg)
        self.btn_eml.setEnabled(True)
        self.btn_eml_all.setEnabled(bool(self.messages))
        self.btn_save_atts.setEnabled(bool(atts))

        if msg["error"]:
            escaped_err = html_lib.escape(msg["error"])
            self.body_browser.setHtml(
                f"<div style='color:{DANGER}; font-family:monospace; padding:12px;'>"
                f"<b>Failed to open file:</b><br>{escaped_err}</div>"
            )
        elif msg["body_html"]:
            self.body_browser.setHtml(adapt_html_for_dark_mode(msg["body_html"]))
        else:
            plain_escaped = html_lib.escape(msg["body_plain"] or "(No content)")
            self.body_browser.setHtml(
                f"<pre style='color:{TEXT_PRIMARY}; font-family:\"Segoe UI\",sans-serif; "
                f"font-size:13px; white-space:pre-wrap; margin:0;'>{plain_escaped}</pre>"
            )

        self.att_tree.clear()
        self.tabs.setTabText(1, f"📎 Attachments ({len(atts)})")
        for att_idx, att in enumerate(atts):
            tree_item = QTreeWidgetItem([att["name"], format_size(att["size"])])
            tree_item.setData(0, Qt.ItemDataRole.UserRole, att_idx)
            self.att_tree.addTopLevelItem(tree_item)

        self.info_lbl.setText(f"{len(atts)} attachment{'s' if len(atts) != 1 else ''}" if atts else "")
        self.status.showMessage(f"Viewing: {msg['filename']}")

    def _show_welcome(self):
        self.tabs.hide()
        self.welcome.show()
        self.info_lbl.setText("")
        for btn in (self.btn_save_msg, self.btn_eml, self.btn_save_atts, self.btn_eml_all):
            btn.setEnabled(False)

    def _save_msg_current(self):
        if not self.current:
            return
        cached_path = self.current.get("cached_path", "")
        if not os.path.exists(cached_path):
            QMessageBox.warning(self, "Not Found", "Cached .msg file could not be found.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save MSG File", self.current["filename"], "Outlook MSG (*.msg)"
        )
        if path:
            try:
                shutil.copy2(cached_path, path)
                self.status.showMessage(f"Saved MSG copy: {path}")
            except Exception as exc:
                QMessageBox.critical(self, "Error", str(exc))

    def _export_eml_current(self):
        if not self.current:
            return
        default_name = os.path.splitext(self.current["filename"])[0] + ".eml"
        path, _ = QFileDialog.getSaveFileName(self, "Export as EML", default_name, "EML (*.eml)")
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(msg_to_eml(self.current))
                self.status.showMessage(f"Exported: {path}")
            except Exception as exc:
                QMessageBox.critical(self, "Error", str(exc))

    def _export_eml_all(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Destination Folder for EML Files")
        if not folder:
            return
        success = 0
        for msg in self.messages:
            try:
                safe_name = re.sub(r'[<>:"/\\|?*]', "_", msg["filename"])
                out = os.path.join(folder, os.path.splitext(safe_name)[0] + ".eml")
                if os.path.exists(out):
                    base, ext = os.path.splitext(out)
                    out = f"{base}_{success}{ext}"
                with open(out, "w", encoding="utf-8") as f:
                    f.write(msg_to_eml(msg))
                success += 1
            except Exception:
                pass
        QMessageBox.information(
            self, "Export Complete", f"Converted {success} of {len(self.messages)} file(s) to EML."
        )
        self.status.showMessage(f"Batch export finished: {folder}")

    def _save_attachments(self, msg: dict | None):
        if not msg or not msg["attachments"]:
            return
        folder = QFileDialog.getExistingDirectory(self, "Select Folder to Save Attachments")
        if not folder:
            return
        saved = 0
        for att in msg["attachments"]:
            try:
                safe_name = re.sub(r'[<>:"/\\|?*]', "_", att["name"])
                out = os.path.join(folder, safe_name)
                with open(out, "wb") as f:
                    f.write(att["data"])
                saved += 1
            except Exception:
                pass
        self.status.showMessage(f"Saved {saved} attachment(s) to {folder}")

    def _save_selected_att(self):
        if not self.current:
            return
        selected = self.att_tree.selectedItems()
        if not selected:
            return
        att_idx = selected[0].data(0, Qt.ItemDataRole.UserRole)
        atts = self.current["attachments"]
        if att_idx is None or not (0 <= att_idx < len(atts)):
            return
        att = atts[att_idx]
        path, _ = QFileDialog.getSaveFileName(self, "Save Attachment", att["name"])
        if path:
            try:
                with open(path, "wb") as f:
                    f.write(att["data"])
                self.status.showMessage(f"Saved attachment: {path}")
            except Exception as exc:
                QMessageBox.critical(self, "Error", str(exc))


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("MSG Viewer")
    app.setStyleSheet(STYLESHEET)

    cli_files = [a for a in sys.argv[1:] if a.lower().endswith(".msg") and os.path.exists(a)]
    window = MSGViewer(initial_files=cli_files)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
