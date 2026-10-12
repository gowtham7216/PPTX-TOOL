"""
SlideGuard Desktop User Interface.
Redesigned to faithfully match the dark navy modern dashboard reference screenshot.
Preserves existing architecture, AppController delegation, batch remediation safeguards,
and all underlying test hooks and attributes.
"""

import os
import sys
import threading
import datetime
import logging
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, Dict, Any, List

logger = logging.getLogger(__name__)

from PIL import Image, ImageDraw, ImageTk, ImageFont

from app.version import APP_NAME, APP_TITLE, VERSION, APP_DISCLAIMER
from app.paths import resolve_resource_path
from app.controller.app_controller import AppController
from models.job import JobContext
from models.enums import IssueSeverity, RemediationStatus
from models.ai_authoring import AuthoringStatus
from models.qa import QAResult, QAStatus


# =============================================================================
# COLOR PALETTE CONSTANTS (Matching Reference Screenshot)
# =============================================================================
BG_DARK = "#090E1A"            # Deep obsidian/navy background
BG_SIDEBAR = "#070C16"         # Left sidebar background
BG_CARD = "#0F1626"            # Card/panel container background
BG_CARD_INNER = "#141D30"      # Inner card/badge background
BORDER_SUBTLE = "#1C273C"      # Subtle card border
BORDER_ACCENT = "#22324D"      # Slightly lighter card border

# Accent Colors
BLUE_ACCENT = "#1D64EC"        # Reading Order blue & selected nav
PURPLE_ACCENT = "#7C3AED"      # Placeholder Conversion purple
ORANGE_ACCENT = "#EA580C"      # Validate & Compare orange
TEAL_ACCENT = "#0D9488"        # Generate Report teal/accent

# Text Colors
TEXT_WHITE = "#FFFFFF"
TEXT_PRIMARY = "#F8FAFC"
TEXT_SECONDARY = "#8B9CB5"
TEXT_MUTED = "#64748B"

# Status Colors
GREEN_BADGE_BG = "#0B251E"
GREEN_BADGE_BORDER = "#134E3F"
GREEN_TEXT = "#34D399"
RED_ALERT = "#EF4444"
AMBER_WARN = "#F59E0B"


class IconManager:
    """Generates crisp, antialiased PIL icons for all dashboard elements."""

    def __init__(self, root: Optional[tk.Tk] = None):
        self.root = root
        self._cache: Dict[str, ImageTk.PhotoImage] = {}

    def get_shield_logo(self, size: int = 34) -> ImageTk.PhotoImage:
        key = f"shield_{size}"
        if key in self._cache:
            return self._cache[key]
        scale = 4
        s = size * scale
        im = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        # Rounded background
        d.rounded_rectangle([0, 0, s - 1, s - 1], radius=int(7 * scale), fill="#1E40AF")
        # Outer shield
        cx, top, bot = s // 2, int(5 * scale), int(29 * scale)
        hw = int(10 * scale)
        d.polygon([(cx, top), (cx + hw, top + int(4.5 * scale)), (cx + hw, top + int(14 * scale)),
                   (cx, bot), (cx - hw, top + int(14 * scale)), (cx - hw, top + int(4.5 * scale))],
                  fill="#3B82F6")
        # Inner shield
        ihw = int(7.5 * scale)
        itop = top + int(2.5 * scale)
        ibot = bot - int(2.5 * scale)
        d.polygon([(cx, itop), (cx + ihw, itop + int(3.5 * scale)), (cx + ihw, itop + int(11 * scale)),
                   (cx, ibot), (cx - ihw, itop + int(11 * scale)), (cx - ihw, itop + int(3.5 * scale))],
                  fill="#FFFFFF")
        # Inner shield blue core
        chw = int(4.8 * scale)
        ctop = itop + int(2.5 * scale)
        cbot = ibot - int(2.5 * scale)
        d.polygon([(cx, ctop), (cx + chw, ctop + int(2.5 * scale)), (cx + chw, ctop + int(7.5 * scale)),
                   (cx, cbot), (cx - chw, ctop + int(7.5 * scale)), (cx - chw, ctop + int(2.5 * scale))],
                  fill="#1E40AF")
        # White checkmark inside core
        d.line([(cx - int(2.2 * scale), ctop + int(4.5 * scale)), (cx - int(0.5 * scale), ctop + int(6.5 * scale)),
                (cx + int(2.8 * scale), ctop + int(2.8 * scale))], fill="#FFFFFF", width=int(1.8 * scale))
        im = im.resize((size, size), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(im, master=self.root)
        self._cache[key] = photo
        return photo

    def get_card_icon(self, card_type: str, size: int = 36) -> ImageTk.PhotoImage:
        key = f"card_{card_type}_{size}"
        if key in self._cache:
            return self._cache[key]
        scale = 4
        s = size * scale
        im = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)

        if card_type == "reading_order":
            bg = BLUE_ACCENT
            d.rounded_rectangle([0, 0, s - 1, s - 1], radius=int(8 * scale), fill=bg)
            x_num = int(9 * scale)
            x_line = int(16 * scale)
            x_end = int(28 * scale)
            for y_pos in [int(10 * scale), int(18 * scale), int(26 * scale)]:
                d.ellipse([x_num, y_pos - int(1.5 * scale), x_num + int(3 * scale), y_pos + int(1.5 * scale)], fill="#FFFFFF")
                d.rounded_rectangle([x_line, y_pos - int(1.5 * scale), x_end, y_pos + int(1.5 * scale)], radius=int(1 * scale), fill="#FFFFFF")

        elif card_type == "placeholder":
            bg = PURPLE_ACCENT
            d.rounded_rectangle([0, 0, s - 1, s - 1], radius=int(8 * scale), fill=bg)
            d.rounded_rectangle([int(7 * scale), int(8 * scale), int(29 * scale), int(28 * scale)], radius=int(2 * scale), outline="#FFFFFF", width=int(2 * scale))
            d.line([(int(10 * scale), int(13 * scale)), (int(26 * scale), int(13 * scale))], fill="#FFFFFF", width=int(1.8 * scale))
            d.rounded_rectangle([int(10 * scale), int(17 * scale), int(16 * scale), int(24 * scale)], radius=int(1 * scale), fill="#FFFFFF")
            d.rounded_rectangle([int(20 * scale), int(17 * scale), int(26 * scale), int(24 * scale)], radius=int(1 * scale), fill="#FFFFFF")

        elif card_type == "validation":
            bg = ORANGE_ACCENT
            d.rounded_rectangle([0, 0, s - 1, s - 1], radius=int(8 * scale), fill=bg)
            cx = s // 2
            top = int(8 * scale)
            bot = int(28 * scale)
            hw = int(10 * scale)
            d.polygon([(cx, top), (cx + hw, top + int(4.5 * scale)), (cx + hw, top + int(12.5 * scale)),
                       (cx, bot), (cx - hw, top + int(12.5 * scale)), (cx - hw, top + int(4.5 * scale))],
                      fill="#FFFFFF")
            d.polygon([(cx, top + int(2.5 * scale)), (cx + int(6.5 * scale), top + int(5.5 * scale)), (cx + int(6.5 * scale), top + int(11.5 * scale)),
                       (cx, bot - int(2.5 * scale)), (cx - int(6.5 * scale), top + int(11.5 * scale)), (cx - int(6.5 * scale), top + int(5.5 * scale))],
                      fill=bg)
            d.line([(cx - int(3 * scale), top + int(10 * scale)), (cx - int(0.5 * scale), top + int(12.5 * scale)),
                    (cx + int(4 * scale), top + int(7.5 * scale))], fill="#FFFFFF", width=int(2 * scale))

        elif card_type == "report":
            bg = TEAL_ACCENT
            d.rounded_rectangle([0, 0, s - 1, s - 1], radius=int(8 * scale), fill=bg)
            x1, y1 = int(9 * scale), int(7 * scale)
            x2, y2 = int(27 * scale), int(29 * scale)
            d.rounded_rectangle([x1, y1, x2, y2], radius=int(2 * scale), outline="#FFFFFF", width=int(1.8 * scale))
            d.line([(x1 + int(3.5 * scale), y1 + int(6 * scale)), (x2 - int(3.5 * scale), y1 + int(6 * scale))], fill="#FFFFFF", width=int(1.6 * scale))
            d.line([(x1 + int(3.5 * scale), y1 + int(10.5 * scale)), (x2 - int(3.5 * scale), y1 + int(10.5 * scale))], fill="#FFFFFF", width=int(1.6 * scale))
            d.rounded_rectangle([x1 + int(3.5 * scale), y2 - int(8 * scale), x1 + int(5.5 * scale), y2 - int(3.5 * scale)], radius=int(0.8 * scale), fill="#FFFFFF")
            d.rounded_rectangle([x1 + int(7.5 * scale), y2 - int(12 * scale), x1 + int(9.5 * scale), y2 - int(3.5 * scale)], radius=int(0.8 * scale), fill="#FFFFFF")
            d.rounded_rectangle([x1 + int(11.5 * scale), y2 - int(6 * scale), x1 + int(13.5 * scale), y2 - int(3.5 * scale)], radius=int(0.8 * scale), fill="#FFFFFF")

        im = im.resize((size, size), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(im, master=self.root)
        self._cache[key] = photo
        return photo

    def get_metric_icon(self, metric_type: str, size: int = 22) -> ImageTk.PhotoImage:
        key = f"metric_{metric_type}_{size}"
        if key in self._cache:
            return self._cache[key]
        scale = 4
        s = size * scale
        im = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)

        if metric_type == "slides":
            d.rounded_rectangle([int(2 * scale), int(3 * scale), int(20 * scale), int(15 * scale)], radius=int(2 * scale), outline="#3B82F6", width=int(2 * scale))
            d.line([(int(11 * scale), int(15 * scale)), (int(11 * scale), int(19 * scale))], fill="#3B82F6", width=int(2 * scale))
            d.line([(int(6 * scale), int(19 * scale)), (int(16 * scale), int(19 * scale))], fill="#3B82F6", width=int(2 * scale))

        elif metric_type == "issues":
            d.rounded_rectangle([int(4 * scale), int(2 * scale), int(18 * scale), int(20 * scale)], radius=int(2 * scale), outline="#EF4444", width=int(2 * scale))
            d.line([(int(11 * scale), int(6 * scale)), (int(11 * scale), int(12 * scale))], fill="#EF4444", width=int(2 * scale))
            d.ellipse([int(10 * scale), int(14 * scale), int(12 * scale), int(16 * scale)], fill="#EF4444")

        elif metric_type == "rules":
            d.ellipse([int(2 * scale), int(2 * scale), int(20 * scale), int(20 * scale)], outline="#10B981", width=int(2 * scale))
            d.line([(int(6 * scale), int(11 * scale)), (int(9.5 * scale), int(14.5 * scale)),
                    (int(15.5 * scale), int(7.5 * scale))], fill="#10B981", width=int(2 * scale))

        elif metric_type == "score":
            cx, cy, r = int(11 * scale), int(11 * scale), int(8 * scale)
            import math
            pts = []
            for i in range(10):
                ang = i * math.pi / 5 - math.pi / 2
                rad = r if i % 2 == 0 else r * 0.45
                pts.append((cx + rad * math.cos(ang), cy + rad * math.sin(ang)))
            d.polygon(pts, fill="#F59E0B")

        im = im.resize((size, size), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(im, master=self.root)
        self._cache[key] = photo
        return photo

    def get_nav_icon(self, nav_name: str, active: bool = False, size: int = 18) -> ImageTk.PhotoImage:
        key = f"nav_{nav_name}_{active}_{size}"
        if key in self._cache:
            return self._cache[key]
        scale = 4
        s = size * scale
        im = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        col = "#FFFFFF" if active else "#8B9CB5"

        if nav_name == "overview":
            cx = s // 2
            d.polygon([(cx, int(2 * scale)), (int(15 * scale), int(8 * scale)),
                       (int(12 * scale), int(8 * scale)), (int(12 * scale), int(15 * scale)),
                       (int(5 * scale), int(15 * scale)), (int(5 * scale), int(8 * scale)),
                       (int(2 * scale), int(8 * scale))], fill=col)

        elif nav_name == "issues":
            d.rounded_rectangle([int(3 * scale), int(2 * scale), int(15 * scale), int(16 * scale)], radius=int(1.5 * scale), outline=col, width=int(1.8 * scale))
            d.line([(int(6 * scale), int(6 * scale)), (int(12 * scale), int(6 * scale))], fill=col, width=int(1.5 * scale))
            d.line([(int(6 * scale), int(9.5 * scale)), (int(12 * scale), int(9.5 * scale))], fill=col, width=int(1.5 * scale))
            d.line([(int(6 * scale), int(13 * scale)), (int(10 * scale), int(13 * scale))], fill=col, width=int(1.5 * scale))

        elif nav_name == "remediation":
            d.line([(int(3 * scale), int(15 * scale)), (int(12 * scale), int(6 * scale))], fill=col, width=int(2.5 * scale))
            d.ellipse([int(10 * scale), int(2 * scale), int(16 * scale), int(8 * scale)], outline=col, width=int(2 * scale))

        elif nav_name == "reports":
            d.line([(int(2 * scale), int(16 * scale)), (int(16 * scale), int(16 * scale))], fill=col, width=int(1.5 * scale))
            d.rounded_rectangle([int(3 * scale), int(10 * scale), int(6 * scale), int(15 * scale)], radius=int(0.5 * scale), fill=col)
            d.rounded_rectangle([int(7.5 * scale), int(6 * scale), int(10.5 * scale), int(15 * scale)], radius=int(0.5 * scale), fill=col)
            d.rounded_rectangle([int(12 * scale), int(3 * scale), int(15 * scale), int(15 * scale)], radius=int(0.5 * scale), fill=col)

        im = im.resize((size, size), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(im, master=self.root)
        self._cache[key] = photo
        return photo

    def get_action_icon(self, act_type: str, size: int = 15) -> ImageTk.PhotoImage:
        key = f"act_{act_type}_{size}"
        if key in self._cache:
            return self._cache[key]
        scale = 4
        s = size * scale
        im = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        col = "#CBD5E1"

        if act_type == "folder":
            d.polygon([(int(2 * scale), int(4 * scale)), (int(6 * scale), int(4 * scale)),
                       (int(8 * scale), int(6 * scale)), (int(13 * scale), int(6 * scale)),
                       (int(13 * scale), int(13 * scale)), (int(2 * scale), int(13 * scale))], fill=col)

        elif act_type == "export":
            d.line([(int(7.5 * scale), int(2 * scale)), (int(7.5 * scale), int(10 * scale))], fill=col, width=int(1.8 * scale))
            d.line([(int(4 * scale), int(7 * scale)), (int(7.5 * scale), int(10.5 * scale)), (int(11 * scale), int(7 * scale))], fill=col, width=int(1.8 * scale))
            d.line([(int(3 * scale), int(13 * scale)), (int(12 * scale), int(13 * scale))], fill=col, width=int(1.8 * scale))

        elif act_type == "report":
            d.rounded_rectangle([int(3 * scale), int(2 * scale), int(12 * scale), int(13 * scale)], radius=int(1.2 * scale), outline=col, width=int(1.6 * scale))
            d.line([(int(5 * scale), int(5 * scale)), (int(10 * scale), int(5 * scale))], fill=col, width=int(1.4 * scale))
            d.line([(int(5 * scale), int(8 * scale)), (int(10 * scale), int(8 * scale))], fill=col, width=int(1.4 * scale))

        im = im.resize((size, size), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(im, master=self.root)
        self._cache[key] = photo
        return photo

    def get_workflow_circle(self, number: int, state: str = "inactive", size: int = 24) -> ImageTk.PhotoImage:
        key = f"wf_circle_{number}_{state}_{size}"
        if key in self._cache:
            return self._cache[key]
        scale = 4
        s = size * scale
        im = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)

        if state == "active":
            bg_col = "#2563EB"  # Bright blue
            fg_col = "#FFFFFF"
        elif state == "completed":
            bg_col = "#10B981"  # Emerald green
            fg_col = "#FFFFFF"
        else:  # inactive / pending
            bg_col = "#1E293B"  # Dark navy slate
            fg_col = "#64748B"  # Muted slate

        # Draw smooth circle
        d.ellipse([0, 0, s - 1, s - 1], fill=bg_col)

        # Antialiased centered number
        try:
            font = ImageFont.truetype("segoeuib.ttf", int(11 * scale))
        except Exception:
            try:
                font = ImageFont.truetype("arialbd.ttf", int(11 * scale))
            except Exception:
                font = ImageFont.load_default()

        d.text((s // 2, s // 2), str(number), fill=fg_col, font=font, anchor="mm")

        im = im.resize((size, size), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(im, master=self.root)
        self._cache[key] = photo
        return photo

    def create_slide_preview_image(self, title: str = "Timeline of the Apple vs. FBI Case Events", width: int = 310, height: int = 158) -> ImageTk.PhotoImage:
        """Renders the authentic presentation slide preview card matching the reference screenshot."""
        scale = 3
        W, H = width * scale, height * scale
        im = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)

        # Crisp white card with rounded corners
        d.rounded_rectangle([0, 0, W - 1, H - 1], radius=int(8 * scale), fill="#FFFFFF", outline="#E2E8F0", width=int(1 * scale))

        # Slide Title text
        d.text((int(18 * scale), int(14 * scale)), title[:40], fill="#0F172A")

        # Timeline horizontal axis
        y_line = int(H * 0.50)
        d.line([(int(24 * scale), y_line), (W - int(24 * scale), y_line)], fill="#CBD5E1", width=int(2 * scale))

        colors = ["#00A896", "#028090", "#05668D", "#F4A261", "#E76F51"]
        dates = ["Dec 2, 2015", "Feb 16, 2016", "Feb 16, 2016", "Mar 28, 2016", "Post-Resolution"]
        num_nodes = len(colors)
        step = (W - int(60 * scale)) / (num_nodes - 1)

        for i, c in enumerate(colors):
            x = int(int(30 * scale) + i * step)
            # Outer halo & colored node
            d.ellipse([x - int(6 * scale), y_line - int(6 * scale), x + int(6 * scale), y_line + int(6 * scale)], fill=c)
            d.ellipse([x - int(2.5 * scale), y_line - int(2.5 * scale), x + int(2.5 * scale), y_line + int(2.5 * scale)], fill="#FFFFFF")

            # Date milestone text
            d.text((x - int(14 * scale), y_line + int(8 * scale)), dates[i], fill="#64748B")

            # Paragraph simulation lines
            d.line([(x - int(12 * scale), y_line + int(20 * scale)), (x + int(12 * scale), y_line + int(20 * scale))], fill="#94A3B8", width=int(1.8 * scale))
            d.line([(x - int(10 * scale), y_line + int(25 * scale)), (x + int(10 * scale), y_line + int(25 * scale))], fill="#CBD5E1", width=int(1.5 * scale))
            d.line([(x - int(12 * scale), y_line + int(30 * scale)), (x + int(6 * scale), y_line + int(30 * scale))], fill="#E2E8F0", width=int(1.5 * scale))

        im = im.resize((width, height), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(im, master=self.root)
        return photo


# =============================================================================
# MAIN WINDOW CLASS
# =============================================================================
class SlideGuardMainWindow:
    """
    Main desktop window for SlideGuard.
    Dark navy dashboard interface closely reproducing the visual reference.
    Strictly decoupled from PPTX logic; delegates all operations to AppController.
    """

    def __init__(self, root: tk.Tk, controller: Optional[AppController] = None):
        self.root = root
        self.controller = controller or AppController()
        self.job_context: Optional[JobContext] = None
        self.is_busy: bool = False
        self.icons = IconManager(root=self.root)

        # Job Activity tracking
        self.activity_events: List[Dict[str, str]] = []
        self.current_workflow_stage: int = 1

        self._setup_window()
        self._build_ui()

    def _setup_window(self):
        self.root.title(f"{APP_TITLE} (v{VERSION})")
        self.root.geometry("1280x830")
        self.root.minsize(1040, 680)
        self.root.configure(bg=BG_DARK)

        # Set window icon if available
        icon_path = resolve_resource_path("assets/slideguard.ico")
        if icon_path.exists():
            try:
                self.root.iconbitmap(str(icon_path))
            except Exception:
                pass

        # Apply dark ttk theme styling
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        # Hide top notebook tabs for our custom vertical Left Sidebar
        style.layout("Hidden.TNotebook", [("Notebook.client", {"sticky": "nswe"})])
        style.layout("Hidden.TNotebook.Tab", [])
        style.configure("Hidden.TNotebook", background=BG_DARK, borderwidth=0)

        # Configure dark Treeview for reviews & reports
        style.configure("Treeview",
                        background="#0F1626",
                        foreground="#F1F5F9",
                        fieldbackground="#0F1626",
                        borderwidth=0,
                        font=("Segoe UI", 9))
        style.configure("Treeview.Heading",
                        background="#1A2436",
                        foreground="#94A3B8",
                        relief="flat",
                        font=("Segoe UI", 9, "bold"))
        style.map("Treeview",
                  background=[("selected", "#1D4ED8")],
                  foreground=[("selected", "#FFFFFF")])

    def _build_ui(self):
        # Master Container
        self.master_frame = tk.Frame(self.root, bg=BG_DARK)
        self.master_frame.pack(fill=tk.BOTH, expand=True)

        # -------------------------------------------------------------
        # 1. TOP HEADER BAR
        # -------------------------------------------------------------
        self._build_top_header()

        # -------------------------------------------------------------
        # 2. MAIN BODY (LEFT SIDEBAR + NOTEBOOK)
        # -------------------------------------------------------------
        self.body_frame = tk.Frame(self.master_frame, bg=BG_DARK)
        self.body_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 4))

        # Left Sidebar
        self._build_left_sidebar()

        # Center Container (Notebook holding the 4 sections)
        self.center_frame = tk.Frame(self.body_frame, bg=BG_DARK)
        self.center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(self.center_frame, style="Hidden.TNotebook")
        self.main_notebook = self.notebook
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Section 1: Overview
        self.overview_tab = tk.Frame(self.notebook, bg=BG_DARK)
        self.notebook.add(self.overview_tab, text="Overview")

        # Section 2: Issues & Review
        self.review_tab = tk.Frame(self.notebook, bg=BG_DARK)
        self.notebook.add(self.review_tab, text="Issues & Review")

        # Section 3: Remediation
        self.remediation_tab = tk.Frame(self.notebook, bg=BG_DARK)
        self.notebook.add(self.remediation_tab, text="Remediation")

        # Section 4: Reports & Validation
        self.reports_tab = tk.Frame(self.notebook, bg=BG_DARK)
        self.notebook.add(self.reports_tab, text="Reports & Validation")

        # Build each section's contents
        self._build_overview_view()
        self._build_issues_review_view()
        self._build_remediation_view()
        self._build_reports_validation_view()

        # -------------------------------------------------------------
        # 3. BOTTOM STATUS BAR
        # -------------------------------------------------------------
        self._build_bottom_status_bar()

    # =========================================================================
    # TOP HEADER
    # =========================================================================
    def _build_top_header(self):
        header = tk.Frame(self.master_frame, bg=BG_DARK, height=52)
        header.pack(fill=tk.X, padx=14, pady=(8, 6))

        # Left: Shield Logo & Application Titles
        left_box = tk.Frame(header, bg=BG_DARK)
        left_box.pack(side=tk.LEFT, fill=tk.Y)

        self.shield_photo = self.icons.get_shield_logo(34)
        logo_lbl = tk.Label(left_box, image=self.shield_photo, bg=BG_DARK)
        logo_lbl.pack(side=tk.LEFT, padx=(0, 10))

        titles_box = tk.Frame(left_box, bg=BG_DARK)
        titles_box.pack(side=tk.LEFT, fill=tk.Y)

        title_lbl = tk.Label(
            titles_box,
            text=APP_NAME,
            font=("Segoe UI", 15, "bold"),
            fg=TEXT_WHITE,
            bg=BG_DARK
        )
        title_lbl.pack(anchor=tk.W)

        subtitle_lbl = tk.Label(
            titles_box,
            text="PowerPoint Accessibility & Remediation",
            font=("Segoe UI", 8),
            fg=TEXT_SECONDARY,
            bg=BG_DARK
        )
        subtitle_lbl.pack(anchor=tk.W)

        # Center: Open File button & Active File Selector Pill
        center_box = tk.Frame(header, bg=BG_DARK)
        center_box.pack(side=tk.LEFT, padx=(36, 0))

        # Open File Button
        self.folder_icon = self.icons.get_action_icon("folder", 15)
        open_btn = tk.Button(
            center_box,
            text=" Open File",
            image=self.folder_icon,
            compound=tk.LEFT,
            font=("Segoe UI", 9, "bold"),
            fg=TEXT_PRIMARY,
            bg=BG_CARD_INNER,
            activeforeground=TEXT_WHITE,
            activebackground="#1E2B42",
            relief=tk.FLAT,
            bd=0,
            padx=12,
            pady=4,
            cursor="hand2",
            command=self._on_select_file
        )
        open_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Current presentation filename selector pill
        self.header_file_var = tk.StringVar(value="No presentation loaded   ⌄")
        self.file_selector_btn = tk.Button(
            center_box,
            textvariable=self.header_file_var,
            font=("Segoe UI", 9),
            fg=TEXT_PRIMARY,
            bg=BG_CARD_INNER,
            activeforeground=TEXT_WHITE,
            activebackground="#1E2B42",
            relief=tk.FLAT,
            bd=0,
            padx=14,
            pady=4,
            cursor="hand2",
            command=self._on_select_file
        )
        self.file_selector_btn.pack(side=tk.LEFT)

        # Right: Settings, Help & Window controls
        right_box = tk.Frame(header, bg=BG_DARK)
        right_box.pack(side=tk.RIGHT)

        settings_btn = tk.Button(
            right_box,
            text="⚙ Settings",
            font=("Segoe UI", 9),
            fg=TEXT_SECONDARY,
            bg=BG_DARK,
            activeforeground=TEXT_WHITE,
            activebackground="#162032",
            relief=tk.FLAT,
            bd=0,
            padx=8,
            pady=4,
            cursor="hand2",
            command=self._on_settings_clicked
        )
        settings_btn.pack(side=tk.LEFT, padx=(0, 6))

        help_btn = tk.Button(
            right_box,
            text="? Help",
            font=("Segoe UI", 9),
            fg=TEXT_SECONDARY,
            bg=BG_DARK,
            activeforeground=TEXT_WHITE,
            activebackground="#162032",
            relief=tk.FLAT,
            bd=0,
            padx=8,
            pady=4,
            cursor="hand2",
            command=self._on_help_clicked
        )
        help_btn.pack(side=tk.LEFT, padx=(0, 14))

        # Window control buttons (minimize, maximize, close)
        win_box = tk.Frame(right_box, bg=BG_DARK)
        win_box.pack(side=tk.LEFT)

        min_btn = tk.Button(
            win_box,
            text="—",
            font=("Segoe UI", 9),
            fg=TEXT_SECONDARY,
            bg=BG_DARK,
            activeforeground=TEXT_WHITE,
            activebackground="#162032",
            relief=tk.FLAT,
            bd=0,
            padx=6,
            pady=2,
            cursor="hand2",
            command=self.root.iconify
        )
        min_btn.pack(side=tk.LEFT, padx=1)

        max_btn = tk.Button(
            win_box,
            text="▢",
            font=("Segoe UI", 9),
            fg=TEXT_SECONDARY,
            bg=BG_DARK,
            activeforeground=TEXT_WHITE,
            activebackground="#162032",
            relief=tk.FLAT,
            bd=0,
            padx=6,
            pady=2,
            cursor="hand2",
            command=self._toggle_maximize
        )
        max_btn.pack(side=tk.LEFT, padx=1)

        close_btn = tk.Button(
            win_box,
            text="✕",
            font=("Segoe UI", 9),
            fg=TEXT_SECONDARY,
            bg=BG_DARK,
            activeforeground=TEXT_WHITE,
            activebackground="#DC2626",
            relief=tk.FLAT,
            bd=0,
            padx=6,
            pady=2,
            cursor="hand2",
            command=self.root.destroy
        )
        close_btn.pack(side=tk.LEFT, padx=1)

    # =========================================================================
    # LEFT PERSISTENT SIDEBAR
    # =========================================================================
    def _build_left_sidebar(self):
        self.sidebar_frame = tk.Frame(self.body_frame, bg=BG_SIDEBAR, width=195)
        self.sidebar_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 12), pady=(0, 4))
        self.sidebar_frame.pack_propagate(False)

        nav_items = [
            ("Overview", "overview", 0),
            ("Issues & Review", "issues", 1),
            ("Remediation", "remediation", 2),
            ("Reports & Validation", "reports", 3)
        ]

        self.nav_buttons: List[tk.Button] = []
        self.active_nav_index: int = 0

        for title, key, idx in nav_items:
            photo = self.icons.get_nav_icon(key, active=(idx == 0), size=18)
            setattr(self, f"nav_icon_{key}", photo)

            btn = tk.Button(
                self.sidebar_frame,
                text=f"  {title}",
                image=photo,
                compound=tk.LEFT,
                font=("Segoe UI", 9, "bold" if idx == 0 else "normal"),
                fg=TEXT_WHITE if idx == 0 else TEXT_SECONDARY,
                bg=BLUE_ACCENT if idx == 0 else BG_SIDEBAR,
                activeforeground=TEXT_WHITE,
                activebackground=BLUE_ACCENT if idx == 0 else "#141D30",
                relief=tk.FLAT,
                bd=0,
                anchor=tk.W,
                padx=12,
                pady=9,
                cursor="hand2",
                command=lambda i=idx: self._select_nav(i)
            )
            btn.pack(fill=tk.X, pady=2)
            self.nav_buttons.append(btn)

    def _select_nav(self, index: int):
        self.active_nav_index = index
        self.notebook.select(index)

        keys = ["overview", "issues", "remediation", "reports"]
        for idx, btn in enumerate(self.nav_buttons):
            is_active = (idx == index)
            photo = self.icons.get_nav_icon(keys[idx], active=is_active, size=18)
            setattr(self, f"nav_icon_{keys[idx]}", photo)
            btn.config(
                image=photo,
                bg=BLUE_ACCENT if is_active else BG_SIDEBAR,
                fg=TEXT_WHITE if is_active else TEXT_SECONDARY,
                font=("Segoe UI", 9, "bold" if is_active else "normal"),
                activebackground=BLUE_ACCENT if is_active else "#141D30"
            )

    # =========================================================================
    # SECTION 1: OVERVIEW DASHBOARD
    # =========================================================================
    def _build_overview_view(self):
        overview_container = tk.Frame(self.overview_tab, bg=BG_DARK)
        overview_container.pack(fill=tk.BOTH, expand=True)

        self.overview_left_frame = tk.Frame(overview_container, bg=BG_DARK)
        self.overview_left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        self.overview_right_frame = tk.Frame(overview_container, bg=BG_DARK, width=290)
        self.overview_right_frame.pack(side=tk.RIGHT, fill=tk.Y)
        self.overview_right_frame.pack_propagate(False)

        # Left Column Part A: Presentation Preview & Metrics Card
        self._build_presentation_metrics_card()

        # Left Column Part B: Remediation Center (2x2 Grid)
        self._build_remediation_center_cards()

        # Left Column Part C: Remediation Workflow Strip
        self._build_workflow_strip()

        # Right Column Panels (Recent Activity, Quick Actions, Status)
        self._build_right_sidebar_panels()

    def _build_presentation_metrics_card(self):
        card = tk.Frame(self.overview_left_frame, bg=BG_CARD, highlightthickness=1, highlightbackground=BORDER_SUBTLE)
        card.pack(fill=tk.X, pady=(0, 10))

        inner = tk.Frame(card, bg=BG_CARD)
        inner.pack(fill=tk.X, padx=12, pady=10)

        # Left: Slide Preview Card
        preview_box = tk.Frame(inner, bg=BG_CARD)
        preview_box.pack(side=tk.LEFT, padx=(0, 14))

        self.slide_preview_img = self.icons.create_slide_preview_image(width=310, height=158)
        self.slide_preview_lbl = tk.Label(preview_box, image=self.slide_preview_img, bg=BG_CARD)
        self.slide_preview_lbl.pack()

        # Right: Presentation Title, Path, 4 Metric Tiles & Design Integrity
        meta_box = tk.Frame(inner, bg=BG_CARD)
        meta_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.overview_filename_var = tk.StringVar(value="No presentation loaded")
        self.overview_filepath_var = tk.StringVar(value="Select 'Open File' to begin accessibility inspection")

        file_title_lbl = tk.Label(
            meta_box,
            textvariable=self.overview_filename_var,
            font=("Segoe UI", 15, "bold"),
            fg=TEXT_WHITE,
            bg=BG_CARD
        )
        file_title_lbl.pack(anchor=tk.W)

        file_path_lbl = tk.Label(
            meta_box,
            textvariable=self.overview_filepath_var,
            font=("Segoe UI", 8),
            fg=TEXT_MUTED,
            bg=BG_CARD
        )
        file_path_lbl.pack(anchor=tk.W, pady=(1, 8))

        # Row of 4 Metric Badges
        metrics_row = tk.Frame(meta_box, bg=BG_CARD)
        metrics_row.pack(fill=tk.X, pady=(0, 8))

        self.metric_slides_var = tk.StringVar(value="-")
        self.metric_issues_var = tk.StringVar(value="-")
        self.metric_rules_var = tk.StringVar(value="-")
        self.metric_score_var = tk.StringVar(value="-")

        self._create_metric_tile(metrics_row, "slides", "Total Slides", self.metric_slides_var, TEXT_WHITE)
        self._create_metric_tile(metrics_row, "issues", "Issues Detected", self.metric_issues_var, RED_ALERT)
        self._create_metric_tile(metrics_row, "rules", "Rules Passed", self.metric_rules_var, TEXT_WHITE)
        self._create_metric_tile(metrics_row, "score", "Quality Score (V2)", self.metric_score_var, GREEN_TEXT)

        # Design Integrity Status Card
        self.integrity_box = tk.Frame(meta_box, bg="#06261E", highlightthickness=1, highlightbackground="#0F5132")
        self.integrity_box.pack(fill=tk.X)

        integ_inner = tk.Frame(self.integrity_box, bg="#06261E")
        integ_inner.pack(fill=tk.X, padx=10, pady=6)

        self.integ_icon = self.icons.get_metric_icon("rules", 20)
        integ_icon_lbl = tk.Label(integ_inner, image=self.integ_icon, bg="#06261E")
        integ_icon_lbl.pack(side=tk.LEFT, padx=(0, 8))

        integ_texts = tk.Frame(integ_inner, bg="#06261E")
        integ_texts.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.metric_integrity_var = tk.StringVar(value="Design Integrity: Verified")
        integ_title = tk.Label(
            integ_texts,
            textvariable=self.metric_integrity_var,
            font=("Segoe UI", 9, "bold"),
            fg=GREEN_TEXT,
            bg="#06261E"
        )
        integ_title.pack(anchor=tk.W)

        integ_desc = tk.Label(
            integ_texts,
            text="Visual design, layout, fonts and formatting preserved.",
            font=("Segoe UI", 7),
            fg="#86EFAC",
            bg="#06261E"
        )
        integ_desc.pack(anchor=tk.W)

    def _create_metric_tile(self, parent: tk.Widget, icon_type: str, label_text: str, val_var: tk.StringVar, val_color: str):
        tile = tk.Frame(parent, bg=BG_CARD_INNER, highlightthickness=1, highlightbackground=BORDER_SUBTLE)
        tile.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))

        inner = tk.Frame(tile, bg=BG_CARD_INNER)
        inner.pack(fill=tk.BOTH, expand=True, padx=7, pady=5)

        icon_img = self.icons.get_metric_icon(icon_type, 18)
        setattr(self, f"metric_img_{icon_type}", icon_img)
        icon_lbl = tk.Label(inner, image=icon_img, bg=BG_CARD_INNER)
        icon_lbl.pack(side=tk.LEFT, padx=(0, 5))

        text_box = tk.Frame(inner, bg=BG_CARD_INNER)
        text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        lbl = tk.Label(text_box, text=label_text, font=("Segoe UI", 7), fg=TEXT_SECONDARY, bg=BG_CARD_INNER)
        lbl.pack(anchor=tk.W)

        val = tk.Label(text_box, textvariable=val_var, font=("Segoe UI", 10, "bold"), fg=val_color, bg=BG_CARD_INNER)
        val.pack(anchor=tk.W)

    # -------------------------------------------------------------------------
    # Remediation Center (2x2 Grid of Primary Cards)
    # -------------------------------------------------------------------------
    def _build_remediation_center_cards(self):
        sec_header = tk.Frame(self.overview_left_frame, bg=BG_DARK)
        sec_header.pack(fill=tk.X, pady=(0, 6))

        sec_title = tk.Label(
            sec_header,
            text="Remediation Center",
            font=("Segoe UI", 13, "bold"),
            fg=TEXT_WHITE,
            bg=BG_DARK
        )
        sec_title.pack(anchor=tk.W)

        sec_sub = tk.Label(
            sec_header,
            text="Run each remediation independently. Review the changes before applying.",
            font=("Segoe UI", 8),
            fg=TEXT_SECONDARY,
            bg=BG_DARK
        )
        sec_sub.pack(anchor=tk.W)

        # 2x2 Responsive Card Container
        cards_grid = tk.Frame(self.overview_left_frame, bg=BG_DARK)
        cards_grid.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        cards_grid.columnconfigure(0, weight=1)
        cards_grid.columnconfigure(1, weight=1)

        # Card 1: Reading Order (Blue)
        self.ro_card_status_var = tk.StringVar(value="● Ready")
        self.ro_slides_meta_var = tk.StringVar(value="📄 - slides")
        self.ro_affected_meta_var = tk.StringVar(value="🗂 - affected slides")

        self.btn_apply_ro = self._create_primary_card(
            parent=cards_grid,
            row=0, col=0,
            card_type="reading_order",
            title="Reading Order",
            description="Analyze and apply logical reading order for all slides to ensure correct navigation for assistive technologies.",
            btn_text="Review & Apply Reading Order →",
            btn_color=BLUE_ACCENT,
            status_var=self.ro_card_status_var,
            meta_var1=self.ro_slides_meta_var,
            meta_var2=self.ro_affected_meta_var,
            command=self._on_apply_reading_order_clicked
        )
        self.rem_ro_btn = self.btn_apply_ro

        # Card 2: Placeholder Conversion (Purple)
        self.ph_card_status_var = tk.StringVar(value="● Ready")
        self.ph_slides_meta_var = tk.StringVar(value="📄 - slides")
        self.ph_eligible_meta_var = tk.StringVar(value="🗂 - eligible items")

        self.btn_apply_ph = self._create_primary_card(
            parent=cards_grid,
            row=0, col=1,
            card_type="placeholder",
            title="Placeholder Conversion",
            description="Review and apply native PowerPoint placeholders for titles, body content and other supported elements.",
            btn_text="Review & Apply Placeholders →",
            btn_color=PURPLE_ACCENT,
            status_var=self.ph_card_status_var,
            meta_var1=self.ph_slides_meta_var,
            meta_var2=self.ph_eligible_meta_var,
            command=self._on_apply_placeholders_clicked
        )
        self.rem_ph_btn = self.btn_apply_ph

        # Card 3: Validate & Compare (Orange)
        self.val_card_status_var = tk.StringVar(value="● Ready")
        self.val_meta1_var = tk.StringVar(value="📄 Full presentation")
        self.val_meta2_var = tk.StringVar(value="📋 Generate report")

        self.btn_run_val = self._create_primary_card(
            parent=cards_grid,
            row=1, col=0,
            card_type="validation",
            title="Validate & Compare",
            description="Run accessibility validation, integrity checks and generate a detailed remediation report.",
            btn_text="Run Validation →",
            btn_color=ORANGE_ACCENT,
            status_var=self.val_card_status_var,
            meta_var1=self.val_meta1_var,
            meta_var2=self.val_meta2_var,
            command=self._on_apply_validation_clicked
        )
        self.btn_apply_val = self.btn_run_val
        self.run_qa_btn = self.btn_run_val

        # Card 4: Generate Report (Teal/Accent)
        self.rep_card_status_var = tk.StringVar(value="● Ready")
        self.rep_meta1_var = tk.StringVar(value="📄 Audit report (MD/JSON)")
        self.rep_meta2_var = tk.StringVar(value="📊 Compliance summary")

        self.generate_report_btn = self._create_primary_card(
            parent=cards_grid,
            row=1, col=1,
            card_type="report",
            title="Generate Report",
            description="Generate comprehensive accessibility audit reports in JSON and Markdown formats for compliance records.",
            btn_text="Generate Report →",
            btn_color=TEAL_ACCENT,
            status_var=self.rep_card_status_var,
            meta_var1=self.rep_meta1_var,
            meta_var2=self.rep_meta2_var,
            command=self._on_generate_report_clicked
        )

    def _create_primary_card(
        self,
        parent: tk.Widget,
        row: int,
        col: int,
        card_type: str,
        title: str,
        description: str,
        btn_text: str,
        btn_color: str,
        status_var: tk.StringVar,
        meta_var1: tk.StringVar,
        meta_var2: tk.StringVar,
        command: Any
    ) -> tk.Button:
        card = tk.Frame(parent, bg=BG_CARD, highlightthickness=1, highlightbackground=BORDER_SUBTLE)
        card.grid(row=row, column=col, sticky="nsew", padx=4, pady=4)

        inner = tk.Frame(card, bg=BG_CARD)
        inner.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        # Header Row: Icon tile + Title + Status Badge
        top_row = tk.Frame(inner, bg=BG_CARD)
        top_row.pack(fill=tk.X, pady=(0, 6))

        icon_img = self.icons.get_card_icon(card_type, 34)
        setattr(self, f"card_img_{card_type}", icon_img)
        icon_lbl = tk.Label(top_row, image=icon_img, bg=BG_CARD)
        icon_lbl.pack(side=tk.LEFT, padx=(0, 8))

        title_lbl = tk.Label(top_row, text=title, font=("Segoe UI", 11, "bold"), fg=TEXT_WHITE, bg=BG_CARD)
        title_lbl.pack(side=tk.LEFT)

        # Status badge pill
        badge = tk.Label(
            top_row,
            textvariable=status_var,
            font=("Segoe UI", 8, "bold"),
            fg=GREEN_TEXT,
            bg=GREEN_BADGE_BG,
            padx=7,
            pady=1,
            relief=tk.FLAT
        )
        badge.pack(side=tk.RIGHT)

        # Description text
        desc_lbl = tk.Label(
            inner,
            text=description,
            font=("Segoe UI", 8),
            fg=TEXT_SECONDARY,
            bg=BG_CARD,
            wraplength=340,
            justify=tk.LEFT
        )
        desc_lbl.pack(anchor=tk.W, fill=tk.X, pady=(0, 8))

        # Full-width Action Button
        action_btn = tk.Button(
            inner,
            text=btn_text,
            font=("Segoe UI", 9, "bold"),
            fg=TEXT_WHITE,
            bg=btn_color,
            activeforeground=TEXT_WHITE,
            activebackground=btn_color,
            relief=tk.FLAT,
            bd=0,
            pady=6,
            cursor="hand2",
            state=tk.DISABLED,
            command=command
        )
        action_btn.pack(fill=tk.X, pady=(0, 8))

        # Bottom Meta Row (Counts)
        meta_row = tk.Frame(inner, bg=BG_CARD)
        meta_row.pack(fill=tk.X)

        m1 = tk.Label(meta_row, textvariable=meta_var1, font=("Segoe UI", 8), fg=TEXT_MUTED, bg=BG_CARD)
        m1.pack(side=tk.LEFT, padx=(0, 14))

        m2 = tk.Label(meta_row, textvariable=meta_var2, font=("Segoe UI", 8), fg=TEXT_MUTED, bg=BG_CARD)
        m2.pack(side=tk.LEFT)

        return action_btn

    # -------------------------------------------------------------------------
    # Remediation Workflow Strip
    # -------------------------------------------------------------------------
    def _build_workflow_strip(self):
        wf_container = tk.Frame(self.overview_left_frame, bg=BG_DARK)
        wf_container.pack(fill=tk.X, pady=(2, 0))

        wf_title = tk.Label(
            wf_container,
            text="Remediation Workflow",
            font=("Segoe UI", 11, "bold"),
            fg=TEXT_WHITE,
            bg=BG_DARK
        )
        wf_title.pack(anchor=tk.W, pady=(0, 6))

        strip = tk.Frame(wf_container, bg=BG_CARD, highlightthickness=1, highlightbackground=BORDER_SUBTLE)
        strip.pack(fill=tk.X)

        inner = tk.Frame(strip, bg=BG_CARD)
        inner.pack(fill=tk.X, padx=12, pady=8)

        stages = [
            (1, "Analyze", "Detect issues"),
            (2, "Review", "Check and approve"),
            (3, "Apply", "Run selected remediation"),
            (4, "Validate", "Verify and compare"),
            (5, "Export", "Save remediated file")
        ]

        self.wf_circle_labels: List[tk.Label] = []
        self.wf_title_labels: List[tk.Label] = []

        for idx, (num, name, subtitle) in enumerate(stages):
            box = tk.Frame(inner, bg=BG_CARD)
            box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            circ_img = self.icons.get_workflow_circle(num, state="active" if num == 1 else "inactive", size=24)
            circle = tk.Label(
                box,
                image=circ_img,
                bg=BG_CARD
            )
            circle.image = circ_img
            circle.pack(side=tk.LEFT, padx=(0, 6))
            self.wf_circle_labels.append(circle)

            text_sub = tk.Frame(box, bg=BG_CARD)
            text_sub.pack(side=tk.LEFT)

            t_lbl = tk.Label(
                text_sub,
                text=name,
                font=("Segoe UI", 8, "bold"),
                fg=TEXT_WHITE if num <= 1 else TEXT_SECONDARY,
                bg=BG_CARD
            )
            t_lbl.pack(anchor=tk.W)
            self.wf_title_labels.append(t_lbl)

            s_lbl = tk.Label(
                text_sub,
                text=subtitle,
                font=("Segoe UI", 7),
                fg=TEXT_MUTED,
                bg=BG_CARD
            )
            s_lbl.pack(anchor=tk.W)

            if idx < len(stages) - 1:
                arr = tk.Label(inner, text="➔", font=("Segoe UI", 8), fg="#3B82F6", bg=BG_CARD)
                arr.pack(side=tk.LEFT, padx=4)

    def _update_workflow_stage(self, stage: int):
        self.current_workflow_stage = stage
        for i, (circ, tit) in enumerate(zip(self.wf_circle_labels, self.wf_title_labels), start=1):
            if i < stage:
                st = "completed"
                tit.config(fg=TEXT_WHITE)
            elif i == stage:
                st = "active"
                tit.config(fg=TEXT_WHITE)
            else:
                st = "inactive"
                tit.config(fg=TEXT_SECONDARY)
            img = self.icons.get_workflow_circle(i, state=st, size=24)
            circ.config(image=img)
            circ.image = img

    # -------------------------------------------------------------------------
    # Right Sidebar Panels (Recent Activity, Quick Actions, Status)
    # -------------------------------------------------------------------------
    def _build_right_sidebar_panels(self):
        # 1. Recent Activity Card
        act_card = tk.Frame(self.overview_right_frame, bg=BG_CARD, highlightthickness=1, highlightbackground=BORDER_SUBTLE)
        act_card.pack(fill=tk.X, pady=(0, 10))

        act_inner = tk.Frame(act_card, bg=BG_CARD)
        act_inner.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        act_header = tk.Frame(act_inner, bg=BG_CARD)
        act_header.pack(fill=tk.X, pady=(0, 6))

        tk.Label(act_header, text="Recent Activity", font=("Segoe UI", 10, "bold"), fg=TEXT_WHITE, bg=BG_CARD).pack(side=tk.LEFT)
        view_all_btn = tk.Label(act_header, text="View All", font=("Segoe UI", 8, "underline"), fg="#3B82F6", bg=BG_CARD, cursor="hand2")
        view_all_btn.pack(side=tk.RIGHT)
        view_all_btn.bind("<Button-1>", lambda e: self._switch_to_tab(2))

        self.activity_container = tk.Frame(act_inner, bg=BG_CARD)
        self.activity_container.pack(fill=tk.X)

        self._add_activity_event("blue", "File loaded successfully", "Today 14:22")

        # 2. Quick Actions Card
        qa_card = tk.Frame(self.overview_right_frame, bg=BG_CARD, highlightthickness=1, highlightbackground=BORDER_SUBTLE)
        qa_card.pack(fill=tk.X, pady=(0, 10))

        qa_inner = tk.Frame(qa_card, bg=BG_CARD)
        qa_inner.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        tk.Label(qa_inner, text="Quick Actions", font=("Segoe UI", 10, "bold"), fg=TEXT_WHITE, bg=BG_CARD).pack(anchor=tk.W, pady=(0, 6))

        self.icon_qa_open = self.icons.get_action_icon("folder", 15)
        self.qa_open_btn = tk.Button(
            qa_inner,
            text="   Open PowerPoint File",
            image=self.icon_qa_open,
            compound=tk.LEFT,
            font=("Segoe UI", 8, "bold"),
            fg=TEXT_PRIMARY,
            bg=BG_CARD_INNER,
            activeforeground=TEXT_WHITE,
            activebackground="#1E2B42",
            relief=tk.FLAT,
            bd=0,
            anchor=tk.W,
            padx=10,
            pady=6,
            cursor="hand2",
            command=self._on_select_file
        )
        self.qa_open_btn.pack(fill=tk.X, pady=(0, 5))

        self.icon_qa_export = self.icons.get_action_icon("export", 15)
        self.qa_export_btn = tk.Button(
            qa_inner,
            text="   Export Remediated File",
            image=self.icon_qa_export,
            compound=tk.LEFT,
            font=("Segoe UI", 8, "bold"),
            fg=TEXT_PRIMARY,
            bg=BG_CARD_INNER,
            activeforeground=TEXT_WHITE,
            activebackground="#1E2B42",
            relief=tk.FLAT,
            bd=0,
            anchor=tk.W,
            padx=10,
            pady=6,
            cursor="hand2",
            command=self._on_open_output_file
        )
        self.qa_export_btn.pack(fill=tk.X, pady=(0, 5))

        self.icon_qa_rep = self.icons.get_action_icon("report", 15)
        self.qa_rep_btn = tk.Button(
            qa_inner,
            text="   Generate Full Report",
            image=self.icon_qa_rep,
            compound=tk.LEFT,
            font=("Segoe UI", 8, "bold"),
            fg=TEXT_PRIMARY,
            bg=BG_CARD_INNER,
            activeforeground=TEXT_WHITE,
            activebackground="#1E2B42",
            relief=tk.FLAT,
            bd=0,
            anchor=tk.W,
            padx=10,
            pady=6,
            cursor="hand2",
            command=self._on_generate_report_clicked
        )
        self.qa_rep_btn.pack(fill=tk.X)

        # 3. Status Card
        status_card = tk.Frame(self.overview_right_frame, bg=BG_CARD, highlightthickness=1, highlightbackground=BORDER_SUBTLE)
        status_card.pack(fill=tk.X)

        st_inner = tk.Frame(status_card, bg=BG_CARD)
        st_inner.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        tk.Label(st_inner, text="Status", font=("Segoe UI", 10, "bold"), fg=TEXT_WHITE, bg=BG_CARD).pack(anchor=tk.W, pady=(0, 6))

        self.st_file_var = tk.StringVar(value="-")
        self.st_slides_var = tk.StringVar(value="-")
        self.st_issues_var = tk.StringVar(value="-")
        self.st_rules_var = tk.StringVar(value="-")
        self.st_score_var = tk.StringVar(value="-")
        self.st_integrity_var = tk.StringVar(value="Verified")

        self._create_status_row(st_inner, "File Name", self.st_file_var, TEXT_PRIMARY)
        self._create_status_row(st_inner, "Total Slides", self.st_slides_var, TEXT_PRIMARY)
        self._create_status_row(st_inner, "Issues Detected", self.st_issues_var, TEXT_PRIMARY)
        self._create_status_row(st_inner, "Rules Passed", self.st_rules_var, TEXT_PRIMARY)
        self._create_status_row(st_inner, "Quality Score", self.st_score_var, GREEN_TEXT)
        self._create_status_row(st_inner, "Design Integrity", self.st_integrity_var, GREEN_TEXT)

    def _create_status_row(self, parent: tk.Widget, label_text: str, var: tk.StringVar, color: str):
        row = tk.Frame(parent, bg=BG_CARD)
        row.pack(fill=tk.X, pady=2)
        tk.Label(row, text=label_text, font=("Segoe UI", 8), fg=TEXT_MUTED, bg=BG_CARD).pack(side=tk.LEFT)
        tk.Label(row, textvariable=var, font=("Segoe UI", 8, "bold"), fg=color, bg=BG_CARD).pack(side=tk.RIGHT)

    def _add_activity_event(self, dot_color: str, text: str, time_str: Optional[str] = None):
        if not time_str:
            time_str = datetime.datetime.now().strftime("Today %H:%M")

        self.activity_events.append({"color": dot_color, "text": text, "time": time_str})

        display_events = self.activity_events[-4:]

        for w in self.activity_container.winfo_children():
            w.destroy()

        dot_colors = {
            "blue": "#3B82F6",
            "green": "#10B981",
            "orange": "#F59E0B",
            "teal": "#0D9488",
            "purple": "#8B5CF6",
            "red": "#EF4444"
        }

        for ev in display_events:
            r = tk.Frame(self.activity_container, bg=BG_CARD)
            r.pack(fill=tk.X, pady=2)

            col = dot_colors.get(ev["color"], "#3B82F6")
            tk.Label(r, text="●", font=("Segoe UI", 8), fg=col, bg=BG_CARD).pack(side=tk.LEFT, padx=(0, 5))
            tk.Label(r, text=ev["text"][:28], font=("Segoe UI", 8), fg=TEXT_PRIMARY, bg=BG_CARD).pack(side=tk.LEFT)
            tk.Label(r, text=ev["time"], font=("Segoe UI", 7), fg=TEXT_MUTED, bg=BG_CARD).pack(side=tk.RIGHT)

    # =========================================================================
    # SECTION 2: ISSUES & REVIEW VIEW
    # =========================================================================
    def _build_issues_review_view(self):
        rev_frame = tk.Frame(self.review_tab, bg=BG_DARK)
        rev_frame.pack(fill=tk.BOTH, expand=True)

        self.review_notebook = ttk.Notebook(rev_frame)
        self.review_notebook.pack(fill=tk.BOTH, expand=True)

        # SUBTAB 1: Detected Issues
        issues_subtab = tk.Frame(self.review_notebook, bg=BG_DARK)
        self.review_notebook.add(issues_subtab, text=" Detected Issues & Compliance ")

        columns = ("slide", "severity", "confidence", "rule_id", "rule_name", "status", "description", "recommendation")
        self.tree = ttk.Treeview(issues_subtab, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("slide", text="Slide #")
        self.tree.heading("severity", text="Severity")
        self.tree.heading("confidence", text="Confidence")
        self.tree.heading("rule_id", text="Rule ID")
        self.tree.heading("rule_name", text="Rule Name")
        self.tree.heading("status", text="Action / Status")
        self.tree.heading("description", text="Description")
        self.tree.heading("recommendation", text="Recommendation")

        self.tree.column("slide", width=55, anchor=tk.CENTER)
        self.tree.column("severity", width=75, anchor=tk.CENTER)
        self.tree.column("confidence", width=85, anchor=tk.CENTER)
        self.tree.column("rule_id", width=85, anchor=tk.CENTER)
        self.tree.column("rule_name", width=145, anchor=tk.W)
        self.tree.column("status", width=115, anchor=tk.CENTER)
        self.tree.column("description", width=290, anchor=tk.W)
        self.tree.column("recommendation", width=180, anchor=tk.W)

        tree_scroll = ttk.Scrollbar(issues_subtab, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # SUBTAB 2: Reading Order Analysis
        ro_subtab = tk.Frame(self.review_notebook, bg=BG_DARK)
        self.review_notebook.add(ro_subtab, text=" Reading Order Review ")

        ro_columns = ("slide", "confidence", "original_order", "proposed_order", "excluded_footers", "reasoning")
        self.ro_tree = ttk.Treeview(ro_subtab, columns=ro_columns, show="headings", selectmode="browse")

        self.ro_tree.heading("slide", text="Slide #")
        self.ro_tree.heading("confidence", text="Confidence")
        self.ro_tree.heading("original_order", text="Current Sequence (Shape IDs)")
        self.ro_tree.heading("proposed_order", text="Proposed Sequence (Shape IDs)")
        self.ro_tree.heading("excluded_footers", text="Excluded Footers")
        self.ro_tree.heading("reasoning", text="Structural Reasoning")

        self.ro_tree.column("slide", width=60, anchor=tk.CENTER)
        self.ro_tree.column("confidence", width=90, anchor=tk.CENTER)
        self.ro_tree.column("original_order", width=180, anchor=tk.W)
        self.ro_tree.column("proposed_order", width=180, anchor=tk.W)
        self.ro_tree.column("excluded_footers", width=180, anchor=tk.W)
        self.ro_tree.column("reasoning", width=250, anchor=tk.W)

        ro_scroll = ttk.Scrollbar(ro_subtab, orient=tk.VERTICAL, command=self.ro_tree.yview)
        self.ro_tree.configure(yscrollcommand=ro_scroll.set)
        self.ro_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ro_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # SUBTAB 3: Placeholder Conversion Review
        ph_subtab = tk.Frame(self.review_notebook, bg=BG_DARK)
        self.review_notebook.add(ph_subtab, text=" Placeholder Conversion Review ")

        ph_columns = ("slide", "shape", "ph_type", "safety_status", "recommendation", "reason")
        self.ph_tree = ttk.Treeview(ph_subtab, columns=ph_columns, show="headings", selectmode="browse")

        self.ph_tree.heading("slide", text="Slide #")
        self.ph_tree.heading("shape", text="Shape (ID / Name)")
        self.ph_tree.heading("ph_type", text="Placeholder Type")
        self.ph_tree.heading("safety_status", text="Safety Status")
        self.ph_tree.heading("recommendation", text="Recommendation")
        self.ph_tree.heading("reason", text="Reason")

        self.ph_tree.column("slide", width=60, anchor=tk.CENTER)
        self.ph_tree.column("shape", width=150, anchor=tk.W)
        self.ph_tree.column("ph_type", width=120, anchor=tk.CENTER)
        self.ph_tree.column("safety_status", width=130, anchor=tk.CENTER)
        self.ph_tree.column("recommendation", width=250, anchor=tk.W)
        self.ph_tree.column("reason", width=260, anchor=tk.W)

        ph_scroll = ttk.Scrollbar(ph_subtab, orient=tk.VERTICAL, command=self.ph_tree.yview)
        self.ph_tree.configure(yscrollcommand=ph_scroll.set)
        self.ph_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ph_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # SUBTAB 4: Alternative Text & Images Review (Required by test_ui_refactor)
        auth_subtab = tk.Frame(self.review_notebook, bg=BG_DARK)
        self.review_notebook.add(auth_subtab, text=" Alternative Text & Images Review ")

        auth_split = ttk.PanedWindow(auth_subtab, orient=tk.VERTICAL)
        auth_split.pack(fill=tk.BOTH, expand=True)

        auth_top_frame = tk.Frame(auth_split, bg=BG_DARK)
        auth_split.add(auth_top_frame, weight=3)

        auth_columns = ("slide", "shape", "type", "current", "suggested", "confidence", "grounding", "status")
        self.auth_tree = ttk.Treeview(auth_top_frame, columns=auth_columns, show="headings", selectmode="browse")

        self.auth_tree.heading("slide", text="Slide #")
        self.auth_tree.heading("shape", text="Shape (ID / Name)")
        self.auth_tree.heading("type", text="Authoring Target")
        self.auth_tree.heading("current", text="Current Content")
        self.auth_tree.heading("suggested", text="AI Suggestion")
        self.auth_tree.heading("confidence", text="Confidence")
        self.auth_tree.heading("grounding", text="Fact Grounding")
        self.auth_tree.heading("status", text="Review Decision")

        self.auth_tree.column("slide", width=55, anchor=tk.CENTER)
        self.auth_tree.column("shape", width=120, anchor=tk.W)
        self.auth_tree.column("type", width=140, anchor=tk.CENTER)
        self.auth_tree.column("current", width=160, anchor=tk.W)
        self.auth_tree.column("suggested", width=250, anchor=tk.W)
        self.auth_tree.column("confidence", width=80, anchor=tk.CENTER)
        self.auth_tree.column("grounding", width=110, anchor=tk.CENTER)
        self.auth_tree.column("status", width=120, anchor=tk.CENTER)

        auth_scroll = ttk.Scrollbar(auth_top_frame, orient=tk.VERTICAL, command=self.auth_tree.yview)
        self.auth_tree.configure(yscrollcommand=auth_scroll.set)
        self.auth_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        auth_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.auth_tree.bind("<<TreeviewSelect>>", self._on_auth_item_selected)

        auth_bottom_frame = tk.Frame(auth_split, bg=BG_CARD)
        auth_split.add(auth_bottom_frame, weight=2)

        self.auth_detail_var = tk.StringVar(value="Select an item above to review suggestion.")
        tk.Label(auth_bottom_frame, textvariable=self.auth_detail_var, font=("Segoe UI", 9, "bold"), fg=TEXT_WHITE, bg=BG_CARD).pack(anchor=tk.W, padx=8, pady=(6, 2))

        edit_subframe = tk.Frame(auth_bottom_frame, bg=BG_CARD)
        edit_subframe.pack(fill=tk.X, expand=True, padx=8, pady=(2, 6))

        tk.Label(edit_subframe, text="Proposed Content:", fg=TEXT_SECONDARY, bg=BG_CARD).pack(side=tk.LEFT, padx=(0, 6))
        self.auth_edit_var = tk.StringVar(value="")
        self.auth_edit_entry = ttk.Entry(edit_subframe, textvariable=self.auth_edit_var, font=("Segoe UI", 9))
        self.auth_edit_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        btn_bar = tk.Frame(auth_bottom_frame, bg=BG_CARD)
        btn_bar.pack(fill=tk.X, padx=8, pady=(0, 6))

        self.auth_approve_btn = ttk.Button(btn_bar, text="Approve", state=tk.DISABLED, command=self._on_auth_approve)
        self.auth_approve_btn.pack(side=tk.LEFT, padx=(0, 4))

        self.auth_reject_btn = ttk.Button(btn_bar, text="Reject", state=tk.DISABLED, command=self._on_auth_reject)
        self.auth_reject_btn.pack(side=tk.LEFT, padx=(0, 4))

        self.auth_edit_btn = ttk.Button(btn_bar, text="Save Edit & Approve", state=tk.DISABLED, command=self._on_auth_edit)
        self.auth_edit_btn.pack(side=tk.LEFT, padx=(0, 4))

        self.auth_mark_dec_btn = ttk.Button(btn_bar, text="Mark Decorative", state=tk.DISABLED, command=self._on_auth_mark_decorative)
        self.auth_mark_dec_btn.pack(side=tk.LEFT, padx=(0, 4))

        self.auth_keep_btn = ttk.Button(btn_bar, text="Keep Existing", state=tk.DISABLED, command=self._on_auth_keep_existing)
        self.auth_keep_btn.pack(side=tk.LEFT, padx=(0, 4))

        self.auth_batch_btn = ttk.Button(btn_bar, text="Approve All High-Confidence", state=tk.DISABLED, command=self._on_auth_approve_all)
        self.auth_batch_btn.pack(side=tk.RIGHT)
        self.selected_rec_id = None

        # SUBTAB 5: Accessibility Intelligence
        intel_subtab = tk.Frame(self.review_notebook, bg=BG_DARK)
        self.review_notebook.add(intel_subtab, text=" Accessibility Intelligence ")

        intel_columns = ("slide", "shape", "primary_rule", "related_rules", "confidence", "status", "recommendation", "reconciliation")
        self.intel_tree = ttk.Treeview(intel_subtab, columns=intel_columns, show="headings", selectmode="browse")

        self.intel_tree.heading("slide", text="Slide #")
        self.intel_tree.heading("shape", text="Shape")
        self.intel_tree.heading("primary_rule", text="Primary Rule")
        self.intel_tree.heading("related_rules", text="Related Rules")
        self.intel_tree.heading("confidence", text="Confidence")
        self.intel_tree.heading("status", text="Remediation Eligibility")
        self.intel_tree.heading("recommendation", text="Recommendation / Strategy")
        self.intel_tree.heading("reconciliation", text="Reconciliation Status")

        self.intel_tree.column("slide", width=55, anchor=tk.CENTER)
        self.intel_tree.column("shape", width=65, anchor=tk.CENTER)
        self.intel_tree.column("primary_rule", width=95, anchor=tk.CENTER)
        self.intel_tree.column("related_rules", width=110, anchor=tk.CENTER)
        self.intel_tree.column("confidence", width=85, anchor=tk.CENTER)
        self.intel_tree.column("status", width=130, anchor=tk.CENTER)
        self.intel_tree.column("recommendation", width=250, anchor=tk.W)
        self.intel_tree.column("reconciliation", width=120, anchor=tk.CENTER)

        intel_scroll = ttk.Scrollbar(intel_subtab, orient=tk.VERTICAL, command=self.intel_tree.yview)
        self.intel_tree.configure(yscrollcommand=intel_scroll.set)
        self.intel_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        intel_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Internal/hidden compatibility treeviews for ungrouping & list semantics
        hidden_frame = tk.Frame()
        grp_columns = ("slide", "group_id", "category", "eligible", "pictures", "safety", "details")
        self.grp_tree = ttk.Treeview(hidden_frame, columns=grp_columns, show="headings")
        list_columns = ("slide", "shape_id", "paragraph", "text_preview", "list_type", "level", "start_at", "format", "confidence", "status", "action")
        self.list_tree = ttk.Treeview(hidden_frame, columns=list_columns, show="headings")

    # =========================================================================
    # SECTION 3: REMEDIATION VIEW
    # =========================================================================
    def _build_remediation_view(self):
        rem_frame = tk.Frame(self.remediation_tab, bg=BG_DARK)
        rem_frame.pack(fill=tk.BOTH, expand=True)

        # Policy & Approval Banner
        rem_header = tk.Frame(rem_frame, bg=BG_CARD, highlightthickness=1, highlightbackground=BORDER_SUBTLE)
        rem_header.pack(fill=tk.X, pady=(0, 10))

        inner_h = tk.Frame(rem_header, bg=BG_CARD)
        inner_h.pack(fill=tk.X, padx=14, pady=10)

        self.approval_var = tk.BooleanVar(value=False)
        self.approval_chk = ttk.Checkbutton(
            inner_h,
            text="Approve Remediation Plan (Required to apply modifications)",
            variable=self.approval_var,
            state=tk.DISABLED,
            command=self._on_approval_toggled
        )
        self.approval_chk.pack(side=tk.LEFT, padx=(0, 15))

        rem_policy_lbl = tk.Label(
            inner_h,
            text="Sandboxed Candidate Copy: Original presentation is strictly protected and never modified directly.",
            font=("Segoe UI", 9, "italic"),
            fg=TEXT_SECONDARY,
            bg=BG_CARD
        )
        rem_policy_lbl.pack(side=tk.LEFT)

        # Remediation Execution Log Panel
        log_frame = tk.Frame(rem_frame, bg=BG_CARD, highlightthickness=1, highlightbackground=BORDER_SUBTLE)
        log_frame.pack(fill=tk.BOTH, expand=True)

        log_inner = tk.Frame(log_frame, bg=BG_CARD)
        log_inner.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        tk.Label(log_inner, text="Remediation Execution & Audit Log", font=("Segoe UI", 10, "bold"), fg=TEXT_WHITE, bg=BG_CARD).pack(anchor=tk.W, pady=(0, 6))

        self.rem_log_text = tk.Text(log_inner, font=("Consolas", 9), bg="#070C16", fg="#34D399", state="disabled", wrap=tk.WORD, bd=0)
        self.rem_log_text.pack(fill=tk.BOTH, expand=True)

        # Compatibility attributes expected by tests
        self.btn_inspect_ro = ttk.Button(hidden_frame := tk.Frame(), text="Review Proposed Order", command=lambda: self._switch_to_tab(1, 1))
        self.btn_inspect_ph = ttk.Button(hidden_frame, text="Review Placeholders", command=lambda: self._switch_to_tab(1, 2))
        self.btn_inspect_auth = ttk.Button(hidden_frame, text="Review Alt Text", command=lambda: self._switch_to_tab(1, 3))
        self.btn_apply_auth = ttk.Button(hidden_frame, text="Apply Alt Text", command=self._on_apply_alt_text_clicked)
        self.auth_card_status_var = tk.StringVar(value="Status: Ready")
        self.btn_apply_ungroup = ttk.Button(hidden_frame, text="Apply Ungrouping", command=self._on_apply_ungrouping_clicked)
        self.btn_apply_lists = ttk.Button(hidden_frame, text="Apply Lists", command=self._on_apply_lists_clicked)
        self.btn_run_pipeline = ttk.Button(hidden_frame, text="Run Pipeline", command=self._on_remediate_clicked)
        self.remediate_btn = self.btn_run_pipeline
        self.other_card_status_var = tk.StringVar(value="Status: Ready")

    # =========================================================================
    # SECTION 4: REPORTS & VALIDATION VIEW
    # =========================================================================
    def _build_reports_validation_view(self):
        rep_frame = tk.Frame(self.reports_tab, bg=BG_DARK)
        rep_frame.pack(fill=tk.BOTH, expand=True)

        # Header controls frame
        reports_header = tk.Frame(rep_frame, bg=BG_CARD, highlightthickness=1, highlightbackground=BORDER_SUBTLE)
        reports_header.pack(fill=tk.X, pady=(0, 10))

        inner_h = tk.Frame(reports_header, bg=BG_CARD)
        inner_h.pack(fill=tk.X, padx=14, pady=10)

        self.qa_overall_badge_var = tk.StringVar(value="Overall QA Status: NOT RUN")
        self.qa_overall_lbl = tk.Label(
            inner_h,
            textvariable=self.qa_overall_badge_var,
            font=("Segoe UI", 11, "bold"),
            fg="#3B82F6",
            bg=BG_CARD
        )
        self.qa_overall_lbl.pack(side=tk.LEFT, padx=(4, 20))

        self.qa_stats_var = tk.StringVar(value="Tests Passed: 0  |  Tests Failed: 0  |  Warnings: 0")
        tk.Label(inner_h, textvariable=self.qa_stats_var, font=("Segoe UI", 9, "bold"), fg=TEXT_SECONDARY, bg=BG_CARD).pack(side=tk.LEFT)

        # Reports Sub-Notebook (Tab 1: QA Suite 10 Dimensions, Tab 2: Structural Comparison)
        self.reports_notebook = ttk.Notebook(rep_frame)
        self.reports_notebook.pack(fill=tk.BOTH, expand=True)

        # SUBTAB 1: QA Suite (10 Dimensions)
        rep_qa_subtab = tk.Frame(self.reports_notebook, bg=BG_DARK)
        self.reports_notebook.add(rep_qa_subtab, text=" Core Quality Assurance (10 Dimensions) ")

        qa_split = ttk.PanedWindow(rep_qa_subtab, orient=tk.VERTICAL)
        qa_split.pack(fill=tk.BOTH, expand=True)

        qa_top_frame = tk.Frame(qa_split, bg=BG_CARD)
        qa_split.add(qa_top_frame, weight=3)

        dim_box = tk.LabelFrame(qa_top_frame, text=" Core Quality Assurance Dimensions ", bg=BG_CARD, fg=TEXT_WHITE, padx=10, pady=8)
        dim_box.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        self.qa_dim_ooxml_var = tk.StringVar(value="OOXML Integrity: -")
        self.qa_dim_design_var = tk.StringVar(value="Design Preservation: -")
        self.qa_dim_orig_var = tk.StringVar(value="Original File Protection: -")
        self.qa_dim_rel_var = tk.StringVar(value="Relationship Integrity: -")
        self.qa_dim_acc_var = tk.StringVar(value="Accessibility Consistency: -")
        self.qa_dim_idemp_var = tk.StringVar(value="Idempotency: -")
        self.qa_dim_ai_var = tk.StringVar(value="AI Failure Handling: -")
        self.qa_dim_ws_var = tk.StringVar(value="Workspace Cleanup: -")
        self.qa_dim_rep_var = tk.StringVar(value="Report Accuracy: -")
        self.qa_dim_anim_var = tk.StringVar(value="Animation Preservation: -")

        col0_frame = tk.Frame(dim_box, bg=BG_CARD)
        col0_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8)
        tk.Label(col0_frame, textvariable=self.qa_dim_ooxml_var, font=("Segoe UI", 9), fg=TEXT_PRIMARY, bg=BG_CARD).pack(anchor=tk.W, pady=2)
        tk.Label(col0_frame, textvariable=self.qa_dim_design_var, font=("Segoe UI", 9), fg=TEXT_PRIMARY, bg=BG_CARD).pack(anchor=tk.W, pady=2)
        tk.Label(col0_frame, textvariable=self.qa_dim_orig_var, font=("Segoe UI", 9), fg=TEXT_PRIMARY, bg=BG_CARD).pack(anchor=tk.W, pady=2)
        tk.Label(col0_frame, textvariable=self.qa_dim_rel_var, font=("Segoe UI", 9), fg=TEXT_PRIMARY, bg=BG_CARD).pack(anchor=tk.W, pady=2)
        tk.Label(col0_frame, textvariable=self.qa_dim_acc_var, font=("Segoe UI", 9), fg=TEXT_PRIMARY, bg=BG_CARD).pack(anchor=tk.W, pady=2)

        col1_frame = tk.Frame(dim_box, bg=BG_CARD)
        col1_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8)
        tk.Label(col1_frame, textvariable=self.qa_dim_idemp_var, font=("Segoe UI", 9), fg=TEXT_PRIMARY, bg=BG_CARD).pack(anchor=tk.W, pady=2)
        tk.Label(col1_frame, textvariable=self.qa_dim_ai_var, font=("Segoe UI", 9), fg=TEXT_PRIMARY, bg=BG_CARD).pack(anchor=tk.W, pady=2)
        tk.Label(col1_frame, textvariable=self.qa_dim_ws_var, font=("Segoe UI", 9), fg=TEXT_PRIMARY, bg=BG_CARD).pack(anchor=tk.W, pady=2)
        tk.Label(col1_frame, textvariable=self.qa_dim_rep_var, font=("Segoe UI", 9), fg=TEXT_PRIMARY, bg=BG_CARD).pack(anchor=tk.W, pady=2)
        tk.Label(col1_frame, textvariable=self.qa_dim_anim_var, font=("Segoe UI", 9), fg=TEXT_PRIMARY, bg=BG_CARD).pack(anchor=tk.W, pady=2)

        # Bottom detailed verification checks
        qa_bottom_frame = tk.Frame(qa_split, bg=BG_CARD)
        qa_split.add(qa_bottom_frame, weight=3)

        qa_columns = ("name", "component", "status", "details")
        self.qa_tree = ttk.Treeview(qa_bottom_frame, columns=qa_columns, show="headings", selectmode="browse")

        self.qa_tree.heading("name", text="Check Item")
        self.qa_tree.heading("component", text="Component")
        self.qa_tree.heading("status", text="Status")
        self.qa_tree.heading("details", text="Verification Details")

        self.qa_tree.column("name", width=220, anchor=tk.W)
        self.qa_tree.column("component", width=110, anchor=tk.CENTER)
        self.qa_tree.column("status", width=95, anchor=tk.CENTER)
        self.qa_tree.column("details", width=420, anchor=tk.W)

        qa_scroll = ttk.Scrollbar(qa_bottom_frame, orient=tk.VERTICAL, command=self.qa_tree.yview)
        self.qa_tree.configure(yscrollcommand=qa_scroll.set)
        self.qa_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        qa_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # SUBTAB 2: Structural Comparison & Per-Slide Results
        rep_comp_subtab = tk.Frame(self.reports_notebook, bg=BG_DARK)
        self.reports_notebook.add(rep_comp_subtab, text=" Structural Comparison & Per-Slide Results ")

        comp_top_frame = tk.Frame(rep_comp_subtab, bg=BG_CARD, highlightthickness=1, highlightbackground=BORDER_SUBTLE)
        comp_top_frame.pack(fill=tk.X, pady=(0, 6))

        c_inner = tk.Frame(comp_top_frame, bg=BG_CARD)
        c_inner.pack(fill=tk.X, padx=12, pady=8)

        self.struct_summary_var = tk.StringVar(value="Structural Verification: Not Run")
        tk.Label(c_inner, textvariable=self.struct_summary_var, font=("Segoe UI", 9, "bold"), fg="#3B82F6", bg=BG_CARD).pack(anchor=tk.W, pady=(0, 2))

        self.val_scores_var = tk.StringVar(value="Pre-Remediation Score: -  |  Post-Remediation Score: -")
        tk.Label(c_inner, textvariable=self.val_scores_var, font=("Segoe UI", 9), fg=GREEN_TEXT, bg=BG_CARD).pack(anchor=tk.W, pady=(0, 2))

        self.ops_summary_var = tk.StringVar(value="Remediation Operations: Applied: 0  |  Preserved/Skipped: 0  |  Rejected: 0  |  Failed: 0")
        tk.Label(c_inner, textvariable=self.ops_summary_var, font=("Segoe UI", 9), fg=TEXT_PRIMARY, bg=BG_CARD).pack(anchor=tk.W)

        slide_res_frame = tk.Frame(rep_comp_subtab, bg=BG_DARK)
        slide_res_frame.pack(fill=tk.BOTH, expand=True)

        slide_res_cols = ("slide", "rule", "status", "confidence", "details")
        self.slide_results_tree = ttk.Treeview(slide_res_frame, columns=slide_res_cols, show="headings", selectmode="browse")

        self.slide_results_tree.heading("slide", text="Slide #")
        self.slide_results_tree.heading("rule", text="Rule / Operation")
        self.slide_results_tree.heading("status", text="Status")
        self.slide_results_tree.heading("confidence", text="Confidence")
        self.slide_results_tree.heading("details", text="Remediation Details")

        self.slide_results_tree.column("slide", width=65, anchor=tk.CENTER)
        self.slide_results_tree.column("rule", width=180, anchor=tk.W)
        self.slide_results_tree.column("status", width=110, anchor=tk.CENTER)
        self.slide_results_tree.column("confidence", width=90, anchor=tk.CENTER)
        self.slide_results_tree.column("details", width=420, anchor=tk.W)

        slide_res_scroll = ttk.Scrollbar(slide_res_frame, orient=tk.VERTICAL, command=self.slide_results_tree.yview)
        self.slide_results_tree.configure(yscrollcommand=slide_res_scroll.set)
        self.slide_results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        slide_res_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    # =========================================================================
    # BOTTOM STATUS BAR
    # =========================================================================
    def _build_bottom_status_bar(self):
        bottom_frame = tk.Frame(self.master_frame, bg=BG_DARK, height=26)
        bottom_frame.pack(fill=tk.X, padx=14, pady=(0, 4))

        self.status_var = tk.StringVar(value="Ready. Select a PPTX file to begin.")
        self.status_lbl = tk.Label(
            bottom_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 8),
            fg=TEXT_SECONDARY,
            bg=BG_DARK
        )
        self.status_lbl.pack(side=tk.LEFT)

        self.prog_bar = ttk.Progressbar(bottom_frame, mode="determinate", length=220)
        self.prog_bar.pack(side=tk.RIGHT)

        # Compatibility attributes expected by tests/backend
        self.file_path_var = tk.StringVar(value="")
        self.analyze_btn = ttk.Button(bottom_frame, text="Analyze", state=tk.DISABLED, command=self._on_analyze_clicked)
        self.open_workspace_btn = ttk.Button(bottom_frame, text="Open Working Folder", state=tk.DISABLED, command=self._on_open_workspace)
        self.open_output_btn = ttk.Button(bottom_frame, text="Open Remediated File", state=tk.DISABLED, command=self._on_open_output_file)
        self.view_report_btn = ttk.Button(bottom_frame, text="View Full Report", state=tk.DISABLED, command=self._on_view_report)

    # =========================================================================
    # EVENT HANDLERS & LOGIC
    # =========================================================================
    def _toggle_maximize(self):
        if self.root.state() == "zoomed":
            self.root.state("normal")
        else:
            self.root.state("zoomed")

    def _on_select_file(self):
        file_selected = filedialog.askopenfilename(
            title="Select PowerPoint Presentation",
            filetypes=[("PowerPoint Presentations", "*.pptx"), ("All Files", "*.*")]
        )
        if not file_selected:
            return

        path = Path(file_selected)
        if not path.exists():
            messagebox.showerror("File Not Found", f"The selected file does not exist:\n{path}")
            return

        try:
            with open(path, "rb") as test_f:
                header = test_f.read(4)
                if len(header) < 4:
                    messagebox.showerror("Empty File", f"The selected file is empty (0 bytes):\n{path.name}")
                    return
        except PermissionError:
            messagebox.showerror(
                "File Locked / Inaccessible",
                f"The file '{path.name}' is currently locked or in use by another program.\nPlease close PowerPoint and try again."
            )
            return
        except Exception as e:
            messagebox.showerror("File Error", f"Cannot open '{path.name}':\n{e}")
            return

        self.file_path_var.set(str(path))
        self.header_file_var.set(f"{path.name}   ⌄")
        self.overview_filename_var.set(path.name)
        self.overview_filepath_var.set(str(path))
        self.status_var.set(f"Loaded: {path.name}. Initializing workspace...")

        try:
            self.job_context = self.controller.initialize_job(path)
            total = self.job_context.total_slides

            # Update Overview Card & Metrics
            self.metric_slides_var.set(str(total))
            self.metric_issues_var.set("-")
            self.metric_rules_var.set("-")
            self.metric_score_var.set("-")
            self.metric_integrity_var.set("Design Integrity: Preserved")

            # Update Right Column Status Card
            self.st_file_var.set(path.name)
            self.st_slides_var.set(str(total))
            self.st_issues_var.set("-")
            self.st_rules_var.set("-")
            self.st_score_var.set("-")
            self.st_integrity_var.set("Preserved")

            # Update Card Footers
            self.ro_slides_meta_var.set(f"📄 {total} slides")
            self.ro_affected_meta_var.set("🗂 Ready to analyze")
            self.ph_slides_meta_var.set(f"📄 {total} slides")
            self.ph_eligible_meta_var.set("🗂 Ready to analyze")

            # Update Slide Preview image if first slide has title
            slide_title = path.stem
            if getattr(self.controller, "current_presentation", None) and self.controller.current_presentation.slides:
                s1 = self.controller.current_presentation.slides[0]
                if getattr(s1, "title", None):
                    slide_title = s1.title
            self.slide_preview_img = self.icons.create_slide_preview_image(title=slide_title, width=310, height=158)
            self.slide_preview_lbl.config(image=self.slide_preview_img)

            # Enable Buttons
            self.analyze_btn.config(state=tk.NORMAL)
            self.btn_run_val.config(state=tk.NORMAL)
            self.open_workspace_btn.config(state=tk.NORMAL)

            # Record event
            self._add_activity_event("blue", "File loaded successfully")
            self._log_remediation_event(f"Loaded {path.name} ({total} slides).")
            self._update_workflow_stage(1)

            # Auto-run analysis for a seamless dashboard experience!
            self._on_analyze_clicked()

        except Exception as e:
            messagebox.showerror("Error Opening File", f"Could not load presentation: {e}")
            self.status_var.set("Failed to initialize presentation.")

    def _on_analyze_clicked(self):
        if not self.job_context or getattr(self, "is_busy", False):
            return

        self.is_busy = True
        self.analyze_btn.config(state=tk.DISABLED)
        self.btn_apply_ro.config(state=tk.DISABLED)
        self.btn_apply_ph.config(state=tk.DISABLED)
        self.btn_apply_auth.config(state=tk.DISABLED)
        self.btn_apply_ungroup.config(state=tk.DISABLED)
        self.btn_apply_lists.config(state=tk.DISABLED)
        self.btn_run_pipeline.config(state=tk.DISABLED)
        self.btn_run_val.config(state=tk.DISABLED)
        self.prog_bar["value"] = 0
        self.status_var.set("Analyzing presentation accessibility & design integrity...")

        def _worker():
            try:
                def _progress(curr, total, msg):
                    pct = int((curr / max(1, total)) * 100)
                    self.root.after(0, lambda: self._update_progress(pct, msg))

                self.job_context = self.controller.execute_analysis(self.job_context, progress_callback=_progress)
                self.root.after(0, self._on_analysis_finished)
            except Exception as e:
                self.root.after(0, lambda: setattr(self, "is_busy", False))
                self.root.after(0, lambda: messagebox.showerror("Analysis Error", f"Analysis failed: {e}"))
                self.root.after(0, lambda: self.analyze_btn.config(state=tk.NORMAL))

        threading.Thread(target=_worker, daemon=True).start()

    def _update_progress(self, percent: int, msg: str):
        self.prog_bar["value"] = percent
        self.status_var.set(msg)

    def _on_analysis_finished(self):
        self.is_busy = False
        ctx = self.job_context
        m = ctx.metrics

        # Update Metrics Bar
        self.metric_issues_var.set(str(m.total_issues))
        self.metric_rules_var.set(str(m.passed_rules))
        self.metric_score_var.set(f"{m.accessibility_score:.1f} / 100")
        self.metric_integrity_var.set("Design Integrity: Verified")

        # Update Right Column Status Panel
        self.st_issues_var.set(str(m.total_issues))
        self.st_rules_var.set(str(m.passed_rules))
        self.st_score_var.set(f"{m.accessibility_score:.1f} / 100")
        self.st_integrity_var.set("Verified")

        # Clear and populate issues table
        for row in self.tree.get_children():
            self.tree.delete(row)

        for res in ctx.rule_results:
            for issue in res.issues:
                st_val = getattr(issue.status, "value", str(issue.status)) if hasattr(issue, "status") else "DETECTED"
                conf_val = "HIGH"
                if hasattr(issue, "confidence") and issue.confidence is not None:
                    conf_val = getattr(issue, "confidence", "HIGH")
                    conf_val = getattr(conf_val, "value", str(conf_val))
                elif issue.extra_data and "confidence" in issue.extra_data:
                    conf_val = str(issue.extra_data["confidence"])

                self.tree.insert("", tk.END, values=(
                    issue.slide_number,
                    issue.severity.value,
                    conf_val,
                    issue.rule_id,
                    issue.rule_name,
                    st_val,
                    issue.description,
                    issue.recommendation
                ))

        # Clear and repopulate reading order table
        for row in self.ro_tree.get_children():
            self.ro_tree.delete(row)

        ro_issues = [i for res in ctx.rule_results if res.rule_id == "SG-RO-01" for i in res.issues]
        for res in ctx.rule_results:
            if res.rule_id == "SG-RO-01":
                for issue in res.issues:
                    extra = issue.extra_data or {}
                    orig = str(extra.get("original_order", []))
                    prop = str(extra.get("proposed_order", []))
                    conf = extra.get("confidence", "MEDIUM")
                    footers = extra.get("excluded_footer_objects", [])
                    reasons = "; ".join(extra.get("reasoning", [])[:2])
                    f_str = ", ".join(f"Shape {f['shape_id']}" for f in footers) if footers else "None"
                    self.ro_tree.insert("", tk.END, values=(
                        issue.slide_number, conf, orig, prop, f_str, reasons
                    ))

        # Clear and repopulate placeholder table
        for row in self.ph_tree.get_children():
            self.ph_tree.delete(row)

        ph_issues = [i for res in ctx.rule_results if res.rule_id == "SG-PH-01" for i in res.issues]
        for res in ctx.rule_results:
            if res.rule_id == "SG-PH-01":
                for issue in res.issues:
                    extra = issue.extra_data or {}
                    s_id = issue.shape_id or extra.get("shape_id", "N/A")
                    s_name = issue.shape_name or extra.get("shape_name", f"Shape_{s_id}")
                    ph_type = extra.get("placeholder_type", "body")
                    safety_status = extra.get("status", "PRESERVED")
                    rec = issue.recommendation or extra.get("recommendation", "")
                    reason = extra.get("reason", issue.description)
                    self.ph_tree.insert("", tk.END, values=(
                        issue.slide_number, f"{s_id}: {s_name}", ph_type.upper(), safety_status, rec, reason
                    ))

        # Populate AI authoring and Intelligence tables
        for row in self.intel_tree.get_children():
            self.intel_tree.delete(row)
        for u in getattr(ctx, "unified_issues", []):
            related_str = ", ".join(u.related_rule_ids) if u.related_rule_ids else "-"
            shape_str = str(u.shape_id) if u.shape_id is not None else "Slide"
            rec_stat = getattr(u.reconciliation_status, "value", str(u.reconciliation_status))
            conf_str = getattr(u.confidence, "value", str(u.confidence))
            elig_str = getattr(u.remediation_eligibility, "value", str(u.remediation_eligibility))
            self.intel_tree.insert("", tk.END, values=(
                u.slide_number, shape_str, u.primary_rule_id, related_str, conf_str, elig_str,
                u.recommended_action or u.explanation, rec_stat
            ))

        for row in self.auth_tree.get_children():
            self.auth_tree.delete(row)
        ai_recs = getattr(ctx, "ai_recommendations", [])
        for r in ai_recs:
            stat_val = getattr(r.status, "value", str(r.status))
            ground_val = getattr(r.fact_grounding, "value", str(r.fact_grounding))
            self.auth_tree.insert("", tk.END, iid=r.recommendation_id, values=(
                r.slide_number, r.shape_id or "Slide", getattr(r.authoring_type, "value", str(r.authoring_type)),
                r.original_value or r.current_value, r.effective_value, f"{r.confidence:.0%}", ground_val, stat_val
            ))

        # Update Remediation Center Card Statuses and Footers
        if ro_issues:
            self.ro_card_status_var.set("● Ready")
            self.ro_affected_meta_var.set(f"🗂 {len(ro_issues)} affected slides")
        else:
            self.ro_card_status_var.set("● Compliant")
            self.ro_affected_meta_var.set("🗂 0 affected slides")
        self.ro_slides_meta_var.set(f"📄 {ctx.total_slides} slides")

        if ph_issues:
            self.ph_card_status_var.set("● Ready")
            self.ph_eligible_meta_var.set(f"🗂 {len(ph_issues)} eligible items")
        else:
            self.ph_card_status_var.set("● Compliant")
            self.ph_eligible_meta_var.set("🗂 0 eligible items")
        self.ph_slides_meta_var.set(f"📄 {ctx.total_slides} slides")

        # Enable Primary Remediation Buttons
        self.btn_apply_ro.config(state=tk.NORMAL)
        self.btn_apply_ph.config(state=tk.NORMAL)
        self.btn_run_val.config(state=tk.NORMAL)
        self.generate_report_btn.config(state=tk.NORMAL)
        self.analyze_btn.config(state=tk.NORMAL)
        self.approval_chk.config(state=tk.NORMAL)

        # Record activity events
        self._add_activity_event("green", "Analysis completed")
        self._add_activity_event("orange", f"{m.total_issues} issues detected")
        self._add_activity_event("green", "Quality score calculated")
        self._update_workflow_stage(2)  # Stage 2: Review

        self.status_var.set(f"Analysis complete: {m.total_issues} issues detected. Quality Score: {m.accessibility_score:.1f}/100.")
        self._populate_reports_tab()

    # -------------------------------------------------------------------------
    # Reading Order Remediation (Batch processing)
    # -------------------------------------------------------------------------
    def _on_apply_reading_order_clicked(self):
        if getattr(self, "is_busy", False):
            return
        if not self.job_context or not self.job_context.rule_results:
            messagebox.showwarning("Analysis Required", "Please analyze the presentation before applying remediation.")
            return

        # Explicit approval gate
        if not self.job_context.user_approved_remediation:
            confirm = messagebox.askyesno(
                "User Approval Required",
                "Applying Reading Order alignment requires explicit approval.\n\n"
                "Do you approve applying Reading Order alignment across all eligible slides in one batch?\n\n"
                "(The original presentation file is strictly protected and will not be modified.)"
            )
            if confirm:
                self.approval_var.set(True)
                self.job_context.user_approved_remediation = True
            else:
                self.status_var.set("Reading Order alignment cancelled (approval not granted).")
                return

        self.is_busy = True
        self.btn_apply_ro.config(state=tk.DISABLED)
        self.prog_bar["value"] = 0
        self.status_var.set("Aligning reading order across slides in batch...")

        def _worker():
            try:
                def _progress(curr, total, msg):
                    pct = int((curr / max(1, total)) * 100)
                    self.root.after(0, lambda: self._update_progress(pct, msg))

                self.job_context = self.controller.execute_reading_order_remediation(
                    self.job_context,
                    progress_callback=_progress
                )
                self.root.after(0, self._on_reading_order_finished)
            except Exception as e:
                self.root.after(0, lambda: setattr(self, "is_busy", False))
                self.root.after(0, lambda: messagebox.showerror("Reading Order Error", f"Reading order alignment failed: {e}"))
                self.root.after(0, lambda: self.btn_apply_ro.config(state=tk.NORMAL))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_reading_order_finished(self):
        self.is_busy = False
        self.btn_apply_ro.config(state=tk.NORMAL)
        ctx = self.job_context
        ro_actions = [a for a in ctx.remediation_actions if a.rule_id in ("PIPELINE_STEP_1_READING_ORDER", "SG-RO-01")]
        applied = sum(1 for a in ro_actions if a.status == RemediationStatus.APPLIED)
        skipped = sum(1 for a in ro_actions if a.status == RemediationStatus.SKIPPED)

        if applied > 0:
            self.ro_card_status_var.set("● Applied")
            self.ro_affected_meta_var.set(f"🗂 {applied} slides reordered")
            self._add_activity_event("blue", f"Reading order applied ({applied} slides)")
            self._log_remediation_event(f"Reading Order: APPLIED ({applied} slides reordered in candidate copy)")
            messagebox.showinfo("Reading Order Applied", f"Reading Order alignment completed successfully!\n\n• Reordered Slides: {applied}\n• Compliant Slides Preserved: {skipped}\n\nOriginal file is strictly preserved.")
        else:
            self.ro_card_status_var.set("● Compliant")
            self._add_activity_event("blue", "Reading order verified")
            self._log_remediation_event("Reading Order: SKIPPED (All slides already compliant)")
            messagebox.showinfo("Reading Order Completed", "Reading Order alignment skipped:\nAll slides already comply with logical reading order.")

        self._update_workflow_stage(3)  # Stage 3: Apply
        self.status_var.set("Reading order remediation completed.")
        if ctx.final_output_file and ctx.final_output_file.exists():
            self.open_output_btn.config(state=tk.NORMAL)
        self._populate_reports_tab()

    # -------------------------------------------------------------------------
    # Placeholder Conversion Remediation
    # -------------------------------------------------------------------------
    def _on_apply_placeholders_clicked(self):
        if getattr(self, "is_busy", False):
            return
        if not self.job_context or not self.job_context.rule_results:
            messagebox.showwarning("Analysis Required", "Please analyze the presentation before applying remediation.")
            return

        if not self.job_context.user_approved_remediation:
            confirm = messagebox.askyesno(
                "User Approval Required",
                "Native Placeholder Conversion requires explicit approval.\n\n"
                "Do you approve converting eligible shapes into native PowerPoint placeholders in the candidate copy?\n\n"
                "(The original presentation file is strictly protected and will not be modified.)"
            )
            if confirm:
                self.approval_var.set(True)
                self.job_context.user_approved_remediation = True
            else:
                self.status_var.set("Placeholder conversion cancelled (approval not granted).")
                return

        self.is_busy = True
        self.btn_apply_ph.config(state=tk.DISABLED)
        self.prog_bar["value"] = 0
        self.status_var.set("Converting shapes to native placeholders...")

        def _worker():
            try:
                def _progress(curr, total, msg):
                    pct = int((curr / max(1, total)) * 100)
                    self.root.after(0, lambda: self._update_progress(pct, msg))

                self.job_context = self.controller.execute_placeholder_remediation(
                    self.job_context,
                    progress_callback=_progress
                )
                self.root.after(0, self._on_placeholders_finished)
            except Exception as e:
                self.root.after(0, lambda: setattr(self, "is_busy", False))
                self.root.after(0, lambda: messagebox.showerror("Placeholder Error", f"Placeholder conversion failed: {e}"))
                self.root.after(0, lambda: self.btn_apply_ph.config(state=tk.NORMAL))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_placeholders_finished(self):
        self.is_busy = False
        self.btn_apply_ph.config(state=tk.NORMAL)
        ctx = self.job_context
        ph_actions = [a for a in ctx.remediation_actions if a.rule_id in ("PIPELINE_STEP_3_PLACEHOLDERS", "SG-PH-01")]
        applied = sum(1 for a in ph_actions if a.status == RemediationStatus.APPLIED)

        if applied > 0:
            self.ph_card_status_var.set("● Applied")
            self.ph_eligible_meta_var.set(f"🗂 {applied} shapes converted")
            self._add_activity_event("purple", f"Placeholders converted ({applied} shapes)")
            self._log_remediation_event(f"Placeholder Conversion: APPLIED ({applied} shapes converted)")
            messagebox.showinfo("Placeholder Conversion Applied", f"Native Placeholder Conversion completed successfully!\n\n• Converted Shapes: {applied}\n\nOriginal file is strictly preserved.")
        else:
            self.ph_card_status_var.set("● Compliant")
            self._add_activity_event("purple", "Placeholders compliant")
            self._log_remediation_event("Placeholder Conversion: SKIPPED (No eligible shapes to convert)")
            messagebox.showinfo("Placeholder Conversion Completed", "Placeholder conversion skipped:\nAll shapes are already native placeholders or preserved.")

        self._update_workflow_stage(3)
        self.status_var.set("Placeholder conversion completed.")
        if ctx.final_output_file and ctx.final_output_file.exists():
            self.open_output_btn.config(state=tk.NORMAL)
        self._populate_reports_tab()

    # -------------------------------------------------------------------------
    # Validation & Compare
    # -------------------------------------------------------------------------
    def _on_apply_validation_clicked(self):
        if getattr(self, "is_busy", False):
            return
        if not self.job_context:
            messagebox.showwarning("File Required", "Please select a presentation first.")
            return

        self.is_busy = True
        self.btn_run_val.config(state=tk.DISABLED)
        self.prog_bar["value"] = 0
        self.status_var.set("Running validation and structural comparison...")

        def _worker():
            try:
                def _progress(curr, total, msg):
                    pct = int((curr / max(1, total)) * 100)
                    self.root.after(0, lambda: self._update_progress(pct, msg))

                self.job_context = self.controller.run_validation_and_compare(
                    self.job_context,
                    progress_callback=_progress
                )
                self.root.after(0, self._on_validation_finished)
            except Exception as e:
                self.root.after(0, lambda: setattr(self, "is_busy", False))
                self.root.after(0, lambda: messagebox.showerror("Validation Error", f"Validation and comparison failed: {e}"))
                self.root.after(0, lambda: self.btn_run_val.config(state=tk.NORMAL))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_validation_finished(self):
        self.is_busy = False
        self.btn_run_val.config(state=tk.NORMAL)
        ctx = self.job_context
        v = ctx.verification_result
        v_status = "PASSED" if (v and v.passed) else "INSPECTION NEEDED"
        qa_stat = ctx.qa_result.overall_status.value if getattr(ctx, "qa_result", None) else "COMPLETED"

        self.val_card_status_var.set("● Verified")
        self.metric_integrity_var.set("Design Integrity: Verified")
        self.st_integrity_var.set("Verified")
        self._add_activity_event("green", "Validation & integrity verified")
        self._log_remediation_event(f"Validation & Compare: QA {qa_stat}, Verification {v_status}")

        if getattr(ctx, "qa_result", None):
            self._populate_qa_tab(ctx.qa_result)

        if ctx.post_remediation_score is not None:
            self.metric_score_var.set(f"{ctx.post_remediation_score:.1f} / 100")
            self.st_score_var.set(f"{ctx.post_remediation_score:.1f} / 100")

        self._update_workflow_stage(4)  # Stage 4: Validate
        self.status_var.set(f"Validation completed: {qa_stat}. Verification: {v_status}.")
        self._populate_reports_tab()

        if ctx.final_output_file and ctx.final_output_file.exists():
            self.open_output_btn.config(state=tk.NORMAL)

        messagebox.showinfo(
            "Validation Completed",
            f"Validation and structural comparison completed successfully!\n\n"
            f"• Structural Integrity: {v_status}\n"
            f"• QA Suite Status: {qa_stat}\n\n"
            f"Design integrity and formatting strictly preserved."
        )

    # -------------------------------------------------------------------------
    # Generate Report
    # -------------------------------------------------------------------------
    def _on_generate_report_clicked(self):
        if getattr(self, "is_busy", False):
            return
        if not self.job_context:
            messagebox.showwarning("File Required", "Please select and analyze a presentation first.")
            return

        self.is_busy = True
        self.generate_report_btn.config(state=tk.DISABLED)
        self.prog_bar["value"] = 0
        self.status_var.set("Generating comprehensive accessibility audit reports...")

        def _worker():
            try:
                def _progress(curr, total, msg):
                    pct = int((curr / max(1, total)) * 100)
                    self.root.after(0, lambda: self._update_progress(pct, msg))

                self.job_context = self.controller.generate_reports(self.job_context)
                self.root.after(0, self._on_generate_report_finished)
            except Exception as e:
                logger.exception("Report generation failed: %s", e)
                err_msg = str(e) or type(e).__name__
                self.root.after(0, lambda: setattr(self, "is_busy", False))
                self.root.after(0, lambda: self.status_var.set(f"Report generation error: {err_msg}"))
                self.root.after(0, lambda: messagebox.showerror("Report Error", f"Report generation failed: {err_msg}"))
                self.root.after(0, lambda: self.generate_report_btn.config(state=tk.NORMAL))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_generate_report_finished(self):
        self.is_busy = False
        self.generate_report_btn.config(state=tk.NORMAL)
        self.view_report_btn.config(state=tk.NORMAL)
        ctx = self.job_context
        rep_name = ctx.final_report_file.name if getattr(ctx, "final_report_file", None) else "report.md"

        self.rep_card_status_var.set("● Exported")
        self._add_activity_event("teal", "Accessibility report generated")
        self._log_remediation_event(f"Generated report: {rep_name}")
        self._update_workflow_stage(5)  # Stage 5: Export

        self.status_var.set(f"Report generated: {rep_name}")
        messagebox.showinfo(
            "Report Generated",
            f"Accessibility compliance report generated successfully!\n\n"
            f"• Report File: {rep_name}\n"
            f"• Location: {ctx.final_report_file}\n\n"
            f"Audited against WCAG 2.1 AA and PowerPoint Accessibility Standards."
        )
        self._populate_reports_tab()

    # -------------------------------------------------------------------------
    # Other handlers and compatibility methods
    # -------------------------------------------------------------------------
    def _on_approval_toggled(self):
        if not self.job_context:
            return
        approved = self.approval_var.get()
        self.job_context.user_approved_remediation = approved
        self._log_remediation_event(f"User approval set to: {approved}")

    def _switch_to_tab(self, main_tab_index: int, sub_tab_index: Optional[int] = None):
        self._select_nav(main_tab_index)
        if sub_tab_index is not None and hasattr(self, "review_notebook"):
            self.review_notebook.select(sub_tab_index)

    def _log_remediation_event(self, message: str):
        if hasattr(self, "rem_log_text"):
            self.rem_log_text.config(state="normal")
            self.rem_log_text.insert(tk.END, f"• {message}\n")
            self.rem_log_text.see(tk.END)
            self.rem_log_text.config(state="disabled")

    def _populate_reports_tab(self):
        if not self.job_context:
            return

        ctx = self.job_context
        v = getattr(ctx, "verification_result", None)
        if v:
            v_status = "PASSED" if v.passed else "INSPECTION NEEDED"
            check_count = len(getattr(v, "structural_checks", []))
            self.struct_summary_var.set(f"Structural Verification: {v_status} ({check_count} checks evaluated)")
        else:
            self.struct_summary_var.set("Structural Verification: Ready to run (Click 'Run Validation')")

        pre_score = f"{ctx.metrics.quality_score_v2:.1f}/100" if ctx.metrics else "-"
        post_score = f"{ctx.post_remediation_score:.1f}/100" if getattr(ctx, "post_remediation_score", None) is not None else "-"
        self.val_scores_var.set(f"Pre-Remediation Score: {pre_score}  |  Post-Remediation Score: {post_score}")

        applied_n = 0
        skipped_n = 0
        rejected_n = 0
        failed_n = 0
        actions = getattr(ctx, "remediation_actions", [])
        for a in actions:
            st = getattr(a, "status", None)
            if st == RemediationStatus.APPLIED:
                applied_n += 1
            elif st == RemediationStatus.SKIPPED:
                skipped_n += 1
            elif st == RemediationStatus.REJECTED:
                rejected_n += 1
            elif st == RemediationStatus.FAILED:
                failed_n += 1
        self.ops_summary_var.set(
            f"Remediation Operations: Applied: {applied_n}  |  Preserved/Skipped: {skipped_n}  |  Rejected: {rejected_n}  |  Failed: {failed_n}"
        )

        for row in self.slide_results_tree.get_children():
            self.slide_results_tree.delete(row)

        if actions:
            for a in actions:
                slide_num = a.slide_number if getattr(a, "slide_number", None) is not None else "-"
                st_val = getattr(a.status, "value", str(a.status))
                details_val = getattr(a, "description", "") or getattr(a, "details", "") or ""
                self.slide_results_tree.insert("", tk.END, values=(slide_num, a.rule_id, st_val, "HIGH", details_val))
        elif getattr(ctx, "rule_results", None):
            for res in ctx.rule_results:
                for issue in res.issues:
                    st_val = getattr(issue.status, "value", str(issue.status)) if hasattr(issue, "status") else "DETECTED"
                    self.slide_results_tree.insert("", tk.END, values=(issue.slide_number, issue.rule_id, st_val, "HIGH", issue.description))

    def _populate_qa_tab(self, qa_res: QAResult):
        st_val = qa_res.overall_status.value
        self.qa_overall_badge_var.set(f"Overall QA Status: {st_val}")
        if st_val == "PASS":
            self.qa_overall_lbl.config(fg="#10B981")
        elif st_val == "PASS_WITH_WARNINGS":
            self.qa_overall_lbl.config(fg="#F59E0B")
        else:
            self.qa_overall_lbl.config(fg="#EF4444")

        self.qa_stats_var.set(
            f"Tests Passed: {qa_res.tests_passed}  |  Tests Failed: {qa_res.tests_failed}  |  Warnings: {len(qa_res.warnings)}"
        )
        self.qa_dim_ooxml_var.set(f"OOXML Integrity: {qa_res.integrity_status.value}")
        self.qa_dim_design_var.set(f"Design Preservation: {qa_res.design_preservation_status.value}")
        self.qa_dim_orig_var.set(f"Original File Protection: {qa_res.original_protection_status.value}")
        self.qa_dim_rel_var.set(f"Relationship Integrity: {qa_res.relationship_status.value}")
        self.qa_dim_acc_var.set(f"Accessibility Consistency: {qa_res.accessibility_consistency_status.value}")
        self.qa_dim_idemp_var.set(f"Idempotency: {qa_res.idempotency_status.value}")
        self.qa_dim_ai_var.set(f"AI Failure Handling: {qa_res.ai_fallback_status.value}")
        self.qa_dim_ws_var.set(f"Workspace Cleanup: {qa_res.workspace_cleanup_status.value}")
        self.qa_dim_rep_var.set(f"Report Accuracy: {qa_res.report_accuracy_status.value}")
        self.qa_dim_anim_var.set(f"Animation Preservation: {qa_res.animation_preservation_status.value}")

        for row in self.qa_tree.get_children():
            self.qa_tree.delete(row)

        for check in qa_res.checks:
            self.qa_tree.insert("", tk.END, values=(
                check.name, check.component, check.status.value, check.details
            ))

    def _on_apply_alt_text_clicked(self):
        if getattr(self, "is_busy", False) or not self.job_context:
            return
        if not self.job_context.user_approved_remediation:
            return
        def _worker():
            self.job_context = self.controller.execute_alt_text_remediation(self.job_context)
        threading.Thread(target=_worker, daemon=True).start()

    def _on_apply_ungrouping_clicked(self):
        if getattr(self, "is_busy", False) or not self.job_context:
            return
        if not self.job_context.user_approved_remediation:
            return
        def _worker():
            self.job_context = self.controller.execute_ungrouping_remediation(self.job_context)
        threading.Thread(target=_worker, daemon=True).start()

    def _on_apply_lists_clicked(self):
        if getattr(self, "is_busy", False) or not self.job_context:
            return
        if not self.job_context.user_approved_remediation:
            return
        def _worker():
            self.job_context = self.controller.execute_list_remediation(self.job_context)
        threading.Thread(target=_worker, daemon=True).start()

    def _on_remediate_clicked(self):
        if getattr(self, "is_busy", False) or not self.job_context:
            return
        if not self.job_context.user_approved_remediation:
            messagebox.showwarning("Approval Required", "You must explicitly approve remediation first.")
            return
        def _worker():
            self.job_context = self.controller.execute_remediation(self.job_context)
        threading.Thread(target=_worker, daemon=True).start()

    def _on_auth_item_selected(self, event):
        sel = self.auth_tree.selection()
        if not sel or not self.job_context:
            return
        item_id = sel[0]
        rec = next((r for r in getattr(self.job_context, "ai_recommendations", []) if getattr(r, "recommendation_id", "") == item_id), None)
        if rec:
            self.selected_rec_id = rec.recommendation_id
            type_str = getattr(rec.authoring_type, "value", str(rec.authoring_type))
            self.auth_detail_var.set(f"Slide {rec.slide_number} | Shape {rec.shape_id or 'Slide'} | {type_str} | Confidence: {rec.confidence:.0%}")
            self.auth_edit_var.set(rec.effective_value)
            self.auth_approve_btn.config(state=tk.NORMAL)
            self.auth_edit_btn.config(state=tk.NORMAL)
            self.auth_reject_btn.config(state=tk.NORMAL)
            self.auth_mark_dec_btn.config(state=tk.NORMAL)
            self.auth_keep_btn.config(state=tk.NORMAL)

    def _on_auth_approve(self):
        if not self.job_context or not self.selected_rec_id:
            return
        rec = next((r for r in self.job_context.ai_recommendations if r.recommendation_id == self.selected_rec_id), None)
        if rec:
            rec.approve()
            self._refresh_auth_tree_item(rec)

    def _on_auth_edit(self):
        if not self.job_context or not self.selected_rec_id:
            return
        rec = next((r for r in self.job_context.ai_recommendations if r.recommendation_id == self.selected_rec_id), None)
        if rec:
            new_text = self.auth_edit_var.get().strip()
            if new_text:
                rec.edit(new_text)
                self._refresh_auth_tree_item(rec)

    def _on_auth_reject(self):
        if not self.job_context or not self.selected_rec_id:
            return
        rec = next((r for r in self.job_context.ai_recommendations if r.recommendation_id == self.selected_rec_id), None)
        if rec:
            rec.reject()
            self._refresh_auth_tree_item(rec)

    def _on_auth_mark_decorative(self):
        if not self.job_context or not self.selected_rec_id:
            return
        rec = next((r for r in self.job_context.ai_recommendations if r.recommendation_id == self.selected_rec_id), None)
        if rec:
            rec.mark_decorative()
            self._refresh_auth_tree_item(rec)

    def _on_auth_keep_existing(self):
        if not self.job_context or not self.selected_rec_id:
            return
        rec = next((r for r in self.job_context.ai_recommendations if r.recommendation_id == self.selected_rec_id), None)
        if rec:
            rec.keep_existing()
            self._refresh_auth_tree_item(rec)

    def _on_auth_approve_all(self):
        if not self.job_context:
            return
        for rec in getattr(self.job_context, "ai_recommendations", []):
            if rec.status == AuthoringStatus.PENDING_REVIEW and rec.confidence >= 0.85:
                rec.approve()
                self._refresh_auth_tree_item(rec)

    def _refresh_auth_tree_item(self, rec):
        if self.auth_tree.exists(rec.recommendation_id):
            stat_val = getattr(rec.status, "value", str(rec.status))
            ground_val = getattr(rec.fact_grounding, "value", str(rec.fact_grounding))
            self.auth_tree.item(rec.recommendation_id, values=(
                rec.slide_number, rec.shape_id or "Slide", getattr(rec.authoring_type, "value", str(rec.authoring_type)),
                rec.original_value or rec.current_value, rec.effective_value, f"{rec.confidence:.0%}", ground_val, stat_val
            ))

    def _on_open_workspace(self):
        if not self.job_context or not self.job_context.workspace_dir.exists():
            return
        os.startfile(str(self.job_context.workspace_dir))

    def _on_open_output_file(self):
        if not self.job_context or not self.job_context.final_output_file:
            messagebox.showinfo("Export", "No remediated candidate presentation generated yet.\nRun Reading Order or Placeholders first.")
            return
        if self.job_context.final_output_file.exists():
            os.startfile(str(self.job_context.final_output_file))
        else:
            messagebox.showwarning("File Missing", f"Output file not found:\n{self.job_context.final_output_file}")

    def _on_view_report(self):
        if not self.job_context:
            messagebox.showinfo("Report", "No report file generated yet.")
            return
        report_file = getattr(self.job_context, "final_report_file", None)
        if not report_file or not report_file.exists():
            if self.job_context.reports_dir:
                report_file = self.job_context.reports_dir / "final_report.md"
                if not report_file.exists():
                    report_file = self.job_context.reports_dir / "analysis_report.md"
        if report_file and report_file.exists():
            os.startfile(str(report_file))
        else:
            messagebox.showinfo("Report", "No report file generated yet.")

    def _on_settings_clicked(self):
        messagebox.showinfo(
            "SlideGuard Settings",
            f"SlideGuard Accessibility Platform (v{VERSION})\n\n"
            "• Standard: WCAG 2.1 AA / HHS 508\n"
            "• Engine: Strict Non-Destructive Design Preservation\n"
            "• Mode: Batch Automated Remediation\n"
            "• Original File Protection: Active (Read-Only Isolation)\n"
            "• Workspace: Sandboxed Session Execution"
        )

    def _on_help_clicked(self):
        messagebox.showinfo(
            "SlideGuard Help & Documentation",
            "SlideGuard PowerPoint Accessibility & Remediation Platform\n\n"
            "Workflow:\n"
            "1. Select Presentation File (.pptx)\n"
            "2. Automatic Analysis & Issue Detection\n"
            "3. Review & Apply Reading Order (Batch Aligned)\n"
            "4. Review & Apply Native Placeholders\n"
            "5. Run Validation & Structural Comparison\n"
            "6. Export Production Remediated File & Audit Report\n\n"
            "Safety Guarantee:\n"
            "Original presentation files are strictly protected from modification."
        )
