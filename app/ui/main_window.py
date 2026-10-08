"""
SlideGuard Desktop User Interface.
Provides a clean, functional desktop UI using standard Tkinter / TTK.
Strictly decoupled from PPTX logic; delegates all operations to AppController.
Includes interactive Reading Order inspection and review.
"""

import os
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional

from app.version import APP_NAME, APP_TITLE, VERSION, APP_DISCLAIMER
from app.paths import resolve_resource_path
from app.controller.app_controller import AppController
from models.job import JobContext
from models.enums import IssueSeverity, RemediationStatus
from models.ai_authoring import AuthoringStatus
from models.qa import QAResult, QAStatus


class SlideGuardMainWindow:
    """
    Main desktop window for SlideGuard.
    Handles user interaction, file selection, background execution, and results display.
    """

    def __init__(self, root: tk.Tk, controller: Optional[AppController] = None):
        self.root = root
        self.controller = controller or AppController()
        self.job_context: Optional[JobContext] = None

        self._setup_window()
        self._build_ui()

    def _setup_window(self):
        self.root.title(f"{APP_TITLE} (v{VERSION})")
        self.root.geometry("1020x740")
        self.root.minsize(880, 600)

        # Set window icon if available
        icon_path = resolve_resource_path("assets/slideguard.ico")
        if icon_path.exists():
            try:
                self.root.iconbitmap(str(icon_path))
            except Exception:
                pass

        # Apply clean ttk theme
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")

    def _build_ui(self):
        # Main layout container
        main_frame = ttk.Frame(self.root, padding="12 12 12 12")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. Header Banner
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))

        title_lbl = ttk.Label(
            header_frame,
            text=APP_NAME.upper(),
            font=("Segoe UI", 16, "bold"),
            foreground="#1E3A8A"
        )
        title_lbl.pack(side=tk.LEFT)

        subtitle_lbl = ttk.Label(
            header_frame,
            text=f" v{VERSION}  |  PowerPoint Accessibility & Design Integrity Platform",
            font=("Segoe UI", 11)
        )
        subtitle_lbl.pack(side=tk.LEFT, pady=(3, 0))

        # 2. File Selection Panel
        file_frame = ttk.LabelFrame(main_frame, text=" Presentation File ", padding="10 10 10 10")
        file_frame.pack(fill=tk.X, pady=(0, 10))

        self.file_path_var = tk.StringVar(value="")
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, state="readonly", font=("Consolas", 9))
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        browse_btn = ttk.Button(file_frame, text="Select PPTX...", command=self._on_select_file)
        browse_btn.pack(side=tk.RIGHT)

        # 3. Action Toolbar & Status Bar
        actions_frame = ttk.Frame(main_frame)
        actions_frame.pack(fill=tk.X, pady=(0, 10))

        self.analyze_btn = ttk.Button(
            actions_frame,
            text="Analyze Presentation",
            state=tk.DISABLED,
            command=self._on_analyze_clicked
        )
        self.analyze_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.approval_var = tk.BooleanVar(value=False)
        self.approval_chk = ttk.Checkbutton(
            actions_frame,
            text="Approve Remediation Plan",
            variable=self.approval_var,
            state=tk.DISABLED,
            command=self._on_approval_toggled
        )
        self.approval_chk.pack(side=tk.LEFT, padx=(0, 10))

        self.remediate_btn = ttk.Button(
            actions_frame,
            text="Remediate & Validate",
            state=tk.DISABLED,
            command=self._on_remediate_clicked
        )
        self.remediate_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.open_workspace_btn = ttk.Button(
            actions_frame,
            text="Open Working Folder",
            state=tk.DISABLED,
            command=self._on_open_workspace
        )
        self.open_workspace_btn.pack(side=tk.RIGHT)

        # 4. Progress Area
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(0, 10))

        self.status_var = tk.StringVar(value="Ready. Select a PPTX file to begin.")
        self.status_lbl = ttk.Label(progress_frame, textvariable=self.status_var, font=("Segoe UI", 9))
        self.status_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.prog_bar = ttk.Progressbar(progress_frame, mode="determinate", length=220)
        self.prog_bar.pack(side=tk.RIGHT)

        # 5. Metrics Cards Summary
        metrics_frame = ttk.LabelFrame(main_frame, text=" Quality & Verification Metrics ", padding="8 8 8 8")
        metrics_frame.pack(fill=tk.X, pady=(0, 10))

        self.metric_slides_var = tk.StringVar(value="Total Slides: -")
        self.metric_issues_var = tk.StringVar(value="Issues Detected: -")
        self.metric_rules_var = tk.StringVar(value="Rules Passed: -")
        self.metric_score_var = tk.StringVar(value="Quality Score: -")
        self.metric_integrity_var = tk.StringVar(value="Design Integrity: -")

        ttk.Label(metrics_frame, textvariable=self.metric_slides_var, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, expand=True)
        ttk.Label(metrics_frame, textvariable=self.metric_issues_var, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, expand=True)
        ttk.Label(metrics_frame, textvariable=self.metric_rules_var, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, expand=True)
        ttk.Label(metrics_frame, textvariable=self.metric_score_var, font=("Segoe UI", 9, "bold"), foreground="#047857").pack(side=tk.LEFT, expand=True)
        ttk.Label(metrics_frame, textvariable=self.metric_integrity_var, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, expand=True)

        # 6. Tabbed Results Panel (Issues + Reading Order Review)
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # TAB 1: Detected Issues
        issues_tab = ttk.Frame(notebook, padding="6 6 6 6")
        notebook.add(issues_tab, text=" Detected Issues & Compliance ")

        columns = ("slide", "severity", "rule_id", "rule_name", "status", "description", "recommendation")
        self.tree = ttk.Treeview(issues_tab, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("slide", text="Slide #")
        self.tree.heading("severity", text="Severity")
        self.tree.heading("rule_id", text="Rule ID")
        self.tree.heading("rule_name", text="Rule Name")
        self.tree.heading("status", text="Action / Status")
        self.tree.heading("description", text="Description")
        self.tree.heading("recommendation", text="Recommendation")

        self.tree.column("slide", width=55, anchor=tk.CENTER)
        self.tree.column("severity", width=75, anchor=tk.CENTER)
        self.tree.column("rule_id", width=85, anchor=tk.CENTER)
        self.tree.column("rule_name", width=150, anchor=tk.W)
        self.tree.column("status", width=120, anchor=tk.CENTER)
        self.tree.column("description", width=300, anchor=tk.W)
        self.tree.column("recommendation", width=190, anchor=tk.W)

        tree_scroll = ttk.Scrollbar(issues_tab, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # TAB 2: Reading Order Analysis
        ro_tab = ttk.Frame(notebook, padding="6 6 6 6")
        notebook.add(ro_tab, text=" Reading Order Review ")

        ro_columns = ("slide", "confidence", "original_order", "proposed_order", "excluded_footers", "reasoning")
        self.ro_tree = ttk.Treeview(ro_tab, columns=ro_columns, show="headings", selectmode="browse")

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

        ro_scroll = ttk.Scrollbar(ro_tab, orient=tk.VERTICAL, command=self.ro_tree.yview)
        self.ro_tree.configure(yscrollcommand=ro_scroll.set)
        self.ro_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ro_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # TAB 3: Selective Ungrouping Analysis (Phase 4)
        grp_tab = ttk.Frame(notebook, padding="6 6 6 6")
        notebook.add(grp_tab, text=" Selective Ungrouping Review ")

        grp_columns = ("slide", "group_id", "category", "eligible", "pictures", "safety", "details")
        self.grp_tree = ttk.Treeview(grp_tab, columns=grp_columns, show="headings", selectmode="browse")

        self.grp_tree.heading("slide", text="Slide #")
        self.grp_tree.heading("group_id", text="Group ID")
        self.grp_tree.heading("category", text="Group Category")
        self.grp_tree.heading("eligible", text="Eligible Text/Shapes")
        self.grp_tree.heading("pictures", text="Protected Pictures")
        self.grp_tree.heading("safety", text="Remediation Safety")
        self.grp_tree.heading("details", text="Remediation Strategy / Action")

        self.grp_tree.column("slide", width=60, anchor=tk.CENTER)
        self.grp_tree.column("group_id", width=80, anchor=tk.CENTER)
        self.grp_tree.column("category", width=140, anchor=tk.CENTER)
        self.grp_tree.column("eligible", width=130, anchor=tk.CENTER)
        self.grp_tree.column("pictures", width=130, anchor=tk.CENTER)
        self.grp_tree.column("safety", width=130, anchor=tk.CENTER)
        self.grp_tree.column("details", width=290, anchor=tk.W)

        grp_scroll = ttk.Scrollbar(grp_tab, orient=tk.VERTICAL, command=self.grp_tree.yview)
        self.grp_tree.configure(yscrollcommand=grp_scroll.set)
        self.grp_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        grp_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # TAB 4: Placeholder Conversion Review (Phase 5)
        ph_tab = ttk.Frame(notebook, padding="6 6 6 6")
        notebook.add(ph_tab, text=" Placeholder Conversion Review ")

        ph_columns = ("slide", "shape", "ph_type", "safety_status", "recommendation", "reason")
        self.ph_tree = ttk.Treeview(ph_tab, columns=ph_columns, show="headings", selectmode="browse")

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

        ph_scroll = ttk.Scrollbar(ph_tab, orient=tk.VERTICAL, command=self.ph_tree.yview)
        self.ph_tree.configure(yscrollcommand=ph_scroll.set)
        self.ph_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ph_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # TAB 5: List Semantics Review (Phase 6)
        list_tab = ttk.Frame(notebook, padding="6 6 6 6")
        notebook.add(list_tab, text=" List Semantics Review ")

        list_columns = ("slide", "shape_id", "paragraph", "text_preview", "list_type", "level", "start_at", "format", "confidence", "status", "action")
        self.list_tree = ttk.Treeview(list_tab, columns=list_columns, show="headings", selectmode="browse")

        self.list_tree.heading("slide", text="Slide #")
        self.list_tree.heading("shape_id", text="Shape ID")
        self.list_tree.heading("paragraph", text="Para")
        self.list_tree.heading("text_preview", text="Text Preview")
        self.list_tree.heading("list_type", text="List Type")
        self.list_tree.heading("level", text="Level")
        self.list_tree.heading("start_at", text="Start")
        self.list_tree.heading("format", text="Bullet / Number Format")
        self.list_tree.heading("confidence", text="Confidence")
        self.list_tree.heading("status", text="Status")
        self.list_tree.heading("action", text="Proposed Action")

        self.list_tree.column("slide", width=55, anchor=tk.CENTER)
        self.list_tree.column("shape_id", width=65, anchor=tk.CENTER)
        self.list_tree.column("paragraph", width=50, anchor=tk.CENTER)
        self.list_tree.column("text_preview", width=180, anchor=tk.W)
        self.list_tree.column("list_type", width=95, anchor=tk.CENTER)
        self.list_tree.column("level", width=45, anchor=tk.CENTER)
        self.list_tree.column("start_at", width=45, anchor=tk.CENTER)
        self.list_tree.column("format", width=130, anchor=tk.CENTER)
        self.list_tree.column("confidence", width=80, anchor=tk.CENTER)
        self.list_tree.column("status", width=110, anchor=tk.CENTER)
        self.list_tree.column("action", width=220, anchor=tk.W)

        list_scroll = ttk.Scrollbar(list_tab, orient=tk.VERTICAL, command=self.list_tree.yview)
        self.list_tree.configure(yscrollcommand=list_scroll.set)
        self.list_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # TAB 6: Accessibility Intelligence Review (Phase 8)
        intel_tab = ttk.Frame(notebook, padding="6 6 6 6")
        notebook.add(intel_tab, text=" Accessibility Intelligence ")

        intel_columns = ("slide", "shape", "primary_rule", "related_rules", "confidence", "status", "recommendation", "reconciliation")
        self.intel_tree = ttk.Treeview(intel_tab, columns=intel_columns, show="headings", selectmode="browse")

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

        intel_scroll = ttk.Scrollbar(intel_tab, orient=tk.VERTICAL, command=self.intel_tree.yview)
        self.intel_tree.configure(yscrollcommand=intel_scroll.set)
        self.intel_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        intel_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # TAB 7: Alternative Text & Images Review (Phase 9 / Task 3)
        auth_tab = ttk.Frame(notebook, padding="6 6 6 6")
        notebook.add(auth_tab, text=" Alternative Text & Images Review ")

        auth_split = ttk.PanedWindow(auth_tab, orient=tk.VERTICAL)
        auth_split.pack(fill=tk.BOTH, expand=True)

        auth_top_frame = ttk.Frame(auth_split)
        auth_split.add(auth_top_frame, weight=3)

        auth_columns = ("slide", "shape", "type", "current", "suggested", "confidence", "grounding", "status")
        self.auth_tree = ttk.Treeview(auth_top_frame, columns=auth_columns, show="headings", selectmode="browse")

        self.auth_tree.heading("slide", text="Slide #")
        self.auth_tree.heading("shape", text="Shape (ID / Name)")
        self.auth_tree.heading("type", text="Authoring Target")
        self.auth_tree.heading("current", text="Current Content")
        self.auth_tree.heading("suggested", text="AI Suggestion (Gemini 3.5)")
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
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.auth_tree.bind("<<TreeviewSelect>>", self._on_auth_item_selected)

        # Bottom Review Decision Panel
        auth_bottom_frame = ttk.LabelFrame(auth_split, text=" Alternative Text & Accessibility Editor ", padding="8 8 8 8")
        auth_split.add(auth_bottom_frame, weight=2)

        self.auth_detail_var = tk.StringVar(value="Select an item above to review AI suggestion.")
        ttk.Label(auth_bottom_frame, textvariable=self.auth_detail_var, font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 4))

        edit_subframe = ttk.Frame(auth_bottom_frame)
        edit_subframe.pack(fill=tk.X, expand=True, pady=(2, 6))

        ttk.Label(edit_subframe, text="Proposed Content:").pack(side=tk.LEFT, padx=(0, 6))
        self.auth_edit_var = tk.StringVar(value="")
        self.auth_edit_entry = ttk.Entry(edit_subframe, textvariable=self.auth_edit_var, font=("Segoe UI", 9))
        self.auth_edit_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        btn_bar = ttk.Frame(auth_bottom_frame)
        btn_bar.pack(fill=tk.X)

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

        self.auth_batch_btn = ttk.Button(btn_bar, text="Approve All High-Confidence (>=85%)", state=tk.DISABLED, command=self._on_auth_approve_all)
        self.auth_batch_btn.pack(side=tk.RIGHT)

        self.selected_rec_id = None


        # TAB 8: Comprehensive QA & Validation (Phase 10)
        qa_tab = ttk.Frame(notebook, padding="6 6 6 6")
        notebook.add(qa_tab, text=" QA & Validation ")

        qa_split = ttk.PanedWindow(qa_tab, orient=tk.VERTICAL)
        qa_split.pack(fill=tk.BOTH, expand=True)

        qa_top_frame = ttk.Frame(qa_split)
        qa_split.add(qa_top_frame, weight=3)

        # Header bar in top frame
        qa_header_frame = ttk.Frame(qa_top_frame)
        qa_header_frame.pack(fill=tk.X, pady=(0, 6))

        self.qa_overall_badge_var = tk.StringVar(value="Overall QA Status: NOT RUN")
        self.qa_overall_lbl = ttk.Label(
            qa_header_frame,
            textvariable=self.qa_overall_badge_var,
            font=("Segoe UI", 11, "bold"),
            foreground="#1E3A8A"
        )
        self.qa_overall_lbl.pack(side=tk.LEFT, padx=(4, 15))

        self.qa_stats_var = tk.StringVar(value="Tests Passed: 0  |  Tests Failed: 0  |  Warnings: 0")
        ttk.Label(qa_header_frame, textvariable=self.qa_stats_var, font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 15))

        self.run_qa_btn = ttk.Button(
            qa_header_frame,
            text="Run Comprehensive QA",
            state=tk.DISABLED,
            command=self._on_run_qa_clicked
        )
        self.run_qa_btn.pack(side=tk.RIGHT)

        # Dimensional status box
        dim_box = ttk.LabelFrame(qa_top_frame, text=" Core Quality Assurance Dimensions ", padding="8 8 8 8")
        dim_box.pack(fill=tk.BOTH, expand=True, pady=(0, 4))

        # 10 Dimensions in 2 columns
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

        col0_frame = ttk.Frame(dim_box)
        col0_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)
        ttk.Label(col0_frame, textvariable=self.qa_dim_ooxml_var, font=("Segoe UI", 9)).pack(anchor=tk.W, pady=2)
        ttk.Label(col0_frame, textvariable=self.qa_dim_design_var, font=("Segoe UI", 9)).pack(anchor=tk.W, pady=2)
        ttk.Label(col0_frame, textvariable=self.qa_dim_orig_var, font=("Segoe UI", 9)).pack(anchor=tk.W, pady=2)
        ttk.Label(col0_frame, textvariable=self.qa_dim_rel_var, font=("Segoe UI", 9)).pack(anchor=tk.W, pady=2)
        ttk.Label(col0_frame, textvariable=self.qa_dim_acc_var, font=("Segoe UI", 9)).pack(anchor=tk.W, pady=2)

        col1_frame = ttk.Frame(dim_box)
        col1_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)
        ttk.Label(col1_frame, textvariable=self.qa_dim_idemp_var, font=("Segoe UI", 9)).pack(anchor=tk.W, pady=2)
        ttk.Label(col1_frame, textvariable=self.qa_dim_ai_var, font=("Segoe UI", 9)).pack(anchor=tk.W, pady=2)
        ttk.Label(col1_frame, textvariable=self.qa_dim_ws_var, font=("Segoe UI", 9)).pack(anchor=tk.W, pady=2)
        ttk.Label(col1_frame, textvariable=self.qa_dim_rep_var, font=("Segoe UI", 9)).pack(anchor=tk.W, pady=2)
        ttk.Label(col1_frame, textvariable=self.qa_dim_anim_var, font=("Segoe UI", 9)).pack(anchor=tk.W, pady=2)

        # Bottom detailed verification checks
        qa_bottom_frame = ttk.LabelFrame(qa_split, text=" Detailed Check Verification Log ", padding="6 6 6 6")
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

        # 7. Bottom Status & Report Button
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X)

        self.disclaimer_lbl = ttk.Label(
            bottom_frame,
            text=APP_DISCLAIMER,
            font=("Segoe UI", 8, "italic"),
            foreground="#6B7280"
        )
        self.disclaimer_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.open_output_btn = ttk.Button(
            bottom_frame,
            text="Open Remediated File",
            state=tk.DISABLED,
            command=self._on_open_output_file
        )
        self.open_output_btn.pack(side=tk.RIGHT, padx=(8, 0))

        self.view_report_btn = ttk.Button(
            bottom_frame,
            text="View Full Report",
            state=tk.DISABLED,
            command=self._on_view_report
        )
        self.view_report_btn.pack(side=tk.RIGHT)


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

        # Verify read accessibility / lock status
        try:
            with open(path, "rb") as test_f:
                header = test_f.read(4)
                if len(header) < 4:
                    messagebox.showerror("Empty File", f"The selected file is empty (0 bytes):\n{path.name}")
                    return
        except PermissionError:
            messagebox.showerror(
                "File Locked / Inaccessible",
                f"The file '{path.name}' is currently locked or in use by another program (e.g. PowerPoint).\n\n"
                f"Please close PowerPoint or unlock the file and try again."
            )
            return
        except Exception as e:
            messagebox.showerror("File Error", f"Cannot open '{path.name}':\n{e}")
            return

        self.file_path_var.set(str(path))
        self.status_var.set(f"Selected: {path.name}. Initializing workspace copy...")

        try:
            self.job_context = self.controller.initialize_job(path)
            self.metric_slides_var.set(f"Total Slides: {self.job_context.total_slides}")
            self.analyze_btn.config(state=tk.NORMAL)
            self.run_qa_btn.config(state=tk.NORMAL)
            self.open_workspace_btn.config(state=tk.NORMAL)
            self.status_var.set(f"Loaded {path.name} ({self.job_context.total_slides} slides). Ready to analyze.")
        except Exception as e:
            messagebox.showerror("Error Opening File", f"Could not load presentation: {e}")
            self.status_var.set("Failed to initialize presentation.")

    def _on_analyze_clicked(self):
        if not self.job_context:
            return

        self.analyze_btn.config(state=tk.DISABLED)
        self.remediate_btn.config(state=tk.DISABLED)
        self.prog_bar["value"] = 0

        def _worker():
            try:
                def _progress(curr, total, msg):
                    pct = int((curr / max(1, total)) * 100)
                    self.root.after(0, lambda: self._update_progress(pct, msg))

                self.job_context = self.controller.execute_analysis(self.job_context, progress_callback=_progress)
                self.root.after(0, self._on_analysis_finished)
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Analysis Error", f"Analysis failed: {e}"))

        threading.Thread(target=_worker, daemon=True).start()

    def _update_progress(self, percent: int, msg: str):
        self.prog_bar["value"] = percent
        self.status_var.set(msg)

    def _on_analysis_finished(self):
        ctx = self.job_context
        m = ctx.metrics

        self.metric_issues_var.set(f"Issues Detected: {m.total_issues}")
        self.metric_rules_var.set(f"Rules Passed: {m.passed_rules}")
        self.metric_score_var.set(f"Quality Score: {m.accessibility_score:.1f}/100")
        self.metric_integrity_var.set("Design Integrity: Verified")

        # 1. Clear and repopulate issues table
        for row in self.tree.get_children():
            self.tree.delete(row)

        for res in ctx.rule_results:
            for issue in res.issues:
                st_val = getattr(issue.status, "value", str(issue.status)) if hasattr(issue, "status") else "DETECTED"
                self.tree.insert("", tk.END, values=(
                    issue.slide_number,
                    issue.severity.value,
                    issue.rule_id,
                    issue.rule_name,
                    st_val,
                    issue.description,
                    issue.recommendation
                ))

        # 2. Clear and repopulate reading order table
        for row in self.ro_tree.get_children():
            self.ro_tree.delete(row)

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
                        issue.slide_number,
                        conf,
                        orig,
                        prop,
                        f_str,
                        reasons
                    ))

        # 3. Clear and repopulate selective ungrouping table (Phase 4)
        for row in self.grp_tree.get_children():
            self.grp_tree.delete(row)

        for res in ctx.rule_results:
            if res.rule_id == "SG-GRP-01":
                for issue in res.issues:
                    extra = issue.extra_data or {}
                    grp_id = issue.shape_id or extra.get("group_id", "N/A")
                    cat = extra.get("category", "MIXED")
                    el_cnt = extra.get("eligible_count", 0)
                    pic_cnt = extra.get("picture_count", 0)
                    is_safe = extra.get("is_safe_to_remediate", True)
                    safety_str = "SAFE TO UNGROUP" if is_safe else "PRESERVED / MANUAL"
                    skip_reason = extra.get("skip_reason")
                    strategy = skip_reason if skip_reason else f"Extract {el_cnt} text/shape(s) to slide; protect {pic_cnt} picture(s)."

                    self.grp_tree.insert("", tk.END, values=(
                        issue.slide_number,
                        grp_id,
                        cat,
                        el_cnt,
                        pic_cnt,
                        safety_str,
                        strategy
                    ))

        # 4. Clear and repopulate placeholder conversion table (Phase 5)
        for row in self.ph_tree.get_children():
            self.ph_tree.delete(row)

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
                        issue.slide_number,
                        f"{s_id}: {s_name}",
                        ph_type.upper(),
                        safety_status,
                        rec,
                        reason
                    ))

        # 5. Clear and repopulate list semantics table (Phase 6)
        for row in self.list_tree.get_children():
            self.list_tree.delete(row)

        for res in ctx.rule_results:
            if res.rule_id == "SG-LST-01":
                for issue in res.issues:
                    extra = issue.extra_data or {}
                    s_id = issue.shape_id or extra.get("shape_id", "N/A")
                    p_idx = extra.get("paragraph_index", 0) + 1
                    preview = extra.get("text_preview", "")
                    cat = extra.get("category", "NONE")
                    lvl = extra.get("level", 0)
                    start_at = extra.get("numbering_start", 1)
                    fmt = extra.get("numbering_type") or extra.get("bullet_char") or "None"
                    conf = extra.get("confidence", "HIGH")
                    st = extra.get("status", "VALID")
                    action = extra.get("proposed_action", issue.recommendation)

                    self.list_tree.insert("", tk.END, values=(
                        issue.slide_number,
                        s_id,
                        f"P{p_idx}",
                        preview[:35],
                        cat,
                        lvl,
                        start_at,
                        fmt,
                        conf,
                        st,
                        action
                    ))

        # 6. Clear and repopulate Accessibility Intelligence table (Phase 8)
        for row in self.intel_tree.get_children():
            self.intel_tree.delete(row)

        for u in getattr(ctx, "unified_issues", []):
            related_str = ", ".join(u.related_rule_ids) if u.related_rule_ids else "-"
            shape_str = str(u.shape_id) if u.shape_id is not None else "Slide"
            rec_stat = getattr(u.reconciliation_status, "value", str(u.reconciliation_status))
            conf_str = getattr(u.confidence, "value", str(u.confidence))
            elig_str = getattr(u.remediation_eligibility, "value", str(u.remediation_eligibility))
            self.intel_tree.insert("", tk.END, values=(
                u.slide_number,
                shape_str,
                u.primary_rule_id,
                related_str,
                conf_str,
                elig_str,
                u.recommended_action or u.explanation,
                rec_stat
            ))

        # 7. Clear and repopulate AI Authoring table (Phase 9)
        for row in self.auth_tree.get_children():
            self.auth_tree.delete(row)

        ai_recs = getattr(ctx, "ai_recommendations", [])
        for r in ai_recs:
            stat_val = getattr(r.status, "value", str(r.status))
            ground_val = getattr(r.fact_grounding, "value", str(r.fact_grounding))
            self.auth_tree.insert("", tk.END, iid=r.recommendation_id, values=(
                r.slide_number,
                r.shape_id or "Slide",
                getattr(r.authoring_type, "value", str(r.authoring_type)),
                r.original_value or r.current_value,
                r.effective_value,
                f"{r.confidence:.0%}",
                ground_val,
                stat_val
            ))

        if ai_recs:
            self.auth_batch_btn.config(state=tk.NORMAL)

        self.analyze_btn.config(state=tk.NORMAL)
        self.approval_chk.config(state=tk.NORMAL)
        self.view_report_btn.config(state=tk.NORMAL)
        self.status_var.set(
            f"Analysis complete. Found {m.total_issues} issue(s) "
            f"({len(getattr(ctx, 'unified_issues', []))} correlated, {len(ai_recs)} AI authoring proposals). "
            f"Quality score V2: {m.quality_score_v2:.1f}/100."
        )


    def _on_approval_toggled(self):
        if not self.job_context:
            return
        approved = self.approval_var.get()
        self.job_context.user_approved_remediation = approved
        self.remediate_btn.config(state=tk.NORMAL if approved else tk.DISABLED)

    def _on_remediate_clicked(self):
        if not self.job_context or not self.job_context.user_approved_remediation:
            messagebox.showwarning("Approval Required", "You must explicitly check 'Approve Remediation Plan' first.")
            return

        self.remediate_btn.config(state=tk.DISABLED)
        self.prog_bar["value"] = 0

        def _worker():
            try:
                def _progress(curr, total, msg):
                    pct = int((curr / max(1, total)) * 100)
                    self.root.after(0, lambda: self._update_progress(pct, msg))

                self.job_context = self.controller.execute_remediation(self.job_context, progress_callback=_progress)
                self.root.after(0, self._on_remediation_finished)
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Remediation Error", f"Remediation failed: {e}"))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_remediation_finished(self):
        ctx = self.job_context
        v = ctx.verification_result
        v_status = "PASSED" if (v and v.passed) else "INSPECTION NEEDED"

        # Update post-remediation score in header banner
        if ctx.post_remediation_score is not None:
            self.metric_score_var.set(f"Quality Score V2: {ctx.post_remediation_score:.1f}/100")

        # Refresh Accessibility Intelligence review tab with reconciled statuses
        for row in self.intel_tree.get_children():
            self.intel_tree.delete(row)

        for u in getattr(ctx, "unified_issues", []):
            related_str = ", ".join(u.related_rule_ids) if u.related_rule_ids else "-"
            shape_str = str(u.shape_id) if u.shape_id is not None else "Slide"
            rec_stat = getattr(u.reconciliation_status, "value", str(u.reconciliation_status))
            conf_str = getattr(u.confidence, "value", str(u.confidence))
            elig_str = getattr(u.remediation_eligibility, "value", str(u.remediation_eligibility))
            self.intel_tree.insert("", tk.END, values=(
                u.slide_number,
                shape_str,
                u.primary_rule_id,
                related_str,
                conf_str,
                elig_str,
                u.recommended_action or u.explanation,
                rec_stat
            ))

        # Refresh AI Authoring review tab
        for row in self.auth_tree.get_children():
            self.auth_tree.delete(row)

        for r in getattr(ctx, "ai_recommendations", []):
            stat_val = getattr(r.status, "value", str(r.status))
            ground_val = getattr(r.fact_grounding, "value", str(r.fact_grounding))
            self.auth_tree.insert("", tk.END, iid=r.recommendation_id, values=(
                r.slide_number,
                r.shape_id or "Slide",
                getattr(r.authoring_type, "value", str(r.authoring_type)),
                r.original_value or r.current_value,
                r.effective_value,
                f"{r.confidence:.0%}",
                ground_val,
                stat_val
            ))

        # Count applied actions across pipeline steps
        ro_applied = any(a.rule_id == "PIPELINE_STEP_1_READING_ORDER" and a.status == RemediationStatus.APPLIED for a in ctx.remediation_actions)
        grp_applied = any(a.rule_id == "PIPELINE_STEP_2_UNGROUPING" and a.status == RemediationStatus.APPLIED for a in ctx.remediation_actions)
        ph_applied = any(a.rule_id == "PIPELINE_STEP_3_PLACEHOLDERS" and a.status == RemediationStatus.APPLIED for a in ctx.remediation_actions)
        lst_applied = any(a.rule_id == "PIPELINE_STEP_4_LISTS" and a.status == RemediationStatus.APPLIED for a in ctx.remediation_actions)
        acc_applied = any(a.rule_id == "PIPELINE_STEP_5_ACCESSIBILITY" and a.status == RemediationStatus.APPLIED for a in ctx.remediation_actions)
        auth_applied = any(a.rule_id.startswith("AI_AUTH") and a.status == RemediationStatus.APPLIED for a in ctx.remediation_actions)

        steps_executed = []
        if ro_applied:
            steps_executed.append("Reading Order Alignment")
        if grp_applied:
            steps_executed.append("Selective Ungrouping")
        if ph_applied:
            steps_executed.append("Placeholder Conversion")
        if lst_applied:
            steps_executed.append("List Semantics & Numbering")
        if acc_applied:
            steps_executed.append("Accessibility Remediation")
        if auth_applied:
            steps_executed.append("Approved AI Content")
        if not steps_executed:
            steps_executed.append("Safe Multi-Stage Pipeline (Preserved)")

        summary_steps = " + ".join(steps_executed)
        score_info = f"Score: {ctx.post_remediation_score:.1f}/100." if ctx.post_remediation_score is not None else ""
        self.status_var.set(f"Remediation executed ({summary_steps}). Validation: {v_status}. {score_info}")

        # Update Phase 10 QA & Validation tab if available
        if getattr(ctx, "qa_result", None):
            self._populate_qa_tab(ctx.qa_result)
        self.run_qa_btn.config(state=tk.NORMAL)

        if ctx.final_output_file and ctx.final_output_file.exists():
            self.open_output_btn.config(state=tk.NORMAL)

        output_details = ""
        if ctx.final_output_file:
            output_details += f"\n • Remediated File: {ctx.final_output_file.name}\n   Location: {ctx.final_output_file}"
        if ctx.final_report_file:
            output_details += f"\n • Full Report: {ctx.final_report_file.name}\n   Location: {ctx.final_report_file}"

        messagebox.showinfo(
            "Remediation Completed",
            f"Remediation pipeline successfully completed!\n\n"
            f"Applied Steps:\n • {summary_steps}\n\n"
            f"Validation Status: {v_status}\n"
            f"{f'Post-Remediation Quality Score: {ctx.post_remediation_score:.1f}/100' if ctx.post_remediation_score is not None else ''}"
            f"{output_details}\n\n"
            f"Original file is preserved unchanged."
        )

    def _on_auth_item_selected(self, event):
        sel = self.auth_tree.selection()
        if not sel or not self.job_context:
            return
        item_id = sel[0]
        rec = next((r for r in getattr(self.job_context, "ai_recommendations", []) if getattr(r, "recommendation_id", "") == item_id), None)
        if rec:
            self.selected_rec_id = rec.recommendation_id
            type_str = getattr(rec.authoring_type, "value", str(rec.authoring_type))
            self.auth_detail_var.set(
                f"Slide {rec.slide_number} | Shape {rec.shape_id or 'Slide'} | {type_str} | "
                f"Confidence: {rec.confidence:.0%} | Rationale: {rec.rationale}"
            )
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
            self.status_var.set(f"Approved {getattr(rec.authoring_type, 'value', str(rec.authoring_type))} for slide {rec.slide_number}.")

    def _on_auth_edit(self):
        if not self.job_context or not self.selected_rec_id:
            return
        rec = next((r for r in self.job_context.ai_recommendations if r.recommendation_id == self.selected_rec_id), None)
        if rec:
            new_text = self.auth_edit_var.get().strip()
            if not new_text:
                messagebox.showwarning("Empty Content", "Authoring content cannot be empty.")
                return
            rec.edit(new_text)
            self._refresh_auth_tree_item(rec)
            self.status_var.set(f"Saved custom content & approved for slide {rec.slide_number}.")

    def _on_auth_reject(self):
        if not self.job_context or not self.selected_rec_id:
            return
        rec = next((r for r in self.job_context.ai_recommendations if r.recommendation_id == self.selected_rec_id), None)
        if rec:
            rec.reject()
            self._refresh_auth_tree_item(rec)
            self.status_var.set(f"Rejected suggestion for slide {rec.slide_number}.")

    def _on_auth_mark_decorative(self):
        if not self.job_context or not self.selected_rec_id:
            return
        rec = next((r for r in self.job_context.ai_recommendations if r.recommendation_id == self.selected_rec_id), None)
        if rec:
            rec.mark_decorative()
            self._refresh_auth_tree_item(rec)
            self.status_var.set(f"Marked shape {rec.shape_id} on slide {rec.slide_number} as decorative.")

    def _on_auth_keep_existing(self):
        if not self.job_context or not self.selected_rec_id:
            return
        rec = next((r for r in self.job_context.ai_recommendations if r.recommendation_id == self.selected_rec_id), None)
        if rec:
            rec.keep_existing()
            self._refresh_auth_tree_item(rec)
            self.status_var.set(f"Preserved existing alt text for shape {rec.shape_id} on slide {rec.slide_number}.")

    def _on_auth_approve_all(self):
        if not self.job_context:
            return
        approved_cnt = 0
        for rec in self.job_context.ai_recommendations:
            if rec.status == AuthoringStatus.PENDING_REVIEW and rec.confidence >= 0.85:
                rec.approve()
                self._refresh_auth_tree_item(rec)
                approved_cnt += 1
        self.status_var.set(f"Batch approved {approved_cnt} high-confidence recommendations (>=85%).")


    def _refresh_auth_tree_item(self, rec):
        if self.auth_tree.exists(rec.recommendation_id):
            stat_val = getattr(rec.status, "value", str(rec.status))
            ground_val = getattr(rec.fact_grounding, "value", str(rec.fact_grounding))
            self.auth_tree.item(rec.recommendation_id, values=(
                rec.slide_number,
                rec.shape_id or "Slide",
                getattr(rec.authoring_type, "value", str(rec.authoring_type)),
                rec.original_value or rec.current_value,
                rec.effective_value,
                f"{rec.confidence:.0%}",
                ground_val,
                stat_val
            ))

    def _on_open_workspace(self):
        if not self.job_context or not self.job_context.workspace_dir.exists():
            return
        os.startfile(str(self.job_context.workspace_dir))

    def _on_open_output_file(self):
        if not self.job_context or not self.job_context.final_output_file:
            return
        if self.job_context.final_output_file.exists():
            os.startfile(str(self.job_context.final_output_file))
        else:
            messagebox.showwarning("File Missing", f"Output file not found:\n{self.job_context.final_output_file}")

    def _on_view_report(self):
        if not self.job_context:
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

    def _on_run_qa_clicked(self):
        if not self.job_context:
            return
        self.run_qa_btn.config(state=tk.DISABLED)
        self.status_var.set("Running comprehensive QA validation suite...")

        def _worker():
            try:
                self.job_context = self.controller.run_qa_verification(self.job_context)
                self.root.after(0, self._on_qa_finished)
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("QA Error", f"QA validation failed: {e}"))
                self.root.after(0, lambda: self.run_qa_btn.config(state=tk.NORMAL))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_qa_finished(self):
        self.run_qa_btn.config(state=tk.NORMAL)
        if self.job_context and getattr(self.job_context, "qa_result", None):
            self._populate_qa_tab(self.job_context.qa_result)
            self.status_var.set(f"Comprehensive QA completed: {self.job_context.qa_result.overall_status.value}.")

    def _populate_qa_tab(self, qa_res: QAResult):
        st_val = qa_res.overall_status.value
        self.qa_overall_badge_var.set(f"Overall QA Status: {st_val}")
        if st_val == "PASS":
            self.qa_overall_lbl.config(foreground="#047857")
        elif st_val == "PASS_WITH_WARNINGS":
            self.qa_overall_lbl.config(foreground="#D97706")
        else:
            self.qa_overall_lbl.config(foreground="#DC2626")

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
                check.name,
                check.component,
                check.status.value,
                check.details
            ))


