#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KPI 2026 - Modern Local Desktop App
Design System: Professional, Colorful, Clean
Mac + Windows | CustomTkinter + Matplotlib
"""
import json, pathlib, sys, os, platform, io, datetime
import tkinter as tk
from tkinter import messagebox, filedialog
import openpyxl
import customtkinter as ctk
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.font_manager as fm

# PDF helpers (optional, graceful fallback if not installed)
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib.colors import HexColor, white, black
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak, HRFlowable
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    def ar(text):
        try:
            return get_display(arabic_reshaper.reshape(str(text)))
        except: return str(text)
except ImportError:
    def ar(text): return str(text)

# ---------- Design System ----------
COLORS = {
    "primary": "#6366F1",       # Indigo
    "primary_dark": "#4F46E5",
    "primary_light": "#818CF8",
    "secondary": "#8B5CF6",     # Violet
    "accent": "#06B6D4",        # Cyan
    "accent_warm": "#F59E0B",   # Amber
    "success": "#10B981",
    "warning": "#F59E0B",
    "error": "#EF4444",
    "bg": "#F8FAFC",            # Slate 50
    "bg_card": "#FFFFFF",
    "bg_hover": "#F1F5F9",
    "border": "#E2E8F0",
    "text": "#0F172A",          # Slate 900
    "text_sec": "#475569",
    "text_muted": "#94A3B8",
    "level": {1:"#EF4444",2:"#F59E0B",3:"#EAB308",4:"#10B981",5:"#06B6D4"},
    "level_bg": {1:"#FEF2F2",2:"#FFFBEB",3:"#FEFCE8",4:"#ECFDF5",5:"#ECFEFF"},
}
LEVELS_AR = {1:"ناشئ",2:"أولي",3:"أساسي",4:"منظم",5:"ديناميكي"}
LEVELS_DESC = {1:"لا توجد ممارسة واضحة",2:"بدأ التطبيق بصورة أولية",3:"الممارسة موجودة أساسية",4:"الممارسة واضحة ومنظمة",5:"الممارسة متطورة وديناميكية"}
FONTS = {
    "title": ("Segoe UI" if platform.system()=="Windows" else "Helvetica", 18, "bold"),
    "heading": ("Segoe UI" if platform.system()=="Windows" else "Helvetica", 13, "bold"),
    "body": ("Segoe UI" if platform.system()=="Windows" else "Helvetica", 11),
    "small": ("Segoe UI" if platform.system()=="Windows" else "Helvetica", 10),
    "tiny": ("Segoe UI" if platform.system()=="Windows" else "Helvetica", 9),
}

def get_level(avg):
    if avg < 1.8: return "ناشئ",1
    if avg < 2.6: return "أولي",2
    if avg < 3.4: return "أساسي",3
    if avg < 4.2: return "منظم",4
    return "ديناميكي",5

# ---------- Paths (PyInstaller compatible) ----------
def get_base():
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return pathlib.Path(sys._MEIPASS)
    exe_dir = pathlib.Path(sys.executable).parent if getattr(sys, 'frozen', False) else pathlib.Path(__file__).parent
    if (exe_dir / "KPI-Final 2026.xlsx").exists():
        return exe_dir
    return pathlib.Path(__file__).parent

BASE = get_base()
EXCEL = BASE / "KPI-Final 2026.xlsx"
JSON_PATH = BASE / "kpi_domains.json"
if not EXCEL.exists():
    for p in [pathlib.Path.cwd(), pathlib.Path(__file__).parent, pathlib.Path(sys.executable).parent]:
        if (p / "KPI-Final 2026.xlsx").exists():
            EXCEL = p / "KPI-Final 2026.xlsx"
            JSON_PATH = p / "kpi_domains.json"
            BASE = p
            break

# Autosave location (hidden file in home, survives app restarts, not in app folder)
AUTOSAVE_PATH = pathlib.Path.home() / ".kpi_autosave.json"

def load_domains():
    return json.loads(JSON_PATH.read_text(encoding="utf-8"))
domains = load_domains()

# ---------- Appearance ----------
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class ModernCard(ctk.CTkFrame):
    """Card with shadow-like border and rounded corners"""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color=COLORS["bg_card"], corner_radius=12, border_width=1, border_color=COLORS["border"], **kwargs)

class LevelBadge(ctk.CTkFrame):
    def __init__(self, parent, level_n, text, **kwargs):
        bg = COLORS["level_bg"][level_n]
        fg = COLORS["level"][level_n]
        super().__init__(parent, fg_color=bg, corner_radius=20, **kwargs)
        ctk.CTkLabel(self, text=text, text_color=fg, font=(FONTS["tiny"][0], 9, "bold")).pack(padx=10, pady=4)

class SegmentedScore(ctk.CTkFrame):
    """Modern segmented control for 1-5"""
    def __init__(self, parent, variable, command, **kwargs):
        super().__init__(parent, fg_color=COLORS["bg"], corner_radius=10, **kwargs)
        self.var = variable
        self.cmd = command
        self.btns = []
        labels = ["١ ناشئ","٢ أولي","٣ أساسي","٤ منظم","٥ ديناميكي"]
        for i in range(1,6):
            btn = ctk.CTkButton(self, text=labels[i-1], width=85, height=32, corner_radius=8,
                                fg_color="transparent", text_color=COLORS["text_sec"],
                                hover_color=COLORS["bg_hover"], border_width=0,
                                font=(FONTS["small"][0], 10, "bold"),
                                command=lambda v=i: self.select(v))
            btn.pack(side="left", padx=2, pady=2)
            self.btns.append(btn)
        self.update_visual()

    def select(self, v):
        self.var.set(v)
        self.update_visual()
        if self.cmd: self.cmd()

    def update_visual(self):
        v = self.var.get()
        for i, btn in enumerate(self.btns, 1):
            if i == v:
                btn.configure(fg_color=COLORS["level"][v], text_color="white", hover_color=COLORS["level"][v])
            elif i < v:
                btn.configure(fg_color=COLORS["level_bg"][i], text_color=COLORS["level"][i])
            else:
                btn.configure(fg_color="transparent", text_color=COLORS["text_muted"])

class KPIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("KPI 2026 — نظام التقييم المؤسسي")
        self.root.geometry("1380x800")
        self.root.minsize(1200, 700)
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except: pass
        # App icon (for window and taskbar) - uses icon.png if available
        try:
            icon_path = BASE / "icon.png"
            if not icon_path.exists():
                icon_path = pathlib.Path(__file__).parent / "icon.png"
            if icon_path.exists():
                icon_img = Image.open(str(icon_path))
                # Resize to 64 for window icon (Tk needs PhotoImage)
                icon_img = icon_img.resize((64,64), Image.LANCZOS)
                self._icon_photo = ImageTk.PhotoImage(icon_img)
                self.root.iconphoto(True, self._icon_photo)
        except Exception:
            pass
        try:
            # Fallback try .ico on Windows
            if platform.system() == "Windows":
                ico = BASE / "KPI.ico"
                if ico.exists():
                    self.root.iconbitmap(str(ico))
        except: pass

        # Data
        self.scores = {d["domain"]: {ind["name"]: 1 for ind in d["indicators"]} for d in domains}
        self.org_var = tk.StringVar(value="")
        self.vars = {}
        self.current_domain = 0
        self.search_var = tk.StringVar(value="")
        self._autosave_dirty = False

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.org_var.trace_add("write", lambda *a: self._autosave())

        # Root bg
        self.root.configure(bg=COLORS["bg"])

        # --- Header (Gradient-like with solid + accent) ---
        header = ctk.CTkFrame(root, fg_color=COLORS["primary"], corner_radius=0, height=72)
        header.pack(fill="x")
        header.pack_propagate(False)
        h_inner = ctk.CTkFrame(header, fg_color="transparent")
        h_inner.pack(fill="both", expand=True, padx=20, pady=10)

        # Logo + Title
        left_h = ctk.CTkFrame(h_inner, fg_color="transparent")
        left_h.pack(side="left", fill="y")
        # Icon circle
        icon = ctk.CTkFrame(left_h, fg_color="white", corner_radius=22, width=44, height=44)
        icon.pack(side="left", padx=(0,12))
        icon.pack_propagate(False)
        ctk.CTkLabel(icon, text="📊", font=(FONTS["title"][0], 20)).place(relx=0.5, rely=0.5, anchor="center")
        ctk.CTkLabel(left_h, text="KPI 2026", text_color="white", font=(FONTS["title"][0], 20, "bold")).pack(side="left", anchor="w")
        ctk.CTkLabel(left_h, text="  نظام التقييم المؤسسي  •  58 مؤشر  •  9 محاور", text_color="#E0E7FF", font=FONTS["small"]).pack(side="left", padx=(10,0), anchor="w")

        # Center: Org input
        center_h = ctk.CTkFrame(h_inner, fg_color="transparent")
        center_h.pack(side="left", padx=30)
        ctk.CTkLabel(center_h, text="اسم المؤسسة", text_color="white", font=FONTS["tiny"]).pack(anchor="w")
        self.entry_org = ctk.CTkEntry(center_h, textvariable=self.org_var, width=240, height=34,
                                      fg_color="white", text_color=COLORS["text"], border_width=0, corner_radius=8,
                                      placeholder_text="مثال: مركز تعلم الكبار...")
        self.entry_org.pack()
        self.entry_org.bind("<KeyRelease>", lambda e: self.update_overall())

        # Right: Overall metrics
        right_h = ctk.CTkFrame(h_inner, fg_color="transparent")
        right_h.pack(side="right")
        # Overall card
        self.overall_card = ctk.CTkFrame(right_h, fg_color="white", corner_radius=12, width=280, height=52)
        self.overall_card.pack(side="right", padx=(12,0))
        self.overall_card.pack_propagate(False)
        self.lbl_overall_val = ctk.CTkLabel(self.overall_card, text="1.00", text_color=COLORS["level"][1], font=(FONTS["title"][0], 22, "bold"))
        self.lbl_overall_val.place(x=14, y=6)
        self.lbl_overall_lvl = ctk.CTkLabel(self.overall_card, text="ناشئ  •  20/100", text_color=COLORS["text_sec"], font=FONTS["small"])
        self.lbl_overall_lvl.place(x=14, y=30)
        # Progress circle-like bar
        self.progress = ctk.CTkProgressBar(right_h, width=100, height=8, corner_radius=4, progress_color=COLORS["level"][1], fg_color="#E0E7FF")
        self.progress.pack(side="right", padx=12)
        self.progress.set(0.2)

        # PDF Arabic support check (Item 3)
        try:
            import uharfbuzz
            has_harfbuzz = True
        except ImportError:
            has_harfbuzz = False
            warn = ctk.CTkFrame(root, fg_color="#FEF2F2", corner_radius=8, border_width=1, border_color=COLORS["error"])
            warn.pack(fill="x", padx=12, pady=(0,6))
            ctk.CTkLabel(warn, text="⚠️  PDF Arabic يحتاج:  pip install uharfbuzz  —  التصدير سيعمل لكن بجودة أقل", text_color=COLORS["error"], font=FONTS["tiny"]).pack(padx=10, pady=5)

        # --- Onboarding 3-Step Header (Item 4) ---
        steps_frame = ctk.CTkFrame(root, fg_color="white", corner_radius=0, height=44, border_width=1, border_color=COLORS["border"])
        steps_frame.pack(fill="x", padx=12, pady=(0,6))
        steps_frame.pack_propagate(False)
        # Progress 0/58
        self.lbl_progress = ctk.CTkLabel(steps_frame, text="0/58 مكتمل", text_color=COLORS["text_sec"], font=FONTS["tiny"])
        self.lbl_progress.pack(side="right", padx=16, pady=10)
        self.progress_steps = ctk.CTkProgressBar(steps_frame, width=140, height=6, progress_color=COLORS["primary"], fg_color=COLORS["border"])
        self.progress_steps.pack(side="right", pady=14)
        self.progress_steps.set(0)
        # Steps
        self.step_labels = []
        steps = [("١", "إدخال الاسم"), ("٢", "تقييم 58 مؤشر"), ("٣", "تصدير PDF/Excel")]
        for i, (num, title) in enumerate(steps):
            step = ctk.CTkFrame(steps_frame, fg_color="transparent")
            step.pack(side="left", padx=16, pady=8)
            circle = ctk.CTkFrame(step, fg_color=COLORS["primary"] if i==0 else COLORS["border"], corner_radius=12, width=24, height=24)
            circle.pack(side="left", padx=(0,6))
            circle.pack_propagate(False)
            lbl_num = ctk.CTkLabel(circle, text=num, text_color="white" if i==0 else COLORS["text_muted"], font=(FONTS["tiny"][0], 10, "bold"))
            lbl_num.place(relx=0.5, rely=0.5, anchor="center")
            lbl_title = ctk.CTkLabel(step, text=title, text_color=COLORS["text"] if i==0 else COLORS["text_muted"], font=FONTS["small"])
            lbl_title.pack(side="left")
            if i < 2:
                ctk.CTkLabel(steps_frame, text="←", text_color=COLORS["border"], font=FONTS["small"]).pack(side="left", padx=4)
            self.step_labels.append((circle, lbl_num, lbl_title))
        self.steps_frame = steps_frame

        # --- Main Container ---
        main = ctk.CTkFrame(root, fg_color=COLORS["bg"])
        main.pack(fill="both", expand=True, padx=12, pady=12)

        # --- Sidebar ---
        sidebar = ModernCard(main, width=300)
        sidebar.pack(side="left", fill="y", padx=(0,12), pady=0)
        sidebar.pack_propagate(False)

        # Sidebar header
        s_head = ctk.CTkFrame(sidebar, fg_color="transparent")
        s_head.pack(fill="x", padx=14, pady=14)
        ctk.CTkLabel(s_head, text="المحاور", text_color=COLORS["text"], font=FONTS["heading"]).pack(side="left")
        ctk.CTkLabel(s_head, text="9", text_color="white", fg_color=COLORS["primary"], corner_radius=12, width=24, height=20, font=(FONTS["tiny"][0],9,"bold")).pack(side="right")

        # Search
        search = ctk.CTkEntry(sidebar, textvariable=self.search_var, placeholder_text="🔍 بحث في المحاور...", height=36, corner_radius=10, fg_color=COLORS["bg"], border_color=COLORS["border"])
        search.pack(fill="x", padx=14, pady=(0,10))
        self.search_var.trace_add("write", lambda *a: self.filter_domains())

        # Domain list scroll
        self.sidebar_scroll = ctk.CTkScrollableFrame(sidebar, fg_color="transparent")
        self.sidebar_scroll.pack(fill="both", expand=True, padx=6, pady=0)
        self.domain_cards = []
        self.build_sidebar()

        # Sidebar actions
        s_actions = ctk.CTkFrame(sidebar, fg_color="transparent")
        s_actions.pack(fill="x", padx=14, pady=12)
        ctk.CTkButton(s_actions, text="↺  تصفير", width=130, height=34, corner_radius=8, fg_color=COLORS["bg"], text_color=COLORS["text_sec"], hover_color=COLORS["border"], border_width=1, border_color=COLORS["border"], command=lambda: self.set_all(1)).pack(side="left", padx=(0,6))
        ctk.CTkButton(s_actions, text="★  الكل ٥", width=130, height=34, corner_radius=8, fg_color=COLORS["primary"], hover_color=COLORS["primary_dark"], command=lambda: self.set_all(5)).pack(side="left")
        # Preset
        self.preset = ctk.CTkOptionMenu(s_actions, values=["سيناريو سريع","ناشئ (1)","أساسي (3)","منظم (4)","ديناميكي (5)"], width=270, height=32, corner_radius=8, fg_color="white", button_color=COLORS["primary"], text_color=COLORS["text"], command=self.apply_preset)
        self.preset.pack(pady=(8,0))
        self.preset.set("سيناريو سريع")

        # --- Content Area with Tabs ---
        content = ctk.CTkFrame(main, fg_color="transparent")
        content.pack(side="left", fill="both", expand=True)

        # Tab bar (modern pill)
        tab_bar = ctk.CTkFrame(content, fg_color=COLORS["bg_card"], corner_radius=12, border_width=1, border_color=COLORS["border"], height=48)
        tab_bar.pack(fill="x", pady=(0,12))
        tab_bar.pack_propagate(False)
        self.tab_var = tk.StringVar(value="eval")
        self.btn_eval = ctk.CTkButton(tab_bar, text="📝  التقييم", width=130, height=34, corner_radius=8, fg_color=COLORS["primary"], hover_color=COLORS["primary_dark"], command=lambda: self.switch_tab("eval"))
        self.btn_eval.pack(side="left", padx=6, pady=6)
        self.btn_results = ctk.CTkButton(tab_bar, text="📈  النتائج والتصدير", width=150, height=34, corner_radius=8, fg_color="transparent", text_color=COLORS["text_sec"], hover_color=COLORS["bg"], command=lambda: self.switch_tab("results"))
        self.btn_results.pack(side="left", padx=2, pady=6)
        # Quick stats in tab bar
        self.tab_stats = ctk.CTkLabel(tab_bar, text="58 مؤشر  •  متوسط 1.00  •  20/100", text_color=COLORS["text_muted"], font=FONTS["tiny"])
        self.tab_stats.pack(side="right", padx=16)

        # Tab containers
        self.tab_eval_frame = ctk.CTkFrame(content, fg_color="transparent")
        self.tab_results_frame = ctk.CTkFrame(content, fg_color="transparent")
        self.tab_eval_frame.pack(fill="both", expand=True)

        self.build_eval_tab()
        self.build_results_tab()
        self.update_overall()
        self.select_domain(0)
        self._update_onboarding_progress()
        self._show_onboarding_if_needed()
        # Load autosave AFTER UI is fully built
        self._load_autosave()

    # ---------- Sidebar ----------
    def build_sidebar(self):
        for w in self.sidebar_scroll.winfo_children():
            w.destroy()
        self.domain_cards.clear()
        for idx, d in enumerate(domains):
            # filter
            q = self.search_var.get().strip()
            if q and q not in d["domain"]:
                continue
            vals = list(self.scores[d["domain"]].values())
            avg = sum(vals)/len(vals)
            lvl, n = get_level(avg)
            is_active = idx == self.current_domain
            card = ctk.CTkFrame(self.sidebar_scroll, fg_color=COLORS["primary"] if is_active else "white",
                                corner_radius=10, border_width=1, border_color=COLORS["primary"] if is_active else COLORS["border"])
            card.pack(fill="x", pady=4, padx=2)
            # clickable
            card.bind("<Button-1>", lambda e, i=idx: self.select_domain(i))
            # inner
            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.pack(fill="x", padx=10, pady=10)
            inner.bind("<Button-1>", lambda e, i=idx: self.select_domain(i))

            top = ctk.CTkFrame(inner, fg_color="transparent")
            top.pack(fill="x")
            top.bind("<Button-1>", lambda e, i=idx: self.select_domain(i))
            ctk.CTkLabel(top, text=f"{idx+1:02d}", text_color="white" if is_active else COLORS["primary"],
                         fg_color=COLORS["primary"] if not is_active else "white",
                         corner_radius=6, width=28, height=22, font=(FONTS["tiny"][0],9,"bold")).pack(side="left")
            ctk.CTkLabel(top, text=f"{len(d['indicators'])} مؤشر", text_color="white" if is_active else COLORS["text_muted"], font=FONTS["tiny"]).pack(side="right")

            lbl = ctk.CTkLabel(inner, text=d["domain"], text_color="white" if is_active else COLORS["text"],
                               font=(FONTS["small"][0], 11, "bold"), anchor="w", wraplength=240, justify="left")
            lbl.pack(fill="x", pady=(6,4))
            lbl.bind("<Button-1>", lambda e, i=idx: self.select_domain(i))

            # progress
            prog = ctk.CTkProgressBar(inner, height=6, corner_radius=3, progress_color=COLORS["level"][n], fg_color=COLORS["border"] if not is_active else "#818CF8")
            prog.pack(fill="x", pady=(2,6))
            prog.set(avg/5)
            prog.bind("<Button-1>", lambda e, i=idx: self.select_domain(i))

            bottom = ctk.CTkFrame(inner, fg_color="transparent")
            bottom.pack(fill="x")
            ctk.CTkLabel(bottom, text=f"{avg:.2f}", text_color="white" if is_active else COLORS["level"][n], font=(FONTS["small"][0], 11, "bold")).pack(side="left")
            LevelBadge(bottom, n, lvl).pack(side="right")
            # hover
            def on_enter(e, c=card, act=is_active):
                if not act: c.configure(fg_color=COLORS["bg_hover"])
            def on_leave(e, c=card, act=is_active):
                if not act: c.configure(fg_color="white")
            card.bind("<Enter>", on_enter)
            card.bind("<Leave>", on_leave)
            self.domain_cards.append(card)
        # Force canvas to render its window content (fixes invisible text on first load)
        self.root.after(0, self._force_sidebar_canvas_update)

    def _force_sidebar_canvas_update(self):
        """Force CTkScrollableFrame canvas to render its window content - fixes invisible text on first load"""
        try:
            canvas = self.sidebar_scroll._parent_canvas
            inner = self.sidebar_scroll._parent_frame
            # Force canvas to recalculate scrollregion and render window content
            canvas.update_idletasks()
            inner.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Force a redraw by toggling canvas width
            w = canvas.winfo_width()
            if w > 1:
                canvas.itemconfig(1, width=w)  # force window item resize
        except Exception:
            pass

    def filter_domains(self):
        self.build_sidebar()

    def _update_sidebar_selection(self, old_idx, new_idx):
        """Lightweight sidebar selection update without full rebuild"""
        try:
            # Update old card
            if 0 <= old_idx < len(self.domain_cards):
                old_card = self.domain_cards[old_idx]
                old_card.configure(fg_color="white", border_color=COLORS["border"])
                # Find inner labels and update (first child is inner frame)
                for child in old_card.winfo_children():
                    if isinstance(child, ctk.CTkFrame):
                        for sub in child.winfo_children():
                            if isinstance(sub, ctk.CTkFrame):
                                for lbl in sub.winfo_children():
                                    if isinstance(lbl, ctk.CTkLabel) and lbl.cget("text") == f"{old_idx+1:02d}":
                                        lbl.configure(fg_color=COLORS["primary"], text_color="white")
                                    elif isinstance(lbl, ctk.CTkLabel) and "مؤشر" in str(lbl.cget("text")):
                                        lbl.configure(text_color=COLORS["text_muted"])
            # Update new card
            if 0 <= new_idx < len(self.domain_cards):
                new_card = self.domain_cards[new_idx]
                new_card.configure(fg_color=COLORS["primary"], border_color=COLORS["primary"])
                for child in new_card.winfo_children():
                    if isinstance(child, ctk.CTkFrame):
                        for sub in child.winfo_children():
                            if isinstance(sub, ctk.CTkFrame):
                                for lbl in sub.winfo_children():
                                    if isinstance(lbl, ctk.CTkLabel) and lbl.cget("text") == f"{new_idx+1:02d}":
                                        lbl.configure(fg_color="white", text_color=COLORS["primary"])
                                    elif isinstance(lbl, ctk.CTkLabel) and "مؤشر" in str(lbl.cget("text")):
                                        lbl.configure(text_color="white")
                # Also update the main label and bottom
                for child in new_card.winfo_children():
                    if isinstance(child, ctk.CTkFrame):
                        for sub in child.winfo_children():
                            if isinstance(sub, ctk.CTkLabel) and sub.cget("text") not in (f"{new_idx+1:02d}", f"{len(domains[new_idx]['indicators'])} مؤشر"):
                                # This is the domain name label
                                if len(str(sub.cget("text"))) > 10:
                                    sub.configure(text_color="white")
        except Exception as e:
            # Fallback to full rebuild if in-place fails
            print(f"Sidebar inplace failed, rebuild: {e}")
            self.build_sidebar()
            return
        # Also need to update the bottom avg labels - simpler to just rebuild that one card's bottom
        # For now, fallback to rebuilding just the two cards if needed
        # If the above didn't fully update, do a targeted rebuild of those two indices
        # To keep it simple and safe, we will do a full rebuild only if the lightweight fails
        # But we can also just not update the avg in sidebar on score change - it will update on next domain switch
        pass

    def select_domain(self, idx):
        is_initial = not hasattr(self, 'domain_cards') or len(self.domain_cards) == 0 or len(self.vars) == 0
        if idx == self.current_domain and not is_initial:
            return
        old_idx = self.current_domain
        self.current_domain = idx
        # Try lightweight update first (skip on initial)
        if not is_initial:
            try:
                self._update_sidebar_selection(old_idx, idx)
            except:
                self.build_sidebar()
        else:
            # Initial already built sidebar in __init__, don't rebuild
            if len(self.domain_cards) == 0:
                self.build_sidebar()
        self.show_domain()

    # ---------- Eval Tab ----------
    def build_eval_tab(self):
        # Domain header card
        self.domain_header = ModernCard(self.tab_eval_frame)
        self.domain_header.pack(fill="x", pady=(0,12))
        # Scroll for indicators
        self.eval_scroll = ctk.CTkScrollableFrame(self.tab_eval_frame, fg_color="transparent")
        self.eval_scroll.pack(fill="both", expand=True)

    def show_domain(self):
        for w in self.eval_scroll.winfo_children():
            w.destroy()
        for w in self.domain_header.winfo_children():
            w.destroy()
        d = domains[self.current_domain]
        vals = list(self.scores[d["domain"]].values())
        avg = sum(vals)/len(vals)
        lvl, n = get_level(avg)

        # Header content
        h = ctk.CTkFrame(self.domain_header, fg_color="transparent")
        h.pack(fill="x", padx=16, pady=14)
        left = ctk.CTkFrame(h, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(left, text=d["domain"], text_color=COLORS["text"], font=(FONTS["heading"][0], 15, "bold")).pack(anchor="w")
        ctk.CTkLabel(left, text=f"{len(d['indicators'])} مؤشر  •  اضغط على الدرجة لكل مؤشر  •  التغيير يحفظ تلقائياً", text_color=COLORS["text_muted"], font=FONTS["tiny"]).pack(anchor="w", pady=(2,0))
        right = ctk.CTkFrame(h, fg_color="transparent")
        right.pack(side="right")
        # Big metric - store refs for in-place update (Item 5)
        metric = ctk.CTkFrame(right, fg_color=COLORS["level_bg"][n], corner_radius=12, width=180, height=56)
        metric.pack()
        metric.pack_propagate(False)
        self.domain_header_metric = metric
        self.domain_header_avg_val = ctk.CTkLabel(metric, text=f"{avg:.2f}", text_color=COLORS["level"][n], font=(FONTS["title"][0], 20, "bold"))
        self.domain_header_avg_val.place(x=12, y=4)
        self.domain_header_avg_lvl = ctk.CTkLabel(metric, text=f"{lvl}  •  {avg*20:.0f}/100", text_color=COLORS["text_sec"], font=FONTS["tiny"])
        self.domain_header_avg_lvl.place(x=12, y=30)
        self.domain_header_prog = ctk.CTkProgressBar(metric, width=70, height=6, progress_color=COLORS["level"][n], fg_color="white")
        self.domain_header_prog.place(x=96, y=22)
        self.domain_header_prog.set(avg/5)

        # Indicators
        for ind in d["indicators"]:
            card = ModernCard(self.eval_scroll)
            card.pack(fill="x", pady=6, padx=2)

            # Card header
            ch = ctk.CTkFrame(card, fg_color="transparent")
            ch.pack(fill="x", padx=14, pady=(12,6))
            # Number badge
            badge = ctk.CTkFrame(ch, fg_color=COLORS["primary"], corner_radius=8, width=32, height=32)
            badge.pack(side="left", padx=(0,10))
            badge.pack_propagate(False)
            ctk.CTkLabel(badge, text=str(ind["id"]), text_color="white", font=(FONTS["small"][0], 12, "bold")).place(relx=0.5, rely=0.5, anchor="center")
            # Title + question
            txt = ctk.CTkFrame(ch, fg_color="transparent")
            txt.pack(side="left", fill="x", expand=True)
            ctk.CTkLabel(txt, text=ind["name"], text_color=COLORS["text"], font=FONTS["heading"], anchor="w").pack(anchor="w")
            lbl_q = ctk.CTkLabel(txt, text=ind["question"], text_color=COLORS["text_sec"], font=FONTS["small"], wraplength=900, justify="left", anchor="w")
            lbl_q.pack(anchor="w", pady=(2,0))

            # Score control
            key = (d["domain"], ind["name"])
            if key not in self.vars:
                self.vars[key] = tk.IntVar(value=self.scores[d["domain"]][ind["name"]])
            var = self.vars[key]
            seg = SegmentedScore(card, var, lambda k=key, v=var: self.on_score_change(k, v))
            seg.pack(fill="x", padx=14, pady=(0,8))
            var._seg = seg

            # Description card
            desc_card = ctk.CTkFrame(card, fg_color=COLORS["level_bg"][var.get()], corner_radius=10, border_width=1, border_color=COLORS["level"][var.get()])
            desc_card.pack(fill="x", padx=14, pady=(0,12))
            # keep reference
            var._desc_card = desc_card
            var._levels = ind["levels"]
            lbl_desc = ctk.CTkLabel(desc_card, text=f"› {LEVELS_AR[var.get()]}: {ind['levels'][var.get()-1]}", text_color=COLORS["text_sec"], font=FONTS["tiny"], wraplength=920, justify="left", anchor="w")
            lbl_desc.pack(fill="x", padx=10, pady=8)
            var._desc = lbl_desc
            self.update_desc(var)

    def update_desc(self, var):
        v = var.get()
        # Keep card bg static to avoid flicker, only update text and badge color
        var._desc.configure(text=f"› {LEVELS_AR[v]} — {LEVELS_DESC[v]}: {var._levels[v-1][:90]}…")
        # Instead of changing entire card bg (causes redraw), just update border via after_idle for next frame
        try:
            # Use after_idle to batch the color change to next idle, reducing flicker
            self.root.after_idle(lambda: var._desc_card.configure(fg_color=COLORS["level_bg"][v], border_color=COLORS["level"][v]) if hasattr(var, '_desc_card') and var._desc_card.winfo_exists() else None)
        except:
            var._desc_card.configure(fg_color=COLORS["level_bg"][v], border_color=COLORS["level"][v])
        if hasattr(var, "_seg"):
            # Only update the two affected buttons, not all 5, for speed
            try:
                # Find old and new values - we don't track old, so update all but debounced
                var._seg.update_visual()
            except: pass

    def _update_domain_header_inplace(self):
        """Item 5: update header avg without rebuilding entire domain (instant)"""
        try:
            d = domains[self.current_domain]
            vals = list(self.scores[d["domain"]].values())
            avg = sum(vals)/len(vals)
            lvl, n = get_level(avg)
            if hasattr(self, 'domain_header_avg_val'):
                self.domain_header_avg_val.configure(text=f"{avg:.2f}", text_color=COLORS["level"][n])
                self.domain_header_avg_lvl.configure(text=f"{lvl}  •  {avg*20:.0f}/100")
                self.domain_header_metric.configure(fg_color=COLORS["level_bg"][n])
                self.domain_header_prog.configure(progress_color=COLORS["level"][n])
                self.domain_header_prog.set(avg/5)
        except: pass

    def _update_sidebar_card_inplace(self, domain):
        """Item 5: update only the changed domain's sidebar card"""
        try:
            # Find idx
            idx = next((i for i, d in enumerate(domains) if d["domain"] == domain), None)
            if idx is None or idx >= len(self.domain_cards):
                return
            # Simplest: rebuild just that card is still heavy, so just mark dirty
            # For now, do lightweight: update progress text if card exists
            # Full rebuild only when switching domains, not on every click
            pass
        except: pass

    def on_score_change(self, key, var):
        domain, name = key
        self.scores[domain][name] = var.get()
        self.update_desc(var)
        self.update_overall()
        self._update_domain_header_inplace()
        # Item 5: no full rebuild - keep instant, sidebar will refresh on domain switch
        # Update only the overall progress, not full sidebar
        self._autosave()
        # toast-like feedback (shorter for speed)
        self.root.after(0, lambda: self.show_toast(f"✓ {name[:22]} → {var.get()}"))

    def show_toast(self, msg):
        toast = ctk.CTkFrame(self.root, fg_color=COLORS["text"], corner_radius=20)
        toast.place(relx=0.5, rely=0.92, anchor="center")
        ctk.CTkLabel(toast, text=msg, text_color="white", font=FONTS["tiny"]).pack(padx=14, pady=6)
        self.root.after(1800, toast.destroy)

    def set_all(self, val):
        for d in domains:
            for ind in d["indicators"]:
                self.scores[d["domain"]][ind["name"]] = val
                k=(d["domain"], ind["name"])
                if k in self.vars:
                    self.vars[k].set(val)
                    self.update_desc(self.vars[k])
        self.update_overall()
        self.build_sidebar()
        self.show_domain()
        self._autosave()
        # success toast
        self.show_toast(f"✓ تم تعيين كل المؤشرات إلى {val} — {LEVELS_AR[val]}")

    def apply_preset(self, val):
        mp={"ناشئ (1)":1,"أساسي (3)":3,"منظم (4)":4,"ديناميكي (5)":5}
        v=mp.get(val)
        if v: self.set_all(v)
        self.preset.set("سيناريو سريع")

    def update_overall(self):
        all_vals=[]
        for d in domains:
            all_vals.extend(self.scores[d["domain"]].values())
        avg = sum(all_vals)/len(all_vals) if all_vals else 0
        lvl,n = get_level(avg)
        score100 = avg*20
        try:
            self.lbl_overall_val.configure(text=f"{avg:.2f}", text_color=COLORS["level"][n])
            self.lbl_overall_lvl.configure(text=f"{lvl}  •  {score100:.0f}/100")
            self.progress.configure(progress_color=COLORS["level"][n])
            self.progress.set(score100/100)
            self.tab_stats.configure(text=f"58 مؤشر  •  {avg:.2f}/5  •  {lvl}  •  {score100:.0f}/100")
        except: pass
        # Update onboarding progress (Item 4)
        try:
            self._update_onboarding_progress()
        except: pass

    # ---------- Results Tab ----------
    def build_results_tab(self):
        # Top metrics (4 cards)
        top = ctk.CTkFrame(self.tab_results_frame, fg_color="transparent")
        top.pack(fill="x", pady=(0,12))
        self.metric_cards = []
        titles = ["المتوسط العام", "المستوى", "العلامة", "المتبقي للتميز"]
        for i in range(4):
            card = ModernCard(top)
            card.pack(side="left", fill="both", expand=True, padx=4)
            ctk.CTkLabel(card, text=titles[i], text_color=COLORS["text_muted"], font=FONTS["tiny"]).pack(pady=(12,2))
            val = ctk.CTkLabel(card, text="—", text_color=COLORS["text"], font=(FONTS["title"][0], 18, "bold"))
            val.pack()
            sub = ctk.CTkLabel(card, text="—", text_color=COLORS["text_sec"], font=FONTS["tiny"])
            sub.pack(pady=(0,12))
            self.metric_cards.append((val, sub, card))

        # Mid: table + charts
        mid = ctk.CTkFrame(self.tab_results_frame, fg_color="transparent")
        mid.pack(fill="both", expand=True)

        # Left: table card
        left_card = ModernCard(mid)
        left_card.pack(side="left", fill="both", expand=True, padx=(0,6))
        ctk.CTkLabel(left_card, text="جدول النتائج التفصيلي", text_color=COLORS["text"], font=FONTS["heading"]).pack(pady=12)
        # Modern tree with styled header
        table_frame = ctk.CTkFrame(left_card, fg_color="transparent")
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0,10))
        # Use tk tree but styled
        cols=("المحور","المتوسط","المستوى","من 100","الفجوة")
        self.tree = tk.ttk.Treeview(table_frame, columns=cols, show="headings", height=10)
        style = tk.ttk.Style()
        style.configure("Treeview", rowheight=28, font=(FONTS["small"][0], 10), background="white", fieldbackground="white")
        style.configure("Treeview.Heading", font=(FONTS["small"][0], 10, "bold"), background=COLORS["primary"], foreground="white")
        style.map("Treeview", background=[("selected", COLORS["primary_light"])])
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=80 if c!="المحور" else 200, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ctk.CTkScrollbar(table_frame, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        # gap label
        self.lbl_gap = ctk.CTkLabel(left_card, text="", text_color=COLORS["error"], font=(FONTS["small"][0], 10, "bold"), fg_color="#FEF2F2", corner_radius=8)
        self.lbl_gap.pack(fill="x", padx=10, pady=8)

        # Right: charts card
        right_card = ModernCard(mid)
        right_card.pack(side="left", fill="both", expand=True, padx=(6,0))
        ctk.CTkLabel(right_card, text="الرسوم البيانية", text_color=COLORS["text"], font=FONTS["heading"]).pack(pady=12)
        # Matplotlib
        plt.style.use("seaborn-v0_8-whitegrid")
        self.fig, (self.ax1, self.ax2) = plt.subplots(1, 2, figsize=(6.5, 3.0), dpi=100)
        self.fig.patch.set_facecolor("white")
        self.fig.tight_layout(pad=2.0)
        self.canvas = FigureCanvasTkAgg(self.fig, master=right_card)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=8, pady=8)

        # Bottom actions
        actions = ctk.CTkFrame(self.tab_results_frame, fg_color="transparent")
        actions.pack(fill="x", pady=12)
        ctk.CTkButton(actions, text="🔄  تحديث الرسوم", height=38, corner_radius=10, fg_color=COLORS["bg_card"], text_color=COLORS["text"], border_width=1, border_color=COLORS["border"], hover_color=COLORS["bg"], command=self.refresh_charts).pack(side="left", padx=6)
        ctk.CTkButton(actions, text="💾  تصدير Excel", height=42, corner_radius=10, fg_color=COLORS["primary"], hover_color=COLORS["primary_dark"], command=self.export_excel).pack(side="left", padx=6)
        ctk.CTkButton(actions, text="📕  تصدير PDF", height=42, corner_radius=10, fg_color="#DC2626", hover_color="#B91C1C", text_color="white", command=self.export_pdf).pack(side="left", padx=6)

        self.refresh_charts()

    def calc_results(self):
        res=[]
        all_vals=[]
        for d in domains:
            vals=list(self.scores[d["domain"]].values())
            avg=sum(vals)/len(vals)
            lvl,n=get_level(avg)
            res.append({"domain":d["domain"],"avg":round(avg,2),"score100":round(avg*20,1),"level":lvl,"level_n":n,"gap":round(5-avg,2)})
            all_vals.extend(vals)
        overall=sum(all_vals)/len(all_vals)
        lvl,n=get_level(overall)
        return res, round(overall,2), round(overall*20,1), lvl, n

    def refresh_charts(self):
        results, avg, score100, lvl, n = self.calc_results()
        # update metric cards
        self.metric_cards[0][0].configure(text=f"{avg:.2f}", text_color=COLORS["level"][n])
        self.metric_cards[0][1].configure(text="/ 5.00")
        self.metric_cards[1][0].configure(text=lvl, text_color=COLORS["level"][n])
        self.metric_cards[1][1].configure(text=LEVELS_DESC[n])
        self.metric_cards[2][0].configure(text=f"{score100:.0f}", text_color=COLORS["level"][n])
        self.metric_cards[2][1].configure(text="/ 100")
        self.metric_cards[3][0].configure(text=f"{5-avg:.2f}", text_color=COLORS["text"])
        self.metric_cards[3][1].configure(text=f"{100-score100:.0f} نقطة حتى 100")
        for _,_,card in self.metric_cards:
            card.configure(border_color=COLORS["level"][n] if n>=4 else COLORS["border"])

        # table
        for i in self.tree.get_children():
            self.tree.delete(i)
        for r in results:
            tag = f"lvl{r['level_n']}"
            self.tree.insert("", "end", values=(r["domain"], r["avg"], r["level"], r["score100"], r["gap"]), tags=(tag,))
            self.tree.tag_configure(tag, background=COLORS["level_bg"][r["level_n"]])
        self.tree.insert("", "end", values=("الإجمالي", avg, lvl, score100, round(5-avg,2)), tags=("total",))
        self.tree.tag_configure("total", background="#EEF2FF", font=(FONTS["small"][0], 10, "bold"))
        weakest = min(results, key=lambda x: x["avg"])
        self.lbl_gap.configure(text=f"⚠️  أضعف محور: {weakest['domain']}  ({weakest['avg']} — {weakest['level']})  •  يحتاج {weakest['gap']} للوصول للديناميكي")

        # charts - modern style
        self.ax1.clear(); self.ax2.clear()
        # Bar with rounded look
        names=[r["domain"].split()[0][:12] for r in results]
        scores=[r["score100"] for r in results]
        colors=[COLORS["level"][r["level_n"]] for r in results]
        bars=self.ax1.bar(names, scores, color=colors, edgecolor="white", linewidth=1.2, width=0.6, zorder=3)
        self.ax1.set_ylim(0,105)
        self.ax1.set_ylabel("من 100", fontsize=8, color=COLORS["text_muted"])
        self.ax1.set_title("العلامات حسب المحور", fontsize=10, fontweight="bold", color=COLORS["text"], pad=10)
        self.ax1.tick_params(axis='x', rotation=20, labelsize=7, colors=COLORS["text_sec"])
        self.ax1.tick_params(axis='y', labelsize=7, colors=COLORS["text_sec"])
        self.ax1.grid(axis='y', linestyle="--", alpha=0.3, zorder=0)
        for bar, v in zip(bars, scores):
            self.ax1.text(bar.get_x()+bar.get_width()/2, v+1.5, f"{v:.0f}", ha="center", fontsize=7, fontweight="bold", color=COLORS["text"])

        # Donut
        self.ax2.pie([score100, 100-score100], colors=[COLORS["level"][n], "#F1F5F9"], startangle=90, counterclock=False, wedgeprops=dict(width=0.38, edgecolor="white", linewidth=2))
        self.ax2.text(0,0, f"{score100:.0f}", ha="center", va="center", fontsize=20, fontweight="bold", color=COLORS["level"][n])
        self.ax2.text(0,-0.18, lvl, ha="center", va="center", fontsize=9, color=COLORS["text_sec"])
        self.ax2.set_title(f"العلامة النهائية  {avg:.2f}/5", fontsize=10, fontweight="bold", color=COLORS["text"], pad=10)

        self.fig.tight_layout()
        self.canvas.draw()
        self.update_overall()

    def switch_tab(self, tab):
        if tab=="eval":
            self.tab_results_frame.pack_forget()
            self.tab_eval_frame.pack(fill="both", expand=True)
            self.btn_eval.configure(fg_color=COLORS["primary"], text_color="white")
            self.btn_results.configure(fg_color="transparent", text_color=COLORS["text_sec"])
        else:
            self.tab_eval_frame.pack_forget()
            self.tab_results_frame.pack(fill="both", expand=True)
            self.btn_results.configure(fg_color=COLORS["primary"], text_color="white")
            self.btn_eval.configure(fg_color="transparent", text_color=COLORS["text_sec"])
            self.refresh_charts()

    # ---------- PDF Export ----------
    def export_pdf(self):
        if not self._validate_for_export():
            return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"KPI-2026-{self.org_var.get() or 'report'}.pdf", filetypes=[("PDF","*.pdf")])
        if not path: return
        try:
            # Prefer fpdf2 with HarfBuzz shaping (correct visual Arabic via OCR test)
            # Fallback to reportlab if fpdf2 not available
            try:
                self._build_pdf_fpdf(path)
            except Exception as e:
                print(f"fpdf2 failed, fallback to reportlab: {e}")
                import traceback
                traceback.print_exc()
                if not HAS_REPORTLAB:
                    raise Exception("fpdf2 failed and reportlab not installed. pip install fpdf2 uharfbuzz reportlab")
                self._build_pdf_reportlab(path)
            messagebox.showinfo("تم ✓", f"تم إنشاء PDF بنجاح:\n{path}")
            self.show_toast("✓ PDF تم إنشاؤه")
            try:
                if platform.system()=="Darwin": os.system(f'open "{path}"')
                elif platform.system()=="Windows": os.startfile(path)
                else: os.system(f'xdg-open "{path}"')
            except: pass
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("خطأ PDF", str(e))

    def _build_pdf_fpdf(self, out_path):
        """Professional PDF via fpdf2 + uharfbuzz - correct visual Arabic (verified via OCR)"""
        from fpdf import FPDF
        try:
            import uharfbuzz
            has_shaping = True
        except ImportError:
            has_shaping = False
        results, avg, score100, lvl, n = self.calc_results()
        # Find fonts
        noto_reg = BASE / "NotoNaskhArabic-Regular.ttf"
        noto_bold = BASE / "NotoNaskhArabic-Bold.ttf"
        if not noto_reg.exists():
            for p in [pathlib.Path.cwd(), pathlib.Path(__file__).parent, pathlib.Path(sys.executable).parent]:
                if (p / "NotoNaskhArabic-Regular.ttf").exists():
                    noto_reg = p / "NotoNaskhArabic-Regular.ttf"
                    noto_bold = p / "NotoNaskhArabic-Bold.ttf"
                    break
        pdf = FPDF(orientation='P', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=True, margin=12)
        # Register fonts
        if noto_reg.exists():
            pdf.add_font("Noto", "", str(noto_reg))
            pdf.add_font("Noto", "B", str(noto_bold))
            font = "Noto"
        else:
            pdf.add_font("Helvetica", "", "")
            font = "Helvetica"
        if has_shaping:
            try:
                pdf.set_text_shaping(True)
            except: pass
        # Use DejaVu for numbers fallback if needed, but Noto handles numbers

        def hex_to_rgb(h):
            h=h.lstrip("#")
            return tuple(int(h[i:i+2],16) for i in (0,2,4))

        def cell(text, w, h, align='C', style='', size=9, bg=None, border=False, color=None):
            pdf.set_font(font, style, size)
            if bg:
                r,g,b = hex_to_rgb(bg)
                pdf.set_fill_color(r,g,b)
            if color:
                r,g,b = hex_to_rgb(color)
                pdf.set_text_color(r,g,b)
            else:
                pdf.set_text_color(15,23,42)
            pdf.cell(w, h, text, align=align, border=border, fill=bg is not None, new_x="RIGHT", new_y="TOP")

        pdf.add_page()
        # Header bar (indigo)
        hdr_r, hdr_g, hdr_b = hex_to_rgb(COLORS["primary"])
        pdf.set_fill_color(hdr_r, hdr_g, hdr_b)
        pdf.rect(10, 8, 190, 14, 'F')
        pdf.set_y(10)
        pdf.set_font(font, "B", 11)
        pdf.set_text_color(255,255,255)
        pdf.cell(190, 10, f'KPI 2026  |  {self.org_var.get() or "تقرير التقييم المؤسسي"}', align='C')
        pdf.ln(18)
        # Title
        pdf.set_font(font, "B", 18)
        pdf.set_text_color(hdr_r, hdr_g, hdr_b)
        pdf.cell(0, 10, "نظام التقييم المؤسسي — 58 مؤشر  •  9 محاور", align='C', new_x="LMARGIN", new_y="NEXT")
        pdf.set_font(font, "", 8)
        pdf.set_text_color(100,116,139)
        pdf.cell(0, 6, f'{self.org_var.get() or "المؤسسة"}  •  {datetime.date.today().isoformat()}  •  KPI-Final 2026', align='C', new_x="LMARGIN", new_y="NEXT")
        pdf.ln(4)
        # Overall metrics - 4 boxes
        pdf.set_draw_color(*hex_to_rgb(COLORS["border"]))
        y0 = pdf.get_y()
        box_w = 46
        box_h = 18
        x0 = 11
        metrics = [
            (f"{avg:.2f}", "/ 5.00", "المتوسط العام"),
            (lvl, LEVELS_DESC[n], "المستوى"),
            (f"{score100:.0f}", "/ 100", "العلامة"),
            (f"{5-avg:.2f}", f"{100-score100:.0f} نقطة", "المتبقي"),
        ]
        for i, (val, sub, title) in enumerate(metrics):
            x = x0 + i*(box_w+2)
            # Bg for value
            lr,lg,lb = hex_to_rgb(COLORS["level"][n] if i<3 else COLORS["text"])
            # Card
            pdf.set_fill_color(248,250,252)
            pdf.set_draw_color(*hex_to_rgb(COLORS["border"]))
            pdf.rect(x, y0, box_w, box_h, 'DF')
            # Value
            pdf.set_xy(x, y0+2)
            pdf.set_font(font, "B", 13)
            pdf.set_text_color(lr,lg,lb)
            pdf.cell(box_w, 6, val, align='C')
            pdf.set_xy(x, y0+8)
            pdf.set_font(font, "", 7)
            pdf.set_text_color(100,116,139)
            pdf.cell(box_w, 4, sub, align='C')
            pdf.set_xy(x, y0+12)
            pdf.set_font(font, "", 7)
            pdf.set_text_color(71,85,105)
            pdf.cell(box_w, 4, title, align='C')
        pdf.set_y(y0+box_h+6)
        # Summary table header
        pdf.set_font(font, "B", 11)
        pdf.set_text_color(hdr_r, hdr_g, hdr_b)
        pdf.cell(0, 8, "ملخص النتائج حسب المحور", align='R', new_x="LMARGIN", new_y="NEXT")
        pdf.ln(1)
        # Table
        col_w = [68, 18, 22, 28, 22, 22]
        headers = ["المحور", "المؤشرات", "المتوسط", "المستوى", "من 100", "الفجوة"]
        # Header row
        pdf.set_fill_color(hdr_r, hdr_g, hdr_b)
        pdf.set_text_color(255,255,255)
        pdf.set_font(font, "B", 7)
        for i, h in enumerate(headers):
            pdf.cell(col_w[i], 8, h, align='C', border=1, fill=True)
        pdf.ln()
        # Rows
        pdf.set_font(font, "", 7)
        for idx, r in enumerate(results):
            # Alternate row color
            if idx % 2 == 0:
                pdf.set_fill_color(255,255,255)
            else:
                pdf.set_fill_color(248,250,252)
            # Level color for text
            lr,lg,lb = hex_to_rgb(COLORS["level"][r["level_n"]])
            # Domain
            pdf.set_text_color(15,23,42)
            pdf.cell(col_w[0], 7, r["domain"], align='R', border=1, fill=True)
            pdf.cell(col_w[1], 7, str(len([d for d in domains if d["domain"]==r["domain"]][0]["indicators"])), align='C', border=1, fill=True)
            pdf.cell(col_w[2], 7, f"{r['avg']:.2f}", align='C', border=1, fill=True)
            # Level with bg
            pdf.set_fill_color(*hex_to_rgb(COLORS["level_bg"][r["level_n"]]))
            pdf.set_text_color(lr,lg,lb)
            pdf.set_font(font, "B", 7)
            pdf.cell(col_w[3], 7, r["level"], align='C', border=1, fill=True)
            pdf.set_font(font, "", 7)
            pdf.set_text_color(15,23,42)
            if idx % 2 == 0:
                pdf.set_fill_color(255,255,255)
            else:
                pdf.set_fill_color(248,250,252)
            pdf.cell(col_w[4], 7, f"{r['score100']:.0f}", align='C', border=1, fill=True)
            pdf.cell(col_w[5], 7, f"{r['gap']:.2f}", align='C', border=1, fill=True)
            pdf.ln()
        # Overall
        pdf.set_fill_color(238,242,255)
        pdf.set_font(font, "B", 7)
        pdf.cell(col_w[0], 7, "الإجمالي", align='R', border=1, fill=True)
        pdf.cell(col_w[1], 7, "58", align='C', border=1, fill=True)
        pdf.cell(col_w[2], 7, f"{avg:.2f}", align='C', border=1, fill=True)
        pdf.cell(col_w[3], 7, lvl, align='C', border=1, fill=True)
        pdf.cell(col_w[4], 7, f"{score100:.0f}", align='C', border=1, fill=True)
        pdf.cell(col_w[5], 7, f"{5-avg:.2f}", align='C', border=1, fill=True)
        pdf.ln(8)
        # Weakest
        weakest = min(results, key=lambda x: x["avg"])
        pdf.set_fill_color(254,242,242)
        pdf.set_draw_color(*hex_to_rgb(COLORS["error"]))
        pdf.set_font(font, "", 7)
        pdf.set_text_color(*hex_to_rgb(COLORS["error"]))
        pdf.cell(0, 7, f"  ⚠  أضعف محور: {weakest['domain']}  ({weakest['avg']} — {weakest['level']})  •  يحتاج {weakest['gap']} للوصول للديناميكي  ", align='R', border=1, fill=True)
        pdf.ln(10)
        # Charts
        pdf.set_font(font, "B", 11)
        pdf.set_text_color(hdr_r, hdr_g, hdr_b)
        pdf.cell(0, 8, "الرسوم البيانية", align='R', new_x="LMARGIN", new_y="NEXT")
        # Generate charts as images
        try:
            fig1, ax1 = plt.subplots(figsize=(7, 2.6), dpi=170)
            fig1.patch.set_facecolor("white")
            names = [r["domain"].split()[0][:12] for r in results]
            scores = [r["score100"] for r in results]
            cols = [COLORS["level"][r["level_n"]] for r in results]
            bars = ax1.bar(names, scores, color=cols, edgecolor="white", linewidth=1, width=0.6)
            ax1.set_ylim(0,105)
            ax1.set_ylabel("من 100", fontsize=7)
            ax1.tick_params(axis='x', rotation=18, labelsize=6)
            ax1.grid(axis='y', linestyle="--", alpha=0.3)
            for b,v in zip(bars, scores):
                ax1.text(b.get_x()+b.get_width()/2, v+1, f"{v:.0f}", ha="center", fontsize=6, fontweight="bold")
            fig1.tight_layout()
            buf1 = io.BytesIO()
            fig1.savefig(buf1, format="png", bbox_inches="tight", facecolor="white", dpi=170)
            plt.close(fig1)
            buf1.seek(0)
            # Save temp
            tmp1 = pathlib.Path("/tmp/kpi_bar.png")
            tmp1.write_bytes(buf1.getvalue())
            pdf.image(str(tmp1), w=180, h=65)
            pdf.ln(2)
            # Donut
            fig2, ax2 = plt.subplots(figsize=(3, 3), dpi=170)
            fig2.patch.set_facecolor("white")
            ax2.pie([score100, 100-score100], colors=[COLORS["level"][n], "#F1F5F9"], startangle=90, counterclock=False, wedgeprops=dict(width=0.38, edgecolor="white", linewidth=2))
            ax2.text(0,0, f"{score100:.0f}", ha="center", va="center", fontsize=18, fontweight="bold", color=COLORS["level"][n])
            ax2.text(0,-0.18, lvl, ha="center", va="center", fontsize=8)
            ax2.set_title(f"العلامة النهائية  {avg:.2f}/5", fontsize=9, fontweight="bold")
            fig2.tight_layout()
            buf2 = io.BytesIO()
            fig2.savefig(buf2, format="png", bbox_inches="tight", facecolor="white", dpi=170)
            plt.close(fig2)
            buf2.seek(0)
            tmp2 = pathlib.Path("/tmp/kpi_donut.png")
            tmp2.write_bytes(buf2.getvalue())
            # Center
            x = (210 - 60)/2
            pdf.image(str(tmp2), x=x, w=60, h=60)
        except Exception as e:
            pdf.set_font(font, "", 7)
            pdf.cell(0, 6, f"تعذر إنشاء الرسوم: {e}", align='C')
        pdf.add_page()
        # Details per domain
        pdf.set_font(font, "B", 11)
        pdf.set_text_color(hdr_r, hdr_g, hdr_b)
        pdf.cell(0, 8, "تفاصيل المؤشرات", align='R', new_x="LMARGIN", new_y="NEXT")
        pdf.ln(2)
        for d in domains:
            vals = list(self.scores[d["domain"]].values())
            av = sum(vals)/len(vals)
            lv, nn = get_level(av)
            # Domain header
            r1,g1,b1 = hex_to_rgb(COLORS["level"][nn])
            pr,pg,pb = hex_to_rgb(COLORS["level_bg"][nn])
            pdf.set_fill_color(pr,pg,pb)
            pdf.set_draw_color(r1,g1,b1)
            pdf.set_font(font, "B", 9)
            pdf.set_text_color(r1,g1,b1)
            pdf.cell(0, 8, f"  {d['domain']}  —  {len(d['indicators'])} مؤشر  •  {av:.2f} {lv} • {av*20:.0f}/100  ", align='R', border=1, fill=True)
            pdf.ln(8)
            # Table header
            pdf.set_fill_color(241,245,249)
            pdf.set_text_color(15,23,42)
            pdf.set_font(font, "B", 6.5)
            hdr_w = [10, 42, 14, 20, 104]
            hdrs = ["#", "المؤشر", "الدرجة", "المستوى", "الوصف"]
            for i, h in enumerate(hdrs):
                pdf.cell(hdr_w[i], 7, h, align='C', border=1, fill=True)
            pdf.ln()
            pdf.set_font(font, "", 6.5)
            for ind in d["indicators"]:
                sc = self.scores[d["domain"]][ind["name"]]
                lv2, nn2 = get_level(sc)
                # Alternate
                if ind["id"] % 2 == 0:
                    pdf.set_fill_color(255,255,255)
                else:
                    pdf.set_fill_color(248,250,252)
                # Check page break
                if pdf.get_y() > 265:
                    pdf.add_page()
                h = 7
                # Calculate needed height for description
                desc = ind["levels"][sc-1][:140] + "…"
                # Simple: use multi_cell for description, but need row height
                # For simplicity, fixed height + truncate
                pdf.set_text_color(15,23,42)
                pdf.cell(hdr_w[0], h, str(ind["id"]), align='C', border=1, fill=True)
                pdf.cell(hdr_w[1], h, ind["name"][:22], align='R', border=1, fill=True)
                # Score with color
                pdf.set_text_color(*hex_to_rgb(COLORS["level"][nn2]))
                pdf.set_font(font, "B", 7)
                pdf.cell(hdr_w[2], h, str(sc), align='C', border=1, fill=True)
                pdf.set_font(font, "", 6.5)
                pdf.set_fill_color(*hex_to_rgb(COLORS["level_bg"][nn2]))
                pdf.set_text_color(*hex_to_rgb(COLORS["level"][nn2]))
                pdf.cell(hdr_w[3], h, lv2, align='C', border=1, fill=True)
                pdf.set_text_color(71,85,105)
                if ind["id"] % 2 == 0:
                    pdf.set_fill_color(255,255,255)
                else:
                    pdf.set_fill_color(248,250,252)
                pdf.cell(hdr_w[4], h, desc[:55], align='R', border=1, fill=True)
                pdf.ln()
            pdf.ln(4)
        # Footer
        pdf.set_font(font, "", 6)
        pdf.set_text_color(148,163,184)
        pdf.cell(0, 6, f"تم إنشاء هذا التقرير تلقائياً بواسطة نظام KPI 2026  •  {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}  •  {self.org_var.get() or '—'}", align='C')
        pdf.output(out_path)

    def _build_pdf_reportlab(self, out_path):
        results, avg, score100, lvl, n = self.calc_results()
        # Colors
        P = HexColor(COLORS["primary"])
        Pdark = HexColor(COLORS["primary_dark"])
        level_colors = {k: HexColor(v) for k,v in COLORS["level"].items()}
        level_bg = {k: HexColor(v) for k,v in COLORS["level_bg"].items()}
        # Register premium Arabic font (Noto Naskh - Google, bundled)
        font_name = "Helvetica"
        font_bold = "Helvetica-Bold"
        # Prefer bundled Noto Naskh Arabic (high quality, works Mac+Windows)
        try:
            noto_reg = BASE / "NotoNaskhArabic-Regular.ttf"
            noto_bold = BASE / "NotoNaskhArabic-Bold.ttf"
            # Also check cwd and exe dir
            if not noto_reg.exists():
                for p in [pathlib.Path.cwd(), pathlib.Path(__file__).parent, pathlib.Path(sys.executable).parent]:
                    if (p / "NotoNaskhArabic-Regular.ttf").exists():
                        noto_reg = p / "NotoNaskhArabic-Regular.ttf"
                        noto_bold = p / "NotoNaskhArabic-Bold.ttf"
                        break
            if noto_reg.exists() and noto_bold.exists():
                pdfmetrics.registerFont(TTFont("NotoNaskh", str(noto_reg)))
                pdfmetrics.registerFont(TTFont("NotoNaskh-Bold", str(noto_bold)))
                # Map for bold fallback
                pdfmetrics.registerFontFamily("NotoNaskh", normal="NotoNaskh", bold="NotoNaskh-Bold")
                font_name = "NotoNaskh"
                font_bold = "NotoNaskh-Bold"
            else:
                # Fallback: Geeza Pro (Mac) or DejaVu
                candidates = [
                    ("/System/Library/Fonts/GeezaPro.ttc", 0),
                    ("/System/Library/Fonts/Supplemental/Arial Unicode.ttf", None),
                ]
                for fp, idx in candidates:
                    if pathlib.Path(fp).exists():
                        try:
                            if idx is not None:
                                pdfmetrics.registerFont(TTFont("Arabic", fp, subfontIndex=idx))
                            else:
                                pdfmetrics.registerFont(TTFont("Arabic", fp))
                            font_name = "Arabic"
                            font_bold = "Arabic"
                            break
                        except: continue
                if font_name == "Helvetica":
                    try:
                        import matplotlib.font_manager as mfm
                        dejavu = mfm.findfont("DejaVu Sans")
                        if dejavu and pathlib.Path(dejavu).exists():
                            pdfmetrics.registerFont(TTFont("Arabic", dejavu))
                            font_name = "Arabic"
                            font_bold = "Arabic"
                    except: pass
        except Exception as e:
            print(f"Font register fallback: {e}")
            pass

        doc = SimpleDocTemplate(out_path, pagesize=A4, topMargin=14*mm, bottomMargin=14*mm, leftMargin=14*mm, rightMargin=14*mm,
                                title=f"KPI 2026 - {self.org_var.get()}", author="KPI System")
        styles = getSampleStyleSheet()
        s_title = ParagraphStyle("Title", parent=styles["Title"], fontName=font_bold, fontSize=22, textColor=P, alignment=TA_CENTER, spaceAfter=2*mm)
        s_sub = ParagraphStyle("Sub", parent=styles["Normal"], fontName=font_name, fontSize=9, textColor=HexColor(COLORS["text_muted"]), alignment=TA_CENTER, spaceAfter=4*mm)
        s_h1 = ParagraphStyle("H1", parent=styles["Heading1"], fontName=font_bold, fontSize=13, textColor=Pdark, spaceBefore=6*mm, spaceAfter=3*mm, alignment=TA_RIGHT, borderPadding=(0,0,6))
        s_h2 = ParagraphStyle("H2", parent=styles["Heading2"], fontName=font_bold, fontSize=10, textColor=HexColor(COLORS["text"]), spaceBefore=4*mm, spaceAfter=2*mm, alignment=TA_RIGHT)
        s_body = ParagraphStyle("Body", parent=styles["Normal"], fontName=font_name, fontSize=8, textColor=HexColor(COLORS["text_sec"]), leading=12, alignment=TA_RIGHT, wordWrap="CJK")
        s_cell = ParagraphStyle("Cell", parent=s_body, fontSize=7.5, leading=10, alignment=TA_CENTER)
        s_cell_right = ParagraphStyle("CellR", parent=s_cell, alignment=TA_RIGHT)
        s_badge = ParagraphStyle("Badge", parent=s_cell, fontName=font_bold, fontSize=7, textColor=white, alignment=TA_CENTER)

        story = []

        # Helper to wrap text - use LOGICAL order (no reshaping) for proper Arabic
        # Noto font + ReportLab handles shaping via embedded font;  would create
        # presentation forms (ﻬﻤﻴﻴﻘﺗ) which copy as reversed visual and look "not good"
        def Ptxt(txt, style=s_body):
            # For PDF, store logical Arabic directly - font shaping handles joining
            # Only use  if you need visual fallback for old viewers
            return Paragraph(str(txt), style)
        # Keep ar for reference but don't use for main content
        def PtxtVisual(txt, style=s_body):
            return Paragraph(txt, style)

        # --- Cover Header ---
        # Top bar
        header_data = [[Ptxt(f'<font color="white"><b>KPI 2026</b></font>', ParagraphStyle("h", parent=s_body, textColor=white, alignment=TA_CENTER, fontName=font_bold)),
                        Ptxt(f'<font color="white">{self.org_var.get() or "تقرير التقييم المؤسسي"}</font>', ParagraphStyle("h2", parent=s_body, textColor=white, alignment=TA_CENTER))]]
        t = Table(header_data, colWidths=[35*mm, 155*mm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), P),
            ("BOX", (0,0), (-1,-1), 0.5, Pdark),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING", (0,0), (-1,-1), 4), ("RIGHTPADDING", (0,0), (-1,-1), 4),
            ("TOPPADDING", (0,0), (-1,-1), 6), ("BOTTOMPADDING", (0,0), (-1,-1), 6),
            ("ROUNDEDCORNERS", [6,6,6,6]),
        ]))
        story.append(t)
        story.append(Spacer(1, 6*mm))

        story.append(Ptxt("نظام التقييم المؤسسي — 58 مؤشر  •  9 محاور", s_title))
        story.append(Ptxt(f"{self.org_var.get() or 'المؤسسة'}  •  {datetime.date.today().isoformat()}  •  KPI-Final 2026", s_sub))
        story.append(HRFlowable(width="100%", thickness=0.5, color=HexColor(COLORS["border"]), spaceAfter=4*mm))

        # --- Overall Metrics (4 cards) ---
        lvl_color = level_colors[n]
        metrics = [
            [Ptxt(f'<font size=16 color="{COLORS["level"][n]}"><b>{avg:.2f}</b></font><br/><font size=7 color="{COLORS["text_muted"]}">/ 5.00</font>', s_cell),
             Ptxt(f'<font size=14 color="{COLORS["level"][n]}"><b>{lvl}</b></font><br/><font size=7 color="{COLORS["text_muted"]}">{LEVELS_DESC[n]}</font>', s_cell),
             Ptxt(f'<font size=16 color="{COLORS["level"][n]}"><b>{score100:.0f}</b></font><br/><font size=7 color="{COLORS["text_muted"]}">/ 100</font>', s_cell),
             Ptxt(f'<font size=14 color="{COLORS["text"]}"><b>{5-avg:.2f}</b></font><br/><font size=7 color="{COLORS["text_muted"]}">{"المتبقي للتميز"}</font>', s_cell)],
            [Ptxt("المتوسط العام", s_cell), Ptxt("المستوى", s_cell), Ptxt("العلامة", s_cell), Ptxt("الفجوة", s_cell)]
        ]
        t = Table(metrics, colWidths=[45*mm, 45*mm, 45*mm, 45*mm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), HexColor(COLORS["bg"])),
            ("BACKGROUND", (0,1), (-1,1), white),
            ("BOX", (0,0), (-1,-1), 0.7, HexColor(COLORS["border"])),
            ("INNERGRID", (0,0), (-1,-1), 0.4, HexColor(COLORS["border"])),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING", (0,0), (-1,-1), 6), ("BOTTOMPADDING", (0,0), (-1,-1), 6),
            ("ROUNDEDCORNERS", [8,8,8,8]),
        ]))
        story.append(t)
        story.append(Spacer(1, 5*mm))

        # --- Summary Table (9 domains) ---
        story.append(Ptxt("ملخص النتائج حسب المحور", s_h1))
        hdr = [Ptxt("المحور", s_cell), Ptxt("المؤشرات", s_cell), Ptxt("المتوسط", s_cell), Ptxt("المستوى", s_cell), Ptxt("من 100", s_cell), Ptxt("الفجوة", s_cell)]
        data = [hdr]
        for r in results:
            # level badge cell with background
            data.append([
                Ptxt(r["domain"], s_cell_right),
                Ptxt(str(next((d for d in domains if d["domain"]==r["domain"]), {"indicators":[]})["indicators"].__len__() or r.get("count",5)), s_cell),
                Ptxt(f"{r['avg']:.2f}", s_cell),
                Ptxt(r["level"], ParagraphStyle("b", parent=s_badge, backColor=level_colors[r["level_n"]], borderPadding=(2,4,2))),
                Ptxt(f"{r['score100']:.0f}", s_cell),
                Ptxt(f"{r['gap']:.2f}", s_cell),
            ])
        # overall row
        data.append([Ptxt(f"<b>{'الإجمالي'}</b>", s_cell), Ptxt(f"<b>58</b>", s_cell), Ptxt(f"<b>{avg:.2f}</b>", s_cell), Ptxt(f"<b>{lvl}</b>", s_cell), Ptxt(f"<b>{score100:.0f}</b>", s_cell), Ptxt(f"<b>{5-avg:.2f}</b>", s_cell)])
        colW = [58*mm, 20*mm, 22*mm, 28*mm, 22*mm, 22*mm]
        t = Table(data, colWidths=colW, repeatRows=1)
        style = TableStyle([
            ("BACKGROUND", (0,0), (-1,0), P),
            ("TEXTCOLOR", (0,0), (-1,0), white),
            ("FONTNAME", (0,0), (-1,0), font_bold),
            ("FONTSIZE", (0,0), (-1,0), 8),
            ("ALIGN", (0,0), (-1,-1), "CENTER"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("GRID", (0,0), (-1,-1), 0.4, HexColor(COLORS["border"])),
            ("ROWBACKGROUNDS", (0,1), (-1,-2), [white, HexColor(COLORS["bg"])]),
            ("BACKGROUND", (0,-1), (-1,-1), HexColor("#EEF2FF")),
            ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ("LEFTPADDING", (0,0), (-1,-1), 3), ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ])
        # color level column
        for i, r in enumerate(results, 1):
            style.add("BACKGROUND", (3,i), (3,i), level_bg[r["level_n"]])
            style.add("TEXTCOLOR", (3,i), (3,i), level_colors[r["level_n"]])
        t.setStyle(style)
        story.append(t)
        # weakest highlight
        weakest = min(results, key=lambda x: x["avg"])
        story.append(Spacer(1, 3*mm))
        warn_style = ParagraphStyle("warn", parent=s_body, backColor=HexColor("#FEF2F2"), borderColor=HexColor(COLORS["error"]), borderWidth=0.5, borderPadding=(6,6,6), textColor=HexColor(COLORS["error"]), fontSize=8, alignment=TA_RIGHT)
        story.append(Paragraph(f"⚠ أضعف محور: {weakest['domain']}  ({weakest['avg']} — {weakest['level']})  •  يحتاج {weakest['gap']} للوصول للديناميكي", warn_style))
        story.append(Spacer(1, 6*mm))

        # --- Charts ---
        story.append(Ptxt("الرسوم البيانية", s_h1))
        # Generate charts to images
        try:
            # Bar
            fig1, ax1 = plt.subplots(figsize=(7, 2.8), dpi=150)
            fig1.patch.set_facecolor("white")
            names = [r["domain"].split()[0][:12] for r in results]
            scores = [r["score100"] for r in results]
            cols = [COLORS["level"][r["level_n"]] for r in results]
            bars = ax1.bar(names, scores, color=cols, edgecolor="white", linewidth=1, width=0.6)
            ax1.set_ylim(0,105)
            ax1.set_ylabel("من 100", fontsize=8)
            ax1.set_title("العلامات حسب المحور", fontsize=10, fontweight="bold")
            ax1.tick_params(axis='x', rotation=18, labelsize=7)
            ax1.grid(axis='y', linestyle="--", alpha=0.3)
            for b,v in zip(bars, scores):
                ax1.text(b.get_x()+b.get_width()/2, v+1, f"{v:.0f}", ha="center", fontsize=7, fontweight="bold")
            fig1.tight_layout()
            buf1 = io.BytesIO()
            fig1.savefig(buf1, format="png", bbox_inches="tight", facecolor="white")
            plt.close(fig1)
            buf1.seek(0)
            img1 = RLImage(buf1, width=170*mm, height=68*mm)
            story.append(img1)
            story.append(Spacer(1, 4*mm))
            # Donut
            fig2, ax2 = plt.subplots(figsize=(3.2, 3.2), dpi=150)
            fig2.patch.set_facecolor("white")
            ax2.pie([score100, 100-score100], colors=[COLORS["level"][n], "#F1F5F9"], startangle=90, counterclock=False, wedgeprops=dict(width=0.38, edgecolor="white", linewidth=2))
            ax2.text(0,0, f"{score100:.0f}", ha="center", va="center", fontsize=22, fontweight="bold", color=COLORS["level"][n])
            ax2.text(0,-0.2, lvl, ha="center", va="center", fontsize=9)
            ax2.set_title(f"العلامة النهائية  {avg:.2f}/5", fontsize=10, fontweight="bold")
            fig2.tight_layout()
            buf2 = io.BytesIO()
            fig2.savefig(buf2, format="png", bbox_inches="tight", facecolor="white")
            plt.close(fig2)
            buf2.seek(0)
            img2 = RLImage(buf2, width=65*mm, height=65*mm)
            # center it
            t = Table([[img2]], colWidths=[180*mm])
            t.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER")]))
            story.append(t)
        except Exception as e:
            story.append(Ptxt(f"تعذر إنشاء الرسوم: {e}", s_body))
        story.append(Spacer(1, 6*mm))

        # --- Details per domain ---
        story.append(Ptxt("تفاصيل المؤشرات", s_h1))
        for d in domains:
            vals = list(self.scores[d["domain"]].values())
            av = sum(vals)/len(vals)
            lv, nn = get_level(av)
            # domain header
            hdr_data = [[Ptxt(f"<b>{d['domain']}</b>  —  {len(d['indicators'])} {'مؤشر'}  •  {av:.2f} {lv} • {av*20:.0f}/100", ParagraphStyle("dh", parent=s_body, fontName=font_bold, textColor=level_colors[nn], fontSize=9, backColor=level_bg[nn], borderPadding=(6,6,6)))]]
            th = Table(hdr_data, colWidths=[180*mm])
            th.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1), level_bg[nn]), ("BOX",(0,0),(-1,-1),0.5, level_colors[nn]), ("ROUNDEDCORNERS",[6,6,6,6])]))
            story.append(th)
            story.append(Spacer(1, 2*mm))
            # table of indicators
            idata = [[Ptxt("#", s_cell), Ptxt("المؤشر", s_cell), Ptxt("الدرجة", s_cell), Ptxt("المستوى", s_cell), Ptxt("الوصف", s_cell_right)]]
            for ind in d["indicators"]:
                sc = self.scores[d["domain"]][ind["name"]]
                lv2, nn2 = get_level(sc)
                idata.append([
                    Ptxt(str(ind["id"]), s_cell),
                    Ptxt(ind["name"], s_cell_right),
                    Ptxt(f"<b>{sc}</b>", ParagraphStyle("sc", parent=s_cell, textColor=level_colors[nn2], fontName=font_bold)),
                    Ptxt(lv2, ParagraphStyle("lv", parent=s_badge, backColor=level_colors[nn2], textColor=white)),
                    Ptxt(ind["levels"][sc-1][:120] + "…", ParagraphStyle("desc", parent=s_body, fontSize=6, leading=8, alignment=TA_RIGHT)),
                ])
            t = Table(idata, colWidths=[10*mm, 42*mm, 14*mm, 20*mm, 94*mm], repeatRows=1)
            ts = TableStyle([
                ("BACKGROUND",(0,0),(-1,0), HexColor("#F1F5F9")),
                ("TEXTCOLOR",(0,0),(-1,0), HexColor(COLORS["text"])),
                ("FONTNAME",(0,0),(-1,0), font_bold),
                ("FONTSIZE",(0,0),(-1,0),7),
                ("GRID",(0,0),(-1,-1),0.3, HexColor(COLORS["border"])),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[white, HexColor(COLORS["bg"])]),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
                ("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),
            ])
            t.setStyle(ts)
            story.append(t)
            story.append(Spacer(1, 4*mm))

        # Footer
        story.append(HRFlowable(width="100%", thickness=0.5, color=HexColor(COLORS["border"]), spaceAfter=3*mm, spaceBefore=4*mm))
        story.append(Ptxt(f"{'تم إنشاء هذا التقرير تلقائياً بواسطة نظام KPI 2026'}  •  {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}  •  {self.org_var.get() or '—'}", ParagraphStyle("foot", parent=s_body, fontSize=7, textColor=HexColor(COLORS["text_muted"]), alignment=TA_CENTER)))

        def footer(canvas, doc):
            canvas.saveState()
            canvas.setFont(font_name, 7)
            canvas.setFillColor(HexColor(COLORS["text_muted"]))
            canvas.drawCentredString(A4[0]/2, 10*mm, f"صفحة {doc.page}  •  KPI 2026  •  {self.org_var.get() or ''}")
            # top line
            canvas.setStrokeColor(P)
            canvas.setLineWidth(0.7)
            canvas.line(14*mm, A4[1]-10*mm, A4[0]-14*mm, A4[1]-10*mm)
            canvas.restoreState()

        doc.build(story, onFirstPage=footer, onLaterPages=footer)

    # ---------- Exports (Item 2: with validation) ----------
    def export_excel(self):
        if not self._validate_for_export():
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"KPI-2026-{self.org_var.get() or 'result'}.xlsx", filetypes=[("Excel","*.xlsx")])
        if not path: return
        try:
            self._build_excel(path)
            messagebox.showinfo("تم ✓", f"تم الحفظ بنجاح:\n{path}")
            self.show_toast("✓ تم تصدير Excel")
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def _build_excel(self, out_path):
        wb=openpyxl.load_workbook(EXCEL)
        results, avg, score100, lvl, n = self.calc_results()
        res_map={r["domain"]:r for r in results}
        for d in domains:
            ws=wb[d["sheet"]]
            hdr=None
            for idx,row in enumerate(ws.iter_rows(values_only=True),1):
                if row and any("1 - ناشئ" in str(c) for c in row if c):
                    hdr=idx; break
            if not hdr: continue
            r=hdr+1
            for ind in d["indicators"]:
                sc=self.scores[d["domain"]][ind["name"]]
                ws.cell(row=r,column=8,value=sc)
                ws.cell(row=r,column=9,value=sc*20)
                r+=1
            for row in ws.iter_rows(min_row=r,max_row=ws.max_row):
                for cell in row:
                    if cell.value and "متوسط المجال" in str(cell.value):
                        v=sum(self.scores[d["domain"]].values())/len(d["indicators"])
                        ws.cell(row=cell.row,column=cell.column+1,value=round(v,2))
                        ws.cell(row=cell.row,column=cell.column+2,value=round(v*20,1))
        if "النتائج" in wb.sheetnames:
            ws=wb["النتائج"]
            for row in ws.iter_rows(min_row=5,max_row=15):
                dc=row[0].value
                if dc and str(dc).strip():
                    for k,v in res_map.items():
                        if str(dc).strip() in k or k in str(dc).strip():
                            row[2].value=v["avg"]; row[3].value=v["level"]; row[4].value=v["score100"]; break
                    if "الترويج" in str(dc) and res_map.get("الترويج والخبرة في تعلم الكبار"):
                        m=res_map["الترويج والخبرة في تعلم الكبار"]
                        row[2].value=m["avg"]; row[3].value=m["level"]; row[4].value=m["score100"]
            for row in ws.iter_rows(min_row=14,max_row=17):
                if row[0].value and "المتوسط العام" in str(row[0].value):
                    row[2].value=avg; row[3].value=lvl; row[4].value=score100
        wb.save(out_path)

    def export_csv(self):
        import csv
        path=filedialog.asksaveasfilename(defaultextension=".csv", initialfile="KPI-ملخص.csv", filetypes=[("CSV","*.csv")])
        if not path: return
        results, avg, score100, lvl, _ = self.calc_results()
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w=csv.writer(f)
            w.writerow(["المحور","المتوسط","المستوى","العلامة/100","الفجوة"])
            for r in results: w.writerow([r["domain"],r["avg"],r["level"],r["score100"],r["gap"]])
            w.writerow(["الإجمالي",avg,lvl,score100,round(5-avg,2)])
        messagebox.showinfo("تم ✓","تم حفظ CSV")
        self.show_toast("✓ CSV محفوظ")

    def save_json(self):
        path=filedialog.asksaveasfilename(defaultextension=".json", initialfile="KPI-scores.json", filetypes=[("JSON","*.json")])
        if not path: return
        data={"org":self.org_var.get(),"scores":self.scores}
        pathlib.Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        messagebox.showinfo("تم ✓", f"حفظ JSON:\n{path}")
        self.show_toast("✓ JSON محفوظ")

    def load_json(self):
        path=filedialog.askopenfilename(filetypes=[("JSON","*.json")])
        if not path: return
        data=json.loads(pathlib.Path(path).read_text(encoding="utf-8"))
        sc=data.get("scores", data)
        self.scores=sc
        self.org_var.set(data.get("org",""))
        for k,var in self.vars.items():
            dom,name=k
            if dom in sc and name in sc[dom]:
                var.set(sc[dom][name])
                self.update_desc(var)
        self.refresh_charts()
        self.build_sidebar()
        self.show_domain()
        messagebox.showinfo("تم ✓","تم تحميل JSON")
        self.show_toast("✓ تم التحميل")

    def random_demo(self):
        import random
        for d in domains:
            for ind in d["indicators"]:
                v=random.randint(1,5)
                self.scores[d["domain"]][ind["name"]]=v
                k=(d["domain"],ind["name"])
                if k in self.vars:
                    self.vars[k].set(v)
                    self.update_desc(self.vars[k])
        self.refresh_charts()
        self.build_sidebar()
        self.show_domain()
        self.show_toast("🎲 تم توليد بيانات عشوائية")
        self._autosave()

    # ---------- Autosave & Close Handling (Item 1) ----------
    def _has_unsaved_changes(self):
        for d in domains:
            for v in self.scores[d["domain"]].values():
                if v != 1:
                    return True
        return bool(self.org_var.get().strip())

    def _autosave(self):
        try:
            data = {"org": self.org_var.get(), "scores": self.scores, "ts": datetime.datetime.now().isoformat()}
            tmp = AUTOSAVE_PATH.with_suffix(".tmp")
            tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            tmp.replace(AUTOSAVE_PATH)
            self._autosave_dirty = True
        except Exception as e:
            print(f"Autosave failed: {e}")

    def _load_autosave(self):
        try:
            if AUTOSAVE_PATH.exists():
                data = json.loads(AUTOSAVE_PATH.read_text(encoding="utf-8"))
                sc = data.get("scores")
                if sc and isinstance(sc, dict):
                    for d in domains:
                        if d["domain"] in sc:
                            for k, v in sc[d["domain"]].items():
                                if k in self.scores[d["domain"]] and v in (1,2,3,4,5):
                                    self.scores[d["domain"]][k] = v
                                    # Also update the IntVar if it exists
                                    key = (d["domain"], k)
                                    if key in self.vars:
                                        self.vars[key].set(v)
                    self.org_var.set(data.get("org", ""))
                    self._autosave_dirty = False
                    # Update UI to reflect loaded values
                    self.update_overall()
                    self.build_sidebar()
                    self.show_domain()
                    self._update_onboarding_progress()
                    self.root.after(600, lambda: self.show_toast("✓ تم استعادة الجلسة السابقة"))
        except Exception as e:
            print(f"Load autosave failed: {e}")

    def _on_close(self):
        if self._has_unsaved_changes():
            ans = messagebox.askyesnocancel(
                "حفظ العمل؟",
                "لديك تقييم غير محفوظ.\nهل تريد حفظه قبل الإغلاق؟\n\n• نعم = حفظ ومتابعة الإغلاق\n• لا = تجاهل والإغلاق\n• إلغاء = البقاء في التطبيق",
                icon="warning"
            )
            if ans is None:
                return
            if ans:
                try:
                    self._autosave()
                    # Offer explicit file save as well
                    path = filedialog.asksaveasfilename(defaultextension=".json", initialfile=f"KPI-{self.org_var.get().strip() or 'autosave'}.json", filetypes=[("JSON","*.json")])
                    if path:
                        pathlib.Path(path).write_text(json.dumps({"org": self.org_var.get(), "scores": self.scores}, ensure_ascii=False, indent=2), encoding="utf-8")
                except Exception as e:
                    messagebox.showerror("خطأ", f"فشل الحفظ: {e}")
                    return
        self.root.destroy()

    # ---------- Validation (Item 2) ----------
    def _validate_for_export(self):
        # 1. Org name required (friendly warning)
        if not self.org_var.get().strip():
            ans = messagebox.askyesnocancel(
                "اسم المؤسسة مطلوب",
                "لم تدخل اسم المؤسسة.\n\nهل تريد المتابعة بدون اسم؟\n\n• نعم = متابعة بدون اسم\n• لا = العودة لإدخال الاسم\n• إلغاء = إلغاء التصدير",
                icon="warning"
            )
            if ans is None or not ans:
                # Focus the org entry for easy correction
                try:
                    self.entry_org.focus_set()
                    self.entry_org.configure(border_color=COLORS["error"], border_width=2)
                    self.root.after(2000, lambda: self.entry_org.configure(border_color=COLORS["border"], border_width=0))
                except: pass
                return False
        # 2. All 58 still 1 (likely not completed)
        all_one = True
        for d in domains:
            for v in self.scores[d["domain"]].values():
                if v != 1:
                    all_one = False
                    break
            if not all_one:
                break
        if all_one:
            ans = messagebox.askyesnocancel(
                "تأكيد التصدير",
                "جميع المؤشرات لا تزال 1 (ناشئ — 20/100).\n\nيبدو أنك لم تكمل التقييم بعد.\nهل أنت متأكد من التصدير؟\n\n• نعم = تصدير على أي حال\n• لا = العودة لإكمال التقييم\n• إلغاء = إلغاء التصدير",
                icon="warning"
            )
            if ans is None or not ans:
                return False
        return True

    # ---------- Onboarding (Item 4) ----------
    def _show_onboarding_if_needed(self):
        flag = pathlib.Path.home() / ".kpi_onboarded"
        if flag.exists():
            return
        self.root.after(500, self._show_onboarding_modal)

    def _show_onboarding_modal(self):
        modal = ctk.CTkToplevel(self.root)
        modal.title("مرحباً بك")
        modal.geometry("520x380")
        modal.transient(self.root)
        modal.grab_set()
        try:
            modal.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() - 520)//2
            y = self.root.winfo_y() + (self.root.winfo_height() - 380)//2
            modal.geometry(f"520x380+{x}+{y}")
        except: pass
        ctk.CTkLabel(modal, text="مرحباً بك في KPI 2026", font=(FONTS["title"][0], 18, "bold"), text_color=COLORS["primary"]).pack(pady=(20,4))
        ctk.CTkLabel(modal, text="نظام التقييم المؤسسي — 58 مؤشر  •  9 محاور", text_color=COLORS["text_muted"], font=FONTS["small"]).pack()
        ctk.CTkLabel(modal, text="3 خطوات بسيطة:", text_color=COLORS["text"], font=FONTS["heading"]).pack(pady=(14,6))
        steps = [
            ("١", "أدخل اسم المؤسسة", "في الأعلى — حقل أبيض"),
            ("٢", "قيّم كل مؤشر 1-5", "اختر ناشئ → ديناميكي، يحفظ تلقائياً"),
            ("٣", "صدّر النتائج", "PDF أو Excel في تبويب النتائج"),
        ]
        for num, title, desc in steps:
            row = ctk.CTkFrame(modal, fg_color="transparent")
            row.pack(fill="x", padx=24, pady=6)
            circle = ctk.CTkFrame(row, fg_color=COLORS["primary"], corner_radius=16, width=32, height=32)
            circle.pack(side="left", padx=(0,10))
            circle.pack_propagate(False)
            ctk.CTkLabel(circle, text=num, text_color="white", font=(FONTS["heading"][0], 12, "bold")).place(relx=0.5, rely=0.5, anchor="center")
            txt = ctk.CTkFrame(row, fg_color="transparent")
            txt.pack(side="left", anchor="w")
            ctk.CTkLabel(txt, text=title, text_color=COLORS["text"], font=FONTS["heading"], anchor="w").pack(anchor="w")
            ctk.CTkLabel(txt, text=desc, text_color=COLORS["text_muted"], font=FONTS["tiny"], anchor="w").pack(anchor="w")
        var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(modal, text="لا تظهر هذه الرسالة مرة أخرى", variable=var, font=FONTS["tiny"]).pack(pady=10)
        def close():
            if var.get():
                try:
                    (pathlib.Path.home() / ".kpi_onboarded").write_text("1")
                except: pass
            modal.destroy()
        ctk.CTkButton(modal, text="ابدأ التقييم  →", fg_color=COLORS["primary"], hover_color=COLORS["primary_dark"], width=200, height=38, corner_radius=10, command=close).pack(pady=10)
        # Also close on Escape
        modal.bind("<Escape>", lambda e: close())

    def _update_onboarding_progress(self):
        # Count completed (score !=1)
        completed = sum(1 for d in domains for v in self.scores[d["domain"]].values() if v != 1)
        total = 58
        pct = completed / total
        try:
            self.lbl_progress.configure(text=f"{completed}/{total} مكتمل")
            self.progress_steps.set(pct)
            # Update 3 steps: 1 org, 2 eval, 3 export
            has_org = bool(self.org_var.get().strip())
            has_scores = completed > 0
            # Step 1: org
            c1, n1, t1 = self.step_labels[0]
            c1.configure(fg_color=COLORS["success"] if has_org else COLORS["primary"])
            n1.configure(text="✓" if has_org else "١", text_color="white")
            t1.configure(text_color=COLORS["success"] if has_org else COLORS["text"])
            # Step 2: eval
            c2, n2, t2 = self.step_labels[1]
            if completed == total:
                c2.configure(fg_color=COLORS["success"])
                n2.configure(text="✓", text_color="white")
                t2.configure(text_color=COLORS["success"])
            elif has_scores:
                c2.configure(fg_color=COLORS["primary"])
                n2.configure(text="٢", text_color="white")
                t2.configure(text_color=COLORS["text"])
            else:
                c2.configure(fg_color=COLORS["border"])
                n2.configure(text="٢", text_color=COLORS["text_muted"])
                t2.configure(text_color=COLORS["text_muted"])
            # Step 3: export (enabled when has_scores)
            c3, n3, t3 = self.step_labels[2]
            if has_scores:
                c3.configure(fg_color=COLORS["primary"])
                n3.configure(text="٣", text_color="white")
                t3.configure(text_color=COLORS["text"])
            else:
                c3.configure(fg_color=COLORS["border"])
                n3.configure(text="٣", text_color=COLORS["text_muted"])
                t3.configure(text_color=COLORS["text_muted"])
        except Exception as e:
            print(f"Progress update failed: {e}")

if __name__=="__main__":
    ctk.set_appearance_mode("light")
    root = ctk.CTk()
    app = KPIApp(root)
    root.mainloop()
