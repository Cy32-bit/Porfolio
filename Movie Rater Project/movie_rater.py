"""
Movie Rater
────────────────────────────────────────────────────────────────
Stack   : CustomTkinter · Pandas · Openpyxl · Pillow
Features: CRUD · Live Excel sync (mtime) · PyInstaller-safe
          resource paths · integer-only sliders · global
          mousewheel scrolling · Poppins typography
"""

import os
import sys
import time
import threading
from pathlib import Path

import customtkinter as ctk
from tkinter import messagebox, filedialog
import tkinter.ttk as ttk
import pandas as pd
import numpy as np

try:
    from PIL import Image
    PIL_OK = True
except ImportError:
    PIL_OK = False

# ═══════════════════════════════════════════════════════════════
#  RESOURCE PATH  –  works both in source and PyInstaller --onefile
# ═══════════════════════════════════════════════════════════════

def resource_path(relative: str) -> Path:
    """Return absolute path to a bundled resource, compatible with
    PyInstaller one-file mode (_MEIPASS) and normal Python execution."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    return base / relative

# Paths to bundle with PyInstaller (--add-data flags in setup.bat)
LOGO_PNG = resource_path("assets/mrl_noBg.png")
LOGO_ICO = resource_path("assets/mrl_noBg.ico")

# ═══════════════════════════════════════════════════════════════
#  APPEARANCE
# ═══════════════════════════════════════════════════════════════

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# ── Palette ────────────────────────────────────────────────────
GOLD      = "#F5A623"
GOLD_DIM  = "#C8871A"
BG        = "#0D1520"
CARD      = "#162030"
CARD2     = "#1A2840"
INPUT     = "#1C2B3A"
TXT       = "#FFFFFF"
MUTED     = "#8A9BB0"
BORDER    = "#243447"
RED       = "#E05252"
RED2      = "#B83C3C"
GREEN     = "#4CAF82"

# ── Schema ─────────────────────────────────────────────────────
CATS    = ["Plot", "Cinematography", "Acting", "Direction", "Pacing"]
COLUMNS = ["Title", "Year", "Synopsis"] + CATS + ["Overall_Rating"]

CAT_ICON = {
    "Plot":           "📖",
    "Cinematography": "🎥",
    "Acting":         "🎭",
    "Direction":      "🎬",
    "Pacing":         "⚡",
}
CAT_DESC = {
    "Plot":           "Story structure & narrative quality",
    "Cinematography": "Visual storytelling & camera work",
    "Acting":         "Performance & character portrayal",
    "Direction":      "Directorial vision & execution",
    "Pacing":         "Rhythm, flow & scene timing",
}

DEFAULT_DB = Path.home() / "Documents" / "Movie Rating DB.xlsx"

# ═══════════════════════════════════════════════════════════════
#  FONT HELPER
# ═══════════════════════════════════════════════════════════════

def F(size: int, bold: bool = False) -> ctk.CTkFont:
    return ctk.CTkFont(
        family="Poppins",
        size=size,
        weight="bold" if bold else "normal",
    )

# ═══════════════════════════════════════════════════════════════
#  DB HELPERS
# ═══════════════════════════════════════════════════════════════

def db_load(path: Path) -> pd.DataFrame:
    try:
        if not path.exists():
            return pd.DataFrame(columns=COLUMNS)
        if path.suffix == ".csv":
            df = pd.read_csv(path)
        elif path.suffix == ".json":
            df = pd.read_json(path)
        else:
            df = pd.read_excel(path, engine="openpyxl")
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = "" if col in ("Title", "Synopsis", "Year") else 0
        return df[COLUMNS].reset_index(drop=True)
    except Exception as exc:
        messagebox.showerror("Load Error", str(exc))
        return pd.DataFrame(columns=COLUMNS)


def db_save(df: pd.DataFrame, path: Path) -> bool:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.suffix == ".csv":
            df.to_csv(path, index=False)
        elif path.suffix == ".json":
            df.to_json(path, orient="records", indent=2)
        else:
            df.to_excel(path, index=False, engine="openpyxl")
        return True
    except Exception as exc:
        messagebox.showerror("Save Error", str(exc))
        return False


def mtime(path: Path) -> float:
    try:
        return path.stat().st_mtime
    except Exception:
        return 0.0


# ═══════════════════════════════════════════════════════════════
#  GLOBAL MOUSEWHEEL MANAGER
#  Attach once; dispatches events to whichever canvas is under
#  the pointer so scrolling works without clicking first.
# ═══════════════════════════════════════════════════════════════

class WheelManager:
    """Register/unregister scrollable canvases for global wheel dispatch."""

    def __init__(self, root: ctk.CTk):
        self._canvases: list = []
        root.bind_all("<MouseWheel>", self._dispatch, add="+")     # Windows
        root.bind_all("<Button-4>",   self._dispatch, add="+")     # Linux scroll up
        root.bind_all("<Button-5>",   self._dispatch, add="+")     # Linux scroll down

    def register(self, canvas) -> None:
        if canvas not in self._canvases:
            self._canvases.append(canvas)

    def unregister(self, canvas) -> None:
        self._canvases = [c for c in self._canvases if c != canvas]

    def _dispatch(self, event) -> None:
        # Determine scroll delta (Windows vs Linux)
        if event.num == 4:
            delta = 3
        elif event.num == 5:
            delta = -3
        else:
            delta = int(-1 * (event.delta / 120)) * 3   # 3× faster than default

        # Scroll the canvas that contains the widget under the pointer
        widget = event.widget
        for canvas in self._canvases:
            try:
                if self._is_child(widget, canvas):
                    canvas.yview_scroll(delta, "units")
                    return
            except Exception:
                continue
        # Fallback: scroll the topmost registered canvas
        if self._canvases:
            try:
                self._canvases[-1].yview_scroll(delta, "units")
            except Exception:
                pass

    @staticmethod
    def _is_child(widget, canvas) -> bool:
        """Return True if widget is canvas or a descendant of it."""
        try:
            w = widget
            while w:
                if str(w) == str(canvas):
                    return True
                w = w.master
        except Exception:
            pass
        return False


# ═══════════════════════════════════════════════════════════════
#  MAIN WINDOW
# ═══════════════════════════════════════════════════════════════

class MovieRater(ctk.CTk):

    def __init__(self):
        super().__init__()

        # ── Window setup ──────────────────────────────────────
        self.title("Movie Rater")
        self.geometry("900x600")
        self.minsize(820, 540)
        self.configure(fg_color=BG)
        self._set_icon(self)

        # ── State ─────────────────────────────────────────────
        self.db_path     = DEFAULT_DB
        self.df          = db_load(self.db_path)
        self._last_mtime = mtime(self.db_path)
        self._saving     = False          # suppress watcher during own writes

        # ── Global scroll manager ─────────────────────────────
        self.wheel = WheelManager(self)

        # ── Build UI ──────────────────────────────────────────
        self._build_header()
        self._build_tabview()

        # ── Live-sync watcher ─────────────────────────────────
        threading.Thread(target=self._watch_loop, daemon=True).start()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Icon helper ───────────────────────────────────────────

    def _set_icon(self, window):
        if LOGO_ICO.exists():
            try:
                window.iconbitmap(str(LOGO_ICO))
                return
            except Exception:
                pass

    # ── Shared card factory ───────────────────────────────────

    def card(self, parent, **kw) -> ctk.CTkFrame:
        defaults = dict(fg_color=CARD, corner_radius=14,
                        border_width=1, border_color=BORDER)
        defaults.update(kw)
        return ctk.CTkFrame(parent, **defaults)

    # ─────────────────────────────────────────────────────────
    #  HEADER
    # ─────────────────────────────────────────────────────────

    def _build_header(self):
        bar = ctk.CTkFrame(self, fg_color=CARD, corner_radius=0, height=82)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        wrap = ctk.CTkFrame(bar, fg_color="transparent")
        wrap.pack(fill="both", expand=True, padx=28, pady=10)

        # Left — logo + title
        left = ctk.CTkFrame(wrap, fg_color="transparent")
        left.pack(side="left", fill="y")

        logo_shown = False
        if PIL_OK and LOGO_PNG.exists():
            try:
                img = Image.open(LOGO_PNG).convert("RGBA")
                img.thumbnail((54, 54), Image.LANCZOS)
                ctk_img = ctk.CTkImage(light_image=img,
                                       dark_image=img, size=(54, 54))
                ctk.CTkLabel(left, image=ctk_img, text="",
                             width=54).pack(side="left", padx=(0, 14))
                logo_shown = True
            except Exception:
                pass

        if not logo_shown:
            ctk.CTkLabel(left, text="🎬", font=F(32),
                         text_color=GOLD).pack(side="left", padx=(0, 12))

        title_col = ctk.CTkFrame(left, fg_color="transparent")
        title_col.pack(side="left", fill="y", pady=6)
        ctk.CTkLabel(title_col, text="Movie Rater",
                     font=F(22, True), text_color=GOLD).pack(anchor="w")
        ctk.CTkLabel(title_col, text="Rate your cinema experience",
                     font=F(12), text_color=MUTED).pack(anchor="w")

        # Right — count + action buttons
        right = ctk.CTkFrame(wrap, fg_color="transparent")
        right.pack(side="right", fill="y")

        self.count_lbl = ctk.CTkLabel(right,
                                       text=self._count_text(),
                                       font=F(12), text_color=TXT)
        self.count_lbl.pack(side="left", padx=(0, 20))

        for label, cmd in [
            ("🗄   Database", self._change_db),
            ("💾   Save As",  self._save_as),
        ]:
            ctk.CTkButton(
                right, text=label, font=F(11), height=38, width=128,
                fg_color=INPUT, hover_color=CARD2,
                text_color=MUTED, corner_radius=10,
                command=cmd,
            ).pack(side="left", padx=(0, 6))

        ctk.CTkFrame(self, fg_color=BORDER, height=1,
                     corner_radius=0).pack(fill="x")

    # ─────────────────────────────────────────────────────────
    #  TABVIEW
    # ─────────────────────────────────────────────────────────

    def _build_tabview(self):
        self.tabs = ctk.CTkTabview(
            self,
            fg_color=BG,
            segmented_button_fg_color=CARD,
            segmented_button_selected_color=GOLD,
            segmented_button_selected_hover_color=GOLD_DIM,
            segmented_button_unselected_color=CARD,
            segmented_button_unselected_hover_color=INPUT,
            text_color=TXT,
            corner_radius=14,
            command=self._on_tab_switch,
        )
        self.tabs.pack(fill="both", expand=True, padx=18, pady=(14, 16))

        self.T_RATE = "  ☆  Rate Movie  "
        self.T_LIB  = "  ⊞  My Library  "
        self.T_STAT = "  ⊿  Stats  "

        for label in (self.T_RATE, self.T_LIB, self.T_STAT):
            self.tabs.add(label)

        self._build_rate_tab(self.tabs.tab(self.T_RATE))
        self._build_library_tab(self.tabs.tab(self.T_LIB))
        self._build_stats_tab(self.tabs.tab(self.T_STAT))

    def _on_tab_switch(self):
        t = self.tabs.get()
        if t == self.T_LIB:
            self._refresh_library()
        elif t == self.T_STAT:
            self._refresh_stats()

    # ═══════════════════════════════════════════════════════════
    #  RATE TAB
    # ═══════════════════════════════════════════════════════════

    def _build_rate_tab(self, parent):
        parent.configure(fg_color=BG)

        sf = ctk.CTkScrollableFrame(
            parent, fg_color=BG, corner_radius=0,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=MUTED,
        )
        sf.pack(fill="both", expand=True)
        # Register inner canvas for global wheel dispatch
        self.wheel.register(sf._parent_canvas)

        # ── Movie info ────────────────────────────────────────
        c = self.card(sf)
        c.pack(fill="x", pady=(0, 10))
        ci = ctk.CTkFrame(c, fg_color="transparent")
        ci.pack(fill="x", padx=20, pady=18)

        ctk.CTkLabel(ci, text="Rate a Movie",
                     font=F(16, True), text_color=TXT).grid(
                     row=0, column=0, columnspan=4, sticky="w", pady=(0, 14))

        for col_idx, txt in enumerate(("Movie Title", "Release Year")):
            ctk.CTkLabel(ci, text=txt, font=F(11),
                         text_color=MUTED).grid(
                         row=1, column=col_idx * 2, sticky="w",
                         padx=(20 * col_idx, 0))

        self.v_title = ctk.StringVar()
        self.v_year  = ctk.StringVar(value="2024")

        ctk.CTkEntry(
            ci, textvariable=self.v_title,
            placeholder_text="Enter movie name...",
            font=F(12), height=42,
            fg_color=INPUT, border_color=BORDER,
            text_color=TXT, corner_radius=10,
        ).grid(row=2, column=0, columnspan=2, sticky="ew", pady=(5, 0))

        ctk.CTkEntry(
            ci, textvariable=self.v_year,
            font=F(12), height=42, width=160,
            fg_color=INPUT, border_color=BORDER,
            text_color=TXT, corner_radius=10,
        ).grid(row=2, column=2, columnspan=2, sticky="ew",
               pady=(5, 0), padx=(20, 0))

        ci.columnconfigure(0, weight=3)
        ci.columnconfigure(2, weight=1)

        # ── Synopsis ──────────────────────────────────────────
        sc = self.card(sf)
        sc.pack(fill="x", pady=(0, 10))
        si = ctk.CTkFrame(sc, fg_color="transparent")
        si.pack(fill="x", padx=20, pady=16)
        ctk.CTkLabel(si, text="Synopsis  (optional)",
                     font=F(11), text_color=MUTED).pack(anchor="w", pady=(0, 6))
        self.synopsis = ctk.CTkTextbox(
            si, height=78, font=F(11),
            fg_color=INPUT, border_color=BORDER,
            text_color=TXT, corner_radius=10,
            scrollbar_button_color=BORDER,
        )
        self.synopsis.pack(fill="x")

        # ── Category header ───────────────────────────────────
        hc = self.card(sf)
        hc.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(hc, text="★   Rate Each Category  ( 1 – 10, integers )",
                     font=F(13, True), text_color=GOLD).pack(
                     anchor="w", padx=20, pady=13)

        # ── Sliders (integer-only) ────────────────────────────
        self.cat_int_vars: dict[str, ctk.IntVar]   = {}
        self.cat_val_lbls: dict[str, ctk.CTkLabel] = {}
        for cat in CATS:
            self._make_slider_card(sf, cat)

        # ── Overall (float average) ───────────────────────────
        ov = ctk.CTkFrame(sf, fg_color="#192B14",
                          corner_radius=14,
                          border_width=1, border_color="#2D4A1E")
        ov.pack(fill="x", pady=(0, 12))
        ovi = ctk.CTkFrame(ov, fg_color="transparent")
        ovi.pack(fill="x", padx=20, pady=16)

        top = ctk.CTkFrame(ovi, fg_color="transparent")
        top.pack(fill="x")
        ctk.CTkLabel(top, text="★  Overall Rating",
                     font=F(14, True), text_color=TXT).pack(side="left")
        self.overall_lbl = ctk.CTkLabel(top, text="0.0 / 10",
                                         font=F(24, True), text_color=GOLD)
        self.overall_lbl.pack(side="right")
        ctk.CTkLabel(ovi, text="Auto-calculated mean of all category scores",
                     font=F(10), text_color=MUTED).pack(anchor="w", pady=(3, 0))

        # ── Action buttons ────────────────────────────────────
        br = ctk.CTkFrame(sf, fg_color="transparent")
        br.pack(fill="x", pady=(4, 24))

        ctk.CTkButton(
            br, text="✓   Save Rating",
            font=F(13, True), height=48, corner_radius=12,
            fg_color=GOLD, hover_color=GOLD_DIM, text_color="#000000",
            command=self._save_rating,
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkButton(
            br, text="↺   Clear Form",
            font=F(13), height=48, corner_radius=12,
            fg_color=CARD, hover_color=INPUT, text_color=MUTED,
            command=self._clear_form,
        ).pack(side="left")

    # ── Slider card (integer 1–10) ────────────────────────────

    def _make_slider_card(self, parent, cat: str):
        c = self.card(parent)
        c.pack(fill="x", pady=(0, 8))
        ci = ctk.CTkFrame(c, fg_color="transparent")
        ci.pack(fill="x", padx=20, pady=14)

        top = ctk.CTkFrame(ci, fg_color="transparent")
        top.pack(fill="x")
        ctk.CTkLabel(top, text=f"{CAT_ICON[cat]}   {cat}",
                     font=F(12, True), text_color=TXT).pack(side="left")

        val_lbl = ctk.CTkLabel(top, text="0",
                                font=F(20, True), text_color=GOLD)
        val_lbl.pack(side="right")
        self.cat_val_lbls[cat] = val_lbl

        ctk.CTkLabel(ci, text=CAT_DESC[cat],
                     font=F(10), text_color=MUTED).pack(anchor="w", pady=(2, 8))

        int_var = ctk.IntVar(value=0)
        self.cat_int_vars[cat] = int_var

        def _snap(val, _v=int_var, _lbl=val_lbl):
            """Snap float slider value to nearest integer."""
            snapped = round(float(val))
            _v.set(snapped)
            _lbl.configure(text=str(snapped))
            self._update_overall()

        ctk.CTkSlider(
            ci,
            from_=0, to=10,
            number_of_steps=10,       # forces integer positions
            variable=int_var,
            command=_snap,
            button_color=GOLD,
            button_hover_color=GOLD_DIM,
            progress_color=GOLD,
            fg_color=INPUT,
            height=18,
        ).pack(fill="x", pady=(0, 4))

        foot = ctk.CTkFrame(ci, fg_color="transparent")
        foot.pack(fill="x")
        ctk.CTkLabel(foot, text="0 – Poor",
                     font=F(9), text_color=MUTED).pack(side="left")
        ctk.CTkLabel(foot, text="10 – Outstanding",
                     font=F(9), text_color=MUTED).pack(side="right")

    def _update_overall(self):
        avg = sum(v.get() for v in self.cat_int_vars.values()) / len(CATS)
        self.overall_lbl.configure(text=f"{avg:.1f} / 10")

    def _save_rating(self):
        title = self.v_title.get().strip()
        if not title:
            messagebox.showwarning("Missing Title", "Please enter a movie title.")
            return

        try:
            year = int(self.v_year.get().strip()) if self.v_year.get().strip() else ""
        except ValueError:
            messagebox.showwarning("Invalid Year", "Year must be a whole number.")
            return

        synopsis = self.synopsis.get("1.0", "end").strip()
        ratings  = {cat: int(self.cat_int_vars[cat].get()) for cat in CATS}
        overall  = round(sum(ratings.values()) / len(CATS), 2)

        # Duplicate guard
        if not self.df.empty:
            dupes = self.df["Title"].str.strip().str.lower() == title.lower()
            if dupes.any():
                if not messagebox.askyesno(
                        "Duplicate Entry",
                        f'"{title}" already exists.\nAdd a second entry anyway?'):
                    return

        new_row = {"Title": title, "Year": year, "Synopsis": synopsis,
                   **ratings, "Overall_Rating": overall}
        self.df = pd.concat(
            [self.df, pd.DataFrame([new_row])], ignore_index=True)

        self._write_db()
        self._refresh_count()
        messagebox.showinfo("Saved",
                            f'"{title}" saved.\nOverall Rating: {overall} / 10')
        self._clear_form()

    def _clear_form(self):
        """Reset every control to its default zero state."""
        self.v_title.set("")
        self.v_year.set("2024")
        self.synopsis.delete("1.0", "end")
        for cat in CATS:
            self.cat_int_vars[cat].set(0)
            self.cat_val_lbls[cat].configure(text="0")
        self.overall_lbl.configure(text="0.0 / 10")

    # ═══════════════════════════════════════════════════════════
    #  LIBRARY TAB
    # ═══════════════════════════════════════════════════════════

    def _build_library_tab(self, parent):
        parent.configure(fg_color=BG)

        # Header row
        top = ctk.CTkFrame(parent, fg_color="transparent")
        top.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(top, text="My Rated Movies",
                     font=F(15, True), text_color=TXT).pack(side="left")

        ctk.CTkLabel(top, text="Sort by:", font=F(11),
                     text_color=MUTED).pack(side="right", padx=(0, 6))
        self.v_sort = ctk.StringVar(value="Overall_Rating")
        ctk.CTkComboBox(
            top,
            values=["Title", "Year", "Overall_Rating"] + CATS,
            variable=self.v_sort, width=180, height=36, font=F(11),
            fg_color=INPUT, border_color=BORDER,
            button_color=BORDER, button_hover_color=CARD2,
            dropdown_fg_color=CARD, dropdown_hover_color=INPUT,
            text_color=TXT, corner_radius=8,
            command=lambda _: self._refresh_library(),
        ).pack(side="right")

        # Search
        self.v_search = ctk.StringVar()
        self.v_search.trace_add("write", lambda *_: self._refresh_library())
        ctk.CTkEntry(
            parent, textvariable=self.v_search,
            placeholder_text="🔍   Search by title...",
            font=F(11), height=40,
            fg_color=INPUT, border_color=BORDER,
            text_color=TXT, corner_radius=10,
        ).pack(fill="x", pady=(0, 8))

        # Treeview
        tree_wrap = self.card(parent)
        tree_wrap.pack(fill="both", expand=True, pady=(0, 8))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("MR.Treeview",
                        background=CARD,
                        foreground=TXT,
                        fieldbackground=CARD,
                        rowheight=30,
                        font=("Poppins", 10))
        style.configure("MR.Treeview.Heading",
                        background=INPUT,
                        foreground=GOLD,
                        font=("Poppins", 10, "bold"),
                        relief="flat")
        style.map("MR.Treeview",
                  background=[("selected", "#243A5A")],
                  foreground=[("selected", TXT)])

        cols = ("Title", "Year") + tuple(CATS) + ("Overall",)
        self.tree = ttk.Treeview(tree_wrap, columns=cols,
                                  show="headings", selectmode="browse",
                                  style="MR.Treeview")
        vsb = ttk.Scrollbar(tree_wrap, orient="vertical",
                             command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y", padx=(0, 4), pady=4)
        self.tree.pack(fill="both", expand=True, padx=4, pady=4)

        # Mousewheel on treeview
        self.tree.bind("<MouseWheel>",
                       lambda e: self.tree.yview_scroll(
                           int(-1 * (e.delta / 120)) * 3, "units"))

        widths = {"Title": 180, "Year": 55, "Overall": 70}
        for col in cols:
            self.tree.heading(col, text=col, anchor="center")
            self.tree.column(col,
                             width=widths.get(col, 88), anchor="center")
        self.tree.column("Title", anchor="w")

        self.tree.tag_configure("even", background=CARD)
        self.tree.tag_configure("odd",  background="#1A2A3A")
        self.tree.bind("<Double-1>", lambda _: self._edit_selected())

        # Action row
        ar = ctk.CTkFrame(parent, fg_color="transparent")
        ar.pack(fill="x")

        ctk.CTkButton(
            ar, text="✎   Edit Selected",
            font=F(11), height=40, corner_radius=10, width=160,
            fg_color=CARD, hover_color=INPUT, text_color=TXT,
            command=self._edit_selected,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            ar, text="✕   Delete Selected",
            font=F(11), height=40, corner_radius=10, width=168,
            fg_color=RED, hover_color=RED2, text_color=TXT,
            command=self._delete_selected,
        ).pack(side="left")

        self._refresh_library()

    def _refresh_library(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        df = self.df.copy()
        q  = self.v_search.get().strip().lower()
        if q:
            df = df[df["Title"].str.lower().str.contains(q, na=False)]

        col = self.v_sort.get()
        if col in df.columns:
            df = df.sort_values(col, ascending=(col == "Title"))

        for i, (_, r) in enumerate(df.iterrows()):
            tag = "even" if i % 2 == 0 else "odd"
            row_vals = [r.get("Title", ""), r.get("Year", "")]
            for cat in CATS:
                v = r.get(cat, 0)
                try:
                    row_vals.append(int(float(v)))
                except (ValueError, TypeError):
                    row_vals.append("")
            ov = r.get("Overall_Rating", 0)
            try:
                row_vals.append(f"{float(ov):.2f}")
            except (ValueError, TypeError):
                row_vals.append("")
            self.tree.insert("", "end", values=row_vals, tags=(tag,))

    def _selected_title(self) -> str | None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("No Selection",
                                "Select a movie from the list first.")
            return None
        return str(self.tree.item(sel[0])["values"][0])

    # ── Edit selected ─────────────────────────────────────────

    def _edit_selected(self):
        title = self._selected_title()
        if title is None:
            return

        # Find the FIRST matching row index in self.df
        mask = self.df["Title"].astype(str).str.strip() == title.strip()
        if not mask.any():
            messagebox.showwarning("Not Found",
                                   f'"{title}" not found in database.')
            return

        df_idx = int(self.df.index[mask][0])
        row    = self.df.loc[df_idx]
        EditDialog(self, row=row, df_idx=df_idx, on_save=self._apply_edit)

    def _apply_edit(self, df_idx: int, updated: dict):
        """
        Write the updated dict back into self.df at the exact row,
        save to disk, and immediately refresh the UI.
        This ensures no duplicate is created — only the targeted row changes.
        """
        try:
            for col, val in updated.items():
                # If the target column is numeric, coerce the incoming value
                # to a numeric type (numbers -> numbers, 'nan' or invalid -> NaN).
                if col in self.df.columns and pd.api.types.is_numeric_dtype(self.df[col].dtype):
                    coerced = pd.to_numeric(val, errors="coerce")
                    # assign the coerced numeric (may be np.nan) safely
                    self.df.at[df_idx, col] = coerced
                else:
                    self.df.at[df_idx, col] = val

            self._write_db()
            self._refresh_count()
            self._refresh_library()
            messagebox.showinfo("Saved", "Changes saved.")
        except Exception as exc:
            messagebox.showerror("Save Error", str(exc))

    # ── Delete selected ───────────────────────────────────────

    def _delete_selected(self):
        title = self._selected_title()
        if title is None:
            return
        if messagebox.askyesno("Delete Movie",
                f'Permanently delete "{title}"?'):
            self.df = self.df[
                self.df["Title"].astype(str).str.strip() != title.strip()
            ].reset_index(drop=True)
            self._write_db()
            self._refresh_count()
            self._refresh_library()

    # ═══════════════════════════════════════════════════════════
    #  STATS TAB
    # ═══════════════════════════════════════════════════════════

    def _build_stats_tab(self, parent):
        parent.configure(fg_color=BG)

        sf = ctk.CTkScrollableFrame(
            parent, fg_color=BG, corner_radius=0,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=MUTED,
        )
        sf.pack(fill="both", expand=True)
        self.wheel.register(sf._parent_canvas)

        ctk.CTkLabel(sf, text="Your Rating Statistics",
                     font=F(15, True), text_color=TXT).pack(
                     anchor="w", pady=(4, 14))

        # 2×2 summary grid
        grid = ctk.CTkFrame(sf, fg_color="transparent")
        grid.pack(fill="x", pady=(0, 14))
        grid.columnconfigure((0, 1), weight=1)

        self._stat_lbls: dict[str, ctk.CTkLabel] = {}
        stat_meta = [
            ("movies", "🎬", "Movies Rated",   "0",  GOLD,  0, 0),
            ("avg",    "★",  "Average Rating", "–",  GOLD,  0, 1),
            ("top",    "🏆", "Rated 9–10",      "0",  GREEN, 1, 0),
            ("low",    "👎", "Rated 1–3",       "0",  RED,   1, 1),
        ]
        for key, icon, label, default, color, grow, gcol in stat_meta:
            fc = self.card(grid)
            fc.grid(row=grow, column=gcol, sticky="nsew",
                    padx=(0, 8) if gcol == 0 else (8, 0),
                    pady=(0, 8))
            fi = ctk.CTkFrame(fc, fg_color="transparent")
            fi.pack(fill="both", padx=18, pady=16)
            ctk.CTkLabel(fi, text=icon, font=F(28),
                         text_color=color).pack(anchor="w")
            n = ctk.CTkLabel(fi, text=default,
                             font=F(32, True), text_color=color)
            n.pack(anchor="w")
            ctk.CTkLabel(fi, text=label, font=F(10),
                         text_color=MUTED).pack(anchor="w")
            self._stat_lbls[key] = n

        # Category progress bars
        bc = self.card(sf)
        bc.pack(fill="x", pady=(0, 14))
        bi = ctk.CTkFrame(bc, fg_color="transparent")
        bi.pack(fill="x", padx=20, pady=18)
        ctk.CTkLabel(bi, text="Average by Category",
                     font=F(13, True), text_color=TXT).pack(
                     anchor="w", pady=(0, 12))

        self._cat_bars: dict[str, ctk.CTkProgressBar] = {}
        self._cat_bar_lbls: dict[str, ctk.CTkLabel]   = {}
        for cat in CATS:
            row = ctk.CTkFrame(bi, fg_color="transparent")
            row.pack(fill="x", pady=5)
            ctk.CTkLabel(row,
                         text=f"{CAT_ICON[cat]}  {cat}",
                         font=F(11), text_color=TXT,
                         width=170, anchor="w").pack(side="left")
            bar = ctk.CTkProgressBar(row, height=14, corner_radius=7,
                                      fg_color=INPUT, progress_color=GOLD)
            bar.set(0)
            bar.pack(side="left", fill="x", expand=True, padx=10)
            lbl = ctk.CTkLabel(row, text="–",
                                font=F(11, True), text_color=GOLD,
                                width=42, anchor="w")
            lbl.pack(side="left")
            self._cat_bars[cat]     = bar
            self._cat_bar_lbls[cat] = lbl

    def _refresh_stats(self):
        n = len(self.df)
        self._stat_lbls["movies"].configure(text=str(n))

        if n == 0:
            self._stat_lbls["avg"].configure(text="–")
            self._stat_lbls["top"].configure(text="0")
            self._stat_lbls["low"].configure(text="0")
            for cat in CATS:
                self._cat_bars[cat].set(0)
                self._cat_bar_lbls[cat].configure(text="–")
            return

        ov = pd.to_numeric(self.df["Overall_Rating"], errors="coerce")
        self._stat_lbls["avg"].configure(text=f"{ov.mean():.1f}")
        self._stat_lbls["top"].configure(text=str(int((ov >= 9).sum())))
        self._stat_lbls["low"].configure(text=str(int((ov <= 3).sum())))

        for cat in CATS:
            if cat in self.df.columns:
                ca = pd.to_numeric(self.df[cat], errors="coerce").mean()
                if pd.notna(ca):
                    self._cat_bars[cat].set(ca / 10)
                    self._cat_bar_lbls[cat].configure(text=f"{ca:.1f}")

    # ═══════════════════════════════════════════════════════════
    #  DB HELPERS
    # ═══════════════════════════════════════════════════════════

    def _count_text(self) -> str:
        return f"{len(self.df)} movies rated"

    def _refresh_count(self):
        self.count_lbl.configure(text=self._count_text())

    def _write_db(self):
        """Save df to disk and update mtime so the watcher doesn't re-trigger."""
        self._saving = True
        db_save(self.df, self.db_path)
        self._last_mtime = mtime(self.db_path)
        self._saving = False

    def _change_db(self):
        path = filedialog.askopenfilename(
            title="Open Movie Database",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"),
                       ("JSON", "*.json"), ("All", "*.*")])
        if path:
            self.db_path     = Path(path)
            self.df          = db_load(self.db_path)
            self._last_mtime = mtime(self.db_path)
            self._refresh_count()
            self._refresh_library()
            messagebox.showinfo("Database Loaded", f"Loaded:\n{path}")

    def _save_as(self):
        path = filedialog.asksaveasfilename(
            title="Save Database As",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"),
                       ("JSON", "*.json")])
        if path:
            db_save(self.df, Path(path))
            messagebox.showinfo("Saved", f"Exported to:\n{path}")

    # ═══════════════════════════════════════════════════════════
    #  LIVE EXCEL WATCHER  (mtime polling, every 2 s)
    # ═══════════════════════════════════════════════════════════

    def _watch_loop(self):
        while True:
            time.sleep(2)
            if self._saving:
                continue
            try:
                mt = mtime(self.db_path)
                if mt > 0 and mt != self._last_mtime:
                    self._last_mtime = mt
                    self.after(0, self._on_external_change)
            except Exception:
                pass

    def _on_external_change(self):
        fresh = db_load(self.db_path)
        try:
            same = fresh.equals(self.df)
        except Exception:
            same = False
        if not same:
            self.df = fresh
            self._refresh_count()
            t = self.tabs.get()
            if t == self.T_LIB:
                self._refresh_library()
            elif t == self.T_STAT:
                self._refresh_stats()
            orig = self.title()
            self.title("Movie Rater  —  ↺  Database refreshed from disk")
            self.after(2800, lambda: self.title(orig))

    def _on_close(self):
        self._saving = True      # stop watcher thread reacting
        self.destroy()


# ═══════════════════════════════════════════════════════════════
#  EDIT DIALOG
# ═══════════════════════════════════════════════════════════════

class EditDialog(ctk.CTkToplevel):

    def __init__(self, parent: MovieRater, row, df_idx: int, on_save):
        super().__init__(parent)

        self.df_idx  = df_idx
        self.on_save = on_save
        title_str    = str(row.get("Title", ""))

        self.title(f"Edit — {title_str}")
        self.geometry("560x590")
        self.resizable(False, False)
        self.configure(fg_color=BG)
        self.grab_set()
        self.lift()
        self.focus_force()
        parent._set_icon(self)

        ctk.CTkLabel(self, text=f"Editing: {title_str}",
                     font=F(15, True), text_color=TXT).pack(
                     anchor="w", padx=24, pady=(20, 2))
        ctk.CTkLabel(self,
                     text="Adjust ratings or synopsis, then save.",
                     font=F(10), text_color=MUTED).pack(
                     anchor="w", padx=24, pady=(0, 10))

        sf = ctk.CTkScrollableFrame(self, fg_color=BG, corner_radius=0,
                                     scrollbar_button_color=BORDER)
        sf.pack(fill="both", expand=True, padx=20)

        # Synopsis
        ctk.CTkLabel(sf, text="Synopsis",
                     font=F(10), text_color=MUTED).pack(
                     anchor="w", pady=(0, 4))
        self.syn = ctk.CTkTextbox(sf, height=68, font=F(11),
                                   fg_color=INPUT, border_color=BORDER,
                                   text_color=TXT, corner_radius=10)
        self.syn.pack(fill="x", pady=(0, 14))
        self.syn.insert("1.0", str(row.get("Synopsis", "") or ""))

        # Sliders — integer-only, pre-seeded with existing values
        self._edit_int_vars: dict[str, ctk.IntVar]   = {}
        self._edit_val_lbls: dict[str, ctk.CTkLabel] = {}

        for cat in CATS:
            current = 0
            try:
                current = int(round(float(row.get(cat, 0) or 0)))
            except (ValueError, TypeError):
                pass

            fc = ctk.CTkFrame(sf, fg_color=CARD, corner_radius=12,
                              border_width=1, border_color=BORDER)
            fc.pack(fill="x", pady=(0, 8))
            fi = ctk.CTkFrame(fc, fg_color="transparent")
            fi.pack(fill="x", padx=16, pady=12)

            top = ctk.CTkFrame(fi, fg_color="transparent")
            top.pack(fill="x")
            ctk.CTkLabel(top, text=f"{CAT_ICON[cat]}   {cat}",
                         font=F(11, True), text_color=TXT).pack(side="left")
            lbl = ctk.CTkLabel(top, text=str(current),
                                font=F(16, True), text_color=GOLD)
            lbl.pack(side="right")
            self._edit_val_lbls[cat] = lbl

            var = ctk.IntVar(value=current)
            self._edit_int_vars[cat] = var

            def _snap(val, _v=var, _lbl=lbl):
                s = round(float(val))
                _v.set(s)
                _lbl.configure(text=str(s))

            ctk.CTkSlider(
                fi,
                from_=0, to=10,
                number_of_steps=10,
                variable=var,
                command=_snap,
                button_color=GOLD,
                button_hover_color=GOLD_DIM,
                progress_color=GOLD,
                fg_color=INPUT,
                height=16,
            ).pack(fill="x", pady=(8, 0))

        # Buttons
        br = ctk.CTkFrame(self, fg_color="transparent")
        br.pack(fill="x", padx=20, pady=14)

        ctk.CTkButton(
            br, text="✓   Save Changes",
            font=F(12, True), height=44, corner_radius=11,
            fg_color=GOLD, hover_color=GOLD_DIM, text_color="#000000",
            command=self._commit,
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkButton(
            br, text="Cancel",
            font=F(12), height=44, corner_radius=11,
            fg_color=CARD, hover_color=INPUT, text_color=MUTED,
            command=self.destroy,
        ).pack(side="left")

    def _commit(self):
        synopsis = self.syn.get("1.0", "end").strip()
        updated  = {"Synopsis": synopsis}
        for cat in CATS:
            updated[cat] = int(self._edit_int_vars[cat].get())
        updated["Overall_Rating"] = round(
            sum(updated[c] for c in CATS) / len(CATS), 2)

        self.on_save(self.df_idx, updated)
        self.destroy()


# ═══════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = MovieRater()
    app.mainloop()