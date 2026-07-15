"""Advisee Document Filler, standalone desktop app.

A step-by-step wizard. No browser, no web server, no network at all.
Works on Windows and Linux with the same code.

Developed by Yaksharo a.k.a Ezer
"""
import os
import queue
import sys
import threading
import traceback
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

# Heavy libraries (pdfplumber, python-docx, docxcompose) are loaded
# lazily on first use so the window appears immediately.
grade_parser = None
filler = None


def _load_engine():
    global grade_parser, filler
    if grade_parser is None:
        import parser as _gp
        import filler as _fl
        grade_parser = _gp
        filler = _fl


def resource_path(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


APP_TITLE = "Advisee Document Filler"
APP_VERSION = "1.0"
DEVELOPER = "Yaksharo a.k.a Ezer"

try:
    import darkdetect
except ImportError:
    darkdetect = None

THEMES = {
    "light": {
        "bg": "#f6f8f6", "card": "#ffffff", "fg": "#1f2937",
        "muted": "#6b7280", "line": "#d6dbd6",
        "accent": "#14532d", "accent_fg": "#ffffff",
        "accent_hover": "#1c6b3a", "field": "#ffffff",
        "header": "#14532d", "header_fg": "#ffffff",
    },
    "dark": {
        "bg": "#161b18", "card": "#1f2621", "fg": "#e5e9e6",
        "muted": "#9aa39c", "line": "#3a423c",
        "accent": "#2f9e5b", "accent_fg": "#0b0f0c",
        "accent_hover": "#3cb96e", "field": "#262e28",
        "header": "#0f1512", "header_fg": "#e5e9e6",
    },
}


def system_theme():
    if darkdetect is not None:
        try:
            return "dark" if darkdetect.isDark() else "light"
        except Exception:
            pass
    return "light"


def safe_name(name):
    import re
    return re.sub(r"[^A-Za-z0-9 _.\-]", "", name).strip().replace(" ", "_")


def clean_dir_path(raw):
    p = (raw or "").strip().strip('"').strip("'")
    p = os.path.expanduser(os.path.expandvars(p))
    return os.path.normpath(p) if p else ""


class Wizard(tk.Tk):
    STEPS = ["PDF Files", "Students", "Documents", "Faculty", "Generate"]

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self._center_window(920, 680)
        self.minsize(800, 600)

        # state gathered across steps
        self.pdf_paths = []
        self.merged = {}
        self.term_keys = []
        self.student_vars = {}
        self.term_vars = {}
        self.doc_appraisal = tk.BooleanVar(value=True)
        self.doc_report = tk.BooleanVar(value=True)
        self.output_mode = tk.StringVar(value="individual")
        self.trim_rows = tk.BooleanVar(value=True)
        self.adviser = tk.StringVar(value="JUAN M. DELA CRUZ")
        self.faculty_entries = {}
        self.out_dir = tk.StringVar(value=os.path.join(
            os.path.expanduser("~"), "AdviseeDocuments"))

        self.step = 0
        self.busy = False
        self._msgq = queue.Queue()   # thread-safe UI updates
        self._poll_queue()
        # Custom themed titlebar on Windows (borderless with min/max/close).
        # On Linux we keep the native decorations so the titlebar always
        # matches the desktop environment (GNOME, KDE, XFCE...) exactly.
        self._custom_titlebar = sys.platform.startswith("win")
        self._is_zoomed = False
        self._normal_geometry = None
        self._drag_origin = None
        self.font_size = 12                 # bigger default for legibility
        self.theme_name = system_theme()    # follows the OS on startup
        self._themed_plain = []             # non-ttk widgets to re-color

        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        # Segoe MDL2 Assets is the icon font real Windows titlebars use;
        # its glyphs are uniform in size, unlike generic Unicode symbols.
        try:
            self._mdl2 = "Segoe MDL2 Assets" in tkfont.families()
        except tk.TclError:
            self._mdl2 = False
        if self._mdl2:
            self._glyphs = {"min": "\uE921", "max": "\uE922",
                            "restore": "\uE923", "close": "\uE8BB"}
        else:
            self._glyphs = {"min": "\u2500", "max": "\u25a1",
                            "restore": "\u25a1", "close": "\u2715"}
        self._set_app_icon()
        self._init_fonts()
        self._build_chrome()
        self.apply_theme()
        if self._custom_titlebar:
            self._enable_custom_titlebar()
        self._show_step()

    # ---------------------------------------------------- window geometry
    def _center_window(self, w, h):
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 3)   # slightly above center, like most apps
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _work_area(self):
        """Usable desktop area (screen minus taskbar) on Windows."""
        try:
            import ctypes
            from ctypes import wintypes
            rect = wintypes.RECT()
            SPI_GETWORKAREA = 0x0030
            ctypes.windll.user32.SystemParametersInfoW(
                SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
            return (rect.left, rect.top,
                    rect.right - rect.left, rect.bottom - rect.top)
        except Exception:
            return (0, 0, self.winfo_screenwidth(),
                    self.winfo_screenheight())

    # ---------------------------------------------------- custom titlebar
    def _enable_custom_titlebar(self):
        self.overrideredirect(True)
        self.update_idletasks()
        self._set_appwindow_style()
        # style changes only take effect after the window is re-shown
        self.wm_withdraw()
        self.after(30, self._reshow_on_top)
        self.bind("<Map>", self._on_map)

    def _set_appwindow_style(self):
        """Borderless windows lose their taskbar button and alt-tab entry
        on Windows; restore both with the WS_EX_APPWINDOW style."""
        try:
            import ctypes
            GWL_EXSTYLE = -20
            WS_EX_APPWINDOW = 0x00040000
            WS_EX_TOOLWINDOW = 0x00000080
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            style = (style & ~WS_EX_TOOLWINDOW) | WS_EX_APPWINDOW
            ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
        except Exception:
            pass

    def _reshow_on_top(self):
        """Show the window in front of other programs on launch."""
        self.wm_deiconify()
        self.lift()
        try:
            self.attributes("-topmost", True)
            self.after(700, lambda: self.attributes("-topmost", False))
        except tk.TclError:
            pass
        self.focus_force()

    def _on_map(self, _event=None):
        # restore borderless mode after un-minimizing from the taskbar,
        # and re-apply the taskbar style (toggling decorations can drop it)
        if self._custom_titlebar and not self.overrideredirect():
            self.overrideredirect(True)
            self.update_idletasks()
            self._set_appwindow_style()
            self.lift()

    def _minimize(self):
        # iconify() needs decorations on Windows, so drop them briefly
        self.overrideredirect(False)
        self.iconify()

    def _toggle_zoom(self):
        if self._is_zoomed:
            self.state("normal")
            if self._normal_geometry:
                self.geometry(self._normal_geometry)
            self._is_zoomed = False
            self._btn_zoom.config(text=self._glyphs["max"])
        else:
            self._normal_geometry = self.geometry()
            x, y, w, h = self._work_area()
            self.geometry(f"{w}x{h}+{x}+{y}")
            self._is_zoomed = True
            self._btn_zoom.config(text=self._glyphs["restore"])

    def _start_move(self, event):
        if self._is_zoomed:
            self._drag_origin = None
            return
        self._drag_origin = (event.x_root, event.y_root,
                             self.winfo_x(), self.winfo_y())

    def _do_move(self, event):
        if not self._drag_origin:
            return
        ox, oy, wx, wy = self._drag_origin
        self.geometry(f"+{wx + event.x_root - ox}+{wy + event.y_root - oy}")

    # --------------------------------------------------------------- icon
    def _set_app_icon(self):
        """Use the bundled logo as the window/taskbar icon on both OSes."""
        try:
            png = resource_path(os.path.join("assets", "logo.png"))
            if os.path.exists(png):
                self._icon_img = tk.PhotoImage(file=png)
                self.iconphoto(True, self._icon_img)
        except tk.TclError:
            pass
        if sys.platform.startswith("win"):
            try:
                ico = resource_path(os.path.join("assets", "logo.ico"))
                if os.path.exists(ico):
                    self.iconbitmap(default=ico)
            except tk.TclError:
                pass

    # -------------------------------------------------------------- fonts
    def _init_fonts(self):
        for name in ("TkDefaultFont", "TkTextFont", "TkHeadingFont",
                     "TkMenuFont", "TkFixedFont"):
            try:
                tkfont.nametofont(name).configure(size=self.font_size)
            except tk.TclError:
                pass

    def change_font(self, delta):
        self.font_size = max(9, min(18, self.font_size + delta))
        self._init_fonts()
        self.apply_theme()
        self._show_step()

    # -------------------------------------------------------------- theme
    def toggle_theme(self):
        self.theme_name = "dark" if self.theme_name == "light" else "light"
        self.apply_theme()
        self._show_step()

    def apply_theme(self):
        t = THEMES[self.theme_name]
        s = self.style
        base = ("", self.font_size)
        s.configure(".", background=t["bg"], foreground=t["fg"],
                    fieldbackground=t["field"], font=base,
                    bordercolor=t["line"], troughcolor=t["card"])
        s.configure("TFrame", background=t["bg"])
        s.configure("Card.TFrame", background=t["card"])
        s.configure("TLabel", background=t["bg"], foreground=t["fg"])
        s.configure("Card.TLabel", background=t["card"], foreground=t["fg"])
        s.configure("Muted.TLabel", background=t["bg"],
                    foreground=t["muted"])
        s.configure("Header.TFrame", background=t["header"])
        s.configure("Header.TLabel", background=t["header"],
                    foreground=t["header_fg"],
                    font=("", self.font_size + 3, "bold"))
        s.configure("HeaderSub.TLabel", background=t["header"],
                    foreground=t["header_fg"])
        s.configure("Step.TLabel", background=t["bg"], foreground=t["fg"],
                    font=("", self.font_size + 2, "bold"))
        s.configure("TLabelframe", background=t["card"],
                    bordercolor=t["line"])
        s.configure("TLabelframe.Label", background=t["card"],
                    foreground=t["accent"],
                    font=("", self.font_size, "bold"))
        s.configure("TCheckbutton", background=t["card"],
                    foreground=t["fg"])
        s.map("TCheckbutton", background=[("active", t["card"])])
        s.configure("TRadiobutton", background=t["card"],
                    foreground=t["fg"])
        s.map("TRadiobutton", background=[("active", t["card"])])
        s.configure("TButton", background=t["card"], foreground=t["fg"],
                    bordercolor=t["line"], focusthickness=1, padding=6)
        s.map("TButton", background=[("active", t["line"])])
        s.configure("Accent.TButton", background=t["accent"],
                    foreground=t["accent_fg"], padding=8,
                    font=("", self.font_size, "bold"))
        s.map("Accent.TButton",
              background=[("active", t["accent_hover"]),
                          ("disabled", t["line"])])
        s.configure("TEntry", fieldbackground=t["field"],
                    foreground=t["fg"], insertcolor=t["fg"])
        s.configure("TProgressbar", background=t["accent"],
                    troughcolor=t["card"], bordercolor=t["line"])
        tbfont = (("Segoe MDL2 Assets", max(8, self.font_size - 3))
                  if getattr(self, "_mdl2", False)
                  else ("", max(8, self.font_size - 2)))
        s.configure("Title.TButton", background=t["header"],
                    foreground=t["header_fg"], borderwidth=0,
                    padding=(10, 8), font=tbfont)
        s.map("Title.TButton", background=[("active", t["line"])])
        s.configure("TitleClose.TButton", background=t["header"],
                    foreground=t["header_fg"], borderwidth=0,
                    padding=(10, 8), font=tbfont)
        s.map("TitleClose.TButton",
              background=[("active", "#e81123")],
              foreground=[("active", "#ffffff")])
        s.configure("Vertical.TScrollbar", background=t["card"],
                    troughcolor=t["bg"], bordercolor=t["line"],
                    arrowcolor=t["fg"])
        self.configure(background=t["bg"])
        if getattr(self, "_custom_titlebar", False):
            try:
                self.configure(highlightthickness=1,
                               highlightbackground=t["line"],
                               highlightcolor=t["line"])
            except tk.TclError:
                pass
        for w, kind in list(self._themed_plain):
            try:
                if kind == "listbox":
                    w.configure(bg=t["field"], fg=t["fg"],
                                highlightbackground=t["line"],
                                selectbackground=t["accent"],
                                selectforeground=t["accent_fg"])
                elif kind == "canvas":
                    w.configure(bg=t["card"], highlightthickness=0)
            except tk.TclError:
                pass
        self._update_theme_button()

    def _update_theme_button(self):
        if hasattr(self, "btn_theme"):
            nxt = "Dark" if self.theme_name == "light" else "Light"
            self.btn_theme.config(text=f"{nxt} mode")

    # -------------------------------------------------------------- chrome
    def _build_chrome(self):
        if self._custom_titlebar:
            tb = ttk.Frame(self, style="Header.TFrame")
            tb.pack(fill="x")
            self._titlebar = tb
            ttk.Label(tb, text=f"  {APP_TITLE}",
                      style="HeaderSub.TLabel").pack(side="left", pady=4)
            ttk.Button(tb, text=self._glyphs["close"], width=5,
                       style="TitleClose.TButton",
                       command=self.destroy).pack(side="right", fill="y")
            self._btn_zoom = ttk.Button(tb, text=self._glyphs["max"],
                                        width=5, style="Title.TButton",
                                        command=self._toggle_zoom)
            self._btn_zoom.pack(side="right", fill="y")
            ttk.Button(tb, text=self._glyphs["min"], width=5,
                       style="Title.TButton",
                       command=self._minimize).pack(side="right", fill="y")
            for w in (tb, tb.winfo_children()[0]):
                w.bind("<ButtonPress-1>", self._start_move)
                w.bind("<B1-Motion>", self._do_move)
                w.bind("<Double-Button-1>", lambda e: self._toggle_zoom())

        header = ttk.Frame(self, style="Header.TFrame", padding=(16, 10))
        header.pack(fill="x")
        left = ttk.Frame(header, style="Header.TFrame")
        left.pack(side="left")
        ttk.Label(left, text=APP_TITLE,
                  style="Header.TLabel").pack(anchor="w")
        ttk.Label(left, text="UNP CCIT adviser tools",
                  style="HeaderSub.TLabel").pack(anchor="w")
        right = ttk.Frame(header, style="Header.TFrame")
        right.pack(side="right")
        ttk.Button(right, text="A-", width=3,
                   command=lambda: self.change_font(-1)).pack(side="left",
                                                              padx=2)
        ttk.Button(right, text="A+", width=3,
                   command=lambda: self.change_font(1)).pack(side="left",
                                                             padx=2)
        self.btn_theme = ttk.Button(right, command=self.toggle_theme)
        self.btn_theme.pack(side="left", padx=8)
        ttk.Button(right, text="About",
                   command=self.show_about).pack(side="left")

        self.header_lbl = ttk.Label(self, text="", style="Step.TLabel")
        self.header_lbl.pack(anchor="w", padx=16, pady=(12, 4))
        self.progress_lbl = ttk.Label(self, text="", style="Muted.TLabel")
        self.progress_lbl.pack(anchor="w", padx=16)

        self.body = ttk.Frame(self)
        self.body.pack(fill="both", expand=True, padx=16, pady=8)

        nav = ttk.Frame(self, padding=(16, 10))
        nav.pack(fill="x", side="bottom")
        self.btn_back = ttk.Button(nav, text="< Back", command=self.go_back)
        self.btn_back.pack(side="left")
        if self._custom_titlebar:
            ttk.Sizegrip(nav).pack(side="right", padx=(6, 0))
        self.btn_next = ttk.Button(nav, text="Next >",
                                   style="Accent.TButton",
                                   command=self.go_next)
        self.btn_next.pack(side="right")
        self.status = ttk.Label(nav, text="", style="Muted.TLabel")
        self.status.pack(side="left", padx=16)

    def show_about(self):
        top = tk.Toplevel(self)
        top.title(f"About {APP_TITLE}")
        top.resizable(False, False)
        top.transient(self)
        t = THEMES[self.theme_name]
        top.configure(background=t["bg"])
        frame = ttk.Frame(top, padding=24)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=APP_TITLE,
                  font=("", self.font_size + 4, "bold")).pack(anchor="w")
        ttk.Label(frame, text=f"Version {APP_VERSION}",
                  style="Muted.TLabel").pack(anchor="w", pady=(0, 12))
        ttk.Label(frame, justify="left", text=(
            "Fills the UNP CCIT Appraisal Sheet (VPAA-CCIT-QF-05) and\n"
            "Report of Rating (VPAA-CCIT-QF-09) for every advisee, using\n"
            "the Periodic Grades Listing PDFs from the student portal.\n\n"
            "Everything runs locally on this computer. No data is sent\n"
            "anywhere.")).pack(anchor="w")
        ttk.Label(frame, text=f"\nDeveloped by {DEVELOPER}",
                  font=("", self.font_size, "bold")).pack(anchor="w")
        ttk.Button(frame, text="Close", style="Accent.TButton",
                   command=top.destroy).pack(anchor="e", pady=(16, 0))
        top.update_idletasks()
        tw, th = top.winfo_reqwidth(), top.winfo_reqheight()
        px = self.winfo_x() + (self.winfo_width() - tw) // 2
        py = self.winfo_y() + (self.winfo_height() - th) // 2
        top.geometry(f"+{max(0, px)}+{max(0, py)}")
        top.grab_set()

    def _clear_body(self):
        self._themed_plain = []
        for w in self.body.winfo_children():
            w.destroy()

    def _show_step(self):
        self._clear_body()
        self.header_lbl.config(
            text=f"Step {self.step + 1}: {self.STEPS[self.step]}")
        done = "\u25cf" * (self.step + 1) + "\u25cb" * (4 - self.step)
        self.progress_lbl.config(text=f"{done}   {self.step + 1} of 5")
        self.btn_back.config(state="normal" if self.step > 0 else "disabled")
        self.btn_next.config(
            text="Next >", state="disabled" if self.step == 4 else "normal")
        [self._step_files, self._step_students, self._step_documents,
         self._step_faculty, self._step_generate][self.step]()
        self.apply_theme()

    def go_back(self):
        if self.busy:
            return
        if self.step > 0:
            self.step -= 1
            self._show_step()

    def go_next(self):
        if self.busy or self.step == 4:
            return
        if self.step == 0:
            self._parse_async()
            return
        ok = [None, self._leave_students, self._leave_documents,
              lambda: True][self.step]
        if ok():
            self.step += 1
            self._show_step()

    # ------------------------------------------------------- step 1: files
    def _step_files(self):
        f = self.body
        ttk.Label(f, text="Select one or more Periodic Grades Listing "
                          "PDFs. You can pick several terms at once.\n"
                          "Students are matched across PDFs by ID "
                          "number.").pack(anchor="w", pady=6)
        ttk.Button(f, text="Add PDF files...", style="Accent.TButton",
                   command=self._pick_pdfs).pack(anchor="w", pady=6)
        self.file_list = tk.Listbox(f, height=8, borderwidth=1,
                                    relief="solid", activestyle="none")
        self.file_list.pack(fill="both", expand=True, pady=6)
        self._themed_plain.append((self.file_list, "listbox"))
        for p in self.pdf_paths:
            self.file_list.insert("end", p)
        ttk.Button(f, text="Remove selected",
                   command=self._remove_pdf).pack(anchor="w")

    def _pick_pdfs(self):
        paths = filedialog.askopenfilenames(
            title="Select grade listing PDFs",
            filetypes=[("PDF files", "*.pdf")])
        for p in paths:
            if p not in self.pdf_paths:
                self.pdf_paths.append(p)
                self.file_list.insert("end", p)

    def _remove_pdf(self):
        for i in reversed(self.file_list.curselection()):
            self.pdf_paths.pop(i)
            self.file_list.delete(i)

    def _parse_async(self):
        """Parse PDFs on a worker thread so the window never freezes."""
        if not self.pdf_paths:
            messagebox.showwarning(APP_TITLE, "Please add at least one PDF.")
            return
        self.busy = True
        self.btn_next.config(state="disabled")
        self.btn_back.config(state="disabled")
        self.status.config(text="Reading PDFs, please wait...")
        paths = list(self.pdf_paths)

        def work():
            try:
                _load_engine()
                parsed = [grade_parser.parse_pdf(p) for p in paths]
                merged = grade_parser.merge_terms(parsed)
                self._ui(lambda: self._parse_done(merged, None))
            except Exception as e:
                msg = str(e)
                self._ui(lambda: self._parse_done(None, msg))

        threading.Thread(target=work, daemon=True).start()

    def _parse_done(self, merged, error):
        self.busy = False
        self.status.config(text="")
        self.btn_next.config(state="normal")
        self.btn_back.config(state="normal" if self.step > 0
                             else "disabled")
        if error:
            messagebox.showerror(APP_TITLE,
                                 f"Could not parse a PDF:\n{error}")
            return
        if not merged:
            messagebox.showwarning(
                APP_TITLE, "No students found. Is this a Periodic Grades "
                           "Listing PDF from the portal?")
            return
        self.merged = merged
        self.term_keys = sorted({k for s in merged.values()
                                 for k in s["terms"]})
        self.student_vars = {}
        self.term_vars = {}
        self.faculty_entries = {}
        self.step = 1
        self._show_step()

    # ---------------------------------------------------- step 2: students
    def _step_students(self):
        f = self.body
        students = sorted(self.merged.values(), key=lambda s: s["name"])
        ttk.Label(f, text=f"Parsed {len(students)} students. Untick anyone "
                          "you want to skip.").pack(anchor="w", pady=6)
        bar = ttk.Frame(f)
        bar.pack(anchor="w", pady=(0, 6))
        ttk.Button(bar, text="Select all",
                   command=lambda: self._set_all(True)).pack(side="left")
        ttk.Button(bar, text="Select none",
                   command=lambda: self._set_all(False)).pack(side="left",
                                                              padx=6)
        canvas, inner = self._scroll_area(f)
        hdr = ttk.Frame(inner, style="Card.TFrame")
        hdr.pack(fill="x")
        for txt, w in [("", 3), ("Name", 38), ("ID", 11), ("Course", 8),
                       ("Terms", 6), ("Subjects", 8)]:
            ttk.Label(hdr, text=txt, width=w, style="Card.TLabel",
                      font=("", self.font_size, "bold")).pack(side="left")
        for s in students:
            var = self.student_vars.setdefault(s["id"],
                                               tk.BooleanVar(value=True))
            row = ttk.Frame(inner, style="Card.TFrame")
            row.pack(fill="x")
            ttk.Checkbutton(row, variable=var, width=2).pack(side="left")
            n_subj = sum(len(v) for v in s["terms"].values())
            for txt, w in [(s["name"], 38), (s["id"], 11),
                           (s["course"], 8), (len(s["terms"]), 6),
                           (n_subj, 8)]:
                ttk.Label(row, text=txt, width=w,
                          style="Card.TLabel").pack(side="left")

    def _set_all(self, value):
        for var in self.student_vars.values():
            var.set(value)

    def _leave_students(self):
        if not any(v.get() for v in self.student_vars.values()):
            messagebox.showwarning(APP_TITLE, "Select at least one student.")
            return False
        return True

    # --------------------------------------------------- step 3: documents
    def _step_documents(self):
        f = self.body
        box1 = ttk.LabelFrame(f, text="Documents to generate", padding=10)
        box1.pack(fill="x", pady=6)
        ttk.Checkbutton(box1, text="Appraisal Sheet (one per student, all "
                                   "terms)",
                        variable=self.doc_appraisal).pack(anchor="w", pady=2)
        ttk.Checkbutton(box1, text="Report of Rating (one per student per "
                                   "term)",
                        variable=self.doc_report).pack(anchor="w", pady=2)

        box2 = ttk.LabelFrame(f, text="Terms to include (Report of Rating)",
                              padding=10)
        box2.pack(fill="x", pady=6)
        for key in self.term_keys:
            var = self.term_vars.setdefault(key, tk.BooleanVar(value=True))
            ttk.Checkbutton(box2, text=f"{key[0]} {key[1]}",
                            variable=var).pack(anchor="w", pady=2)
        ttk.Label(box2, text="The Appraisal Sheet always includes every "
                             "term found in the PDFs.",
                  style="Card.TLabel").pack(anchor="w", pady=2)

        box3 = ttk.LabelFrame(f, text="Output format", padding=10)
        box3.pack(fill="x", pady=6)
        ttk.Radiobutton(box3, text="Individual files (one docx per student)",
                        variable=self.output_mode,
                        value="individual").pack(anchor="w", pady=2)
        ttk.Radiobutton(box3, text="One batch file (all students merged, "
                                   "one per page)",
                        variable=self.output_mode,
                        value="combined").pack(anchor="w", pady=2)
        ttk.Checkbutton(box3, text="Remove unused blank rows in the Report "
                                   "of Rating table",
                        variable=self.trim_rows).pack(anchor="w", pady=2)

    def _leave_documents(self):
        if not (self.doc_appraisal.get() or self.doc_report.get()):
            messagebox.showwarning(APP_TITLE,
                                   "Pick at least one document type.")
            return False
        if self.doc_report.get() and self.term_vars and \
                not any(v.get() for v in self.term_vars.values()):
            messagebox.showwarning(APP_TITLE,
                                   "Pick at least one term for the Report "
                                   "of Rating, or untick it.")
            return False
        return True

    # ----------------------------------------------------- step 4: faculty
    def _subject_codes(self):
        codes = {}
        for s in self.merged.values():
            for subjects in s["terms"].values():
                for subj in subjects:
                    codes.setdefault(subj["code"], subj["title"])
        return sorted(codes.items())

    def _step_faculty(self):
        f = self.body
        ttk.Label(f, text="Adviser name (printed on the Report of "
                          "Rating):").pack(anchor="w", pady=4)
        ttk.Entry(f, textvariable=self.adviser,
                  width=44).pack(anchor="w")
        ttk.Label(f, text="\nFaculty per subject. These codes were found "
                          "in your PDFs. Leave a name blank\nand the "
                          "column stays blank in the "
                          "documents.").pack(anchor="w", pady=4)
        canvas, inner = self._scroll_area(f)
        hdr = ttk.Frame(inner, style="Card.TFrame")
        hdr.pack(fill="x")
        for txt, w in [("Code", 11), ("Descriptive Title", 42),
                       ("Faculty name", 28)]:
            ttk.Label(hdr, text=txt, width=w, style="Card.TLabel",
                      font=("", self.font_size, "bold")).pack(side="left")
        for code, title in self._subject_codes():
            var = self.faculty_entries.setdefault(code, tk.StringVar())
            row = ttk.Frame(inner, style="Card.TFrame")
            row.pack(fill="x", pady=1)
            ttk.Label(row, text=code, width=11,
                      style="Card.TLabel").pack(side="left")
            ttk.Label(row, text=title[:46], width=42,
                      style="Card.TLabel").pack(side="left")
            ttk.Entry(row, textvariable=var, width=30).pack(side="left")

    # ---------------------------------------------------- step 5: generate
    def _step_generate(self):
        f = self.body
        n_students = sum(1 for v in self.student_vars.values() if v.get())
        picked_terms = [k for k, v in self.term_vars.items() if v.get()]
        docs = []
        if self.doc_appraisal.get():
            docs.append("Appraisal Sheet")
        if self.doc_report.get():
            docs.append("Report of Rating")
        mode = ("individual files" if self.output_mode.get() == "individual"
                else "one batch file per document type")
        summary = (f"Students: {n_students}\n"
                   f"Documents: {', '.join(docs)}\n"
                   f"Terms for ROR: "
                   f"{', '.join(f'{t} {sy}' for t, sy in picked_terms)}\n"
                   f"Output: {mode}\n"
                   f"Trim blank ROR rows: "
                   f"{'yes' if self.trim_rows.get() else 'no'}")
        box = ttk.LabelFrame(f, text="Review your choices", padding=10)
        box.pack(fill="x", pady=6)
        ttk.Label(box, text=summary, justify="left",
                  style="Card.TLabel").pack(anchor="w")

        out = ttk.Frame(f)
        out.pack(fill="x", pady=8)
        ttk.Label(out, text="Save to folder:").pack(side="left")
        ttk.Entry(out, textvariable=self.out_dir,
                  width=48).pack(side="left", padx=6)
        ttk.Button(out, text="Browse...",
                   command=self._pick_out_dir).pack(side="left")

        self.progress = ttk.Progressbar(f, mode="determinate")
        self.progress.pack(fill="x", pady=10)
        self.gen_status = ttk.Label(f, text="", style="Muted.TLabel")
        self.gen_status.pack(anchor="w")

        self.btn_generate = ttk.Button(f, text="Generate documents",
                                       style="Accent.TButton",
                                       command=self._generate_clicked)
        self.btn_generate.pack(pady=10)
        self.btn_open = ttk.Button(f, text="Open output folder",
                                   command=self._open_out_dir,
                                   state="disabled")
        self.btn_open.pack()

    def _pick_out_dir(self):
        d = filedialog.askdirectory(title="Choose output folder")
        if d:
            self.out_dir.set(d)

    def _open_out_dir(self):
        path = clean_dir_path(self.out_dir.get())
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # noqa
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}" >/dev/null 2>&1 &')
        except OSError as e:
            messagebox.showerror(APP_TITLE, f"Could not open folder:\n{e}")

    def _generate_clicked(self):
        # snapshot EVERYTHING on the main thread; the worker never touches
        # tk variables (reading them off-thread breaks on Windows)
        out_root = clean_dir_path(self.out_dir.get())
        if not out_root:
            messagebox.showwarning(APP_TITLE, "Choose an output folder.")
            return
        try:
            os.makedirs(out_root, exist_ok=True)
            probe = os.path.join(out_root, ".write_test")
            with open(probe, "w") as fh:
                fh.write("ok")
            os.remove(probe)
        except OSError as e:
            messagebox.showerror(
                APP_TITLE,
                f"Cannot write to this folder:\n{out_root}\n\n{e}\n\n"
                "Pick another folder (use Browse to avoid typos).")
            return

        cfg = {
            "out_root": out_root,
            "students": [s for sid, s in sorted(
                self.merged.items(), key=lambda kv: kv[1]["name"])
                if self.student_vars.get(sid) is not None
                and self.student_vars[sid].get()],
            "faculty_map": {c.replace(" ", "").upper(): v.get().strip()
                            for c, v in self.faculty_entries.items()
                            if v.get().strip()},
            "adviser": self.adviser.get().strip(),
            "trim": self.trim_rows.get(),
            "picked": {k for k, v in self.term_vars.items() if v.get()},
            "mode": self.output_mode.get(),
            "appraisal": self.doc_appraisal.get(),
            "report": self.doc_report.get(),
            "term_keys": list(self.term_keys),
        }
        self.btn_generate.config(state="disabled")
        self.btn_open.config(state="disabled")
        threading.Thread(target=self._generate, args=(cfg,),
                         daemon=True).start()

    def _generate(self, cfg):
        try:
            _load_engine()
            students = cfg["students"]
            jobs = []
            if cfg["mode"] == "combined":
                if cfg["appraisal"]:
                    jobs.append(("appraisal_all", None))
                if cfg["report"]:
                    for key in cfg["term_keys"]:
                        if key in cfg["picked"]:
                            jobs.append(("report_all", key))
            else:
                for s in students:
                    if cfg["appraisal"]:
                        jobs.append(("appraisal", s))
                    if cfg["report"]:
                        for key in sorted(s["terms"]):
                            if key in cfg["picked"]:
                                jobs.append(("report", (s, key)))

            total = len(jobs)
            self._ui(lambda: self.progress.config(maximum=max(total, 1),
                                                  value=0))
            done = 0
            for kind, payload in jobs:
                self._run_job(kind, payload, cfg)
                done += 1
                self._ui(lambda d=done: (
                    self.progress.config(value=d),
                    self.gen_status.config(
                        text=f"Generated {d} of {total}")))

            self._ui(lambda: (
                self.gen_status.config(
                    text=f"Done. Output written to {cfg['out_root']}"),
                self.btn_open.config(state="normal"),
                self.btn_generate.config(state="normal")))
        except Exception:
            err = traceback.format_exc(limit=3)
            self._ui(lambda: (
                self.btn_generate.config(state="normal"),
                messagebox.showerror(APP_TITLE,
                                     f"Generation failed:\n{err}")))

    def _run_job(self, kind, payload, cfg):
        out_root = cfg["out_root"]
        fm, adv, trim = cfg["faculty_map"], cfg["adviser"], cfg["trim"]
        if kind == "appraisal":
            s = payload
            folder = os.path.join(out_root, "Appraisal_Sheets")
            os.makedirs(folder, exist_ok=True)
            buf = filler.fill_appraisal(s, fm)
            path = os.path.join(folder,
                                f"{safe_name(s['name'])}_{s['id']}.docx")
            with open(path, "wb") as fh:
                fh.write(buf.read())
        elif kind == "report":
            s, (term, sy) = payload
            tslug = term.replace(" ", "")
            folder = os.path.join(out_root, "Reports_of_Rating",
                                  f"{tslug}_{sy}")
            os.makedirs(folder, exist_ok=True)
            buf = filler.fill_report(s, term, sy, s["terms"][(term, sy)],
                                     fm, adv, trim)
            path = os.path.join(folder,
                                f"{safe_name(s['name'])}_{s['id']}.docx")
            with open(path, "wb") as fh:
                fh.write(buf.read())
        elif kind == "appraisal_all":
            docs = [filler.build_appraisal(s, fm) for s in cfg["students"]]
            buf = filler.combine_documents(docs)
            if buf:
                with open(os.path.join(out_root,
                                       "Appraisal_Sheets_ALL.docx"),
                          "wb") as fh:
                    fh.write(buf.read())
        elif kind == "report_all":
            term, sy = payload
            docs = [filler.build_report(s, term, sy, s["terms"][(term, sy)],
                                        fm, adv, trim)
                    for s in cfg["students"] if (term, sy) in s["terms"]]
            buf = filler.combine_documents(docs)
            if buf:
                tslug = term.replace(" ", "")
                fname = f"Report_of_Rating_{tslug}_{sy}_ALL.docx"
                with open(os.path.join(out_root, fname), "wb") as fh:
                    fh.write(buf.read())

    # ------------------------------------------------------------ helpers
    def _ui(self, fn):
        """Schedule a UI update from any thread (thread-safe)."""
        self._msgq.put(fn)

    def _poll_queue(self):
        try:
            while True:
                fn = self._msgq.get_nowait()
                try:
                    fn()
                except tk.TclError:
                    pass
        except queue.Empty:
            pass
        self.after(50, self._poll_queue)

    def _scroll_area(self, parent):
        holder = ttk.Frame(parent, style="Card.TFrame")
        holder.pack(fill="both", expand=True, pady=6)
        canvas = tk.Canvas(holder, highlightthickness=0)
        vsb = ttk.Scrollbar(holder, orient="vertical",
                            command=canvas.yview)
        inner = ttk.Frame(canvas, style="Card.TFrame")
        inner.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._themed_plain.append((canvas, "canvas"))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(
            int(-e.delta / 120), "units"))
        canvas.bind_all("<Button-4>",
                        lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>",
                        lambda e: canvas.yview_scroll(1, "units"))
        return canvas, inner


if __name__ == "__main__":
    Wizard().mainloop()
