import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import json
import os

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    from pystray import Icon as TrayIcon, Menu as TrayMenu, MenuItem as TrayMenuItem
    import threading
    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(SCRIPT_DIR, "commands.json")
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")

# ---------------------------------------------------------------------------
# Theme — light blue / grey / white
# ---------------------------------------------------------------------------
THEME = {
    "bg":              "#F0F4F8",
    "sidebar_bg":      "#E2E8F0",
    "sidebar_active":  "#CBD5E1",
    "sidebar_fg":      "#334155",
    "sidebar_sep":     "#CBD5E1",
    "content_bg":      "#F8FAFC",
    "card_bg":         "#FFFFFF",
    "card_border":     "#E2E8F0",
    "card_hover":      "#EFF6FF",
    "card_hover_border": "#93C5FD",
    "card_shadow":     "#E2E8F0",
    "accent":          "#3B82F6",
    "accent_hover":    "#2563EB",
    "accent_fg":       "#FFFFFF",
    "accent_light":    "#DBEAFE",
    "btn_bg":          "#E2E8F0",
    "btn_fg":          "#1E293B",
    "btn_danger":      "#EF4444",
    "btn_danger_hover":"#DC2626",
    "text_primary":    "#1E293B",
    "text_secondary":  "#64748B",
    "text_muted":      "#94A3B8",
    "search_bg":       "#FFFFFF",
    "search_border":   "#CBD5E1",
    "status_bg":       "#E2E8F0",
    "toast_bg":        "#1E293B",
    "toast_fg":        "#F8FAFC",
    "output_bg":       "#1E293B",
    "output_fg":       "#E2E8F0",
    "output_header":   "#334155",
    "badge_bg":        "#DBEAFE",
    "badge_fg":        "#1D4ED8",
}

FONT_FAMILY = "Segoe UI"
FONTS = {
    "heading":   (FONT_FAMILY, 13, "bold"),
    "title":     (FONT_FAMILY, 11, "bold"),
    "body":      (FONT_FAMILY, 10),
    "small":     (FONT_FAMILY, 9),
    "card_name": (FONT_FAMILY, 10, "bold"),
    "card_cmd":  (FONT_FAMILY, 8),
    "card_btn":  (FONT_FAMILY, 9),
    "fab":       (FONT_FAMILY, 18, "bold"),
    "status":    (FONT_FAMILY, 9),
    "toast":     (FONT_FAMILY, 10),
    "sidebar":   (FONT_FAMILY, 10),
    "output":    ("Consolas", 9),
    "badge":     (FONT_FAMILY, 8),
    "empty":     (FONT_FAMILY, 12),
    "empty_sub": (FONT_FAMILY, 10),
    "search_icon": (FONT_FAMILY, 12),
}

CARD_IMAGE_SIZE = (160, 100)
CARD_MAX_NAME = 22
COLUMNS = 3


# ---------------------------------------------------------------------------
# Tooltip helper
# ---------------------------------------------------------------------------
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self._tw = None
        widget.bind("<Enter>", self._show, add="+")
        widget.bind("<Leave>", self._hide, add="+")

    def _show(self, event):
        if self._tw:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._tw = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(
            tw, text=self.text, font=FONTS["card_cmd"],
            bg=THEME["toast_bg"], fg=THEME["toast_fg"],
            padx=8, pady=4, wraplength=300, justify="left"
        )
        lbl.pack()

    def _hide(self, event):
        if self._tw:
            self._tw.destroy()
            self._tw = None

    def update_text(self, text):
        self.text = text


# ---------------------------------------------------------------------------
# PlaceholderEntry
# ---------------------------------------------------------------------------
class PlaceholderEntry(tk.Entry):
    def __init__(self, master, placeholder="", **kw):
        self._ph = placeholder
        self._ph_color = THEME["text_muted"]
        self._fg = kw.get("fg", THEME["text_primary"])
        self._has_placeholder = False
        super().__init__(master, **kw)
        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        self._on_focus_out(None)

    def _on_focus_in(self, event):
        if self._has_placeholder:
            self._has_placeholder = False
            self.config(fg=self._fg)
            # Temporarily block the trace while clearing placeholder
            var = self.cget("textvariable")
            self.config(textvariable="")
            self.delete(0, tk.END)
            self.config(textvariable=var)

    def _on_focus_out(self, event):
        if not self.get():
            self._has_placeholder = True
            # Temporarily disconnect the StringVar so inserting placeholder
            # does not fire the trace callback
            var = self.cget("textvariable")
            self.config(textvariable="")
            self.delete(0, tk.END)
            self.insert(0, self._ph)
            self.config(fg=self._ph_color)
            self.config(textvariable=var)


class PowerShellApp:
    # -----------------------------------------------------------------------
    # Init
    # -----------------------------------------------------------------------
    def __init__(self, root):
        self.root = root
        self.root.title("PowerShell Command Runner")
        self.root.configure(bg=THEME["bg"])
        self.root.minsize(750, 500)

        self._load_geometry()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.commands = self.load_commands()
        self.custom_categories = self._load_custom_categories()
        self._drag_data = {"idx": None}
        self._visible_count = 0

        self.create_menubar()

        # ---- Status bar (pack first so it stays at bottom) ----
        self.status_bar = tk.Frame(root, bg=THEME["status_bg"], height=28)
        self.status_bar.pack(side="bottom", fill="x")
        self.status_bar.pack_propagate(False)

        self._status_left = tk.Label(
            self.status_bar, text="", font=FONTS["status"],
            bg=THEME["status_bg"], fg=THEME["text_secondary"], anchor="w"
        )
        self._status_left.pack(side="left", padx=12)
        self._status_right = tk.Label(
            self.status_bar, text="Ctrl+N  New  |  Ctrl+F  Find  |  F5  Refresh",
            font=FONTS["status"], bg=THEME["status_bg"], fg=THEME["text_muted"], anchor="e"
        )
        self._status_right.pack(side="right", padx=12)

        # ---- Output panel (hidden, pack before main so it's above status) ----
        self._output_visible = False
        self.output_frame = tk.Frame(root, bg=THEME["output_bg"])

        output_header = tk.Frame(self.output_frame, bg=THEME["output_header"])
        output_header.pack(fill="x")
        tk.Label(
            output_header, text="  Output", font=FONTS["small"],
            bg=THEME["output_header"], fg=THEME["output_fg"], anchor="w"
        ).pack(side="left", padx=4, pady=3)
        clear_btn = tk.Label(
            output_header, text="  Clear  ", font=FONTS["small"],
            bg=THEME["output_header"], fg=THEME["text_muted"], cursor="hand2"
        )
        clear_btn.pack(side="right", padx=8, pady=3)
        clear_btn.bind("<Button-1>", lambda e: self._clear_output())
        close_out_btn = tk.Label(
            output_header, text="  X  ", font=FONTS["small"],
            bg=THEME["output_header"], fg=THEME["text_muted"], cursor="hand2"
        )
        close_out_btn.pack(side="right", pady=3)
        close_out_btn.bind("<Button-1>", lambda e: (self._output_toggle_var.set(False), self.toggle_output()))

        self.output_text = tk.Text(
            self.output_frame, bg=THEME["output_bg"], fg=THEME["output_fg"],
            font=FONTS["output"], wrap="word", relief="flat", state="disabled",
            insertbackground=THEME["output_fg"], padx=10, pady=6, height=8
        )
        output_scroll = ttk.Scrollbar(self.output_frame, orient="vertical", command=self.output_text.yview)
        self.output_text.configure(yscrollcommand=output_scroll.set)
        self.output_text.pack(side="left", fill="both", expand=True)
        output_scroll.pack(side="right", fill="y")

        # ---- Main container ----
        main_container = tk.Frame(root, bg=THEME["bg"])
        main_container.pack(fill="both", expand=True)

        # ---- Sidebar ----
        self.sidebar = tk.Frame(main_container, width=180, bg=THEME["sidebar_bg"])
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        sidebar_title = tk.Label(
            self.sidebar, text="PS Runner", font=FONTS["title"],
            bg=THEME["sidebar_bg"], fg=THEME["sidebar_fg"], anchor="w"
        )
        sidebar_title.pack(fill="x", padx=14, pady=(16, 8))

        # Thin separator under title
        tk.Frame(self.sidebar, bg=THEME["sidebar_sep"], height=1).pack(fill="x", padx=14, pady=(0, 6))

        self._active_category = "Home"
        self._search_mode = False
        self._sidebar_buttons = {}
        self._sidebar_extras = []  # separators and other non-button widgets to clean up
        self._build_sidebar()

        # Separator line between sidebar and content
        tk.Frame(main_container, bg=THEME["sidebar_sep"], width=1).pack(side="left", fill="y")

        # ---- Content area ----
        self.content_area = tk.Frame(main_container, bg=THEME["content_bg"])
        self.content_area.pack(side="left", fill="both", expand=True)

        # Search bar (hidden by default, shown when Search is clicked)
        self.search_frame = tk.Frame(self.content_area, bg=THEME["content_bg"])

        search_icon = tk.Label(
            self.search_frame, text="\U0001F50D", font=FONTS["search_icon"],
            bg=THEME["content_bg"], fg=THEME["text_muted"]
        )
        search_icon.pack(side="left", padx=(0, 6))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.update_list)
        self._search_after_id = None

        self.search_entry = PlaceholderEntry(
            self.search_frame, placeholder="Search commands...",
            textvariable=self.search_var, font=FONTS["body"],
            bg=THEME["search_bg"], fg=THEME["text_primary"],
            relief="solid", bd=1, insertbackground=THEME["text_primary"]
        )
        self.search_entry.pack(side="left", fill="x", expand=True, ipady=5)

        # Close search button
        close_search = tk.Label(
            self.search_frame, text=" \u2715 ", font=FONTS["body"],
            bg=THEME["content_bg"], fg=THEME["text_muted"], cursor="hand2"
        )
        close_search.pack(side="right", padx=(6, 0))
        close_search.bind("<Button-1>", lambda e: self._exit_search())
        self.root.bind("<Escape>", lambda e: self._exit_search() if self._search_mode else None)

        # Scrollable card grid
        self.canvas = tk.Canvas(self.content_area, bg=THEME["content_bg"], highlightthickness=0, bd=0)
        self.scrollbar = ttk.Scrollbar(self.content_area, orient="vertical", command=self.canvas.yview)
        self.buttons_frame = tk.Frame(self.canvas, bg=THEME["content_bg"])

        self.buttons_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._canvas_win = self.canvas.create_window((0, 0), window=self.buttons_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        self.canvas.pack(side="left", fill="both", expand=True, padx=(20, 0), pady=(4, 10))
        self.scrollbar.pack(side="right", fill="y", pady=4)

        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        # Card cache
        self._card_widgets = []
        self._image_refs = []
        self._card_tooltips = []

        # Empty state label (shown when no cards visible)
        self._empty_frame = tk.Frame(self.buttons_frame, bg=THEME["content_bg"])
        tk.Label(
            self._empty_frame, text="No commands yet",
            font=FONTS["empty"], bg=THEME["content_bg"], fg=THEME["text_muted"]
        ).pack(pady=(60, 4))
        tk.Label(
            self._empty_frame, text="Click  +  or press Ctrl+N to add your first command",
            font=FONTS["empty_sub"], bg=THEME["content_bg"], fg=THEME["text_muted"]
        ).pack()

        # ---- FAB ----
        self.fab = tk.Canvas(root, width=54, height=54, bg=THEME["content_bg"], highlightthickness=0, bd=0)
        self.fab.place(relx=0.93, rely=0.86, anchor="center")
        self._draw_fab()
        self.fab.bind("<Button-1>", lambda e: self.add_new_command())
        self.fab.bind("<Enter>", lambda e: self._draw_fab(hover=True))
        self.fab.bind("<Leave>", lambda e: self._draw_fab(hover=False))

        # ---- Toast overlay ----
        self._toast_label = tk.Label(
            root, text="", font=FONTS["toast"], bg=THEME["toast_bg"],
            fg=THEME["toast_fg"], padx=18, pady=8
        )
        self._toast_after_id = None

        # Build
        self._build_cards()
        self.refresh_buttons()
        self._update_status()

        # System tray (optional)
        self._tray_icon = None
        if HAS_TRAY and HAS_PIL:
            self._setup_tray()

    # -----------------------------------------------------------------------
    # FAB
    # -----------------------------------------------------------------------
    def _draw_fab(self, hover=False):
        self.fab.delete("all")
        color = THEME["accent_hover"] if hover else THEME["accent"]
        # Shadow
        self.fab.create_oval(4, 5, 52, 53, fill=THEME["card_shadow"], outline="")
        # Circle
        self.fab.create_oval(2, 2, 50, 50, fill=color, outline="")
        self.fab.create_text(26, 25, text="+", fill=THEME["accent_fg"], font=FONTS["fab"])

    # -----------------------------------------------------------------------
    # Canvas resize
    # -----------------------------------------------------------------------
    def _on_canvas_resize(self, event):
        self.canvas.itemconfig(self._canvas_win, width=event.width)

    # -----------------------------------------------------------------------
    # Geometry persistence
    # -----------------------------------------------------------------------
    def _load_geometry(self):
        try:
            with open(CONFIG_FILE, "r") as f:
                cfg = json.load(f)
            geo = cfg.get("geometry")
            if geo:
                self.root.geometry(geo)
                return
        except Exception:
            pass
        self.root.geometry("960x680")

    def _save_geometry(self):
        cfg = {}
        try:
            with open(CONFIG_FILE, "r") as f:
                cfg = json.load(f)
        except Exception:
            pass
        cfg["geometry"] = self.root.geometry()
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(cfg, f)
        except Exception:
            pass

    def _on_close(self):
        self._save_geometry()
        if self._tray_icon:
            self._tray_icon.stop()
        self.root.destroy()

    # -----------------------------------------------------------------------
    # Sidebar
    # -----------------------------------------------------------------------
    def _get_categories(self):
        cats = set()
        for item in self.commands:
            cat = item.get("category", "")
            if cat:
                cats.add(cat)
        # Merge with custom categories
        for cat_info in self.custom_categories:
            cats.add(cat_info["name"])
        return sorted(cats)

    def _get_category_icon(self, cat_name):
        for cat_info in self.custom_categories:
            if cat_info["name"] == cat_name:
                return cat_info.get("icon", "")
        return ""

    def _load_custom_categories(self):
        try:
            with open(CONFIG_FILE, "r") as f:
                cfg = json.load(f)
            cats = cfg.get("categories", [])
            if isinstance(cats, list):
                return cats
        except Exception:
            pass
        return []

    def _save_custom_categories(self):
        cfg = {}
        try:
            with open(CONFIG_FILE, "r") as f:
                cfg = json.load(f)
        except Exception:
            pass
        cfg["categories"] = self.custom_categories
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(cfg, f)
        except Exception:
            pass

    def _build_sidebar(self):
        for w in list(self._sidebar_buttons.values()):
            if isinstance(w, tuple):
                w[0].destroy()
            else:
                w.destroy()
        self._sidebar_buttons.clear()
        for w in self._sidebar_extras:
            w.destroy()
        self._sidebar_extras.clear()

        # Fixed items: Search, Home
        for fixed in ("Search", "Home"):
            if fixed == "Home":
                icon = "\u2302"
                count = len(self.commands)
            else:
                icon = "\U0001F50D"
                count = None

            container = tk.Frame(self.sidebar, bg=THEME["sidebar_bg"], cursor="hand2")
            container.pack(fill="x")

            lbl = tk.Label(
                container, text=f" {icon}  {fixed}", font=FONTS["sidebar"],
                bg=THEME["sidebar_bg"], fg=THEME["sidebar_fg"], anchor="w"
            )
            lbl.pack(side="left", fill="x", expand=True, padx=(14, 0), pady=7)

            if count is not None:
                count_lbl = tk.Label(
                    container, text=str(count), font=FONTS["badge"],
                    bg=THEME["sidebar_bg"], fg=THEME["text_muted"], anchor="e"
                )
                count_lbl.pack(side="right", padx=(0, 14), pady=7)
            else:
                count_lbl = None

            click_widgets = [w for w in (container, lbl, count_lbl) if w is not None]
            for w in click_widgets:
                if fixed == "Search":
                    w.bind("<Button-1>", lambda e: self._activate_search())
                else:
                    w.bind("<Button-1>", lambda e, c=fixed: self._select_category(c))
                w.bind("<Enter>", lambda e, cn=container, lb=lbl, cl=count_lbl, c=fixed: self._sidebar_hover(cn, lb, cl, c, True))
                w.bind("<Leave>", lambda e, cn=container, lb=lbl, cl=count_lbl, c=fixed: self._sidebar_hover(cn, lb, cl, c, False))

            self._sidebar_buttons[fixed] = (container, lbl, count_lbl)

        # Separator before categories
        cat_sep = tk.Frame(self.sidebar, bg=THEME["sidebar_sep"], height=1)
        cat_sep.pack(fill="x", padx=14, pady=6)
        self._sidebar_extras.append(cat_sep)

        # Dynamic categories
        categories = self._get_categories()
        for cat in categories:
            count = sum(1 for c in self.commands if c.get("category", "") == cat)
            cat_icon = self._get_category_icon(cat)

            container = tk.Frame(self.sidebar, bg=THEME["sidebar_bg"], cursor="hand2")
            container.pack(fill="x")

            display_text = f" {cat_icon}  {cat}" if cat_icon else f"  {cat}"
            lbl = tk.Label(
                container, text=display_text, font=FONTS["sidebar"],
                bg=THEME["sidebar_bg"], fg=THEME["sidebar_fg"], anchor="w"
            )
            lbl.pack(side="left", fill="x", expand=True, padx=(14, 0), pady=7)

            count_lbl = tk.Label(
                container, text=str(count), font=FONTS["badge"],
                bg=THEME["sidebar_bg"], fg=THEME["text_muted"], anchor="e"
            )
            count_lbl.pack(side="right", padx=(0, 14), pady=7)

            for w in (container, lbl, count_lbl):
                w.bind("<Button-1>", lambda e, c=cat: self._select_category(c))
                w.bind("<Enter>", lambda e, cn=container, lb=lbl, cl=count_lbl, c=cat: self._sidebar_hover(cn, lb, cl, c, True))
                w.bind("<Leave>", lambda e, cn=container, lb=lbl, cl=count_lbl, c=cat: self._sidebar_hover(cn, lb, cl, c, False))

            self._sidebar_buttons[cat] = (container, lbl, count_lbl)

        # Spacer to push "Add Category" to bottom
        spacer = tk.Frame(self.sidebar, bg=THEME["sidebar_bg"])
        spacer.pack(fill="both", expand=True)
        self._sidebar_extras.append(spacer)

        # "Add Category" button at bottom
        add_cat_sep = tk.Frame(self.sidebar, bg=THEME["sidebar_sep"], height=1)
        add_cat_sep.pack(fill="x", padx=14, pady=(6, 0))
        self._sidebar_extras.append(add_cat_sep)

        add_cat_btn = tk.Frame(self.sidebar, bg=THEME["sidebar_bg"], cursor="hand2")
        add_cat_btn.pack(fill="x", pady=(0, 8))
        self._sidebar_extras.append(add_cat_btn)

        add_cat_lbl = tk.Label(
            add_cat_btn, text=" +  Add Category", font=FONTS["sidebar"],
            bg=THEME["sidebar_bg"], fg=THEME["accent"], anchor="w", cursor="hand2"
        )
        add_cat_lbl.pack(fill="x", padx=14, pady=8)

        for w in (add_cat_btn, add_cat_lbl):
            w.bind("<Button-1>", lambda e: self.add_category())
            w.bind("<Enter>", lambda e: add_cat_lbl.configure(bg=THEME["sidebar_active"]) or add_cat_btn.configure(bg=THEME["sidebar_active"]))
            w.bind("<Leave>", lambda e: add_cat_lbl.configure(bg=THEME["sidebar_bg"]) or add_cat_btn.configure(bg=THEME["sidebar_bg"]))

        self._highlight_sidebar()

    def _sidebar_hover(self, container, lbl, count_lbl, cat, entering):
        # Don't hover the active item
        is_active = (cat == "Search" and self._search_mode) or (cat != "Search" and cat == self._active_category and not self._search_mode)
        if is_active:
            return
        bg = THEME["sidebar_active"] if entering else THEME["sidebar_bg"]
        for w in (container, lbl, count_lbl):
            if w is not None:
                w.configure(bg=bg)

    def _select_category(self, cat):
        self._search_mode = False
        self._active_category = cat
        # Hide search bar and clear search text
        self.search_frame.pack_forget()
        self.search_var.set("")
        self._highlight_sidebar()
        self.refresh_buttons()

    def _activate_search(self):
        self._search_mode = True
        self._highlight_sidebar()
        # Show the search bar inline
        self.search_frame.pack(fill="x", padx=20, pady=(14, 6), before=self.canvas)
        self.search_entry.focus_set()
        self.search_entry.select_range(0, tk.END)

    def _exit_search(self):
        if not self._search_mode:
            return
        self._search_mode = False
        self.search_frame.pack_forget()
        self.search_var.set("")
        self._highlight_sidebar()
        self.refresh_buttons()

    def _highlight_sidebar(self):
        for name, widgets in self._sidebar_buttons.items():
            container, lbl, count_lbl = widgets
            is_active = (name == "Search" and self._search_mode) or (name != "Search" and name == self._active_category and not self._search_mode)
            if is_active:
                container.configure(bg=THEME["accent_light"])
                lbl.configure(bg=THEME["accent_light"], fg=THEME["accent"])
                if count_lbl:
                    count_lbl.configure(bg=THEME["accent_light"], fg=THEME["accent"])
            else:
                container.configure(bg=THEME["sidebar_bg"])
                lbl.configure(bg=THEME["sidebar_bg"], fg=THEME["sidebar_fg"])
                if count_lbl:
                    count_lbl.configure(bg=THEME["sidebar_bg"], fg=THEME["text_muted"])

    # -----------------------------------------------------------------------
    # Menubar
    # -----------------------------------------------------------------------
    def create_menubar(self):
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Command", accelerator="Ctrl+N", command=self.add_new_command)
        file_menu.add_separator()
        file_menu.add_command(label="Import Commands...", command=self.import_commands)
        file_menu.add_command(label="Export Commands...", command=self.export_commands)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", accelerator="Alt+F4", command=self._on_close)
        menubar.add_cascade(label="File", menu=file_menu)

        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Find", accelerator="Ctrl+F", command=self.focus_search)
        edit_menu.add_separator()
        edit_menu.add_command(label="Delete All Commands", command=self.delete_all_commands)
        menubar.add_cascade(label="Edit", menu=edit_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        self._sidebar_visible = tk.BooleanVar(value=True)
        view_menu.add_checkbutton(label="Sidebar", variable=self._sidebar_visible, command=self.toggle_sidebar)
        self._output_toggle_var = tk.BooleanVar(value=False)
        view_menu.add_checkbutton(label="Output Panel", variable=self._output_toggle_var, command=self.toggle_output)
        view_menu.add_separator()
        view_menu.add_command(label="Refresh", accelerator="F5", command=self.refresh_all)
        menubar.add_cascade(label="View", menu=view_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Keyboard Shortcuts", command=self.show_shortcuts)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)

        self.root.bind("<Control-n>", lambda e: self.add_new_command())
        self.root.bind("<Control-N>", lambda e: self.add_new_command())
        self.root.bind("<Control-f>", lambda e: self.focus_search())
        self.root.bind("<Control-F>", lambda e: self.focus_search())
        self.root.bind("<F5>", lambda e: self.refresh_all())

    # -----------------------------------------------------------------------
    # Data I/O
    # -----------------------------------------------------------------------
    def load_commands(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r") as f:
                    data = json.load(f)
                if isinstance(data, list):
                    return data
            except (json.JSONDecodeError, IOError):
                messagebox.showwarning("Warning", "commands.json is corrupted. Starting with empty list.")
        return []

    def save_commands(self):
        with open(DATA_FILE, "w") as f:
            json.dump(self.commands, f)

    def import_commands(self):
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "r") as f:
                data = json.load(f)
            if not isinstance(data, list):
                messagebox.showerror("Import Error", "File does not contain a valid command list.")
                return
            count = 0
            for item in data:
                if isinstance(item, dict) and "name" in item and "cmd" in item:
                    self.commands.append(item)
                    count += 1
            if count:
                self.save_commands()
                self._rebuild()
            self._toast(f"Imported {count} command(s)")
        except (json.JSONDecodeError, IOError) as e:
            messagebox.showerror("Import Error", f"Could not read file:\n{e}")

    def export_commands(self):
        if not self.commands:
            self._toast("No commands to export")
            return
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "w") as f:
                json.dump(self.commands, f, indent=2)
            self._toast(f"Exported {len(self.commands)} command(s)")
        except IOError as e:
            messagebox.showerror("Export Error", f"Could not write file:\n{e}")

    # -----------------------------------------------------------------------
    # Actions
    # -----------------------------------------------------------------------
    def focus_search(self):
        self._activate_search()

    def delete_all_commands(self):
        if not self.commands:
            self._toast("No commands to delete")
            return
        if messagebox.askyesno("Confirm", f"Delete all {len(self.commands)} command(s)? This cannot be undone."):
            self.commands.clear()
            self.save_commands()
            self._rebuild()
            self._toast("All commands deleted")

    def toggle_sidebar(self):
        if self._sidebar_visible.get():
            self.sidebar.pack(side="left", fill="y", before=self.content_area)
        else:
            self.sidebar.pack_forget()

    def toggle_output(self):
        self._output_visible = self._output_toggle_var.get()
        if self._output_visible:
            self.output_frame.pack(side="bottom", fill="x", before=self.status_bar)
        else:
            self.output_frame.pack_forget()

    def refresh_all(self):
        self.commands = self.load_commands()
        self.custom_categories = self._load_custom_categories()
        self._rebuild()
        self._toast("Refreshed")

    def _rebuild(self):
        self._build_sidebar()
        self._build_cards()
        self.refresh_buttons(self.search_var.get().lower())
        self._update_status()

    def show_shortcuts(self):
        shortcuts = (
            "Ctrl+N\t\tNew Command\n"
            "Ctrl+F\t\tFind / Focus Search\n"
            "F5\t\tRefresh\n"
            "Escape\t\tClose Dialog\n"
            "Alt+F4\t\tExit\n\n"
            "Right-click any card for more options."
        )
        messagebox.showinfo("Keyboard Shortcuts", shortcuts)

    def show_about(self):
        messagebox.showinfo(
            "About",
            "PowerShell Command Runner\n\n"
            "A desktop utility to organise and launch\n"
            "PowerShell commands with one click.\n\n"
            "Tip: Right-click cards to Run, Edit,\n"
            "Duplicate, or Delete."
        )

    # -----------------------------------------------------------------------
    # Add Category dialog
    # -----------------------------------------------------------------------
    CATEGORY_ICONS = [
        "\U0001F4C1", "\U0001F4C2", "\U0001F5C2", "\U0001F4BB", "\U0001F5A5",
        "\U0001F310", "\U0001F512", "\U0001F513", "\U0001F527", "\U0001F6E0",
        "\u2699", "\U0001F4E6", "\U0001F4CA", "\U0001F4C8", "\U0001F4DD",
        "\U0001F3AE", "\U0001F3B5", "\U0001F4F7", "\U0001F4F1", "\U0001F4E7",
        "\u2B50", "\U0001F525", "\U0001F680", "\U0001F4A1", "\U0001F50C",
        "\U0001F4BE", "\U0001F4BF", "\U0001F4C5", "\U0001F3E0", "\U0001F6A9",
    ]

    def add_category(self):
        dlg = self._themed_dialog("Add Category", width=380, height=340)

        self._themed_label(dlg, "Category Name:", 0)
        name_var = tk.StringVar()
        name_entry = self._themed_entry(dlg, name_var, 0)
        name_entry.focus_set()

        self._themed_label(dlg, "Select Icon:", 1)

        # Icon selector grid
        icon_frame = tk.Frame(dlg, bg=THEME["bg"])
        icon_frame.grid(row=2, column=0, columnspan=2, padx=16, pady=8, sticky="ew")

        selected_icon = tk.StringVar(value="")
        icon_labels = []

        def select_icon(icon, lbl):
            selected_icon.set(icon)
            for il in icon_labels:
                il.configure(bg=THEME["bg"], relief="flat")
            lbl.configure(bg=THEME["accent_light"], relief="solid")

        col = 0
        row = 0
        for icon in self.CATEGORY_ICONS:
            lbl = tk.Label(
                icon_frame, text=icon, font=(FONT_FAMILY, 14),
                bg=THEME["bg"], fg=THEME["text_primary"],
                width=3, height=1, cursor="hand2", relief="flat", bd=1
            )
            lbl.grid(row=row, column=col, padx=2, pady=2)
            lbl.bind("<Button-1>", lambda e, ic=icon, lb=lbl: select_icon(ic, lb))
            icon_labels.append(lbl)
            col += 1
            if col >= 10:
                col = 0
                row += 1

        btn_frame = tk.Frame(dlg, bg=THEME["bg"])
        btn_frame.grid(row=3, column=0, columnspan=2, pady=16)

        def create(event=None):
            cat_name = name_var.get().strip()
            if not cat_name:
                messagebox.showwarning("Warning", "Category name cannot be empty.", parent=dlg)
                return
            # Check for duplicate
            existing = {c["name"] for c in self.custom_categories}
            if cat_name in existing:
                messagebox.showwarning("Warning", f"Category '{cat_name}' already exists.", parent=dlg)
                return
            new_cat = {"name": cat_name}
            icon = selected_icon.get()
            if icon:
                new_cat["icon"] = icon
            self.custom_categories.append(new_cat)
            self._save_custom_categories()
            self._rebuild()
            self._toast(f"Category '{cat_name}' created")
            dlg.destroy()

        dlg.bind("<Return>", create)
        self._themed_button(btn_frame, "Create", create, "accent").pack(side="left", padx=8)
        self._themed_button(btn_frame, "Cancel", dlg.destroy).pack(side="left", padx=8)

    # -----------------------------------------------------------------------
    # Status bar
    # -----------------------------------------------------------------------
    def _update_status(self):
        total = len(self.commands)
        shown = self._visible_count
        cat = self._active_category
        if self._search_mode:
            text = f"Showing {shown} of {total} command(s)  —  Search"
        elif shown == total:
            text = f"{total} command(s)"
        else:
            text = f"Showing {shown} of {total} command(s)"
        if cat not in ("Home", "Search") and not self._search_mode:
            text += f"  —  {cat}"
        self._status_left.config(text=text)

    # -----------------------------------------------------------------------
    # Toast
    # -----------------------------------------------------------------------
    def _toast(self, message, duration=2500):
        if self._toast_after_id:
            self.root.after_cancel(self._toast_after_id)
            self._toast_label.place_forget()
        self._toast_label.config(text=f"  {message}  ")
        self._toast_label.place(relx=0.5, rely=0.93, anchor="center")
        self._toast_label.lift()
        self._toast_after_id = self.root.after(duration, self._hide_toast)

    def _hide_toast(self):
        self._toast_label.place_forget()
        self._toast_after_id = None

    # -----------------------------------------------------------------------
    # Dialog helpers
    # -----------------------------------------------------------------------
    def _themed_dialog(self, title, width=520, height=360):
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.resizable(False, False)
        dlg.configure(bg=THEME["bg"])
        dlg.grab_set()
        dlg.bind("<Escape>", lambda e: dlg.destroy())

        # Center on parent
        self.root.update_idletasks()
        px = self.root.winfo_rootx() + (self.root.winfo_width() - width) // 2
        py = self.root.winfo_rooty() + (self.root.winfo_height() - height) // 2
        dlg.geometry(f"{width}x{height}+{px}+{py}")
        return dlg

    def _themed_label(self, parent, text, row, col=0, **kw):
        lbl = tk.Label(parent, text=text, font=FONTS["body"], bg=THEME["bg"], fg=THEME["text_primary"])
        lbl.grid(row=row, column=col, sticky="w", padx=16, pady=(12 if row == 0 else 5, 2), **kw)
        return lbl

    def _themed_entry(self, parent, var, row, col=1, width=34, **kw):
        ent = tk.Entry(
            parent, textvariable=var, width=width, font=FONTS["body"],
            bg=THEME["search_bg"], fg=THEME["text_primary"], relief="solid", bd=1,
            insertbackground=THEME["text_primary"]
        )
        ent.grid(row=row, column=col, padx=16, pady=(12 if row == 0 else 5, 2), sticky="ew", **kw)
        return ent

    def _themed_button(self, parent, text, command, style="normal"):
        colors = {
            "normal":  (THEME["btn_bg"], THEME["btn_fg"]),
            "accent":  (THEME["accent"], THEME["accent_fg"]),
            "danger":  (THEME["btn_danger"], THEME["accent_fg"]),
        }
        bg, fg = colors.get(style, colors["normal"])
        hover_colors = {
            "normal":  THEME["sidebar_active"],
            "accent":  THEME["accent_hover"],
            "danger":  THEME["btn_danger_hover"],
        }
        hover_bg = hover_colors.get(style, THEME["sidebar_active"])

        btn = tk.Button(
            parent, text=text, command=command, font=FONTS["card_btn"],
            bg=bg, fg=fg, activebackground=hover_bg, activeforeground=fg,
            relief="flat", padx=16, pady=5, cursor="hand2", bd=0
        )
        btn.bind("<Enter>", lambda e, b=btn, h=hover_bg: b.configure(bg=h))
        btn.bind("<Leave>", lambda e, b=btn, o=bg: b.configure(bg=o))
        return btn

    # -----------------------------------------------------------------------
    # Add command dialog
    # -----------------------------------------------------------------------
    def add_new_command(self):
        dlg = self._themed_dialog("New Command", height=460)
        dlg.columnconfigure(1, weight=1)

        self._themed_label(dlg, "Name:", 0)
        name_var = tk.StringVar()
        name_entry = self._themed_entry(dlg, name_var, 0)
        name_entry.focus_set()

        self._themed_label(dlg, "Command:", 1)
        cmd_var = tk.StringVar()
        self._themed_entry(dlg, cmd_var, 1)

        self._themed_label(dlg, "Image:", 2)
        img_var = tk.StringVar()
        self._themed_entry(dlg, img_var, 2, width=22)

        self._themed_label(dlg, "Category:", 3)
        cat_var = tk.StringVar()

        # Scrollable category selector
        cat_frame = tk.Frame(dlg, bg=THEME["bg"])
        cat_frame.grid(row=3, column=1, padx=16, pady=5, sticky="ew")

        cat_canvas = tk.Canvas(cat_frame, bg=THEME["card_bg"], highlightthickness=1,
                               highlightbackground=THEME["card_border"], height=70)
        cat_scrollbar = ttk.Scrollbar(cat_frame, orient="horizontal", command=cat_canvas.xview)
        cat_inner = tk.Frame(cat_canvas, bg=THEME["card_bg"])

        cat_canvas.configure(xscrollcommand=cat_scrollbar.set)
        cat_scrollbar.pack(side="bottom", fill="x")
        cat_canvas.pack(side="top", fill="x", expand=True)
        cat_canvas.create_window((0, 0), window=cat_inner, anchor="nw")
        cat_inner.bind("<Configure>", lambda e: cat_canvas.configure(scrollregion=cat_canvas.bbox("all")))

        cat_pill_labels = []
        categories = self._get_categories()

        def select_cat(cat_name):
            if cat_var.get() == cat_name:
                cat_var.set("")
            else:
                cat_var.set(cat_name)
            _refresh_pills()

        def _refresh_pills():
            sel = cat_var.get()
            for pill_lbl, pill_cat in cat_pill_labels:
                if pill_cat == sel:
                    pill_lbl.configure(bg=THEME["accent"], fg=THEME["accent_fg"])
                else:
                    pill_lbl.configure(bg=THEME["badge_bg"], fg=THEME["badge_fg"])

        for cat_name in categories:
            icon = self._get_category_icon(cat_name)
            pill_text = f" {icon} {cat_name} " if icon else f" {cat_name} "
            pill = tk.Label(
                cat_inner, text=pill_text, font=FONTS["badge"],
                bg=THEME["badge_bg"], fg=THEME["badge_fg"],
                padx=8, pady=4, cursor="hand2", relief="flat", bd=0
            )
            pill.pack(side="left", padx=3, pady=6)
            pill.bind("<Button-1>", lambda e, cn=cat_name: select_cat(cn))
            cat_pill_labels.append((pill, cat_name))

        _refresh_pills()

        def browse_image():
            path = filedialog.askopenfilename(parent=dlg, filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All files", "*.*")])
            if path:
                img_var.set(path)

        browse_btn = self._themed_button(dlg, "Browse...", browse_image)
        browse_btn.grid(row=2, column=2, padx=(0, 16), pady=5)

        btn_frame = tk.Frame(dlg, bg=THEME["bg"])
        btn_frame.grid(row=5, column=0, columnspan=3, pady=20)

        def create(event=None):
            new_name = name_var.get().strip()
            new_cmd = cmd_var.get().strip()
            if not new_name or not new_cmd:
                messagebox.showwarning("Warning", "Name and command cannot be empty.", parent=dlg)
                return
            new_item = {"name": new_name, "cmd": new_cmd}
            img_path = img_var.get().strip()
            if img_path:
                new_item["image"] = img_path
            cat = cat_var.get().strip()
            if cat:
                new_item["category"] = cat
            self.commands.append(new_item)
            self.save_commands()
            self._rebuild()
            self._toast(f"Created '{new_name}'")
            dlg.destroy()

        dlg.bind("<Return>", create)
        self._themed_button(btn_frame, "Create", create, "accent").pack(side="left", padx=8)
        self._themed_button(btn_frame, "Cancel", dlg.destroy).pack(side="left", padx=8)

    # -----------------------------------------------------------------------
    # Edit command dialog
    # -----------------------------------------------------------------------
    def edit_command(self, idx):
        item = self.commands[idx]
        dlg = self._themed_dialog(f"Edit — {item['name']}", height=460)
        dlg.columnconfigure(1, weight=1)

        self._themed_label(dlg, "Name:", 0)
        name_var = tk.StringVar(value=item["name"])
        self._themed_entry(dlg, name_var, 0)

        self._themed_label(dlg, "Command:", 1)
        cmd_var = tk.StringVar(value=item["cmd"])
        self._themed_entry(dlg, cmd_var, 1)

        self._themed_label(dlg, "Image:", 2)
        img_var = tk.StringVar(value=item.get("image", ""))
        self._themed_entry(dlg, img_var, 2, width=22)

        self._themed_label(dlg, "Category:", 3)
        cat_var = tk.StringVar(value=item.get("category", ""))

        # Scrollable category selector
        cat_frame = tk.Frame(dlg, bg=THEME["bg"])
        cat_frame.grid(row=3, column=1, padx=16, pady=5, sticky="ew")

        cat_canvas = tk.Canvas(cat_frame, bg=THEME["card_bg"], highlightthickness=1,
                               highlightbackground=THEME["card_border"], height=70)
        cat_scrollbar = ttk.Scrollbar(cat_frame, orient="horizontal", command=cat_canvas.xview)
        cat_inner = tk.Frame(cat_canvas, bg=THEME["card_bg"])

        cat_canvas.configure(xscrollcommand=cat_scrollbar.set)
        cat_scrollbar.pack(side="bottom", fill="x")
        cat_canvas.pack(side="top", fill="x", expand=True)
        cat_canvas_win = cat_canvas.create_window((0, 0), window=cat_inner, anchor="nw")
        cat_inner.bind("<Configure>", lambda e: cat_canvas.configure(scrollregion=cat_canvas.bbox("all")))

        cat_pill_labels = []
        categories = self._get_categories()
        current_cat = item.get("category", "")

        def select_cat(cat_name):
            if cat_var.get() == cat_name:
                cat_var.set("")
            else:
                cat_var.set(cat_name)
            _refresh_pills()

        def _refresh_pills():
            sel = cat_var.get()
            for pill_lbl, pill_cat in cat_pill_labels:
                if pill_cat == sel:
                    pill_lbl.configure(bg=THEME["accent"], fg=THEME["accent_fg"])
                else:
                    pill_lbl.configure(bg=THEME["badge_bg"], fg=THEME["badge_fg"])

        for cat_name in categories:
            icon = self._get_category_icon(cat_name)
            pill_text = f" {icon} {cat_name} " if icon else f" {cat_name} "
            pill = tk.Label(
                cat_inner, text=pill_text, font=FONTS["badge"],
                bg=THEME["badge_bg"], fg=THEME["badge_fg"],
                padx=8, pady=4, cursor="hand2", relief="flat", bd=0
            )
            pill.pack(side="left", padx=3, pady=6)
            pill.bind("<Button-1>", lambda e, cn=cat_name: select_cat(cn))
            cat_pill_labels.append((pill, cat_name))

        _refresh_pills()

        def browse_image():
            path = filedialog.askopenfilename(parent=dlg, filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All files", "*.*")])
            if path:
                img_var.set(path)

        browse_btn = self._themed_button(dlg, "Browse...", browse_image)
        browse_btn.grid(row=2, column=2, padx=(0, 16), pady=5)

        btn_frame = tk.Frame(dlg, bg=THEME["bg"])
        btn_frame.grid(row=5, column=0, columnspan=3, pady=20)

        def save(event=None):
            new_name = name_var.get().strip()
            new_cmd = cmd_var.get().strip()
            if not new_name or not new_cmd:
                messagebox.showwarning("Warning", "Name and command cannot be empty.", parent=dlg)
                return
            item["name"] = new_name
            item["cmd"] = new_cmd
            img_path = img_var.get().strip()
            if img_path:
                item["image"] = img_path
            elif "image" in item:
                del item["image"]
            cat = cat_var.get().strip()
            if cat:
                item["category"] = cat
            elif "category" in item:
                del item["category"]
            self.save_commands()
            self._rebuild()
            self._toast(f"Saved '{new_name}'")
            dlg.destroy()

        def delete():
            if messagebox.askyesno("Confirm Delete", f"Delete '{item['name']}'?", parent=dlg):
                self.commands.pop(idx)
                self.save_commands()
                self._rebuild()
                self._toast("Deleted command")
                dlg.destroy()

        dlg.bind("<Return>", save)
        self._themed_button(btn_frame, "Save", save, "accent").pack(side="left", padx=8)
        self._themed_button(btn_frame, "Delete", delete, "danger").pack(side="left", padx=8)
        self._themed_button(btn_frame, "Cancel", dlg.destroy).pack(side="left", padx=8)

    # -----------------------------------------------------------------------
    # Run command
    # -----------------------------------------------------------------------
    def run_powershell(self, command):
        if not messagebox.askyesno("Confirm", f"Run this command?\n\n{command}"):
            return
        self._toast("Running...")
        try:
            proc = subprocess.Popen(
                ["powershell", "-Command", command],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT
            )
            self.root.after(100, lambda: self._poll_output(proc))
        except FileNotFoundError:
            messagebox.showerror("Error", "PowerShell not found. Is it installed and on PATH?")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _poll_output(self, proc):
        ret = proc.poll()
        chunk = b""
        if proc.stdout:
            try:
                # Non-blocking: read whatever is available
                if hasattr(proc.stdout, "read1"):
                    chunk = proc.stdout.read1(4096)
                else:
                    chunk = proc.stdout.read(4096) if ret is not None else b""
            except Exception:
                pass
        if chunk:
            text = chunk.decode("utf-8", errors="replace")
            self.output_text.configure(state="normal")
            self.output_text.insert("end", text)
            self.output_text.see("end")
            self.output_text.configure(state="disabled")
            if not self._output_visible:
                self._output_toggle_var.set(True)
                self.toggle_output()
        if ret is None:
            self.root.after(150, lambda: self._poll_output(proc))
        else:
            # Read any remaining
            if proc.stdout:
                try:
                    remaining = proc.stdout.read()
                    if remaining:
                        text = remaining.decode("utf-8", errors="replace")
                        self.output_text.configure(state="normal")
                        self.output_text.insert("end", text)
                        self.output_text.configure(state="disabled")
                except Exception:
                    pass
            self.output_text.configure(state="normal")
            self.output_text.insert("end", f"\n{'─' * 40}\nProcess exited with code {ret}\n\n")
            self.output_text.see("end")
            self.output_text.configure(state="disabled")
            self._toast("Command finished")

    def _clear_output(self):
        self.output_text.configure(state="normal")
        self.output_text.delete("1.0", "end")
        self.output_text.configure(state="disabled")

    # -----------------------------------------------------------------------
    # Right-click context menu
    # -----------------------------------------------------------------------
    def _show_card_menu(self, event, idx):
        menu = tk.Menu(self.root, tearoff=0, font=FONTS["small"])
        item = self.commands[idx]
        menu.add_command(label="  Run", command=lambda: self.run_powershell(item["cmd"]))
        menu.add_command(label="  Edit", command=lambda: self.edit_command(idx))
        menu.add_command(label="  Duplicate", command=lambda: self._duplicate_command(idx))
        menu.add_separator()
        menu.add_command(label="  Delete", command=lambda: self._delete_command(idx))
        menu.tk_popup(event.x_root, event.y_root)

    def _duplicate_command(self, idx):
        item = self.commands[idx]
        copy = dict(item)
        copy["name"] = item["name"] + " (copy)"
        self.commands.insert(idx + 1, copy)
        self.save_commands()
        self._rebuild()
        self._toast(f"Duplicated '{item['name']}'")

    def _delete_command(self, idx):
        name = self.commands[idx]["name"]
        if messagebox.askyesno("Confirm Delete", f"Delete '{name}'?"):
            self.commands.pop(idx)
            self.save_commands()
            self._rebuild()
            self._toast(f"Deleted '{name}'")

    # -----------------------------------------------------------------------
    # Search
    # -----------------------------------------------------------------------
    def update_list(self, *args):
        if self._search_after_id is not None:
            self.root.after_cancel(self._search_after_id)
        self._search_after_id = self.root.after(200, self._do_search)

    def _do_search(self):
        self._search_after_id = None
        self.refresh_buttons(self.search_var.get().lower())

    # -----------------------------------------------------------------------
    # Image loading
    # -----------------------------------------------------------------------
    def _load_thumbnail(self, path):
        if not HAS_PIL or not path or not os.path.isfile(path):
            return None
        try:
            img = Image.open(path)
            # Crop to aspect ratio then resize for uniform cards
            target_w, target_h = CARD_IMAGE_SIZE
            img_ratio = img.width / img.height
            target_ratio = target_w / target_h
            if img_ratio > target_ratio:
                new_h = img.height
                new_w = int(new_h * target_ratio)
                left = (img.width - new_w) // 2
                img = img.crop((left, 0, left + new_w, new_h))
            else:
                new_w = img.width
                new_h = int(new_w / target_ratio)
                top = (img.height - new_h) // 2
                img = img.crop((0, top, new_w, top + new_h))
            img = img.resize(CARD_IMAGE_SIZE, Image.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception:
            return None

    # -----------------------------------------------------------------------
    # Card building
    # -----------------------------------------------------------------------
    def _build_cards(self):
        for frame, _ in self._card_widgets:
            frame.destroy()
        self._card_widgets = []
        self._image_refs = []
        self._card_tooltips = []

        for idx, item in enumerate(self.commands):
            card = tk.Frame(
                self.buttons_frame, bg=THEME["card_bg"], bd=0,
                highlightthickness=1, highlightbackground=THEME["card_border"],
                padx=12, pady=10
            )

            # Collect non-button children to recolor on hover
            hover_targets = []

            # Image
            photo = self._load_thumbnail(item.get("image"))
            if photo:
                self._image_refs.append(photo)
                lbl_img = tk.Label(card, image=photo, bg=THEME["card_bg"], bd=0)
            else:
                lbl_img = tk.Label(
                    card, text="\U0001F4BB", font=(FONT_FAMILY, 28),
                    bg=THEME["accent_light"], fg=THEME["accent"],
                    width=12, height=3
                )
            lbl_img.pack(pady=(0, 8), fill="x")
            lbl_img.bind("<Button-3>", lambda e, i=idx: self._show_card_menu(e, i))
            hover_targets.append(lbl_img)

            # Name (truncated if long)
            display_name = item["name"]
            if len(display_name) > CARD_MAX_NAME:
                display_name = display_name[:CARD_MAX_NAME - 1] + "\u2026"
            name_lbl = tk.Label(card, text=display_name, font=FONTS["card_name"], bg=THEME["card_bg"], fg=THEME["text_primary"])
            name_lbl.pack()
            hover_targets.append(name_lbl)

            # Command preview
            cmd_preview = item["cmd"]
            if len(cmd_preview) > 35:
                cmd_preview = cmd_preview[:34] + "\u2026"
            cmd_lbl = tk.Label(card, text=cmd_preview, font=FONTS["card_cmd"], bg=THEME["card_bg"], fg=THEME["text_muted"])
            cmd_lbl.pack(pady=(1, 3))
            hover_targets.append(cmd_lbl)

            # Category badge
            cat = item.get("category", "")
            if cat:
                cat_lbl = tk.Label(
                    card, text=f" {cat} ", font=FONTS["badge"],
                    bg=THEME["badge_bg"], fg=THEME["badge_fg"], padx=6, pady=1
                )
                cat_lbl.pack(pady=(0, 4))

            # Buttons row
            btn_row = tk.Frame(card, bg=THEME["card_bg"])
            btn_row.pack(pady=(6, 0))
            hover_targets.append(btn_row)

            run_btn = self._themed_button(btn_row, "\u25B6  Run", lambda c=item["cmd"]: self.run_powershell(c), "accent")
            run_btn.pack(side="left", padx=3)

            edit_btn = self._themed_button(btn_row, "Edit", lambda i=idx: self.edit_command(i))
            edit_btn.pack(side="left", padx=3)

            # Hover effect — only recolor non-button labels
            def on_enter(e, c=card, targets=hover_targets):
                c.configure(highlightbackground=THEME["card_hover_border"], bg=THEME["card_hover"])
                for t in targets:
                    try:
                        if not isinstance(t, tk.Button):
                            t.configure(bg=THEME["card_hover"])
                    except tk.TclError:
                        pass

            def on_leave(e, c=card, targets=hover_targets):
                c.configure(highlightbackground=THEME["card_border"], bg=THEME["card_bg"])
                for t in targets:
                    try:
                        if not isinstance(t, tk.Button):
                            t.configure(bg=THEME["card_bg"])
                    except tk.TclError:
                        pass

            card.bind("<Enter>", on_enter)
            card.bind("<Leave>", on_leave)
            card.bind("<Button-3>", lambda e, i=idx: self._show_card_menu(e, i))

            # Drag-and-drop
            card.bind("<ButtonPress-1>", lambda e, i=idx: self._drag_start(i))
            card.bind("<B1-Motion>", self._drag_motion)
            card.bind("<ButtonRelease-1>", self._drag_end)

            # Tooltip with full command
            tt = ToolTip(card, item["cmd"])
            self._card_tooltips.append(tt)

            self._card_widgets.append((card, idx))

    # -----------------------------------------------------------------------
    # Card grid layout
    # -----------------------------------------------------------------------
    def refresh_buttons(self, filter_text=""):
        self._empty_frame.grid_forget()
        for frame, _ in self._card_widgets:
            frame.grid_forget()

        row = 0
        col = 0
        count = 0
        for frame, idx in self._card_widgets:
            item = self.commands[idx]
            if self._active_category not in ("Home", "Search"):
                if item.get("category", "") != self._active_category:
                    continue
            if filter_text and filter_text not in item["name"].lower():
                continue

            frame.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")
            self.buttons_frame.columnconfigure(col, weight=1)
            col += 1
            count += 1
            if col >= COLUMNS:
                col = 0
                row += 1

        self._visible_count = count

        if count == 0:
            self._empty_frame.grid(row=0, column=0, columnspan=COLUMNS, sticky="nsew")

        self._update_status()

    # -----------------------------------------------------------------------
    # Drag and drop
    # -----------------------------------------------------------------------
    def _drag_start(self, idx):
        self._drag_data["idx"] = idx

    def _drag_motion(self, event):
        pass

    def _drag_end(self, event):
        src = self._drag_data["idx"]
        if src is None:
            return
        widget = event.widget.winfo_containing(event.x_root, event.y_root)
        target_idx = None
        for frame, idx in self._card_widgets:
            if widget is frame or (widget is not None and str(widget).startswith(str(frame))):
                target_idx = idx
                break
        if target_idx is not None and target_idx != src:
            item = self.commands.pop(src)
            self.commands.insert(target_idx, item)
            self.save_commands()
            self._build_cards()
            self.refresh_buttons(self.search_var.get().lower())
            self._toast("Reordered")
        self._drag_data["idx"] = None

    # -----------------------------------------------------------------------
    # System tray
    # -----------------------------------------------------------------------
    def _setup_tray(self):
        try:
            img = Image.new("RGB", (64, 64), THEME["accent"])
            menu = TrayMenu(
                TrayMenuItem("Show", self._tray_show),
                TrayMenuItem("Exit", self._tray_exit),
            )
            self._tray_icon = TrayIcon("PSRunner", img, "PowerShell Command Runner", menu)
            threading.Thread(target=self._tray_icon.run, daemon=True).start()
        except Exception:
            self._tray_icon = None

    def _tray_show(self, icon=None, menu_item=None):
        self.root.after(0, self.root.deiconify)

    def _tray_exit(self, icon=None, menu_item=None):
        if self._tray_icon:
            self._tray_icon.stop()
        self.root.after(0, self._on_close)


if __name__ == "__main__":
    root = tk.Tk()
    app = PowerShellApp(root)
    root.mainloop()
