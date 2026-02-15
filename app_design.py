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
    "bg":              "#F0F4F8",   # main background – very light blue-grey
    "sidebar_bg":      "#E2E8F0",   # sidebar – soft grey
    "sidebar_active":  "#CBD5E1",   # sidebar hover / selected
    "sidebar_fg":      "#334155",   # sidebar text
    "content_bg":      "#FFFFFF",   # content pane
    "card_bg":         "#FFFFFF",   # card background
    "card_border":     "#CBD5E1",   # card border
    "card_hover":      "#EBF2FA",   # card hover tint
    "accent":          "#3B82F6",   # primary accent – light blue
    "accent_hover":    "#2563EB",   # accent hover – darker blue
    "accent_fg":       "#FFFFFF",   # text on accent
    "btn_bg":          "#E2E8F0",   # secondary button bg
    "btn_fg":          "#1E293B",   # secondary button text
    "btn_danger":      "#EF4444",   # danger (delete) red
    "btn_danger_hover":"#DC2626",
    "text_primary":    "#1E293B",   # main text
    "text_secondary":  "#64748B",   # muted text
    "search_bg":       "#F8FAFC",   # search entry bg
    "search_border":   "#CBD5E1",
    "status_bg":       "#E2E8F0",   # status bar
    "toast_bg":        "#334155",   # toast background
    "toast_fg":        "#F8FAFC",   # toast text
    "output_bg":       "#1E293B",   # output viewer background
    "output_fg":       "#E2E8F0",   # output viewer text
}

FONT_FAMILY = "Segoe UI"
FONTS = {
    "title":     (FONT_FAMILY, 11, "bold"),
    "body":      (FONT_FAMILY, 10),
    "small":     (FONT_FAMILY, 9),
    "card_name": (FONT_FAMILY, 10, "bold"),
    "card_btn":  (FONT_FAMILY, 9),
    "fab":       (FONT_FAMILY, 18, "bold"),
    "status":    (FONT_FAMILY, 9),
    "toast":     (FONT_FAMILY, 10),
    "sidebar":   (FONT_FAMILY, 10),
    "output":    ("Consolas", 9),
}

CARD_IMAGE_SIZE = (100, 100)
COLUMNS = 3


class PowerShellApp:
    # -----------------------------------------------------------------------
    # Init
    # -----------------------------------------------------------------------
    def __init__(self, root):
        self.root = root
        self.root.title("PowerShell Command Runner")
        self.root.configure(bg=THEME["bg"])
        self.root.minsize(700, 450)

        # Restore saved window geometry
        self._load_geometry()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Data
        self.commands = self.load_commands()
        self._drag_data = {"idx": None}

        # Menubar
        self.create_menubar()

        # ---- Main container ----
        main_container = tk.Frame(root, bg=THEME["bg"])
        main_container.pack(fill="both", expand=True)

        # ---- Sidebar ----
        self.sidebar = tk.Frame(main_container, width=170, bg=THEME["sidebar_bg"])
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        sidebar_title = tk.Label(
            self.sidebar, text="Categories", font=FONTS["title"],
            bg=THEME["sidebar_bg"], fg=THEME["sidebar_fg"], anchor="w"
        )
        sidebar_title.pack(fill="x", padx=12, pady=(14, 6))

        self._active_category = "All"
        self._sidebar_buttons = {}
        self._build_sidebar()

        # ---- Content area ----
        self.content_area = tk.Frame(main_container, bg=THEME["content_bg"])
        self.content_area.pack(side="left", fill="both", expand=True)

        # Search bar
        search_frame = tk.Frame(self.content_area, bg=THEME["content_bg"])
        search_frame.pack(fill="x", padx=16, pady=(12, 4))

        tk.Label(
            search_frame, text="Search", font=FONTS["body"],
            bg=THEME["content_bg"], fg=THEME["text_secondary"]
        ).pack(side="left", padx=(0, 8))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.update_list)
        self._search_after_id = None

        self.search_entry = tk.Entry(
            search_frame, textvariable=self.search_var, font=FONTS["body"],
            bg=THEME["search_bg"], fg=THEME["text_primary"],
            relief="solid", bd=1, insertbackground=THEME["text_primary"]
        )
        self.search_entry.pack(side="left", fill="x", expand=True, ipady=4)

        # Scrollable card grid
        self.canvas = tk.Canvas(self.content_area, bg=THEME["content_bg"], highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.content_area, orient="vertical", command=self.canvas.yview)
        self.buttons_frame = tk.Frame(self.canvas, bg=THEME["content_bg"])

        self.buttons_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._canvas_win = self.canvas.create_window((0, 0), window=self.buttons_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        self.canvas.pack(side="left", fill="both", expand=True, padx=(16, 0), pady=8)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        # Card cache
        self._card_widgets = []
        self._image_refs = []

        # ---- FAB ----
        self.fab = tk.Canvas(root, width=52, height=52, bg=THEME["content_bg"], highlightthickness=0, bd=0)
        self.fab.place(relx=0.93, rely=0.88, anchor="center")
        self._draw_fab()
        self.fab.bind("<Button-1>", lambda e: self.add_new_command())
        self.fab.bind("<Enter>", lambda e: self._draw_fab(hover=True))
        self.fab.bind("<Leave>", lambda e: self._draw_fab(hover=False))

        # ---- Status bar ----
        self.status_bar = tk.Label(
            root, text="", font=FONTS["status"], bg=THEME["status_bg"],
            fg=THEME["text_secondary"], anchor="w", padx=12, pady=4
        )
        self.status_bar.pack(side="bottom", fill="x")
        self._update_status()

        # ---- Toast overlay ----
        self._toast_label = tk.Label(
            root, text="", font=FONTS["toast"], bg=THEME["toast_bg"],
            fg=THEME["toast_fg"], padx=16, pady=8
        )
        self._toast_after_id = None

        # ---- Output panel (hidden by default) ----
        self._output_visible = False
        self.output_frame = tk.Frame(root, bg=THEME["output_bg"], height=150)
        self.output_text = tk.Text(
            self.output_frame, bg=THEME["output_bg"], fg=THEME["output_fg"],
            font=FONTS["output"], wrap="word", relief="flat", state="disabled",
            insertbackground=THEME["output_fg"], padx=8, pady=6
        )
        output_scroll = ttk.Scrollbar(self.output_frame, orient="vertical", command=self.output_text.yview)
        self.output_text.configure(yscrollcommand=output_scroll.set)
        self.output_text.pack(side="left", fill="both", expand=True)
        output_scroll.pack(side="right", fill="y")

        # Build cards
        self._build_cards()
        self.refresh_buttons()

        # System tray (optional)
        self._tray_icon = None
        if HAS_TRAY and HAS_PIL:
            self._setup_tray()

    # -----------------------------------------------------------------------
    # FAB drawing
    # -----------------------------------------------------------------------
    def _draw_fab(self, hover=False):
        self.fab.delete("all")
        color = THEME["accent_hover"] if hover else THEME["accent"]
        self.fab.create_oval(2, 2, 50, 50, fill=color, outline=color)
        self.fab.create_text(26, 26, text="+", fill=THEME["accent_fg"], font=FONTS["fab"])

    # -----------------------------------------------------------------------
    # Canvas resize — keep inner frame width in sync
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
        self.root.geometry("900x650")

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
    # Sidebar — categories from command tags
    # -----------------------------------------------------------------------
    def _get_categories(self):
        cats = set()
        for item in self.commands:
            cat = item.get("category", "")
            if cat:
                cats.add(cat)
        return sorted(cats)

    def _build_sidebar(self):
        for w in list(self._sidebar_buttons.values()):
            w.destroy()
        self._sidebar_buttons.clear()

        categories = ["All"] + self._get_categories()
        for cat in categories:
            btn = tk.Label(
                self.sidebar, text=cat, font=FONTS["sidebar"],
                bg=THEME["sidebar_bg"], fg=THEME["sidebar_fg"],
                anchor="w", padx=16, pady=6, cursor="hand2"
            )
            btn.pack(fill="x")
            btn.bind("<Button-1>", lambda e, c=cat: self._select_category(c))
            btn.bind("<Enter>", lambda e, b=btn: b.configure(bg=THEME["sidebar_active"]) if b.cget("bg") != THEME["accent"] else None)
            btn.bind("<Leave>", lambda e, b=btn: b.configure(bg=THEME["sidebar_bg"]) if b.cget("bg") != THEME["accent"] else None)
            self._sidebar_buttons[cat] = btn

        self._highlight_category()

    def _select_category(self, cat):
        self._active_category = cat
        self._highlight_category()
        self.refresh_buttons(self.search_var.get().lower())

    def _highlight_category(self):
        for name, btn in self._sidebar_buttons.items():
            if name == self._active_category:
                btn.configure(bg=THEME["accent"], fg=THEME["accent_fg"])
            else:
                btn.configure(bg=THEME["sidebar_bg"], fg=THEME["sidebar_fg"])

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
        self.search_entry.focus_set()
        self.search_entry.select_range(0, tk.END)

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
        self._rebuild()
        self._toast("Refreshed")

    def _rebuild(self):
        self._build_sidebar()
        self._build_cards()
        self.refresh_buttons(self.search_var.get().lower())
        self._update_status()

    def show_shortcuts(self):
        shortcuts = "Ctrl+N\tNew Command\nCtrl+F\tFind / Focus Search\nF5\tRefresh\nAlt+F4\tExit"
        messagebox.showinfo("Keyboard Shortcuts", shortcuts)

    def show_about(self):
        messagebox.showinfo(
            "About",
            "PowerShell Command Runner\n\n"
            "A desktop utility to save and run\n"
            "PowerShell commands with one click."
        )

    # -----------------------------------------------------------------------
    # Status bar
    # -----------------------------------------------------------------------
    def _update_status(self):
        n = len(self.commands)
        cat = self._active_category
        self.status_bar.config(text=f"{n} command(s)  |  Category: {cat}")

    # -----------------------------------------------------------------------
    # Toast notifications
    # -----------------------------------------------------------------------
    def _toast(self, message, duration=2500):
        if self._toast_after_id:
            self.root.after_cancel(self._toast_after_id)
            self._toast_label.place_forget()
        self._toast_label.config(text=message)
        self._toast_label.place(relx=0.5, rely=0.92, anchor="center")
        self._toast_label.lift()
        self._toast_after_id = self.root.after(duration, self._hide_toast)

    def _hide_toast(self):
        self._toast_label.place_forget()
        self._toast_after_id = None

    # -----------------------------------------------------------------------
    # Dialogs — themed helper
    # -----------------------------------------------------------------------
    def _themed_dialog(self, title, width=440, height=340):
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.geometry(f"{width}x{height}")
        dlg.resizable(False, False)
        dlg.configure(bg=THEME["bg"])
        dlg.grab_set()
        return dlg

    def _themed_label(self, parent, text, row, col=0, **kw):
        lbl = tk.Label(parent, text=text, font=FONTS["body"], bg=THEME["bg"], fg=THEME["text_primary"])
        lbl.grid(row=row, column=col, sticky="w", padx=14, pady=(10 if row == 0 else 4, 2), **kw)
        return lbl

    def _themed_entry(self, parent, var, row, col=1, width=34, **kw):
        ent = tk.Entry(
            parent, textvariable=var, width=width, font=FONTS["body"],
            bg=THEME["search_bg"], fg=THEME["text_primary"], relief="solid", bd=1,
            insertbackground=THEME["text_primary"]
        )
        ent.grid(row=row, column=col, padx=14, pady=(10 if row == 0 else 4, 2), sticky="ew", **kw)
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
            relief="flat", padx=14, pady=4, cursor="hand2", bd=0
        )
        btn.bind("<Enter>", lambda e: btn.configure(bg=hover_bg))
        btn.bind("<Leave>", lambda e: btn.configure(bg=bg))
        return btn

    # -----------------------------------------------------------------------
    # Add new command dialog
    # -----------------------------------------------------------------------
    def add_new_command(self):
        dlg = self._themed_dialog("New Command")

        self._themed_label(dlg, "Name:", 0)
        name_var = tk.StringVar()
        self._themed_entry(dlg, name_var, 0)

        self._themed_label(dlg, "Command:", 1)
        cmd_var = tk.StringVar()
        self._themed_entry(dlg, cmd_var, 1)

        self._themed_label(dlg, "Image:", 2)
        img_var = tk.StringVar()
        self._themed_entry(dlg, img_var, 2, width=22)

        self._themed_label(dlg, "Category:", 3)
        cat_var = tk.StringVar()
        self._themed_entry(dlg, cat_var, 3)

        def browse_image():
            path = filedialog.askopenfilename(parent=dlg, filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All files", "*.*")])
            if path:
                img_var.set(path)

        browse_btn = self._themed_button(dlg, "Browse...", browse_image)
        browse_btn.grid(row=2, column=1, sticky="e", padx=14, pady=4)

        btn_frame = tk.Frame(dlg, bg=THEME["bg"])
        btn_frame.grid(row=5, column=0, columnspan=2, pady=18)

        def create():
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

        self._themed_button(btn_frame, "Create", create, "accent").pack(side="left", padx=8)
        self._themed_button(btn_frame, "Cancel", dlg.destroy).pack(side="left", padx=8)

    # -----------------------------------------------------------------------
    # Edit command dialog
    # -----------------------------------------------------------------------
    def edit_command(self, idx):
        item = self.commands[idx]
        dlg = self._themed_dialog(f"Edit: {item['name']}")

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
        self._themed_entry(dlg, cat_var, 3)

        def browse_image():
            path = filedialog.askopenfilename(parent=dlg, filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All files", "*.*")])
            if path:
                img_var.set(path)

        browse_btn = self._themed_button(dlg, "Browse...", browse_image)
        browse_btn.grid(row=2, column=1, sticky="e", padx=14, pady=4)

        btn_frame = tk.Frame(dlg, bg=THEME["bg"])
        btn_frame.grid(row=5, column=0, columnspan=2, pady=18)

        def save():
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
                self._toast(f"Deleted command")
                dlg.destroy()

        self._themed_button(btn_frame, "Save", save, "accent").pack(side="left", padx=8)
        self._themed_button(btn_frame, "Delete", delete, "danger").pack(side="left", padx=8)
        self._themed_button(btn_frame, "Cancel", dlg.destroy).pack(side="left", padx=8)

    # -----------------------------------------------------------------------
    # Run command
    # -----------------------------------------------------------------------
    def run_powershell(self, command):
        if not messagebox.askyesno("Confirm", f"Run this command?\n\n{command}"):
            return
        self._toast("Running command...")
        try:
            proc = subprocess.Popen(
                ["powershell", "-Command", command],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True
            )
            self.root.after(100, lambda: self._poll_output(proc))
        except FileNotFoundError:
            messagebox.showerror("Error", "PowerShell not found. Is it installed and on PATH?")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _poll_output(self, proc):
        ret = proc.poll()
        try:
            chunk = proc.stdout.read(4096) if proc.stdout else ""
        except Exception:
            chunk = ""
        if chunk:
            self.output_text.configure(state="normal")
            self.output_text.insert("end", chunk)
            self.output_text.see("end")
            self.output_text.configure(state="disabled")
            if not self._output_visible:
                self._output_toggle_var.set(True)
                self.toggle_output()
        if ret is None:
            self.root.after(200, lambda: self._poll_output(proc))
        else:
            self.output_text.configure(state="normal")
            self.output_text.insert("end", f"\n--- exited with code {ret} ---\n")
            self.output_text.see("end")
            self.output_text.configure(state="disabled")
            self._toast("Command finished")

    # -----------------------------------------------------------------------
    # Right-click context menu
    # -----------------------------------------------------------------------
    def _show_card_menu(self, event, idx):
        menu = tk.Menu(self.root, tearoff=0, font=FONTS["small"])
        item = self.commands[idx]
        menu.add_command(label="Run", command=lambda: self.run_powershell(item["cmd"]))
        menu.add_command(label="Edit", command=lambda: self.edit_command(idx))
        menu.add_command(label="Duplicate", command=lambda: self._duplicate_command(idx))
        menu.add_separator()
        menu.add_command(label="Delete", command=lambda: self._delete_command(idx))
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

        for idx, item in enumerate(self.commands):
            card = tk.Frame(self.buttons_frame, bg=THEME["card_bg"], bd=0, highlightthickness=1, highlightbackground=THEME["card_border"], padx=8, pady=8)

            # Hover effect
            def on_enter(e, c=card):
                c.configure(bg=THEME["card_hover"])
                for child in c.winfo_children():
                    try:
                        child.configure(bg=THEME["card_hover"])
                    except tk.TclError:
                        pass

            def on_leave(e, c=card):
                c.configure(bg=THEME["card_bg"])
                for child in c.winfo_children():
                    try:
                        child.configure(bg=THEME["card_bg"])
                    except tk.TclError:
                        pass

            card.bind("<Enter>", on_enter)
            card.bind("<Leave>", on_leave)

            # Right-click
            card.bind("<Button-3>", lambda e, i=idx: self._show_card_menu(e, i))

            # Drag-and-drop
            card.bind("<ButtonPress-1>", lambda e, i=idx: self._drag_start(i))
            card.bind("<B1-Motion>", self._drag_motion)
            card.bind("<ButtonRelease-1>", self._drag_end)

            # Image
            photo = self._load_thumbnail(item.get("image"))
            if photo:
                self._image_refs.append(photo)
                lbl_img = tk.Label(card, image=photo, bg=THEME["card_bg"], bd=0)
            else:
                lbl_img = tk.Label(card, text="No Image", font=FONTS["small"], bg=THEME["sidebar_bg"], fg=THEME["text_secondary"], width=14, height=6)
            lbl_img.pack(pady=(0, 6))
            lbl_img.bind("<Button-3>", lambda e, i=idx: self._show_card_menu(e, i))

            # Name label
            name_lbl = tk.Label(card, text=item["name"], font=FONTS["card_name"], bg=THEME["card_bg"], fg=THEME["text_primary"])
            name_lbl.pack()

            # Category tag
            cat = item.get("category", "")
            if cat:
                cat_lbl = tk.Label(card, text=cat, font=FONTS["small"], bg=THEME["sidebar_active"], fg=THEME["text_secondary"], padx=6, pady=1)
                cat_lbl.pack(pady=(2, 4))

            # Buttons row
            btn_row = tk.Frame(card, bg=THEME["card_bg"])
            btn_row.pack(pady=(6, 0))

            run_btn = self._themed_button(btn_row, "Run", lambda c=item["cmd"]: self.run_powershell(c), "accent")
            run_btn.pack(side="left", padx=3)

            edit_btn = self._themed_button(btn_row, "Edit", lambda i=idx: self.edit_command(i))
            edit_btn.pack(side="left", padx=3)

            self._card_widgets.append((card, idx))

    # -----------------------------------------------------------------------
    # Card grid layout
    # -----------------------------------------------------------------------
    def refresh_buttons(self, filter_text=""):
        for frame, _ in self._card_widgets:
            frame.grid_forget()

        row = 0
        col = 0
        for frame, idx in self._card_widgets:
            item = self.commands[idx]
            # Category filter
            if self._active_category != "All":
                if item.get("category", "") != self._active_category:
                    continue
            # Text filter
            if filter_text and filter_text not in item["name"].lower():
                continue

            frame.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")
            self.buttons_frame.columnconfigure(col, weight=1)
            col += 1
            if col >= COLUMNS:
                col = 0
                row += 1

        self._update_status()

    # -----------------------------------------------------------------------
    # Drag and drop reordering
    # -----------------------------------------------------------------------
    def _drag_start(self, idx):
        self._drag_data["idx"] = idx

    def _drag_motion(self, event):
        pass  # visual feedback could be added later

    def _drag_end(self, event):
        src = self._drag_data["idx"]
        if src is None:
            return
        # Determine which card we're over
        widget = event.widget.winfo_containing(event.x_root, event.y_root)
        target_idx = None
        for frame, idx in self._card_widgets:
            if widget is frame or widget is not None and str(widget).startswith(str(frame)):
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
