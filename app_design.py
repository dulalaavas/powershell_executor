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

# File to save custom commands (resolve relative to script location)
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "commands.json")

class PowerShellApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Desktop Application Design")
        self.root.geometry("800x600")

        # Data Storage
        self.commands = self.load_commands()

        # 1. Menubar (File, Edit, View, Help)
        self.create_menubar()

        # Main Layout: Sidebar (Left) and Content (Right)
        main_container = tk.Frame(root)
        main_container.pack(fill="both", expand=True)

        # 2. Sidebar for Categories
        self.sidebar = tk.Frame(main_container, width=150, bg="lightgray", relief="sunken")
        self.sidebar.pack(side="left", fill="y")
        
        tk.Label(self.sidebar, text="Home", bg="lightgray", anchor="w").pack(fill="x", padx=5, pady=5)
        tk.Label(self.sidebar, text="<Category 1>", bg="lightgray", anchor="w").pack(fill="x", padx=5, pady=5)
        tk.Label(self.sidebar, text="<Category 2>", bg="lightgray", anchor="w").pack(fill="x", padx=5, pady=5)

        # 3. Content Area
        self.content_area = tk.Frame(main_container, bg="white")
        self.content_area.pack(side="right", fill="both", expand=True)

        # Search Bar
        search_frame = tk.Frame(self.content_area, bg="white")
        search_frame.pack(fill="x", padx=10, pady=5)
        tk.Label(search_frame, text="Search:", bg="white").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.update_list)
        self._search_after_id = None
        self.search_entry = tk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=5)

        # Scrollable Button Grid Area
        self.canvas = tk.Canvas(self.content_area, bg="white", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.content_area, orient="vertical", command=self.canvas.yview)
        self.buttons_frame = tk.Frame(self.canvas, bg="white")

        self.buttons_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.buttons_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        self.scrollbar.pack(side="right", fill="y")

        # Enable mouse wheel scrolling
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        # Cache for command card widgets: list of (frame, index) tuples
        self._card_widgets = []
        # Keep references to PhotoImage objects so they aren't garbage collected
        self._image_refs = []

        # 4. Floating Action Button (+)
        fab = tk.Button(root, text="+", bg="blue", fg="white", font=("Arial", 16), command=self.add_new_command)
        fab.place(relx=0.92, rely=0.9, width=50, height=50)

        self._build_cards()
        self.refresh_buttons()

    def create_menubar(self):
        menubar = tk.Menu(self.root)

        # --- File ---
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Command", accelerator="Ctrl+N", command=self.add_new_command)
        file_menu.add_separator()
        file_menu.add_command(label="Import Commands...", command=self.import_commands)
        file_menu.add_command(label="Export Commands...", command=self.export_commands)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", accelerator="Alt+F4", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        # --- Edit ---
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Find", accelerator="Ctrl+F", command=self.focus_search)
        edit_menu.add_separator()
        edit_menu.add_command(label="Delete All Commands", command=self.delete_all_commands)
        menubar.add_cascade(label="Edit", menu=edit_menu)

        # --- View ---
        view_menu = tk.Menu(menubar, tearoff=0)
        self._sidebar_visible = tk.BooleanVar(value=True)
        view_menu.add_checkbutton(label="Sidebar", variable=self._sidebar_visible, command=self.toggle_sidebar)
        view_menu.add_separator()
        view_menu.add_command(label="Refresh", accelerator="F5", command=self.refresh_all)
        menubar.add_cascade(label="View", menu=view_menu)

        # --- Help ---
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Keyboard Shortcuts", command=self.show_shortcuts)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)

        # --- Keyboard bindings ---
        self.root.bind("<Control-n>", lambda e: self.add_new_command())
        self.root.bind("<Control-N>", lambda e: self.add_new_command())
        self.root.bind("<Control-f>", lambda e: self.focus_search())
        self.root.bind("<Control-F>", lambda e: self.focus_search())
        self.root.bind("<F5>", lambda e: self.refresh_all())

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
        path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
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
                self._build_cards()
                self.refresh_buttons(self.search_var.get().lower())
            messagebox.showinfo("Import", f"Imported {count} command(s).")
        except (json.JSONDecodeError, IOError) as e:
            messagebox.showerror("Import Error", f"Could not read file:\n{e}")

    def export_commands(self):
        if not self.commands:
            messagebox.showinfo("Export", "No commands to export.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "w") as f:
                json.dump(self.commands, f, indent=2)
            messagebox.showinfo("Export", f"Exported {len(self.commands)} command(s).")
        except IOError as e:
            messagebox.showerror("Export Error", f"Could not write file:\n{e}")

    def focus_search(self):
        self.search_entry.focus_set()
        self.search_entry.select_range(0, tk.END)

    def delete_all_commands(self):
        if not self.commands:
            messagebox.showinfo("Delete All", "No commands to delete.")
            return
        if messagebox.askyesno("Confirm", f"Delete all {len(self.commands)} command(s)? This cannot be undone."):
            self.commands.clear()
            self.save_commands()
            self._build_cards()
            self.refresh_buttons()

    def toggle_sidebar(self):
        if self._sidebar_visible.get():
            self.sidebar.pack(side="left", fill="y", before=self.content_area)
        else:
            self.sidebar.pack_forget()

    def refresh_all(self):
        self.commands = self.load_commands()
        self._build_cards()
        self.refresh_buttons(self.search_var.get().lower())

    def show_shortcuts(self):
        shortcuts = (
            "Ctrl+N\tNew Command\n"
            "Ctrl+F\tFind / Focus Search\n"
            "F5\tRefresh\n"
            "Alt+F4\tExit"
        )
        messagebox.showinfo("Keyboard Shortcuts", shortcuts)

    def show_about(self):
        messagebox.showinfo(
            "About",
            "PowerShell Command Runner\n\n"
            "A desktop utility to save and run\n"
            "PowerShell commands with one click."
        )

    def add_new_command(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("New Command")
        dlg.geometry("400x300")
        dlg.resizable(False, False)
        dlg.grab_set()

        # --- Name ---
        tk.Label(dlg, text="Name:").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 2))
        name_var = tk.StringVar()
        tk.Entry(dlg, textvariable=name_var, width=40).grid(row=0, column=1, padx=10, pady=(10, 2))

        # --- Command ---
        tk.Label(dlg, text="Command:").grid(row=1, column=0, sticky="w", padx=10, pady=2)
        cmd_var = tk.StringVar()
        tk.Entry(dlg, textvariable=cmd_var, width=40).grid(row=1, column=1, padx=10, pady=2)

        # --- Image ---
        tk.Label(dlg, text="Image:").grid(row=2, column=0, sticky="w", padx=10, pady=2)
        img_var = tk.StringVar()
        tk.Entry(dlg, textvariable=img_var, width=28).grid(row=2, column=1, sticky="w", padx=10, pady=2)

        def browse_image():
            filetypes = [("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All files", "*.*")]
            path = filedialog.askopenfilename(parent=dlg, filetypes=filetypes)
            if path:
                img_var.set(path)

        tk.Button(dlg, text="Browse...", command=browse_image).grid(row=2, column=1, sticky="e", padx=10, pady=2)

        # --- Create / Cancel buttons ---
        btn_frame = tk.Frame(dlg)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=20)

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
            self.commands.append(new_item)
            self.save_commands()
            self._build_cards()
            self.refresh_buttons(self.search_var.get().lower())
            dlg.destroy()

        tk.Button(btn_frame, text="Create", width=10, command=create).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Cancel", width=10, command=dlg.destroy).pack(side="left", padx=10)

    def run_powershell(self, command):
        if not messagebox.askyesno("Confirm", f"Run this command?\n\n{command}"):
            return
        try:
            subprocess.Popen(["powershell", "-Command", command])
        except FileNotFoundError:
            messagebox.showerror("Error", "PowerShell not found. Is it installed and on PATH?")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def edit_command(self, idx):
        """Open an edit dialog for the command at the given index."""
        item = self.commands[idx]

        dlg = tk.Toplevel(self.root)
        dlg.title(f"Edit: {item['name']}")
        dlg.geometry("400x300")
        dlg.resizable(False, False)
        dlg.grab_set()

        # --- Name ---
        tk.Label(dlg, text="Name:").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 2))
        name_var = tk.StringVar(value=item["name"])
        tk.Entry(dlg, textvariable=name_var, width=40).grid(row=0, column=1, padx=10, pady=(10, 2))

        # --- Command ---
        tk.Label(dlg, text="Command:").grid(row=1, column=0, sticky="w", padx=10, pady=2)
        cmd_var = tk.StringVar(value=item["cmd"])
        tk.Entry(dlg, textvariable=cmd_var, width=40).grid(row=1, column=1, padx=10, pady=2)

        # --- Image ---
        tk.Label(dlg, text="Image:").grid(row=2, column=0, sticky="w", padx=10, pady=2)
        img_var = tk.StringVar(value=item.get("image", ""))
        img_entry = tk.Entry(dlg, textvariable=img_var, width=28)
        img_entry.grid(row=2, column=1, sticky="w", padx=10, pady=2)

        def browse_image():
            filetypes = [("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All files", "*.*")]
            path = filedialog.askopenfilename(parent=dlg, filetypes=filetypes)
            if path:
                img_var.set(path)

        tk.Button(dlg, text="Browse...", command=browse_image).grid(row=2, column=1, sticky="e", padx=10, pady=2)

        # --- Save / Delete buttons ---
        btn_frame = tk.Frame(dlg)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=20)

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
            self.save_commands()
            self._build_cards()
            self.refresh_buttons(self.search_var.get().lower())
            dlg.destroy()

        def delete():
            if messagebox.askyesno("Confirm Delete", f"Delete '{item['name']}'?", parent=dlg):
                self.commands.pop(idx)
                self.save_commands()
                self._build_cards()
                self.refresh_buttons(self.search_var.get().lower())
                dlg.destroy()

        tk.Button(btn_frame, text="Save", width=10, command=save).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Delete", width=10, fg="red", command=delete).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Cancel", width=10, command=dlg.destroy).pack(side="left", padx=10)

    def update_list(self, *args):
        if self._search_after_id is not None:
            self.root.after_cancel(self._search_after_id)
        self._search_after_id = self.root.after(200, self._do_search)

    def _do_search(self):
        self._search_after_id = None
        search_term = self.search_var.get().lower()
        self.refresh_buttons(search_term)

    def _load_thumbnail(self, path):
        """Load an image file and return a PhotoImage thumbnail, or None on failure."""
        if not HAS_PIL or not path or not os.path.isfile(path):
            return None
        try:
            img = Image.open(path)
            img.thumbnail((80, 80))
            return ImageTk.PhotoImage(img)
        except Exception:
            return None

    def _build_cards(self):
        """Pre-build a widget card for each command. Called once at startup and after adding commands."""
        for frame, _ in self._card_widgets:
            frame.destroy()
        self._card_widgets = []
        self._image_refs = []

        for idx, item in enumerate(self.commands):
            frame = tk.Frame(self.buttons_frame, bd=2, relief="groove", padx=5, pady=5)

            # Image area — show thumbnail if set, otherwise placeholder
            photo = self._load_thumbnail(item.get("image"))
            if photo:
                self._image_refs.append(photo)
                lbl_img = tk.Label(frame, image=photo, width=80, height=80)
            else:
                lbl_img = tk.Label(frame, text="[IMG]", bg="gray", width=10, height=5)
            lbl_img.pack()

            btn = tk.Button(frame, text=f"Run {item['name']}", command=lambda c=item["cmd"]: self.run_powershell(c))
            btn.pack(pady=(5, 2))

            edit_btn = tk.Button(frame, text="Edit", command=lambda i=idx: self.edit_command(i))
            edit_btn.pack(pady=(0, 5))

            self._card_widgets.append((frame, idx))

    def refresh_buttons(self, filter_text=""):
        """Show/hide pre-built cards based on filter. No widget creation or destruction."""
        # Hide all cards first
        for frame, _ in self._card_widgets:
            frame.grid_forget()

        # Show matching cards in grid layout
        row = 0
        col = 0
        for frame, idx in self._card_widgets:
            if filter_text in self.commands[idx]["name"].lower():
                frame.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")
                col += 1
                if col > 2:
                    col = 0
                    row += 1

if __name__ == "__main__":
    root = tk.Tk()
    app = PowerShellApp(root)
    root.mainloop()
