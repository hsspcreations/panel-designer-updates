__version__ = "2025.08.31"

import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox, ttk
import tkinter.font as tkfont
import json
import requests
import shutil
import pandas as pd
import numpy as np
import os
import sys
import gspread
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas
import subprocess
from gspread_formatting import format_cell_range, CellFormat, Color, TextFormat
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from PIL import Image, ImageTk
from collections import defaultdict


# ====== Persisted Version Helpers (Injected) ======
import os, sys, time, shutil, requests

APPDATA_FOLDER = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "PanelDesigner")
os.makedirs(APPDATA_FOLDER, exist_ok=True)
VERSION_FILE = os.path.join(APPDATA_FOLDER, "version.txt")

def get_installed_version():
    try:
        with open(VERSION_FILE, "r", encoding="utf-8") as f:
            return f.read().strip() or "0"
    except Exception:
        return "0"

def set_installed_version(v):
    try:
        with open(VERSION_FILE, "w", encoding="utf-8") as f:
            f.write(str(v))
    except Exception:
        pass

def fetch_remote_version():
    try:
        url = "https://raw.githubusercontent.com/hsspcreations/panel-designer-updates/refs/heads/main/version.txt"
        r = requests.get(url, timeout=6)
        if r.status_code == 200:
            return r.text.strip()
    except Exception:
        return None
    return None
# ====== End Injected Helpers ======


BREAKER_FILE = "breaker_types.json"
APPDATA_FOLDER = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "PanelDesigner")
PANELS_FOLDER = os.path.join(APPDATA_FOLDER, "panels")
os.makedirs(PANELS_FOLDER, exist_ok=True)
TOKEN_FILE = os.path.join(APPDATA_FOLDER, "token.json")


def update_software():
    remote_ver = fetch_remote_version()
    if not remote_ver:
        try:
            from tkinter import messagebox
            messagebox.showerror("Update", "Could not fetch remote version.")
        except Exception: pass
        return
    current_ver = get_installed_version()
    if str(remote_ver) <= str(current_ver):
        try:
            from tkinter import messagebox
            messagebox.showinfo("Update", "Already up to date.")
        except Exception: pass
        return

    import time
    messagebox.showinfo("Updater", f"Current version: {__version__}\nDownloading latest version...")
    UPDATE_URL = "https://raw.githubusercontent.com/hsspcreations/panel-designer-updates/refs/heads/main/main.py"
    try:
        url = f"{UPDATE_URL}?t={int(time.time())}"  # bypass cache
        response = requests.get(url, stream=True)
        if response.status_code == 200:
            backup_path = "main_backup.py"
            if os.path.exists("main.py"):
                shutil.copy("main.py", backup_path)  # backup old version
            with open("main.py", "wb") as f:
                f.write(response.content)
            messagebox.showinfo("Update Complete", "Software updated successfully. Restarting now...")
            os.execl(sys.executable, sys.executable, *sys.argv)  # restart app
        else:
            messagebox.showerror("Update Failed", "Could not download the update.")
    except Exception as e:
        messagebox.showerror("Error", f"Update failed: {e}")


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller"""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


CREDENTIALS_FILE = resource_path("credentials.json")
BUSBAR_DATA_FILE = resource_path("Copy of Quotation Calculator v1.5.csv")

CUBICLE_SIZES = [
    "600mm x 1800mm", "800mm x 1800mm", "1000mm x 1800mm",
    "600mm x 2000mm", "800mm x 2000mm", "1000mm x 2000mm",
    "600mm x 400mm", "800mm x 400mm", "1000mm x 400mm",
    "600mm x 600mm", "800mm x 600mm", "1000mm x 600mm"
]

SECTION_NAMES = ["Breaker", "ELR/EFR", "PFR", "Power Analyzer/Energy Meter", "Indicator Light", "SPD"]
SCALE = 0.2

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]


class Tooltip:
    def __init__(self, canvas, text):
        self.canvas = canvas
        self.text = text
        self.tip_window = None

    def show(self, x, y):
        self.hide()
        self.tip_window = tk.Toplevel(self.canvas)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.geometry(f"+{x+15}+{y+15}")
        label = tk.Label(self.tip_window, text=self.text, background="yellow", relief="solid", borderwidth=1, font=("Arial", 8))
        label.pack()

    def hide(self):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


class PanelDesigner:
    def __init__(self, root, customer, project, ref):
        self.root = root
        self.customer = customer
        self.project = project
        self.ref = ref
        self.root.title(f"Panel Designer - {project}")
        try:
            self.root.iconbitmap(resource_path("Hssp.ico"))
        except Exception:
            pass

        self.breaker_types = self.load_breaker_types()
        self.busbar_data = self.load_busbar_data()
        self.saved_panels = self.load_saved_panels()
        self.panel_name = None
        self.panel_depth = None  # store panel depth (mm)
        self.cubicles = []
        self.busbars = []
        self.tooltip = None
        self.icon_image = None
        self.drag_data = {"item": None, "x": 0, "y": 0}
        self.undo_stack = []
        self.footer_ids = []  # track footer elements for theme refresh

        # THEME STATE
        self.is_dark_mode = False
        self.style = ttk.Style()
        self.palette = self.get_palette("light")

        top_frame = tk.Frame(root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        self.panel_var = tk.StringVar()
        self.panel_var.set("Select Panel" if self.saved_panels else "No Panels")
        self.panel_menu = tk.OptionMenu(top_frame, self.panel_var, *(["Select Panel"] + self.saved_panels),
                                        command=self.on_panel_select)
        self.panel_menu.pack(side=tk.LEFT, padx=5)

        tk.Button(top_frame, text="Create Panel", command=self.create_panel).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Add Cubicle", command=self.add_cubicle).pack(side=tk.LEFT, padx=5)

        # === NEW: Single dropdown for the three busbar actions ===
        busbar_mb = tk.Menubutton(top_frame, text="➕ Busbar", relief=tk.RAISED, borderwidth=1)
        busbar_menu = tk.Menu(busbar_mb, tearoff=False)
        busbar_menu.add_command(label="Add Vertical Busbar", command=self.add_vertical_busbar_form)
        busbar_menu.add_command(label="Add Horizontal Busbar", command=self.add_horizontal_busbar_form)
        busbar_menu.add_command(label="Add Busbar Terminal", command=self.add_busbar_terminal_form)
        busbar_mb.configure(menu=busbar_menu)
        busbar_mb.pack(side=tk.LEFT, padx=5)

        tk.Button(top_frame, text="Upload Breaker Types", command=self.load_breaker_excel).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Save Panel", command=self.save_panel).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Generate BOM", command=self.generate_bom).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Undo", command=self.undo_last_action).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Update Software", command=update_software).pack(side=tk.LEFT, padx=5)

        # === NEW: Dark/Light mode toggle as a Checkbutton ===
        self.dark_mode_var = tk.BooleanVar(value=False)
        self.dark_toggle = ttk.Checkbutton(top_frame, text="Dark Mode", variable=self.dark_mode_var, command=self.toggle_theme_check)
        self.dark_toggle.pack(side=tk.RIGHT, padx=5)

        canvas_frame = tk.Frame(root)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.canvas = tk.Canvas(canvas_frame, bg="lightgray", width=1000, height=600, scrollregion=(0, 0, 5000, 3000))
        h_scroll = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        v_scroll = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.apply_theme()
        self.add_bottom_right_info()

    def add_bottom_right_info(self):
        # clear previous footer items if any
        if hasattr(self, "footer_ids") and self.footer_ids:
            for _id in self.footer_ids:
                try:
                    self.canvas.delete(_id)
                except Exception:
                    pass
            self.footer_ids = []
        canvas_width = int(self.canvas['width'])
        canvas_height = int(self.canvas['height'])
        padding = 10

        icon_path = resource_path("Hssp.ico")
        if os.path.exists(icon_path):
            try:
                img = Image.open(icon_path)
                img = img.resize((32, 32))
                self.icon_image = ImageTk.PhotoImage(img)
                self.footer_ids.append(self.canvas.create_image(canvas_width - 40, canvas_height - 60, image=self.icon_image, anchor="se"))
            except Exception:
                pass

        self.footer_ids.append(self.canvas.create_text(canvas_width - padding, canvas_height - 30,
                                text="hsspcreations@gmail.com", font=("Arial", 10), fill=self.palette["muted_text"], anchor="se"))
        self.footer_ids.append(self.canvas.create_text(canvas_width - padding, canvas_height - 10,
                                text="0764319139", font=("Arial", 10), fill=self.palette["muted_text"], anchor="se"))

    def load_breaker_types(self):
        if os.path.exists(BREAKER_FILE):
            try:
                with open(BREAKER_FILE, "r") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def save_breaker_types(self):
        with open(BREAKER_FILE, "w") as f:
            json.dump(self.breaker_types, f)

    def load_busbar_data(self):
        try:
            df = pd.read_csv(BUSBAR_DATA_FILE)
            return df
        except FileNotFoundError:
            messagebox.showerror("Error", f"Busbar data file not found at: {BUSBAR_DATA_FILE}")
            return pd.DataFrame()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load busbar data: {e}")
            return pd.DataFrame()

    def project_key(self):
        return f"{self.customer}_{self.project}_{self.ref}"

    def load_saved_panels(self):
        os.makedirs(PANELS_FOLDER, exist_ok=True)
        panels = []
        for file in os.listdir(PANELS_FOLDER):
            if file.endswith(".json"):
                try:
                    with open(f"{PANELS_FOLDER}/{file}", "r") as f:
                        data = json.load(f)
                        pinfo = data.get("project_info", {})
                        if (pinfo.get("customer") == self.customer and
                                pinfo.get("project") == self.project and
                                pinfo.get("ref") == self.ref):
                            panels.append(file[:-5])
                except Exception:
                    continue
        return panels

    def refresh_panel_menu(self):
        self.saved_panels = self.load_saved_panels()
        menu = self.panel_menu["menu"]
        menu.delete(0, "end")
        if self.saved_panels:
            menu.add_command(label="Select Panel", command=lambda: None)
            for panel in self.saved_panels:
                menu.add_command(label=panel, command=lambda value=panel: self.on_panel_select(value))
            self.panel_var.set(self.panel_name if self.panel_name in self.saved_panels else "Select Panel")
        else:
            menu.add_command(label="No Panels", command=lambda: None)
            self.panel_var.set("No Panels")

    def on_panel_select(self, selected_panel):
        if selected_panel not in ("No Panels", "Select Panel"):
            self.panel_var.set(selected_panel)
            self.load_panel(selected_panel)

    def create_panel(self):
        # Combined form for panel name and depth in a single window
        top = tk.Toplevel(self.root)
        top.title("Create New Panel")
        top.grab_set()

        frm = tk.Frame(top, padx=10, pady=10)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text="Panel Name:").grid(row=0, column=0, sticky="e", pady=5, padx=5)
        name_var = tk.StringVar()
        tk.Entry(frm, textvariable=name_var, width=30).grid(row=0, column=1, pady=5, padx=5)

        tk.Label(frm, text="Panel Depth (mm):").grid(row=1, column=0, sticky="e", pady=5, padx=5)
        depth_var = tk.StringVar(value="600")
        tk.Entry(frm, textvariable=depth_var, width=10).grid(row=1, column=1, sticky="w", pady=5, padx=5)

        def confirm():
            name = name_var.get().strip()
            if not name:
                messagebox.showerror("Invalid Name", "Please enter a panel name.")
                return
            try:
                depth = int(depth_var.get())
                if depth <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid Depth", "Depth must be a positive integer (mm).")
                return

            self.panel_name = name
            self.panel_depth = depth

            self.cubicles.clear()
            self.busbars.clear()
            self.canvas.delete("all")
            self.panel_var.set(name)
            self.apply_theme()
            self.add_bottom_right_info()

            try:
                self.canvas.create_text(20, 10, text=f"Depth: {self.panel_depth} mm", anchor="nw", font=("Arial", 10, "bold"))
            except Exception:
                pass

            top.destroy()

        btns = tk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(btns, text="Create", command=confirm).pack(side="left", padx=5)
        ttk.Button(btns, text="Cancel", command=top.destroy).pack(side="left", padx=5)

        top.wait_window()

    def add_cubicle(self):
        if not self.panel_name:
            messagebox.showwarning("No Panel", "Please create or select a panel first.")
            return

        top = tk.Toplevel(self.root)
        top.title("Select Cubicle Size")
        tk.Label(top, text="Select Cubicle Size:").pack(pady=5)

        selected_size = tk.StringVar(value=CUBICLE_SIZES[0])
        combo = ttk.Combobox(top, textvariable=selected_size, values=CUBICLE_SIZES, state="readonly")
        combo.pack(padx=10, pady=5)

        def on_confirm():
            size = selected_size.get()
            width, height = map(int, size.replace("mm", "").split("x"))
            w, h = width * SCALE, height * SCALE

            if not self.cubicles:
                x, y = 50, 50
            else:
                last = self.cubicles[-1]
                x = last["x"] + last["width"] * SCALE
                y = last["y"]

            rect = self.canvas.create_rectangle(x, y, x + w, y + h, fill=self.palette["cubicle_fill"], outline=self.palette["cubicle_outline"], width=3)

            cubicle_data = {"id": rect, "width": width, "height": height, "x": x, "y": y, "compartments": []}
            self.cubicles.append(cubicle_data)
            self.undo_stack.append({"type": "add_cubicle", "cubicle": cubicle_data})

            self.ask_compartments(cubicle_data)
            top.destroy()

        tk.Button(top, text="Add Cubicle", command=on_confirm).pack(pady=5)
        top.grab_set()
        top.wait_window()

    def delete_selected_cubicle(self):
        if not self.cubicles:
            messagebox.showwarning("Delete Cubicle", "No cubicles to delete.")
            return
        cubicle = self.cubicles.pop()
        self.canvas.delete(cubicle["id"])
        for comp in cubicle["compartments"]:
            for sec in comp["sections"]:
                self.canvas.delete(sec["id"])
        messagebox.showinfo("Delete Cubicle", "Last added cubicle deleted successfully!")

    def ask_compartments(self, cubicle):
        num = simpledialog.askinteger("Compartments", "Enter number of compartments:", minvalue=1)
        if num:
            self.create_compartments(cubicle, num)

    def create_compartments(self, cubicle, num):
        coords = self.canvas.coords(cubicle["id"])
        x1, y1, x2, y2 = coords
        compartment_height = (y2 - y1) / num

        for _ in range(num):
            comp_y1 = y1 + len(cubicle["compartments"]) * compartment_height
            comp_y2 = comp_y1 + compartment_height
            compartment = {"sections": []}

            section_width = (x2 - x1) / len(SECTION_NAMES)
            for j, section_name in enumerate(SECTION_NAMES):
                sec_x1 = x1 + j * section_width
                sec_x2 = sec_x1 + section_width
                section_rect = self.canvas.create_rectangle(sec_x1, comp_y1, sec_x2, comp_y2, fill=self.palette["section_empty"], outline=self.palette["section_outline"])
                self.canvas.tag_bind(section_rect, "<Button-1>",
                                     lambda e, s=section_name, r=section_rect, c=compartment: self.select_item(s, r, c))
                compartment["sections"].append({"name": section_name, "id": section_rect, "item": None})

            cubicle["compartments"].append(compartment)

    def select_item(self, section_name, section_rect, compartment):
        self.show_search_popup(section_name, section_rect, compartment)

    # ---------- TEXT FITTING HELPERS ----------
    def _compute_text_layout(self, section_rect, font_name=("Arial", 6)):
        coords = self.canvas.coords(section_rect)
        x1, y1, x2, y2 = coords
        width = max(1, x2 - x1)
        height = max(1, y2 - y1)

        size = font_name[1] if isinstance(font_name, tuple) else 6
        # Attempt to find a font size that fits at least one column
        for fs in range(int(size), 3, -1):
            fnt = tkfont.Font(family=font_name[0] if isinstance(font_name, tuple) else "Arial", size=fs)
            line_h = max(1, fnt.metrics("linespace"))
            char_w = max(1, fnt.measure("W"))
            max_lines = max(1, int(height // line_h))
            if max_lines < 1:
                continue
            # For a single column: need char_w width
            if char_w <= width:
                return fnt, line_h, char_w, max_lines
        # Fallback minimal font
        fnt = tkfont.Font(family="Arial", size=4)
        line_h = max(1, fnt.metrics("linespace"))
        char_w = max(1, fnt.measure("W"))
        max_lines = max(1, int(height // line_h))
        return fnt, line_h, char_w, max_lines

    def _split_text_into_columns(self, text, max_lines):
        # Split into chunks (columns) of length max_lines
        chunks = []
        text = str(text)
        for i in range(0, len(text), max_lines):
            chunks.append(text[i:i + max_lines])
        return chunks

    def draw_vertical_text_in_section(self, section, text, desc):
        """Draw vertical, wrapped text that fits inside the section rectangle.
        Stores the created text item ids in section['item']['text_ids'].
        """
        section_rect = section["id"]
        # Remove existing text ids if any
        try:
            old_text_ids = section.get("item", {}).get("text_ids", [])
            for tid in old_text_ids or []:
                self.canvas.delete(tid)
        except Exception:
            pass

        coords = self.canvas.coords(section_rect)
        x1, y1, x2, y2 = coords
        width = x2 - x1

        fnt, line_h, char_w, max_lines = self._compute_text_layout(section_rect, ("Arial", 6))
        if max_lines < 1:
            max_lines = 1
        columns = self._split_text_into_columns(text, max_lines)

        # Compute how many columns fit horizontally; if not all fit, truncate with ellipsis
        col_gap = max(2, int(char_w * 0.5))
        total_needed = len(columns) * char_w + (len(columns) - 1) * col_gap
        max_cols_fit = max(1, int((width + col_gap) // (char_w + col_gap)))
        draw_columns = columns[:max_cols_fit]
        truncated = len(columns) > max_cols_fit

        # Center the columns horizontally
        draw_width = len(draw_columns) * char_w + (len(draw_columns) - 1) * col_gap
        start_x = x1 + (width - draw_width) / 2 + char_w / 2
        center_y = (y1 + y2) / 2

        text_ids = []
        for idx, chunk in enumerate(draw_columns):
            col_text = "\n".join(list(chunk))
            tx = start_x + idx * (char_w + col_gap)
            tid = self.canvas.create_text(tx, center_y, text=col_text, font=fnt, fill=self.palette["text"], anchor="center", justify="center")
            self.canvas.tag_bind(tid, "<Enter>", lambda e, d=desc: self.show_tooltip(e, d))
            self.canvas.tag_bind(tid, "<Leave>", lambda e: self.hide_tooltip())
            text_ids.append(tid)

        # If truncated, draw a tiny ellipsis at the far right
        if truncated:
            ellipsis_id = self.canvas.create_text(x2 - 2, y1 + 2, text="…", font=("Arial", max(5, fnt.cget("size") - 1)), fill=self.palette["text"], anchor="ne")
            text_ids.append(ellipsis_id)

        # Save text ids with the item for future cleanup/undo
        if section.get("item"):
            section["item"]["text_ids"] = text_ids

        return text_ids

    def show_search_popup(self, section_name, section_rect, compartment):
        popup = tk.Toplevel(self.root)
        popup.title(f"Select {section_name}")
        popup.geometry("300x400")

        search_var = tk.StringVar()
        search_entry = tk.Entry(popup, textvariable=search_var)
        search_entry.pack(fill=tk.X, padx=5, pady=5)

        listbox = tk.Listbox(popup)
        listbox.pack(fill=tk.BOTH, expand=True)

        def update_list(*args):
            search_text = search_var.get().lower()
            listbox.delete(0, tk.END)
            for model, desc in self.breaker_types.items():
                if search_text in desc.lower() or search_text in model.lower():
                    listbox.insert(tk.END, f"{desc} ({model})")

        search_var.trace("w", update_list)
        update_list()

        def select_item():
            if listbox.curselection():
                selected = listbox.get(listbox.curselection())
                model = selected.split("(")[-1].strip(")")
                desc = selected.split("(")[0].strip()

                self.canvas.itemconfig(section_rect, fill=self.palette["section_selected"])
                # find the section object
                target_section = None
                for section in compartment["sections"]:
                    if section["id"] == section_rect:
                        target_section = section
                        break
                if target_section is None:
                    popup.destroy()
                    return

                # remove any previous text in this section
                try:
                    old_ids = target_section.get("item", {}).get("text_ids", [])
                    for tid in old_ids or []:
                        self.canvas.delete(tid)
                except Exception:
                    pass

                previous_item = target_section.get("item")
                target_section["item"] = {"model": model, "desc": desc, "text_ids": []}

                # draw new text vertically to fit
                new_text_ids = self.draw_vertical_text_in_section(target_section, model, desc)
                target_section["item"]["text_ids"] = new_text_ids

                # push undo
                self.undo_stack.append({
                    "type": "select_component",
                    "section": target_section,
                    "previous_item": previous_item,
                    "new_text_ids": new_text_ids,
                    "rect_id": section_rect
                })

                popup.destroy()

        tk.Button(popup, text="Select", command=select_item).pack(pady=5)

    def show_tooltip(self, event, text):
        self.tooltip = Tooltip(self.canvas, text)
        self.tooltip.show(event.x_root, event.y_root)

    def hide_tooltip(self):
        if self.tooltip:
            self.tooltip.hide()

    def add_busbar_terminal_form(self):
        form = tk.Toplevel(self.root)
        form.title("Add Busbar Terminal")

        tk.Label(form, text="Busbar Size:").grid(row=0, column=0, padx=5, pady=5)
        busbar_sizes = [
            "20x6 Busbar (5.5m Length) LVT",
            "25x10 Busbar (5.5m Length) LVT",
            "32x10 Cu Busbar (5.5m Length) LVT",
            "40x10 Cu Busbar (5.5m Length) LVT",
            "50x10 Cu Busbar (5.5m Length) LVT",
            "63x10 Cu Busbar (5.5m Length) LVT",
            "75x10 Cu Busbar (5.5m Length) LVT",
            "80x10 Cu Busbar (5.5m Length) LVT",
            "100x10 Cu Busbar (5.5m Length) LVT"
        ]
        size_var = tk.StringVar(value=busbar_sizes[0])
        ttk.Combobox(form, textvariable=size_var, values=busbar_sizes, state="readonly").grid(row=0, column=1, padx=5, pady=5)

        tk.Label(form, text="No. of Runs:").grid(row=1, column=0, padx=5, pady=5)
        runs_var = tk.StringVar(value="1")
        runs_entry = tk.Entry(form, textvariable=runs_var)
        runs_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(form, text="Phase:").grid(row=2, column=0, padx=5, pady=5)
        phase_var = tk.StringVar(value="Single Phase")
        ttk.Combobox(form, textvariable=phase_var, values=["Single Phase", "Three Phase"], state="readonly").grid(row=2, column=1, padx=5, pady=5)

        tk.Label(form, text="Type:").grid(row=3, column=0, padx=5, pady=5)
        type_var = tk.StringVar(value="Horizontal")
        ttk.Combobox(form, textvariable=type_var, values=["Horizontal", "Vertical"], state="readonly").grid(row=3, column=1, padx=5, pady=5)

        def submit():
            try:
                runs_int = int(runs_var.get())
                if runs_int <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid Input", "No. of Runs must be a positive integer.")
                return
            self.spawn_busbar_terminal(size_var.get(), runs_int, phase_var.get(), type_var.get())
            form.destroy()

        tk.Button(form, text="Add", command=submit).grid(row=4, column=0, columnspan=2, pady=10)
        form.grab_set()
        form.wait_window()

    def spawn_busbar_terminal(self, busbar_size, no_of_runs, phase, busbar_type):
        if busbar_type.lower() == "horizontal":
            x1, x2 = 50, 250
            y = 150
            line_id = self.canvas.create_line(x1, y, x2, y, fill=self.palette["busbar_terminal"], width=4)
            coords = [x1, y, x2, y]
        else:
            x = 200
            y1, y2 = 50, 300
            line_id = self.canvas.create_line(x, y1, x, y2, fill=self.palette["busbar_terminal"], width=4)
            coords = [x, y1, x, y2]

        busbar_data = {
            "id": line_id,
            "type": busbar_type.lower(),
            "coords": coords,
            "amperage": None,
            "current_density": None,
            "phase": phase,
            "busbar_size": busbar_size,
            "no_of_runs": int(no_of_runs)
        }
        self.busbars.append(busbar_data)
        self.undo_stack.append({"type": "add_busbar", "busbar": busbar_data})
        self.make_busbar_draggable(line_id, busbar_type.lower())
        self.make_busbar_resizable(line_id, busbar_type.lower())

    def add_vertical_busbar_form(self):
        form = tk.Toplevel(self.root)
        form.title("Add Vertical Busbar")

        tk.Label(form, text="Amperage (A):").grid(row=0, column=0, padx=5, pady=5)
        amp_var = tk.IntVar(value=100)
        tk.Entry(form, textvariable=amp_var).grid(row=0, column=1, padx=5, pady=5)

        tk.Label(form, text="Current Density (A/mm²):").grid(row=1, column=0, padx=5, pady=5)
        cd_var = tk.DoubleVar(value=2.5)
        tk.Entry(form, textvariable=cd_var).grid(row=1, column=1, padx=5, pady=5)

        tk.Label(form, text="Phase:").grid(row=2, column=0, padx=5, pady=5)
        phase_var = tk.StringVar(value="Single Phase")
        ttk.Combobox(form, textvariable=phase_var, values=["Single Phase", "Three Phase"], state="readonly").grid(row=2, column=1, padx=5, pady=5)

        def submit():
            self.spawn_vertical_busbar(amp_var.get(), cd_var.get(), phase_var.get())
            form.destroy()

        tk.Button(form, text="Add", command=submit).grid(row=3, column=0, columnspan=2, pady=10)
        form.grab_set()
        form.wait_window()

    def spawn_vertical_busbar(self, amperage, current_density, phase):
        x = 150
        y1, y2 = 50, 300
        line_id = self.canvas.create_line(x, y1, x, y2, fill=self.palette["busbar"], width=4)
        self.canvas.tag_raise(line_id)
        busbar_data = {
            "id": line_id,
            "type": "vertical",
            "coords": [x, y1, x, y2],
            "amperage": amperage,
            "current_density": current_density,
            "phase": phase
        }
        self.busbars.append(busbar_data)
        self.undo_stack.append({"type": "add_busbar", "busbar": busbar_data})
        self.make_busbar_draggable(line_id, "vertical")
        self.make_busbar_resizable(line_id, "vertical")

    def add_horizontal_busbar_form(self):
        form = tk.Toplevel(self.root)
        form.title("Add Horizontal Busbar")

        tk.Label(form, text="Amperage (A):").grid(row=0, column=0, padx=5, pady=5)
        amp_var = tk.IntVar(value=100)
        tk.Entry(form, textvariable=amp_var).grid(row=0, column=1, padx=5, pady=5)

        tk.Label(form, text="Current Density (A/mm²):").grid(row=1, column=0, padx=5, pady=5)
        cd_var = tk.DoubleVar(value=2.5)
        tk.Entry(form, textvariable=cd_var).grid(row=1, column=1, padx=5, pady=5)

        tk.Label(form, text="Phase:").grid(row=2, column=0, padx=5, pady=5)
        phase_var = tk.StringVar(value="Single Phase")
        ttk.Combobox(form, textvariable=phase_var, values=["Single Phase", "Three Phase"], state="readonly").grid(row=2, column=1, padx=5, pady=5)

        def submit():
            self.spawn_horizontal_busbar(amp_var.get(), cd_var.get(), phase_var.get())
            form.destroy()

        tk.Button(form, text="Add", command=submit).grid(row=3, column=0, columnspan=2, pady=10)
        form.grab_set()
        form.wait_window()

    def spawn_horizontal_busbar(self, amperage, current_density, phase):
        x1, x2 = 50, 250
        y = 100
        line_id = self.canvas.create_line(x1, y, x2, y, fill=self.palette["busbar"], width=4)
        self.canvas.tag_raise(line_id)
        busbar_data = {
            "id": line_id,
            "type": "horizontal",
            "coords": [x1, y, x2, y],
            "amperage": amperage,
            "current_density": current_density,
            "phase": phase
        }
        self.busbars.append(busbar_data)
        self.undo_stack.append({"type": "add_busbar", "busbar": busbar_data})
        self.make_busbar_draggable(line_id, "horizontal")
        self.make_busbar_resizable(line_id, "horizontal")

    def make_busbar_draggable(self, line_id, busbar_type):
        def on_press(event):
            self.drag_data["item"] = line_id
            self.drag_data["x"] = event.x
            self.drag_data["y"] = event.y

        def on_release(event):
            self.drag_data["item"] = None

        def on_move(event):
            if self.drag_data["item"] == line_id:
                dx = event.x - self.drag_data["x"]
                dy = event.y - self.drag_data["y"]
                self.drag_data["x"] = event.x
                self.drag_data["y"] = event.y
                coords = self.canvas.coords(line_id)
                new_coords = [coords[0] + dx, coords[1] + dy, coords[2] + dx, coords[3] + dy]
                self.canvas.coords(line_id, *new_coords)
                for b in self.busbars:
                    if b["id"] == line_id:
                        b["coords"] = new_coords
                        break

        self.canvas.tag_bind(line_id, "<ButtonPress-1>", on_press)
        self.canvas.tag_bind(line_id, "<ButtonRelease-1>", on_release)
        self.canvas.tag_bind(line_id, "<B1-Motion>", on_move)

    def make_busbar_resizable(self, line_id, busbar_type):
        handle_size = 6
        coords = self.canvas.coords(line_id)
        handle_id = self.canvas.create_rectangle(coords[2] - handle_size, coords[3] - handle_size,
                                                 coords[2] + handle_size, coords[3] + handle_size, fill=self.palette["handle"], tags=("handle",))

        def on_press(event):
            self.drag_data["item"] = handle_id
            self.drag_data["x"] = event.x
            self.drag_data["y"] = event.y

        def on_release(event):
            self.drag_data["item"] = None

        def on_move(event):
            if self.drag_data["item"] == handle_id:
                dy = event.y - self.drag_data["y"]
                dx = event.x - self.drag_data["x"]
                self.drag_data["x"] = event.x
                self.drag_data["y"] = event.y
                coords = self.canvas.coords(line_id)
                if busbar_type == "vertical":
                    coords[3] += dy
                else:
                    coords[2] += dx
                self.canvas.coords(line_id, *coords)
                self.canvas.coords(handle_id, coords[2] - handle_size, coords[3] - handle_size,
                                   coords[2] + handle_size, coords[3] + handle_size)
                for b in self.busbars:
                    if b["id"] == line_id:
                        b["coords"] = coords
                        break

        self.canvas.tag_bind(handle_id, "<ButtonPress-1>", on_press)
        self.canvas.tag_bind(handle_id, "<ButtonRelease-1>", on_release)
        self.canvas.tag_bind(handle_id, "<B1-Motion>", on_move)

    def load_panel(self, name):
        self.panel_name = name
        self.canvas.delete("all")
        self.cubicles.clear()
        self.busbars.clear()

        with open(f"{PANELS_FOLDER}/{name}.json", "r") as f:
            panel_data = json.load(f)

        self.panel_depth = panel_data.get("panel_depth")

        for cub in panel_data.get("cubicles", []):
            rect = self.canvas.create_rectangle(*cub["coords"], fill=self.palette["cubicle_fill"], outline=self.palette["cubicle_outline"], width=3)
            cubicle_data = {
                "id": rect,
                "width": cub["width"],
                "height": cub["height"],
                "x": cub["coords"][0],
                "y": cub["coords"][1],
                "compartments": []
            }
            self.cubicles.append(cubicle_data)
            self.undo_stack.append({"type": "add_cubicle", "cubicle": cubicle_data})

            self.create_compartments(cubicle_data, len(cub["compartments"]))

            for comp_idx, saved_comp in enumerate(cub["compartments"]):
                for sec_idx, saved_sec in enumerate(saved_comp["sections"]):
                    item = saved_sec.get("item")
                    if item:
                        section = cubicle_data["compartments"][comp_idx]["sections"][sec_idx]
                        self.canvas.itemconfig(section["id"], fill=self.palette["section_selected"])
                        section["item"] = {"model": item["model"], "desc": item.get("desc", ""), "text_ids": []}
                        self.draw_vertical_text_in_section(section, item["model"], item.get("desc", ""))

        for busbar in panel_data.get("busbars", []):
            line = self.canvas.create_line(*busbar["coords"], fill="orange", width=3)
            self.canvas.tag_raise(line)
            busbar["id"] = line
            self.busbars.append(busbar)
            self.make_busbar_draggable(line, busbar["type"])
            self.make_busbar_resizable(line, busbar["type"])

        self.apply_theme()
        self.add_bottom_right_info()

        if self.panel_depth:
            try:
                self.canvas.create_text(20, 10, text=f"Depth: {self.panel_depth} mm", anchor="nw", font=("Arial", 10, "bold"))
            except Exception:
                pass

    def save_panel(self):
        if not self.panel_name:
            messagebox.showwarning("No Panel", "Please create or select a panel first.")
            return

        panel_data = {
            "project_info": {"customer": self.customer, "project": self.project, "ref": self.ref},
            "panel_depth": self.panel_depth,
            "cubicles": [],
            "busbars": self.busbars
        }

        for cub in self.cubicles:
            cub_data = {
                "coords": self.canvas.coords(cub["id"]),
                "width": cub["width"],
                "height": cub["height"],
                "color": self.canvas.itemcget(cub["id"], "fill"),
                "compartments": []
            }
            for comp in cub["compartments"]:
                comp_data = {"sections": []}
                for sec in comp["sections"]:
                    # Don't persist text_ids to keep save small
                    item = sec["item"]
                    if item:
                        comp_data["sections"].append({"name": sec["name"], "item": {"model": item["model"], "desc": item.get("desc", "")}})
                    else:
                        comp_data["sections"].append({"name": sec["name"], "item": None})
                cub_data["compartments"].append(comp_data)
            panel_data["cubicles"].append(cub_data)

        with open(f"{PANELS_FOLDER}/{self.panel_name}.json", "w") as f:
            json.dump(panel_data, f)

        messagebox.showinfo("Saved", f"Panel '{self.panel_name}' saved successfully!")
        self.refresh_panel_menu()

    def load_breaker_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path)
            if "Model No" not in df.columns or "Description" not in df.columns:
                messagebox.showerror("Error", "Excel must have 'Model No' and 'Description' columns.")
                return
            added = 0
            for _, row in df.iterrows():
                model = str(row["Model No"]).strip()
                desc = str(row["Description"]).strip()
                if model and model not in self.breaker_types:
                    self.breaker_types[model] = desc
                    added += 1
            if added > 0:
                self.save_breaker_types()
            messagebox.showinfo("Loaded", f"Added {added} new breaker types.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel: {e}")

    def generate_bom(self):
        if not self.cubicles:
            messagebox.showwarning("Generate BOM", "Please add cubicles and components first.")
            return

        # Google Sheets sync
        creds = get_credentials()
        client = gspread.authorize(creds)
        spreadsheet = None

        spreadsheet_name = f"{self.customer}_{self.project}_{self.ref}"
        try:
            spreadsheet = client.open(spreadsheet_name)
        except gspread.SpreadsheetNotFound:
            spreadsheet = client.create(spreadsheet_name)

        sheet_title = self.panel_name
        try:
            ws = spreadsheet.worksheet(sheet_title)
        except gspread.WorksheetNotFound:
            ws = spreadsheet.add_worksheet(title=sheet_title, rows="200", cols="20")

        data = [["Cubicle (X,Y)"] + SECTION_NAMES]
        for cub_idx, cub in enumerate(self.cubicles, start=1):
            for comp_idx, comp in enumerate(cub["compartments"], start=1):
                row = [f"{cub_idx},{comp_idx}"]
                for section_name in SECTION_NAMES:
                    item = next((sec["item"] for sec in comp["sections"] if sec["name"] == section_name and sec["item"]), None)
                    row.append(item["model"] if item else "")
                data.append(row)

        data.append([])
        data.append(["Busbars"])
        data.append(["Type", "Amperage (A)", "Current Density (A/mm²)", "Coordinates (x1, y1, x2, y2)", "Phase", "Busbar Size", "No. of Runs", "Busbar Length (mm)"])
        for bus in self.busbars:
            coords = tuple(map(int, bus["coords"]))
            length = (coords[2] - coords[0]) if bus["type"] == "horizontal" else (coords[3] - coords[1])
            data.append([bus.get("type"), bus.get("amperage"), bus.get("current_density"), str(coords), bus.get("phase"), bus.get("busbar_size", ""), bus.get("no_of_runs", ""), length])

        ws.update(values=data, range_name="A1")

        # Totals across project
        part_totals = defaultdict(lambda: {"desc": "", "total": 0, "panels": defaultdict(int)})
        category_totals = defaultdict(lambda: defaultdict(lambda: {"desc": "", "total": 0, "panels": defaultdict(int)}))
        busbar_totals = defaultdict(lambda: {"total": 0, "panels": defaultdict(int), "desc": ""})

        panel_files = [f for f in os.listdir(PANELS_FOLDER) if f.endswith(".json")]
        relevant_panels = []
        no_match_counter = 0

        for fname in panel_files:
            fpath = os.path.join(PANELS_FOLDER, fname)
            with open(fpath, "r") as f:
                panel_data = json.load(f)
                info = panel_data.get("project_info", {})
                if info.get("customer") == self.customer and info.get("project") == self.project and info.get("ref") == self.ref:
                    pname = fname[:-5]
                    relevant_panels.append(pname)

                    for cub in panel_data.get("cubicles", []):
                        for comp in cub.get("compartments", []):
                            for sec in comp.get("sections", []):
                                item = sec.get("item")
                                if item:
                                    model = item["model"]
                                    desc = item.get("desc", "")
                                    category = sec.get("name", "Others")
                                    part_totals[model]["desc"] = desc
                                    part_totals[model]["total"] += 1
                                    part_totals[model]["panels"][pname] += 1

                                    cat_bucket = category_totals[category][model]
                                    cat_bucket["desc"] = desc
                                    cat_bucket["total"] += 1
                                    cat_bucket["panels"][pname] += 1

                    for busbar in panel_data.get("busbars", []):
                        amp = busbar.get("amperage")
                        cd = busbar.get("current_density")
                        coords = busbar.get("coords", [0, 0, 0, 0])
                        phase = busbar.get("phase", "Single Phase")
                        busbar_size_str = busbar.get("busbar_size", "")
                        try:
                            no_of_runs = int(busbar.get("no_of_runs", 1))
                        except Exception:
                            no_of_runs = 1

                        length = (coords[2] - coords[0]) if busbar.get("type") == "horizontal" else (coords[3] - coords[1])

                        if busbar_size_str:
                            bus_part_no = busbar_size_str
                            bus_desc = busbar_size_str
                            qty = max(0, int(length)) * no_of_runs
                            busbar_totals[bus_part_no]["desc"] = bus_desc
                            busbar_totals[bus_part_no]["total"] += qty
                            busbar_totals[bus_part_no]["panels"][pname] += qty

                        elif cd is not None and amp is not None and cd > 0 and amp > 0:
                            area_needed = amp / cd
                            nearest_busbar = self.find_nearest_highest_busbar(area_needed)

                            if nearest_busbar is None:
                                bus_part_no = f"NO_MATCH_{no_match_counter}"
                                bus_desc = f"No match for Phase={phase}, Amperage={amp}, CD={cd}, AreaNeeded={area_needed:.2f}"
                                qty = 0
                                no_match_counter += 1
                            else:
                                bus_part_no = nearest_busbar["Part no"]
                                bus_desc = nearest_busbar["Item description"]
                                bus_runs = int(nearest_busbar["No. of runs"]) if "No. of runs" in nearest_busbar else 1
                                base_qty = max(0, int(length)) * bus_runs
                                multiplier = 2 if phase == "Single Phase" else 4
                                qty = base_qty * multiplier

                            busbar_totals[bus_part_no]["desc"] = bus_desc
                            busbar_totals[bus_part_no]["total"] += qty
                            busbar_totals[bus_part_no]["panels"][pname] += qty

        # Total BOM sheet
        try:
            total_ws = spreadsheet.worksheet("Total BOM")
        except gspread.WorksheetNotFound:
            total_ws = spreadsheet.add_worksheet(title="Total BOM", rows="200", cols="30")

        header = ["Part No.", "Description", "Total Qty"] + relevant_panels
        total_data = [header]

        for model, info in part_totals.items():
            row = [model, info["desc"], info["total"]]
            for pname in relevant_panels:
                row.append(info["panels"].get(pname, 0))
            total_data.append(row)

        total_data.append([])
        total_data.append(["Busbar Materials"])
        total_data.append(header)
        for model, info in busbar_totals.items():
            row = [model, info["desc"], info["total"]]
            for pname in relevant_panels:
                row.append(info["panels"].get(pname, 0))
            total_data.append(row)

        serializable_data = []
        for row in total_data:
            new_row = []
            for item in row:
                if isinstance(item, np.generic):
                    new_row.append(item.item())
                else:
                    new_row.append(item)
            serializable_data.append(new_row)

        total_ws.clear()
        total_ws.update(values=serializable_data, range_name="A1")
        header_format = {"backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8}, "horizontalAlignment": "CENTER", "textFormat": {"bold": True}}
        total_ws.format("A1:Z1", header_format)

        # Grouped PDF
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        project_folder = os.path.join(desktop_path, self.project)
        os.makedirs(project_folder, exist_ok=True)

        pdf_path = os.path.join(project_folder, "Total_BOM.pdf")
        doc = SimpleDocTemplate(pdf_path, pagesize=A4, rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)

        styles = getSampleStyleSheet()
        elements = []

        header_table_data = []
        logo_path = resource_path("VLPP.ico")
        if os.path.exists(logo_path):
            header_logo = RLImage(logo_path, width=40, height=40)
        else:
            header_logo = Paragraph("", styles["Normal"])

        header_email = Paragraph("<b>venora@gmail.com</b>", styles["Normal"])
        header_table_data.append([header_logo, header_email])

        header_table = Table(header_table_data, colWidths=[60, 440])
        header_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (1, 0), (1, 0), "RIGHT")]))
        elements.append(header_table)
        elements.append(Spacer(1, 8))

        project_style = ParagraphStyle("ProjectInfo", parent=styles["Normal"], fontSize=10, leading=13, spaceAfter=6)
        project_info_text = (f"<b>Customer:</b> {self.customer}<br/>"
                             f"<b>Project:</b> {self.project}<br/>"
                             f"<b>Reference:</b> {self.ref}")
        project_info_para = Paragraph(project_info_text, project_style)
        elements.append(project_info_para)
        elements.append(Spacer(1, 8))

        title = Paragraph("<b>Total Bill of Materials (BOM)</b>", styles["Title"])
        elements.append(title)
        elements.append(Spacer(1, 12))

        def build_table(rows):
            table = Table(rows, repeatRows=1)
            table_style = TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
            ])
            table.setStyle(table_style)
            for i in range(1, len(rows)):
                if i % 2 == 0:
                    table.setStyle(TableStyle([("BACKGROUND", (0, i), (-1, i), colors.whitesmoke)]))
                else:
                    table.setStyle(TableStyle([("BACKGROUND", (0, i), (-1, i), colors.beige)]))
            return table

        ordered_categories = ["Breaker", "ELR/EFR", "PFR", "Power Analyzer/Energy Meter", "Indicator Light", "SPD"]
        for cat in ordered_categories:
            items = category_totals.get(cat, {})
            if not items:
                continue
            elements.append(Paragraph(f"<b>{cat}</b>", styles["Heading2"]))
            header_row = ["Part No.", "Description", "Total Qty"] + relevant_panels
            rows = [header_row]
            for model, info in items.items():
                row = [model, info["desc"], int(info["total"])]
                for pname in relevant_panels:
                    row.append(int(info["panels"].get(pname, 0)))
                rows.append(row)
            elements.append(build_table(rows))
            elements.append(Spacer(1, 12))

        if busbar_totals:
            elements.append(Paragraph("<b>Busbar Materials</b>", styles["Heading2"]))
            header_row = ["Part No.", "Description", "Total Qty"] + relevant_panels
            rows = [header_row]
            for model, info in busbar_totals.items():
                row = [model, info["desc"], int(info["total"])]
                for pname in relevant_panels:
                    row.append(int(info["panels"].get(pname, 0)))
                rows.append(row)
            elements.append(build_table(rows))

        doc.build(elements)

        try:
            if os.name == "nt":
                os.startfile(pdf_path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", pdf_path])
            else:
                subprocess.Popen(["xdg-open", pdf_path])
        except Exception as e:
            print("Could not open PDF automatically:", e)

        messagebox.showinfo("PDF Saved", f"Total BOM PDF saved to:\n{pdf_path}")
        total_ws.format("A1:Z1", header_format)
        messagebox.showinfo("BOM Generated", "BOM added to Google Sheets and grouped PDF created!")

    def find_nearest_highest_busbar(self, area_value):
        if self.busbar_data.empty:
            return None
        try:
            filtered = self.busbar_data[self.busbar_data['Area (sqmm)'] >= area_value]
        except Exception:
            return None
        if filtered.empty:
            return None
        match_row = filtered.loc[filtered['Area (sqmm)'].idxmin()]
        return match_row

    def undo_last_action(self):
        if not self.undo_stack:
            messagebox.showinfo("Undo", "Nothing to undo.")
            return

        action = self.undo_stack.pop()
        if action["type"] == "add_cubicle":
            self.canvas.delete(action["cubicle"]["id"])
            for comp in action["cubicle"]["compartments"]:
                for sec in comp["sections"]:
                    self.canvas.delete(sec["id"])
            self.cubicles.remove(action["cubicle"])
        elif action["type"] == "add_busbar":
            self.canvas.delete(action["busbar"]["id"])
            self.busbars.remove(action["busbar"])
        elif action["type"] == "select_component":
            section = action["section"]
            # remove current text ids
            try:
                for tid in action.get("new_text_ids", []):
                    self.canvas.delete(tid)
            except Exception:
                pass
            section["item"] = action["previous_item"]
            # recolor
            self.canvas.itemconfig(action["rect_id"], fill=self.palette["section_empty"] if not action["previous_item"] else self.palette["section_selected"])
            # if previous item existed, redraw its text
            if action["previous_item"]:
                section["item"] = {"model": action["previous_item"]["model"], "desc": action["previous_item"].get("desc", ""), "text_ids": []}
                tids = self.draw_vertical_text_in_section(section, action["previous_item"]["model"], action["previous_item"].get("desc", ""))
                section["item"]["text_ids"] = tids
        messagebox.showinfo("Undo", "Last action undone.")

    # ================= THEME HELPERS =================
    def get_palette(self, mode="light"):
        if mode == "dark":
            return {
                "bg": "#1f1f1f",
                "canvas_bg": "#0f1115",
                "text": "#e5e7eb",
                "muted_text": "#9ca3af",
                "cubicle_fill": "#1f2a44",
                "cubicle_outline": "#9ca3af",
                "section_empty": "#111827",
                "section_selected": "#166534",
                "section_outline": "#374151",
                "busbar": "#f59e0b",
                "busbar_terminal": "#a855f7",
                "handle": "#ef4444",
            }
        else:
            return {
                "bg": "#ffffff",
                "canvas_bg": "lightgray",
                "text": "#111827",
                "muted_text": "#374151",
                "cubicle_fill": "lightblue",
                "cubicle_outline": "#111827",
                "section_empty": "#ffffff",
                "section_selected": "#bbf7d0",
                "section_outline": "#d1d5db",
                "busbar": "orange",
                "busbar_terminal": "purple",
                "handle": "red",
            }

    def apply_theme(self):
        try:
            self.style.theme_use("clam" if self.is_dark_mode else "default")
        except Exception:
            pass
        try:
            self.root.configure(bg=self.palette["bg"])
        except Exception:
            pass
        try:
            self.canvas.configure(bg=self.palette["canvas_bg"])
        except Exception:
            pass

        for cub in self.cubicles:
            try:
                self.canvas.itemconfig(cub["id"], fill=self.palette["cubicle_fill"], outline=self.palette["cubicle_outline"])
            except Exception:
                pass
            for comp in cub["compartments"]:
                for sec in comp["sections"]:
                    try:
                        fill = self.palette["section_selected"] if sec.get("item") else self.palette["section_empty"]
                        self.canvas.itemconfig(sec["id"], fill=fill, outline=self.palette["section_outline"])
                        # recolor text if exists
                        if sec.get("item") and sec["item"].get("text_ids"):
                            for tid in sec["item"]["text_ids"]:
                                self.canvas.itemconfig(tid, fill=self.palette["text"])
                    except Exception:
                        pass

        for b in self.busbars:
            try:
                color = self.palette["busbar_terminal"] if b.get("busbar_size") else self.palette["busbar"]
                self.canvas.itemconfig(b["id"], fill=color)
            except Exception:
                pass

        try:
            for h in self.canvas.find_withtag("handle"):
                self.canvas.itemconfig(h, fill=self.palette["handle"])
        except Exception:
            pass

        try:
            self.add_bottom_right_info()
        except Exception:
            pass

        # sync toggle label to current mode
        try:
            self.dark_mode_var.set(self.is_dark_mode)
            self.dark_toggle.config(text="Dark Mode" if not self.is_dark_mode else "Light Mode")
        except Exception:
            pass

    def set_light_mode(self):
        self.is_dark_mode = False
        self.palette = self.get_palette("light")
        self.apply_theme()

    def set_dark_mode(self):
        self.is_dark_mode = True
        self.palette = self.get_palette("dark")
        self.apply_theme()

    def toggle_theme(self):
        if self.is_dark_mode:
            self.set_light_mode()
        else:
            self.set_dark_mode()

    def toggle_theme_check(self):
        # Called by the Checkbutton
        if self.dark_mode_var.get():
            self.set_dark_mode()
        else:
            self.set_light_mode()


def get_credentials():
    creds = None
    try:
        if os.path.exists(TOKEN_FILE):
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

        if creds and creds.valid:
            return creds

        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                with open(TOKEN_FILE, "w") as token:
                    token.write(creds.to_json())
                return creds
            except Exception as e:
                print("Refresh failed, regenerating token.json:", e)
                try:
                    os.remove(TOKEN_FILE)
                except Exception:
                    pass

        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())
        return creds

    except Exception as e:
        print("Credential error, regenerating:", e)
        try:
            if os.path.exists(TOKEN_FILE):
                os.remove(TOKEN_FILE)
        except Exception:
            pass
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())
        return creds


def load_all_projects():
    projects = {}
    os.makedirs(PANELS_FOLDER, exist_ok=True)
    for file in os.listdir(PANELS_FOLDER):
        if file.endswith(".json"):
            try:
                with open(f"{PANELS_FOLDER}/{file}", "r") as f:
                    data = json.load(f)
                    pinfo = data.get("project_info", {})
                    c, p, r = pinfo.get("customer", "").strip(), pinfo.get("project", "").strip(), pinfo.get("ref", "").strip()
                    if c and p and r:
                        key = (c, p, r)
                        display_name = f"{c} | {p} | {r}"
                        projects[key] = display_name
            except Exception:
                continue
    return sorted(projects.values()), projects


def startup_screen():
    root = tk.Tk()
    root.title("Welcome")
    root.configure(bg="#faebd7")
    root.geometry("500x400")
    root.resizable(False, False)

    result = {}

    # --- Frame ---
    frm = tk.Frame(root, bg="#faebd7")
    frm.pack(expand=True, fill="both")

    # --- Logo ---
    try:
        logo = tk.PhotoImage(file=resource_path("VLPP.ico"))
        logo_label = tk.Label(frm, image=logo, bg="#faebd7")
        logo_label.image = logo
        logo_label.pack(pady=10)
    except Exception:
        pass

    tk.Label(frm, text="WELCOME", font=("Arial", 16, "bold"), bg="#faebd7").pack()
    tk.Label(frm, text="PANEL DESIGNER", font=("Arial", 12), bg="#faebd7").pack(pady=5)

    # --- Button Actions ---
    def show_create_new():
        frm.pack_forget()
        create_frame()

    def show_open_file():
        frm.pack_forget()
        open_frame()

    tk.Button(frm, text="CREATE NEW", width=20, command=show_create_new).pack(pady=10)
    tk.Button(frm, text="OPEN FILE", width=20, command=show_open_file).pack(pady=5)

    # --- Footer ---
    footer = tk.Frame(root, bg="#faebd7")
    footer.pack(side="bottom", anchor="se", padx=10, pady=5)
    tk.Label(footer, text="venora@gmail.com", bg="#faebd7", font=("Arial", 9)).pack(anchor="e")
    tk.Label(footer, text="0764319139", bg="#faebd7", font=("Arial", 9)).pack(anchor="e")

    # --- Create New Frame ---
    def create_frame():
        newf = tk.Frame(root, bg="#faebd7")
        newf.pack(expand=True)

        tk.Label(newf, text="PANEL DESIGNER", font=("Arial", 14, "bold"), bg="#faebd7").grid(row=0, column=0, columnspan=2, pady=10)

        tk.Label(newf, text="Project Name:", bg="#faebd7").grid(row=1, column=0, sticky="e", pady=5, padx=5)
        proj_var = tk.StringVar()
        tk.Entry(newf, textvariable=proj_var).grid(row=1, column=1, pady=5)

        tk.Label(newf, text="Customer Name:", bg="#faebd7").grid(row=2, column=0, sticky="e", pady=5, padx=5)
        cust_var = tk.StringVar()
        tk.Entry(newf, textvariable=cust_var).grid(row=2, column=1, pady=5)

        tk.Label(newf, text="Our Reference:", bg="#faebd7").grid(row=3, column=0, sticky="e", pady=5, padx=5)
        ref_var = tk.StringVar()
        tk.Entry(newf, textvariable=ref_var).grid(row=3, column=1, pady=5)

        def create_action():
            if not proj_var.get().strip() or not cust_var.get().strip() or not ref_var.get().strip():
                messagebox.showerror("Error", "Fill all fields")
                return
            result["project"] = proj_var.get().strip()
            result["customer"] = cust_var.get().strip()
            result["ref"] = ref_var.get().strip()
            root.destroy()

        tk.Button(newf, text="CREATE", command=create_action).grid(row=4, column=0, pady=15)
        tk.Button(newf, text="BACK", command=lambda:[newf.destroy(), frm.pack(expand=True)]).grid(row=4, column=1)

    # --- Open File Frame ---
    def open_frame():
        of = tk.Frame(root, bg="#faebd7")
        of.pack(expand=True)

        tk.Label(of, text="Open Project", font=("Arial", 14, "bold"), bg="#faebd7").pack(pady=10)

        project_names, project_map = load_all_projects()
        selected_var = tk.StringVar()
        combo = ttk.Combobox(of, textvariable=selected_var, values=project_names, state="readonly", width=40)
        combo.pack(pady=10)

        def open_action():
            sel = selected_var.get()
            for key, name in project_map.items():
                if name == sel:
                    c, p, r = key
                    result["customer"] = c
                    result["project"] = p
                    result["ref"] = r
                    root.destroy()
                    return
            messagebox.showerror("Error", "Select a project")

        tk.Button(of, text="OPEN", command=open_action).pack(pady=10)
        tk.Button(of, text="BACK", command=lambda:[of.destroy(), frm.pack(expand=True)]).pack()

    root.protocol("WM_DELETE_WINDOW", lambda: exit(0))
    root.mainloop()
    return result


if __name__ == "__main__":
    project_info = startup_screen()
    root = tk.Tk()
    window_width, window_height = 1200, 700
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = int((screen_width / 2) - (window_width / 2))
    y = int((screen_height / 2) - (window_height / 2))
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.minsize(1000, 600)
    app = PanelDesigner(root, project_info["customer"], project_info["project"], project_info["ref"])
    root.mainloop()
