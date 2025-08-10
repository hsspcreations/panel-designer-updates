import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox, ttk
import json
import pandas as pd
import numpy as np # <-- Added import for numpy
import os
import sys
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from PIL import Image, ImageTk
from collections import defaultdict

BREAKER_FILE = "breaker_types.json"
APPDATA_FOLDER = os.path.join(os.environ.get("APPDATA"), "PanelDesigner")
PANELS_FOLDER = os.path.join(APPDATA_FOLDER, "panels")
os.makedirs(PANELS_FOLDER, exist_ok=True)
TOKEN_FILE = "token.json"

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
        self.root.iconbitmap(resource_path("Hssp.ico"))

        self.breaker_types = self.load_breaker_types()
        self.busbar_data = self.load_busbar_data()
        self.saved_panels = self.load_saved_panels()
        self.panel_name = None
        self.cubicles = []
        self.busbars = []
        self.tooltip = None
        self.icon_image = None
        self.drag_data = {"item": None, "x": 0, "y": 0}
        self.undo_stack = []

        top_frame = tk.Frame(root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        self.panel_var = tk.StringVar()
        self.panel_var.set("Select Panel" if self.saved_panels else "No Panels")
        self.panel_menu = tk.OptionMenu(top_frame, self.panel_var, *(["Select Panel"] + self.saved_panels),
                                        command=self.on_panel_select)
        self.panel_menu.pack(side=tk.LEFT, padx=5)

        tk.Button(top_frame, text="Create Panel", command=self.create_panel).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Add Cubicle", command=self.add_cubicle).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Add Vertical Busbar", command=self.add_vertical_busbar_form).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Add Horizontal Busbar", command=self.add_horizontal_busbar_form).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Upload Breaker Types", command=self.load_breaker_excel).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Save Panel", command=self.save_panel).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Generate BOM", command=self.generate_bom).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Undo", command=self.undo_last_action).pack(side=tk.LEFT, padx=5)

        canvas_frame = tk.Frame(root)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.canvas = tk.Canvas(canvas_frame, bg="lightgray", width=1000, height=600, scrollregion=(0, 0, 5000, 3000))
        h_scroll = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        v_scroll = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.add_bottom_right_info()

    def add_bottom_right_info(self):
        canvas_width = int(self.canvas['width'])
        canvas_height = int(self.canvas['height'])
        padding = 10

        icon_path = resource_path("Hssp.ico")
        if os.path.exists(icon_path):
            try:
                img = Image.open(icon_path)
                img = img.resize((32, 32))
                self.icon_image = ImageTk.PhotoImage(img)
                self.canvas.create_image(canvas_width - 40, canvas_height - 60, image=self.icon_image, anchor="se")
            except:
                pass

        self.canvas.create_text(canvas_width - padding, canvas_height - 30,
                                text="hsspcreations@gmail.com", font=("Arial", 10), fill="black", anchor="se")
        self.canvas.create_text(canvas_width - padding, canvas_height - 10,
                                text="0764319139", font=("Arial", 10), fill="black", anchor="se")

    def load_breaker_types(self):
        if os.path.exists(BREAKER_FILE):
            try:
                with open(BREAKER_FILE, "r") as f:
                    return json.load(f)
            except:
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
                except:
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
        name = simpledialog.askstring("Panel Name", "Enter a name for the new panel:")
        if name:
            self.panel_name = name
            self.cubicles.clear()
            self.busbars.clear()
            self.canvas.delete("all")
            self.panel_var.set(name)
            self.add_bottom_right_info()

    def add_cubicle(self):
        if not self.panel_name:
            messagebox.showwarning("No Panel", "Please create or select a panel first.")
            return

        # Create the cubicle size selection window
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
                x = last["x"] + last["width"] * SCALE  # No gap between cubicles
                y = last["y"]

            # Use a fixed color
            rect = self.canvas.create_rectangle(x, y, x + w, y + h, fill="lightblue", outline="black", width=3)

            cubicle_data = { "id": rect, "width": width, "height": height, "x": x, "y": y, "compartments": [] }
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

        # Select the last added cubicle to delete for simplicity
        cubicle = self.cubicles.pop()
        self.canvas.delete(cubicle["id"])

        # Delete all compartments and their sections
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
                section_rect = self.canvas.create_rectangle(sec_x1, comp_y1, sec_x2, comp_y2, fill="white")
                self.canvas.tag_bind(section_rect, "<Button-1>",
                                     lambda e, s=section_name, r=section_rect, c=compartment: self.select_item(s, r, c))
                compartment["sections"].append({"name": section_name, "id": section_rect, "item": None})

            cubicle["compartments"].append(compartment)

    def select_item(self, section_name, section_rect, compartment):
        self.show_search_popup(section_name, section_rect, compartment)

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

                self.canvas.itemconfig(section_rect, fill="lightgreen")
                coords = self.canvas.coords(section_rect)
                x = (coords[0] + coords[2]) / 2
                y = (coords[1] + coords[3]) / 2
                vertical_text = "\n".join(model)
                text_id = self.canvas.create_text(x, y, text=vertical_text, font=("Arial", 6))
                self.canvas.tag_bind(text_id, "<Enter>", lambda e, d=desc: self.show_tooltip(e, d))
                self.canvas.tag_bind(text_id, "<Leave>", lambda e: self.hide_tooltip())

                
                for section in compartment["sections"]:
                    if section["id"] == section_rect:
                        previous_item = section["item"]
                        section["item"] = {"model": model, "desc": desc}
                        self.undo_stack.append({
                            "type": "select_component",
                            "section": section,
                            "previous_item": previous_item,
                            "text_id": text_id,
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
        line_id = self.canvas.create_line(x, y1, x, y2, fill="orange", width=4)
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
        line_id = self.canvas.create_line(x1, y, x2, y, fill="orange", width=4)
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
        if busbar_type == "vertical":
            handle_id = self.canvas.create_rectangle(coords[2] - handle_size, coords[3] - handle_size,
                                                     coords[2] + handle_size, coords[3] + handle_size, fill="red")
        else:
            handle_id = self.canvas.create_rectangle(coords[2] - handle_size, coords[3] - handle_size,
                                                     coords[2] + handle_size, coords[3] + handle_size, fill="red")

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

        for cub in panel_data.get("cubicles", []):
            rect = self.canvas.create_rectangle(*cub["coords"], fill=cub.get("color", "lightblue"))
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

            # Create compartments
            self.create_compartments(cubicle_data, len(cub["compartments"]))

            # Restore breakers
            for comp_idx, saved_comp in enumerate(cub["compartments"]):
                for sec_idx, saved_sec in enumerate(saved_comp["sections"]):
                    item = saved_sec.get("item")
                    if item:
                        section = cubicle_data["compartments"][comp_idx]["sections"][sec_idx]
                        self.canvas.itemconfig(section["id"], fill="lightgreen")
                        coords = self.canvas.coords(section["id"])
                        x = (coords[0] + coords[2]) / 2
                        y = (coords[1] + coords[3]) / 2
                        text_id = self.canvas.create_text(x, y, text="\n".join(item["model"]), font=("Arial", 6))
                        self.canvas.tag_bind(text_id, "<Enter>", lambda e, d=item["desc"]: self.show_tooltip(e, d))
                        self.canvas.tag_bind(text_id, "<Leave>", lambda e: self.hide_tooltip())
                        section["item"] = item

        # Restore busbars
        for busbar in panel_data.get("busbars", []):
            line = self.canvas.create_line(*busbar["coords"], fill="orange", width=3)
            self.canvas.tag_raise(line)
            busbar["id"] = line
            self.busbars.append(busbar)
            self.make_busbar_draggable(line, busbar["type"])
            self.make_busbar_resizable(line, busbar["type"])

        self.add_bottom_right_info()
    def save_panel(self):
        if not self.panel_name:
            messagebox.showwarning("No Panel", "Please create or select a panel first.")
            return

        panel_data = {
            "project_info": {"customer": self.customer, "project": self.project, "ref": self.ref},
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
                    comp_data["sections"].append({"name": sec["name"], "item": sec["item"]})
                cub_data["compartments"].append(comp_data)
            panel_data["cubicles"].append(cub_data)

        with open(f"{PANELS_FOLDER}/{self.panel_name}.json", "w") as f:
            json.dump(panel_data, f)

        messagebox.showinfo("Saved", f"Panel '{self.panel_name}' saved successfully!")
        self.refresh_panel_menu()
        self.generate_bom()

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

    # Corrected generate_bom method
    def generate_bom(self):
        if not self.cubicles:
            messagebox.showwarning("Generate BOM", "Please add cubicles and components first.")
            return

        # Initialize the Google Sheets client and spreadsheet variable.
        # This prevents the NameError from occurring.
        creds = get_credentials()
        client = gspread.authorize(creds)
        spreadsheet = None  # Ensure spreadsheet is always defined.

        spreadsheet_name = f"{self.customer}_{self.project}_{self.ref}"
        try:
            spreadsheet = client.open(spreadsheet_name)
        except gspread.SpreadsheetNotFound:
            spreadsheet = client.create(spreadsheet_name)

        # --- 1. Current Panel BOM ---
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
        data.append(["Type", "Amperage (A)", "Current Density (A/mm²)", "Coordinates (x1, y1, x2, y2)", "Phase", "Busbar Length (mm)"])
        for bus in self.busbars:
            coords = tuple(map(int, bus["coords"]))
            length = (coords[2] - coords[0]) if bus["type"] == "horizontal" else (coords[3] - coords[1])
            data.append([bus["type"], bus["amperage"], bus["current_density"], str(coords), bus["phase"], length])
        
        # Fixed: Use named arguments for the update() method
        ws.update(values=data, range_name="A1")

        # --- 2. Total BOM Aggregation ---
        
        part_totals = defaultdict(lambda: {"desc": "", "total": 0, "panels": defaultdict(int)})
        busbar_totals = defaultdict(lambda: {"total": 0, "panels": defaultdict(int), "desc": "", "no_of_runs": 0})
        
        panel_files = [f for f in os.listdir(PANELS_FOLDER) if f.endswith(".json")]
        relevant_panels = []
        for fname in panel_files:
            fpath = os.path.join(PANELS_FOLDER, fname)
            with open(fpath, "r") as f:
                panel_data = json.load(f)
                info = panel_data.get("project_info", {})
                if info.get("customer") == self.customer and info.get("project") == self.project and info.get("ref") == self.ref:
                    pname = fname[:-5]
                    relevant_panels.append(pname)
                    
                    # Aggregate breakers
                    for cub in panel_data.get("cubicles", []):
                        for comp in cub.get("compartments", []):
                            for sec in comp.get("sections", []):
                                item = sec.get("item")
                                if item:
                                    model = item["model"]
                                    desc = item["desc"]
                                    part_totals[model]["desc"] = desc
                                    part_totals[model]["total"] += 1
                                    part_totals[model]["panels"][pname] += 1
                    
                    # Aggregate busbars
                    for busbar in panel_data.get("busbars", []):
                        amp = busbar["amperage"]
                        cd = busbar["current_density"]
                        coords = busbar["coords"]
                        phase = busbar["phase"]
                        
                        # Calculate busbar length
                        length = (coords[2] - coords[0]) if busbar["type"] == "horizontal" else (coords[3] - coords[1])
                        
                        if cd > 0 and amp > 0:
                            # Step 1: Area needed
                            area_needed = amp / cd
                            print(f"📏 Area needed: {amp} ÷ {cd} = {area_needed}")

                            # Step 2: Nearest highest busbar
                            nearest_busbar = self.find_nearest_highest_busbar(area_needed)
                            
                            if nearest_busbar is None:
                                bus_part_no = f"NO_MATCH_{no_match_counter}"
                                bus_desc = f"No match for Phase={phase}, Amperage={amp}, CD={cd}, AreaNeeded={area_needed:.2f}"
                                qty = 0
                                no_match_counter += 1
                            else:
                                bus_part_no = nearest_busbar["Part no"]
                                bus_desc = nearest_busbar["Item description"]
                                bus_runs = nearest_busbar["No. of runs"]
                                print(f"🔍 Match: {bus_part_no}, Area={nearest_busbar['Area (sqmm)']}, Runs={bus_runs}")

                                # Step 3: Base qty = length × runs
                                base_qty = length * bus_runs
                                print(f"📐 Base qty: {length} × {bus_runs} → {base_qty}")

                                # Step 4: Apply phase multiplier
                                multiplier = 2 if phase == "Single Phase" else 4
                                qty = base_qty * multiplier
                                print(f"✅ Final qty: {base_qty} × {multiplier} → {qty}")

                            busbar_totals[bus_part_no]["desc"] = bus_desc
                            busbar_totals[bus_part_no]["total"] += qty
                            busbar_totals[bus_part_no]["panels"][pname] += qty

        # Create or update "Total BOM" sheet
        try:
            total_ws = spreadsheet.worksheet("Total BOM")
        except gspread.WorksheetNotFound:
            total_ws = spreadsheet.add_worksheet(title="Total BOM", rows="200", cols="30")

        header = ["Part No.", "Description", "Total Qty"] + relevant_panels
        total_data = [header]
        
        # Add breaker totals
        for model, info in part_totals.items():
            row = [model, info["desc"], info["total"]]
            for pname in relevant_panels:
                row.append(info["panels"].get(pname, 0))
            total_data.append(row)
        
        # Add a separator and busbar totals
        total_data.append([])
        total_data.append(["Busbar Materials"])
        total_data.append(header)
        for model, info in busbar_totals.items():
            row = [model, info["desc"], info["total"]]
            for pname in relevant_panels:
                row.append(info["panels"].get(pname, 0))
            total_data.append(row)

        # Fixed: Convert numpy types to native Python types
        serializable_data = []
        for row in total_data:
            new_row = []
            for item in row:
                if isinstance(item, np.int64):
                    new_row.append(int(item))
                else:
                    new_row.append(item)
            serializable_data.append(new_row)

        total_ws.clear()
        
        # Fixed: Use named arguments for the update() method
        total_ws.update(values=serializable_data, range_name="A1")

        # Gray header formatting
        header_format = {
            "backgroundColor": {
                "red": 0.8,
                "green": 0.8,
                "blue": 0.8
            },
            "horizontalAlignment": "CENTER",
            "textFormat": {
                "bold": True
            }
        }
        total_ws.format("A1:Z1", header_format)
        total_ws.format(f"A{len(total_data) - len(busbar_totals) + 1}:Z{len(total_data) - len(busbar_totals) + 1}", header_format)
        
        messagebox.showinfo("BOM Generated", "BOM added to Google Sheets, including 'Total BOM' sheet!")
    
    def find_nearest_highest_busbar(self, area_value):
        if self.busbar_data.empty:
            print("❌ Busbar data is empty!")
            return None

        # Filter to nearest highest or equal
        filtered = self.busbar_data[self.busbar_data['Area (sqmm)'] >= area_value]
        if filtered.empty:
            print(f"⚠️ No match found for required area {area_value}")
            return None
        
        match_row = filtered.loc[filtered['Area (sqmm)'].idxmin()]
        print(f"🔍 Match for target {area_value}: {match_row['Part no']} | {match_row['Item description']} | Area={match_row['Area (sqmm)']} | Runs={match_row['No. of runs']}")
        return match_row
        
        filtered_data = self.busbar_data[self.busbar_data['Area (sqmm)'] >= area_value]
        if filtered_data.empty:
            return pd.Series()
        
        return filtered_data.loc[filtered_data['Area (sqmm)'].idxmin()]


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
            section["item"] = action["previous_item"]
            self.canvas.delete(action["text_id"])
            self.canvas.itemconfig(action["rect_id"], fill="white" if not action["previous_item"] else "lightgreen")
        messagebox.showinfo("Undo", "Last action undone.")


def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
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
            except:
                continue
    return sorted(projects.values()), projects


def startup_screen():
    root = tk.Tk()
    root.title("Project Info")

    tk.Label(root, text="Customer:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    customer_var = tk.StringVar()
    tk.Entry(root, textvariable=customer_var).grid(row=0, column=1, padx=5, pady=5)

    tk.Label(root, text="Project:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    project_var = tk.StringVar()
    tk.Entry(root, textvariable=project_var).grid(row=1, column=1, padx=5, pady=5)

    tk.Label(root, text="Our Ref:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
    ref_var = tk.StringVar()
    tk.Entry(root, textvariable=ref_var).grid(row=2, column=1, padx=5, pady=5)

    project_names, project_map = load_all_projects()

    selected_project_var = tk.StringVar()
    tk.Label(root, text="Open Existing Project:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
    project_dropdown = ttk.Combobox(root, textvariable=selected_project_var, values=project_names, state="readonly")
    project_dropdown.grid(row=3, column=1, padx=5, pady=5)

    action_var = tk.StringVar(value="create")
    tk.Radiobutton(root, text="Create New Project", variable=action_var, value="create").grid(row=4, column=0, columnspan=2, pady=5)
    tk.Radiobutton(root, text="Open Existing Project", variable=action_var, value="open").grid(row=5, column=0, columnspan=2, pady=5)

    result = {}

    def on_confirm():
        if action_var.get() == "create":
            if not customer_var.get().strip() or not project_var.get().strip() or not ref_var.get().strip():
                messagebox.showerror("Error", "Please fill all fields.")
                return
            result["customer"] = customer_var.get().strip()
            result["project"] = project_var.get().strip()
            result["ref"] = ref_var.get().strip()
            root.destroy()
        else:
            selected = selected_project_var.get()
            for key, display_name in project_map.items():
                if display_name == selected:
                    c, p, r = key
                    result["customer"] = c
                    result["project"] = p
                    result["ref"] = r
                    root.destroy()
                    return
            messagebox.showerror("Error", "Please select an existing project.")

    tk.Button(root, text="Confirm", command=on_confirm).grid(row=6, column=0, columnspan=2, pady=10)

    root.protocol("WM_DELETE_WINDOW", lambda: exit(0))
    root.mainloop()

    return result


if __name__ == "__main__":
    project_info = startup_screen()
    root = tk.Tk()
    app = PanelDesigner(root, project_info["customer"], project_info["project"], project_info["ref"])
    root.mainloop()

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
            section["item"] = action["previous_item"]
            self.canvas.delete(action["text_id"])
            self.canvas.itemconfig(action["rect_id"], fill="white" if not action["previous_item"] else "lightgreen")
        messagebox.showinfo("Undo", "Last action undone.")