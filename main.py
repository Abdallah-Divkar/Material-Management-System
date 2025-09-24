import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import os
import sys

# Add project root to path
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)

from modules.delivery_note import DeliveryNoteGenerator
from modules.material_list import MaterialListGenerator
from modules.dispatch_note import DispatchNoteGenerator

class HomePage(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MTS Material Management Tool")
        self.geometry("1000x700")
        self.minsize(800, 600)
        self.configure(bg="#00A695")  # header/footer green
        
        self.center_window()
        self.create_widgets()
        self.bind_keys()
    
    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        # ---------- Header ----------
        header_frame = tk.Frame(self, bg="#00A695")
        header_frame.pack(pady=(20, 30), fill="x")

        logo_title_frame = tk.Frame(header_frame, bg="#00A695")
        logo_title_frame.pack(anchor="center")

        # Logo
        try:
            logo_path = os.path.join("assets", "mts_logo.png")
            if os.path.exists(logo_path):
                logo_image = Image.open(logo_path).resize((80, 80))
                self.logo_photo = ImageTk.PhotoImage(logo_image)
                logo_label = tk.Label(logo_title_frame, image=self.logo_photo, bg="#00A695")
                logo_label.pack(side="left", padx=(0, 15))
        except Exception as e:
            print(f"Error loading logo: {e}")

        # Title
        title_label = tk.Label(
            logo_title_frame,
            text="Material Management Tool",
            font=("Arial", 28, "bold"),
            fg="white", bg="#00A695"
        )
        title_label.pack(side="left", anchor="center")

        # Subtitle
        subtitle_label = tk.Label(
            header_frame,
            text="Streamline your delivery, dispatch, and inventory documentation.",
            font=("Arial", 13, "italic"),
            fg="white", bg="#00A695"
        )
        subtitle_label.pack(pady=(8, 0))

        prompt_label = tk.Label(
            header_frame,
            text="Select a module to begin",
            font=("Arial", 17, "bold"),
            fg="white", bg="#00A695"
        )
        prompt_label.pack(pady=(15, 0))

        # ---------- Modules ----------
        modules_container = tk.Frame(self, bg="#E6F2E6")
        modules_container.pack(pady=25, expand=True)

        inner_frame = tk.Frame(modules_container, bg="#00A695")
        inner_frame.pack(expand=True)
        
        modules = [
            {"name": "Delivery Note", "description": "Generate delivery notes\nfor customer shipments", "command": self.open_delivery_note, "icon": "delivery.png"},
            {"name": "Dispatch Note", "description": "Generate dispatch notes\nfor outgoing materials", "command": self.open_dispatch_note, "icon": "dispatch.png"},
            {"name": "Material List", "description": "Create and manage\nmaterial inventory lists", "command": self.open_material_list, "icon": "material.png"}
        ]

        for i, module in enumerate(modules):
            self.create_module_card(inner_frame, module, col=i)

        # Exit button
        exit_btn = tk.Button(
            self, text="Exit Application",
            font=("Arial", 12, "bold"),
            bg="#E57373", fg="white",
            relief="flat", padx=25, pady=10,
            command=self.quit,
            cursor="hand2"
        )
        exit_btn.pack(pady=(0, 25))
        self.add_hover_effect(exit_btn, "#F08080", "#E57373")

        # Footer
        footer_frame = tk.Frame(self, bg="#00A695")
        footer_frame.pack(side="bottom", fill="x", pady=10)
        footer_label = tk.Label(
            footer_frame,
            text="Version 1.0.0   |   Developed by Abdallah Divker   |   © 2025 MTS Company",
            font=("Arial", 10),
            fg="#F5F5F5", bg="#00A695"
        )
        footer_label.pack(pady=5)

    def create_module_card(self, parent, module_config, col):
        card = tk.Frame(parent, bg="#76DEC8", padx=20, pady=20, bd=2, relief="raised")
        card.grid(row=0, column=col, padx=20, sticky="nsew")
        parent.grid_columnconfigure(col, weight=1)

        # Icon
        try:
            icon_path = os.path.join("assets", module_config["icon"])
            if os.path.exists(icon_path):
                img = Image.open(icon_path).resize((120, 120))
                icon_photo = ImageTk.PhotoImage(img)
                icon_label = tk.Label(card, image=icon_photo, bg="#76DEC8")
                icon_label.image = icon_photo
                icon_label.pack(pady=10)
        except Exception as e:
            print(f"Error loading icon: {e}")

        # Module Button
        btn = tk.Button(
            card, text=module_config["name"],
            font=("Arial", 13, "bold"),
            bg="#38778E", fg="white",
            relief="flat", width=18,
            command=module_config["command"],
            cursor="hand2"
        )
        btn.pack(pady=10)
        self.add_hover_effect(btn, "#66B5BB", "#386C8E")

        # Description
        desc_label = tk.Label(
            card, text=module_config["description"],
            font=("Arial", 10),
            fg="white", bg="#76DEC8",
            justify="center"
        )
        desc_label.pack()

    def add_hover_effect(self, widget, hover_color, normal_color):
        def on_enter(e): widget.config(bg=hover_color)
        def on_leave(e): widget.config(bg=normal_color)
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    def bind_keys(self):
        self.bind_all("<Tab>", lambda e: self.focus_next_widget(e))

    def focus_next_widget(self, event):
        event.widget.tk_focusNext().focus()
        return "break"

    def open_delivery_note(self):
        self.withdraw()
        app = DeliveryNoteGenerator(self)
        app.protocol("WM_DELETE_WINDOW", lambda: self.on_module_close(app))

    def open_material_list(self):
        self.withdraw()
        app = MaterialListGenerator(self)
        app.protocol("WM_DELETE_WINDOW", lambda: self.on_module_close(app))

    def open_dispatch_note(self):
        self.withdraw()
        app = DispatchNoteGenerator(self)
        app.protocol("WM_DELETE_WINDOW", lambda: self.on_module_close(app))

    def on_module_close(self, module_window):
        module_window.destroy()
        self.deiconify()


if __name__ == "__main__":
    app = HomePage()
    app.mainloop()
