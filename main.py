import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import os
import sys

# Add the project root to Python path
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)

from modules.delivery_note import DeliveryNoteGenerator
from modules.material_list import MaterialListGenerator
from modules.dispatch_note import DispatchNoteGenerator


class HomePage(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MTS Material Management System")
        self.geometry("800x600")
        self.configure(bg="#00A651")
        
        # Center the window
        self.center_window()
        
        self.create_widgets()
    
    def center_window(self):
        """Center the window on screen"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        # Main container
        main_frame = tk.Frame(self, bg="#00A651")
        main_frame.pack(expand=True, fill="both", padx=50, pady=50)
        
        # Header section
        header_frame = tk.Frame(main_frame, bg="#00A651")
        header_frame.pack(pady=(0, 40))
        
        # Logo
        try:
            logo_path = os.path.join("assets", "mts_logo.png")
            if os.path.exists(logo_path):
                logo_image = Image.open(logo_path)
                logo_image = logo_image.resize((120, 120))
                self.logo_photo = ImageTk.PhotoImage(logo_image)
                
                logo_label = tk.Label(header_frame, image=self.logo_photo, bg="#00A651")
                logo_label.pack(pady=(0, 20))
        except Exception as e:
            print(f"Error loading logo: {e}")
        
        # Title
        title_label = tk.Label(
            header_frame,
            text="MTS Material Management System",
            font=("Arial", 24, "bold"),
            fg="white",
            bg="#00A651"
        )
        title_label.pack()
        
        subtitle_label = tk.Label(
            header_frame,
            text="Choose a module to begin",
            font=("Arial", 12),
            fg="white",
            bg="#00A651"
        )
        subtitle_label.pack(pady=(5, 0))
        
        # Modules section
        modules_frame = tk.Frame(main_frame, bg="#00A651")
        modules_frame.pack(expand=True)
        
        # Module buttons container
        buttons_frame = tk.Frame(modules_frame, bg="#00A651")
        buttons_frame.pack(expand=True)
        
        # Module configurations
        modules = [
            {
                "name": "Delivery Note Generator",
                "description": "Generate delivery notes for customer shipments",
                "command": self.open_delivery_note,
                "color": "#1E88E5"
            },
            {
                "name": "Material List Generator", 
                "description": "Create and manage material inventory lists",
                "command": self.open_material_list,
                "color": "#43A047"
            },
            {
                "name": "Dispatch Note Generator",
                "description": "Generate dispatch notes for outgoing materials",
                "command": self.open_dispatch_note,
                "color": "#FB8C00"
            }
        ]
        
        # Create module buttons
        for i, module in enumerate(modules):
            self.create_module_button(buttons_frame, module, row=i)
        
        # Footer
        footer_frame = tk.Frame(main_frame, bg="#00A651")
        footer_frame.pack(side="bottom", pady=(40, 0))
        
        exit_btn = tk.Button(
            footer_frame,
            text="Exit Application",
            font=("Arial", 10),
            bg="#DC3545",
            fg="white",
            relief="flat",
            padx=20,
            pady=5,
            command=self.quit
        )
        exit_btn.pack()
    
    def create_module_button(self, parent, module_config, row):
        """Create a styled module button"""
        # Main button frame
        btn_frame = tk.Frame(parent, bg="#00A651")
        btn_frame.pack(pady=15, fill="x")
        
        # Button
        button = tk.Button(
            btn_frame,
            text=module_config["name"],
            font=("Arial", 14, "bold"),
            bg=module_config["color"],
            fg="white",
            relief="flat",
            padx=30,
            pady=15,
            width=25,
            command=module_config["command"],
            cursor="hand2"
        )
        button.pack()
        
        # Description
        desc_label = tk.Label(
            btn_frame,
            text=module_config["description"],
            font=("Arial", 10),
            fg="white",
            bg="#00A651"
        )
        desc_label.pack(pady=(5, 0))
        
        # Hover effects
        def on_enter(e):
            button.config(bg=self.darken_color(module_config["color"]))
        
        def on_leave(e):
            button.config(bg=module_config["color"])
        
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)
    
    def darken_color(self, hex_color):
        """Darken a hex color for hover effect"""
        # Remove # if present
        hex_color = hex_color.lstrip('#')
        
        # Convert to RGB
        rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        
        # Darken by reducing each component by 20
        darkened = tuple(max(0, c - 20) for c in rgb)
        
        # Convert back to hex
        return f"#{darkened[0]:02x}{darkened[1]:02x}{darkened[2]:02x}"
    
    def open_delivery_note(self):
        """Open Delivery Note Generator"""
        self.withdraw()  # Hide homepage
        app = DeliveryNoteGenerator(self)
        app.protocol("WM_DELETE_WINDOW", lambda: self.on_module_close(app))
    
    def open_material_list(self):
        """Open Material List Generator"""
        self.withdraw()  # Hide homepage
        app = MaterialListGenerator(self)
        app.protocol("WM_DELETE_WINDOW", lambda: self.on_module_close(app))
    
    def open_dispatch_note(self):
        """Open Dispatch Note Generator"""
        self.withdraw()  # Hide homepage
        app = DispatchNoteGenerator(self)
        app.protocol("WM_DELETE_WINDOW", lambda: self.on_module_close(app))
    
    def on_module_close(self, module_window):
        """Handle module window closing"""
        module_window.destroy()
        self.deiconify()  # Show homepage again


if __name__ == "__main__":
    app = HomePage()
    app.mainloop()