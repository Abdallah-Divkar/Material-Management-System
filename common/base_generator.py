import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import pandas as pd
from abc import ABC, abstractmethod
#from modules.delivery_note import create_delivery_info_inline

from common.excel_handler import get_products, get_product_details
from common.utils import format_qty, format_price, format_weight

class BaseGenerator(tk.Toplevel, ABC):
    """Base class for all generator modules"""
    
    def __init__(self, parent, title="Generator"):
        super().__init__(parent)
        self.parent = parent
        self.title(title)
        self.geometry("1000x800")
        self.configure(bg="#00A651")
        
        # Common attributes
        self.products = get_products()  # This will now try to load from cache first
        self.currency_unit = tk.StringVar(value="SAR")
        self.selected_items = []
        self.combo_display_list = []
        
        # Build display list for combobox
        self.build_combo_display_list()

        self.main_frame = parent
        #self.main_frame.pack(fill="both", expand=True)

        # ✅ Create scrollable canvas + frame here
        '''self.canvas = tk.Canvas(self.main_frame)
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient="horizontal", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")'''
        
        # Create the UI
        self.create_base_widgets()
        #self.create_custom_widgets()  # Override in subclasses
        
        # Focus on this window
        self.focus_set()
        self.grab_set()
    
    def build_combo_display_list(self):
        """Build the display list for the combobox"""
        self.combo_display_list = [
            f"{p['Part Number']} - {p['Description']}" 
            for p in self.products
        ]
    
    def create_base_widgets(self):
        """Create the base UI components common to all generators"""
        
        # Create main canvas and scrollbar for the entire window
        self.main_canvas = tk.Canvas(self, bg="#00A651", highlightthickness=0)
        self.main_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)

        # Pack the main canvas and scrollbar
        self.main_canvas.pack(side="left", fill="both", expand=True, padx=15, pady=5)
        self.main_scrollbar.pack(side="right", fill="y", pady=15)

        # Create main frame inside the canvas
        self.main_frame = tk.Frame(self.main_canvas, bg="#00A651")
        self.main_canvas_frame = self.main_canvas.create_window((0, 0), window=self.main_frame, anchor="nw")

        # Bind canvas configurations
        self.main_frame.bind("<Configure>", self.on_main_frame_configure)
        self.main_canvas.bind("<Configure>", self.on_main_canvas_configure)
        
        # Title section
        self.create_title_section()
        
        # Search section
        self.create_search_section()
        
        # Details section
        self.create_details_section()
        
        # Item management section
        self.create_item_management_section()
        
        # Items treeview
        self.create_items_treeview()
        
        # Control buttons
        self.create_control_buttons()
    
    def create_title_section(self, left_frame_callback=None):
        """Header: left reserved for delivery info, right has logo + title"""
        header_frame = tk.Frame(self.main_frame, bg="#00A651")
        header_frame.grid(row=0, column=0, sticky="ew")  # or pack if needed
        self.header_frame = header_frame  # store for child to use
        header_frame.grid_columnconfigure(0, weight=1)  # left
        header_frame.grid_columnconfigure(1, weight=1)  # right
        
        # Left: Delivery Info
        if left_frame_callback:
            self.delivery_info_frame = left_frame_callback(header_frame)
        else:
            self.delivery_info_frame = tk.Frame(header_frame, bg="#F0F0F0")
            self.delivery_info_frame.grid(row=0, column=0, sticky="nsew")
        
        # Right side: Logo + Title
        logo_title_frame = tk.Frame(header_frame, bg="#00A651")
        logo_title_frame.grid(row=0, column=1, sticky="e")
        
        try:
            logo_path = os.path.join("assets", "mts_logo.png")
            if os.path.exists(logo_path):
                logo_image = Image.open(logo_path).resize((80, 80))
                self.logo_photo = ImageTk.PhotoImage(logo_image)
                logo_label = tk.Label(logo_title_frame, image=self.logo_photo, bg="#00A651")
                logo_label.pack(side="left", padx=(0, 10))
        except Exception as e:
            print(f"Error loading logo: {e}")

        title_text = getattr(self, 'module_title', 'Generator')
        title = tk.Label(logo_title_frame, text=title_text, bg="#00A651",
                        fg="white", font=("Arial", 28, "bold"))
        title.pack(side="left")


    def create_search_section(self):
        """Create the search/product selection section"""
        search_frame = tk.Frame(self.main_frame, bg="#00A651")
        search_frame.grid(row=1, column=0, pady=10, sticky="w")

        
        tk.Label(
            search_frame, 
            text="Search Products (Part Number, Description):",
            bg="#00A651", 
            fg="white", 
            font=("Arial", 12)
        ).pack(side="left", padx=(0, 10))

        #self.search_var = tk.StringVar()
        self.combo_var = tk.StringVar()
        self.combo = ttk.Combobox(
            search_frame, 
            values=self.combo_display_list, 
            textvariable=self.combo_var, 
            width=50
        )
        self.combo.pack(side="left", pady=5)
        self.combo.bind('<KeyRelease>', self.on_keyrelease)
        self.combo.bind('<<ComboboxSelected>>', self.on_item_selected)
        self.combo.bind('<Return>', self.on_enter_pressed)

        ttk.Button(search_frame, text="Upload File", command=self.upload_file).pack(padx=10)
    
    def create_details_section(self):
        """Create the product details section"""
        # Create container frame for both details and treeview
        self.container_frame = tk.Frame(self.main_frame, bg="#F0F0F0")
        self.container_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)

        # Left side - Details section
        details_container = tk.Frame(self.container_frame, bg="#F0F0F0")
        details_container.pack(side='left', fill='both', padx=(0, 10))

        # Details frame with fixed width
        outer_frame = tk.Frame(details_container, bg="#F0F0F0", width=400, height=500)
        outer_frame.pack(fill='both', expand=True)
        outer_frame.propagate(False)

        # Add canvas and scrollbar for details
        self.canvas = tk.Canvas(outer_frame, bg="#F0F0F0", width=400)
        scrollbar = ttk.Scrollbar(outer_frame, orient="vertical", command=self.canvas.yview)

        # Detail frame (managed by canvas only)
        self.detail_frame = tk.Frame(self.canvas, bg="#F0F0F0", relief="sunken", borderwidth=2)

        # Configure canvas
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Create window in canvas for detail_frame
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.detail_frame, anchor="nw")

        # Bind canvas configurations
        self.detail_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)

        # Add middle buttons section
        self.create_middle_buttons()

        self.detail_labels = {}

    def create_middle_buttons(self):
        """Create buttons column between details and treeview"""
        buttons_frame = tk.Frame(self.container_frame, bg="#F0F0F0", width=50)
        buttons_frame.pack(side='left', fill='y', padx=10)
        
        # Add spacing at top
        tk.Frame(buttons_frame, height=20, bg="#F0F0F0").pack()
        
        # Add to Tree button
        self.add_to_tree_btn = ttk.Button(
            buttons_frame,
            text="+",
            width=3,
            command=self.add_selected_items
        )
        self.add_to_tree_btn.pack(pady=5)
        
        # Remove from Tree button
        self.remove_from_tree_btn = ttk.Button(
            buttons_frame,
            text="-",
            width=3,
            command=self.remove_selected_item
        )
        self.remove_from_tree_btn.pack(pady=5)
        
        # Move Up button
        self.move_up_btn = ttk.Button(
            buttons_frame,
            text="↑",
            width=3,
            command=self.move_item_up
        )
        self.move_up_btn.pack(pady=5)
        
        # Move Down button
        self.move_down_btn = ttk.Button(
            buttons_frame,
            text="↓",
            width=3,
            command=self.move_item_down
        )
        self.move_down_btn.pack(pady=5)

        self.clear_details_btn = ttk.Button(
            buttons_frame, 
            text="Clear All",  # Changed text to be more clear
            width=8,  # Set a fixed width
            command=self.clear_details
        )
        self.clear_details_btn.pack(pady=5)
        
        # Add spacing at bottom
        tk.Frame(buttons_frame, height=20, bg="#F0F0F0").pack()

    # Add these new methods for moving items
    def move_item_up(self):
        """Move selected item up in the treeview"""
        selected = self.item_tree.selection()
        if not selected:
            return
            
        for item in selected:
            idx = self.item_tree.index(item)
            if idx > 0:
                prev = self.item_tree.prev(item)
                self.item_tree.move(item, self.item_tree.parent(item), self.item_tree.index(prev))

    def move_item_down(self):
        """Move selected item down in the treeview"""
        selected = self.item_tree.selection()
        if not selected:
            return
            
        for item in reversed(selected):
            next_item = self.item_tree.next(item)
            if next_item:
                self.item_tree.move(item, self.item_tree.parent(item), self.item_tree.index(next_item))

    def create_items_treeview(self):
        """Create the items treeview - to be customized by subclasses"""
        columns = self.get_treeview_columns()
        
        # Create a frame to hold the Treeview and Scrollbar
        tree_frame = ttk.Frame(self.container_frame)  # Use container_frame instead of main_frame
        tree_frame.pack(side='left', fill="both", expand=True)

        # Create vertical scrollbar
        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        tree_scrollbar.pack(side="right", fill="y")

        # Create Treeview and attach scrollbar
        self.item_tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show='headings',
            yscrollcommand=tree_scrollbar.set
        )
        self.item_tree.pack(side="left", fill="both", expand=True)

        # Configure scrollbar
        tree_scrollbar.config(command=self.item_tree.yview)

        # Set up headings and column widths
        for col in columns:
            self.item_tree.heading(col, text=col)
            width = self.get_column_width(col)
            self.item_tree.column(col, width=width)

        # Add double-click binding for editing
        self.item_tree.bind('<Double-1>', self.on_double_click)

    def create_control_buttons(self):
        """Create the control buttons section"""
        btn_frame = tk.Frame(self.main_frame, bg="#00A651")
        btn_frame.grid(row=4, column=0, pady=5, sticky="ew")


        ttk.Button(btn_frame, text="Reset", command=self.reset_items).grid(row=0, column=1, padx=10)
        
        # Export button text customizable by subclass
        export_text = getattr(self, 'export_button_text', 'Export to Excel')
        ttk.Button(btn_frame, text=export_text, command=self.export_to_excel).grid(row=0, column=2, padx=10)
        
        if hasattr(self, 'export_template'):
            ttk.Button(btn_frame, text="Export as Template", command=self.export_template).grid(row=0, column=3, padx=10)

        ttk.Button(btn_frame, text="Back to Home", command=self.return_home).grid(row=0, column=4, padx=10)

        self.btn_frame = btn_frame
    
    # Abstract methods to be implemented by subclasses
    @abstractmethod
    def create_custom_widgets(self):
        """Create custom widgets specific to the generator type"""
        pass
    
    @abstractmethod
    def get_treeview_columns(self):
        """Return the columns for the treeview"""
        pass
    
    @abstractmethod
    def get_export_data(self):
        """Return data formatted for export"""
        pass
    
    # Common functionality methods
    def get_column_width(self, col):
        """Get column width for treeview"""
        width_map = {
            "Part Number": 100,
            "Description": 200,
            "Qty": 80,
            "Supplier": 100,
            "Unit Price": 100,
            "Weight": 80,
            "Customer": 120,
            "Delivery Date": 100,
            "Location": 120,
            "Notes": 150
        }
        return width_map.get(col, 100)
    
    def on_keyrelease(self, event):
        """Handle key release - maintains original behavior of showing details while typing"""
        typed = self.combo_var.get().lower()
        
        # Check if add_btn exists before trying to configure it  
        if hasattr(self, 'add_btn'):
            self.add_btn.config(state='disabled')
        
        if not typed:
            self.combo['values'] = self.combo_display_list
            self.clear_details()  # Only clear when search is empty
            return

        # Filter products based on user input
        filtered_products = [
            p for p in self.products
            if (typed in str(p['Part Number']).lower() or
                typed in str(p['Description']).lower())
        ]
        
        filtered_display = [
            f"{p['Part Number']} - {p['Description']}" for p in filtered_products
        ]
        self.combo['values'] = filtered_display

        # Clear existing details and show all matches
        self.clear_details()
        
        # Show all matches in the details frame (just like before)
        if filtered_products:
            for product in filtered_products:
                self.show_details(product)

    def on_item_selected(self, event):
        """Handle item selection from combobox"""
        selection = self.combo_var.get()
        if not selection:
            return

        # Extract part number from selection
        part_number = str(selection.split(' - ')[0].strip())
        print(f"Selected raw combobox text: '{selection}'")
        print(f"Extracted part number: '{part_number}'")
        product = next((p for p in self.products if str(p.get('Part Number', '')).strip() == part_number), None)
        
        if product:
            self.show_details(product)
            # Check if add_btn exists before trying to configure it
            if hasattr(self, 'add_btn'):
                self.add_btn.config(state='normal')
        else:
            messagebox.showerror("Error", f"Product with part number '{part_number}' not found.")
            # Check if add_btn exists before trying to configure it
            if hasattr(self, 'add_btn'):
                self.add_btn.config(state='disabled')

    def add_item(self, frame_to_remove=None):
        """Add currently selected product to the items tree"""
        if not hasattr(self, 'current_product') or self.current_product is None:
            messagebox.showwarning("Warning", "No valid product selected.")
            return

        p = self.current_product
        try:
            # Get formatted values
            values = self.format_item_for_tree(p)
            
            # Insert into tree
            self.item_tree.insert("", "end", values=values)
            self.selected_items.append(p)
            self.update_remove_button_state()

            # Clear selection
            self.combo.set('')
            
            if frame_to_remove:
                frame_to_remove.destroy()
            
            # Check if add_btn exists before trying to configure it
            if hasattr(self, 'add_btn'):
                self.add_btn.config(state='disabled')

        except ValueError as e:
            messagebox.showerror("Error", f"Invalid data format: {e}")
    

    '''def on_item_selected(self, event):
        self.on_enter_pressed(event)'''

    '''def on_item_selected(self, event):
        """When user selects an item from dropdown, show its details."""
        selection = self.combo.get()
        if not selection:
            return

        # Extract part number from selection and convert to string
        part_number = str(selection.split(' - ')[0].strip())
        product = get_product_details(part_number)
        
        if product:
            self.show_details(product)
            self.add_btn.config(state='normal')
        else:
            messagebox.showerror("Error", f"Product with part number '{part_number}' not found.")
            self.add_btn.config(state='disabled')'''

    def clear_details(self):
        """Clear all widgets in detail frame"""
        if self.detail_frame.winfo_children():  # Only show dialog if there are items to clear
            if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all details?"):
                for child in self.detail_frame.winfo_children():
                    child.destroy()
                
                self.detail_labels = {}
                self.current_product = None
                self.combo.set('')  # Clear the combobox selection
                
                # Update button states
                if hasattr(self, 'add_btn'):
                    self.add_btn.config(state='disabled')
                if hasattr(self, 'add_to_tree_btn'):
                    self.add_to_tree_btn.config(state='disabled')
        else:
            messagebox.showinfo("Info", "No details to clear")

    def show_details(self, product):
        """Show product details in horizontal layout with checkbox"""
        frame = ttk.Frame(self.detail_frame, borderwidth=1, relief="solid", padding=5)
        frame.pack(fill='x', pady=2, padx=5)

        # Store the product reference
        frame.product = product

        # Create checkbox variable for this product
        checkbox_var = tk.BooleanVar()
        frame.checkbox_var = checkbox_var

        # Checkbox
        checkbox = ttk.Checkbutton(frame, variable=checkbox_var)
        checkbox.pack(side='left', padx=(0, 5))

        # Part Number (bold, fixed width)
        part_label = ttk.Label(
            frame, 
            text=f"{product['Part Number']}", 
            font=('Arial', 9, 'bold'),
            width=20
        )
        part_label.pack(side='left', padx=(0, 10))

        # Description (truncated if too long)
        description = str(product['Description'])
        if len(description) > 50:
            description = description[:47] + "..."
        
        desc_label = ttk.Label(
            frame, 
            text=description, 
            font=('Arial', 9),
            width=45
        )
        desc_label.pack(side='left', padx=(0, 10))

        # Clear button
        clear_btn = ttk.Button(
            frame,
            text="✕",
            width=3,
            command=lambda f=frame: f.destroy()
        )
        clear_btn.pack(side='right', padx=(5, 0))

    def add_selected_items(self):
        """Add all checked items to the tree"""
        added_count = 0
        frames_to_remove = []
        
        # Check all frames in detail_frame for selected checkboxes
        for child in self.detail_frame.winfo_children():
            if hasattr(child, 'checkbox_var') and child.checkbox_var.get():
                # This frame has a checked checkbox
                product = child.product
                try:
                    # Get formatted values
                    values = self.format_item_for_tree(product)
                    
                    # Insert into tree
                    self.item_tree.insert("", "end", values=values)
                    self.selected_items.append(product)
                    added_count += 1
                    frames_to_remove.append(child)
                    
                except ValueError as e:
                    messagebox.showerror("Error", f"Invalid data format for {product.get('Part Number', 'Unknown')}: {e}")
        
        # Remove the frames of added items
        for frame in frames_to_remove:
            frame.destroy()
        
        if added_count > 0:
            self.update_remove_button_state()
            self.combo.set('')  # Clear search
            messagebox.showinfo("Items Added", f"Added {added_count} item(s) to the list.")
        else:
            messagebox.showwarning("No Selection", "Please select items using checkboxes first.")

    def create_item_management_section(self):
        """Create the item management buttons section with checkbox support"""
        append_item_frame = tk.Frame(self.main_frame, bg='#00A651')
        append_item_frame.grid(row=3, column=0, pady=10, sticky="w")


        # Add Selected Items button (replaces individual Add Item button)
        self.add_selected_btn = ttk.Button(
            append_item_frame, 
            text="+", 
            command=self.add_selected_items
        )
        self.add_selected_btn.grid(row=0, column=1, padx=10)

        # Remove Selected button
        self.remove_btn = ttk.Button(
            append_item_frame, 
            text="-", 
            command=self.remove_selected_item
        )
        self.remove_btn.grid(row=0, column=2, padx=10)
        self.remove_btn.config(state='disabled')

        # Clear All Details button
        clear_details_btn = ttk.Button(
            append_item_frame, 
            text="Clear Details", 
            command=self.clear_details
        )
        clear_details_btn.grid(row=0, column=3, padx=10)

    
    '''def select_for_list(self, product, frame):
        """Add selected product to list"""
        self.current_product = product
        self.add_item(frame)
    
    def add_item(self, frame_to_remove=None):
        """Add currently selected product to the items tree"""
        if not hasattr(self, 'current_product') or self.current_product is None:
            messagebox.showwarning("Warning", "No valid product selected.")
            return

        p = self.current_product
        try:
            # Get formatted values
            values = self.format_item_for_tree(p)
            
            # Insert into tree
            self.item_tree.insert("", "end", values=values)
            self.selected_items.append(p)
            self.update_remove_button_state()

            # Clear selection
            self.combo.set('')
            
            if frame_to_remove:
                frame_to_remove.destroy()
            
            self.add_item.config(state='disabled')

        except ValueError as e:
            messagebox.showerror("Error", f"Invalid data format: {e}")
'''
    def format_item_for_tree(self, product):
        """Format product data for tree display - override in subclasses"""
        qty = format_qty(product.get('Qty', 1))
        price = format_price(product.get('Unit Price', 0), self.currency_unit.get())
        weight = format_weight(product.get('Weight', 0))
        
        return (
            product['Part Number'],
            product['Description'],
            qty,
            product.get('Supplier', ''),
            price,
            weight
        )
    
    def remove_selected_item(self):
        """Remove selected items from the tree"""
        all_items = self.item_tree.get_children()
        num_items = len(all_items)

        if num_items == 0:
            messagebox.showwarning("No Items", "There are no items to remove.")
            self.remove_btn.config(state='disabled')
            return

        if num_items == 1:
            # Only one item - remove it immediately
            self.item_tree.delete(all_items[0])
            if self.selected_items:
                self.selected_items.pop(0)
            self.update_remove_button_state()
            return

        # Multiple items - require selection
        selected = self.item_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select item(s) to remove.")
            return

        # Remove selected items
        indices_to_remove = sorted([self.item_tree.index(item) for item in selected], reverse=True)

        for item_id in selected:
            self.item_tree.delete(item_id)

        # Remove from selected_items list
        for index in indices_to_remove:
            if index < len(self.selected_items):
                self.selected_items.pop(index)

        self.update_remove_button_state()
    
    def reset_items(self):
        """Reset all items"""
        self.selected_items.clear()
        for row in self.item_tree.get_children():
            self.item_tree.delete(row)
        self.update_remove_button_state()
    
    def update_remove_button_state(self):
        """Update remove button state based on items"""
        if self.item_tree.get_children():
            self.remove_btn.config(state='normal')
        else:
            self.remove_btn.config(state='disabled')
    
    def export_to_excel(self):
        """Export items to Excel - uses get_export_data() from subclass"""
        if not self.item_tree.get_children():
            messagebox.showinfo("No Data", "No items to export.")
            return

        # Get export data from subclass
        try:
            data = self.get_export_data()
            
            if not data:
                messagebox.showinfo("No Valid Data", "No valid rows to export.")
                return

            # File selection
            default_name = getattr(self, 'default_filename', 'export.xlsx')
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                title=f"Save {getattr(self, 'module_title', 'Export')} As",
                initialfile=default_name
            )

            
            print(f"User selected file path: {file_path}")

            if not file_path:
                print("Export cancelled by user (no file selected).")
                return
            # Check if file exists and ask for overwrite confirmation
            if os.path.exists(file_path):
                confirm = messagebox.askyesno(
                    "Confirm Overwrite",
                    f"The file '{os.path.basename(file_path)}' already exists.\nDo you want to overwrite it?"
                )
                if not confirm:
                    print("User declined to overwrite existing file.")
                    return

            # Proceed with export
            df_new = pd.DataFrame(data)
            if os.path.exists(file_path):
                # Load existing and combine
                existing_df = pd.read_excel(file_path)
                df_combined = pd.concat([existing_df, df_new], ignore_index=True)
                df_combined.to_excel(file_path, index=False)
            else:
                df_new.to_excel(file_path, index=False)

            messagebox.showinfo("Export Successful", f"File saved to:\n{file_path}")
            print(f"Export successful to {file_path}")

        except Exception as e:
            messagebox.showerror("Export Failed", f"Could not save file:\n{e}")
            print(f"Export failed with error: {e}")

    def upload_file(self):
        """Upload product list file and cache it"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                self.products = df.to_dict(orient='records')
                
                # Save to cache
                from common.excel_handler import save_products_cache
                if save_products_cache(self.products):
                    print("Products cached successfully")
                
                # Update display list
                self.build_combo_display_list()
                self.combo['values'] = self.combo_display_list
                messagebox.showinfo("Success", "Product list uploaded and cached successfully.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read Excel file:\n{str(e)}")
    
    def on_double_click(self, event):
        """Handle double-click on treeview item to edit"""
        item = self.item_tree.identify_row(event.y)
        column = self.item_tree.identify_column(event.x)

        if not item:
            return
            
        column_id = int(column[1]) - 1
        column_name = self.item_tree["columns"][column_id]
        
        # Only allow editing Qty and unit price columns by default
        if column_name not in ["Qty", "Unit Price"]:
            return
            
        current_value = self.item_tree.item(item)['values'][column_id]
        
        # Create entry widget for editing
        entry = ttk.Entry(self.item_tree)
        entry.insert(0, current_value)
        
        bbox = self.item_tree.bbox(item, column)
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        
        def on_enter(e):
            new_value = entry.get().strip()
            try:
                if column_name == "Qty":
                    qty_str = ''.join(filter(str.isdigit, new_value))
                    if not qty_str:
                        messagebox.showerror("Invalid Input", "Quantity must be a positive number")
                        return
                    
                    qty_val = int(qty_str)
                    if qty_val <= 0:
                        messagebox.showerror("Invalid Input", "Quantity must be greater than 0")
                        return
                    
                    new_value = format_qty(qty_val)
                
                values = list(self.item_tree.item(item)['values'])
                values[column_id] = new_value
                self.item_tree.item(item, values=values)
                
            except ValueError:
                messagebox.showerror("Invalid Input", f"Please enter a valid {column_name.lower()}")
            finally:
                entry.destroy()
        
        entry.bind('<Return>', on_enter)
        entry.bind('<Escape>', lambda e: entry.destroy())
        entry.focus_set()
    
    def on_enter_pressed(self, event):
        """Handle Enter key in search combobox - now opens dropdown"""
        # Open the dropdown when Enter is pressed
        self.combo.event_generate('<Button-1>')
    
    def return_home(self):
        """Return to homepage"""
        self.destroy()
        self.parent.deiconify()
    
    # Canvas scroll methods
    def on_frame_configure(self, event=None):
        """Reset the scroll region to encompass the inner frame"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_canvas_configure(self, event):
        """When canvas is resized, resize the inner frame to match"""
        width = event.width
        self.canvas.itemconfig(self.canvas_frame, width=width)

    def on_main_frame_configure(self, event=None):
        """Reset the scroll region to encompass the main frame"""
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

    def on_main_canvas_configure(self, event):
        """When main canvas is resized, resize the main frame to match"""
        width = event.width
        self.main_canvas.itemconfig(self.main_canvas_frame, width=width)