"""
Material List Generator Module
"""
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from common.base_generator import BaseGenerator
from common.utils import format_qty, format_price


class MaterialListGenerator(BaseGenerator):
    def __init__(self, parent):
        self.module_title = "Material List Generator"
        self.export_button_text = "Export Material List"
        self.default_filename = f"material_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        # Initialize your StringVars
        self.project_name_var = tk.StringVar()
        self.project_code_var = tk.StringVar()
        self.prepared_by_var = tk.StringVar()
        self.list_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.project_location_var = tk.StringVar()
        self.category_filter = tk.StringVar(value="All")

        # FIRST: call super to initialize BaseGenerator and self.main_frame
        super().__init__(parent, self.module_title)

        # Now, create a scrollable canvas inside the existing self.main_frame
        self.canvas = tk.Canvas(self.main_frame)
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Pack canvas and scrollbar inside self.main_frame
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Create an inner frame inside canvas that will hold your widgets
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # Bind mouse wheel scrolling for convenience
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

        # Now create your custom widgets inside the scrollable scrollable_frame
        self.create_custom_widgets()

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    def create_custom_widgets(self):
        self.create_info_section()
        self.create_filter_section()
        # You can add more sections/buttons here

    def create_info_section(self):
        info_frame = ttk.LabelFrame(self.scrollable_frame, text="Project Information", padding=(10, 10))
        info_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(info_frame, text="Project Name:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(info_frame, textvariable=self.project_name_var, width=40).grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(info_frame, text="Project Code:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(info_frame, textvariable=self.project_code_var, width=40).grid(row=1, column=1, padx=5, pady=2)

        ttk.Label(info_frame, text="List Date (YYYY-MM-DD):").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(info_frame, textvariable=self.list_date_var, width=40).grid(row=2, column=1, padx=5, pady=2)

        ttk.Label(info_frame, text="Prepared By:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(info_frame, textvariable=self.prepared_by_var).grid(row=3, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(info_frame, text="Project Location:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(info_frame, textvariable=self.project_location_var).grid(row=4, column=1, sticky="ew", padx=5, pady=5)

    def create_filter_section(self):
        filter_frame = ttk.LabelFrame(self.scrollable_frame, text="Filters", padding=(10, 10))
        filter_frame.pack(fill="x", padx=15, pady=5)

        ttk.Label(filter_frame, text="Category:").pack(side="left", padx=5)

        category_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.category_filter,
            values=["All", "Steel", "Concrete", "Electrical", "Piping", "Safety", "Tools"],
            width=15,
            state="readonly"
        )
        category_combo.pack(side="left", padx=5)
        category_combo.bind('<<ComboboxSelected>>', self.on_category_filter_change)

        ttk.Button(filter_frame, text="Generate Summary", command=self.show_summary).pack(side="right", padx=5)

    
    def get_treeview_columns(self):
        """Return columns specific to material lists"""
        return ("Part Number", "Description", "Category", "Qty", "Unit", "Unit Price", "Total Price", "Supplier")
    
    def format_item_for_tree(self, product):
        """Format product data for material list tree display"""
        qty = int(str(product.get('Qty', 1)).split()[0]) if 'pcs' in str(product.get('Qty', 1)) else product.get('Qty', 1)
        unit_price = product.get('Unit Price', 0)
        total_price = qty * unit_price
        
        return (
            product['Part Number'],
            product['Description'],
            product.get('Category', 'General'),  # Add category if available
            format_qty(qty),
            "pcs",  # Default unit
            format_price(unit_price, self.currency_unit.get()),
            format_price(total_price, self.currency_unit.get()),
            product.get('Supplier', '')
        )
    
    def on_category_filter_change(self, event=None):
        """Handle category filter change"""
        selected_category = self.category_filter.get()
        
        if selected_category == "All":
            # Show all items
            for item in self.item_tree.get_children():
                self.item_tree.set(item, "Category", self.item_tree.set(item, "Category"))
        else:
            # This is a simple implementation - in a real app you might want to
            # filter the displayed items or highlight matching categories
            pass
    
    def show_summary(self):
        """Show material list summary"""
        if not self.item_tree.get_children():
            tk.messagebox.showinfo("No Data", "No materials in the list to summarize.")
            return
        
        # Calculate totals
        total_items = len(self.item_tree.get_children())
        total_value = 0
        total_weight = 0
        categories = {}
        
        for row in self.item_tree.get_children():
            vals = self.item_tree.item(row)['values']
            try:
                # Extract total price (column 6)
                price_str = vals[6].split()[0]
                total_value += float(price_str)
                
                # Count categories
                category = vals[2]
                categories[category] = categories.get(category, 0) + 1
                
            except (ValueError, IndexError):
                continue
        
        # Create summary window
        summary_window = tk.Toplevel(self)
        summary_window.title("Material List Summary")
        summary_window.geometry("400x300")
        summary_window.configure(bg="#F0F0F0")
        
        # Summary content
        tk.Label(
            summary_window, 
            text="Material List Summary", 
            font=("Arial", 16, "bold"),
            bg="#F0F0F0"
        ).pack(pady=10)
        
        summary_text = f"""
            Project: {self.project_name_var.get()}
            Project Code: {self.project_code_var.get()}
            Date: {self.list_date_var.get()}

            Total Items: {total_items}
            Total Value: {total_value:.2f} SAR

            Categories:
            """
        
        for category, count in categories.items():
            summary_text += f"  • {category}: {count} items\n"
        
        text_widget = tk.Text(summary_window, height=15, width=50, bg="white")
        text_widget.pack(padx=20, pady=10, fill="both", expand=True)
        text_widget.insert("1.0", summary_text)
        text_widget.config(state="disabled")
        
        ttk.Button(
            summary_window, 
            text="Close", 
            command=summary_window.destroy
        ).pack(pady=10)
    
    def get_export_data(self):
        """Return data formatted for material list export"""
        data = []
        
        project_name = self.project_name_var.get().strip()
        project_code = self.project_code_var.get().strip()
        list_date = self.list_date_var.get().strip()
        prepared_by = self.prepared_by_var.get().strip()
        project_location = self.project_location_var.get().strip()
        # Validate project information
        #project_name = self.project_name.get().strip()
        if not project_name:
            raise ValueError("Project name is required")
        
        list_date = self.list_date_var.get().strip()
        if not list_date:
            raise ValueError("List date is required")
        
        # Process each item in the tree
        for row in self.item_tree.get_children():
            vals = self.item_tree.item(row)['values']
            
            # Basic validation
            if not vals[0] or not vals[1]:  # Part Number and Description
                continue
            
            try:
                qty = int(vals[3].split()[0])  # Extract number from "5 pcs"
                unit_price = float(vals[5].split()[0])  # Extract number from "10.50 SAR"
                total_price = float(vals[6].split()[0])  # Extract number from "52.50 SAR"
            except (ValueError, IndexError):
                continue
            
            row_data = {
                'Project Name': project_name,
                'Project Code': self.project_code_var.get(),
                'Project Location': project_location,
                'List Date': list_date,
                'Prepared By': self.prepared_by_var.get(),
                'Part Number': vals[0],
                'Description': vals[1],
                'Category': vals[2],
                'Qty': qty,
                'Unit': vals[4],
                'Unit Price (SAR)': unit_price,
                'Total Price (SAR)': total_price,
                'Supplier': vals[7]
            }
            data.append(row_data)
        
        return data
    
    def get_column_width(self, col):
        """Get column width for material list treeview"""
        width_map = {
            "Part Number": 100,
            "Description": 180,
            "Category": 80,
            "Qty": 60,
            "Unit": 50,
            "Unit Price": 80,
            "Total Price": 90,
            "Supplier": 120
        }
        return width_map.get(col, 100)


# Test the module independently
if __name__ == "__main__":
    # Create a simple test window
    root = tk.Tk()
    root.withdraw()  # Hide root window
    
    app = MaterialListGenerator(root)
    app.mainloop()