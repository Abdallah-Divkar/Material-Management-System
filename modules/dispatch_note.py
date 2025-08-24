"""
Dispatch Note Generator Module
"""
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import sys
import os

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from common.base_generator import BaseGenerator
from common.utils import format_qty, format_price, format_weight


class DispatchNoteGenerator(BaseGenerator):
    """Dispatch Note Generator - for outgoing material dispatches"""
    
    def __init__(self, parent):
        # Set module-specific attributes before calling super().__init__
        self.module_title = "Dispatch Note Generator"
        self.export_button_text = "Export Dispatch Note"
        self.default_filename = f"dispatch_note_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        # Dispatch-specific data
        self.dispatch_date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.dispatch_to = tk.StringVar()
        self.destination = tk.StringVar()
        self.transport_method = tk.StringVar()
        self.tracking_number = tk.StringVar()
        self.priority = tk.StringVar(value="Normal")
        self.dispatched_by = tk.StringVar()
        self.special_instructions = tk.StringVar()
        
        super().__init__(parent, "Dispatch Note Generator")
    
    def create_custom_widgets(self):
        """Create dispatch note specific widgets"""
        self.create_dispatch_info_section()
    
    def create_dispatch_info_section(self):
        """Create section for dispatch-specific information"""
        info_frame = tk.LabelFrame(
            self.main_frame, 
            text="Dispatch Information", 
            bg="#F0F0F0", 
            font=("Arial", 12, "bold")
        )
        info_frame.pack(fill="x", padx=15, pady=10)
        
        # Grid layout for form fields
        # Row 0: Dispatch Date and Dispatch To
        tk.Label(info_frame, text="Dispatch Date:", bg="#F0F0F0").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.date_entry = tk.Entry(info_frame, textvariable=self.dispatch_date, width=15)
        self.date_entry.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(info_frame, text="Dispatch To:", bg="#F0F0F0").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.dispatch_to_entry = tk.Entry(info_frame, textvariable=self.dispatch_to, width=25)
        self.dispatch_to_entry.grid(row=0, column=3, padx=5, pady=5)
        
        # Row 1: Destination and Transport Method
        tk.Label(info_frame, text="Destination:", bg="#F0F0F0").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.destination_entry = tk.Entry(info_frame, textvariable=self.destination, width=15)
        self.destination_entry.grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(info_frame, text="Transport Method:", bg="#F0F0F0").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        transport_combo = ttk.Combobox(
            info_frame, 
            textvariable=self.transport_method,
            values=["Truck", "Van", "Courier", "Air Freight", "Sea Freight", "Personal Pickup"],
            width=22,
            state="readonly"
        )
        transport_combo.grid(row=1, column=3, padx=5, pady=5)
        
        # Row 2: Priority and Tracking Number
        tk.Label(info_frame, text="Priority:", bg="#F0F0F0").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        priority_combo = ttk.Combobox(
            info_frame,
            textvariable=self.priority,
            values=["Low", "Normal", "High", "Urgent"],
            width=12,
            state="readonly"
        )
        priority_combo.grid(row=2, column=1, padx=5, pady=5)
        
        tk.Label(info_frame, text="Tracking Number:", bg="#F0F0F0").grid(row=2, column=2, sticky="w", padx=5, pady=5)
        self.tracking_entry = tk.Entry(info_frame, textvariable=self.tracking_number, width=25)
        self.tracking_entry.grid(row=2, column=3, padx=5, pady=5)
        
        # Row 3: Dispatched By and Special Instructions
        tk.Label(info_frame, text="Dispatched By:", bg="#F0F0F0").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.dispatched_by_entry = tk.Entry(info_frame, textvariable=self.dispatched_by, width=15)
        self.dispatched_by_entry.grid(row=3, column=1, padx=5, pady=5)
        
        tk.Label(info_frame, text="Special Instructions:", bg="#F0F0F0").grid(row=3, column=2, sticky="w", padx=5, pady=5)
        self.instructions_entry = tk.Entry(info_frame, textvariable=self.special_instructions, width=25)
        self.instructions_entry.grid(row=3, column=3, padx=5, pady=5)
        
        # Add dispatch status tracking
        self.create_status_section(info_frame)
    
    def create_status_section(self, parent):
        """Create status tracking section"""
        status_frame = tk.LabelFrame(
            parent, 
            text="Status Tracking", 
            bg="#F0F0F0", 
            font=("Arial", 10, "bold")
        )
        status_frame.grid(row=4, column=0, columnspan=4, sticky="ew", padx=5, pady=10)
        
        # Status buttons
        self.dispatch_status = tk.StringVar(value="Prepared")
        
        status_options = ["Prepared", "In Transit", "Delivered", "Returned"]
        
        for i, status in enumerate(status_options):
            rb = tk.Radiobutton(
                status_frame,
                text=status,
                variable=self.dispatch_status,
                value=status,
                bg="#F0F0F0"
            )
            rb.grid(row=0, column=i, padx=10, pady=5)
    
    def get_treeview_columns(self):
        """Return columns specific to dispatch notes"""
        return ("Part Number", "Description", "Qty", "Unit Price", "Total Value", "Weight", "Status", "Notes")
    
    def format_item_for_tree(self, product):
        """Format product data for dispatch note tree display"""
        qty = product.get('Qty', 1)
        if isinstance(qty, str) and 'pcs' in qty:
            qty = int(qty.split()[0])
        
        unit_price = product.get('Unit Price', 0)
        total_value = qty * unit_price
        
        return (
            product['Part Number'],
            product['Description'],
            format_qty(qty),
            format_price(unit_price, self.currency_unit.get()),
            format_price(total_value, self.currency_unit.get()),
            format_weight(product.get('Weight', 0)),
            "Ready",  # Default status for dispatch
            ""  # Empty notes field
        )
    
    def get_export_data(self):
        """Return data formatted for dispatch note export"""
        data = []
        
        # Validate dispatch information
        dispatch_to = self.dispatch_to.get().strip()
        if not dispatch_to:
            raise ValueError("Dispatch To field is required")
        
        dispatch_date = self.dispatch_date.get().strip()
        if not dispatch_date:
            raise ValueError("Dispatch date is required")
        
        # Process each item in the tree
        for row in self.item_tree.get_children():
            vals = self.item_tree.item(row)['values']
            
            # Basic validation
            if not vals[0] or not vals[1]:  # Part Number and Description
                continue
            
            try:
                qty = int(vals[2].split()[0])  # Extract number from "5 pcs"
                unit_price = float(vals[3].split()[0])  # Extract number from "10.50 SAR"
                total_value = float(vals[4].split()[0])  # Extract number from "52.50 SAR"
                weight = float(vals[5].split()[0])  # Extract number from "2.500 kg"
            except (ValueError, IndexError):
                continue
            
            # Calculate totals
            total_weight = qty * weight
            
            row_data = {
                'Dispatch Date': dispatch_date,
                'Dispatch To': dispatch_to,
                'Destination': self.destination.get(),
                'Transport Method': self.transport_method.get(),
                'Priority': self.priority.get(),
                'Tracking Number': self.tracking_number.get(),
                'Dispatched By': self.dispatched_by.get(),
                'Overall Status': self.dispatch_status.get(),
                'Part Number': vals[0],
                'Description': vals[1],
                'Qty': qty,
                'Unit Price (SAR)': unit_price,
                'Total Value (SAR)': total_value,
                'Unit Weight (kg)': round(weight, 3),
                'Total Weight (kg)': round(total_weight, 3),
                'Item Status': vals[6] if len(vals) > 6 else 'Ready',
                'Item Notes': vals[7] if len(vals) > 7 else '',
                'Special Instructions': self.special_instructions.get()
            }
            data.append(row_data)
        
        return data
    
    def get_column_width(self, col):
        """Get column width for dispatch note treeview"""
        width_map = {
            "Part Number": 100,
            "Description": 180,
            "Qty": 60,
            "Unit Price": 80,
            "Total Value": 90,
            "Weight": 80,
            "Status": 70,
            "Notes": 100
        }
        return width_map.get(col, 100)
    
    def on_double_click(self, event):
        """Override to allow editing Notes and Status columns"""
        item = self.item_tree.identify_row(event.y)
        column = self.item_tree.identify_column(event.x)

        if not item:
            return
            
        column_id = int(column[1]) - 1
        column_name = self.item_tree["columns"][column_id]
        
        # Allow editing Qty, Status, and Notes columns
        if column_name not in ["Qty", "Status", "Notes"]:
            return
            
        current_value = self.item_tree.item(item)['values'][column_id]
        
        # Create entry widget for editing
        if column_name == "Status":
            # Use combobox for status
            widget = ttk.Combobox(
                self.item_tree,
                values=["Ready", "Packed", "Shipped", "Delivered", "Issue"],
                state="readonly"
            )
            widget.set(current_value)
        else:
            # Regular entry for other fields
            widget = ttk.Entry(self.item_tree)
            widget.insert(0, current_value)
        
        bbox = self.item_tree.bbox(item, column)
        widget.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        
        def on_enter(e):
            new_value = widget.get().strip()
            try:
                if column_name == "Qty":
                    qty_str = ''.join(filter(str.isdigit, new_value))
                    if not qty_str:
                        tk.messagebox.showerror("Invalid Input", "Quantity must be a positive number")
                        return
                    
                    qty_val = int(qty_str)
                    if qty_val <= 0:
                        tk.messagebox.showerror("Invalid Input", "Quantity must be greater than 0")
                        return
                    
                    new_value = format_qty(qty_val)
                
                values = list(self.item_tree.item(item)['values'])
                values[column_id] = new_value
                self.item_tree.item(item, values=values)
                
            except ValueError:
                tk.messagebox.showerror("Invalid Input", f"Please enter a valid {column_name.lower()}")
            finally:
                widget.destroy()
        
        widget.bind('<Return>', on_enter)
        widget.bind('<Escape>', lambda e: widget.destroy())
        widget.focus_set()


# Test the module independently
if __name__ == "__main__":
    # Create a simple test window
    root = tk.Tk()
    root.withdraw()  # Hide root window
    
    app = DispatchNoteGenerator(root)
    app.mainloop()