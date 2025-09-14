"""
Delivery Note Generator Module

This module provides functionality for generating delivery notes with proper
title section layout and consistent styling throughout the application.

Author: Material Management System
Version: 2.0
Last Modified: 2024
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import sys
import os
import json
import tkinter.filedialog as fd
from datetime import datetime
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from common.utils import replace_placeholder_in_paragraph, save_to_json, load_from_json, parse_float_from_string

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from common.base_generator import BaseGenerator
from common.utils import format_qty, format_price, format_weight

DELIVERY_INFO_FILE = "delivery_info.json"

class DeliveryNoteGenerator(BaseGenerator):
    """
    Delivery Note Generator - for incoming material deliveries
    
    This class extends BaseGenerator to provide delivery note specific functionality
    with improved title section layout and consistent UI structure.
    """
    
    def __init__(self, parent):
        """Initialize the Delivery Note Generator with proper title configuration."""
        # Set module-specific attributes before calling super().__init__
        self.module_title = "Delivery Note Generator"
        self.export_button_text = "Export Delivery Note"
        #self.default_filename = f"delivery_note_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        # Initialize parent class
        super().__init__(parent, "Delivery Note Generator")

        #self.saved_delivery_info = {}  # Initialize storage
        # Initialize delivery-specific data
        # Set system date automatically
        #customer_entry = customer_entry.get()
        self.delivery_date_var = tk.StringVar(value=datetime.now().strftime("%d-%m-%y"))
        self.notes = tk.StringVar()
        #self.default_filename = f"DN0{self.delivery_date_var.get()}-{self.customer_entry.get()}.xlsx"

        # Create the title section with delivery info callback
        self.create_title_section(left_frame_callback=self.create_delivery_info_inline)
        #self.load_delivery_info()


        # Create custom widgets
        self.create_custom_widgets()

    def create_title_section(self, left_frame_callback=None):
        """
        Create an improved title section with proper layout structure.
        
        Args:
            left_frame_callback: Optional callback for creating left frame content
        """
        # Create main header frame with consistent styling
        header_frame = tk.Frame(self.main_frame, bg="#00A651", height=120)
        header_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=(5, 0))
        
        # Configure column weights for responsive layout
        header_frame.grid_columnconfigure(0, weight=1)  # Left side (delivery info)
        header_frame.grid_columnconfigure(1, weight=0)  # Right side (logo + title)
        
        # Store header frame reference for child classes
        self.header_frame = header_frame
        
        # Left side: Delivery Information (if callback provided)
        if left_frame_callback:
            self.delivery_info_frame = left_frame_callback(header_frame)
        else:
            # Default empty frame if no callback provided
            self.delivery_info_frame = tk.Frame(header_frame, bg="#F0F0F0")
            self.delivery_info_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # Right side: Logo and Title section
        self.create_logo_title_section(header_frame)

    def create_logo_title_section(self, parent_frame):
        """
        Create the logo and title section with consistent styling.
        
        Args:
            parent_frame: Parent frame to contain logo and title elements
        """
        logo_title_frame = tk.Frame(parent_frame, bg="#00A651")
        logo_title_frame.grid(row=0, column=1, sticky="ne", padx=(10, 0), pady=10)
        
        # Logo section
        try:
            logo_path = os.path.join("assets", "mts_logo.png")
            if os.path.exists(logo_path):
                from PIL import Image, ImageTk
                logo_image = Image.open(logo_path).resize((80, 80), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_image)
                logo_label = tk.Label(
                    logo_title_frame, 
                    image=self.logo_photo, 
                    bg="#00A651",
                    relief="flat"
                )
                logo_label.pack(side="left", padx=(0, 15), pady=5)
        except Exception as e:
            print(f"Warning: Could not load logo - {e}")
            # Create placeholder for logo space
            placeholder = tk.Frame(logo_title_frame, width=80, height=80, bg="#00A651")
            placeholder.pack(side="left", padx=(0, 15), pady=5)
        
        # Title section
        title_text = getattr(self, 'module_title', 'Generator')
        title_label = tk.Label(
            logo_title_frame,
            text=title_text,
            bg="#00A651",
            fg="white",
            font=("Arial", 24, "bold"),
            anchor="e",
            justify="right"
        )
        title_label.pack(side="left", pady=5)

    def create_delivery_info_inline(self, parent_frame):
        """
        Create compact delivery information section for the header
        with updated field names and empty entries.
        
        Args:
            parent_frame: Parent frame to contain delivery info
            
        Returns:
            tk.LabelFrame: The created delivery info frame
        """
        info_frame = tk.LabelFrame(
            parent_frame,
            text="Delivery Information",
            bg="#F0F0F0",
            font=("Arial", 11, "bold"),
            labelanchor="nw",
            relief="raised",
            bd=2
        )
        info_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=5)
        
        # Configure grid weights for responsive layout
        for i in range(4):
            info_frame.grid_columnconfigure(i, weight=1)
        
        # Row 0: Customer/Company and Delivery Note Number / Ref
        self.create_info_field(info_frame, "Client:", 0, 0)
        self.customer_entry = self.create_info_entry(info_frame, "", 0, 1)
        
        #self.create_info_field(info_frame, "Delivery Note Number / Ref:", 0, 2)
        #self.delivery_ref_entry = self.create_info_entry(info_frame, "", 0, 3)
        
        # Row 1: Address
        self.create_info_field(info_frame, "Address:", 0, 2)
        self.address_entry = self.create_info_entry(info_frame, "", 0, 3)
        
        # Row 2: Phone and Fax
        self.create_info_field(info_frame, "Phone Number:", 2, 0)
        self.phone_entry = self.create_info_entry(info_frame, "", 2, 1)
        
        self.create_info_field(info_frame, "Fax:", 2, 2)
        self.fax_entry = self.create_info_entry(info_frame, "", 2, 3)
        
        # Row 3: Incharge / Contact Person
        self.create_info_field(info_frame, "Incharge:", 3, 0)
        self.incharge_entry = self.create_info_entry(info_frame, "", 3, 1)
        self.create_info_field(info_frame, "Contact  Number:", 3, 2)
        self.contact_number_entry = self.create_info_entry(info_frame, "", 3, 3)
        
        # Row 4: Customer PO Ref and Quotation
        self.create_info_field(info_frame, "Customer PO Ref:", 4, 0)
        self.po_ref_entry = self.create_info_entry(info_frame, "", 4, 1)
        
        self.create_info_field(info_frame, "Quotation:", 4, 2)
        self.quotation_entry = self.create_info_entry(info_frame, "", 4, 3)
        
        # Row 5: Project
        self.create_info_field(info_frame, "Project:", 5, 0)
        self.project_entry = self.create_info_entry(info_frame, "", 5, 1)
        self.create_info_field(info_frame, "Delivery Date:", 5, 2)
        self.delivery_date = self.create_info_entry(info_frame, self.delivery_date_var.get(), 5, 3, textvariable=self.delivery_date_var)
        
        return info_frame

    def create_info_field(self, parent, text, row, col, width=12):
        """
        Create a consistent label field for delivery information.
        
        Args:
            parent: Parent widget
            text: Label text
            row: Grid row
            col: Grid column
            width: Label width
        """
        label = tk.Label(
            parent,
            text=text,
            bg="#F0F0F0",
            font=("Arial", 9, "normal"),
            anchor="w",
            width=width
        )
        label.grid(row=row, column=col, sticky="w", padx=3, pady=2)
        return label

    def create_info_entry(self, parent, default_value="", row=0, col=0, textvariable=None, width=18):
        """
        Create a consistent entry field for delivery information.
        
        Args:
            parent: Parent widget
            default_value: Default entry value
            row: Grid row
            col: Grid column
            textvariable: Optional StringVar for the entry
            width: Entry width
            
        Returns:
            tk.Entry: The created entry widget
        """
        entry = tk.Entry(
            parent,
            font=("Arial", 9),
            width=width,
            relief="sunken",
            bd=1,
            textvariable=textvariable
        )
        if not textvariable and default_value:
            entry.insert(0, default_value)
        entry.grid(row=row, column=col, sticky="ew", padx=3, pady=2)
        return entry

    '''def save_delivery_info(self):
        """Store current delivery info entries for reuse and persist to JSON file"""
        self.saved_delivery_info = {
            "Customer": self.customer_entry.get(),
            "Delivery Ref": self.delivery_ref_entry.get(),
            "Address": self.address_entry.get(),
            "Phone": self.phone_entry.get(),
            "Fax": self.fax_entry.get(),
            "Incharge": self.incharge_entry.get(),
            "Customer PO Ref": self.po_ref_entry.get(),
            "Quotation": self.quotation_entry.get(),
            "Project": self.project_entry.get(),
            "Notes": self.notes.get()  # use StringVar
        }

        try:
            with open(DELIVERY_INFO_FILE, "w") as f:
                json.dump(self.saved_delivery_info, f, indent=4)
        except Exception as e:
            print(f"Error saving delivery info to file: {e}")

    def load_delivery_info(self):
        """Restore saved delivery info from JSON file into entries"""
        if os.path.exists(DELIVERY_INFO_FILE):
            try:
                with open(DELIVERY_INFO_FILE, "r") as f:
                    self.saved_delivery_info = json.load(f)
            except Exception as e:
                print(f"Error loading delivery info from file: {e}")
                self.saved_delivery_info = {}
        else:
            self.saved_delivery_info = {}

        if not self.saved_delivery_info:
            return

        self.customer_entry.delete(0, tk.END)
        self.customer_entry.insert(0, self.saved_delivery_info.get("Customer", ""))

        self.delivery_ref_entry.delete(0, tk.END)
        self.delivery_ref_entry.insert(0, self.saved_delivery_info.get("Delivery Ref", ""))

        self.address_entry.delete(0, tk.END)
        self.address_entry.insert(0, self.saved_delivery_info.get("Address", ""))

        self.phone_entry.delete(0, tk.END)
        self.phone_entry.insert(0, self.saved_delivery_info.get("Phone", ""))

        self.fax_entry.delete(0, tk.END)
        self.fax_entry.insert(0, self.saved_delivery_info.get("Fax", ""))

        self.incharge_entry.delete(0, tk.END)
        self.incharge_entry.insert(0, self.saved_delivery_info.get("Incharge", ""))

        self.po_ref_entry.delete(0, tk.END)
        self.po_ref_entry.insert(0, self.saved_delivery_info.get("Customer PO Ref", ""))

        self.quotation_entry.delete(0, tk.END)
        self.quotation_entry.insert(0, self.saved_delivery_info.get("Quotation", ""))

        self.project_entry.delete(0, tk.END)
        self.project_entry.insert(0, self.saved_delivery_info.get("Project", ""))

        self.notes.set(self.saved_delivery_info.get("Notes", ""))'''

    def create_custom_widgets(self):
        """Create delivery note specific widgets."""
        # This method ensures treeview columns are properly configured
        self.get_treeview_columns()
        
        # Any additional custom widgets can be added here
        # For example, additional buttons or fields specific to delivery notes

    def get_treeview_columns(self):
        """Return columns specific to delivery notes."""
        return ("Part Number", "Description", "Qty", "Supplier", "Unit Price", "Weight", "Status")
    
    def format_item_for_tree(self, product):
        """
        Format product data for delivery note tree display.
        
        Args:
            product: Product dictionary containing item details
            
        Returns:
            tuple: Formatted values for tree display
        """
        qty = format_qty(product.get('Qty', 1))
        price = format_price(product.get('Unit Price', 0), self.currency_unit.get())
        weight = format_weight(product.get('Weight', 0))
        
        return (
            product['Part Number'],
            product['Description'],
            qty,
            product.get('Supplier', ''),
            price,
            weight,
            "Pending"  # Default status for delivery
        )
    
    def get_export_data(self):
        """
        Return data formatted for delivery note export.
        """
        data = []

        # Debug: check delivery info
        customer = self.customer_entry.get().strip()
        print(f"[DEBUG] Customer entry: '{customer}'")
        if not customer:
            raise ValueError("Customer name is required")

        delivery_date = self.delivery_date.get().strip()
        print(f"[DEBUG] Delivery date entry: '{delivery_date}'")
        if not delivery_date:
            raise ValueError("Delivery date is required")

        # Debug: check treeview rows
        children = self.item_tree.get_children()
        print(f"[DEBUG] Number of rows in treeview: {len(children)}")
        if not children:
            print("[DEBUG] No items found in treeview!")
            return []

        # Process each item in the tree
        for idx, row in enumerate(children):
            vals = self.item_tree.item(row)['values']
            print(f"[DEBUG] Row {idx} values: {vals} (types: {[type(v) for v in vals]})")

            # Basic validation
            if not vals[0] or not vals[1]:
                print(f"[DEBUG] Skipping row {idx}: missing Part Number or Description")
                continue

            try:
                qty = int(str(vals[2]).split()[0])  # still "1 pcs"
                price = parse_float_from_string(vals[4])  # "USD 108.09" -> 108.09
                weight = parse_float_from_string(vals[5])  # "20.000 kg" -> 20.0
            except Exception as e:
                print(f"[DEBUG] Skipping row {idx} due to parse error: {e}")
                continue


            total_price = qty * price
            total_weight = qty * weight

            row_data = {
                'Delivery Note Date': self.delivery_date.get().strip(),
                'Customer': self.customer_entry.get().strip(),
                #'Delivery Ref': self.delivery_ref_entry.get().strip(),
                'Address': self.address_entry.get().strip(),
                'Phone': self.phone_entry.get().strip(),
                'Fax': self.fax_entry.get().strip(),
                'Incharge': self.incharge_entry.get().strip(),
                'Customer PO Ref': self.po_ref_entry.get().strip(),
                'Quotation': self.quotation_entry.get().strip(),
                'Project': self.project_entry.get().strip(),
                'Contact Number': self.contact_number_entry.get().strip(),  # <- added
                'Part Number': vals[0],
                'Description': vals[1],
                'Supplier': vals[3],
                'Qty': qty,
                'Unit Price (SAR)': price,
                'Total Price (SAR)': round(total_price, 2),
                'Unit Weight (kg)': round(weight, 3),
                'Total Weight (kg)': round(total_weight, 3),
                'Status': vals[6] if len(vals) > 6 else 'Pending',
                'Notes': self.notes.get().strip()
            }


            data.append(row_data)

        print(f"[DEBUG] Total exportable rows: {len(data)}")
        return data

    
    '''def get_export_data(self):
        """
        Return data formatted for delivery note export.
        
        Returns:
            list: List of dictionaries containing export data
            
        Raises:
            ValueError: If required fields are missing
        """
        data = []
        
        # Validate delivery information
        customer = self.customer_entry.get().strip()
        if not customer:
            raise ValueError("Customer name is required")
        
        delivery_date = self.delivery_date.get().strip()
        if not delivery_date:
            raise ValueError("Delivery date is required")
        
        # Process each item in the tree
        for row in self.item_tree.get_children():
            vals = self.item_tree.item(row)['values']
            
            # Basic validation
            if not vals[0] or not vals[1]:  # Part Number and Description
                continue
            
            try:
                qty = int(vals[2].split()[0])  # Extract number from "5 pcs"
                price = float(vals[4].split()[0])  # Extract number from "10.50 SAR"
                weight = float(vals[5].split()[0])  # Extract number from "2.500 kg"
            except (ValueError, IndexError):
                continue
            
            # Calculate totals
            total_price = qty * price
            total_weight = qty * weight
            
            row_data = {
                'Delivery Note Date': delivery_date,
                'Customer': customer,
                'Company': self.company_entry.get().strip(),
                'Address': self.address_entry.get().strip(),
                'Phone': self.phone_entry.get().strip(),
                'Fax': self.fax_entry.get().strip(),
                'Customer Number': self.customer_num_entry.get().strip(),
                'Quotation ID': self.qid_entry.get().strip(),
                'Project Name': self.project_entry.get().strip(),
                'Part Number': vals[0],
                'Description': vals[1],
                'Supplier': vals[3],
                'Qty': qty,
                'Unit Price (SAR)': price,
                'Total Price (SAR)': round(total_price, 2),
                'Unit Weight (kg)': round(weight, 3),
                'Total Weight (kg)': round(total_weight, 3),
                'Status': vals[6] if len(vals) > 6 else 'Pending',
                'Notes': self.notes.get().strip()
            }
            data.append(row_data)
        
        return data'''
    
    def export_template(self):
        """
        Export delivery note template with required columns.
        """
        try:
            export_data = self.get_export_data()

            if not export_data:
                raise ValueError("No data to export. Please add items to the delivery note.")

            # Select template file
            filepath = fd.askopenfilename(
                filetypes=[("Word Documents", "*.docx")],
                title="Select Word Template"
            )

            if not filepath:
                return

            doc = Document(filepath)

            # Mapping of placeholders to actual values
            placeholders = {
                "Customer": self.customer_entry.get().strip(),
                "Address": self.address_entry.get().strip(),
                "Phone_Num": self.phone_entry.get().strip(),
                "Fax": self.fax_entry.get().strip(),
                "Incharge": self.incharge_entry.get().strip(),
                "Customer_PO": self.po_ref_entry.get().strip(),
                "Quotation": self.quotation_entry.get().strip(),
                "Project_Name": self.project_entry.get().strip(),
                "Contact_Num": self.contact_number_entry.get().strip(),
                "Date": self.delivery_date.get().strip()
            }

            # Replace placeholders in document paragraphs, headers, footers
            for paragraph in doc.paragraphs:
                for key, value in placeholders.items():
                    replace_placeholder_in_paragraph(paragraph, f"{{{key}}}", value)
            for section in doc.sections:
                for header_paragraph in section.header.paragraphs:
                    for key, value in placeholders.items():
                        replace_placeholder_in_paragraph(header_paragraph, f"{{{key}}}", value)
                for footer_paragraph in section.footer.paragraphs:
                    for key, value in placeholders.items():
                        replace_placeholder_in_paragraph(footer_paragraph, f"{{{key}}}", value)

            # Populate item table
            self.populate_item_table(doc, export_data)

            # Save the populated document with desired default filename
            current_date = datetime.now().strftime("%d-%m-%y")
            client_name = self.customer_entry.get().strip() or "Client"
            default_filename = f"DN0{current_date}-{client_name}.docx"

            save_path = fd.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx")],
                initialfile=default_filename
            )

            if save_path:
                doc.save(save_path)
                save_to_json(export_data)
                messagebox.showinfo("Success", f"Delivery note exported successfully to:\n{save_path}")

        except Exception as e:
            messagebox.showerror("Export Failed", f"Failed to export delivery note:\n{str(e)}")

    def populate_item_table(self, doc, export_data):
        """
        Populate the item table in the Word document.
        
        Args:
            doc: Word document object
            export_data: List of item data dictionaries
        """
        if not doc.tables:
            return

        # Find the correct table (assuming first table with proper headers)
        item_table = None
        for table in doc.tables:
            if len(table.rows) > 0:
                headers = [cell.text.strip().lower() for cell in table.rows[0].cells]
                if 'no.' in headers and 'item' in headers and 'description' in headers:
                    item_table = table
                    break

        if not item_table:
            raise ValueError("Could not find the item table in the Word template.")

        # Clear existing rows (except header)
        for row in item_table.rows[1:]:
            table_element = item_table._tbl
            table_element.remove(row._tr)

        # Add new rows for each item
        for i, item in enumerate(export_data, start=1):
            row_cells = item_table.add_row().cells
            if len(row_cells) >= 4:
                row_cells[0].text = str(i)
                row_cells[1].text = str(item['Part Number'])
                row_cells[2].text = str(item['Description'])
                row_cells[3].text = str(item['Qty'])

    def get_column_width(self, col):
        """
        Get column width for delivery note treeview.
        
        Args:
            col: Column name
            
        Returns:
            int: Column width in pixels
        """
        width_map = {
            "Part Number": 100,
            "Description": 200,
            "Qty": 60,
            "Supplier": 120,
            "Unit Price": 80,
            "Weight": 80,
            "Status": 80
        }
        return width_map.get(col, 100)


# Test the module independently
if __name__ == "__main__":
    # Create a simple test window
    root = tk.Tk()
    root.withdraw()  # Hide root window
    
    try:
        app = DeliveryNoteGenerator(root)
        app.mainloop()
    except Exception as e:
        print(f"Error running delivery note generator: {e}")
        messagebox.showerror("Error", f"Failed to start application:\n{e}")
    finally:
        root.destroy()