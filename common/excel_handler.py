"""
Excel file handling functions
"""
import pandas as pd
import os
import json

# Default file name - can be overridden
DEFAULT_FILE_NAME = "Price_List.xlsx"
CACHE_FILE = "product_cache.json"

def get_products(file_path=None):
    """Return all products as a list of dictionaries"""
    try:
        # First try to load from cache if no file path provided
        if not file_path:
            cached_products = load_products_cache()
            if cached_products:
                return cached_products
        
        # Use provided file path or default
        excel_file = file_path or DEFAULT_FILE_NAME
        
        if not os.path.exists(excel_file):
            # Return sample data if file doesn't exist
            return get_sample_products()
        
        df = pd.read_excel(excel_file)
        # Clean unnamed columns
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
        products = df.to_dict(orient='records')
        
        # Save to cache if successful
        if products:
            save_products_cache(products)
            
        return products
    
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return get_sample_products()

def get_product_details(part_number, file_path=None):
    """Get product details by part number"""
    try:
        search_part = str(part_number).strip()
        products = get_products(file_path)
        
        # Find matching product
        for product in products:
            if str(product.get('Part Number', '')).strip() == search_part:
                return product
        
        return None
    
    except Exception as e:
        print(f"Error getting product details: {e}")
        return None

def get_sample_products():
    """Return sample product data for testing"""
    return [
        {
            'Part Number': 'P001',
            'Description': 'Steel Pipe 10mm',
            'Supplier': 'ABC Steel Co.',
            'Unit Price': 25.50,
            'Weight': 2.5,
            'Qty': 1
        },
        {
            'Part Number': 'P002', 
            'Description': 'Copper Wire 5m',
            'Supplier': 'ElectroTech Ltd.',
            'Unit Price': 15.75,
            'Weight': 0.8,
            'Qty': 1
        },
        {
            'Part Number': 'P003',
            'Description': 'Concrete Block',
            'Supplier': 'BuildMax Inc.',
            'Unit Price': 8.25,
            'Weight': 15.0,
            'Qty': 1
        },
        {
            'Part Number': 'P004',
            'Description': 'Paint Bucket 5L',
            'Supplier': 'ColorWorks',
            'Unit Price': 45.00,
            'Weight': 5.2,
            'Qty': 1
        },
        {
            'Part Number': 'P005',
            'Description': 'Safety Helmet',
            'Supplier': 'SafeGuard Pro',
            'Unit Price': 32.80,
            'Weight': 0.4,
            'Qty': 1
        }
    ]

def validate_excel_file(file_path):
    """Validate Excel file format"""
    try:
        df = pd.read_excel(file_path)
        
        required_columns = ['Part Number', 'Description']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            return False, f"Missing required columns: {', '.join(missing_columns)}"
        
        return True, "File is valid"
    
    except Exception as e:
        return False, f"Error reading file: {str(e)}"

def export_to_excel(data, file_path, sheet_name='Sheet1'):
    """Export data to Excel file"""
    try:
        df = pd.DataFrame(data)
        
        # Check if file exists
        if os.path.exists(file_path):
            # Append to existing file
            with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay') as writer:
                existing_df = pd.read_excel(file_path)
                combined_df = pd.concat([existing_df, df], ignore_index=True)
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            # Create new file
            df.to_excel(file_path, sheet_name=sheet_name, index=False)
        
        return True, "Export successful"
    
    except Exception as e:
        return False, f"Export failed: {str(e)}"

def save_products_cache(products):
    """Save products to cache file"""
    try:
        cache_dir = os.path.join(os.path.dirname(__file__), "cache")
        if not os.path.exists(cache_dir):
            os.makedirs(cache_dir)
            
        cache_path = os.path.join(cache_dir, CACHE_FILE)
        with open(cache_path, 'w') as f:
            json.dump(products, f)
        return True
    except Exception as e:
        print(f"Error saving cache: {e}")
        return False

def load_products_cache():
    """Load products from cache file"""
    try:
        cache_path = os.path.join(os.path.dirname(__file__), "cache", CACHE_FILE)
        if os.path.exists(cache_path):
            with open(cache_path, 'r') as f:
                return json.load(f)
    except Exception as e:
        print(f"Error loading cache: {e}")
    return None