"""
Utility functions for formatting and validation
"""

def format_qty(qty):
    """Format quantity for display"""
    try:
        qty_val = int(qty) if qty else 1
        return f"{qty_val} pcs"
    except (ValueError, TypeError):
        return "1 pcs"

def format_price(price, currency="SAR"):
    """Format price for display"""
    try:
        price_val = float(price) if price else 0.0
        return f"{price_val:.2f} {currency}"
    except (ValueError, TypeError):
        return f"0.00 {currency}"

def format_weight(weight):
    """Format weight for display"""
    try:
        weight_val = float(weight) if weight else 0.0
        return f"{weight_val:.3f} kg"
    except (ValueError, TypeError):
        return "0.000 kg"

def validate_numeric(value, field_name, min_val=0):
    """Validate numeric input"""
    try:
        num_val = float(value)
        if num_val < min_val:
            raise ValueError(f"{field_name} must be greater than or equal to {min_val}")
        return num_val
    except (ValueError, TypeError):
        raise ValueError(f"Invalid {field_name}: must be a number")

def clean_string(value):
    """Clean string input"""
    if not value:
        return ""
    return str(value).strip()

def parse_qty_from_display(qty_display):
    """Parse quantity from display format '5 pcs' -> 5"""
    try:
        return int(qty_display.split()[0])
    except (ValueError, IndexError):
        return 1

def parse_price_from_display(price_display):
    """Parse price from display format '10.50 SAR' -> 10.50"""
    try:
        return float(price_display.split()[0])
    except (ValueError, IndexError):
        return 0.0

def parse_weight_from_display(weight_display):
    """Parse weight from display format '2.500 kg' -> 2.500"""
    try:
        return float(weight_display.split()[0])
    except (ValueError, IndexError):
        return 0.0
    
def replace_placeholder_in_paragraph(paragraph, placeholder, value):
    if placeholder in paragraph.text:
        full_text = ''.join(run.text for run in paragraph.runs)
        new_text = full_text.replace(placeholder, value)
        # Clear all runs
        for run in paragraph.runs:
            run.text = ''
        # Set new text in first run
        if paragraph.runs:
            paragraph.runs[0].text = new_text