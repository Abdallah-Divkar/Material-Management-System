"""
Utility functions for formatting and validation
"""
import json
import os
import re

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
    
def format_currency(value, currency="SAR"):
    """Format numbers with the given currency and 2 decimals"""
    try:
        value = float(value) if value not in (None, "") else 0.0
        return f"{currency} {value:,.2f}"
    except (ValueError, TypeError):
        return f"{currency} 0.00"


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

def save_to_json(data, filename="delivery_cache.json"):
    try:
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving JSON: {e}")
        return False

def load_from_json(filename="delivery_cache.json"):
    if os.path.exists(filename):
        try:
            with open(filename, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading JSON: {e}")
            return []
        
def parse_float_from_string(s):
    """
    Extract the first float number from a string.
    Returns 0.0 if no valid number found.
    """
    match = re.search(r"[-+]?\d*\.\d+|\d+", str(s))
    if match:
        return float(match.group())
    return 0.0