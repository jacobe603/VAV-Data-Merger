"""Utility functions for data processing and normalization."""
import pandas as pd
import json
import decimal
from datetime import datetime


class CustomJSONEncoder(json.JSONEncoder):
    """Custom JSON encoder to handle various data types."""

    def default(self, obj):
        if isinstance(obj, decimal.Decimal):
            return float(obj)
        elif isinstance(obj, datetime):
            return obj.isoformat()
        elif obj is None:
            return None
        elif isinstance(obj, (bytes, bytearray)):
            return str(obj, 'utf-8', errors='ignore')
        try:
            return super().default(obj)
        except TypeError:
            return str(obj)


def safe_string_convert(value):
    """Safely convert any value to a JSON-safe string.

    Args:
        value: Any value to convert

    Returns:
        Converted value or None if invalid
    """
    if value is None:
        return None
    elif pd.isna(value):  # Handle pandas NaN values
        return None
    elif isinstance(value, str):
        # Handle string NaN values too
        if str(value).lower() in ['nan', 'n/a', '']:
            return None
        # Remove any problematic characters
        return value.encode('ascii', 'ignore').decode('ascii')
    elif isinstance(value, (int, bool)):
        return value
    elif isinstance(value, float):
        # Handle float NaN values
        if pd.isna(value) or str(value).lower() == 'nan':
            return None
        return value
    elif isinstance(value, decimal.Decimal):
        return float(value)
    elif isinstance(value, datetime):
        return value.isoformat()
    else:
        # Convert to string and clean up
        try:
            str_val = str(value)
            if str_val.lower() in ['nan', 'n/a', '']:
                return None
            return str_val.encode('ascii', 'ignore').decode('ascii')
        except:
            return None


def normalize_tag_format(tag):
    """Convert tag formats between Excel (V-1-1) and TW2 (V-1-01) formats.

    Args:
        tag: Tag string to normalize

    Returns:
        Normalized tag with zero-padded numbers
    """
    if not tag or not isinstance(tag, str):
        return tag

    # Split the tag into parts
    parts = tag.split('-')
    if len(parts) >= 3:
        # Pad the last part with leading zero if needed
        try:
            last_num = int(parts[-1])
            if last_num < 10:
                parts[-1] = f"{last_num:02d}"
            return '-'.join(parts)
        except:
            return tag
    return tag


def normalize_unit_tag(tag):
    """Normalize unit tags by adding zero-padding: V-1-1 -> V-1-01, V-1-12 -> V-1-12.

    Args:
        tag: Unit tag string

    Returns:
        Normalized tag with zero-padded final number
    """
    if not tag:
        return tag

    tag_str = str(tag).strip()

    # Look for pattern like V-1-1, V-2-3, etc.
    import re
    pattern = r'^([A-Z]+-\d+-)(\d+)$'
    match = re.match(pattern, tag_str)

    if match:
        prefix = match.group(1)  # V-1-
        number = match.group(2)  # 1

        # Pad single digits with zero
        if len(number) == 1:
            return f"{prefix}{number.zfill(2)}"

    return tag_str


def clean_size_value(value):
    """Remove inch marks (") from size values and add zero-padding for numeric sizes.

    Args:
        value: Size value to clean

    Returns:
        Cleaned size value
    """
    if value is None or pd.isna(value):
        return value

    # Convert to string and remove all double quotes
    cleaned = str(value).replace('"', '')

    # Check if this is a simple numeric size that needs zero-padding
    # Handle both integer and string representations
    try:
        # Try to convert to integer to see if it's a simple number
        num_value = int(float(cleaned))
        # If it's a single digit (1-9), zero-pad to 2 digits
        if 1 <= num_value <= 9:
            return f"{num_value:02d}"
        # For larger numbers, return as string but ensure 2 digits minimum
        elif num_value >= 10:
            return str(num_value)
        else:
            return cleaned
    except (ValueError, TypeError):
        # Not a simple number - could be dimensions like "24x16" or "20x18"
        # Just return cleaned (without quotes)
        return cleaned


def normalize_header_text(text):
    """Clean and normalize header text.

    Args:
        text: Header text to normalize

    Returns:
        Cleaned header text
    """
    if text is None or pd.isna(text):
        return ""
    # Convert to string and clean up
    cleaned = str(text).strip()
    # Remove line breaks and normalize whitespace - this is key for "UNIT\nNO."
    cleaned = ' '.join(cleaned.split())
    # Remove special characters that cause issues
    cleaned = cleaned.replace('&', 'and').replace('"', '').replace("'", "")
    return cleaned


def is_probably_header_value(value):
    """Heuristic to decide if a cell looks like part of a header row.

    Args:
        value: Cell value to check

    Returns:
        True if value looks like a header, False otherwise
    """
    if value is None:
        return False

    if isinstance(value, (int, float)):
        return False

    text = str(value).strip()
    if not text:
        return False

    lower_text = text.lower()
    if lower_text in {'n/a', 'na', 'nan'}:
        return False

    if any(char.isdigit() for char in text):
        letters = sum(1 for c in text if c.isalpha())
        digits = sum(1 for c in text if c.isdigit())
        return letters > digits

    return True


def normalize_hw_rows_value(value):
    """Normalize HW Rows value to integer.

    Args:
        value: HW Rows value to normalize

    Returns:
        Integer value or None if invalid
    """
    try:
        if value is None or value == '':
            return None
        if isinstance(value, (int, float)):
            return int(value)
        str_val = str(value).strip()
        if not str_val:
            return None
        return int(float(str_val))
    except Exception:
        return None
