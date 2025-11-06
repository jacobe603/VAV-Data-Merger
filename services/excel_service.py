"""Service for reading and processing Excel files."""
import pandas as pd
from utils.data_utils import (
    safe_string_convert,
    clean_size_value,
    normalize_header_text,
    is_probably_header_value
)


def combine_multi_row_headers(df, header_rows=2, title_row_offset=0):
    """Combine multi-row headers into single header row.

    Args:
        df: DataFrame with raw Excel data
        header_rows: Number of header rows to combine (default 2)
        title_row_offset: Number of title rows to skip before headers (default 0)

    Returns:
        List of combined header strings
    """
    headers = []

    start_row = title_row_offset
    end_row = title_row_offset + header_rows
    header_data = df.iloc[start_row:end_row]

    use_second_row = header_rows > 1 and len(header_data) > 1
    if use_second_row:
        second_row_values = [val for val in header_data.iloc[1].tolist() if not pd.isna(val)]
        if second_row_values:
            header_like_count = sum(1 for val in second_row_values if is_probably_header_value(val))
            if header_like_count < len(second_row_values) * 0.5:
                use_second_row = False

    for col_idx in range(len(df.columns)):
        row1_val = normalize_header_text(header_data.iloc[0, col_idx]) if not pd.isna(header_data.iloc[0, col_idx]) else ""
        row2_val = ""
        if use_second_row and not pd.isna(header_data.iloc[1, col_idx]):
            row2_val = normalize_header_text(header_data.iloc[1, col_idx])

        if row1_val and row2_val:
            combined = f"{row1_val}_{row2_val}"
        elif row1_val:
            combined = row1_val
        elif row2_val:
            combined = row2_val
        else:
            combined = f"Column_{col_idx + 1}"

        headers.append(combined)

    return headers


def map_excel_headers_to_standard(excel_headers):
    """Map Excel headers to our standard field names.

    Args:
        excel_headers: List of Excel header strings

    Returns:
        List of standardized header names
    """
    # Define mapping from combined Excel headers to standard names
    header_mapping = {
        # Unit identification
        'UNIT NO.': 'Unit_No',
        'UNIT_NO.': 'Unit_No',
        'UNIT NO': 'Unit_No',
        'UNIT_NO': 'Unit_No',

        # Manufacturer
        'MANUFACTURER and MODEL NO.': 'Manufacturer_Model',
        'MANUFACTURER & MODEL NO.': 'Manufacturer_Model',
        'MANUFACTURER MODEL NO.': 'Manufacturer_Model',
        'MANUFACTURER and MODEL': 'Manufacturer_Model',

        # Unit size
        'UNIT SIZE': 'Unit_Size',
        'UNIT_SIZE': 'Unit_Size',

        # Dimensions
        'W x L x H': 'Dimensions',
        'Wx Lx H': 'Dimensions',
        'DIMENSIONS': 'Dimensions',

        # Inlet size
        'INLET SIZE': 'Inlet_Size',
        'INLET_SIZE': 'Inlet_Size',

        # Outlet size
        'OUTLET SIZE': 'Outlet_Size',
        'OUTLET_SIZE': 'Outlet_Size',

        # CFM values - handle the multi-row structure
        'CFM_MAX': 'CFM_Max',
        'CFM MAX': 'CFM_Max',
        'CFM_MIN': 'CFM_Min',
        'CFM MIN': 'CFM_Min',
        'CFM_HEAT': 'CFM_Heat',
        'CFM HEAT': 'CFM_Heat',
        # Handle single CFM column that gets combined with sub-headers
        'CFM': 'CFM_Max',  # Default CFM to Max if no sub-header

        # Temperature values
        'EAT': 'EAT',
        'LAT': 'LAT',

        # Other values
        'MBH': 'MBH',
        'TOTAL MBH': 'Total_MBH',
        'TOTAL_MBH': 'Total_MBH',
        'EWT': 'EWT',
        'FLUID': 'Fluid',
        'GPM': 'GPM',
        'MAX WPD': 'Max_WPD',
        'MAX_WPD': 'Max_WPD',
        'WPD': 'WPD',
        'APD': 'APD',
        'NOTES': 'Notes',
        'TAG': 'Unit_No'
    }

    # Map headers with smart CFM handling
    mapped_headers = []

    for i, header in enumerate(excel_headers):
        header_upper = header.upper().strip()

        # Direct mapping first
        if header_upper in header_mapping:
            mapped_headers.append(header_mapping[header_upper])
        # Handle special cases for CFM sub-columns
        elif header_upper == 'MAX':
            # Check if this is likely a CFM sub-column by looking at previous headers
            if i > 0 and ('CFM' in excel_headers[i-1].upper() or any('CFM' in str(excel_headers[j]).upper() for j in range(max(0, i-3), i))):
                mapped_headers.append('CFM_Max')
            else:
                mapped_headers.append('MAX')
        elif header_upper == 'MIN':
            # Check if this is likely a CFM sub-column
            if i > 0 and ('CFM' in excel_headers[i-1].upper() or any('CFM' in str(excel_headers[j]).upper() for j in range(max(0, i-3), i))):
                mapped_headers.append('CFM_Min')
            else:
                mapped_headers.append('MIN')
        elif header_upper == 'HEAT':
            # Check if this is likely a CFM sub-column
            if i > 0 and ('CFM' in excel_headers[i-1].upper() or any('CFM' in str(excel_headers[j]).upper() for j in range(max(0, i-3), i))):
                mapped_headers.append('CFM_Heat')
            else:
                mapped_headers.append('HEAT')
        # Handle empty or generic CFM columns
        elif header_upper == 'CFM' or (header_upper == '' and i > 0 and 'CFM' in excel_headers[i-1].upper()):
            # This is likely the start of CFM multi-column, assign based on position
            mapped_headers.append('CFM_Max')  # First CFM column is usually MAX
        # Default handling - keep original or create generic name
        else:
            # Try partial matching for common patterns
            if 'MANUFACTURER' in header_upper:
                mapped_headers.append('Manufacturer_Model')
            elif 'UNIT' in header_upper and ('SIZE' in header_upper or 'NO' in header_upper):
                if 'SIZE' in header_upper:
                    mapped_headers.append('Unit_Size')
                else:
                    mapped_headers.append('Unit_No')
            elif 'INLET' in header_upper:
                mapped_headers.append('Inlet_Size')
            elif 'OUTLET' in header_upper:
                mapped_headers.append('Outlet_Size')
            elif 'DIMENSION' in header_upper or 'x' in header_upper.lower():
                mapped_headers.append('Dimensions')
            else:
                # Keep original header but clean it up
                clean_header = header.replace(' ', '_').replace('&', 'and')
                mapped_headers.append(clean_header)

    print(f"Header mapping result: {list(zip(excel_headers, mapped_headers))}")
    return mapped_headers


def read_excel_data_safe(file_path, data_start_row=3, header_rows=2, skip_title_row=True):
    """Read Excel data with proper error handling and configurable header detection.

    Args:
        file_path: Path to Excel file
        data_start_row: Row number where data starts (1-based)
        header_rows: Number of header rows to combine
        skip_title_row: Whether to skip the first row as title

    Returns:
        Dictionary with 'success', 'data', 'columns', 'row_count', 'header_info' keys
        or 'success': False with 'error' message
    """
    try:
        print("=" * 50)
        print("EXCEL FILE PROCESSING STARTED")
        print("=" * 50)
        print(f"Attempting to read Excel data from: {file_path}")

        # Read the Excel file without any header assumptions
        df_raw = pd.read_excel(file_path, sheet_name=0, header=None)
        print(f"Raw Excel shape: {df_raw.shape}")

        # Show first 5 rows for debugging
        print("=== FIRST 5 ROWS OF RAW EXCEL ===")
        for i in range(min(5, len(df_raw))):
            row_values = df_raw.iloc[i].tolist()
            print(f"Row {i}: {row_values}")
        print("=== END RAW EXCEL PREVIEW ===")

        # Use configurable header detection
        title_row_offset = 1 if skip_title_row else 0

        print(f"Configuration - Data start row: {data_start_row}, Header rows: {header_rows}, Skip title: {skip_title_row}")
        print(f"Title row offset: {title_row_offset}")

        # Auto-adjust title row offset if the first row is actual headers
        first_row_values = [val for val in df_raw.iloc[0].tolist() if not pd.isna(val)]
        if skip_title_row and first_row_values:
            header_like = sum(1 for val in first_row_values if is_probably_header_value(val))
            if header_like >= max(1, len(first_row_values) // 2):
                title_row_offset = 0
                print('AUTO-DETECT: using first row as headers')

        # Combine multi-row headers based on configuration
        excel_headers = combine_multi_row_headers(df_raw, header_rows=header_rows, title_row_offset=title_row_offset)
        print(f"Combined headers detected: {excel_headers}")

        # Map Excel headers to our standard names
        mapped_headers = map_excel_headers_to_standard(excel_headers)
        print(f"Mapped to standard headers: {mapped_headers}")

        # Extract data starting from the configured row (convert to 0-based index)
        data_start_index = data_start_row - 1
        print(f"Extracting data starting from row {data_start_row} (index {data_start_index})")
        df_data = df_raw.iloc[data_start_index:].reset_index(drop=True)
        df_data.columns = mapped_headers[:len(df_data.columns)]

        # Remove completely empty rows
        df_cleaned = df_data.dropna(how='all').reset_index(drop=True)

        # Headers should already be properly mapped, so Unit_No column should exist
        # Don't try to detect or recreate Unit_No column - trust the header mapping
        print(f"DataFrame columns after mapping: {list(df_cleaned.columns)}")
        if 'Unit_No' in df_cleaned.columns:
            print(f"Unit_No column found with sample values: {df_cleaned['Unit_No'].head().tolist()}")
        else:
            print("WARNING: Unit_No column not found in mapped headers")

        # Remove rows where critical data is missing (check multiple columns)
        # Keep rows that have data in at least Unit_Size or CFM_Max or Manufacturer_Model
        critical_cols = ['Unit_Size', 'CFM_Max', 'Manufacturer_Model']
        available_critical = [col for col in critical_cols if col in df_cleaned.columns]

        if available_critical:
            # Keep rows that have data in at least one critical column
            mask = df_cleaned[available_critical].notna().any(axis=1)
            df_cleaned = df_cleaned[mask].reset_index(drop=True)

        # Clean size fields - remove inch marks
        size_columns = ['Unit_Size', 'Inlet_Size', 'Outlet_Size']
        for col in size_columns:
            if col in df_cleaned.columns:
                df_cleaned[col] = df_cleaned[col].apply(clean_size_value)

        # Convert to safe format
        data = []
        for _, row in df_cleaned.iterrows():
            row_dict = {}
            for col in df_cleaned.columns:
                raw_value = row[col]
                safe_value = safe_string_convert(raw_value)
                row_dict[col] = safe_value
            data.append(row_dict)

        print(f"Successfully read {len(data)} Excel records")

        return {
            'success': True,
            'data': data,
            'columns': list(df_cleaned.columns),
            'row_count': len(data),
            'header_info': {
                'original_headers': excel_headers,
                'combined_headers': excel_headers,  # Same as original for now
                'mapped_headers': mapped_headers,
                'data_start_row': data_start_row
            }
        }

    except Exception as e:
        print(f"Error reading Excel data: {str(e)}")
        return {
            'success': False,
            'error': str(e).encode('ascii', 'ignore').decode('ascii')
        }
