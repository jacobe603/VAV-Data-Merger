"""Service for database update operations."""
import os
import shutil
from datetime import datetime
from utils.db_utils import get_mdb_connection
from utils.data_utils import normalize_tag_format, clean_size_value


def apply_excel_mappings_to_tw2(tw2_file_path, excel_data, mappings):
    """Apply Excel field mappings to TW2 database.

    Args:
        tw2_file_path: Path to the TW2 database file
        excel_data: List of dictionaries with Excel data
        mappings: Dictionary mapping TW2 fields to Excel fields

    Returns:
        Dictionary with 'success', 'updated_records', 'backup_file', 'errors' keys
    """
    try:
        # Create a backup
        backup_path = tw2_file_path + '.backup_' + datetime.now().strftime('%Y%m%d_%H%M%S')
        shutil.copy2(tw2_file_path, backup_path)

        # Connect to the database
        conn = get_mdb_connection(tw2_file_path)
        cursor = conn.cursor()

        updated_records = 0
        errors = []

        # Process each Excel row
        for excel_row in excel_data:
            try:
                # Get the Tag value from Excel to match with tw2
                tag_field = mappings.get('Tag')
                if not tag_field:
                    continue

                tag_value = excel_row.get(tag_field)
                if not tag_value:
                    continue

                # Normalize tag format for matching (V-1-1 -> V-1-01)
                normalized_tag = normalize_tag_format(str(tag_value))
                print(f"Original tag: {tag_value} -> Normalized: {normalized_tag}")

                # Implement batched field updates to avoid SQL parameter limits
                print(f"Debug - All mappings received: {mappings}")

                # Group fields into smaller batches to isolate the problematic field
                field_batches = [
                    ['UnitSize', 'InletSize', 'CFMDesign'],      # Batch 1: Size and design fields
                    ['CFMMinPrime', 'CFMMin'],                   # Batch 2a: CFM Min fields
                    ['HWCFM', 'HeatingPrimaryAirflow'],         # Batch 2b: Heating airflow fields
                    ['HWGPM']                                    # Batch 3: GPM field
                ]

                record_updated = False
                batch_success_count = 0

                for batch_num, batch_fields in enumerate(field_batches, 1):
                    update_fields = []
                    params = []

                    # Build update for this batch
                    for tw2_field in batch_fields:
                        if tw2_field in mappings:
                            excel_field = mappings[tw2_field]
                            if excel_field in excel_row:
                                value = excel_row[excel_field]

                                # Special handling for Unit_Size mapping
                                if excel_field == 'Unit_Size' and tw2_field in ['UnitSize', 'InletSize']:
                                    # Clean and format the value first
                                    cleaned_value = clean_size_value(value)

                                    if tw2_field == 'UnitSize':
                                        # UnitSize always gets the cleaned Unit_Size value
                                        final_value = cleaned_value
                                    elif tw2_field == 'InletSize':
                                        # InletSize: special case when Unit_Size = 40, then InletSize = "24x16"
                                        # Otherwise, InletSize gets the same value as UnitSize
                                        try:
                                            unit_size_num = int(float(str(cleaned_value)))
                                            if unit_size_num == 40:
                                                final_value = "24x16"
                                            else:
                                                final_value = cleaned_value
                                        except (ValueError, TypeError):
                                            # If not a number, use cleaned value as-is
                                            final_value = cleaned_value

                                    update_fields.append(f"[{tw2_field}] = ?")
                                    if final_value is None or (isinstance(final_value, str) and str(final_value).strip() == ''):
                                        params.append(None)
                                    else:
                                        params.append(final_value)
                                    print(f"Debug - Batch {batch_num}: Unit_Size special mapping {tw2_field} = {final_value} (from {value})")

                                else:
                                    # Standard field mapping - apply size cleaning if it's a size field
                                    if tw2_field in ['UnitSize', 'InletSize'] or 'Size' in tw2_field:
                                        cleaned_value = clean_size_value(value)
                                        final_value = cleaned_value
                                    else:
                                        final_value = value

                                    update_fields.append(f"[{tw2_field}] = ?")
                                    # Convert empty strings to None for database
                                    if final_value is None or (isinstance(final_value, str) and str(final_value).strip() == ''):
                                        params.append(None)
                                    else:
                                        params.append(final_value)
                                    print(f"Debug - Batch {batch_num}: Adding field {tw2_field} = {final_value}")

                    # Execute batch if we have fields to update
                    if update_fields:
                        # Add the WHERE clause parameter - use normalized tag for matching
                        params.append(normalized_tag)

                        query = f"""
                            UPDATE tblSchedule
                            SET {', '.join(update_fields)}
                            WHERE [Tag] = ?
                        """

                        print(f"Debug - Batch {batch_num} query: {query}")
                        print(f"Debug - Batch {batch_num} params: {params}")

                        try:
                            cursor.execute(query, params)
                            if cursor.rowcount > 0:
                                batch_success_count += 1
                                record_updated = True
                                print(f"Debug - Batch {batch_num} successful for {tag_value}")
                            else:
                                print(f"Debug - Batch {batch_num} no rows affected for {tag_value}")
                        except Exception as batch_error:
                            error_msg = f"Batch {batch_num} error for {tag_value}: {str(batch_error)}"
                            errors.append(error_msg)
                            print(error_msg)

                if record_updated:
                    updated_records += 1
                    print(f"Debug - Successfully updated {tag_value} with {batch_success_count}/{len(field_batches)} batches")

            except Exception as e:
                error_msg = f"Error updating {tag_value}: {str(e)}"
                errors.append(error_msg)
                print(error_msg)

        # Commit the changes
        conn.commit()
        conn.close()

        return {
            'success': True,
            'updated_records': updated_records,
            'backup_file': backup_path,
            'errors': errors if errors else None
        }

    except Exception as e:
        return {
            'success': False,
            'error': str(e).encode('ascii', 'ignore').decode('ascii')
        }


def save_hw_rows_to_tw2(target_file, edits):
    """Save HW Rows changes to the TW2 file.

    Args:
        target_file: Path to the TW2 database file
        edits: List of dictionaries with 'unit_tag' and 'hw_rows' keys

    Returns:
        Dictionary with 'success', 'updated_count', 'backup_file', 'target_file', 'warnings' keys
    """
    try:
        # Validate edits
        for edit in edits:
            hw_rows = edit.get('hw_rows')
            if hw_rows not in [1, 2, 3, 4]:
                return {
                    'success': False,
                    'error': f'Invalid HW Rows value: {hw_rows}. Must be 1, 2, 3, or 4'
                }

        # Create backup
        backup_path = target_file + '.backup_hw_rows_' + datetime.now().strftime('%Y%m%d_%H%M%S')
        shutil.copy2(target_file, backup_path)

        # Connect to database
        conn = get_mdb_connection(target_file)
        cursor = conn.cursor()

        # Determine which HW Rows columns exist
        hw_rows_columns = []
        try:
            cursor.execute("SELECT * FROM tblSchedule WHERE 1=0")
            column_lookup = {(desc[0] or '').lower(): desc[0] for desc in (cursor.description or [])}

            hwrows_calc_column = column_lookup.get('hwrowscalc') or 'HWRowsCalc'
            hw_rows_columns.append(hwrows_calc_column)

            hwrows_column = column_lookup.get('hwrows')
            if hwrows_column and hwrows_column not in hw_rows_columns:
                hw_rows_columns.append(hwrows_column)

            hwrow_column = column_lookup.get('hwrow')
            if hwrow_column and hwrow_column not in hw_rows_columns:
                hw_rows_columns.append(hwrow_column)
        except Exception as e:
            print(f"HW ROWS: Unable to inspect columns: {e}")
            hw_rows_columns = hw_rows_columns or ['HWRowsCalc']

        updated_count = 0
        errors = []

        # Apply each edit
        for edit in edits:
            unit_tag = edit.get('unit_tag')
            hw_rows = edit.get('hw_rows')

            if unit_tag in (None, ''):
                errors.append('Missing unit tag in edit payload')
                continue

            try:
                clean_tag = unit_tag.split('  ')[0] if '  ' in unit_tag else unit_tag
                hw_rows_value = int(hw_rows)

                set_clauses = [f'[{column}] = ?' for column in hw_rows_columns]
                params = [hw_rows_value] * len(hw_rows_columns)

                update_query = f"UPDATE tblSchedule SET {', '.join(set_clauses)} WHERE [Tag] = ?"
                params.append(clean_tag)
                cursor.execute(update_query, params)

                if cursor.rowcount > 0:
                    updated_count += 1
                    print(f"Updated HW Rows for {clean_tag}: {hw_rows_value}")
                else:
                    errors.append(f"No record found for tag: {clean_tag}")
                    print(f"No record found for tag: {clean_tag}")

            except Exception as e:
                error_msg = f"Error updating {unit_tag}: {str(e)}"
                errors.append(error_msg)
                print(error_msg)

        conn.commit()
        conn.close()

        result = {
            'success': True,
            'updated_count': updated_count,
            'backup_file': os.path.basename(backup_path),
            'target_file': os.path.basename(target_file),
            'target_path': target_file
        }

        if errors:
            result['warnings'] = errors

        return result

    except Exception as e:
        return {
            'success': False,
            'error': f'Failed to save HW Rows: {str(e)}'
        }
