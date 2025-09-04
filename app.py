from flask import Flask, render_template, request, jsonify, Response
from flask_cors import CORS
import pyodbc
import pandas as pd
import os
import json
import decimal
from datetime import datetime
from werkzeug.utils import secure_filename
import shutil

app = Flask(__name__)
CORS(app)

# Configure upload settings
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'tw2', 'mdb'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Ensure upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Store uploaded data temporarily
session_data = {}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

class CustomJSONEncoder(json.JSONEncoder):
    """Custom JSON encoder to handle various data types"""
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

def normalize_tag_format(tag):
    """Convert tag formats between Excel (V-1-1) and TW2 (V-1-01) formats"""
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

def safe_string_convert(value):
    """Safely convert any value to a JSON-safe string"""
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

def get_mdb_connection(file_path):
    """Create a connection to the Access database with error handling"""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Database file not found: {file_path}")
    
    abs_path = os.path.abspath(file_path)
    
    # Try different connection strings
    connection_strings = [
        f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={abs_path};',
        f'DRIVER={{Microsoft Access Driver (*.mdb)}};DBQ={abs_path};',
        f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={abs_path};PWD=;'
    ]
    
    for conn_str in connection_strings:
        try:
            return pyodbc.connect(conn_str)
        except Exception as e:
            continue
    
    raise Exception(f"Failed to connect to database with any driver")

def read_tw2_data_safe(file_path):
    """Read TW2 data using safe methods that avoid cursor.columns()"""
    try:
        print(f"Attempting to read TW2 data from: {file_path}")
        
        conn = get_mdb_connection(file_path)
        cursor = conn.cursor()
        
        # Method 1: Get column names by running a dummy query
        print("Getting column information...")
        cursor.execute("SELECT * FROM tblSchedule WHERE 1=0")
        column_names = [desc[0] for desc in cursor.description]
        print(f"Found {len(column_names)} columns: {column_names[:10]}...")
        
        # Method 2: Get record count
        cursor.execute("SELECT COUNT(*) FROM tblSchedule")
        record_count = cursor.fetchone()[0]
        print(f"Found {record_count} records")
        
        # Method 3: Get actual data
        cursor.execute("SELECT * FROM tblSchedule")
        rows = cursor.fetchall()
        
        # Convert to safe dictionaries
        data = []
        for row in rows:
            row_dict = {}
            for i, column_name in enumerate(column_names):
                try:
                    # Safely convert each value
                    raw_value = row[i]
                    safe_value = safe_string_convert(raw_value)
                    row_dict[column_name] = safe_value
                except Exception as e:
                    print(f"Error converting column {column_name}: {e}")
                    row_dict[column_name] = None
            data.append(row_dict)
        
        conn.close()
        print("Successfully read TW2 data")
        
        return {
            'success': True,
            'data': data,
            'columns': column_names,
            'row_count': record_count
        }
        
    except Exception as e:
        print(f"Error reading TW2 data: {str(e)}")
        import traceback
        traceback.print_exc()
        
        return {
            'success': False,
            'error': str(e).encode('ascii', 'ignore').decode('ascii')
        }

def read_excel_data_safe(file_path):
    """Read Excel data with proper error handling"""
    try:
        print(f"Attempting to read Excel data from: {file_path}")
        
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name=0)
        
        # Skip the first two rows as they seem to be headers
        df_cleaned = df.iloc[2:].reset_index(drop=True)
        
        # Rename columns based on expected structure
        column_names = [
            'Unit_No', 'Manufacturer_Model', 'Unit_Size', 'Dimensions',
            'Inlet_Size', 'Outlet_Size', 'CFM_Max', 'CFM_Min', 'CFM_Heat',
            'EAT', 'LAT', 'Total_MBH', 'EWT', 'Fluid', 'GPM', 'Max_WPD', 'Notes'
        ]
        
        # Ensure we don't exceed the number of columns
        if len(column_names) > len(df_cleaned.columns):
            column_names = column_names[:len(df_cleaned.columns)]
        
        df_cleaned.columns = column_names[:len(df_cleaned.columns)]
        
        # Remove rows where Unit_No is NaN
        df_cleaned = df_cleaned[df_cleaned['Unit_No'].notna()]
        
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
            'row_count': len(data)
        }
        
    except Exception as e:
        print(f"Error reading Excel data: {str(e)}")
        return {
            'success': False,
            'error': str(e).encode('ascii', 'ignore').decode('ascii')
        }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload_tw2', methods=['POST'])
def upload_tw2():
    """Upload and analyze .tw2 file with fixed encoding handling"""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file provided'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'}), 400
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            
            # Save the file
            file.save(filepath)
            abs_filepath = os.path.abspath(filepath)
            
            # Read the tw2 data using safe method
            result = read_tw2_data_safe(abs_filepath)
            
            if result['success']:
                # Store in session data
                session_data['tw2_file'] = abs_filepath
                session_data['tw2_data'] = result['data']
                session_data['tw2_columns'] = result['columns']
                session_data['original_filename'] = file.filename
            
            # Use custom JSON encoding
            return Response(
                json.dumps(result, cls=CustomJSONEncoder, ensure_ascii=True),
                mimetype='application/json'
            )
            
    except Exception as e:
        error_msg = str(e).encode('ascii', 'ignore').decode('ascii')
        return Response(
            json.dumps({'success': False, 'error': error_msg}, ensure_ascii=True),
            mimetype='application/json',
            status=500
        )

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    """Upload and analyze Excel file"""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file provided'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'}), 400
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Read the Excel data
            result = read_excel_data_safe(filepath)
            
            if result['success']:
                # Store in session data
                session_data['excel_file'] = filepath
                session_data['excel_data'] = result['data']
                session_data['excel_columns'] = result['columns']
            
            return Response(
                json.dumps(result, cls=CustomJSONEncoder, ensure_ascii=True),
                mimetype='application/json'
            )
            
    except Exception as e:
        error_msg = str(e).encode('ascii', 'ignore').decode('ascii')
        return Response(
            json.dumps({'success': False, 'error': error_msg}, ensure_ascii=True),
            mimetype='application/json',
            status=500
        )

@app.route('/get_mapping_fields', methods=['GET'])
def get_mapping_fields():
    """Get the fields available for mapping from both files"""
    
    # Target fields in tblSchedule that need to be mapped
    target_fields = [
        'Tag', 'UnitSize', 'InletSize', 'CFMDesign', 'CFMMinPrime',
        'HeatingPrime', 'HWCFM', 'HWGPM'
    ]
    
    # Get tw2 columns if available
    tw2_fields = []
    if 'tw2_columns' in session_data:
        tw2_fields = session_data['tw2_columns']
    
    # Get Excel columns if available
    excel_fields = []
    if 'excel_columns' in session_data:
        excel_fields = session_data['excel_columns']
    
    # Suggested mappings
    suggested_mappings = {
        'Tag': 'Unit_No',
        'UnitSize': 'Unit_Size',
        'InletSize': 'Inlet_Size',
        'CFMDesign': 'CFM_Max',
        'CFMMinPrime': 'CFM_Min',
        'HeatingPrime': 'CFM_Heat',
        'HWCFM': 'CFM_Heat',
        'HWGPM': 'GPM'
    }
    
    result = {
        'target_fields': target_fields,
        'tw2_fields': tw2_fields,
        'excel_fields': excel_fields,
        'suggested_mappings': suggested_mappings
    }
    
    return Response(
        json.dumps(result, cls=CustomJSONEncoder, ensure_ascii=True),
        mimetype='application/json'
    )

@app.route('/apply_mapping', methods=['POST'])
def apply_mapping():
    """Apply the mapping and update the tw2 database"""
    try:
        data = request.json
        mappings = data.get('mappings', {})
        
        if not session_data.get('tw2_file') or not session_data.get('excel_data'):
            return Response(
                json.dumps({'success': False, 'error': 'Files not loaded'}, ensure_ascii=True),
                mimetype='application/json',
                status=400
            )
        
        # Create a backup
        backup_path = session_data['tw2_file'] + '.backup_' + datetime.now().strftime('%Y%m%d_%H%M%S')
        shutil.copy2(session_data['tw2_file'], backup_path)
        
        # Connect to the database
        conn = get_mdb_connection(session_data['tw2_file'])
        cursor = conn.cursor()
        
        updated_records = 0
        errors = []
        
        # Process each Excel row
        for excel_row in session_data['excel_data']:
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
                    ['CFMMinPrime'],                             # Batch 2a: Test CFMMinPrime alone
                    ['HWCFM'],                                   # Batch 2b: Test HWCFM alone  
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
                                update_fields.append(f"[{tw2_field}] = ?")
                                # Convert empty strings to None for database
                                if value is None or (isinstance(value, str) and str(value).strip() == ''):
                                    params.append(None)
                                else:
                                    params.append(value)
                                print(f"Debug - Batch {batch_num}: Adding field {tw2_field} = {value}")
                    
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
        
        result = {
            'success': True,
            'updated_records': updated_records,
            'backup_file': backup_path,
            'errors': errors if errors else None
        }
        
        return Response(
            json.dumps(result, cls=CustomJSONEncoder, ensure_ascii=True),
            mimetype='application/json'
        )
        
    except Exception as e:
        error_msg = str(e).encode('ascii', 'ignore').decode('ascii')
        return Response(
            json.dumps({'success': False, 'error': error_msg}, ensure_ascii=True),
            mimetype='application/json',
            status=500
        )

if __name__ == '__main__':
    app.run(debug=True, port=5004)