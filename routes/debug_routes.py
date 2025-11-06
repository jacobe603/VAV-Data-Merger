"""Debug and testing routes."""
import os
import pandas as pd
from datetime import datetime
from flask import Blueprint, request, jsonify, session
from services.excel_service import combine_multi_row_headers, map_excel_headers_to_standard

debug_bp = Blueprint('debug', __name__)


@debug_bp.route('/debug_excel', methods=['GET'])
def debug_excel():
    """Debug endpoint to show raw Excel data."""
    if 'excel_file' not in session:
        return jsonify({'error': 'No Excel file uploaded'}), 400

    try:
        file_path = session['excel_file']
        df_raw = pd.read_excel(file_path, sheet_name=0, header=None)

        # Get first 5 rows as lists
        debug_info = {
            'file_path': file_path,
            'shape': list(df_raw.shape),
            'first_5_rows': []
        }

        for i in range(min(5, len(df_raw))):
            row_data = df_raw.iloc[i].fillna('').tolist()
            debug_info['first_5_rows'].append({
                'row_index': i,
                'data': row_data
            })

        return jsonify(debug_info)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@debug_bp.route('/debug_headers', methods=['GET'])
def debug_headers():
    """Debug endpoint to show header processing."""
    if 'excel_file' not in session:
        return jsonify({'error': 'No Excel file uploaded'}), 400

    try:
        file_path = session['excel_file']
        df_raw = pd.read_excel(file_path, sheet_name=0, header=None)

        # Show header processing step by step
        title_row_offset = 1  # Skip title row
        header_rows = 2

        # Get raw header rows
        start_row = title_row_offset
        end_row = title_row_offset + header_rows
        header_data = df_raw.iloc[start_row:end_row]

        debug_info = {
            'title_row_offset': title_row_offset,
            'header_rows': header_rows,
            'raw_header_row_1': header_data.iloc[0].fillna('').tolist(),
            'raw_header_row_2': header_data.iloc[1].fillna('').tolist(),
        }

        # Process headers
        excel_headers = combine_multi_row_headers(df_raw, header_rows=header_rows, title_row_offset=title_row_offset)
        mapped_headers = map_excel_headers_to_standard(excel_headers)

        debug_info['combined_headers'] = excel_headers
        debug_info['mapped_headers'] = mapped_headers
        debug_info['header_mapping'] = list(zip(excel_headers, mapped_headers))

        return jsonify(debug_info)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@debug_bp.route('/debug_data', methods=['GET'])
def debug_data():
    """Debug endpoint to show data extraction."""
    if 'excel_file' not in session:
        return jsonify({'error': 'No Excel file uploaded'}), 400

    try:
        file_path = session['excel_file']
        df_raw = pd.read_excel(file_path, sheet_name=0, header=None)

        # Simulate the data extraction with default settings
        data_start_row = 4  # This should be row 4 (index 3)
        data_start_index = data_start_row - 1  # Convert to 0-based = 3

        debug_info = {
            'data_start_row': data_start_row,
            'data_start_index': data_start_index,
            'total_rows': len(df_raw),
            'extracted_data_shape': None,
            'first_3_extracted_rows': []
        }

        # Extract data
        df_data = df_raw.iloc[data_start_index:].reset_index(drop=True)
        debug_info['extracted_data_shape'] = list(df_data.shape)

        # Show first 3 rows of extracted data
        for i in range(min(3, len(df_data))):
            row_data = df_data.iloc[i].fillna('').tolist()
            debug_info['first_3_extracted_rows'].append({
                'extracted_row_index': i,
                'original_row_index': data_start_index + i,
                'data': row_data
            })

        return jsonify(debug_info)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@debug_bp.route('/debug_session', methods=['GET'])
def debug_session():
    """Debug endpoint to inspect current session data."""
    try:
        # Force session creation if it doesn't exist
        if not session:
            session['debug'] = 'session_created'

        session_keys = list(session.keys())
        debug_info = {
            'session_id': request.cookies.get('session', 'No session cookie'),
            'session_keys': session_keys,
            'session_data': {},
            'flask_session_working': True,
            'storage_type': 'filesystem'
        }

        # Show session data with file info
        for key in session_keys:
            if key.endswith('_path'):
                # For file paths, check if they exist
                file_path = session[key]
                debug_info['session_data'][key] = {
                    'path': file_path,
                    'exists': os.path.exists(file_path) if file_path else False,
                    'absolute_path': os.path.abspath(file_path) if file_path else None
                }
            elif key.endswith('_data'):
                # For data, just show count
                data = session[key]
                debug_info['session_data'][key] = f"[{len(data)} records]" if isinstance(data, list) else str(type(data))
            else:
                # For other data, show as is
                debug_info['session_data'][key] = session[key]

        return jsonify(debug_info)

    except Exception as e:
        return jsonify({'error': f'Debug error: {str(e)}'}), 500


@debug_bp.route('/clear_session', methods=['POST'])
def clear_session():
    """Debug endpoint to clear session data."""
    try:
        session.clear()
        return jsonify({'success': True, 'message': 'Session cleared'})
    except Exception as e:
        return jsonify({'error': f'Clear session error: {str(e)}'}), 500


@debug_bp.route('/test_large_session', methods=['POST'])
def test_large_session():
    """Test storing large data in session (Flask-Session validation)."""
    try:
        # Create a large dataset similar to TW2 data
        large_data = []
        for i in range(100):  # 100 records
            record = {}
            for j in range(50):  # 50 columns per record
                record[f'column_{j}'] = f'value_{i}_{j}_' + 'x' * 20  # Long values
            large_data.append(record)

        # Store in session
        session['test_large_data'] = large_data
        session['test_metadata'] = {
            'records': len(large_data),
            'columns': len(large_data[0]) if large_data else 0,
            'test_timestamp': str(datetime.now())
        }

        return jsonify({
            'success': True,
            'message': 'Large data stored in session successfully',
            'data_size': len(str(large_data)),
            'records': len(large_data)
        })
    except Exception as e:
        return jsonify({'error': f'Large session test error: {str(e)}'}), 500
