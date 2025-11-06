"""Service for reading and processing TW2 database files."""
import os
from utils.db_utils import get_mdb_connection
from utils.data_utils import safe_string_convert


def read_tw2_data_safe(file_path):
    """Read TW2 data using safe methods that avoid cursor.columns().

    Args:
        file_path: Path to the TW2/MDB database file

    Returns:
        Dictionary with 'success', 'data', 'columns', 'row_count' keys
        or 'success': False with 'error' message
    """
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


def reload_tw2_data_from_disk(session, preferred_paths=None):
    """Reload TW2 data from disk, updating the session with the latest contents.

    Args:
        session: Flask session object
        preferred_paths: Optional list of (label, path) tuples to try first

    Returns:
        Dictionary with success status and reload information
    """
    from utils.path_utils import sanitize_path

    candidates = []
    seen_paths = set()

    def _add_candidate(label, candidate_path):
        if not candidate_path:
            return
        sanitized = sanitize_path(str(candidate_path).strip())
        if not sanitized:
            return
        normalized = os.path.abspath(sanitized)
        if normalized in seen_paths:
            return
        candidates.append((label, sanitized))
        seen_paths.add(normalized)

    if preferred_paths:
        for label, candidate_path in preferred_paths:
            _add_candidate(label, candidate_path)

    _add_candidate('original', session.get('original_tw2_path'))
    _add_candidate('local', session.get('updated_tw2_path'))
    _add_candidate('tw2', session.get('tw2_file'))

    if not candidates:
        return {'success': False, 'error': 'No TW2 file path available', 'code': 404}

    last_error = {'message': 'Unable to reload TW2 data', 'code': 500}

    for label, candidate_path in candidates:
        if not os.path.exists(candidate_path):
            last_error = {'message': f'File not found at {candidate_path}', 'code': 404}
            continue

        result = read_tw2_data_safe(candidate_path)
        if result.get('success'):
            session['updated_tw2_data'] = result['data']
            session['updated_tw2_columns'] = result['columns']
            session['updated_tw2_records'] = result['row_count']
            session['updated_tw2_filename'] = os.path.basename(candidate_path)
            session['last_tw2_reload_path'] = candidate_path
            session['last_tw2_reload_source'] = label
            session['tw2_last_path'] = candidate_path
            try:
                session['tw2_last_mtime'] = os.path.getmtime(candidate_path)
            except Exception:
                session.pop('tw2_last_mtime', None)
            return {
                'success': True,
                'path': candidate_path,
                'source': label,
                'row_count': result['row_count'],
                'column_count': len(result['columns'])
            }
        else:
            last_error = {
                'message': result.get('error', 'Unknown error while reading TW2 file'),
                'code': 500
            }

    return {'success': False, 'error': last_error['message'], 'code': last_error['code']}
