"""Routes for TW2 file upload and handling."""
import os
import json
from flask import Blueprint, request, jsonify, Response, session, send_file
from werkzeug.utils import secure_filename
from utils.path_utils import sanitize_path, allowed_file
from utils.data_utils import CustomJSONEncoder
from services.tw2_service import read_tw2_data_safe
from models.config import UPLOAD_FOLDER

tw2_bp = Blueprint('tw2', __name__)


@tw2_bp.route('/upload_tw2', methods=['POST'])
def upload_tw2():
    """Upload and analyze .tw2 file with fixed encoding handling."""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file provided'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'}), 400

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)

            # Save the file
            file.save(filepath)
            abs_filepath = os.path.abspath(filepath)

            # Read the tw2 data using safe method
            result = read_tw2_data_safe(abs_filepath)

            if result['success']:
                # Store in session data
                session['tw2_file'] = abs_filepath
                session['tw2_data'] = result['data']
                session['tw2_columns'] = result['columns']
                session['original_filename'] = file.filename

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


@tw2_bp.route('/upload_updated_tw2', methods=['POST'])
def upload_updated_tw2():
    """Handle updated TW2 file upload for performance comparison."""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        if not file.filename.lower().endswith(('.tw2', '.mdb')):
            return jsonify({'error': 'Please upload a TW2 or MDB file'}), 400

        # Get optional original file path from form data
        original_path = sanitize_path(request.form.get('original_path', '').strip())
        print(f"UPLOAD: Original path provided: {original_path}")

        # Save file persistently for refresh functionality
        filename = secure_filename(file.filename)

        # Create a persistent uploads directory
        upload_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'uploads')
        os.makedirs(upload_dir, exist_ok=True)

        # Save with timestamp to avoid conflicts
        import time
        timestamp = str(int(time.time()))
        persistent_filename = f"{timestamp}_{filename}"
        persistent_path = os.path.join(upload_dir, persistent_filename)
        file.save(persistent_path)

        # Read the updated TW2 data
        result = read_tw2_data_safe(persistent_path)

        if result['success']:
            # Store updated TW2 data and file path in session
            session['updated_tw2_data'] = result['data']
            session['updated_tw2_columns'] = result['columns']
            session['updated_tw2_filename'] = filename
            session['updated_tw2_records'] = result['row_count']
            session['updated_tw2_path'] = persistent_path  # Local copy for refresh

            # Store original path if provided for remote refresh capability
            if original_path:
                session['original_tw2_path'] = original_path
                print(f"UPLOAD: Stored original path in session: {original_path}")
            else:
                # Clear original path if not provided
                session.pop('original_tw2_path', None)
                print("UPLOAD: No original path provided, cleared from session")

            return jsonify({
                'success': True,
                'filename': filename,
                'records': result['row_count'],
                'columns': result['columns'][:10],  # First 10 columns for display
                'column_count': len(result['columns']),
                'message': f'Successfully read {result["row_count"]} records with {len(result["columns"])} columns'
            })
        else:
            return jsonify({'error': f'Failed to read TW2 file: {result["error"]}'}), 400

    except Exception as e:
        print(f"Error in upload_updated_tw2: {str(e)}")
        return jsonify({'error': f'Error processing updated TW2 file: {str(e)}'}), 500


@tw2_bp.route('/get_updated_tw2_data', methods=['GET'])
def get_updated_tw2_data():
    """Get updated TW2 data for display."""
    try:
        if not session.get('updated_tw2_data'):
            return jsonify({'error': 'No updated TW2 data loaded'}), 400

        # Return the same structure as the original TW2 data viewer
        return Response(
            json.dumps({
                'success': True,
                'data': session['updated_tw2_data'],
                'columns': session.get('updated_tw2_columns', []),
                'filename': session.get('updated_tw2_filename', 'Unknown'),
                'records': session.get('updated_tw2_records', 0)
            }, cls=CustomJSONEncoder, ensure_ascii=True),
            mimetype='application/json'
        )

    except Exception as e:
        print(f"Error in get_updated_tw2_data: {str(e)}")
        return jsonify({'error': f'Error retrieving updated TW2 data: {str(e)}'}), 500


@tw2_bp.route('/download_merged_tw2', methods=['GET'])
def download_merged_tw2():
    """Download the merged TW2 file."""
    try:
        if not session.get('tw2_file'):
            return jsonify({'error': 'No TW2 file available for download'}), 400

        tw2_file_path = session['tw2_file']
        if not os.path.exists(tw2_file_path):
            return jsonify({'error': 'TW2 file not found'}), 404

        original_filename = os.path.basename(tw2_file_path)
        name_part, ext = os.path.splitext(original_filename)
        download_name = f"{name_part}_merged{ext}"

        return send_file(
            tw2_file_path,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/octet-stream'
        )
    except Exception as e:
        return jsonify({'error': f'Failed to download TW2 file: {str(e)}'}), 500


@tw2_bp.route('/validate_tw2_path', methods=['POST'])
def validate_tw2_path():
    """Validate that a TW2 file path exists and is accessible."""
    try:
        data = request.json
        file_path = sanitize_path(data.get('path', '').strip())

        if not file_path:
            return jsonify({'valid': False, 'error': 'No path provided'})

        print(f"PATH VALIDATION: Checking path: {file_path}")

        # Check if path exists
        if not os.path.exists(file_path):
            return jsonify({
                'valid': False,
                'error': f'Path not found: {file_path}',
                'details': 'File does not exist at the specified location'
            })

        # Check if it's a file (not directory)
        if not os.path.isfile(file_path):
            return jsonify({
                'valid': False,
                'error': f'Path is not a file: {file_path}',
                'details': 'The specified path points to a directory, not a file'
            })

        # Check file extension
        if not file_path.lower().endswith(('.tw2', '.mdb')):
            return jsonify({
                'valid': False,
                'error': 'Invalid file type',
                'details': 'File must be a .tw2 or .mdb file'
            })

        # Try to read the file to ensure it's accessible
        try:
            result = read_tw2_data_safe(file_path)
            if result['success']:
                return jsonify({
                    'valid': True,
                    'message': 'Path is valid and file is readable',
                    'records': result['row_count'],
                    'columns': len(result['columns'])
                })
            else:
                return jsonify({
                    'valid': False,
                    'error': 'File is not readable',
                    'details': f'Error reading TW2 file: {result["error"]}'
                })
        except Exception as e:
            return jsonify({
                'valid': False,
                'error': 'File access error',
                'details': f'Unable to access file: {str(e)}'
            })

    except Exception as e:
        print(f"Error in path validation: {str(e)}")
        return jsonify({'valid': False, 'error': f'Validation error: {str(e)}'}), 500
