"""Routes for Excel file upload and handling."""
import os
import json
from flask import Blueprint, request, jsonify, Response, session
from werkzeug.utils import secure_filename
from utils.path_utils import allowed_file
from utils.data_utils import CustomJSONEncoder
from services.excel_service import read_excel_data_safe
from models.config import UPLOAD_FOLDER

excel_bp = Blueprint('excel', __name__)


@excel_bp.route('/upload_excel', methods=['POST'])
def upload_excel():
    """Upload and analyze Excel file."""
    try:
        print("=" * 50)
        print("UPLOAD_EXCEL ROUTE CALLED")
        print("=" * 50)

        if 'file' not in request.files:
            print("ERROR: No file in request")
            return jsonify({'success': False, 'error': 'No file provided'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'}), 400

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)

            # Get configuration parameters from form data
            data_start_row = int(request.form.get('data_start_row', 3))
            header_rows = int(request.form.get('header_rows', 2))
            skip_title_row = request.form.get('skip_title_row', 'true').lower() == 'true'

            # Read the Excel data with configuration
            result = read_excel_data_safe(filepath, data_start_row=data_start_row,
                                        header_rows=header_rows, skip_title_row=skip_title_row)

            if result['success']:
                # Store in session data
                session['excel_file'] = filepath
                session['excel_data'] = result['data']
                session['excel_columns'] = result['columns']

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


@excel_bp.route('/get_mapping_fields', methods=['GET'])
def get_mapping_fields():
    """Get the fields available for mapping from both files."""
    from models.config import TARGET_FIELDS, SUGGESTED_FIELD_MAPPINGS

    # Get tw2 columns if available
    tw2_fields = session.get('tw2_columns', [])

    # Get Excel columns if available
    excel_fields = session.get('excel_columns', [])

    result = {
        'target_fields': TARGET_FIELDS,
        'tw2_fields': tw2_fields,
        'excel_fields': excel_fields,
        'suggested_mappings': SUGGESTED_FIELD_MAPPINGS
    }

    return Response(
        json.dumps(result, cls=CustomJSONEncoder, ensure_ascii=True),
        mimetype='application/json'
    )
