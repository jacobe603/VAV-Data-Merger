"""Routes for performance comparison and data mapping."""
import json
import logging
from datetime import datetime
from flask import Blueprint, request, jsonify, Response, session, send_file
from utils.data_utils import CustomJSONEncoder
from utils.path_utils import sanitize_path
from utils.db_utils import get_project_name_from_tw2
from services.comparison_service import compare_performance_data
from services.tw2_service import reload_tw2_data_from_disk
from services.database_service import apply_excel_mappings_to_tw2, save_hw_rows_to_tw2
from services.report_service import generate_schedule_data_excel
from models.config import (
    DEFAULT_MBH_LAT_LOWER_MARGIN,
    DEFAULT_MBH_LAT_UPPER_MARGIN,
    DEFAULT_WPD_THRESHOLD,
    DEFAULT_APD_THRESHOLD
)

logger = logging.getLogger('vav_data_merger')
comparison_bp = Blueprint('comparison', __name__)


@comparison_bp.route('/apply_mapping', methods=['POST'])
def apply_mapping():
    """Apply the mapping and update the tw2 database."""
    try:
        data = request.json
        mappings = data.get('mappings', {})

        if not session.get('tw2_file') or not session.get('excel_data'):
            return Response(
                json.dumps({'success': False, 'error': 'Files not loaded'}, ensure_ascii=True),
                mimetype='application/json',
                status=400
            )

        result = apply_excel_mappings_to_tw2(
            session['tw2_file'],
            session['excel_data'],
            mappings
        )

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


@comparison_bp.route('/compare_performance', methods=['POST'])
def compare_performance():
    """Execute performance comparison between Excel and updated TW2 data."""
    try:
        data = request.json or {}
        mbh_lat_lower_margin = float(data.get('mbh_lat_lower_margin', DEFAULT_MBH_LAT_LOWER_MARGIN))
        mbh_lat_upper_margin = float(data.get('mbh_lat_upper_margin', DEFAULT_MBH_LAT_UPPER_MARGIN))
        wpd_threshold = float(data.get('wpd_threshold', DEFAULT_WPD_THRESHOLD))
        apd_threshold = float(data.get('apd_threshold', DEFAULT_APD_THRESHOLD))

        # Check if required data is available
        if not session.get('excel_data'):
            return jsonify({'success': False, 'error': 'Excel data not loaded'}), 400

        reload_info = reload_tw2_data_from_disk(session)
        if not reload_info.get('success'):
            status_code = 404 if reload_info.get('code') == 404 else 500
            return jsonify({'success': False, 'error': 'Unable to reload TW2 data: {}'.format(reload_info.get('error'))}), status_code

        updated_tw2_data = session.get('updated_tw2_data')
        if not updated_tw2_data:
            return jsonify({'success': False, 'error': 'Updated TW2 data not loaded'}), 500

        # Perform comparison
        result = compare_performance_data(
            session['excel_data'],
            updated_tw2_data,
            mbh_lat_lower_margin=mbh_lat_lower_margin,
            mbh_lat_upper_margin=mbh_lat_upper_margin,
            wpd_threshold=wpd_threshold,
            apd_threshold=apd_threshold
        )

        if result['success']:
            return jsonify({
                'success': True,
                'data': {
                    'results': result['results'],
                    'summary': result['summary'],
                    'tw2_path': reload_info.get('path'),
                    'tw2_source': reload_info.get('source'),
                    'tw2_records': reload_info.get('row_count'),
                    'tw2_column_count': reload_info.get('column_count')
                }
            })
        else:
            return jsonify({'success': False, 'error': result['error']}), 500

    except Exception as e:
        logger.exception(f"Error in compare_performance: {str(e)}")
        return jsonify({'success': False, 'error': f'Error during comparison: {str(e)}'}), 500


@comparison_bp.route('/refresh_and_compare', methods=['POST'])
def refresh_and_compare():
    """Refresh TW2 data from stored file path and automatically run comparison."""
    try:
        logger.info("REFRESH: ===== STARTING REFRESH AND COMPARE =====")

        data = request.json or {}
        mbh_lat_lower_margin = float(data.get('mbh_lat_lower_margin', DEFAULT_MBH_LAT_LOWER_MARGIN))
        mbh_lat_upper_margin = float(data.get('mbh_lat_upper_margin', DEFAULT_MBH_LAT_UPPER_MARGIN))
        wpd_threshold = float(data.get('wpd_threshold', DEFAULT_WPD_THRESHOLD))
        apd_threshold = float(data.get('apd_threshold', DEFAULT_APD_THRESHOLD))

        request_original_path = sanitize_path(data.get('original_path')) if data.get('original_path') else None
        preferred_paths = []

        if request_original_path:
            import os
            if os.path.exists(request_original_path):
                session['original_tw2_path'] = request_original_path
                preferred_paths.append(('original', request_original_path))
                logger.info(f"REFRESH: [OK] Using request original path: {request_original_path}")
            else:
                logger.warning(f"REFRESH: [WARNING] Request original path not accessible: {request_original_path}")

        reload_info = reload_tw2_data_from_disk(session, preferred_paths=preferred_paths)
        if not reload_info.get('success'):
            status_code = 404 if reload_info.get('code') == 404 else 500
            logger.error(f"REFRESH: Unable to reload TW2 data: {reload_info.get('error')}")
            return jsonify({'success': False, 'error': reload_info.get('error')}), status_code

        path_source = reload_info.get('source')
        tw2_path = reload_info.get('path')
        logger.info(f"REFRESH: Reloaded TW2 data from {tw2_path} (source: {path_source})")

        if not session.get('excel_data'):
            return jsonify({
                'success': True,
                'data': {
                    'message': 'TW2 data refreshed successfully, but Excel data not loaded for comparison',
                    'tw2_refreshed': True,
                    'comparison_available': False,
                    'path_source': path_source,
                    'tw2_path': tw2_path,
                    'tw2_records': reload_info.get('row_count'),
                    'tw2_column_count': reload_info.get('column_count'),
                    'skipped_read': False
                }
            })

        comparison_result = compare_performance_data(
            session['excel_data'],
            session['updated_tw2_data'],
            mbh_lat_lower_margin=mbh_lat_lower_margin,
            mbh_lat_upper_margin=mbh_lat_upper_margin,
            wpd_threshold=wpd_threshold,
            apd_threshold=apd_threshold
        )

        if comparison_result['success']:
            return jsonify({
                'success': True,
                'data': {
                    'message': 'TW2 data refreshed and comparison completed successfully',
                    'tw2_refreshed': True,
                    'comparison_available': True,
                    'results': comparison_result['results'],
                    'summary': comparison_result['summary'],
                    'path_source': path_source,
                    'tw2_path': tw2_path,
                    'tw2_records': reload_info.get('row_count'),
                    'tw2_column_count': reload_info.get('column_count'),
                    'skipped_read': False
                }
            })
        else:
            return jsonify({
                'success': False,
                'error': 'TW2 data refreshed but comparison failed: {}'.format(comparison_result['error']),
                'data': {
                    'tw2_refreshed': True,
                    'path_source': path_source,
                    'tw2_path': tw2_path
                }
            }), 500

    except Exception as e:
        logger.exception(f"Error in refresh_and_compare: {str(e)}")
        return jsonify({'success': False, 'error': f'Error during refresh and compare: {str(e)}'}), 500


@comparison_bp.route('/save_hw_rows', methods=['POST'])
def save_hw_rows():
    """Save HW Rows changes to the TW2 file."""
    try:
        data = request.get_json(silent=True) or {}
        edits = data.get('edits', [])

        incoming_path = data.get('original_path')
        if incoming_path is not None:
            sanitized_path = sanitize_path(str(incoming_path).strip())
            if sanitized_path:
                session['original_tw2_path'] = sanitized_path
            else:
                return jsonify({
                    'success': False,
                    'error': 'Original TW2 file path is required before saving HW Rows.'
                }), 400

        if not edits:
            return jsonify({'success': False, 'error': 'No edits provided'}), 400

        original_tw2_path = session.get('original_tw2_path')
        if not original_tw2_path:
            return jsonify({
                'success': False,
                'error': 'Original TW2 file path is required before saving HW Rows.'
            }), 400

        import os
        if not os.path.exists(original_tw2_path):
            return jsonify({
                'success': False,
                'error': 'Original TW2 file path is not accessible. Please validate the path and try again.'
            }), 400

        result = save_hw_rows_to_tw2(original_tw2_path, edits)
        return jsonify(result)

    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'Failed to save HW Rows: {str(e)}'
        }), 500


@comparison_bp.route('/export_schedule_data', methods=['POST'])
def export_schedule_data():
    """Export TW2 data as Schedule Data Excel report."""
    try:
        updated_tw2_data = session.get('updated_tw2_data') or session.get('tw2_data')
        if not updated_tw2_data:
            return jsonify({'success': False, 'error': 'TW2 data not loaded'}), 400

        # Get TW2 file path from session (check all possible locations)
        tw2_path = session.get('original_tw2_path') or session.get('updated_tw2_path') or session.get('tw2_file') or ''
        project_name = None

        # Try to get project name from TW2 database first
        if tw2_path:
            project_name = get_project_name_from_tw2(tw2_path)

        # Fallback to filename parsing if database query didn't work
        if not project_name and tw2_path:
            import os
            filename = os.path.splitext(os.path.basename(tw2_path))[0]
            if ' - ' in filename:
                project_name = filename.split(' - ')[0].strip()
            else:
                project_name = filename

        # Final fallback to default name
        if not project_name:
            project_name = 'VAV Schedule Data'

        excel_file = generate_schedule_data_excel(updated_tw2_data, project_name)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"Schedule_Data_{timestamp}.xlsx"

        return send_file(
            excel_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        logger.exception(f"Error in export_schedule_data: {str(e)}")
        return jsonify({'success': False, 'error': f'Error generating report: {str(e)}'}), 500
