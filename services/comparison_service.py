"""Service for performance comparison between Excel and TW2 data."""
import pandas as pd
from utils.data_utils import normalize_unit_tag, normalize_hw_rows_value


def compare_performance_data(excel_data, updated_tw2_data, mbh_lat_lower_margin=15,
                            mbh_lat_upper_margin=25, wpd_threshold=5, apd_threshold=0.25):
    """Compare performance values between Excel and updated TW2 data.

    Args:
        excel_data: List of dictionaries with Excel data
        updated_tw2_data: List of dictionaries with updated TW2 data
        mbh_lat_lower_margin: Lower margin for MBH/LAT comparison (default 15%)
        mbh_lat_upper_margin: Upper margin for MBH/LAT comparison (default 25%)
        wpd_threshold: Water pressure drop threshold (default 5)
        apd_threshold: Air pressure drop threshold (default 0.25)

    Returns:
        Dictionary with 'success', 'results', 'summary' keys
    """
    try:
        comparison_results = []

        # Convert updated TW2 data to DataFrame for easier processing
        updated_tw2_df = pd.DataFrame(updated_tw2_data)
        excel_df = pd.DataFrame(excel_data)

        # Create index mapping for faster lookups with normalized tags
        tw2_index = {}
        for row in updated_tw2_data:
            original_tag = str(row.get('Tag', '')).strip()
            normalized_tag = normalize_unit_tag(original_tag)
            tw2_index[normalized_tag] = row
            # Also keep original tag as backup
            tw2_index[original_tag] = row

        for excel_row in excel_data:
            unit_tag = str(excel_row.get('Unit_No', '')).strip()
            if not unit_tag:
                continue

            # Normalize the Excel unit tag for matching
            normalized_excel_tag = normalize_unit_tag(unit_tag)

            # Find matching TW2 record - try normalized first, then original
            tw2_row = tw2_index.get(normalized_excel_tag) or tw2_index.get(unit_tag)
            if not tw2_row:
                comparison_results.append({
                    'unit_tag': f"{unit_tag} → {normalized_excel_tag}" if normalized_excel_tag != unit_tag else unit_tag,
                    'status': 'Not Found',
                    'excel_mbh': excel_row.get('MBH', 'N/A'),
                    'tw2_mbh': 'N/A',
                    'mbh_diff': 'N/A',
                    'excel_lat': excel_row.get('LAT', 'N/A'),
                    'tw2_lat': 'N/A',
                    'lat_diff': 'N/A',
                    'tw2_wpd': 'N/A',
                    'tw2_apd': 'N/A',
                    'tw2_hw_rows': None,
                    'tw2_hw_rows_raw': None,
                })
                continue

            # Extract values
            excel_mbh = excel_row.get('MBH') or excel_row.get('MBH_Total') or excel_row.get('Total_MBH')
            excel_lat = excel_row.get('LAT') or excel_row.get('Leaving_Air_Temp')

            tw2_mbh = tw2_row.get('HWMBHCalc')
            tw2_lat = tw2_row.get('HWLATCalc')
            tw2_wpd = tw2_row.get('HWPDCalc')
            tw2_apd = tw2_row.get('HWAPDCalc')

            tw2_hw_raw = None
            for hw_key in ('HWRowsCalc', 'HWRows', 'HWRow'):
                candidate = tw2_row.get(hw_key)
                if candidate not in (None, ''):
                    tw2_hw_raw = candidate
                    break
            tw2_hw_rows = normalize_hw_rows_value(tw2_hw_raw)

            # Calculate differences and status
            mbh_diff = 'N/A'
            lat_diff = 'N/A'
            status_flags = []

            # MBH comparison with separate upper/lower margins
            if excel_mbh is not None and tw2_mbh is not None:
                try:
                    excel_mbh_val = float(excel_mbh)
                    tw2_mbh_val = float(tw2_mbh)
                    if excel_mbh_val != 0:
                        mbh_diff = ((tw2_mbh_val - excel_mbh_val) / excel_mbh_val) * 100
                        # Check if outside acceptable range: -15% to +25%
                        if mbh_diff < -mbh_lat_lower_margin:  # Too low (under by more than 15%)
                            status_flags.append(f'MBH {mbh_diff:.1f}% (too low)')
                        elif mbh_diff > mbh_lat_upper_margin:  # Too high (over by more than 25%)
                            status_flags.append(f'MBH {mbh_diff:.1f}% (too high)')
                except (ValueError, TypeError):
                    pass

            # LAT comparison with separate upper/lower margins
            if excel_lat is not None and tw2_lat is not None:
                try:
                    excel_lat_val = float(excel_lat)
                    tw2_lat_val = float(tw2_lat)
                    if excel_lat_val != 0:
                        lat_diff = ((tw2_lat_val - excel_lat_val) / excel_lat_val) * 100
                        # Check if outside acceptable range: -15% to +25%
                        if lat_diff < -mbh_lat_lower_margin:  # Too low (under by more than 15%)
                            status_flags.append(f'LAT {lat_diff:.1f}% (too low)')
                        elif lat_diff > mbh_lat_upper_margin:  # Too high (over by more than 25%)
                            status_flags.append(f'LAT {lat_diff:.1f}% (too high)')
                except (ValueError, TypeError):
                    pass

            # WPD check
            if tw2_wpd is not None:
                try:
                    wpd_val = float(tw2_wpd)
                    if wpd_val > wpd_threshold:
                        status_flags.append(f'WPD {wpd_val:.2f}')
                except (ValueError, TypeError):
                    pass

            # APD check
            if tw2_apd is not None:
                try:
                    apd_val = float(tw2_apd)
                    if apd_val > apd_threshold:
                        status_flags.append(f'APD {apd_val:.2f}')
                except (ValueError, TypeError):
                    pass

            # Determine overall status
            if status_flags:
                status = 'Fail' if any('MBH' in flag or 'LAT' in flag for flag in status_flags) else 'Warning'
            else:
                status = 'Pass'

            comparison_results.append({
                'unit_tag': f"{unit_tag} → {normalized_excel_tag}" if normalized_excel_tag != unit_tag else unit_tag,
                'status': status,
                'status_details': ', '.join(status_flags) if status_flags else 'All within range',
                'excel_mbh': excel_mbh,
                'tw2_mbh': tw2_mbh,
                'mbh_diff': f'{mbh_diff:.1f}%' if isinstance(mbh_diff, (int, float)) else mbh_diff,
                'excel_lat': excel_lat,
                'tw2_lat': tw2_lat,
                'lat_diff': f'{lat_diff:.1f}%' if isinstance(lat_diff, (int, float)) else lat_diff,
                'tw2_wpd': tw2_wpd,
                'tw2_apd': tw2_apd,
                'tw2_hw_rows': tw2_hw_rows,
                'tw2_hw_rows_raw': tw2_hw_raw if tw2_hw_raw not in (None, '') else None
            })

        return {
            'success': True,
            'results': comparison_results,
            'summary': {
                'total': len(comparison_results),
                'pass': len([r for r in comparison_results if r['status'] == 'Pass']),
                'warning': len([r for r in comparison_results if r['status'] == 'Warning']),
                'fail': len([r for r in comparison_results if r['status'] == 'Fail']),
                'not_found': len([r for r in comparison_results if r['status'] == 'Not Found'])
            }
        }

    except Exception as e:
        return {
            'success': False,
            'error': f'Error during performance comparison: {str(e)}'
        }
