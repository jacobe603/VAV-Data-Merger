"""Configuration constants for the application."""
import os

# Application configuration
SECRET_KEY = os.environ.get('VAV_SECRET_KEY', 'vav-data-merger-secret-key-2025')
DEBUG = os.environ.get('DEBUG', 'True').lower() == 'true'
PORT = int(os.environ.get('PORT', 5004))

# Session configuration
SESSION_TYPE = 'filesystem'
SESSION_FILE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'sessions')
SESSION_PERMANENT = False
SESSION_USE_SIGNER = True
SESSION_KEY_PREFIX = 'vav-merger:'

# File upload configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'tw2', 'mdb'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16 MB

# Performance comparison default thresholds
DEFAULT_MBH_LAT_LOWER_MARGIN = 15  # percentage
DEFAULT_MBH_LAT_UPPER_MARGIN = 25  # percentage
DEFAULT_WPD_THRESHOLD = 5  # water pressure drop
DEFAULT_APD_THRESHOLD = 0.25  # air pressure drop

# Excel processing defaults
DEFAULT_DATA_START_ROW = 3
DEFAULT_HEADER_ROWS = 2
DEFAULT_SKIP_TITLE_ROW = True

# Field mappings
SUGGESTED_FIELD_MAPPINGS = {
    'Tag': 'Unit_No',
    'UnitSize': 'Unit_Size',
    'InletSize': 'Unit_Size',  # Both UnitSize and InletSize map to Unit_Size with special logic
    'CFMDesign': 'CFM_Max',
    'CFMMinPrime': 'CFM_Min',
    'CFMMin': 'CFM_Min',  # Alternative field name for CFM Min
    'HeatingPrimaryAirflow': 'CFM_Heat',  # Alternative field name for heating airflow
    'HWCFM': 'CFM_Heat',
    'HWGPM': 'GPM'
}

# Target fields for mapping
TARGET_FIELDS = [
    'Tag', 'UnitSize', 'InletSize', 'CFMDesign', 'CFMMinPrime',
    'HWCFM', 'HWGPM', 'HeatingPrimaryAirflow', 'CFMMin'
]

# Field descriptions for tooltips
FIELD_DESCRIPTIONS = {
    'Tag': 'Unit identifier - Must match Excel Unit_No (e.g., V-1-1 → V-1-01)',
    'UnitSize': 'VAV unit size designation (e.g., 6, 8, 10, 14, 24x16)',
    'InletSize': 'Air inlet size in inches (e.g., 6", 8", 10", 14")',
    'CFMDesign': 'Design air flow rate in CFM - Maximum airflow capacity',
    'CFMMinPrime': 'Minimum primary airflow in CFM when heating/cooling',
    'HeatingPrime': 'Primary airflow during heating mode in CFM',
    'HWCFM': 'Hot water coil airflow in CFM - Airflow through heating coil',
    'HWGPM': 'Hot water flow rate in GPM (gallons per minute)'
}
