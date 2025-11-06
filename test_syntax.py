"""
Syntax and import structure test for refactored code.
Tests that all modules have valid Python syntax and proper import structure.
"""
import sys
import os
import py_compile
import importlib.util

print("=" * 70)
print("VAV Data Merger - Syntax & Structure Validation")
print("=" * 70)

# List of all Python files to check
python_files = [
    'app.py',
    'models/config.py',
    'utils/path_utils.py',
    'utils/data_utils.py',
    'utils/db_utils.py',
    'services/tw2_service.py',
    'services/excel_service.py',
    'services/comparison_service.py',
    'services/database_service.py',
    'services/report_service.py',
    'routes/main_routes.py',
    'routes/tw2_routes.py',
    'routes/excel_routes.py',
    'routes/comparison_routes.py',
    'routes/debug_routes.py',
]

print("\n[TEST 1] Checking Python syntax...")
syntax_errors = []
for file_path in python_files:
    full_path = f'/home/user/VAV-Data-Merger/{file_path}'
    try:
        py_compile.compile(full_path, doraise=True)
        print(f"✓ {file_path}")
    except py_compile.PyCompileError as e:
        print(f"✗ {file_path}: {e}")
        syntax_errors.append((file_path, str(e)))

if syntax_errors:
    print(f"\n❌ Found {len(syntax_errors)} syntax errors!")
    for file_path, error in syntax_errors:
        print(f"  - {file_path}: {error}")
    sys.exit(1)
else:
    print(f"\n✅ All {len(python_files)} files have valid Python syntax!")

print("\n[TEST 2] Checking import structure...")

# Test configuration module (no external deps)
print("\nTesting models/config.py...")
try:
    spec = importlib.util.spec_from_file_location("config", "/home/user/VAV-Data-Merger/models/config.py")
    config = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(config)

    # Check for required attributes
    required_attrs = [
        'SECRET_KEY', 'DEBUG', 'PORT',
        'UPLOAD_FOLDER', 'ALLOWED_EXTENSIONS',
        'DEFAULT_MBH_LAT_LOWER_MARGIN', 'DEFAULT_MBH_LAT_UPPER_MARGIN',
        'TARGET_FIELDS', 'SUGGESTED_FIELD_MAPPINGS', 'FIELD_DESCRIPTIONS'
    ]

    for attr in required_attrs:
        assert hasattr(config, attr), f"Missing required attribute: {attr}"
        print(f"  ✓ {attr} = {getattr(config, attr)[:50] if isinstance(getattr(config, attr), str) else getattr(config, attr)}")

    print("✅ Configuration module validated!")

except Exception as e:
    print(f"❌ Configuration test failed: {e}")
    sys.exit(1)

print("\n[TEST 3] Checking Flask application structure...")
try:
    # Read app.py and check for key components
    with open('/home/user/VAV-Data-Merger/app.py', 'r') as f:
        app_content = f.read()

    required_imports = [
        'from flask import Flask',
        'from flask_session import Session',
        'from routes.main_routes import main_bp',
        'from routes.tw2_routes import tw2_bp',
        'from routes.excel_routes import excel_bp',
        'from routes.comparison_routes import comparison_bp',
        'from routes.debug_routes import debug_bp',
    ]

    for import_line in required_imports:
        if import_line in app_content:
            print(f"  ✓ {import_line}")
        else:
            print(f"  ✗ Missing: {import_line}")
            raise ValueError(f"Missing import: {import_line}")

    # Check for blueprint registration
    required_registrations = [
        'app.register_blueprint(main_bp)',
        'app.register_blueprint(tw2_bp)',
        'app.register_blueprint(excel_bp)',
        'app.register_blueprint(comparison_bp)',
        'app.register_blueprint(debug_bp)',
    ]

    for reg in required_registrations:
        if reg in app_content:
            print(f"  ✓ {reg}")
        else:
            print(f"  ✗ Missing: {reg}")
            raise ValueError(f"Missing registration: {reg}")

    print("✅ Flask application structure validated!")

except Exception as e:
    print(f"❌ Application structure test failed: {e}")
    sys.exit(1)

print("\n[TEST 4] Checking route endpoint definitions...")
route_files = {
    'main_routes.py': ['@main_bp.route(\'/\')'],
    'tw2_routes.py': [
        '@tw2_bp.route(\'/upload_tw2\'',
        '@tw2_bp.route(\'/upload_updated_tw2\'',
        '@tw2_bp.route(\'/get_updated_tw2_data\'',
        '@tw2_bp.route(\'/download_merged_tw2\'',
        '@tw2_bp.route(\'/validate_tw2_path\'',
    ],
    'excel_routes.py': [
        '@excel_bp.route(\'/upload_excel\'',
        '@excel_bp.route(\'/get_mapping_fields\'',
    ],
    'comparison_routes.py': [
        '@comparison_bp.route(\'/apply_mapping\'',
        '@comparison_bp.route(\'/compare_performance\'',
        '@comparison_bp.route(\'/refresh_and_compare\'',
        '@comparison_bp.route(\'/save_hw_rows\'',
        '@comparison_bp.route(\'/export_schedule_data\'',
    ],
    'debug_routes.py': [
        '@debug_bp.route(\'/debug_excel\'',
        '@debug_bp.route(\'/debug_session\'',
    ],
}

for route_file, expected_routes in route_files.items():
    print(f"\nChecking {route_file}...")
    with open(f'/home/user/VAV-Data-Merger/routes/{route_file}', 'r') as f:
        content = f.read()

    for route in expected_routes:
        if route in content:
            print(f"  ✓ {route}")
        else:
            print(f"  ✗ Missing: {route}")
            raise ValueError(f"Missing route in {route_file}: {route}")

print("\n✅ All route endpoints validated!")

print("\n[TEST 5] Checking service function definitions...")
service_functions = {
    'tw2_service.py': ['read_tw2_data_safe', 'reload_tw2_data_from_disk'],
    'excel_service.py': ['combine_multi_row_headers', 'map_excel_headers_to_standard', 'read_excel_data_safe'],
    'comparison_service.py': ['compare_performance_data'],
    'database_service.py': ['apply_excel_mappings_to_tw2', 'save_hw_rows_to_tw2'],
    'report_service.py': ['generate_schedule_data_excel'],
}

for service_file, expected_functions in service_functions.items():
    print(f"\nChecking {service_file}...")
    with open(f'/home/user/VAV-Data-Merger/services/{service_file}', 'r') as f:
        content = f.read()

    for func in expected_functions:
        if f'def {func}(' in content:
            print(f"  ✓ {func}()")
        else:
            print(f"  ✗ Missing: {func}()")
            raise ValueError(f"Missing function in {service_file}: {func}")

print("\n✅ All service functions validated!")

print("\n[TEST 6] Checking utility function definitions...")
utility_functions = {
    'path_utils.py': ['sanitize_path', 'allowed_file'],
    'data_utils.py': [
        'CustomJSONEncoder', 'safe_string_convert', 'normalize_tag_format',
        'normalize_unit_tag', 'clean_size_value', 'normalize_header_text',
        'is_probably_header_value', 'normalize_hw_rows_value'
    ],
    'db_utils.py': ['get_mdb_connection', 'get_project_name_from_tw2'],
}

for util_file, expected_functions in utility_functions.items():
    print(f"\nChecking {util_file}...")
    with open(f'/home/user/VAV-Data-Merger/utils/{util_file}', 'r') as f:
        content = f.read()

    for func in expected_functions:
        # Check for both def and class
        if f'def {func}(' in content or f'class {func}' in content:
            print(f"  ✓ {func}")
        else:
            print(f"  ✗ Missing: {func}")
            raise ValueError(f"Missing function in {util_file}: {func}")

print("\n✅ All utility functions validated!")

print("\n" + "=" * 70)
print("✅ ALL VALIDATION TESTS PASSED!")
print("=" * 70)
print("\nSummary:")
print(f"  - {len(python_files)} Python files validated")
print(f"  - Configuration module verified")
print(f"  - Flask application structure confirmed")
print(f"  - All route endpoints present")
print(f"  - All service functions defined")
print(f"  - All utility functions defined")
print("\n✅ The refactored code structure is correct and ready to run!")
print("\nNote: To fully test functionality, install dependencies with:")
print("  pip install -r requirements.txt")
