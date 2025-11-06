"""
Test script to verify the refactored VAV Data Merger application.
Tests module imports, utility functions, and basic application startup.
"""
import sys
import os

# Add project root to path
sys.path.insert(0, '/home/user/VAV-Data-Merger')

print("=" * 70)
print("VAV Data Merger - Refactoring Test Suite")
print("=" * 70)

# Test 1: Module Imports
print("\n[TEST 1] Testing module imports...")
try:
    from models import config
    print("✓ models.config imported successfully")

    from utils import path_utils, data_utils, db_utils
    print("✓ utils.path_utils imported successfully")
    print("✓ utils.data_utils imported successfully")
    print("✓ utils.db_utils imported successfully")

    from services import tw2_service, excel_service, comparison_service
    from services import database_service, report_service
    print("✓ services.tw2_service imported successfully")
    print("✓ services.excel_service imported successfully")
    print("✓ services.comparison_service imported successfully")
    print("✓ services.database_service imported successfully")
    print("✓ services.report_service imported successfully")

    from routes import main_routes, tw2_routes, excel_routes
    from routes import comparison_routes, debug_routes
    print("✓ routes.main_routes imported successfully")
    print("✓ routes.tw2_routes imported successfully")
    print("✓ routes.excel_routes imported successfully")
    print("✓ routes.comparison_routes imported successfully")
    print("✓ routes.debug_routes imported successfully")

    print("\n✅ All module imports successful!")

except Exception as e:
    print(f"\n❌ Import failed: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Test 2: Utility Functions
print("\n[TEST 2] Testing utility functions...")
try:
    # Test path sanitization
    test_path = '"C:\\Users\\test\\file.txt"'
    sanitized = path_utils.sanitize_path(test_path)
    assert sanitized == 'C:\\Users\\test\\file.txt', "Path sanitization failed"
    print(f"✓ sanitize_path: {test_path} → {sanitized}")

    # Test file extension checking
    assert path_utils.allowed_file('test.xlsx') == True
    assert path_utils.allowed_file('test.txt') == False
    print("✓ allowed_file validation works")

    # Test tag normalization
    tag = "V-1-1"
    normalized = data_utils.normalize_tag_format(tag)
    assert normalized == "V-1-01", f"Expected V-1-01, got {normalized}"
    print(f"✓ normalize_tag_format: {tag} → {normalized}")

    # Test size value cleaning
    size = '8"'
    cleaned = data_utils.clean_size_value(size)
    assert cleaned == "08", f"Expected 08, got {cleaned}"
    print(f"✓ clean_size_value: {size} → {cleaned}")

    # Test HW rows normalization
    assert data_utils.normalize_hw_rows_value(2) == 2
    assert data_utils.normalize_hw_rows_value("3") == 3
    assert data_utils.normalize_hw_rows_value(None) == None
    print("✓ normalize_hw_rows_value works")

    print("\n✅ All utility function tests passed!")

except Exception as e:
    print(f"\n❌ Utility function test failed: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Test 3: Configuration
print("\n[TEST 3] Testing configuration...")
try:
    assert hasattr(config, 'SECRET_KEY'), "SECRET_KEY not in config"
    assert hasattr(config, 'UPLOAD_FOLDER'), "UPLOAD_FOLDER not in config"
    assert hasattr(config, 'ALLOWED_EXTENSIONS'), "ALLOWED_EXTENSIONS not in config"
    assert hasattr(config, 'TARGET_FIELDS'), "TARGET_FIELDS not in config"

    print(f"✓ Config loaded - Port: {config.PORT}, Debug: {config.DEBUG}")
    print(f"✓ Upload folder: {config.UPLOAD_FOLDER}")
    print(f"✓ Allowed extensions: {config.ALLOWED_EXTENSIONS}")
    print(f"✓ Target fields: {len(config.TARGET_FIELDS)} fields")

    print("\n✅ Configuration test passed!")

except Exception as e:
    print(f"\n❌ Configuration test failed: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Test 4: Service Functions with Mock Data
print("\n[TEST 4] Testing service functions with mock data...")
try:
    # Test comparison service with mock data
    excel_data = [
        {'Unit_No': 'V-1-1', 'MBH': 100, 'LAT': 80},
        {'Unit_No': 'V-1-2', 'MBH': 120, 'LAT': 85}
    ]

    tw2_data = [
        {'Tag': 'V-1-01', 'HWMBHCalc': 105, 'HWLATCalc': 82, 'HWPDCalc': 3.5, 'HWAPDCalc': 0.2},
        {'Tag': 'V-1-02', 'HWMBHCalc': 125, 'HWLATCalc': 87, 'HWPDCalc': 4.0, 'HWAPDCalc': 0.15}
    ]

    result = comparison_service.compare_performance_data(excel_data, tw2_data)

    assert result['success'] == True, "Comparison should succeed"
    assert 'results' in result, "Should have results"
    assert 'summary' in result, "Should have summary"
    assert len(result['results']) == 2, "Should have 2 comparison results"

    print("✓ Performance comparison logic works")
    print(f"  - Compared {len(result['results'])} units")
    print(f"  - Summary: {result['summary']}")

    print("\n✅ Service function tests passed!")

except Exception as e:
    print(f"\n❌ Service function test failed: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print("\n" + "=" * 70)
print("✅ ALL TESTS PASSED - Refactored code is working!")
print("=" * 70)
print("\nNext step: Test Flask application startup...")
