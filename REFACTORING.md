# VAV Data Merger - Refactoring Documentation

## Overview

This document describes the refactoring of the VAV Data Merger application from a monolithic structure to a modular architecture for improved maintainability, testability, and scalability.

## Original Structure

**Before refactoring:**
- **app.py** - 2000 lines (all routes, services, utilities)
- **static/app.js** - 1274 lines (all frontend logic)

## New Modular Structure

```
VAV-Data-Merger/
├── app.py                          # Main application entry point (80 lines)
├── app_old.py                      # Original monolithic file (backup)
│
├── models/                         # Configuration and data models
│   ├── __init__.py
│   └── config.py                   # Application configuration constants
│
├── utils/                          # Utility functions
│   ├── __init__.py
│   ├── path_utils.py               # Path and file handling
│   ├── data_utils.py               # Data normalization and conversion
│   └── db_utils.py                 # Database connection utilities
│
├── services/                       # Business logic layer
│   ├── __init__.py
│   ├── tw2_service.py              # TW2 file reading and processing
│   ├── excel_service.py            # Excel file reading and header mapping
│   ├── comparison_service.py       # Performance comparison logic
│   ├── database_service.py         # Database update operations
│   └── report_service.py           # Excel report generation
│
├── routes/                         # Flask route handlers
│   ├── __init__.py
│   ├── main_routes.py              # Main page route
│   ├── tw2_routes.py               # TW2 file upload/handling
│   ├── excel_routes.py             # Excel file upload/handling
│   ├── comparison_routes.py        # Comparison and mapping routes
│   └── debug_routes.py             # Debug endpoints
│
├── templates/                      # HTML templates
│   └── index.html
│
└── static/                         # Static assets
    ├── css/
    ├── js/
    │   ├── app.js                  # (To be refactored into modules)
    │   └── modules/                # Future modular JavaScript files
    └── VAV_Data_Merger_Instructions.md
```

## Module Breakdown

### 1. **app.py** (Main Application - 80 lines)
- Application factory pattern
- Blueprint registration
- Configuration setup
- Session initialization
- Clean entry point

### 2. **models/config.py** (Configuration)
- Application settings (SECRET_KEY, DEBUG, PORT)
- Session configuration
- File upload settings
- Performance comparison thresholds
- Field mappings and descriptions
- Environment variable support

### 3. **utils/** (Utility Functions)

#### **path_utils.py**
- `sanitize_path()` - Remove quotes from file paths
- `allowed_file()` - Validate file extensions

#### **data_utils.py**
- `CustomJSONEncoder` - JSON serialization for special types
- `safe_string_convert()` - Safe data type conversion
- `normalize_tag_format()` - Tag format normalization (V-1-1 → V-1-01)
- `normalize_unit_tag()` - Unit tag normalization with regex
- `clean_size_value()` - Remove inch marks and format sizes
- `normalize_header_text()` - Clean Excel headers
- `is_probably_header_value()` - Heuristic header detection
- `normalize_hw_rows_value()` - Normalize HW Rows values

#### **db_utils.py**
- `get_mdb_connection()` - Create Access database connections
- `get_project_name_from_tw2()` - Query project name from database

### 4. **services/** (Business Logic)

#### **tw2_service.py**
- `read_tw2_data_safe()` - Read TW2 database files safely
- `reload_tw2_data_from_disk()` - Reload TW2 data with fallback paths

#### **excel_service.py**
- `combine_multi_row_headers()` - Merge multi-row Excel headers
- `map_excel_headers_to_standard()` - Map Excel headers to standard names
- `read_excel_data_safe()` - Read and process Excel files

#### **comparison_service.py**
- `compare_performance_data()` - Compare Excel vs TW2 performance metrics

#### **database_service.py**
- `apply_excel_mappings_to_tw2()` - Apply field mappings to database
- `save_hw_rows_to_tw2()` - Update HW Rows in database

#### **report_service.py**
- `generate_schedule_data_excel()` - Generate formatted Excel reports

### 5. **routes/** (HTTP Endpoints)

#### **main_routes.py**
- `GET /` - Main application page

#### **tw2_routes.py**
- `POST /upload_tw2` - Upload TW2 file
- `POST /upload_updated_tw2` - Upload updated TW2 file
- `GET /get_updated_tw2_data` - Retrieve updated TW2 data
- `GET /download_merged_tw2` - Download merged database
- `POST /validate_tw2_path` - Validate file path

#### **excel_routes.py**
- `POST /upload_excel` - Upload Excel file
- `GET /get_mapping_fields` - Get field mapping options

#### **comparison_routes.py**
- `POST /apply_mapping` - Apply field mappings
- `POST /compare_performance` - Run performance comparison
- `POST /refresh_and_compare` - Refresh data and compare
- `POST /save_hw_rows` - Save HW Rows edits
- `POST /export_schedule_data` - Export Excel report

#### **debug_routes.py**
- `GET /debug_excel` - Debug Excel file reading
- `GET /debug_headers` - Debug header processing
- `GET /debug_data` - Debug data extraction
- `GET /debug_session` - Inspect session data
- `POST /clear_session` - Clear session
- `POST /test_large_session` - Test session storage

## Benefits of Refactoring

### 1. **Improved Maintainability**
- Smaller, focused files (80-300 lines each)
- Clear separation of concerns
- Easier to locate and fix bugs
- Better code organization

### 2. **Enhanced Testability**
- Services can be unit tested independently
- Utilities can be tested in isolation
- Routes can be tested with mocked services
- Clear boundaries for mocking

### 3. **Better Scalability**
- Easy to add new routes/services
- Modular components can be reused
- Configuration centralized for easy updates
- Blueprint pattern allows feature modules

### 4. **Improved Collaboration**
- Multiple developers can work on different modules
- Less merge conflicts
- Clear module ownership
- Better code review process

### 5. **Easier Onboarding**
- New developers can understand one module at a time
- Clear structure makes codebase navigation easier
- Self-documenting architecture
- Logical grouping of related functionality

## Migration Notes

### Running the Refactored Application

The refactored application is backwards compatible. No changes needed to:
- Templates (index.html)
- Static assets (CSS, JavaScript)
- Database files
- Excel templates
- User workflows

### Configuration

Environment variables can now be used for configuration:
```bash
export VAV_SECRET_KEY="your-secret-key"
export DEBUG="False"
export PORT="5004"
```

### Testing

To test the refactored application:
```bash
# Activate virtual environment
source .venv/bin/activate  # or .venv\Scripts\activate on Windows

# Run the application
python app.py
```

## Future Improvements

### Phase 2: JavaScript Refactoring (Pending)
- Split `static/app.js` into modules:
  - `file-upload.js` - File upload handling
  - `data-preview.js` - Data table rendering
  - `comparison.js` - Performance comparison
  - `hw-rows-editor.js` - HW Rows editing
  - `utils.js` - Shared utilities

### Phase 3: Testing Infrastructure
- Add pytest for unit tests
- Create test fixtures for sample data
- Add integration tests for routes
- Set up CI/CD pipeline

### Phase 4: Additional Enhancements
- Add type hints (mypy)
- Add docstring validation
- Create API documentation (Swagger/OpenAPI)
- Add logging configuration file
- Implement caching for frequently accessed data

## Version History

- **v2.0.0** (2025-11-06) - Modular architecture refactoring
- **v1.x** - Original monolithic implementation

## Contributors

- Refactoring architecture: Claude (Anthropic)
- Original implementation: VAV Data Merger Team

---

For questions or issues with the refactored codebase, please refer to the README.md or create an issue in the project repository.
