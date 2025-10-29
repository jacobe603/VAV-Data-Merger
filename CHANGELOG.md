# Changelog

All notable changes to this project are documented in this file.

## 2025-10-29

### Fixed
- **Excel header mapping**: Added support for 'TAG' column (maps to 'Unit_No') and additional column mappings ('MBH', 'WPD', 'APD') to improve compatibility with various Excel spreadsheet formats. This fixes the comparison table not displaying results when Excel files use 'TAG' instead of 'Unit No'.

### Added
- **Create Schedule Report feature**: New "Create Report" button that generates Excel spreadsheets matching Titus Teams Schedule Data template format. Features include:
  - Template-based export using `templates/Schedule_Data_Template.xlsx` to preserve exact formatting, borders, and alignment
  - Automatic project name detection from TW2 file path
  - All VAV unit performance data (CFM, MBH, LAT, WPD, APD, etc.)
  - Proper handling of merged cells and formatting from Titus Teams template
  - Additional 7th note documenting hot water performance fluid type (Water, Ethylene Glycol, or Propylene Glycol with percentage)
  - One-click download of formatted report
- **setup.bat**: Windows automated setup script that creates virtual environment and installs dependencies with one click.
- **Enhanced documentation**: Comprehensive troubleshooting section, explicit Python 3.9-3.13 compatibility info, and improved setup instructions.

### Changed
- **requirements.txt**: Updated pandas version constraint to `>=2.1.0` (supports Python 3.13+) and pyodbc to `>=5.2.0` for better dependency compatibility.
- **README.md**: Reorganized and expanded with system requirements section, platform-specific setup instructions, production deployment guide, and detailed troubleshooting section.
- **.gitignore**: Expanded to include virtual environments, IDE files, Python build artifacts, and runtime directories.

## 2025-09-04

### Added
- Original path override in `POST /refresh_and_compare` via `original_path` request field; also saved to session when valid.
- Structured logging (`logging` module) with INFO-level output; replaces ad-hoc prints in refresh/compare routes.
- mtime-based skip for refresh to avoid re-reading unchanged TW2 files; response includes `skipped_read` and `path_source`.

### Changed
- Standardized API responses for `/compare_performance` and `/refresh_and_compare` to a consistent shape:
  - Success: `{ success: true, data: { ... } }`
  - Error: `{ success: false, error: '...' }`
- Extracted large inline script from `templates/index.html` to `static/app.js` and updated the template to include it.
- Updated frontend fetch handlers to read from `data.data` and to pass `original_path` during refresh.

### Notes
- If any external tools consumed these endpoints directly, ensure they read from `data` on success and `error` on failure.
- Refresh endpoint now prefers the provided original file path when accessible; otherwise falls back to the local uploaded copy.
